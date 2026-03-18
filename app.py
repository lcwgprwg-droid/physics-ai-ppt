import os
import re
import cv2
import math
import numpy as np
import tempfile
import io
import time
import gc
import streamlit as st
import fitz  # PyMuPDF
import docx  # Python-docx
from rapidocr_onnxruntime import RapidOCR
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from PIL import Image, ImageEnhance

# ==========================================
# 1. 核心引擎库：全局缓存 OCR 模型
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 增强型图像处理逻辑
# ==========================================
def get_distance(box1_center, box2_center):
    """计算两个点之间的欧式距离"""
    return math.sqrt((box1_center[0] - box2_center[0])**2 + (box1_center[1] - box2_center[1])**2)

def is_box_close(b1, b2, threshold=60):
    """判断两个矩形框是否足够接近以合并"""
    # b = [x1, y1, x2, y2]
    return not (b1[2] < b2[0] - threshold or b1[0] > b2[2] + threshold or 
                b1[3] < b2[1] - threshold or b1[1] > b2[3] + threshold)

def merge_boxes(rects, threshold=50):
    if not rects: return []
    res = []
    while rects:
        curr = rects.pop(0)
        has_merged = False
        for i in range(len(res)):
            if is_box_close(curr, res[i], threshold):
                res[i] = [min(curr[0], res[i][0]), min(curr[1], res[i][1]),
                          max(curr[2], res[i][2]), max(curr[3], res[i][3])]
                has_merged = True
                break
        if not has_merged:
            res.append(curr)
    return res

def extract_images_v2(img_path, ocr_result, out_dir, pdf_page=None):
    """
    针对物理题图优化的提取算法
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 预处理：灰度与边缘检测
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # 创建文字遮罩 - 核心改进：排除掉物理图示中的短字符标签
    text_mask = np.zeros((h_img, w_img), dtype=np.uint8)
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            text = line[1].strip()
            # 物理图示中常见的字母标注通常很短且不含汉字，不遮蔽它们，否则图会断开
            if len(text) <= 2 and not re.search(r'[\u4e00-\u9fa5]', text):
                continue
            cv2.fillPoly(text_mask, [box], 255)

    # 形态学操作：寻找除了文字以外的连通区域
    # 使用Canny算法能更好地抓住细线（如绳子、斜面）
    edges = cv2.Canny(gray, 30, 150)
    edges[text_mask > 0] = 0 # 剔除文字区的边缘干扰
    
    # 动态调整核大小：根据图片宽度适配
    k_size = max(3, int(w_img / 150))
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (k_size, k_size))
    # 闭运算连接断裂线条
    morphed = cv2.morphologyEx(edges, cv2.MORPH_CLOSE, kernel)
    morphed = cv2.dilate(morphed, kernel, iterations=2)

    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    raw_boxes = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        # 物理图面积过滤：太小的是噪点，太大的是背景
        if w > 40 and h > 40 and (w * h) > 2000:
            if w < w_img * 0.9 and h < h_img * 0.9:
                raw_boxes.append([x, y, x + w, y + h])

    final_rects = merge_boxes(raw_boxes, threshold=60)
    
    saved_imgs = []
    for i, box in enumerate(final_rects):
        # 增加 Padding 以包含完整的物理标注
        pad = 15
        x1, y1, x2, y2 = max(0, box[0]-pad), max(0, box[1]-pad), min(w_img, box[2]+pad), min(h_img, box[3]+pad)
        
        p = os.path.join(out_dir, f"fig_{int(time.time()*1000)}_{i}.png")
        
        # 针对 PDF 导出高清原图
        if pdf_page is not None:
            # 这里的坐标缩放需根据 OCR 识别时的 dpi 调整
            scale = h_img / pdf_page.rect.height
            pdf_rect = fitz.Rect(x1/scale, y1/scale, x2/scale, y2/scale)
            pix = pdf_page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), clip=pdf_rect)
            pix.save(p)
        else:
            roi = img[y1:y2, x1:x2]
            cv2.imwrite(p, roi)

        saved_imgs.append({
            "path": p, 
            "center": ((x1+x2)/2, (y1+y2)/2),
            "area": (x2-x1)*(y2-y1)
        })
    
    return saved_imgs

# ==========================================
# 3. 智能题目解析与图文绑定
# ==========================================
def smart_process(img_path, temp_dir, pdf_page=None):
    engine = get_ocr_engine()
    result, _ = engine(img_path)
    
    # 1. 提取图片
    extracted_images = extract_images_v2(img_path, result, temp_dir, pdf_page)
    
    # 2. 识别题目文本（处理分栏）
    img_cv = cv2.imread(img_path)
    h_page, w_page = img_cv.shape[:2]
    mid_x = w_page / 2

    lines = []
    if result:
        for line in result:
            box, text = line[0], line[1]
            cx = (box[0][0] + box[1][0]) / 2
            cy = (box[0][1] + box[2][1]) / 2
            col = 0 if cx < mid_x else 1
            lines.append({"col": col, "y": cy, "text": text.strip(), "box": box})

    # 排序：先分栏再纵向
    lines.sort(key=lambda x: (x['col'], x['y']))

    questions = []
    current_q = None

    # 题号匹配正则（物理题常见题号格式）
    q_pattern = r'^\s*\d+[\.．、\(（]'

    for line in lines:
        text = line['text']
        if not text: continue
        
        # 噪声过滤
        if any(kw in text for kw in ['扫描', '答案', '公众号', '页码', '物理']): continue

        if re.match(q_pattern, text):
            if current_q: questions.append(current_q)
            current_q = {
                "text": text, 
                "center": ((line['box'][0][0]+line['box'][1][0])/2, (line['box'][0][1]+line['box'][2][1])/2),
                "matched_imgs": [],
                "col": line['col']
            }
        else:
            if current_q:
                current_q["text"] += "\n" + text
                # 更新题目中心位置（趋向于题目主体）
                new_y = (current_q["center"][1] + line['y']) / 2
                current_q["center"] = (current_q["center"][0], new_y)

    if current_q: questions.append(current_q)

    # 3. 核心改进：图文引力绑定算法
    for img in extracted_images:
        img_cx, img_cy = img['center']
        img_col = 0 if img_cx < mid_x else 1
        
        best_q = None
        min_dist = float('inf')

        for q in questions:
            # 规则1：同栏优先
            col_penalty = 0 if q['col'] == img_col else 1000
            
            # 规则2：物理图通常在题干下方，计算垂直距离
            # 如果图片在题目下方，dist 为正且较小
            v_dist = img_cy - q['center'][1]
            if v_dist < -50: # 图片在题目上方太远，不合理
                v_dist = abs(v_dist) * 2 
            
            h_dist = abs(img_cx - q['center'][0])
            
            dist = math.sqrt(h_dist**2 + v_dist**2) + col_penalty
            
            if dist < min_dist:
                min_dist = dist
                best_q = q
        
        if best_q:
            best_q['matched_imgs'].append(img['path'])

    return questions

# ==========================================
# 4. PPT 渲染引擎 (保持你的精美样式)
# ==========================================
def set_font_all(run, font_name='微软雅黑'):
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    f = rPr.find(qn('w:rFonts'))
    if f is None:
        f = rPr.makeelement(qn('w:rFonts'))
        rPr.append(f)
    f.set(qn('w:eastAsia'), font_name)

def create_slide(prs, title):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 背景装饰
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(250, 250, 252); bg.line.fill.background()
    # 标题栏
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.6))
    p = tb.text_frame.paragraphs[0]; p.text = title
    p.font.bold = True; p.font.size = Pt(24); p.font.color.rgb = RGBColor(20, 40, 80)
    set_font_all(p.runs[0])
    return slide

def render_ppt(questions, output_buffer):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    for i, q in enumerate(questions):
        slide = create_slide(prs, f"习题精讲 第 {i+1} 题")
        
        # 文本框位置
        has_img = len(q['matched_imgs']) > 0
        txt_width = Inches(8.0) if has_img else Inches(12.0)
        
        # 绘制题干卡片
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.2), txt_width, Inches(5.5))
        shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        shape.line.color.rgb = RGBColor(200, 200, 200)
        
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0.2)
        p = tf.paragraphs[0]
        p.text = q['text']
        p.font.size = Pt(18)
        p.font.color.rgb = RGBColor(40, 40, 40)
        set_font_all(p.runs[0])

        # 绘制图片
        if has_img:
            img_y = 1.2
            for img_p in q['matched_imgs'][:2]: # 最多放两张图，防止重叠
                try:
                    pic = slide.shapes.add_picture(img_p, Inches(8.8), Inches(img_y), width=Inches(4.0))
                    img_y += (pic.height / 914400) + 0.2
                except:
                    pass
                    
    prs.save(output_buffer)

# ==========================================
# 5. Streamlit UI 层
# ==========================================
st.set_page_config(page_title="物理习题 PPT 自动化", layout="wide")
st.title("⚛️ 物理习题教研课件自动生成器")

uploaded_files = st.file_uploader("上传图片或 PDF", accept_multiple_files=True, type=['png', 'jpg', 'pdf'])

if st.button("开始转换", type="primary"):
    if not uploaded_files:
        st.error("请先上传文件")
    else:
        all_qs = []
        with tempfile.TemporaryDirectory() as tmpdir:
            progress = st.progress(0)
            for idx, file in enumerate(uploaded_files):
                st.write(f"正在处理: {file.name}")
                if file.name.endswith('.pdf'):
                    doc = fitz.open(stream=file.read(), filetype="pdf")
                    for p_idx, page in enumerate(doc):
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img_path = os.path.join(tmpdir, f"tmp_{idx}_{p_idx}.png")
                        pix.save(img_path)
                        qs = smart_process(img_path, tmpdir, pdf_page=page)
                        all_qs.extend(qs)
                else:
                    img_path = os.path.join(tmpdir, file.name)
                    with open(img_path, "wb") as f: f.write(file.read())
                    qs = smart_process(img_path, tmpdir)
                    all_qs.extend(qs)
                progress.progress((idx + 1) / len(uploaded_files))
            
            if all_qs:
                buf = io.BytesIO()
                render_ppt(all_qs, buf)
                st.success(f"成功解析 {len(all_qs)} 道题目！")
                st.download_button("下载 PPT 课件", buf.getvalue(), "课件.pptx", "application/vnd.ms-powerpoint")
            else:
                st.error("未能识别出题目，请检查图片清晰度")
