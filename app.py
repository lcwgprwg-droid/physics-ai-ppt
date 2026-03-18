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
from docx import Document  # python-docx
from rapidocr_onnxruntime import RapidOCR
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from PIL import Image

# ==========================================
# 1. 核心引擎与初始化
# ==========================================
@st.cache_resource
def get_ocr_engine():
    """OCR 引擎常驻内存"""
    return RapidOCR()

def set_font_style(run, font_name='微软雅黑'):
    """统一设置中文字体"""
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    f = rPr.find(qn('w:rFonts'))
    if f is None:
        f = rPr.makeelement(qn('w:rFonts'))
        rPr.append(f)
    f.set(qn('w:eastAsia'), font_name)

# ==========================================
# 2. 视觉算法：OpenCV 高级图像分割
# ==========================================
def extract_images_advanced(img_path, ocr_result, out_dir, pdf_page=None):
    """
    针对物理受力图、电路图优化的分割算法
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 预处理：增强对比度并转灰度
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # 1. 精细化文字遮罩：保留物理标注（如 F, v, m, h）
    mask = np.zeros((h_img, w_img), dtype=np.uint8)
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            text = line[1].strip()
            # 物理特化：如果字符极短且不含汉字，认为是图上标注，不遮蔽
            if len(text) <= 2 and not re.search(r'[\u4e00-\u9fa5]', text):
                continue
            cv2.fillPoly(mask, [box], 255)

    # 2. Canny边缘检测：对细线条（绳子、斜面、光路）极其灵敏
    edges = cv2.Canny(gray, 30, 150)
    edges[mask > 0] = 0  # 移除题干文字干扰
    
    # 3. 动态形态学：核大小随图宽动态调整，连接断开的物理线段
    k_size = max(3, int(w_img / 180))
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (k_size, k_size))
    morphed = cv2.morphologyEx(edges, cv2.MORPH_CLOSE, kernel)
    morphed = cv2.dilate(morphed, kernel, iterations=2)

    # 4. 轮廓提取与面积/长宽比过滤
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    raw_boxes = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > 35 and h > 35 and (w * h) > 1800: # 过滤噪点
            if w < w_img * 0.9 and h < h_img * 0.9: # 过滤整页边框
                raw_boxes.append([x, y, x+w, y+h])

    # 5. 框合并：将距离较近的图示零件（如小车和斜面）合并成一张图
    merged = []
    if raw_boxes:
        raw_boxes.sort(key=lambda b: b[1])
        while raw_boxes:
            curr = raw_boxes.pop(0)
            combined = False
            for i in range(len(merged)):
                # 判断两框距离是否小于 60 像素
                if not (curr[2] < merged[i][0]-60 or curr[0] > merged[i][2]+60 or 
                        curr[3] < merged[i][1]-60 or curr[1] > merged[i][3]+60):
                    merged[i] = [min(curr[0], merged[i][0]), min(curr[1], merged[i][1]),
                                 max(curr[2], merged[i][2]), max(curr[3], merged[i][3])]
                    combined = True
                    break
            if not combined: merged.append(curr)

    # 6. 裁切与保存
    results = []
    for i, box in enumerate(merged):
        pad = 12
        x1, y1, x2, y2 = max(0, box[0]-pad), max(0, box[1]-pad), min(w_img, box[2]+pad), min(h_img, box[3]+pad)
        
        save_path = os.path.join(out_dir, f"fig_{int(time.time()*1000)}_{i}.png")
        if pdf_page: # PDF 导出高清图
            scale = h_img / pdf_page.rect.height
            pix = pdf_page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=fitz.Rect(x1/scale, y1/scale, x2/scale, y2/scale))
            pix.save(save_path)
        else: # 普通图片
            roi = img[y1:y2, x1:x2]
            cv2.imwrite(save_path, roi)
            
        results.append({"path": save_path, "center": ((x1+x2)/2, (y1+y2)/2), "col": 0 if (x1+x2)/2 < w_img/2 else 1})
    return results

# ==========================================
# 3. 逻辑层：智能图文绑定
# ==========================================
def bind_logic(questions, images, mid_x):
    """引力模型：图片自动吸附到上方/左方最近的题干"""
    for img in images:
        img_cx, img_cy = img['center']
        best_q = None
        min_dist = float('inf')
        
        for q in questions:
            # 规则：计算题目重心与图片中心的偏移，题目需在图片上方或同栏
            v_dist = img_cy - q['center'][1]
            if v_dist < -30: v_dist = abs(v_dist) * 3 # 惩罚：图片在题目上方不符合逻辑
            
            h_dist = abs(img_cx - q['center'][0])
            col_penalty = 0 if q['col'] == img['col'] else 1000
            
            dist = math.sqrt(h_dist**2 + v_dist**2) + col_penalty
            if dist < min_dist:
                min_dist = dist
                best_q = q
        
        if best_q: best_q['matched_imgs'].append(img['path'])

# ==========================================
# 4. Word 专项解析引擎
# ==========================================
def process_word(file_bytes, temp_dir):
    """
    解析 Word：提取文本段落并尝试提取内嵌图片
    """
    doc = Document(io.BytesIO(file_bytes))
    questions = []
    current_q = None
    q_pattern = r'^\s*\d+[\.．、\(（]'

    # 1. 提取所有段落
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text: continue
        
        if re.match(q_pattern, text):
            if current_q: questions.append(current_q)
            current_q = {"text": text, "matched_imgs": [], "center": (0, 0), "col": 0}
        else:
            if current_q: current_q["text"] += "\n" + text

    if current_q: questions.append(current_q)
    
    # 2. 提取 Word 中的图片并保存
    # 注意：Word 内部图片定位较难，这里采取顺序吸附策略
    img_idx = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_idx += 1
            img_path = os.path.join(temp_dir, f"word_img_{img_idx}.png")
            with open(img_path, "wb") as f:
                f.write(rel.target_part.blob)
            
            # 将 Word 图片尝试分配给最近生成的题目
            if questions:
                # 简单逻辑：按顺序分配给最后一个没图的题目，或平均分配
                target_q = questions[min(img_idx-1, len(questions)-1)]
                target_q['matched_imgs'].append(img_path)
                
    return questions

# ==========================================
# 5. PPT 渲染布局引擎
# ==========================================
def render_to_ppt(questions):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5) # 16:9
    
    c_blue = RGBColor(0, 112, 192)
    c_dark = RGBColor(40, 40, 45)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 1. 背景装饰
        rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        rect.fill.solid(); rect.fill.fore_color.rgb = RGBColor(248, 249, 251); rect.line.fill.background()
        
        # 2. 标题
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.6))
        p = title_box.text_frame.paragraphs[0]
        p.text = f"习题精讲 - 第 {i+1} 题"
        p.font.bold = True; p.font.size = Pt(26); p.font.color.rgb = c_blue
        set_font_style(p.runs[0])

        # 3. 题干文本卡片
        has_img = len(q['matched_imgs']) > 0
        txt_w = Inches(8.2) if has_img else Inches(12.0)
        
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.2), txt_w, Inches(5.8))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = RGBColor(220, 220, 225)
        
        tf = card.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        p = tf.paragraphs[0]
        p.text = q['text']
        p.font.size = Pt(18); p.font.color.rgb = c_dark
        set_font_style(p.runs[0])

        # 4. 图片排版
        if has_img:
            img_y = 1.2
            for img_p in q['matched_imgs'][:2]: # 每页最多放2张图
                try:
                    pic = slide.shapes.add_picture(img_p, Inches(8.9), Inches(img_y), width=Inches(4.0))
                    img_y += (pic.height / 914400) + 0.3
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 6. Streamlit UI 主程序
# ==========================================
st.set_page_config(page_title="物理教研 AI 助手", layout="centered", page_icon="🧪")

st.markdown("""
    <div style='text-align: center;'>
        <h1 style='color: #0070C0;'>🚀 物理习题教研 PPT 自动生成</h1>
        <p>支持 <b>JPG / PNG / PDF / Word</b>，高精度 OpenCV 识别物理图示</p>
    </div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📥 上传你的习题文档（支持多选）", 
                                  type=['png', 'jpg', 'jpeg', 'pdf', 'docx'], 
                                  accept_multiple_files=True)

if st.button("✨ 开始全自动转换", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("请上传文件后再开始转换！")
    else:
        all_questions = []
        progress = st.progress(0)
        status = st.empty()
        
        with tempfile.TemporaryDirectory() as tmpdir:
            engine = get_ocr_engine()
            
            for f_idx, file in enumerate(uploaded_files):
                status.info(f"正在处理: {file.name}...")
                ext = file.name.split('.')[-1].lower()
                
                # --- Word 分支 ---
                if ext == 'docx':
                    qs = process_word(file.read(), tmpdir)
                    all_questions.extend(qs)
                
                # --- PDF 分支 ---
                elif ext == 'pdf':
                    pdf = fitz.open(stream=file.read(), filetype="pdf")
                    for p_idx, page in enumerate(pdf):
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img_path = os.path.join(tmpdir, f"pdf_{f_idx}_{p_idx}.png")
                        pix.save(img_path)
                        
                        # OCR 与 视觉处理
                        res, _ = engine(img_path)
                        imgs = extract_images_advanced(img_path, res, tmpdir, pdf_page=page)
                        
                        # 识别题目
                        img_cv = cv2.imread(img_path)
                        mid_x = img_cv.shape[1] / 2
                        lines = []
                        for line in res:
                            box, text = line[0], line[1]
                            cx, cy = (box[0][0]+box[1][0])/2, (box[0][1]+box[2][1])/2
                            lines.append({"col": 0 if cx < mid_x else 1, "y": cy, "text": text, "center": (cx, cy)})
                        lines.sort(key=lambda x: (x['col'], x['y']))
                        
                        page_qs = []
                        cur_q = None
                        for l in lines:
                            if re.match(r'^\s*\d+[\.．、\(（]', l['text']):
                                if cur_q: page_qs.append(cur_q)
                                cur_q = {"text": l['text'], "matched_imgs": [], "center": l['center'], "col": l['col']}
                            elif cur_q: cur_q['text'] += "\n" + l['text']
                        if cur_q: page_qs.append(cur_q)
                        
                        bind_logic(page_qs, imgs, mid_x)
                        all_questions.extend(page_qs)
                
                # --- 图片分支 ---
                else:
                    img_path = os.path.join(tmpdir, file.name)
                    with open(img_path, "wb") as f: f.write(file.read())
                    res, _ = engine(img_path)
                    imgs = extract_images_advanced(img_path, res, tmpdir)
                    img_cv = cv2.imread(img_path)
                    mid_x = img_cv.shape[1] / 2
                    
                    lines = []
                    for line in res:
                        box, text = line[0], line[1]
                        cx, cy = (box[0][0]+box[1][0])/2, (box[0][1]+box[2][1])/2
                        lines.append({"col": 0 if cx < mid_x else 1, "y": cy, "text": text, "center": (cx, cy)})
                    lines.sort(key=lambda x: (x['col'], x['y']))
                    
                    page_qs = []
                    cur_q = None
                    for l in lines:
                        if re.match(r'^\s*\d+[\.．、\(（]', l['text']):
                            if cur_q: page_qs.append(cur_q)
                            cur_q = {"text": l['text'], "matched_imgs": [], "center": l['center'], "col": l['col']}
                        elif cur_q: cur_q['text'] += "\n" + l['text']
                    if cur_q: page_qs.append(cur_q)
                    
                    bind_logic(page_qs, imgs, mid_x)
                    all_questions.extend(page_qs)

                progress.progress((f_idx + 1) / len(uploaded_files))

            if all_questions:
                status.success(f"✅ 转换完成！共生成 {len(all_questions)} 道题目 PPT。")
                ppt_data = render_to_ppt(all_questions)
                st.download_button(label="📥 点击下载生成好的 PPT 课件", 
                                   data=ppt_data, 
                                   file_name="AI物理教研精讲课件.pptx", 
                                   mime="application/vnd.ms-powerpoint",
                                   use_container_width=True)
            else:
                status.error("未能识别到题目，请确保上传的文件内容包含清晰的题号。")
