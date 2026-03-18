import os
import re
import cv2
import numpy as np
import tempfile
import io
import time
import streamlit as st
import matplotlib.pyplot as plt
import fitz  # PyMuPDF
from docx import Document
from rapidocr_onnxruntime import RapidOCR
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

# ==========================================
# 1. 核心引擎 (RapidOCR)
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：语义分割 (锁定物理线条与独立公式)
# ==========================================
def extract_visual_elements_v8(img_path, ocr_result, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 记录所有文本块坐标
    text_boxes = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            text_boxes.append([x_min, y_min, x_max, y_max])

    # 预处理：增强对比度，专门针对物理题细线
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 中值滤波去除杂点，保护线条
    blurred = cv2.medianBlur(gray, 3)
    binary = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 21, 10)
    
    # 形态学膨胀：把公式字母和物理箭头连成整体块
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (15, 10))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 1. 过滤：排除全页大框 (如截图中那个圆角大矩形)
        if w > w_img * 0.8 or h > h_img * 0.8: continue
        if w < 20 or h < 15: continue 

        # 2. 判定：计算该块与OCR文本的重叠度
        is_pure_text = False
        for tb in text_boxes:
            ix1, iy1 = max(x, tb[0]), max(y, tb[1])
            ix2, iy2 = min(x+w, tb[2]), min(y+h, tb[3])
            if ix1 < ix2 and iy1 < iy2:
                overlap_area = (ix2 - ix1) * (iy2 - iy1)
                # 如果 80% 区域被文本覆盖，判定为普通正文
                if overlap_area / (w * h) > 0.8:
                    is_pure_text = True
                    break
        
        # 3. 物理特化逻辑：如果是公式(独立块)或插图，则提取
        if not is_pure_text or (w * h > 5000):
            roi = img[y:y+h, x:x+w]
            # 过滤几乎空白的区域
            if np.mean(cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)) > 252: continue
            
            f_path = os.path.join(out_dir, f"ele_{int(time.time()*1000)}_{x}.png")
            cv2.imwrite(f_path, roi)
            elements.append({"path": f_path, "y": y + h/2})
            
    return elements

# ==========================================
# 3. PPT 渲染：【修正 AttributeError】
# ==========================================
def set_font_pptx(run, size=18, is_bold=False, color=(40, 40, 40)):
    """
    底层操作 DrawingML 命名空间，修复字体报错
    """
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(*color)
    
    # 强制设置字体（针对 Windows/Office 兼容性）
    run.font.name = '微软雅黑'
    rPr = run._r.get_or_add_rPr()
    # PPT 的中文字体设置在 a:ea (East Asian) 属性中
    rPr.set(qn('a:latin'), '微软雅黑')
    rPr.set(qn('a:ea'), '微软雅黑')

def render_ppt_final(questions, tmp_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    blue = RGBColor(0, 112, 192)
    orange = RGBColor(230, 90, 40)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景底色
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 蓝色侧边指示条
        slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55)).fill.solid().fore_color.rgb = blue
        
        # 标题栏 - 显式左对齐
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        tf_title = title_box.text_frame
        tf_title.paragraphs[0].alignment = PP_ALIGN.LEFT
        run_title = tf_title.paragraphs[0].add_run()
        run_title.text = f"习题精讲 第 {i+1} 题"
        set_font_pptx(run_title, size=26, is_bold=True, color=(20, 40, 80))
        
        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.2)

        # 题干卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.2))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 220, 225)
        
        # 填入文字 - 强制左对齐
        tf = card.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        lines = q['text'].split('\n')
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.2
            r = p.add_run()
            r.text = line.strip()
            set_font_pptx(r, size=18)

        # 思路分析卡片
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.6), txt_w, Inches(1.5))
        card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(220, 220, 225)
        
        tf2 = card2.text_frame
        p2 = tf2.paragraphs[0]
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run(); r2.text = "待补充详细受力分析与列式过程..."; set_font_pptx(r2, size=16, color=(100, 100, 100))

        # 图片投放 (物理图示 + 公式图片)
        if has_img:
            y_cursor = 1.3
            q['imgs'].sort(key=lambda x: x['y'])
            for img_info in q['imgs'][:3]: # 限制 3 张，防止溢出
                try:
                    pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_cursor), width=Inches(3.8))
                    y_cursor += (pic.height / 914400) + 0.2
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. Streamlit 界面
# ==========================================
st.set_page_config(page_title="物理教研课件专家", layout="centered")
st.title("⚛️ 物理题自动解析 (视觉&排版修正版)")

files = st.file_uploader("支持 Word/PDF/图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🔥 立即生成 PPT", type="primary", use_container_width=True):
    if files:
        all_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    cur_q = None
                    for p in doc.paragraphs:
                        t = p.text.strip()
                        if not t: continue
                        if re.match(r'^\d+[\.．、]', t):
                            if cur_q: all_qs.append(cur_q)
                            cur_q = {"text": t, "imgs": []}
                        elif cur_q: cur_q['text'] += "\n" + t
                    if cur_q: all_qs.append(cur_q)
                
                else:
                    paths = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5))
                            p_path = os.path.join(tmp, f"p_{time.time()}.png")
                            pix.save(p_path); paths.append(p_path)
                    else:
                        p_path = os.path.join(tmp, file.name)
                        with open(p_path, "wb") as f: f.write(file.read())
                        paths.append(p_path)
                    
                    for p_p in paths:
                        res, _ = engine(p_p)
                        visuals = extract_visual_elements_v8(p_p, res, tmp)
                        
                        page_qs = []
                        cur_q = None
                        if res:
                            for line in res:
                                txt = line[1].strip()
                                if re.match(r'^\d+[\.．、]', txt):
                                    if cur_q: page_qs.append(cur_q)
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                                elif cur_q: cur_q['text'] += "\n" + txt
                            if cur_q: page_qs.append(cur_q)
                        
                        for v in visuals:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - v['y']))
                                target['imgs'].append(v)
                        all_qs.extend(page_qs)

            if all_qs:
                ppt_data = render_ppt_final(all_qs, tmp)
                st.download_button("📥 下载生成好的 PPT 课件", ppt_data, "物理课件_修正版.pptx", use_container_width=True)
                st.success("课件生成成功！")
