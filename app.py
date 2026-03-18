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
# 1. 核心引擎与公式渲染
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

def render_math_formula(latex_text, save_path):
    """将物理公式转换为高清透明图片，防止OCR乱码或丢失"""
    try:
        if not latex_text: return False
        # 清洗：确保是标准数学格式
        latex_text = latex_text.replace('\n', '').strip()
        display_text = f"${latex_text}$" if not latex_text.startswith('$') else latex_text
        
        fig = plt.figure(figsize=(4, 0.8), dpi=300)
        plt.axis('off')
        plt.text(0.5, 0.5, display_text, size=24, ha='center', va='center', color='#1E2850')
        plt.savefig(save_path, format='png', transparent=True, bbox_inches='tight', pad_inches=0.05)
        plt.close(fig)
        return True
    except:
        return False

# ==========================================
# 2. 视觉算法：基于坐标语义的元素提取 (重构点)
# ==========================================
def extract_visual_elements_v7(img_path, ocr_result, out_dir):
    """
    重构：不再遮盖文字，而是通过重叠面积比例判定
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 获取 OCR 文字区域的坐标字典
    text_rects = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            text_rects.append([x_min, y_min, x_max, y_max])

    # 预处理：保护细线条
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 21, 10)
    
    # 膨胀：将断开的物理图示（如活塞、气缸）连在一起
    kernel = np.ones((15, 15), np.uint8)
    morphed = cv2.dilate(binary, kernel, iterations=1)
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.9 or h > h_img * 0.9: continue # 过滤大外框
        if w < 25 or h < 25: continue # 过滤微小噪点

        # 核心逻辑：计算该轮廓与所有 OCR 文本框的重叠率
        is_text_component = False
        for tr in text_rects:
            # 计算相交矩形
            ix1, iy1 = max(x, tr[0]), max(y, tr[1])
            ix2, iy2 = min(x+w, tr[2]), min(y+h, tr[3])
            if ix1 < ix2 and iy1 < iy2:
                overlap_area = (ix2 - ix1) * (iy2 - iy1)
                # 如果 70% 的面积被判定为已知文字，则不将其视为插图提取
                if overlap_area / (w * h) > 0.7:
                    is_text_component = True
                    break
        
        # 如果不是纯文字，或者是极大的公式块，则提取为图片
        if not is_text_component or (w * h > 8000):
            roi = img[y:y+h, x:x+w]
            # 过滤纯白块
            if np.mean(roi) > 250: continue
            
            f_path = os.path.join(out_dir, f"diag_{int(time.time()*1000)}_{x}.png")
            cv2.imwrite(f_path, roi)
            elements.append({"path": f_path, "y": y + h/2})
            
    return elements

# ==========================================
# 3. PPT 渲染引擎：锁定版式与强制左对齐
# ==========================================
def set_font_style(run, size=18, color=(40, 40, 40), is_bold=False):
    run.font.name = '微软雅黑'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(*color)
    rPr = run._r.get_or_add_rPr()
    f = rPr.find(qn('w:rFonts'))
    if f is None:
        f = rPr.makeelement(qn('w:rFonts'))
        rPr.append(f)
    f.set(qn('w:eastAsia'), '微软雅黑')

def render_physics_ppt(questions, tmp_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    c_blue = RGBColor(0, 112, 192)
    c_orange = RGBColor(230, 90, 40)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 侧边条
        slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55)).fill.solid().fore_color.rgb = c_blue
        
        # 标题 (锁定左对齐)
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(11), Inches(0.8))
        tp = title_box.text_frame.paragraphs[0]; tp.alignment = PP_ALIGN.LEFT
        set_font_style(tp.add_run(), size=26, is_bold=True, color=(20, 40, 80)).text = f"习题精讲 第 {i+1} 题"
        
        has_img = len(q['imgs']) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.2)

        # 原题卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.2))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 220, 225)
        
        # 填充文本 (强制左对齐)
        tf = card.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        for idx, line in enumerate(q['text'].split('\n')):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT; p.line_spacing = 1.2
            set_font_style(p.add_run()).text = line.strip()

        # 物理公式高清补偿 (自动捕获常见物理符号)
        formulas = re.findall(r'([A-Za-z]_\d\s*=\s*\d+\s*[K|m|s|Ω|V|A|J])', q['text'])
        if formulas:
            curr_fy = 4.6
            for fm in formulas[:2]:
                f_path = os.path.join(tmp_dir, f"fm_{time.time()}.png")
                if render_math_formula(fm, f_path):
                    slide.shapes.add_picture(f_path, Inches(0.8), Inches(curr_fy), height=Inches(0.5))
                    curr_fy += 0.6

        # 投放插图与公式切图
        if has_img:
            y_pos = 1.3
            q['imgs'].sort(key=lambda x: x['y'])
            for img_info in q['imgs'][:3]:
                pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_pos), width=Inches(3.8))
                y_pos += (pic.height / 914400) + 0.2
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. 业务控制逻辑
# ==========================================
st.set_page_config(page_title="物理教研 AI 工具", layout="centered")
st.title("⚛️ 物理题自动生成 PPT (最终版)")

uploaded_files = st.file_uploader("支持 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 生成专业级课件", type="primary", use_container_width=True):
    if uploaded_files:
        all_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in uploaded_files:
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
                    # 图片或 PDF 视觉分支
                    img_paths = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5))
                            p_path = os.path.join(tmp, f"p_{time.time()}.png")
                            pix.save(p_path); img_paths.append(p_path)
                    else:
                        p_path = os.path.join(tmp, file.name)
                        with open(p_path, "wb") as f: f.write(file.read())
                        img_paths.append(p_path)
                    
                    for p_path in img_paths:
                        res, _ = engine(p_path)
                        visuals = extract_visual_elements_v7(p_path, res, tmp)
                        
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
                ppt_data = render_physics_ppt(all_qs, tmp)
                st.download_button("📥 下载生成好的 PPT 课件", ppt_data, "物理 AI 精选课件.pptx", use_container_width=True)
