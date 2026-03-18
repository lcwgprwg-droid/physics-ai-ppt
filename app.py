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
# 1. 核心配置：LaTeX 渲染与 OCR
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

def render_latex_to_png(latex_str, out_path):
    """
    将物理公式渲染为高清透明 PNG
    """
    try:
        # 物理公式清洗
        latex_str = latex_str.replace('\n', '').strip()
        if not (latex_str.startswith('$') and latex_str.endswith('$')):
            latex_str = f"${latex_str}$"
        
        fig = plt.figure(figsize=(3, 0.6), dpi=300)
        plt.axis('off')
        plt.text(0.5, 0.5, latex_str, size=22, ha='center', va='center', color='#1E2850')
        plt.savefig(out_path, format='png', transparent=True, bbox_inches='tight', pad_inches=0.05)
        plt.close(fig)
        return True
    except:
        return False

# ==========================================
# 2. 视觉算法重构：非破坏性物理元素提取
# ==========================================
def extract_visual_elements(img_path, ocr_result, out_dir):
    """
    改进算法：不再遮盖文字，而是通过坐标对比排除文字
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 1. 获取 OCR 文本框坐标集合
    text_boxes = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            text_boxes.append([x_min, y_min, x_max, y_max])

    # 2. 增强线条检测
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 专门捕捉物理细线条
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 21, 10)
    
    # 形态学闭合：将断开的受力箭头连起来
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (15, 15))
    morphed = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.9 or h > h_img * 0.9: continue # 剔除外边框
        if w < 25 or h < 25: continue # 允许抓取孤立字母

        # 计算重叠率：如果该区域被 OCR 文本框高度覆盖，则判定为文字，不切图
        is_text = False
        for tb in text_boxes:
            overlap_x1, overlap_y1 = max(x, tb[0]), max(y, tb[1])
            overlap_x2, overlap_y2 = min(x+w, tb[2]), min(y+h, tb[3])
            if overlap_x1 < overlap_x2 and overlap_y1 < overlap_y2:
                overlap_area = (overlap_x2 - overlap_x1) * (overlap_y2 - overlap_y1)
                if overlap_area / (w * h) > 0.7: # 70% 都是文字
                    is_text = True
                    break
        
        if not is_text or (w * h > 8000): # 物理插图或巨型公式
            roi = img[y:y+h, x:x+w]
            f_path = os.path.join(out_dir, f"ele_{int(time.time()*1000)}_{x}.png")
            cv2.imwrite(f_path, roi)
            elements.append({"path": f_path, "y": y + h/2})
            
    return elements

# ==========================================
# 3. PPT 渲染引擎：严格左对齐
# ==========================================
def create_slide_industrial(prs, q_data, title_text, tmp_dir):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 还原老师的审美：灰色底+蓝色条
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    
    # 标题 (左对齐)
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(10), Inches(0.8))
    tp = title_box.text_frame.paragraphs[0]
    tp.alignment = PP_ALIGN.LEFT
    tr = tp.add_run(); tr.text = title_text
    tr.font.name = '微软雅黑'; tr.font.size = Pt(26); tr.font.bold = True
    
    has_img = len(q_data['imgs']) > 0
    txt_w = Inches(8.5) if has_img else Inches(12.2)

    # 原题卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.2))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 220, 225)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    # 填充正文 (逐行左对齐)
    lines = q_data['text'].split('\n')
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.2
        r = p.add_run(); r.text = line.strip()
        r.font.name = '微软雅黑'; r.font.size = Pt(18)
        # 强制东亚字体
        rPr = r._r.get_or_add_rPr()
        f = rPr.get_or_add_rFonts()
        f.set(qn('w:eastAsia'), '微软雅黑')

    # 公式渲染检测 (物理公式补强)
    formula_matches = re.findall(r'([A-Za-z]_\d\s*=\s*\d+\s*[K|m|s|Ω|V|A])', q_data['text'])
    if formula_matches:
        f_y = 4.5
        for fm in formula_matches[:2]:
            f_p = os.path.join(tmp_dir, f"fm_{time.time()}.png")
            if render_latex_to_png(fm, f_p):
                slide.shapes.add_picture(f_p, Inches(0.8), Inches(f_y), height=Inches(0.5))
                f_y += 0.6

    # 投放物理图
    if has_img:
        y_ptr = 1.3
        q_data['imgs'].sort(key=lambda x: x['y'])
        for img_info in q_data['imgs'][:3]:
            pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
            y_ptr += (pic.height / 914400) + 0.2

# ==========================================
# 4. Streamlit UI 业务流
# ==========================================
st.set_page_config(page_title="物理教研课件专家", layout="centered")
st.title("⚛️ 物理习题 AI：高清公式版")

uploaded_files = st.file_uploader("支持 Word/PDF/图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🔥 立即生成专业 PPT", type="primary", use_container_width=True):
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
                    # PDF / Image
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
                        visuals = extract_visual_elements(p_p, res, tmp)
                        
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
                prs = Presentation()
                prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
                for i, q in enumerate(all_qs):
                    create_slide_industrial(prs, q, f"习题精讲 第 {i+1} 题", tmp)
                
                buf = io.BytesIO()
                prs.save(buf)
                st.download_button("📥 下载课件", buf.getvalue(), "物理专业课件.pptx", use_container_width=True)
