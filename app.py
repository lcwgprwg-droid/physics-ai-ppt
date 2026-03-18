import os
import re
import cv2
import math
import tempfile
import io
import time
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
from PIL import Image, ImageEnhance, ImageOps

# ==========================================
# 1. 核心引擎：图像处理与增强（针对物理符号优化）
# ==========================================
def preprocess_for_ocr(img_path):
    """
    针对物理符号和下标进行图像增强：
    1. 放大图片 2. 增强对比度 3. 锐化
    """
    with Image.open(img_path) as img:
        # 转为 RGB 并放大 2 倍 (关键：让下标变大)
        img = img.convert("RGB")
        w, h = img.size
        img = img.resize((w * 2, h * 2), Image.Resampling.LANCZOS)
        
        # 增强对比度：让变浅的变量符号变黑
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.0)
        
        # 锐化
        enhancer = ImageEnhance.Sharpness(img)
        img = enhancer.enhance(2.0)
        
        enhanced_path = img_path.replace(".", "_enhanced.")
        img.save(enhanced_path, quality=95)
        return enhanced_path

def crop_diagrams(img_path, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 6))
    dilated = cv2.dilate(thresh, kernel, iterations=2)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    diagrams = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if h > 40 and w > 40 and (w * h) > 5000:
            roi = img[y:y+h, x:x+w]
            p = os.path.join(out_dir, f"fig_{int(time.time()*1000)}.png")
            cv2.imwrite(p, roi)
            diagrams.append({"path": p, "x_center": x + w/2, "y_center": y + h/2})
    return diagrams

# ==========================================
# 2. 智能 OCR（物理变量保护版）
# ==========================================
def smart_ocr_and_split(img_path, cv_images):
    # 预处理图片：放大并增强
    enhanced_img_path = preprocess_for_ocr(img_path)
    
    # 调低阈值以捕获细小下标
    engine = RapidOCR()
    result, _ = engine(enhanced_img_path)
    if not result: return []

    # 注意：因为图片放大了2倍，这里的坐标需要除以2回到原图坐标
    processed_result = []
    for line in result:
        box, text, score = line[0], line[1], line[2]
        # 缩放坐标还原
        box = [[p[0]/2, p[1]/2] for p in box]
        processed_result.append([box, text, score])

    # 判断单双栏
    all_x = [line[0][0][0] for line in processed_result]
    page_center_x = (max(all_x) + min(all_x)) / 2 if all_x else 500

    sorted_lines = []
    for line in processed_result:
        box, text = line[0], line[1]
        cx = (box[0][0] + box[1][0]) / 2
        cy = (box[0][1] + box[3][1]) / 2
        col_idx = 0 if cx < page_center_x else 1
        sorted_lines.append({"col": col_idx, "y": cy, "box": box, "text": text})

    sorted_lines.sort(key=lambda item: (item['col'], item['y']))
    
    questions, current_q = [], None
    for item in sorted_lines:
        text = item['text'].strip()
        
        # 物理题号识别逻辑
        if re.match(r'^\s*\d+[\.．、\)]', text):
            if current_q: questions.append(current_q)
            current_q = {'text': text, 'y_min': item['box'][0][1], 'y_max': item['box'][2][1], 'matched_imgs': []}
        elif current_q:
            # 物理公式补丁：尝试修复下标连接
            if re.match(r'^[a-zA-Z0-9]$', text): # 如果是孤立的字母或数字，通常是下标
                current_q['text'] += text
            else:
                current_q['text'] += " " + text
            current_q['y_min'] = min(current_q['y_min'], item['box'][0][1])
            current_q['y_max'] = max(current_q['y_max'], item['box'][2][1])

    if current_q: questions.append(current_q)

    # 图像关联
    for img in cv_images:
        best_q, min_dist = None, float('inf')
        for q in questions:
            dist = abs(img['y_center'] - (q['y_min'] + q['y_max'])/2)
            if dist < min_dist: min_dist, best_q = dist, q
        if best_q and min_dist < 400: best_q['matched_imgs'].append(img['path'])
            
    return questions

# ==========================================
# 3. 文档处理路由
# ==========================================
def process_uploaded_files(uploaded_files, temp_dir):
    all_questions = []
    for file in uploaded_files:
        file_bytes = file.read()
        file_ext = file.name.split('.')[-1].lower()

        if file_ext in ['jpg', 'jpeg', 'png']:
            path = os.path.join(temp_dir, file.name)
            with open(path, "wb") as f: f.write(file_bytes)
            all_questions.extend(smart_ocr_and_split(path, crop_diagrams(path, temp_dir)))

        elif file_ext == 'pdf':
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for i in range(len(doc)):
                pix = doc[i].get_pixmap(matrix=fitz.Matrix(2, 2))
                img_path = os.path.join(temp_dir, f"p_{i}.jpg")
                pix.save(img_path)
                all_questions.extend(smart_ocr_and_split(img_path, crop_diagrams(img_path, temp_dir)))

        elif file_ext == 'docx':
            doc = docx.Document(io.BytesIO(file_bytes))
            current_q = None
            for para in doc.paragraphs:
                text = para.text.strip()
                if re.match(r'^\s*\d+[\.．、]', text):
                    if current_q: all_questions.append(current_q)
                    current_q = {'text': text, 'matched_imgs': []}
                elif current_q and text:
                    current_q['text'] += "\n" + text
                if current_q:
                    for run in para.runs:
                        if 'pic:pic' in run._element.xml:
                            rIds = re.findall(r'r:embed="([^"]+)"', run._element.xml)
                            for rId in rIds:
                                try:
                                    img_part = doc.part.related_parts[rId]
                                    img_path = os.path.join(temp_dir, f"w_img_{rId}.png")
                                    with open(img_path, "wb") as f: f.write(img_part.blob)
                                    current_q['matched_imgs'].append(img_path)
                                except: pass
            if current_q: all_questions.append(current_q)
    return all_questions

# ==========================================
# 4. PPT 渲染
# ==========================================
def set_font(run, name='微软雅黑'):
    run.font.name = name
    r = run._r.get_or_add_rPr().find(qn('w:rFonts'))
    if r is None: r = run._r.get_or_add_rPr().makeelement(qn('w:rFonts'))
    r.set(qn('w:eastAsia'), name)

def create_slide(prs, q_text, imgs, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(250, 252, 255); bg.line.fill.background()
    
    # 顶部装饰
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.3), Inches(0.3), Inches(0.1), Inches(0.5))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    tb_title = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(10), Inches(0.6))
    p_title = tb_title.text_frame.paragraphs[0]
    p_title.text = f"核心素养习题精讲 - 题目 {idx}"; p_title.font.size = Pt(24); p_title.font.bold = True; set_font(p_title.runs[0])

    has_img = len(imgs) > 0
    box_w = Inches(8.8) if has_img else Inches(12.5)
    
    # 题目文本框
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.2), box_w, Inches(5.8))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(200, 200, 210)
    
    tf = slide.shapes.add_textbox(Inches(0.6), Inches(1.4), box_w - Inches(0.4), Inches(5.4)).text_frame
    tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]; p.text = q_text; p.font.size = Pt(20); p.line_spacing = 1.3; set_font(p.runs[0])

    if has_img:
        for i, img_p in enumerate(imgs[:2]):
            try: slide.shapes.add_picture(img_p, Inches(9.5), Inches(1.2 + i*3.1), width=Inches(3.5))
            except: pass

def make_master_ppt(questions):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    for i, q in enumerate(questions):
        create_slide(prs, q['text'], q['matched_imgs'], i+1)
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 5. Streamlit UI
# ==========================================
st.set_page_config(page_title="AI 物理教研课件生成器", layout="centered", page_icon="⚛️")

st.markdown("<h1 style='text-align: center; color: #0070C0;'>🚀 AI 物理教研全自动工作站</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #666;'>针对物理公式、下标、符号进行了 OCR 深度优化</p>", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📥 上传 物理试卷/教辅 (图片/PDF/Word)", accept_multiple_files=True, type=['jpg', 'png', 'pdf', 'docx'])

if st.button("✨ 开始生成巅峰排版 PPT", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("请先上传文件")
    else:
        with st.spinner("正在进行超分辨率增强与 OCR 识别..."):
            with tempfile.TemporaryDirectory() as temp_dir:
                questions = process_uploaded_files(uploaded_files, temp_dir)
                if questions:
                    ppt_io = make_master_ppt(questions)
                    st.session_state['ready_ppt'] = ppt_io.getvalue()
                    st.success(f"成功识别 {len(questions)} 道题目！")
                    st.balloons()
                else:
                    st.error("未识别到题目，请尝试提高图片清晰度")

if 'ready_ppt' in st.session_state:
    st.download_button(
        label="⬇️ 点击下载 PPT 课件",
        data=st.session_state['ready_ppt'],
        file_name="AI物理教研精选课件.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
