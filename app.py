import os
import re
import cv2
import numpy as np
import tempfile
import io
import time
import gc
import streamlit as st
import fitz  # PyMuPDF
from rapidocr_onnxruntime import RapidOCR
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from PIL import Image, ImageEnhance
from docx import Document

# ==========================================
# 1. 核心引擎：保持单例
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：精准避开文字框，只抓物理图
# ==========================================
def extract_physics_diagrams(img_path, ocr_result, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 建立文字区域索引，用于“反向排除”
    text_boxes = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            text_boxes.append(box)

    # 预处理
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 使用自适应阈值处理可能的阴影
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
    
    # 膨胀操作：把物理图的线条连起来
    kernel = np.ones((5, 5), np.uint8)
    dilated = cv2.dilate(thresh, kernel, iterations=1)

    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    diagrams = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 核心逻辑：如果这个框太宽或太高（可能是题目的边框），或者里面包含大量文字，就跳过
        if w > w_img * 0.8 or h > h_img * 0.8: continue
        if w < 40 or h < 40: continue
        
        # 检查该矩形内是否包含 OCR 文字块
        contains_much_text = 0
        for tb in text_boxes:
            tx_min, ty_min = np.min(tb, axis=0)
            tx_max, ty_max = np.max(tb, axis=0)
            # 如果文字块在当前矩形内部
            if tx_min >= x and tx_max <= x+w and ty_min >= y and ty_max <= y+h:
                contains_much_text += 1
        
        # 如果矩形内文字超过 3 块，通常是题干本身，不是插图
        if contains_much_text > 3: continue
        
        # 保存真正的物理图示
        roi = img[y:y+h, x:x+w]
        p = os.path.join(out_dir, f"fig_{int(time.time()*1000)}.png")
        cv2.imwrite(p, roi)
        diagrams.append({"path": p, "y": y + h/2, "x": x + w/2})
        
    return diagrams

# ==========================================
# 3. PPT 渲染引擎：【还原你原有的精美设计】
# ==========================================
def set_font(run, font_name='微软雅黑'):
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    f = rPr.find(qn('w:rFonts'))
    if f is None:
        f = rPr.makeelement(qn('w:rFonts'))
        rPr.append(f)
    f.set(qn('w:eastAsia'), font_name)

def create_base_slide(prs, title_text):
    """还原：蓝色装饰条 + 灰色背景"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
    
    # 蓝色侧边条
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11.5), Inches(0.8))
    p = tb.text_frame.paragraphs[0]; p.text = title_text
    p.font.bold = True; p.font.size = Pt(26); p.font.color.rgb = RGBColor(30, 40, 60)
    set_font(p.runs[0])
    return slide

def add_badge_card(slide, x, y, w, h, badge_text, badge_color, content, font_size):
    """还原：白色圆角卡片 + 悬浮标签"""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(230, 230, 235)
    
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.2), y - Inches(0.15), Inches(1.2), Inches(0.35))
    badge.fill.solid(); badge.fill.fore_color.rgb = badge_color; badge.line.fill.background()
    bp = badge.text_frame.paragraphs[0]; bp.text = badge_text; bp.font.bold = True; bp.font.size = Pt(12)
    bp.font.color.rgb = RGBColor(255, 255, 255); bp.alignment = PP_ALIGN.CENTER; set_font(bp.runs[0])
    
    tb = slide.shapes.add_textbox(x + Inches(0.2), y + Inches(0.4), w - Inches(0.4), h - Inches(0.6))
    tf = tb.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    lines = content.split('\n')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line; p.font.size = Pt(font_size); p.font.color.rgb = RGBColor(30, 30, 30); set_font(p.runs[0])

# ==========================================
# 4. 业务逻辑层
# ==========================================
def process_to_final_ppt(uploaded_files, temp_dir):
    engine = get_ocr_engine()
    all_qs = []

    for file in uploaded_files:
        ext = file.name.split('.')[-1].lower()
        if ext in ['jpg', 'png', 'jpeg']:
            path = os.path.join(temp_dir, file.name)
            with open(path, "wb") as f: f.write(file.read())
            
            # OCR 识别
            res, _ = engine(path)
            # 物理图提取（避开文字框逻辑）
            diagrams = extract_physics_diagrams(path, res, temp_dir)
            
            # 题干整合
            current_q = None
            for line in res:
                text = line[1].strip()
                # 识别题号
                if re.match(r'^\d+[\.．、]', text):
                    if current_q: all_qs.append(current_q)
                    current_q = {"text": text, "imgs": [], "y": line[0][0][1]}
                elif current_q:
                    current_q["text"] += "\n" + text
            if current_q: all_qs.append(current_q)
            
            # 关联图示（简单就近原则）
            for d in diagrams:
                if all_qs:
                    # 找到垂直距离最近的题目
                    best_q = min(all_qs, key=lambda q: abs(q['y'] - d['y']))
                    best_q['imgs'].append(d['path'])

    # 开始渲染 PPT
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    c_blue, c_orange = RGBColor(0, 112, 192), RGBColor(230, 90, 40)

    for i, q in enumerate(all_qs):
        slide = create_base_slide(prs, f"习题精讲 第 {i+1} 题")
        has_img = len(q['imgs']) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.0)
        
        # 原题呈现卡片
        add_badge_card(slide, Inches(0.5), Inches(1.3), txt_w, Inches(3.5), "原题呈现", c_blue, q['text'], 18)
        # 解析预留卡片
        add_badge_card(slide, Inches(0.5), Inches(5.2), txt_w, Inches(1.8), "思路分析", c_orange, "此处进行受力分析与列式讲解...", 16)
        
        # 渲染插图
        if has_img:
            img_y = 1.3
            for img_path in q['imgs'][:2]: # 最多两张
                pic = slide.shapes.add_picture(img_path, Inches(9.2), Inches(img_y), width=Inches(3.8))
                img_y += (pic.height / 914400) + 0.2

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 5. Streamlit UI
# ==========================================
st.set_page_config(page_title="物理题自动生成PPT", layout="wide")
st.title("⚛️ 物理教研自动化工作站 (版本修复版)")

files = st.file_uploader("上传真题图片", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])

if st.button("生成精美 PPT 课件", type="primary"):
    if files:
        with tempfile.TemporaryDirectory() as tmp:
            ppt_data = process_to_final_ppt(files, tmp)
            st.download_button("📥 下载 PPT 课件", ppt_data, "物理课件.pptx", use_container_width=True)
    else:
        st.warning("请先上传图片！")
