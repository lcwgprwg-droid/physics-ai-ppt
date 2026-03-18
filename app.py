import os
import re
import cv2
import numpy as np
import tempfile
import io
import time
import streamlit as st
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
# 2. 视觉重构：精准区分【物理图、公式、正文】
# ==========================================
def extract_elements_v5(img_path, ocr_result, out_dir):
    """
    重构逻辑：
    1. 不再涂白，保持原图完整
    2. 物理图：识别为独立大面积区域
    3. 公式：识别为带有数学特征的孤立区域
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 预处理：增强线条对比度
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 使用较小的阈值，保护物理细线条
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 15, 5)
    
    # 获取所有文字块的矩形
    text_boxes = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            text_boxes.append([x_min, y_min, x_max, y_max])

    # 寻找轮廓
    kernel = np.ones((5, 5), np.uint8)
    morphed = cv2.dilate(thresh, kernel, iterations=1)
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    saved_elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 尺寸过滤
        if w > w_img * 0.9 or h > h_img * 0.9: continue 
        if w < 25 or h < 25: continue # 允许更小的字母/符号被捕获

        # 计算该轮廓与所有 OCR 文本框的交集
        is_pure_text = False
        for tb in text_boxes:
            # 计算重叠面积
            ix1, iy1 = max(x, tb[0]), max(y, tb[1])
            ix2, iy2 = min(x+w, tb[2]), min(y+h, tb[3])
            if ix1 < ix2 and iy1 < iy2:
                overlap_area = (ix2 - ix1) * (iy2 - iy1)
                # 如果 80% 的轮廓都被文字框覆盖，说明这就是普通正文，不用切图
                if overlap_area / (w * h) > 0.8:
                    is_pure_text = True
                    break
        
        # 物理题特有：如果这个块很大，或者是孤立的（如 $T_1=300K$），即便它是字，我们也切下来
        # 这样可以防止公式乱码，作为图片展示最稳
        if not is_pure_text or (w * h > 5000): # 面积较大的公式或插图
            roi = img[y:y+h, x:x+w]
            f_path = os.path.join(out_dir, f"element_{int(time.time()*1000)}_{x}.png")
            cv2.imwrite(f_path, roi)
            saved_elements.append({"path": f_path, "y": y + h/2, "x": x + w/2})
        
    return saved_elements

# ==========================================
# 3. PPT 渲染引擎（强制左对齐 + 高清画质）
# ==========================================
def set_text_style(paragraph, size=18, color=(30, 30, 30), is_title=False):
    paragraph.alignment = PP_ALIGN.LEFT # 核心修复：强制左对齐
    paragraph.line_spacing = 1.2
    run = paragraph.add_run()
    run.font.name = '微软雅黑'
    run.font.size = Pt(size if not is_title else 24)
    run.font.bold = is_title
    run.font.color.rgb = RGBColor(*color)
    
    # 修复东亚字体显示
    rPr = run._r.get_or_add_rPr()
    f = rPr.find(qn('w:rFonts'))
    if f is None:
        f = rPr.makeelement(qn('w:rFonts'))
        rPr.append(f)
    f.set(qn('w:eastAsia'), '微软雅黑')
    return run

def create_physics_slide(prs, title_text, q_data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 背景与标题装饰 (保持老师的审美)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
    set_text_style(title_box.text_frame.paragraphs[0], is_title=True).text = title_text

    has_img = len(q_data['imgs']) > 0
    txt_w = Inches(8.5) if has_img else Inches(12.2)

    # 1. 原题卡片
    card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.0))
    card1.fill.solid(); card1.fill.fore_color.rgb = RGBColor(255, 255, 255); card1.line.color.rgb = RGBColor(230, 230, 235)
    
    badge1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(1.1), Inches(1.3), Inches(0.4))
    badge1.fill.solid(); badge1.fill.fore_color.rgb = RGBColor(0, 112, 192); badge1.line.fill.background()
    b1_p = badge1.text_frame.paragraphs[0]; b1_p.alignment = PP_ALIGN.CENTER
    br1 = b1_p.add_run(); br1.text = "原题呈现"; br1.font.size = Pt(13); br1.font.bold = True; br1.font.color.rgb = RGBColor(255, 255, 255)

    tf1 = card1.text_frame
    tf1.word_wrap = True
    # 逐行填入题干
    lines = q_data['text'].split('\n')
    for i, line in enumerate(lines):
        p = tf1.paragraphs[0] if i == 0 else tf1.add_paragraph()
        set_text_style(p, size=18).text = line.strip()

    # 2. 思路分析卡片
    card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.6), txt_w, Inches(1.5))
    card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(230, 230, 235)
    
    badge2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(5.4), Inches(1.3), Inches(0.4))
    badge2.fill.solid(); badge2.fill.fore_color.rgb = RGBColor(230, 90, 40); badge2.line.fill.background()
    b2_p = badge2.text_frame.paragraphs[0]; b2_p.alignment = PP_ALIGN.CENTER
    br2 = b2_p.add_run(); br2.text = "思路分析"; br2.font.size = Pt(13); br2.font.bold = True; br2.font.color.rgb = RGBColor(255, 255, 255)

    p2 = card2.text_frame.paragraphs[0]
    set_text_style(p2, size=16, color=(100, 100, 100)).text = "待补充详细受力分析与列式过程..."

    # 3. 图片与公式投放
    if has_img:
        y_cursor = 1.3
        # 按 y 坐标排序，确保图片和公式顺序自然
        q_data['imgs'].sort(key=lambda x: x['y'])
        for img_info in q_data['imgs'][:3]: # 增加到3张，兼容“图+公式”
            try:
                pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_cursor), width=Inches(3.8))
                y_cursor += (pic.height / 914400) + 0.15
            except: pass

# ==========================================
# 4. 逻辑控制与 Streamlit UI
# ==========================================
st.set_page_config(page_title="物理题 AI 课件工具", layout="centered")
st.title("⚛️ 物理习题自动化 PPT (全要素修复版)")

files = st.file_uploader("支持 Word/PDF/图片", accept_multiple_files=True, type=['png', 'jpg', 'pdf', 'docx'])

if st.button("开始生成精美 PPT", type="primary", use_container_width=True):
    if files:
        all_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                
                # --- Word 解析 ---
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    cur_q = None
                    for p in doc.paragraphs:
                        txt = p.text.strip()
                        if not txt: continue
                        if re.match(r'^\d+[\.．、]', txt):
                            if cur_q: all_qs.append(cur_q)
                            cur_q = {"text": txt, "imgs": []}
                        elif cur_q: cur_q['text'] += "\n" + txt
                    if cur_q: all_qs.append(cur_q)

                # --- PDF/图片 解析 ---
                else:
                    paths = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5))
                            p_p = os.path.join(tmp, f"p_{time.time()}.png")
                            pix.save(p_p); paths.append(p_p)
                    else:
                        p_p = os.path.join(tmp, file.name)
                        with open(p_p, "wb") as f: f.write(file.read())
                        paths.append(p_p)

                    for p_p in paths:
                        res, _ = engine(p_p)
                        # 获取所有插图和公式图
                        elements = extract_elements_v5(p_p, res, tmp)
                        
                        page_qs = []
                        cur_q = None
                        if res:
                            for line in res:
                                txt = line[1].strip()
                                if re.match(r'^\d+[\.．、]', txt):
                                    if cur_q: page_qs.append(cur_q)
                                    cur_q = {"text": txt, "imgs": [], "y": line[0][0][1]}
                                elif cur_q: cur_q['text'] += "\n" + txt
                            if cur_q: page_qs.append(cur_q)

                        # 将元素（图/公式）分配给最近的题目
                        for el in elements:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - el['y']))
                                target['imgs'].append(el)
                        all_qs.extend(page_qs)
            
            if all_qs:
                prs = Presentation()
                prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
                for i, q in enumerate(all_qs):
                    create_physics_slide(prs, f"习题精讲 第 {i+1} 题", q)
                
                buf = io.BytesIO()
                prs.save(buf)
                st.success(f"处理完成！找到 {len(all_qs)} 道题目。")
                st.download_button("📥 下载课件", buf.getvalue(), "物理教研课件.pptx", use_container_width=True)
