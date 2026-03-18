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
# 1. 核心配置与单例模型
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法重构：基于OCR反向遮罩的图像提取
# ==========================================
def extract_pure_diagrams(img_path, ocr_result, out_dir):
    """
    核心重构：先抹除文字，再找图片，彻底解决公式被识别为图片的问题
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 创建一个和原图一样大的纯白遮罩
    clean_canvas = img.copy()
    
    # --- 关键步骤：把所有OCR识别到的文字区域全部涂白 ---
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            # 稍微扩大一点遮罩范围，确保公式的上下标也被盖住
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            cv2.rectangle(clean_canvas, (x_min-5, y_min-5), (x_max+5, y_max+5), (255, 255, 255), -1)

    # 现在的 clean_canvas 只剩下纯插图和边框了
    gray = cv2.cvtColor(clean_canvas, cv2.COLOR_BGR2GRAY)
    # 二值化
    _, thresh = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY_INV)
    
    # 闭运算：连接物理线条
    kernel = np.ones((10, 10), np.uint8)
    morphed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
    
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    diagrams = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 过滤机制
        if w > w_img * 0.9 or h > h_img * 0.9: continue # 过滤外边框
        if w < 30 or h < 30: continue # 过滤噪点
        
        # 截取原图对应位置
        roi = img[y:y+h, x:x+w]
        
        # 检查ROI内的像素密度，防止误切空白区域
        if np.mean(cv2.cvtColor(roi, cv2.COLOR_BGR2GRAY)) > 250: continue

        f_path = os.path.join(out_dir, f"diag_{int(time.time()*1000)}.png")
        cv2.imwrite(f_path, roi)
        diagrams.append({"path": f_path, "y": y + h/2})
        
    return diagrams

# ==========================================
# 3. Word 解析引擎
# ==========================================
def process_word_input(file_bytes, temp_dir):
    doc = Document(io.BytesIO(file_bytes))
    parsed_items = []
    current_q = None
    
    # 记录图片序号，用于关联
    img_counter = 0
    
    for element in doc.element.body:
        # 如果是段落
        if element.tag.endswith('p'):
            para = [p for p in doc.paragraphs if p._element == element][0]
            text = para.text.strip()
            if not text: continue
            
            # 匹配题号
            if re.match(r'^\d+[\.．、]', text):
                if current_q: parsed_items.append(current_q)
                current_q = {"text": text, "imgs": []}
            elif current_q:
                current_q["text"] += "\n" + text
                
        # 如果是嵌入式图片（简单处理Word图片的顺序关联）
        elif element.tag.endswith('drawing') or 'Graphic' in element.tag:
             # Word图片提取逻辑较为复杂，此处简化为寻找rel
             pass

    # 处理Word图片附件
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_counter += 1
            img_p = os.path.join(temp_dir, f"w_img_{img_counter}.png")
            with open(img_p, "wb") as f:
                f.write(rel.target_part.blob)
            if parsed_items: # 顺序关联
                parsed_items[min(img_counter-1, len(parsed_items)-1)]['imgs'].append(img_p)
                
    if current_q and current_q not in parsed_items:
        parsed_items.append(current_q)
    return parsed_items

# ==========================================
# 4. PPT 完美版式渲染
# ==========================================
def render_ppt_professional(questions):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    blue = RGBColor(0, 112, 192)
    orange = RGBColor(230, 90, 40)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 侧边指示条
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar.fill.solid(); bar.fill.fore_color.rgb = blue; bar.line.fill.background()
        
        # 标题
        title = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        tr = title.text_frame.paragraphs[0]
        tr.text = f"习题精讲 第 {i+1} 题"
        tr.font.bold = True; tr.font.size = Pt(26); tr.font.color.rgb = RGBColor(30, 40, 60)
        
        # 题干区域
        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.2)
        
        # 原题呈现卡片
        card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.0))
        card1.fill.solid(); card1.fill.fore_color.rgb = RGBColor(255, 255, 255); card1.line.color.rgb = RGBColor(220, 220, 225)
        
        badge1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(1.15), Inches(1.2), Inches(0.35))
        badge1.fill.solid(); badge1.fill.fore_color.rgb = blue; badge1.line.fill.background()
        b1r = badge1.text_frame.paragraphs[0]
        b1r.text = "原题呈现"; b1r.font.size = Pt(12); b1r.font.bold = True; b1r.font.color.rgb = RGBColor(255, 255, 255); b1r.alignment = PP_ALIGN.CENTER
        
        tf1 = card1.text_frame
        tf1.word_wrap = True; tf1.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        p1 = tf1.paragraphs[0]; p1.text = q['text']; p1.font.size = Pt(18); p1.font.color.rgb = RGBColor(40, 40, 40)
        
        # 思路分析卡片
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.6), txt_w, Inches(1.5))
        card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(220, 220, 225)
        
        badge2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(5.45), Inches(1.2), Inches(0.35))
        badge2.fill.solid(); badge2.fill.fore_color.rgb = orange; badge2.line.fill.background()
        b2r = badge2.text_frame.paragraphs[0]
        b2r.text = "思路分析"; b2r.font.size = Pt(12); b2r.font.bold = True; b2r.font.color.rgb = RGBColor(255, 255, 255); b2r.alignment = PP_ALIGN.CENTER
        
        tf2 = card2.text_frame
        p2 = tf2.paragraphs[0]; p2.text = "待补充详细解析过程..."; p2.font.size = Pt(16); p2.font.color.rgb = RGBColor(80, 80, 80)
        
        # 图片投放
        if has_img:
            y_offset = 1.3
            for img_p in q['imgs'][:2]: # 最多两张
                pic = slide.shapes.add_picture(img_p, Inches(9.2), Inches(y_offset), width=Inches(3.8))
                y_offset += (pic.height / 914400) + 0.2
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 5. Streamlit 入口
# ==========================================
st.set_page_config(page_title="物理AI课件助手", layout="centered")
st.title("⚛️ 物理题自动解析工坊")

files = st.file_uploader("支持 Word/PDF/图片", accept_multiple_files=True, type=['png', 'jpg', 'pdf', 'docx'])

if st.button("开始生成精美 PPT", type="primary"):
    if files:
        all_questions = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                if ext == 'docx':
                    all_questions.extend(process_word_input(file.read(), tmp))
                elif ext == 'pdf':
                    pdf = fitz.open(stream=file.read(), filetype="pdf")
                    for page in pdf:
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        p_path = os.path.join(tmp, f"p_{time.time()}.png")
                        pix.save(p_path)
                        res, _ = engine(p_path)
                        diags = extract_pure_diagrams(p_path, res, tmp)
                        # 简单的坐标聚合逻辑
                        page_qs = []
                        cur_q = None
                        for line in res:
                            txt = line[1].strip()
                            if re.match(r'^\d+[\.．、]', txt):
                                if cur_q: page_qs.append(cur_q)
                                cur_q = {"text": txt, "imgs": [], "y": line[0][0][1]}
                            elif cur_q: cur_q['text'] += "\n" + txt
                        if cur_q: page_qs.append(cur_q)
                        for d in diags:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - d['y']))
                                target['imgs'].append(d['path'])
                        all_questions.extend(page_qs)
                else: # 图片
                    p_path = os.path.join(tmp, file.name)
                    with open(p_path, "wb") as f: f.write(file.read())
                    res, _ = engine(p_path)
                    diags = extract_pure_diagrams(p_path, res, tmp)
                    page_qs = []
                    cur_q = None
                    for line in res:
                        txt = line[1].strip()
                        if re.match(r'^\d+[\.．、]', txt):
                            if cur_q: page_qs.append(cur_q)
                            cur_q = {"text": txt, "imgs": [], "y": line[0][0][1]}
                        elif cur_q: cur_q['text'] += "\n" + txt
                    if cur_q: page_qs.append(cur_q)
                    for d in diags:
                        if page_qs:
                            target = min(page_qs, key=lambda q: abs(q['y'] - d['y']))
                            target['imgs'].append(d['path'])
                    all_questions.extend(page_qs)
            
            if all_questions:
                ppt_buf = render_ppt_professional(all_questions)
                st.download_button("📥 下载 PPT 课件", ppt_buf, "课件.pptx")
