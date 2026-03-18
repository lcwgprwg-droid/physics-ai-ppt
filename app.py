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
# 2. 视觉增强：精准剔除公式与边框
# ==========================================
def extract_diagrams_v4(img_path, ocr_result, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 建立遮罩：完全覆盖文字区，且向外扩充 10 像素
    mask_canvas = np.zeros((h_img, w_img), dtype=np.uint8)
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            # 扩大遮罩，把公式的尾巴、下标彻底盖死
            cv2.rectangle(mask_canvas, (x_min-10, y_min-10), (x_max+10, y_max+10), 255, -1)

    # 预处理
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 245, 255, cv2.THRESH_BINARY_INV)
    
    # 剔除文字遮罩区域
    thresh[mask_canvas > 0] = 0
    
    # 闭运算：连接断开的物理线条
    kernel = np.ones((7, 7), np.uint8)
    morphed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel)
    
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    diagrams = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 1. 尺寸过滤：排除大边框和极小杂点
        if w > w_img * 0.8 or h > h_img * 0.8: continue 
        if w < 40 or h < 40: continue
        
        # 2. 关键：像素密度检测（防公式误杀）
        # 物理图示通常有较多线条，而残留的公式碎片像素点极少
        roi_thresh = morphed[y:y+h, x:x+w]
        black_pixel_ratio = np.sum(roi_thresh == 255) / (w * h)
        if black_pixel_ratio < 0.02: # 如果黑色像素占比低于2%，认为是公式残渣或空白框
            continue

        roi_color = img[y:y+h, x:x+w]
        f_path = os.path.join(out_dir, f"diag_{int(time.time()*1000)}.png")
        cv2.imwrite(f_path, roi_color)
        diagrams.append({"path": f_path, "y": y + h/2})
        
    return diagrams

# ==========================================
# 3. PPT 高级排版引擎（显式左对齐）
# ==========================================
def set_run_font(run, size=18, color=(40, 40, 40), bold=False):
    run.font.name = '微软雅黑'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)
    # 强制设置东亚字体
    rPr = run._r.get_or_add_rPr()
    f = rPr.find(qn('w:rFonts'))
    if f is None:
        f = rPr.makeelement(qn('w:rFonts'))
        rPr.append(f)
    f.set(qn('w:eastAsia'), '微软雅黑')

def render_ppt_final(questions):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    c_blue = RGBColor(0, 112, 192)
    c_orange = RGBColor(230, 90, 40)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 顶部装饰条
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar.fill.solid(); bar.fill.fore_color.rgb = c_blue; bar.line.fill.background()
        
        # 标题 (左对齐)
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        tp = title_box.text_frame.paragraphs[0]
        tp.alignment = PP_ALIGN.LEFT
        tr = tp.add_run()
        tr.text = f"习题精讲 第 {i+1} 题"
        set_run_font(tr, size=26, color=(20, 40, 80), bold=True)
        
        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.3) if has_img else Inches(12.3)
        
        # --- 卡片1：原题呈现 ---
        card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.4), txt_w, Inches(3.8))
        card1.fill.solid(); card1.fill.fore_color.rgb = RGBColor(255, 255, 255); card1.line.color.rgb = RGBColor(220, 220, 225)
        
        # 标签1
        badge1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(1.2), Inches(1.3), Inches(0.4))
        badge1.fill.solid(); badge1.fill.fore_color.rgb = c_blue; badge1.line.fill.background()
        b1p = badge1.text_frame.paragraphs[0]
        b1p.alignment = PP_ALIGN.CENTER
        b1r = b1p.add_run(); b1r.text = "原题呈现"; set_run_font(b1r, size=13, color=(255, 255, 255), bold=True)
        
        # 正文1 (关键：强制左对齐)
        tf1 = card1.text_frame
        tf1.word_wrap = True
        tf1.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        tf1.paragraphs[0].alignment = PP_ALIGN.LEFT
        
        lines = q['text'].split('\n')
        for idx, line in enumerate(lines):
            p = tf1.paragraphs[0] if idx == 0 else tf1.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.2
            r = p.add_run(); r.text = line.strip()
            set_run_font(r, size=18)

        # --- 卡片2：思路分析 ---
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.6), txt_w, Inches(1.5))
        card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(220, 220, 225)
        
        badge2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(5.4), Inches(1.3), Inches(0.4))
        badge2.fill.solid(); badge2.fill.fore_color.rgb = c_orange; badge2.line.fill.background()
        b2p = badge2.text_frame.paragraphs[0]; b2p.alignment = PP_ALIGN.CENTER
        b2r = b2p.add_run(); b2r.text = "思路分析"; set_run_font(b2r, size=13, color=(255, 255, 255), bold=True)
        
        tf2 = card2.text_frame
        tf2.paragraphs[0].alignment = PP_ALIGN.LEFT
        r2 = tf2.paragraphs[0].add_run(); r2.text = "待补充详细解析过程..."; set_run_font(r2, size=16, color=(100, 100, 100))

        # 插图投放
        if has_img:
            y_ptr = 1.4
            for img_p in q['imgs'][:2]:
                try:
                    pic = slide.shapes.add_picture(img_p, Inches(9.0), Inches(y_ptr), width=Inches(4.0))
                    y_ptr += (pic.height / 914400) + 0.2
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. 逻辑控制与 Streamlit UI
# ==========================================
st.set_page_config(page_title="物理教研课件专家", layout="centered")
st.title("⚛️ 物理题自动生成 PPT (视觉修复版)")

files = st.file_uploader("支持 Word/PDF/图片", accept_multiple_files=True, type=['png', 'jpg', 'pdf', 'docx'])

if st.button("开始生成精美 PPT", type="primary", use_container_width=True):
    if files:
        all_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                # 处理 Word
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
                
                # 处理 PDF 和 图片
                else:
                    imgs_to_proc = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                            p_path = os.path.join(tmp, f"p_{time.time()}.png")
                            pix.save(p_path)
                            imgs_to_proc.append(p_path)
                    else:
                        p_path = os.path.join(tmp, file.name)
                        with open(p_path, "wb") as f: f.write(file.read())
                        imgs_to_proc.append(p_path)
                    
                    for p_path in imgs_to_proc:
                        res, _ = engine(p_path)
                        diags = extract_diagrams_v4(p_path, res, tmp)
                        # 题目聚合
                        page_qs = []
                        cur_q = None
                        for line in res:
                            txt = line[1].strip()
                            if re.match(r'^\d+[\.．、]', txt):
                                if cur_q: page_qs.append(cur_q)
                                cur_q = {"text": txt, "imgs": [], "y": line[0][0][1]}
                            elif cur_q: cur_q['text'] += "\n" + txt
                        if cur_q: page_qs.append(cur_q)
                        # 图文匹配
                        for d in diags:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - d['y']))
                                target['imgs'].append(d['path'])
                        all_qs.extend(page_qs)
            
            if all_qs:
                ppt_buf = render_ppt_final(all_qs)
                st.success(f"成功！已处理 {len(all_qs)} 道题目。")
                st.download_button("📥 下载课件", ppt_buf, "物理精选课件.pptx", use_container_width=True)
