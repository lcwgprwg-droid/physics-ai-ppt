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
# 1. 核心 OCR 引擎 (极速单例)
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：物理碎片全量捕获 (V22)
# ==========================================
def extract_visual_elements_v22(img_path, out_dir):
    """
    针对物理题：
    1. 捕捉所有笔迹块（图示、公式、特殊符号）。
    2. 使用 (30, 8) 的宽核，连接横向分布的公式 (如 v=10m/s)。
    3. 排除全页背景大框。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 高灵敏度自适应二值化
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 核心：使用宽矩形核膨胀，专门针对物理公式排版
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (30, 8))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    visuals = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        # 1. 尺寸过滤：排除全页外框 (教辅书外围圆角矩形)
        if w > w_img * 0.9 or h > h_img * 0.9: continue
        # 2. 面积过滤：保留具有视觉意义的块
        if w > 25 and h > 15:
            roi = img[y:y+h, x:x+w]
            # 过滤纯白无效块
            if np.mean(roi) > 252: continue
            
            f_path = os.path.join(out_dir, f"ve_{int(time.time()*1000)}_{x}.png")
            cv2.imwrite(f_path, roi)
            visuals.append({"path": f_path, "y": y + h/2, "area": w * h})
            
    # 按垂直位置排序，同时保证大图优先
    return sorted(visuals, key=lambda v: v['y'])

# ==========================================
# 3. PPT 渲染：【强制左对齐，标准 API】
# ==========================================
def render_master_ppt(questions):
    """
    标准渲染引擎：
    1. 强制左对齐排版。
    2. 兼容中文字体设置。
    3. 绝不使用链式调用。
    """
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(245, 247, 250)
        bg.line.fill.background()
        
        # 装饰条
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(0, 112, 192)
        bar.line.fill.background()
        
        # 标题 (左对齐)
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.alignment = PP_ALIGN.LEFT
        run_t = p_title.add_run()
        run_t.text = f"习题精讲 第 {i+1} 题"
        run_t.font.size = Pt(26)
        run_t.font.bold = True
        run_t.font.color.rgb = RGBColor(20, 40, 80)
        run_t.font.name = '微软雅黑'

        has_vis = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_vis else Inches(12.3)

        # 题干卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.3))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = RGBColor(220, 220, 225)
        
        tf = card.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        # 填充正文 (锁定左对齐)
        raw_text = q.get('text', '内容识别中...').strip()
        lines = raw_text.split('\n')
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.3
            r = p.add_run()
            r.text = line.strip()
            r.font.size = Pt(18)
            r.font.name = '微软雅黑'
            # 底层 XML 中文字体锁定
            try:
                rPr = r._r.get_or_add_rPr()
                ea = rPr.get_or_add_ea()
                ea.set('typeface', '微软雅黑')
            except: pass

        # 思路解析区
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.7), txt_w, Inches(1.4))
        card2.fill.solid()
        card2.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card2.line.color.rgb = RGBColor(220, 220, 225)
        p2 = card2.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run(); r2.text = "思路解析：正在整理中..."; r2.font.size = Pt(16); r2.font.color.rgb = RGBColor(120, 120, 120)

        # 投放插图 (物理图 & 公式原图)
        if has_vis:
            y_ptr = 1.3
            # 按垂直位置排序
            q['imgs'].sort(key=lambda x: x['y'])
            for img_info in q['imgs'][:3]:
                try:
                    slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                    y_ptr += 2.0 
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. 业务流
# ==========================================
st.set_page_config(page_title="物理教研课件 AI", layout="centered")
st.title("⚛️ 物理题自动 PPT (最终修复版)")

files = st.file_uploader("📤 上传资料 (Word/PDF/图片)", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 开启极速转换", type="primary", use_container_width=True):
    if not files:
        st.error("老师，请上传文件。")
    else:
        all_final_qs = []  # 【关键修复】：全局统一变量名
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmpdir:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                
                # --- Word 解析：暴力提取 ---
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    cur_q = None
                    for p in doc.paragraphs:
                        t = p.text.strip()
                        if not t: continue
                        # 兼容题号匹配
                        if re.match(r'^(\d+|[\(（]\d+[\)）])[\.．、\s]', t) or len(t) < 8:
                            if cur_q: all_final_qs.append(cur_q)
                            cur_q = {"text": t, "imgs": []}
                        elif cur_q: cur_q['text'] += "\n" + t
                    if cur_q: all_final_qs.append(cur_q)
                
                # --- 图片/PDF 解析：见字必录 ---
                else:
                    input_paths = []
                    if ext == 'pdf':
                        pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf_doc:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                            p_p = os.path.join(tmpdir, f"p_{time.time()}.png")
                            pix.save(p_p); input_paths.append(p_p)
                    else:
                        p_p = os.path.join(tmpdir, file.name)
                        with open(p_p, "wb") as f: f.write(file.read())
                        input_paths.append(p_p)
                    
                    for p_path in input_paths:
                        res, _ = engine(p_path)
                        # 核心改进：捕获所有视觉碎片
                        visual_elements = extract_visual_elements_v22(p_path, tmpdir)
                        
                        page_qs = []
                        cur_q = None
                        if res:
                            for line in res:
                                txt = line[1].strip()
                                # 只要识别到开头是数字或括号数字，就开启新题
                                if re.match(r'^(\d+|[\(（]\d+[\)）])', txt):
                                    if cur_q: page_qs.append(cur_q)
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                                elif cur_q: cur_q['text'] += "\n" + txt
                                elif not cur_q: # 兜底逻辑
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                            
                            if cur_q: page_qs.append(cur_q)
                        
                        # 空间关联
                        for ve in visual_elements:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - ve['y']))
                                target['imgs'].append(ve)
                        all_final_qs.extend(page_qs)

            if all_final_qs:
                ppt_buf = render_master_ppt(all_final_qs)
                st.download_button("📥 下载课件", ppt_buf, "物理精品课件_最终版.pptx", use_container_width=True)
                st.success(f"识别成功：共生成 {len(all_final_qs)} 页 PPT。")
            else:
                st.error("未能找到任何题目内容，请检查文件格式。")
