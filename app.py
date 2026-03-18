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
# 1. OCR 引擎：保持单例
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：全量捕获（绝不丢图）
# ==========================================
def extract_all_visuals(img_path, out_dir):
    """
    放弃所有复杂的过滤，只要是独立笔迹块，统统抓取。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 二值化，捕捉所有笔迹
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 膨胀，将断开的笔画（如公式字母）连在一起
    kernel = np.ones((10, 10), np.uint8)
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    visuals = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        # 排除全页大背景框（通过占页面比例判定）
        if w > w_img * 0.95 or h > h_img * 0.95: continue
        # 排除极小的噪点
        if w < 20 or h < 20: continue 
        
        roi = img[y:y+h, x:x+w]
        # 过滤掉纯白或非常淡的背景块
        if np.mean(roi) > 250: continue
        
        f_path = os.path.join(out_dir, f"vis_{int(time.time()*1000)}_{x}.png")
        cv2.imwrite(f_path, roi)
        visuals.append({"path": f_path, "y": y + h/2, "area": w * h})
            
    # 按面积降序，如果是物理题，前几个通常是插图或公式块
    return sorted(visuals, key=lambda x: x['area'], reverse=True)

# ==========================================
# 3. PPT 渲染：锁定左对齐，强制字体
# ==========================================
def render_physics_ppt(all_qs):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5) # 16:9
    
    c_blue = RGBColor(0, 112, 192)
    c_white = RGBColor(255, 255, 255)

    for i, q in enumerate(all_qs):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 标题 (左对齐)
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.8))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.alignment = PP_ALIGN.LEFT
        r_title = p_title.add_run()
        r_title.text = f"习题精讲 第 {i+1} 题"
        r_title.font.size, r_title.font.bold, r_title.font.color.rgb = Pt(26), True, RGBColor(20, 40, 80)
        
        # 判定排版
        has_imgs = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_imgs else Inches(12.3)

        # 题干白色圆角卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.2), txt_w, Inches(4.5))
        card.fill.solid(); card.fill.fore_color.rgb = c_white; card.line.color.rgb = RGBColor(220, 220, 225)
        
        tf = card.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        # 填入文字，绝不居中
        content = q.get('text', '未识别到文字内容').strip()
        lines = content.split('\n')
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT # 强制左对齐
            p.line_spacing = 1.3
            r = p.add_run()
            r.text = line
            r.font.name = '微软雅黑'
            r.font.size = Pt(18)
            # 中文字体底层注入
            try:
                rPr = r._r.get_or_add_rPr()
                ea = rPr.get_or_add_ea()
                ea.set('typeface', '微软雅黑')
            except: pass

        # 下方解析预留位
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.9), txt_w, Inches(1.2))
        card2.fill.solid(); card2.fill.fore_color.rgb = c_white; card2.line.color.rgb = RGBColor(220, 220, 225)
        p2 = card2.text_frame.paragraphs[0]; p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run(); r2.text = "思路解析：待补充..."; r2.font.size = Pt(16); r2.font.color.rgb = RGBColor(120, 120, 120)

        # 右侧图示投放
        if has_imgs:
            y_ptr = 1.2
            for img_info in q['imgs'][:3]: # 每页最多放3张最显著的图
                try:
                    slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                    y_ptr += 2.0 
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. Streamlit UI 层 (带详细日志反馈)
# ==========================================
st.set_page_config(page_title="物理课件 AI 最终版", layout="centered")
st.title("⚛️ 物理习题 AI 自动化 (零丢失修复版)")

files = st.file_uploader("📤 上传习题 (Word/PDF/图片)", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 开始转换", type="primary", use_container_width=True):
    if not files:
        st.error("老师，请先上传文件。")
    else:
        all_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmpdir:
            for file in files:
                st.info(f"正在处理: {file.name}")
                ext = file.name.split('.')[-1].lower()
                
                # --- Word 解析：暴力提取所有文字 ---
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    full_text = ""
                    for p in doc.paragraphs:
                        if p.text.strip():
                            full_text += p.text.strip() + "\n"
                    if full_text:
                        # 简单的按题号分割逻辑
                        parts = re.split(r'(\n\d+[\.．、\s]|\n[\(（]\d+[\)）])', "\n" + full_text)
                        cur_q = None
                        for p in parts:
                            if re.match(r'^\n\d+[\.．、\s]|\n[\(（]\d+[\)）]', p):
                                if cur_q: all_qs.append(cur_q)
                                cur_q = {"text": p.strip(), "imgs": []}
                            elif cur_q:
                                cur_q["text"] += p
                            else:
                                cur_q = {"text": p.strip(), "imgs": []}
                        if cur_q: all_qs.append(cur_q)
                
                # --- 图片/PDF 解析：见字必录 ---
                else:
                    paths = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                            p_path = os.path.join(tmpdir, f"p_{time.time()}.png")
                            pix.save(p_path); paths.append(p_path)
                    else:
                        p_path = os.path.join(tmpdir, file.name)
                        with open(p_path, "wb") as f: f.write(file.read())
                        paths.append(p_path)
                    
                    for p_p in paths:
                        res, _ = engine(p_p)
                        visuals = extract_all_visuals(p_p, tmpdir)
                        
                        page_qs = []
                        cur_q = None
                        if res:
                            for line in res:
                                txt = line[1].strip()
                                # 只要识别到数字开头，就开启新题；否则全部叠加
                                if re.match(r'^(\d+|[\(（]\d+[\)）])', txt):
                                    if cur_q: page_qs.append(cur_q)
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                                elif cur_q:
                                    cur_q["text"] += "\n" + txt
                                else:
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                        
                        if cur_q: page_qs.append(cur_q)
                        
                        # 关联图片
                        for v in visuals:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - v['y']))
                                target['imgs'].append(v)
                        all_qs.extend(page_qs)

            if all_qs:
                ppt_data = render_physics_ppt(all_qs)
                st.download_button("📥 下载课件", ppt_data, "物理教研精品课件.pptx", use_container_width=True)
                st.success(f"成功！已解析 {len(all_qs)} 道题目。")
            else:
                st.error("未能识别到内容，请检查文件。")
