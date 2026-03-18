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
# 1. 核心 OCR 引擎
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：暴力抓取全图所有视觉块
# ==========================================
def extract_all_visuals_v21(img_path, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 极大膨胀：确保公式和插图不破碎
    kernel = np.ones((20, 20), np.uint8)
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    visuals = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.96 or h > h_img * 0.96: continue # 过滤全页边框
        if w < 30 or h < 20: continue # 过滤极小杂点
        
        roi = img[y:y+h, x:x+w]
        if np.mean(roi) > 253: continue
        
        f_path = os.path.join(out_dir, f"vis_{int(time.time()*1000)}_{x}.png")
        cv2.imwrite(f_path, roi)
        visuals.append({"path": f_path, "y": y + h/2, "area": w * h})
    
    # 返回面积最大的前 4 个块（通常包含插图和核心公式）
    return sorted(visuals, key=lambda x: x['area'], reverse=True)[:4]

# ==========================================
# 3. PPT 渲染引擎：【极简赋值，绝不空白】
# ==========================================
def render_simple_ppt(questions):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 标题
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.8))
        title_box.text_frame.text = f"习题精讲 第 {i+1} 题"
        p_t = title_box.text_frame.paragraphs[0]
        p_t.alignment = PP_ALIGN.LEFT
        p_t.font.size, p_t.font.bold = Pt(26), True
        p_t.font.color.rgb = RGBColor(20, 40, 80)

        has_imgs = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_imgs else Inches(12.3)

        # 题干卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.2), txt_w, Inches(4.5))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 220, 225)
        
        # --- 核心修复：直接赋值，不进行任何 Run 拆分 ---
        tf = card.text_frame
        tf.word_wrap = True
        tf.text = q.get('text', '识别内容为空').strip()
        
        # 统一格式化卡片内文字
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.3
            for run in p.runs:
                run.font.name = '微软雅黑'
                run.font.size = Pt(18)
                run.font.color.rgb = RGBColor(40, 40, 40)
                # 尝试注入中文字体
                try:
                    rPr = run._r.get_or_add_rPr()
                    ea = rPr.get_or_add_ea()
                    ea.set('typeface', '微软雅黑')
                except: pass

        # 下方预留位
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.9), txt_w, Inches(1.2))
        card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(220, 220, 225)
        card2.text_frame.text = "思路解析：待补充..."
        p2 = card2.text_frame.paragraphs[0]
        p2.font.size, p2.font.color.rgb = Pt(16), RGBColor(120, 120, 120)

        # 投放插图
        if has_imgs:
            y_p = 1.2
            for img_info in q['imgs']:
                try:
                    slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_p), width=Inches(3.8))
                    y_p += 2.0 
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. Streamlit 控制层 (带调试预览)
# ==========================================
st.set_page_config(page_title="物理课件 AI 最终版", layout="centered")
st.title("⚛️ 物理习题 AI 自动化 (强制输出版)")

files = st.file_uploader("📥 上传习题 (Word/PDF/图片)", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 开始转换", type="primary", use_container_width=True):
    if not files:
        st.error("请先上传文件")
    else:
        all_final_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmpdir:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                
                # --- Word 解析：不再搜索题号，直接读取所有段落 ---
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    full_docx_text = ""
                    for p in doc.paragraphs:
                        if p.text.strip():
                            full_docx_text += p.text.strip() + "\n"
                    
                    if full_docx_text:
                        # 尝试拆分，如果拆分失败，就整篇作为一个题目
                        parts = re.split(r'(\n\d+[\.．、]|\n[\(（]\d+[\)）])', "\n" + full_docx_text)
                        if len(parts) <= 1:
                            all_final_qs.append({"text": full_docx_text, "imgs": []})
                        else:
                            cur_q = None
                            for part in parts:
                                if re.match(r'^\n\d+[\.．、]|\n[\(（]\d+[\)）]', part):
                                    if cur_q: all_final_qs.append(cur_q)
                                    cur_q = {"text": part.strip(), "imgs": []}
                                elif cur_q:
                                    cur_q["text"] += "\n" + part
                                else:
                                    cur_q = {"text": part.strip(), "imgs": []}
                            if cur_q: all_questions.append(cur_q)

                # --- 图片/PDF 解析：见字就录 ---
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
                        visuals = extract_all_visuals_v21(p_p, tmpdir)
                        
                        full_page_text = ""
                        if res:
                            for line in res:
                                full_page_text += line[1] + "\n"
                        
                        if full_page_text:
                            # 简单的分页逻辑：每张图片至少生成一页 PPT
                            all_final_qs.append({"text": full_page_text, "imgs": visuals, "y": 0})

            if all_final_qs:
                # 调试预览：在 Streamlit 界面显示提取到的前 100 个字符
                st.write(f"✅ 提取成功，准备渲染 {len(all_final_qs)} 页 PPT...")
                ppt_data = render_simple_ppt(all_final_qs)
                st.download_button("📥 点击下载生成的 PPT", ppt_data, "物理教研课件.pptx", use_container_width=True)
            else:
                st.error("❌ 无法提取任何文字，请确认文件内容不是纯色图片。")
