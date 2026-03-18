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
# 1. 核心引擎
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 专业排版引擎：混合字体渲染
# ==========================================
def add_mixed_styled_text(paragraph, text):
    """
    极高进阶：将文本拆分为中文、物理变量、数字。
    物理变量：Times New Roman + 斜体
    数字/单位：Times New Roman
    中文：微软雅黑
    下标：自动处理为底层 Subscript 属性
    """
    # 匹配规则：
    # 1. 物理变量及下标: ([a-zA-Z])_(\d+|[a-z])
    # 2. 普通英文/数字: ([a-zA-Z0-9\.\+\-\*/=]+)
    # 3. 其他(中文/标点)
    tokens = re.split(r'([a-zA-Z]_\d|[a-zA-Z]_[a-z]|[a-zA-Z]|[0-9\.]+)', text)
    
    for token in tokens:
        if not token: continue
        run = paragraph.add_run()
        
        # 情况 A: 物理下标 (如 T_1, v_0)
        if re.match(r'[a-zA-Z]_\d|[a-zA-Z]_[a-z]', token):
            base, sub = token.split('_')
            # 基础字母斜体
            run.text = base
            run.font.name = 'Times New Roman'
            run.font.italic = True
            # 添加下标 Run
            sub_run = paragraph.add_run()
            sub_run.text = sub
            sub_run.font.name = 'Times New Roman'
            sub_run.font.subscript = True
            sub_run.font.size = Pt(12)
            
        # 情况 B: 单个物理变量 (如 F, m, a)
        elif re.match(r'^[a-zA-Z]$', token):
            run.text = token
            run.font.name = 'Times New Roman'
            run.font.italic = True
            run.font.size = Pt(18)
            
        # 情况 C: 数字或单位 (如 10, 9.8, kg)
        elif re.match(r'^[0-9\.]+$', token):
            run.text = token
            run.font.name = 'Times New Roman'
            run.font.size = Pt(18)
            
        # 情况 D: 中文或其他
        else:
            run.text = token
            run.font.name = '微软雅黑'
            # 锁定中文字体 XML
            rPr = run._r.get_or_add_rPr()
            rPr.set(qn('a:ea'), '微软雅黑')
            run.font.size = Pt(18)

# ==========================================
# 3. OpenCV 视觉算法：文本避让 + 纯图捕获
# ==========================================
def extract_pure_diagrams_v15(img_path, ocr_result, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 建立文字保护区 mask
    text_mask = np.zeros((h_img, w_img), dtype=np.uint8)
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            # 适当扩大保护区，确保物理符号（上下标）不被 OpenCV 触碰
            cv2.fillPoly(text_mask, [box], 255)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 核心：只保留完全没有文字覆盖的线条
    only_diagram = cv2.subtract(binary, text_mask)
    
    # 闭运算：连接插图线条
    kernel = np.ones((10, 10), np.uint8)
    morphed = cv2.morphologyEx(only_diagram, cv2.MORPH_CLOSE, kernel)
    
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    diagrams = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.85 or h > h_img * 0.85: continue
        if w < 40 or h < 40: continue # 过滤掉可能的字符残渣

        roi = img[y:y+h, x:x+w]
        if np.mean(roi) > 253: continue
        
        f_path = os.path.join(out_dir, f"diag_{int(time.time()*1000)}.png")
        cv2.imwrite(f_path, roi)
        diagrams.append({"path": f_path, "y": y + h/2})
            
    return diagrams

# ==========================================
# 4. PPT 渲染：版式锁定
# ==========================================
def render_professional_ppt(questions, tmp_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blue = RGBColor(0, 112, 192)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        # 装饰条
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar.fill.solid(); bar.fill.fore_color.rgb = blue; bar.line.fill.background()
        
        # 标题 (混合排版)
        tb_title = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        p_title = tb_title.text_frame.paragraphs[0]; p_title.alignment = PP_ALIGN.LEFT
        add_mixed_styled_text(p_title, f"习题精讲 第 {i+1} 题")

        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.3)

        # 1. 题干卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.2))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 220, 225)
        
        tf = card.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        # 2. 核心：混合排版填充
        # 对题干内容进行初步清洗，识别下标
        raw_text = q['text'].replace("T1", "T_1").replace("v0", "v_0").replace("Ek", "E_k")
        lines = raw_text.split('\n')
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT; p.line_spacing = 1.3
            add_mixed_styled_text(p, line.strip())

        # 3. 思路分析卡片
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.6), txt_w, Inches(1.5))
        card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(220, 220, 225)
        p2 = card2.text_frame.paragraphs[0]; p2.alignment = PP_ALIGN.LEFT
        add_mixed_styled_text(p2, "待补充详细解析与列式计算过程...")

        # 4. 物理图投影
        if has_img:
            y_ptr = 1.3
            for img_info in q['imgs']:
                pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                y_ptr += (pic.height / 914400) + 0.2
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 5. UI 入口
# ==========================================
st.set_page_config(page_title="物理教研 AI 工具", layout="centered")
st.title("⚛️ 物理题自动 PPT (专业排版增强版)")

files = st.file_uploader("支持 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 生成专业 PPT", type="primary", use_container_width=True):
    if not files:
        st.error("请先上传文件")
    else:
        all_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in files:
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
                    # 图片/PDF 视觉分支
                    paths = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for p in pdf:
                            pix = p.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                            p_p = os.path.join(tmp, f"p_{time.time()}.png")
                            pix.save(p_p); paths.append(p_p)
                    else:
                        p_p = os.path.join(tmp, file.name)
                        with open(p_p, "wb") as f: f.write(file.read())
                        paths.append(p_p)
                    
                    for p_p in paths:
                        res, _ = engine(p_p)
                        # 核心：只抓纯线条图，避开文字/符号
                        diags = extract_pure_diagrams_v15(p_p, res, tmp)
                        
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
                        
                        for d in diags:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - d['y']))
                                target['imgs'].append(d)
                        all_qs.extend(page_qs)

            if all_qs:
                ppt_data = render_professional_ppt(all_qs, tmp)
                st.download_button("📥 下载专业 PPT 课件", ppt_data, "物理教研专家课件.pptx", use_container_width=True)
                st.success("课件生成成功！")
