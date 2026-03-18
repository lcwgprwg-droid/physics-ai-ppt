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
# 1. 核心 OCR 引擎 (极速版)
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 专业排版引擎：物理变量斜体 + 强制左对齐
# ==========================================
def apply_physics_layout(paragraph, text):
    """
    实现物理规范排版：
    变量(v, m, a, t, F)斜体，数字正体。
    强制左对齐。
    """
    paragraph.alignment = PP_ALIGN.LEFT
    paragraph.line_spacing = 1.3
    
    # 简单的物理变量拆分正则
    parts = re.split(r'([a-zA-Z]_\d|[a-zA-Z]|[0-9\.]+)', text)
    
    for part in parts:
        if not part: continue
        run = paragraph.add_run()
        run.text = part
        run.font.size = Pt(18)
        run.font.color.rgb = RGBColor(40, 40, 40)
        
        # 匹配物理变量 (单个字母或带下标) -> 斜体 Times New Roman
        if re.match(r'^[a-zA-Z]$|^[a-zA-Z]_\d$', part):
            run.font.name = 'Times New Roman'
            run.font.italic = True
        # 匹配数字 -> 正体 Times New Roman
        elif re.match(r'^[0-9\.]+$', part):
            run.font.name = 'Times New Roman'
        # 中文 -> 微软雅黑
        else:
            run.font.name = '微软雅黑'
            # 兼容性处理：设置东亚字体
            rPr = run._r.get_or_add_rPr()
            h_fonts = qn('a:ea')
            # 使用更稳健的底层设置
            try:
                etree_obj = rPr.get_or_add_latin()
                etree_obj.set('typeface', '微软雅黑')
            except: pass

# ==========================================
# 3. 视觉算法 V16：保护符号与提取图示
# ==========================================
def extract_diagrams_v16(img_path, ocr_result, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 1. 建立文字屏蔽区（防止公式被切走）
    text_mask = np.zeros((h_img, w_img), dtype=np.uint8)
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            cv2.fillPoly(text_mask, [box], 255)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 只提取 OCR 没认出来的线条区域
    only_diagrams = cv2.subtract(binary, text_mask)
    
    # 形态学：连接物理图线条
    kernel = np.ones((10, 10), np.uint8)
    morphed = cv2.morphologyEx(only_diagrams, cv2.MORPH_CLOSE, kernel)
    
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    diagrams = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.8 or h > h_img * 0.8: continue
        if w < 35 or h < 35: continue 

        roi = img[y:y+h, x:x+w]
        if np.mean(roi) > 252: continue
        
        f_path = os.path.join(out_dir, f"diag_{int(time.time()*1000)}.png")
        cv2.imwrite(f_path, roi)
        diagrams.append({"path": f_path, "y": y + h/2})
            
    return diagrams

# ==========================================
# 4. PPT 生成主引擎 (稳定性修复)
# ==========================================
def make_physics_ppt(all_qs):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5) # 16:9
    
    blue = RGBColor(0, 112, 192)
    
    for i, q in enumerate(all_qs):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 装饰
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar.fill.solid(); bar.fill.fore_color.rgb = blue; bar.line.fill.background()
        
        # 标题
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        apply_physics_layout(title_box.text_frame.paragraphs[0], f"习题精讲 第 {i+1} 题")
        
        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.3)

        # 题干卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.2))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 220, 225)
        
        tf = card.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        lines = q['text'].split('\n')
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            apply_physics_layout(p, line.strip())

        # 解析预留
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.6), txt_w, Inches(1.5))
        card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(220, 220, 225)
        apply_physics_layout(card2.text_frame.paragraphs[0], "待补充详细受力分析与计算过程...")

        # 投放插图
        if has_img:
            y_ptr = 1.3
            for img_info in q['imgs']:
                pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                y_ptr += (pic.height / 914400) + 0.2
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 5. 业务流控制 (兼容各种奇葩格式)
# ==========================================
st.set_page_config(page_title="物理 AI 课件专家", layout="centered")
st.title("⚛️ 物理题自动 PPT (稳定不报错版)")

files = st.file_uploader("支持 Word/PDF/图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 一键生成 PPT", type="primary", use_container_width=True):
    if not files:
        st.error("请上传文件")
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
                        # 兼容：1. 或 (1) 或 1、 或 2025.
                        if re.match(r'^(\d+|[\(（]\d+[\)）])[\.．、\s]', t) or len(t) < 5:
                            if cur_q: all_qs.append(cur_q)
                            cur_q = {"text": t, "imgs": []}
                        elif cur_q: cur_q['text'] += "\n" + t
                    if cur_q: all_qs.append(cur_q)
                
                else:
                    # PDF/图片 处理
                    img_list = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for p in pdf:
                            pix = p.get_pixmap(matrix=fitz.Matrix(2, 2))
                            img_p = os.path.join(tmp, f"page_{time.time()}.png")
                            pix.save(img_p); img_list.append(img_p)
                    else:
                        img_p = os.path.join(tmp, file.name)
                        with open(img_p, "wb") as f: f.write(file.read())
                        img_list.append(img_p)
                    
                    for ip in img_list:
                        res, _ = engine(ip)
                        diags = extract_diagrams_v16(ip, res, tmp)
                        
                        page_qs = []
                        cur_q = None
                        # 兜底：如果整页没匹配到题号，直接把整页文字给 cur_q
                        if res:
                            for line in res:
                                txt = line[1].strip()
                                # 极度宽松匹配：只要开头是数字或括号数字
                                if re.match(r'^(\d+|[\(（]\d+[\)）])', txt):
                                    if cur_q: page_qs.append(cur_q)
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                                elif cur_q:
                                    cur_q['text'] += "\n" + txt
                                elif not cur_q: # 兜底逻辑：第一行如果不是题号，强行开启一个题目
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                        
                        if cur_q: page_qs.append(cur_q)
                        
                        # 关联图片
                        for d in diags:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - d['y']))
                                target['imgs'].append(d)
                        all_qs.extend(page_qs)

            if all_qs:
                ppt_data = make_physics_ppt(all_qs)
                st.download_button("📥 下载生成好的 PPT", ppt_data, "物理教研课件.pptx", use_container_width=True)
                st.success(f"解析成功！共提取 {len(all_qs)} 道题目。")
            else:
                st.error("❌ 未能识别到有效题目，请确保上传的资料文字清晰。")
