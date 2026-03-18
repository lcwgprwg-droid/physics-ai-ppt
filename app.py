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
# 2. 视觉算法：物理碎片聚类抓取 (V20)
# ==========================================
def extract_physics_visuals(img_path, out_dir):
    """
    针对物理 $T_1=300K$ 和细受力图优化：
    1. 不再遮盖文字，直接全图扫描。
    2. 使用横向超长核 (45, 5) 膨胀，确保公式不散。
    3. 排除全页大框，抓取所有独立视觉元素。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 物理细线增强：自适应二值化
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 核心：横向膨胀，专门把横向排列的字母和公式“焊”在一起
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (45, 5))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    visuals = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        # 过滤教辅书自带的大边框
        if w > w_img * 0.85 or h > h_img * 0.85: continue
        # 过滤微小杂点
        if w < 25 or h < 15: continue 
        
        roi = img[y:y+h, x:x+w]
        # 过滤空白
        if np.mean(roi) > 252: continue
        
        f_name = f"phys_vis_{int(time.time()*1000)}_{x}.png"
        f_path = os.path.join(out_dir, f_name)
        cv2.imwrite(f_path, roi)
        visuals.append({"path": f_path, "y": y + h/2})
            
    return visuals

# ==========================================
# 3. PPT 渲染引擎：【原生拆解，绝不报错】
# ==========================================
def render_master_ppt(questions):
    """
    彻底废弃所有链式调用和底层XML注入。
    使用官方最稳健的逐行设置方式。
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    c_blue = RGBColor(0, 112, 192)
    c_orange = RGBColor(230, 90, 40)
    c_white = RGBColor(255, 255, 255)
    c_bg = RGBColor(245, 247, 250)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # --- A. 灰色背景 (逐行赋值) ---
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg_fill = bg.fill
        bg_fill.solid()
        bg_fill.fore_color.rgb = c_bg
        bg.line.fill.background()
        
        # --- B. 蓝色侧边条 ---
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar_fill = bar.fill
        bar_fill.solid()
        bar_fill.fore_color.rgb = c_blue
        bar.line.fill.background()
        
        # --- C. 标题 (左对齐) ---
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        tf_title = title_box.text_frame
        p_title = tf_title.paragraphs[0]
        p_title.alignment = PP_ALIGN.LEFT
        run_title = p_title.add_run()
        run_title.text = f"习题精讲 第 {i+1} 题"
        run_title.font.name = '微软雅黑'
        run_title.font.size = Pt(26)
        run_title.font.bold = True
        run_title.font.color.rgb = RGBColor(20, 40, 80)
        
        # 判定是否有图片
        has_vis = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_vis else Inches(12.3)

        # --- D. 原题卡片 ---
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.2))
        c_fill = card.fill
        c_fill.solid()
        c_fill.fore_color.rgb = c_white
        card.line.color.rgb = RGBColor(220, 220, 225)
        
        tf_q = card.text_frame
        tf_q.word_wrap = True
        tf_q.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        # 逐行填入题干 (锁定左对齐)
        lines = q['text'].split('\n')
        for idx, line in enumerate(lines):
            p = tf_q.paragraphs[0] if idx == 0 else tf_q.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.3
            r = p.add_run()
            r.text = line.strip()
            r.font.name = '微软雅黑'
            r.font.size = Pt(18)
            # 锁定东亚字体
            try:
                rPr = r._r.get_or_add_rPr()
                ea = rPr.get_or_add_ea()
                ea.set('typeface', '微软雅黑')
            except: pass

        # --- E. 思路卡片 ---
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.6), txt_w, Inches(1.5))
        c2_fill = card2.fill
        c2_fill.solid()
        c2_fill.fore_color.rgb = c_white
        card2.line.color.rgb = RGBColor(220, 220, 225)
        
        tf_s = card2.text_frame
        p_s = tf_s.paragraphs[0]
        p_s.alignment = PP_ALIGN.LEFT
        r_s = p_s.add_run()
        r_s.text = "待补充详细受力分析与列式过程..."
        r_s.font.name = '微软雅黑'
        r_s.font.size = Pt(16)
        r_s.font.color.rgb = RGBColor(120, 120, 120)

        # --- F. 图片/公式投放 ---
        if has_vis:
            y_ptr = 1.3
            q['imgs'].sort(key=lambda x: x['y'])
            for img_info in q['imgs'][:3]: # 限制数量防溢出
                try:
                    slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                    # 此处不再尝试 pic.height 链式操作，靠默认纵横比
                    y_ptr += 1.8 
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. 业务控制逻辑
# ==========================================
st.set_page_config(page_title="物理教研课件专家", layout="centered")
st.title("⚛️ 物理题全自动 PPT (最终修正版)")

uploaded_files = st.file_uploader("📤 上传习题资料 (Word/PDF/图片)", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("✨ 立即开始智能转换", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("老师，请上传文件后再开始。")
    else:
        all_final_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmpdir:
            for file in uploaded_files:
                ext = file.name.split('.')[-1].lower()
                
                # --- Word 分支 ---
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    cur_q = None
                    for p in doc.paragraphs:
                        txt = p.text.strip()
                        if not txt: continue
                        # 兼容题号识别 (1. 或 (1) 或 2025.)
                        if re.match(r'^(\d+|[\(（]\d+[\)）])[\.．、\s]', txt) or len(txt) < 6:
                            if cur_q: all_final_qs.append(cur_q)
                            cur_q = {"text": txt, "imgs": []}
                        elif cur_q: cur_q['text'] += "\n" + txt
                    if cur_q: all_final_qs.append(cur_q)
                
                # --- 图片/PDF 视觉分支 ---
                else:
                    input_paths = []
                    if ext == 'pdf':
                        pdf_data = fit_open = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf_data:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                            p_p = os.path.join(tmpdir, f"p_{time.time()}.png")
                            pix.save(p_p); input_paths.append(p_p)
                    else:
                        p_p = os.path.join(tmpdir, file.name)
                        with open(p_p, "wb") as f: f.write(file.read())
                        input_paths.append(p_p)
                    
                    for p_path in input_paths:
                        res, _ = engine(p_path)
                        # 核心改进：抓取图中所有独立物理块 (公式/图)
                        visual_elements = extract_physics_visuals(p_path, tmpdir)
                        
                        page_qs = []
                        cur_q = None
                        if res:
                            for line in res:
                                txt = line[1].strip()
                                # 极度容错题号匹配
                                if re.match(r'^(\d+|[\(（]\d+[\)）])', txt):
                                    if cur_q: page_qs.append(cur_q)
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                                elif cur_q: cur_q['text'] += "\n" + txt
                                elif not cur_q: # 兜底第一题
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
                st.download_button("📥 下载课件", ppt_buf, "物理教研课件_修正版.pptx", use_container_width=True)
                st.success(f"转换成功！共生成 {len(all_final_qs)} 页 PPT。")
            else:
                st.error("未能找到任何题目内容，请检查文件格式。")
