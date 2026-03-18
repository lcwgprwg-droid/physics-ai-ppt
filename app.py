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
from pptx.oxml import parse_xml

# ==========================================
# 1. 核心引擎 (RapidOCR)
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：物理特化型语义分割 (V13)
# ==========================================
def extract_visual_elements_v13(img_path, ocr_result, out_dir):
    """
    针对物理题重构：
    1. 彻底过滤页边大边框。
    2. 基于 OCR 区域的“反向挖洞”：先识别文字，再提取剩下的。
    3. 保护公式：对于孤立的数学特征区，强制作为图片保留。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 提取 OCR 文本框
    text_mask = np.zeros((h_img, w_img), dtype=np.uint8)
    text_boxes = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            text_boxes.append([x_min, y_min, x_max, y_max])
            # 在掩膜中涂白文字区（稍微扩充一点保护边缘）
            cv2.rectangle(text_mask, (x_min-2, y_min-2), (x_max+2, y_max+2), 255, -1)

    # 预处理
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 关键：从二值图中减去文字区，剩下插图和 OCR 漏掉的公式
    no_text_binary = cv2.subtract(binary, text_mask)
    
    # 横向加强膨胀：把物理公式 (T1=300K) 焊接成整体
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 10))
    dilated = cv2.dilate(no_text_binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        # --- 过滤逻辑 ---
        # 1. 过滤教辅书自带的大边框 (宽度超过80%页面)
        if w > w_img * 0.8: continue
        # 2. 过滤杂点
        if w < 20 or h < 15: continue
        
        # 裁剪原图内容
        roi = img[y:y+h, x:x+w]
        # 过滤空白区域
        if np.mean(roi) > 252: continue
        
        f_path = os.path.join(out_dir, f"diag_{int(time.time()*1000)}_{x}.png")
        cv2.imwrite(f_path, roi)
        elements.append({"path": f_path, "y": y + h/2})
            
    return elements

# ==========================================
# 3. PPT 渲染：底层 XML 注入 (彻底解决字体报错)
# ==========================================
def set_run_font_safe(run, font_name="微软雅黑", size=18, color_rgb=(40, 40, 40), is_bold=False):
    """
    使用底层 XML 注入，绕过所有不稳定的 API。
    """
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(*color_rgb)
    
    # PPT 必须通过设置 a:ea (East Asian) 节点来锁定中文字体
    rPr = run._r.get_or_add_rPr()
    # 注入 XML 节点
    xml_str = f'<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="{font_name}"/>'
    rPr.append(parse_xml(xml_str))
    # 同时设置西文字体
    run.font.name = font_name

def render_master_ppt(questions, tmp_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    c_blue = RGBColor(0, 112, 192)
    c_orange = RGBColor(230, 90, 40)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 1. 灰色背景 (分步设置，拒绝链式调用)
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(245, 247, 250)
        bg.line.fill.background()
        
        # 2. 蓝色装饰条
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar.fill.solid()
        bar.fill.fore_color.rgb = c_blue
        bar.line.fill.background()
        
        # 3. 标题 (显式左对齐)
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.alignment = PP_ALIGN.LEFT
        run_title = p_title.add_run()
        run_title.text = f"习题精讲 第 {i+1} 题"
        set_run_font_safe(run_title, size=26, is_bold=True, color_rgb=(20, 40, 80))
        
        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.3)

        # 4. 原题卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.0))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = RGBColor(220, 220, 225)
        
        # 填充正文 (显式左对齐)
        tf = card.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        for idx, line in enumerate(q['text'].split('\n')):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.2
            set_run_font_safe(p.add_run(), size=18).text = line.strip()

        # 5. 思路分析卡片
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.5), txt_w, Inches(1.6))
        card2.fill.solid()
        card2.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card2.line.color.rgb = RGBColor(220, 220, 225)
        
        p2 = card2.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.LEFT
        set_run_font_safe(p2.add_run(), size=16, color_rgb=(100, 100, 100)).text = "待补充详细受力分析与列式讲解过程..."

        # 6. 投放视觉元素 (插图/公式)
        if has_img:
            y_ptr = 1.3
            q['imgs'].sort(key=lambda x: x['y']) # 按垂直位置排版
            for img_info in q['imgs'][:3]:
                try:
                    pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                    y_ptr += (pic.height / 914400) + 0.2
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. 业务流
# ==========================================
st.set_page_config(page_title="物理教研课件专家", layout="centered")
st.title("⚛️ 物理题自动 PPT 生成 (API 深度加固版)")

files = st.file_uploader("支持 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 立即转换并下载", type="primary", use_container_width=True):
    if not files:
        st.error("老师，请上传文件。")
    else:
        all_questions = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                
                # --- Word 解析 ---
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    cur_q = None
                    for p in doc.paragraphs:
                        t = p.text.strip()
                        if not t: continue
                        if re.match(r'^\d+[\.．、]', t):
                            if cur_q: all_questions.append(cur_q)
                            cur_q = {"text": t, "imgs": []}
                        elif cur_q: cur_q['text'] += "\n" + t
                    if cur_q: all_questions.append(cur_q)
                
                # --- 图片/PDF 视觉解析 ---
                else:
                    paths = []
                    if ext == 'pdf':
                        pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf_doc:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5))
                            p_path = os.path.join(tmp, f"p_{time.time()}.png")
                            pix.save(p_path); paths.append(p_path)
                    else:
                        p_path = os.path.join(tmp, file.name)
                        with open(p_path, "wb") as f: f.write(file.read())
                        paths.append(p_path)
                    
                    for p_p in paths:
                        res, _ = engine(p_p)
                        visuals = extract_visual_elements_v13(p_p, res, tmp)
                        
                        page_qs = []
                        cur_q = None
                        if res:
                            for line in res:
                                txt = line[1].strip()
                                # 识别 1. 或 (1) 等题号
                                if re.match(r'^\d+[\.．、]', txt):
                                    if cur_q: page_qs.append(cur_q)
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                                elif cur_q: cur_q['text'] += "\n" + txt
                            if cur_q: page_qs.append(cur_q)
                        
                        for v in visuals:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - v['y']))
                                target['imgs'].append(v)
                        all_questions.extend(page_qs)

            if all_questions:
                ppt_data = render_master_ppt(all_questions, tmp)
                st.download_button("📥 点击下载生成的 PPT 课件", ppt_data, "物理教研精选课件.pptx", use_container_width=True)
                st.success(f"识别成功：共生成 {len(all_questions)} 道题目。")
