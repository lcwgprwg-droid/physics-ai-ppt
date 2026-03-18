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
# 1. 常量配置（设计规范）
# ==========================================
class PPTConfig:
    WIDTH = Inches(13.333)
    HEIGHT = Inches(7.5)
    COLOR_BG = RGBColor(245, 247, 250)
    COLOR_BLUE = RGBColor(0, 112, 192)
    COLOR_TEXT = RGBColor(0, 0, 0)      # 强行锁定黑色，杜绝变白
    COLOR_SUBTEXT = RGBColor(80, 80, 80)
    FONT_MAIN = "微软雅黑"
    FONT_MATH = "Times New Roman"

# ==========================================
# 2. 核心 OCR 引擎 (RapidOCR)
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 3. 视觉工程：物理符号与图像捕获 (V23)
# ==========================================
def extract_physics_vision(img_path, ocr_result, out_dir):
    """
    高级视觉策略：
    1. 计算 OCR 识别框，建立文字保护区。
    2. 利用 Morphological Transformations 连接物理线条。
    3. 提取所有非文字核心区的视觉元素（图+公式）。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 获取文字区域掩膜
    text_mask = np.zeros((h_img, w_img), dtype=np.uint8)
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            cv2.fillPoly(text_mask, [box], 255)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 保护物理细线：降低阈值
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 21, 8)
    
    # 物理特化：横向膨胀合并公式
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 10))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    # 移除被 OCR 确认的密集文字区，剩下孤立的符号和插图
    clean_dilated = cv2.subtract(dilated, text_mask)
    
    contours, _ = cv2.findContours(clean_dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    visual_elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.9 or h > h_img * 0.9: continue # 过滤大边框
        if w < 25 or h < 20: continue # 过滤杂质
        
        roi = img[y:y+h, x:x+w]
        if np.mean(roi) > 252: continue
        
        f_name = f"diag_{int(time.time()*1000)}_{x}.png"
        f_path = os.path.join(out_dir, f_name)
        cv2.imwrite(f_path, roi)
        visual_elements.append({"path": f_path, "y": y + h/2, "area": w * h})
            
    return visual_elements

# ==========================================
# 4. PPT 渲染 Builder (强制视觉规范)
# ==========================================
def add_styled_paragraph(tf, text, is_title=False):
    """
    显式声明每一个段落的属性，确保绝无白字、绝不居中。
    """
    p = tf.add_paragraph() if len(tf.paragraphs) > 0 and tf.paragraphs[0].text else tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = 1.2
    
    run = p.add_run()
    run.text = text.strip()
    run.font.name = PPTConfig.FONT_MAIN
    run.font.size = Pt(24 if is_title else 18)
    run.font.bold = is_title
    run.font.color.rgb = PPTConfig.COLOR_TEXT # 强制黑色
    
    # 强制东亚语言字体注入
    rPr = run._r.get_or_add_rPr()
    ea = rPr.get_or_add_ea()
    ea.set('typeface', PPTConfig.FONT_MAIN)
    return p

def render_physics_master(all_qs):
    prs = Presentation()
    prs.slide_width = PPTConfig.WIDTH
    prs.slide_height = PPTConfig.HEIGHT

    for i, q in enumerate(all_qs):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景底色
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = PPTConfig.COLOR_BG; bg.line.fill.background()
        
        # 蓝色装饰指示器
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar.fill.solid(); bar.fill.fore_color.rgb = PPTConfig.COLOR_BLUE; bar.line.fill.background()
        
        # 页面标题
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        add_styled_paragraph(title_box.text_frame, f"习题精讲 第 {i+1} 题", is_title=True)
        
        has_imgs = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_imgs else Inches(12.3)

        # 1. 题干白色卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.3))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(210, 210, 215)
        
        tf_q = card.text_frame
        tf_q.word_wrap = True
        tf_q.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        # 逐行填入文字
        content_lines = q.get('text', '').split('\n')
        for line in content_lines:
            if line.strip():
                add_styled_paragraph(tf_q, line)

        # 2. 思路分析卡片
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.7), txt_w, Inches(1.4))
        card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(210, 210, 215)
        
        tf_s = card2.text_frame
        p_s = add_styled_paragraph(tf_s, "思路解析：待补充受力分析与公式推导过程...")
        p_s.runs[0].font.color.rgb = PPTConfig.COLOR_SUBTEXT # 辅助文字灰色
        p_s.runs[0].font.size = Pt(16)

        # 3. 投放视觉元素
        if has_imgs:
            y_ptr = 1.3
            q['imgs'].sort(key=lambda x: x['y'])
            for img_info in q['imgs'][:3]: # 取前3个核心图
                try:
                    slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                    y_ptr += 2.0 
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 5. 业务逻辑层：支持混合输入
# ==========================================
st.set_page_config(page_title="高级教研 AI 工具", layout="centered")
st.title("⚛️ AI 物理教研工作站 (程序员视角版)")

uploaded_files = st.file_uploader("📥 上传 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 极速生成 PPT", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("老师，请先上传文件。")
    else:
        all_final_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmpdir:
            for file in uploaded_files:
                st.write(f"正在深度解析: {file.name}")
                ext = file.name.split('.')[-1].lower()
                
                # --- Word 解析：100% 捕获 ---
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    cur_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                    # 鲁棒切分逻辑
                    parts = re.split(r'(\d+[\.．、\s])', cur_text)
                    if len(parts) <= 1:
                        all_final_qs.append({"text": cur_text, "imgs": []})
                    else:
                        temp_q = None
                        for p in parts:
                            if re.match(r'\d+[\.．、\s]', p):
                                if temp_q: all_final_qs.append(temp_q)
                                temp_q = {"text": p, "imgs": []}
                            elif temp_q:
                                temp_q["text"] += p
                        if temp_q: all_final_qs.append(temp_q)
                
                # --- 视觉解析 (PDF/图片) ---
                else:
                    input_paths = []
                    if ext == 'pdf':
                        pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf_doc:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5)) # 提高画质
                            p_p = os.path.join(tmpdir, f"p_{time.time()}.png")
                            pix.save(p_p); input_paths.append(p_p)
                    else:
                        p_p = os.path.join(tmpdir, file.name)
                        with open(p_p, "wb") as f: f.write(file.read())
                        input_paths.append(p_p)
                    
                    for p_path in input_paths:
                        res, _ = engine(p_path)
                        visuals = extract_physics_vision(p_path, res, tmpdir)
                        
                        # 语义化切题
                        page_qs = []
