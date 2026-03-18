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
# 1. 核心 OCR 引擎 (RapidOCR)
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：物理特化型提取 (V18 - 无遮罩稳健版)
# ==========================================
def extract_visual_v18(img_path, out_dir):
    """
    针对物理受力图、电路图、复杂公式优化：
    1. 放弃文字遮罩，防止误伤 $T_1=300K$。
    2. 使用形态学闭运算，将散碎的公式字母和插图线条连接。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    # 自适应阈值，应对教辅书拍照时光照不均
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 膨胀：将公式字母(T1)和受力箭头连成整体块
    # 使用 15x15 的核，保证物理符号不被拆散
    kernel = np.ones((15, 15), np.uint8)
    morphed = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)
    
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 排除全页大框（教辅书边框）
        if w > w_img * 0.85 or h > h_img * 0.85: continue
        # 排除微小噪点
        if w < 30 or h < 30: continue 
        # 排除极其细长的分割线
        if w / h > 15 or h / w > 15: continue

        roi = img[y:y+h, x:x+w]
        # 灰度判定：如果是纯白块则丢弃
        if np.mean(roi) > 250: continue
        
        f_path = os.path.join(out_dir, f"phys_{int(time.time()*1000)}_{x}.png")
        cv2.imwrite(f_path, roi)
        elements.append({"path": f_path, "y": y + h/2})
            
    return elements

# ==========================================
# 3. PPT 渲染：【标准 API 锁定版】
# ==========================================
def safe_set_font(run, size, color_rgb, is_bold=False):
    """标准 API 设置字体，绝不报错"""
    run.font.name = '微软雅黑'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(*color_rgb)

def safe_set_shape(shape, fill_color_rgb=None, line_color_rgb=None):
    """标准 API 分步设置形状，拒绝链式调用"""
    if fill_color_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color_rgb
    if line_color_rgb:
        shape.line.color.rgb = line_color_rgb
    else:
        # 无轮廓
        try:
            shape.line.fill.background()
        except:
            pass

def render_safe_ppt(questions):
    """
    100% 采用标准 API 渲染，保证兼容性
    """
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    c_blue = RGBColor(0, 112, 192)
    c_orange = RGBColor(230, 90, 40)
    c_white = RGBColor(255, 255, 255)
    c_gray_bg = RGBColor(245, 247, 250)
    c_text = RGBColor(40, 40, 40)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 1. 灰色背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        safe_set_shape(bg, fill_color_rgb=c_gray_bg)
        
        # 2. 蓝色装饰条
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        safe_set_shape(bar, fill_color_rgb=c_blue)
        
        # 3. 标题 (左对齐)
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.alignment = PP_ALIGN.LEFT
        run_title = p_title.add_run()
        run_title.text = f"习题精讲 第 {i+1} 题"
        safe_set_font(run_title, 26, (20, 40, 80), is_bold=True)
        
        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.3)

        # 4. 原题呈现卡片
        card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.0))
        safe_set_shape(card1, fill_color_rgb=c_white, line_color_rgb=RGBColor(220, 220, 225))
        
        # 标签1
        badge1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(1.1), Inches(1.3), Inches(0.4))
        safe_set_shape(badge1, fill_color_rgb=c_blue)
        b1p = badge1.text_frame.paragraphs[0]
        b1p.alignment = PP_ALIGN.CENTER
        safe_set_font(b1p.add_run(), 13, (255, 255, 255), is_bold=True).text = "原题呈现"
        
        # 正文 - 强制左对齐
        tf1 = card1.text_frame
        tf1.word_wrap = True
        tf1.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        lines = q['text'].split('\n')
        for idx, line in enumerate(lines):
            p = tf1.paragraphs[0] if idx == 0 else tf1.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.3
            safe_set_font(p.add_run(), 18, (40, 40, 40)).text = line.strip()

        # 5. 思路分析卡片
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.5), txt_w, Inches(1.6))
        safe_set_shape(card2, fill_color_rgb=c_white, line_color_rgb=RGBColor(220, 220, 225))
        
        badge2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(5.3), Inches(1.3), Inches(0.4))
        safe_set_shape(badge2, fill_color_rgb=c_orange)
        b2p = badge2.text_frame.paragraphs[0]
        b2p.alignment = PP_ALIGN.CENTER
        safe_set_font(b2p.add_run(), 13, (255, 255, 255), is_bold=True).text = "思路分析"
        
        p2 = card2.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.LEFT
        safe_set_font(p2.add_run(), 16, (100, 100, 100)).text = "待补充详细解析与受力分析过程..."

        # 6. 图片投放
        if has_img:
            y_ptr = 1.3
            q['imgs'].sort(key=lambda x: x['y'])
            for img_info in q['imgs'][:3]:
                try:
                    pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                    y_ptr += (pic.height / 914400) + 0.2
                except:
                    pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. 业务控制逻辑
# ==========================================
st.set_page_config(page_title="物理题 AI 专家版", layout="centered")
st.title("⚛️ 物理题自动 PPT (工业级稳定版)")

files = st.file_uploader("支持 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 立即转换", type="primary", use_container_width=True):
    if not files:
        st.error("请先上传文件")
    else:
        all_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmpdir:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                
                # --- Word 解析 ---
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    cur_q = None
                    for p in doc.paragraphs:
                        txt = p.text.strip()
                        if not txt: continue
                        # 兼容题号识别
                        if re.match(r'^(\d+|[\(（]\d+[\)）])[\.．、\s]', txt) or len(txt) < 10:
                            if cur_q: all_qs.append(cur_q)
                            cur_q = {"text": txt, "imgs": []}
                        elif cur_q: cur_q['text'] += "\n" + txt
                    if cur_q: all_qs.append(cur_q)
                
                # --- 视觉解析 ---
                else:
                    paths = []
                    if ext == 'pdf':
                        pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf_doc:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                            p_path = os.path.join(tmpdir, f"p_{time.time()}.png")
                            pix.save(p_path); paths.append(p_path)
                    else:
                        p_path = os.path.join(tmpdir, file.name)
                        with open(p_path, "wb") as f: f.write(file.read())
                        paths.append(p_path)
                    
                    for p_p in paths:
                        res, _ = engine(p_p)
                        # 核心改进：回归稳定提取，不丢物理插图
                        visuals = extract_visual_v18(p_p, tmpdir)
                        
                        page_qs = []
                        cur_q = None
                        if res:
                            for line in res:
                                txt = line[1].strip()
                                if re.match(r'^(\d+|[\(（]\d+[\)）])', txt):
                                    if cur_q: page_qs.append(cur_q)
                                    cur_q = {"text": txt, "imgs": [], "y": (line[0][0][1] + line[0][2][1])/2}
                                elif cur_q: cur_q['text'] += "\n" + txt
                            if cur_q: page_qs.append(cur_q)
                        
                        for v in visuals:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - v['y']))
                                target['imgs'].append(v)
                        all_qs.extend(page_qs)

            if all_qs:
                ppt_data = render_safe_ppt(all_qs)
                st.download_button("📥 下载课件", ppt_data, "物理 AI 精品课件.pptx", use_container_width=True)
                st.success(f"成功！已处理 {len(all_qs)} 道题目。")
