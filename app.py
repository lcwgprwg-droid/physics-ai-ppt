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
# 2. 视觉算法：物理特化型语义分割 (V9)
# ==========================================
def extract_visual_elements_v9(img_path, ocr_result, out_dir):
    """
    针对物理题重构：
    1. 极大核膨胀：确保 $T_1=300K$ 这种孤立公式能成块。
    2. 坐标碰撞检测：彻底排除 OCR 已识别的正文文字。
    3. 边框免疫：通过长宽比和面积比过滤掉教辅书的大边框。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 记录 OCR 文字区域
    text_boxes = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            text_boxes.append([x_min, y_min, x_max, y_max])

    # 预处理：灰度 -> 中值滤波 -> 自适应二值化
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.medianBlur(gray, 3)
    binary = cv2.adaptiveThreshold(blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 极大核形态学：横向连接物理公式和图示线条
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 15))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 1. 物理环境过滤：排除全页大框 (教辅书外圈)
        if w > w_img * 0.85 or h > h_img * 0.85: continue
        if w < 20 or h < 15: continue 

        # 2. 坐标碰撞检测：计算该轮廓与 OCR 结果的重叠情况
        is_text_only = False
        for tb in text_boxes:
            # 计算相交区域
            ix1, iy1 = max(x, tb[0]), max(y, tb[1])
            ix2, iy2 = min(x+w, tb[2]), min(y+h, tb[3])
            if ix1 < ix2 and iy1 < iy2:
                intersection = (ix2 - ix1) * (iy2 - iy1)
                # 如果这个块 85% 以上都是 OCR 认出的正文，就不作为图片提取
                if intersection / (w * h) > 0.85:
                    is_text_only = True
                    break
        
        # 3. 提取插图或孤立公式
        if not is_text_only or (w * h > 6000):
            roi = img[y:y+h, x:x+w]
            # 过滤掉几乎纯白的无效区域
            if np.mean(roi) > 252: continue
            
            f_path = os.path.join(out_dir, f"fig_{int(time.time()*1000)}_{x}.png")
            cv2.imwrite(f_path, roi)
            elements.append({"path": f_path, "y": y + h/2})
            
    return elements

# ==========================================
# 3. PPT 渲染：分步声明法 (修复 AttributeError)
# ==========================================
def set_shape_style(shape, bg_color=None, line_color=None):
    """安全地设置形状样式，不再使用链式调用"""
    if bg_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = bg_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()

def set_font_safe(run, size=18, is_bold=False, color=(40, 40, 40)):
    """设置字体及 EastAsia 兼容"""
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(*color)
    run.font.name = '微软雅黑'
    # 强制设置 PPT 底层 XML 字体
    rPr = run._r.get_or_add_rPr()
    for tag in ['a:latin', 'a:ea', 'a:cs']:
        f = rPr.get_or_add_rFonts()
        f.set(qn(tag), '微软雅黑')

def render_master_ppt(questions, tmp_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5) # 16:9
    
    c_blue = RGBColor(0, 112, 192)
    c_orange = RGBColor(230, 90, 40)
    c_white = RGBColor(255, 255, 255)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 1. 灰色背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        set_shape_style(bg, bg_color=RGBColor(245, 247, 250))
        
        # 2. 蓝色装饰条 (修复之前的报错点)
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        set_shape_style(bar, bg_color=c_blue)
        
        # 3. 标题 - 强制左对齐
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        tf_title = title_box.text_frame
        p_title = tf_title.paragraphs[0]
        p_title.alignment = PP_ALIGN.LEFT
        run_title = p_title.add_run()
        run_title.text = f"习题精讲 第 {i+1} 题"
        set_font_safe(run_title, size=26, is_bold=True, color=(20, 40, 80))
        
        # 判定布局
        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.3)

        # 4. 原题卡片
        card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.0))
        set_shape_style(card1, bg_color=c_white, line_color=RGBColor(220, 220, 225))
        
        # 标签1
        badge1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(1.1), Inches(1.3), Inches(0.4))
        set_shape_style(badge1, bg_color=c_blue)
        b1p = badge1.text_frame.paragraphs[0]
        b1p.alignment = PP_ALIGN.CENTER
        set_font_safe(b1p.add_run(), size=13, is_bold=True, color=(255, 255, 255)).text = "原题呈现"
        
        # 填入题干 - 强制左对齐
        tf1 = card1.text_frame
        tf1.word_wrap = True
        tf1.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        for idx, line in enumerate(q['text'].split('\n')):
            p = tf1.paragraphs[0] if idx == 0 else tf1.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.2
            set_font_safe(p.add_run(), size=18).text = line.strip()

        # 5. 思路卡片
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.5), txt_w, Inches(1.6))
        set_shape_style(card2, bg_color=c_white, line_color=RGBColor(220, 220, 225))
        
        badge2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(5.3), Inches(1.3), Inches(0.4))
        set_shape_style(badge2, bg_color=c_orange)
        b2p = badge2.text_frame.paragraphs[0]
        b2p.alignment = PP_ALIGN.CENTER
        set_font_safe(b2p.add_run(), size=13, is_bold=True, color=(255, 255, 255)).text = "思路分析"
        
        p2 = card2.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.LEFT
        set_font_safe(p2.add_run(), size=16, color=(100, 100, 100)).text = "待补充详细受力分析与列式过程..."

        # 6. 投放图片与公式
        if has_img:
            y_ptr = 1.3
            q['imgs'].sort(key=lambda x: x['y'])
            for img_info in q['imgs'][:3]:
                try:
                    pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                    y_ptr += (pic.height / 914400) + 0.2
                except: pass
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. 核心逻辑控制
# ==========================================
st.set_page_config(page_title="物理教研 AI 工具", layout="centered")
st.title("⚛️ 物理题自动 PPT 生成 (完美修正版)")

files = st.file_uploader("支持 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 立即转换", type="primary", use_container_width=True):
    if not files:
        st.error("请先上传文件")
    else:
        all_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                
                # Word 解析
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
                
                # PDF/图片 视觉解析
                else:
                    paths = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for p in pdf:
                            pix = p.get_pixmap(matrix=fitz.Matrix(2.5, 2.5))
                            p_p = os.path.join(tmp, f"p_{time.time()}.png")
                            pix.save(p_p); paths.append(p_p)
                    else:
                        p_p = os.path.join(tmp, file.name)
                        with open(p_p, "wb") as f: f.write(file.read())
                        paths.append(p_p)
                    
                    for p_p in paths:
                        res, _ = engine(p_p)
                        visuals = extract_visual_elements_v9(p_p, res, tmp)
                        
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
                        
                        for v in visuals:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - v['y']))
                                target['imgs'].append(v)
                        all_qs.extend(page_qs)

            if all_qs:
                ppt_data = render_master_ppt(all_qs, tmp)
                st.download_button("📥 下载 PPT 课件", ppt_data, "物理教研课件_精修版.pptx", use_container_width=True)
                st.success("课件生成成功！")
