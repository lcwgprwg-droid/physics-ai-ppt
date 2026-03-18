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
# 2. 视觉算法：物理特化型语义分割 (V11)
# ==========================================
def extract_visual_elements_v11(img_path, ocr_result, out_dir):
    """
    专门针对物理公式和受力图设计的抓取算法：
    1. 不再暴力遮盖文字，而是计算重叠度。
    2. 使用宽矩形核膨胀，专门连接横向分布的公式 (如 T1=300K)。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 获取 OCR 文字区域
    text_boxes = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            text_boxes.append([x_min, y_min, x_max, y_max])

    # 预处理
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # --- 核心改进：针对物理公式的横向膨胀 ---
    # 使用 (40, 10) 的核，这能强行把横向分布的字符和等号“焊接”成一个图片块
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 10))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.9 or h > h_img * 0.9: continue 
        if w < 20 or h < 15: continue 

        # 计算该块与 OCR 正文的重叠面积
        is_text_covered = False
        for tb in text_boxes:
            ix1, iy1 = max(x, tb[0]), max(y, tb[1])
            ix2, iy2 = min(x+w, tb[2]), min(y+h, tb[3])
            if ix1 < ix2 and iy1 < iy2:
                intersection = (ix2 - ix1) * (iy2 - iy1)
                # 如果这个块被 OCR 文字占满了 80% 以上，就认为它只是普通文字，不去切图
                if intersection / (w * h) > 0.8:
                    is_text_covered = True
                    break
        
        # 物理特化：如果是插图，或者它是孤立的大公式，则切图
        if not is_text_covered or (w * h > 6000):
            roi = img[y:y+h, x:x+w]
            if np.mean(roi) > 252: continue # 过滤纯白无效区
            
            f_path = os.path.join(out_dir, f"diag_{int(time.time()*1000)}_{x}.png")
            cv2.imwrite(f_path, roi)
            elements.append({"path": f_path, "y": y + h/2})
            
    return elements

# ==========================================
# 3. PPT 渲染引擎 (底层 XML 修复版)
# ==========================================
def set_font_v11(run, size=18, is_bold=False, color=(40, 40, 40)):
    """
    【核心修复】：彻底解决 AttributeError: get_or_add_rFonts
    使用 PPT 专用的 DrawingML 节点属性
    """
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(*color)
    
    # 强制注入中文字体
    rPr = run._r.get_or_add_rPr()
    # 西文字体
    latin = rPr.get_or_add_latin()
    latin.set('typeface', '微软雅黑')
    # 中文字体 (ea = East Asian)
    ea = rPr.get_or_add_ea()
    ea.set('typeface', '微软雅黑')

def render_ppt_v11(questions, tmp_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    blue = RGBColor(0, 112, 192)
    orange = RGBColor(230, 90, 40)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 蓝色侧边条 (分步设置，严禁链式调用)
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar.fill.solid(); bar.fill.fore_color.rgb = blue; bar.line.fill.background()
        
        # 标题 - 强制左对齐
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        tf_title = title_box.text_frame
        p_title = tf_title.paragraphs[0]
        p_title.alignment = PP_ALIGN.LEFT
        run_title = p_title.add_run()
        run_title.text = f"习题精讲 第 {i+1} 题"
        set_font_v11(run_title, size=26, is_bold=True, color=(20, 40, 80))
        
        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.3)

        # 1. 原题卡片
        card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.0))
        card1.fill.solid(); card1.fill.fore_color.rgb = RGBColor(255, 255, 255); card1.line.color.rgb = RGBColor(220, 220, 225)
        
        badge1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(1.1), Inches(1.3), Inches(0.4))
        badge1.fill.solid(); badge1.fill.fore_color.rgb = blue; badge1.line.fill.background()
        b1p = badge1.text_frame.paragraphs[0]; b1p.alignment = PP_ALIGN.CENTER
        set_font_v11(b1p.add_run(), size=13, is_bold=True, color=(255, 255, 255)).text = "原题呈现"
        
        # 填入文字 - 强制左对齐
        tf1 = card1.text_frame; tf1.word_wrap = True; tf1.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        lines = q['text'].split('\n')
        for idx, line in enumerate(lines):
            p = tf1.paragraphs[0] if idx == 0 else tf1.add_paragraph()
            p.alignment = PP_ALIGN.LEFT; p.line_spacing = 1.2
            set_font_v11(p.add_run(), size=18).text = line.strip()

        # 2. 思路分析卡片
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.5), txt_w, Inches(1.6))
        card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(220, 220, 225)
        
        badge2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(5.3), Inches(1.3), Inches(0.4))
        badge2.fill.solid(); badge2.fill.fore_color.rgb = orange; badge2.line.fill.background()
        b2p = badge2.text_frame.paragraphs[0]; b2p.alignment = PP_ALIGN.CENTER
        set_font_safe_v11 = set_font_v11 # 别名兼容
        set_font_v11(b2p.add_run(), size=13, is_bold=True, color=(255, 255, 255)).text = "思路分析"
        
        p2 = card2.text_frame.paragraphs[0]; p2.alignment = PP_ALIGN.LEFT
        set_font_v11(p2.add_run(), size=16, color=(100, 100, 100)).text = "待补充详细解析过程..."

        # 3. 投放图片与公式
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
# 4. Streamlit 控制层
# ==========================================
st.set_page_config(page_title="物理课件 AI 助手", layout="centered")
st.title("⚛️ 物理题自动 PPT (排版&字体修复版)")

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
                
                # --- Word 分支 ---
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    cur_q = None
                    for p in doc.paragraphs:
                        txt = p.text.strip()
                        if not txt: continue
                        if re.match(r'^\d+[\.．、]', txt):
                            if cur_q: all_qs.append(cur_q)
                            cur_q = {"text": txt, "imgs": []}
                        elif cur_q: cur_q['text'] += "\n" + txt
                    if cur_q: all_qs.append(cur_q)
                
                # --- 视觉分支 ---
                else:
                    paths = []
                    if ext == 'pdf':
                        pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf_doc:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5))
                            p_p = os.path.join(tmp, f"p_{time.time()}.png")
                            pix.save(p_p); paths.append(p_p)
                    else:
                        p_p = os.path.join(tmp, file.name)
                        with open(p_p, "wb") as f: f.write(file.read())
                        paths.append(p_p)
                    
                    for p_p in paths:
                        res, _ = engine(p_p)
                        visuals = extract_visual_elements_v11(p_p, res, tmp)
                        
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
                ppt_data = render_ppt_v11(all_qs, tmp)
                st.download_button("📥 下载 PPT 课件", ppt_data, "物理教研课件_修复版.pptx", use_container_width=True)
                st.success(f"成功识别 {len(all_qs)} 道题目！")
