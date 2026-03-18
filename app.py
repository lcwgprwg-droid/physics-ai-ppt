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
# 2. 视觉算法：确保物理公式和插图 100% 捕获
# ==========================================
def extract_visual_elements_v14(img_path, ocr_result, out_dir):
    """
    针对物理 $T_1=300K$ 这类孤立公式优化的提取算法：
    1. 不再遮盖，直接找轮廓。
    2. 过滤掉被 OCR 识别为普通长文本的区域。
    3. 保留孤立的、无法识别或具有图形特征的块。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 灰度与二值化
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 极大横向膨胀：确保 $T_1=300K$ 这种公式能粘合成一块
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (35, 12))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # 获取 OCR 文字区域
    text_boxes = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            x_min, y_min = np.min(box, axis=0)
            x_max, y_max = np.max(box, axis=0)
            text_boxes.append([x_min, y_min, x_max, y_max])

    elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 排除边框和细碎噪点
        if w > w_img * 0.85 or h > h_img * 0.85: continue
        if w < 20 or h < 15: continue 
        
        # 判定是否为插图或重要公式：如果该区域不与 OCR 的文字块重叠 80% 以上，就切下来
        is_pure_text = False
        for tb in text_boxes:
            ix1, iy1 = max(x, tb[0]), max(y, tb[1])
            ix2, iy2 = min(x+w, tb[2]), min(y+h, tb[3])
            if ix1 < ix2 and iy1 < iy2:
                overlap = (ix2 - ix1) * (iy2 - iy1)
                if overlap / (w * h) > 0.8:
                    is_pure_text = True
                    break
        
        # 提取插图或孤立的大公式
        if not is_pure_text or (w * h > 5000):
            roi = img[y:y+h, x:x+w]
            if np.mean(roi) > 252: continue # 过滤空白
            
            f_path = os.path.join(out_dir, f"fig_{int(time.time()*1000)}.png")
            cv2.imwrite(f_path, roi)
            elements.append({"path": f_path, "y": y + h/2})
            
    return elements

# ==========================================
# 3. PPT 渲染引擎 (回归最稳健的写法)
# ==========================================
def apply_font_settings(run, size=18, is_bold=False, color=(30, 30, 30)):
    """
    稳健设置字体：设置后必须 return run ！！！
    """
    run.font.name = '微软雅黑'
    run.font.size = Pt(size)
    run.font.bold = is_bold
    run.font.color.rgb = RGBColor(*color)
    
    # 锁定中文字体 XML
    rPr = run._r.get_or_add_rPr()
    rFonts = qn('a:ea')
    # 手动添加字体声明，避开 API 报错
    from pptx.oxml import parse_xml
    f_xml = f'<a:ea xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="微软雅黑"/>'
    rPr.append(parse_xml(f_xml))
    
    return run  # 【核心修复】：必须返回对象，否则后续 .text 赋值会崩溃

def render_master_ppt(questions, tmp_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    c_blue = RGBColor(0, 112, 192)
    c_orange = RGBColor(230, 90, 40)

    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(245, 247, 250)
        bg.line.fill.background()
        
        # 装饰条
        bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
        bar.fill.solid()
        bar.fill.fore_color.rgb = c_blue
        bar.line.fill.background()
        
        # 标题 (左对齐)
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11), Inches(0.8))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.alignment = PP_ALIGN.LEFT
        apply_font_settings(p_title.add_run(), size=26, is_bold=True, color=(20, 40, 80)).text = f"习题精讲 第 {i+1} 题"
        
        has_img = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.3)

        # 1. 原题呈现卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.0))
        card.fill.solid()
        card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card.line.color.rgb = RGBColor(220, 220, 225)
        
        # 填充文字 (强制左对齐)
        tf = card.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        lines = q['text'].split('\n')
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT # 【核心改进】：锁定左对齐
            p.line_spacing = 1.2
            apply_font_settings(p.add_run(), size=18).text = line.strip()

        # 2. 思路分析卡片
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.5), txt_w, Inches(1.6))
        card2.fill.solid()
        card2.fill.fore_color.rgb = RGBColor(255, 255, 255)
        card2.line.color.rgb = RGBColor(220, 220, 225)
        
        p2 = card2.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.LEFT
        apply_font_settings(p2.add_run(), size=16, color=(120, 120, 120)).text = "待补充详细解析与列式过程..."

        # 3. 插图投放
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
# 4. Streamlit 逻辑控制
# ==========================================
st.set_page_config(page_title="物理教研课件 AI", layout="centered")
st.title("⚛️ 物理题自动 PPT (稳定版 - 已修复报错)")

files = st.file_uploader("上传 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 立即转换", type="primary", use_container_width=True):
    if not files:
        st.error("请先上传文件。")
    else:
        all_questions = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                
                # Word 处理
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
                
                # 视觉处理
                else:
                    paths = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for page in pdf:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                            p_p = os.path.join(tmp, f"p_{time.time()}.png")
                            pix.save(p_p); paths.append(p_p)
                    else:
                        p_p = os.path.join(tmp, file.name)
                        with open(p_p, "wb") as f: f.write(file.read())
                        paths.append(p_p)
                    
                    for p_p in paths:
                        res, _ = engine(p_p)
                        visuals = extract_visual_elements_v14(p_p, res, tmp)
                        
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
                        all_questions.extend(page_qs)

            if all_questions:
                ppt_data = render_master_ppt(all_questions, tmp)
                st.download_button("📥 下载 PPT 课件", ppt_data, "物理教研课件_稳定修复版.pptx", use_container_width=True)
                st.success(f"成功！已处理 {len(all_questions)} 道题目。")
