import os
import re
import cv2
import math
import tempfile
import io
import time
import streamlit as st
import fitz  # PyMuPDF
import docx  # Python-docx
from rapidocr_onnxruntime import RapidOCR
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from PIL import Image, ImageEnhance

# ==========================================
# 1. 核心引擎：图像处理与图表截取
# ==========================================
def crop_diagrams(img_path, out_dir):
    """从图片中自动检测并裁剪物理插图"""
    img = cv2.imread(img_path)
    if img is None: return []
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (20, 6))
    dilated = cv2.dilate(thresh, kernel, iterations=2)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    diagrams = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        # 过滤过小或比例畸形的区域，锁定物理插图
        if h > 50 and w > 50 and (w * h) > 8000 and 0.1 < (w / float(h)) < 8.0:
            y1, y2 = max(0, y - 15), min(img.shape[0], y + h + 15)
            x1, x2 = max(0, x - 15), min(img.shape[1], x + w + 15)
            diagrams.append({"img": img[y1:y2, x1:x2], "x_center": x + w/2, "y_center": y + h/2})

    saved_imgs = []
    for i, d in enumerate(diagrams):
        p = os.path.join(out_dir, f"fig_{int(time.time()*1000)}_{i}.png")
        roi_rgb = cv2.cvtColor(d["img"], cv2.COLOR_BGR2RGB)
        # 简单的背景白化处理
        mask_bg = (roi_rgb[:,:,0] > 215) & (roi_rgb[:,:,1] > 215) & (roi_rgb[:,:,2] > 215)
        roi_rgb[mask_bg] = [255, 255, 255]
        pil_img = ImageEnhance.Sharpness(ImageEnhance.Contrast(Image.fromarray(roi_rgb)).enhance(1.1)).enhance(1.5)
        pil_img.save(p, quality=95)
        saved_imgs.append({"path": p, "x_center": d["x_center"], "y_center": d["y_center"]})
    return saved_imgs

# ==========================================
# 2. 核心引擎：智能 OCR 与题目切分
# ==========================================
def is_noise(text):
    noise_words = ['复习与提高', 'A组', 'B组', '高中物理', '扫描全能王', 'Page', '页码']
    return any(w in text for w in noise_words) or re.match(r'^\s*\d+\s*$', text)

def smart_ocr_and_split(img_path, cv_images):
    """OCR 识别并根据坐标将题目与插图关联"""
    engine = RapidOCR()
    result, _ = engine(img_path)
    if not result: return []

    # 判断是否为双栏排版
    all_x = [line[0][0][0] for line in result]
    page_center_x = (max(all_x) + min(all_x)) / 2 if all_x else 500

    sorted_lines = []
    for line in result:
        box, text = line[0], line[1]
        cx = (box[0][0] + box[1][0]) / 2
        cy = (box[0][1] + box[3][1]) / 2
        col_idx = 0 if cx < page_center_x else 1
        sorted_lines.append({"col": col_idx, "y": cy, "box": box, "text": text})

    # 先排左栏再排右栏，同栏按 Y 轴排序
    sorted_lines.sort(key=lambda item: (item['col'], item['y']))
    
    questions, current_q = [], None
    for item in sorted_lines:
        text = item['text'].strip()
        if is_noise(text): continue
        
        # 匹配题号开头：如 "1.", "2．", "15、"
        if re.match(r'^\s*\d+[\.．、](?!\d)', text):
            if current_q: questions.append(current_q)
            current_q = {'text': text, 'x_min': item['box'][0][0], 'y_min': item['box'][0][1], 
                         'x_max': item['box'][1][0], 'y_max': item['box'][2][1], 'matched_imgs': []}
        elif current_q:
            current_q['text'] += "\n" + text if re.match(r'^[A-D][\.．、]', text) else text
            current_q['x_min'] = min(current_q['x_min'], item['box'][0][0])
            current_q['y_min'] = min(current_q['y_min'], item['box'][0][1])
            current_q['x_max'] = max(current_q['x_max'], item['box'][1][0])
            current_q['y_max'] = max(current_q['y_max'], item['box'][2][1])

    if current_q: questions.append(current_q)

    # 空间距离算法：将图片归入最近的题目
    for img in cv_images:
        best_q, min_dist = None, float('inf')
        for q in questions:
            q_cy = (q['y_min'] + q['y_max']) / 2
            dist = abs(img['y_center'] - q_cy) # 垂直距离优先
            if dist < min_dist:
                min_dist, best_q = dist, q
        if best_q and min_dist < 500:
            best_q['matched_imgs'].append(img['path'])
            
    return questions

# ==========================================
# 3. 文档处理路由 (支持 Word 图片提取优化)
# ==========================================
def process_uploaded_files(uploaded_files, temp_dir):
    all_questions = []
    for file in uploaded_files:
        file_bytes = file.read()
        file_ext = file.name.split('.')[-1].lower()

        if file_ext in ['jpg', 'jpeg', 'png']:
            path = os.path.join(temp_dir, file.name)
            with open(path, "wb") as f: f.write(file_bytes)
            all_questions.extend(smart_ocr_and_split(path, crop_diagrams(path, temp_dir)))

        elif file_ext == 'pdf':
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for i in range(len(doc)):
                pix = doc[i].get_pixmap(matrix=fitz.Matrix(2, 2))
                img_path = os.path.join(temp_dir, f"p_{i}.jpg")
                pix.save(img_path)
                all_questions.extend(smart_ocr_and_split(img_path, crop_diagrams(img_path, temp_dir)))

        elif file_ext == 'docx':
            doc = docx.Document(io.BytesIO(file_bytes))
            current_q = None
            for para in doc.paragraphs:
                text = para.text.strip()
                if re.match(r'^\s*\d+[\.．、](?!\d)', text):
                    if current_q: all_questions.append(current_q)
                    current_q = {'text': text, 'matched_imgs': []}
                elif current_q and text:
                    current_q['text'] += "\n" + text
                
                # --- Word 图片提取关键逻辑 ---
                if current_q:
                    for run in para.runs:
                        if 'pic:pic' in run._element.xml:
                            rIds = re.findall(r'r:embed="([^"]+)"', run._element.xml)
                            for rId in rIds:
                                try:
                                    img_part = doc.part.related_parts[rId]
                                    img_path = os.path.join(temp_dir, f"w_img_{rId}.png")
                                    with open(img_path, "wb") as f: f.write(img_part.blob)
                                    current_q['matched_imgs'].append(img_path)
                                except: pass
            if current_q: all_questions.append(current_q)
    return all_questions

# ==========================================
# 4. PPT 渲染排版引擎
# ==========================================
def set_font(run, name='微软雅黑'):
    run.font.name = name
    r = run._r.get_or_add_rPr().find(qn('w:rFonts'))
    if r is None: r = run._r.get_or_add_rPr().makeelement(qn('w:rFonts'))
    r.set(qn('w:eastAsia'), name)

def create_slide(prs, title, q_text, imgs, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(248, 250, 253); bg.line.fill.background()
    
    # 标题栏
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.3), Inches(0.3), Inches(0.1), Inches(0.5))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(10), Inches(0.6))
    p = tb.text_frame.paragraphs[0]; p.text = f"习题精讲 - 第 {idx} 题"; p.font.size = Pt(24); p.font.bold = True; set_font(p.runs[0])

    # 左右布局逻辑
    has_img = len(imgs) > 0
    box_w = Inches(8.5) if has_img else Inches(12.5)
    
    # 题目卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.2), box_w, Inches(5.8))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 220, 225)
    
    tf = slide.shapes.add_textbox(Inches(0.6), Inches(1.4), box_w - Inches(0.4), Inches(5.4)).text_frame
    tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    p = tf.paragraphs[0]; p.text = q_text; p.font.size = Pt(18); p.line_spacing = 1.4; set_font(p.runs[0])

    # 图片渲染
    if has_img:
        for i, img_p in enumerate(imgs[:2]): # 每页最多放2张图
            try: slide.shapes.add_picture(img_p, Inches(9.2), Inches(1.2 + i*3.0), width=Inches(3.8))
            except: pass

def make_master_ppt(questions):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    for i, q in enumerate(questions):
        create_slide(prs, "", q['text'], q['matched_imgs'], i+1)
    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer

# ==========================================
# 5. Streamlit 主程序 UI
# ==========================================
st.set_page_config(page_title="AI 物理教研全自动工作站", layout="centered", page_icon="⚛️")

st.markdown("""
<div style='text-align: center; margin-bottom: 30px;'>
    <h1 style='color: #0070C0;'>🚀 AI 物理教研全自动工作站</h1>
    <p style='color: #666;'>支持 <b>图片 / PDF / Word</b> 混合上传，自动提取公式插图并生成巅峰排版 PPT。</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📥 拖拽上传资料（可多选）", accept_multiple_files=True, type=['jpg', 'png', 'pdf', 'docx'])

if st.button("✨ 一键生成精美课件", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("⚠️ 老师，请先上传文件哦！")
    else:
        progress_bar = st.progress(0)
        status = st.empty()
        with tempfile.TemporaryDirectory() as temp_dir:
            status.info("⚙️ 正在运行 OCR 与深度文档解析引擎...")
            questions = process_uploaded_files(uploaded_files, temp_dir)
            progress_bar.progress(60)
            
            if not questions:
                st.error("❌ 未能识别到有效题目，请检查文件清晰度。")
            else:
                status.info(f"✅ 成功提取 {len(questions)} 道题目，正在渲染 PPT...")
                ppt_io = make_master_ppt(questions)
                st.session_state['ready_ppt'] = ppt_io.getvalue()
                progress_bar.progress(100)
                status.success("🎉 课件生成成功！")
                st.balloons()

# 下载区域
if 'ready_ppt' in st.session_state:
    st.write("---")
    st.download_button(
        label="⬇️ 点击下载生成的 PPT 课件",
        data=st.session_state['ready_ppt'],
        file_name=f"AI物理教研课件_{int(time.time())}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
