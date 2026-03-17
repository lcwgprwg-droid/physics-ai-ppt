import os
import re
import cv2
import math
import numpy as np
import tempfile
import io
import time
import gc  # 【新增】极限内存回收模块
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
from PIL import Image, ImageEnhance, ImageDraw

# ==========================================
# 核心引擎库 (视觉 + OCR 逆向遮罩算法)
# ==========================================

def merge_rects(rects, x_margin=40, y_margin=40):
    if not rects: return[]
    boxes = [[r[0], r[1], r[0]+r[2], r[1]+r[3]] for r in rects] 

    def is_close(b1, b2):
        return not (b1[2] < b2[0] - x_margin or b1[0] > b2[2] + x_margin or 
                    b1[3] < b2[1] - y_margin or b1[1] > b2[3] + y_margin)

    merged =[]
    while boxes:
        box = boxes.pop(0)
        has_merged = True
        while has_merged:
            has_merged = False
            for i in range(len(boxes)-1, -1, -1):
                other = boxes[i]
                if is_close(box, other):
                    box = [min(box[0], other[0]), min(box[1], other[1]),
                           max(box[2], other[2]), max(box[3], other[3])]
                    boxes.pop(i)
                    has_merged = True
        merged.append(box)
    return [(b[0], b[1], b[2]-b[0], b[3]-b[1]) for b in merged]

def extract_images_with_ocr_mask(img_path, ocr_result, out_dir):
    img = cv2.imread(img_path)
    if img is None: return[]

    # 【防 OOM 优化】：处理完中间变量后立刻用 del 删除，释放内存
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    mask = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 15, 4)
    del gray  # 释放内存

    if ocr_result:
        for line in ocr_result:
            box = line[0]
            xs, ys = [int(pt[0]) for pt in box], [int(pt[1]) for pt in box]
            x_min, x_max = max(0, min(xs)), min(img.shape[1], max(xs))
            y_min, y_max = max(0, min(ys)), min(img.shape[0], max(ys))
            
            pad = 4
            cv2.rectangle(mask, (max(0, x_min-pad), max(0, y_min-pad)), 
                          (min(img.shape[1], x_max+pad), min(img.shape[0], y_max+pad)), (0, 0, 0), -1)

    kernel_dilate = np.ones((12, 12), np.uint8)
    dilated = cv2.dilate(mask, kernel_dilate, iterations=2)
    del mask  # 释放内存

    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    del dilated  # 释放内存
    gc.collect() # 强制执行一次垃圾回收，防止服务器崩溃

    diagram_boxes =[]
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > 30 and h > 30 and (w * h) > 1200:
            if w < img.shape[1] * 0.85 and h < img.shape[0] * 0.85:
                diagram_boxes.append((x, y, w, h))

    merged_boxes = merge_rects(diagram_boxes, x_margin=45, y_margin=45)

    saved_imgs =[]
    for i, (x, y, w, h) in enumerate(merged_boxes):
        pad_x, pad_y = 25, 25
        x1, y1 = max(0, x - pad_x), max(0, y - pad_y)
        x2, y2 = min(img.shape[1], x + w + pad_x), min(img.shape[0], y + h + pad_y)
        roi = img[y1:y2, x1:x2]

        p = os.path.join(out_dir, f"fig_{int(time.time()*1000)}_{i}.png")
        roi_rgb = cv2.cvtColor(roi, cv2.COLOR_BGR2RGB)
        pil_img = ImageEnhance.Sharpness(ImageEnhance.Contrast(Image.fromarray(roi_rgb)).enhance(1.2)).enhance(1.8)
        pil_img.save(p, quality=95)
        saved_imgs.append({"path": p, "x_center": x + w/2, "y_center": y + h/2})

    del img
    gc.collect()
    return saved_imgs

def post_process_text(text):
    text = re.sub(r'光敏电阻符号是[^\，。]*(?=，|。|$)', '光敏电阻符号是[             ] ', text)
    text = re.sub(r'电磁开关符号是[^\，。]*(?=，|。|$)', '电磁开关符号是[             ] ', text)
    return text

def is_noise(text):
    noise_words =['复习与提高', 'A组', 'B组', '高中物理', '必修', '选择性', '扫描全能王']
    for w in noise_words:
        if w in text: return True
    if re.match(r'^\s*\d+\s*$', text): return True
    return False

def smart_ocr_and_split(img_path, temp_dir):
    engine = RapidOCR()
    result, _ = engine(img_path)
    gc.collect() # 识别完文字后释放模型内存
    
    cv_images = extract_images_with_ocr_mask(img_path, result, temp_dir)
    if not result: return[]

    img_cv = cv2.imread(img_path)
    page_center_x = img_cv.shape[1] / 2
    del img_cv
    gc.collect()

    sorted_lines =[]
    for line in result:
        box, text = line[0], line[1]
        cx, cy = (box[0][0] + box[1][0]) / 2, (box[0][1] + box[3][1]) / 2
        col_idx = 0 if cx < page_center_x else 1
        sorted_lines.append({"col": col_idx, "y": cy, "box": box, "text": text})

    sorted_lines.sort(key=lambda item: (item['col'], item['y']))
    questions, current_q =[], None

    for item in sorted_lines:
        text = item['text'].strip()
        box = item['box']
        x_left, x_right, y_top, y_bottom = box[0][0], box[1][0], box[0][1], box[2][1]

        if is_noise(text): continue
        if re.match(r'^\s*\d+[\.．、](?!\d)', text):
            if current_q: questions.append(current_q)
            current_q = {'text': text, 'x_min': x_left, 'x_max': x_right, 'y_min': y_top, 'y_max': y_bottom, 'matched_imgs':[]}
        else:
            if current_q is None: continue
            if re.match(r'^\s*([A-D][\.．、]|\(\d+\)|①|②|③)', text): current_q['text'] += '\n' + text
            else: current_q['text'] += text
            current_q['x_min'], current_q['x_max'] = min(current_q['x_min'], x_left), max(current_q['x_max'], x_right)
            current_q['y_min'], current_q['y_max'] = min(current_q['y_min'], y_top), max(current_q['y_max'], y_bottom)

    if current_q: questions.append(current_q)
    
    for q in questions:
        q['text'] = post_process_text(q['text'])
        q_cx = (q['x_min'] + q['x_max']) / 2
        q['col'] = 0 if q_cx < page_center_x else 1

    for img in cv_images:
        img_col = 0 if img['x_center'] < page_center_x else 1
        col_questions =[q for q in questions if q['col'] == img_col]
        
        best_q = None
        img_y = img['y_center']

        for i, q in enumerate(col_questions):
            q_top = q['y_min']
            q_bottom = col_questions[i+1]['y_min'] if i < len(col_questions) - 1 else float('inf')
            if q_top - 60 <= img_y <= q_bottom + 60:
                best_q = q
                break
                
        if not best_q and col_questions:
            best_q = min(col_questions, key=lambda q: abs(q['y_max'] - img_y))
            
        if best_q:
            best_q['matched_imgs'].append(img['path'])

    return questions

# ==========================================
# 文档路由引擎
# ==========================================
def process_uploaded_files(uploaded_files, temp_dir):
    all_questions =[]
    for file in uploaded_files:
        file_bytes = file.read()
        file_ext = file.name.split('.')[-1].lower()

        if file_ext in['jpg', 'jpeg', 'png']:
            temp_img_path = os.path.join(temp_dir, file.name)
            with open(temp_img_path, "wb") as f: f.write(file_bytes)
            qs = smart_ocr_and_split(temp_img_path, temp_dir)
            all_questions.extend(qs)

        elif file_ext == 'pdf':
            pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page_num in range(len(pdf_doc)):
                # 【防 OOM 优化】：将清晰度矩阵降为 1.2，大幅降低内存占用，画质依然足够
                pix = pdf_doc.load_page(page_num).get_pixmap(matrix=fitz.Matrix(1.2, 1.2))
                temp_img_path = os.path.join(temp_dir, f"pdf_{file.name}_page_{page_num}.jpg")
                pix.save(temp_img_path)
                pix = None # 释放对象
                gc.collect()
                
                qs = smart_ocr_and_split(temp_img_path, temp_dir)
                all_questions.extend(qs)

        elif file_ext == 'docx':
            doc = docx.Document(io.BytesIO(file_bytes))
            current_q = None
            for para in doc.paragraphs:
                text = para.text.strip()
                if not text: continue
                if re.match(r'^\s*\d+[\.．、](?!\d)', text):
                    if current_q: all_questions.append(current_q)
                    current_q = {'text': text, 'matched_imgs':[]}
                else:
                    if current_q: current_q['text'] += '\n' + text
            if current_q: all_questions.append(current_q)

    return all_questions

# ==========================================
# PPT 渲染排版引擎 (智能多图排版布局)
# ==========================================
def set_font(run, font_name='微软雅黑'):
    run.font.name = font_name
    rPr = run._r.get_or_add_rPr()
    f = rPr.find(qn('w:rFonts'))
    if f is None:
        f = rPr.makeelement(qn('w:rFonts'))
        rPr.append(f)
    f.set(qn('w:eastAsia'), font_name)

def create_base_slide(prs, title_text):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11.5), Inches(0.8))
    p = tb.text_frame.paragraphs[0]; p.text = title_text
    p.font.bold = True; p.font.size = Pt(26); p.font.color.rgb = RGBColor(30, 40, 60); set_font(p.runs[0])
    return slide

def add_badge_card(slide, x, y, w, h, badge_text, badge_color, content, font_size, line_spacing=1.3):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(230, 230, 235)
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.2), y + Inches(0.15), Inches(1.2), Inches(0.35))
    badge.fill.solid(); badge.fill.fore_color.rgb = badge_color; badge.line.fill.background()
    bp = badge.text_frame.paragraphs[0]; bp.text = badge_text; bp.font.bold = True; bp.font.size = Pt(12)
    bp.font.color.rgb = RGBColor(255, 255, 255); bp.alignment = PP_ALIGN.CENTER; set_font(bp.runs[0])
    
    tb = slide.shapes.add_textbox(x + Inches(0.1), y + Inches(0.55), w - Inches(0.2), h - Inches(0.6))
    tf = tb.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for i, line in enumerate(content.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line; p.font.size = Pt(font_size); p.font.color.rgb = RGBColor(30, 30, 30); p.line_spacing = line_spacing; set_font(p.runs[0])

def render_images_on_slide(slide, img_paths):
    if not img_paths: return
    num_imgs = len(img_paths)
    start_y = 1.2
    
    target_width = 4.0 if num_imgs == 1 else (3.8 if num_imgs == 2 else 3.0)

    for img_path in img_paths:
        try:
            pic = slide.shapes.add_picture(img_path, Inches(8.9), Inches(start_y), width=Inches(target_width))
            pic_height_in = pic.height / 914400.0 
            
            if start_y + pic_height_in > 7.3:
                pic.height = int((7.3 - start_y) * 914400)
                pic.width = int(pic.height * (Image.open(img_path).width / Image.open(img_path).height))
                pic_height_in = pic.height / 914400.0
                
            start_y += pic_height_in + 0.15 
        except Exception as e:
            print(f"Error rendering image: {e}")

def make_master_ppt(questions_data):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    c_blue, c_orange = RGBColor(0, 112, 192), RGBColor(230, 90, 40)
    global_idx = 1

    for q in questions_data:
        matched_imgs = q.get('matched_imgs',[])
        has_imgs = len(matched_imgs) > 0
        text_w = Inches(8.3) if has_imgs else Inches(12.0)
        visual_lines = q['text'].count('\n') + (len(q['text']) / 32)
        q_font, q_space = (14, 1.2) if visual_lines > 15 else ((16, 1.3) if visual_lines > 10 else (18, 1.4))

        if len(q['text']) > 120:
            slide1 = create_base_slide(prs, f"习题精讲 - 第 {global_idx} 题")
            add_badge_card(slide1, Inches(0.4), Inches(1.2), text_w, Inches(5.8), "原题呈现", c_blue, q['text'], q_font, q_space)
            render_images_on_slide(slide1, matched_imgs)
            
            slide2 = create_base_slide(prs, f"习题精讲 - 第 {global_idx} 题 (解析)")
            add_badge_card(slide2, Inches(0.4), Inches(1.2), text_w, Inches(5.8), "深度解析", c_orange, "待补充解析过程...", 18, 1.4)
            render_images_on_slide(slide2, matched_imgs)
        else:
            slide = create_base_slide(prs, f"习题精讲 - 第 {global_idx} 题")
            add_badge_card(slide, Inches(0.4), Inches(1.2), text_w, Inches(4.0), "原题呈现", c_blue, q['text'], q_font, q_space)
            add_badge_card(slide, Inches(0.4), Inches(5.4), text_w, Inches(1.8), "深度解析", c_orange, "待补充解析过程...", 16, 1.3)
            render_images_on_slide(slide, matched_imgs)
            
        global_idx += 1

    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer

# ==========================================
# Streamlit 现代化 Web UI (防OOM内存崩溃 + 终极状态机)
# ==========================================
st.set_page_config(page_title="AI 物理/生物教研课件工作站", layout="centered", page_icon="⚛️")

# 【核心机制】：工业级 UI 状态机，绝不丢按钮
if 'app_state' not in st.session_state:
    st.session_state['app_state'] = 'idle'
if 'ready_ppt' not in st.session_state:
    st.session_state['ready_ppt'] = None

st.markdown("""
<div style='text-align: center; margin-bottom: 30px;'>
    <h1 style='color: #0070C0;'>🚀 AI 教研全自动工作站</h1>
    <p style='color: #666;'>支持同时上传多张 <b>教辅照片 / PDF / Word文档</b>，一键生成带有视觉配图与自动分页的巅峰排版 PPT。</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "📥 拖拽上传题库资料（支持双栏排版与高强度图片识别）",
    accept_multiple_files=True,
    type=['jpg', 'jpeg', 'png', 'pdf', 'docx']
)

# 按钮只负责“修改状态”，不嵌套任何复杂逻辑！
if st.button("✨ 一键生成精美 PPT", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("⚠️ 老师，请先上传文件哦！")
    else:
        st.session_state['app_state'] = 'processing'
        # 强制网页刷新一次，脱离按钮嵌套
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()

# 后台静默处理区
if st.session_state['app_state'] == 'processing':
    progress_bar = st.progress(0)
    status_text = st.empty()
    status_text.info("⚙️ 启动深度视觉与OCR引擎... (开启内存保护模式，请稍候)")
    
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            final_questions = process_uploaded_files(uploaded_files, temp_dir)
            progress_bar.progress(60)

            if not final_questions:
                status_text.error("❌ 抱歉，未能识别到有效题目，请检查图片或PDF内容。")
                st.session_state['app_state'] = 'error'
            else:
                status_text.info(f"✅ 成功提取 {len(final_questions)} 道大题！正在渲染排版...")
                ppt_io = make_master_ppt(final_questions)
                
                # 保存数据，并标记成功
                st.session_state['ready_ppt'] = ppt_io.getvalue()
                st.session_state['app_state'] = 'success'
                progress_bar.progress(100)
                
        # 处理完成后，再次强制网页刷新，让画面直接跳到下载按钮！
        try:
            st.rerun()
        except AttributeError:
            st.experimental_rerun()
            
    except Exception as e:
        progress_bar.empty()
        status_text.error(f"❌ 运行过程中发生崩溃：")
        import traceback
        st.code(traceback.format_exc())
        st.session_state['app_state'] = 'error'

# 最外层独立挂载下载区，只要状态是 success，按钮永远存在！
if st.session_state['app_state'] == 'success' and st.session_state['ready_ppt']:
    st.success("🎉 大功告成！不仅精准剥离了图片，还完美匹配了多图！课件已生成！")
    st.download_button(
        label="⬇️ 点击这里下载生成的 PPT 课件",
        data=st.session_state['ready_ppt'],
        file_name="核心素养习题精讲(AI生成版).pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
