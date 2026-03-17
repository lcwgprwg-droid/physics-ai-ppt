import os
import time
import cv2
import re
import math
import tempfile
import io
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
# 核心引擎库 (继承我们最巅峰的 视觉 + OCR 排版引擎)
# ==========================================
def crop_diagrams(img_path, out_dir):
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
        if h > 50 and w > 50 and (w * h) > 8000 and 0.1 < (w / float(h)) < 8.0:
            y1, y2 = max(0, y - 15), min(img.shape[0], y + h + 15)
            x1, x2 = max(0, x - 15), min(img.shape[1], x + w + 15)
            diagrams.append({"img": img[y1:y2, x1:x2], "x_center": x + w / 2, "y_center": y + h / 2})

    saved_imgs = []
    for i, d in enumerate(diagrams):
        p = os.path.join(out_dir, f"fig_{int(time.time() * 1000)}_{i}.png")
        roi_rgb = cv2.cvtColor(d["img"], cv2.COLOR_BGR2RGB)
        mask_bg = (roi_rgb[:, :, 0] > 215) & (roi_rgb[:, :, 1] > 215) & (roi_rgb[:, :, 2] > 215)
        roi_rgb[mask_bg] = [255, 255, 255]
        pil_img = ImageEnhance.Sharpness(ImageEnhance.Contrast(Image.fromarray(roi_rgb)).enhance(1.2)).enhance(1.8)
        pil_img.save(p, quality=95)
        saved_imgs.append({"path": p, "x_center": d["x_center"], "y_center": d["y_center"]})
    return saved_imgs


def post_process_text(text):
    text = re.sub(r'光敏电阻符号是[^\，。]*(?=，|。|$)', '光敏电阻符号是[             ] ', text)
    text = re.sub(r'电磁开关符号是[^\，。]*(?=，|。|$)', '电磁开关符号是 [             ] ', text)
    return text


def is_noise(text):
    noise_words = ['复习与提高', 'A组', 'B组', '高中物理', '必修', '选择性', '扫描全能王']
    for w in noise_words:
        if w in text: return True
    if re.match(r'^\s*\d+\s*$', text): return True
    return False


def smart_ocr_and_split(img_path, cv_images):
    engine = RapidOCR()
    result, _ = engine(img_path)
    if not result: return []

    max_x = max([line[0][1][0] for line in result])
    page_center_x = max_x / 2

    sorted_lines = []
    for line in result:
        box, text = line[0], line[1]
        cx, cy = (box[0][0] + box[1][0]) / 2, (box[0][1] + box[3][1]) / 2
        col_idx = 0 if cx < page_center_x else 1
        sorted_lines.append({"col": col_idx, "y": cy, "box": box, "text": text})

    sorted_lines.sort(key=lambda item: (item['col'], item['y']))

    questions = []
    current_q = None

    for item in sorted_lines:
        text = item['text'].strip()
        box = item['box']
        x_left, x_right, y_top, y_bottom = box[0][0], box[1][0], box[0][1], box[2][1]

        if is_noise(text): continue

        if re.match(r'^\s*\d+[\.．、](?!\d)', text):
            if current_q: questions.append(current_q)
            current_q = {'text': text, 'x_min': x_left, 'x_max': x_right, 'y_min': y_top, 'y_max': y_bottom,
                         'matched_imgs': []}
        else:
            if current_q is None: continue
            if re.match(r'^\s*([A-D][\.．、]|\(\d+\)|①|②|③)', text):
                current_q['text'] += '\n' + text
            else:
                current_q['text'] += text
            current_q['x_min'], current_q['x_max'] = min(current_q['x_min'], x_left), max(current_q['x_max'], x_right)
            current_q['y_min'], current_q['y_max'] = min(current_q['y_min'], y_top), max(current_q['y_max'], y_bottom)

    if current_q: questions.append(current_q)

    for q in questions:
        q['text'] = post_process_text(q['text'])

    for img in cv_images:
        best_q = None
        min_dist = float('inf')
        for q in questions:
            q_cx, q_cy = (q['x_min'] + q['x_max']) / 2, (q['y_min'] + q['y_max']) / 2
            dist = math.sqrt((img['x_center'] - q_cx) ** 2 + (img['y_center'] - q_cy) ** 2)
            if dist < min_dist:
                min_dist = dist
                best_q = q
        if best_q and min_dist < 600:
            best_q['matched_imgs'].append(img['path'])

    return questions


# ==========================================
# 文档解析路由器 (处理 多图 / PDF / DOCX)
# ==========================================
def process_uploaded_files(uploaded_files, temp_dir):
    all_questions = []

    for file in uploaded_files:
        file_bytes = file.read()
        file_ext = file.name.split('.')[-1].lower()

        # 1. 处理图片 (JPG/PNG)
        if file_ext in ['jpg', 'jpeg', 'png']:
            temp_img_path = os.path.join(temp_dir, file.name)
            with open(temp_img_path, "wb") as f:
                f.write(file_bytes)

            cv_imgs = crop_diagrams(temp_img_path, temp_dir)
            qs = smart_ocr_and_split(temp_img_path, cv_imgs)
            all_questions.extend(qs)

        # 2. 处理 PDF (将每一页转为高清图片喂给视觉引擎)
        elif file_ext == 'pdf':
            pdf_doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page_num in range(len(pdf_doc)):
                page = pdf_doc.load_page(page_num)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2倍高清渲染
                temp_img_path = os.path.join(temp_dir, f"pdf_{file.name}_page_{page_num}.jpg")
                pix.save(temp_img_path)

                cv_imgs = crop_diagrams(temp_img_path, temp_dir)
                qs = smart_ocr_and_split(temp_img_path, cv_imgs)
                all_questions.extend(qs)

        # 3. 处理 Word 文档 (DOCX) - 直接提取文本
        elif file_ext == 'docx':
            doc = docx.Document(io.BytesIO(file_bytes))
            current_q = None
            for para in doc.paragraphs:
                text = para.text.strip()
                if not text: continue
                # 如果是新题
                if re.match(r'^\s*\d+[\.．、](?!\d)', text):
                    if current_q: all_questions.append(current_q)
                    current_q = {'text': text, 'matched_imgs': []}  # Word 暂不处理内联图片匹配
                else:
                    if current_q: current_q['text'] += '\n' + text
            if current_q: all_questions.append(current_q)

    return all_questions


# ==========================================
# PPT 渲染引擎 (卡片式排版)
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
    bg.fill.solid();
    bg.fill.fore_color.rgb = RGBColor(245, 247, 250);
    bg.line.fill.background()
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
    bar.fill.solid();
    bar.fill.fore_color.rgb = RGBColor(0, 112, 192);
    bar.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11.5), Inches(0.8))
    p = tb.text_frame.paragraphs[0];
    p.text = title_text
    p.font.bold = True;
    p.font.size = Pt(26);
    p.font.color.rgb = RGBColor(30, 40, 60);
    set_font(p.runs[0])
    return slide


def add_badge_card(slide, x, y, w, h, badge_text, badge_color, content, font_size, line_spacing=1.3):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid();
    card.fill.fore_color.rgb = RGBColor(255, 255, 255);
    card.line.color.rgb = RGBColor(230, 230, 235)

    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.2), y + Inches(0.15), Inches(1.2),
                                   Inches(0.35))
    badge.fill.solid();
    badge.fill.fore_color.rgb = badge_color;
    badge.line.fill.background()
    bp = badge.text_frame.paragraphs[0];
    bp.text = badge_text;
    bp.font.bold = True;
    bp.font.size = Pt(12)
    bp.font.color.rgb = RGBColor(255, 255, 255);
    bp.alignment = PP_ALIGN.CENTER;
    set_font(bp.runs[0])

    tb = slide.shapes.add_textbox(x + Inches(0.1), y + Inches(0.55), w - Inches(0.2), h - Inches(0.6))
    tf = tb.text_frame;
    tf.word_wrap = True;
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

    for i, line in enumerate(content.split('\n')):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line;
        p.font.size = Pt(font_size);
        p.font.color.rgb = RGBColor(30, 30, 30);
        p.line_spacing = line_spacing;
        set_font(p.runs[0])


def make_master_ppt(questions_data):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    c_blue, c_orange = RGBColor(0, 112, 192), RGBColor(230, 90, 40)

    # 全局题号计数器（因为可能上传了多个文件）
    global_idx = 1

    for q in questions_data:
        has_imgs = len(q.get('matched_imgs', [])) > 0
        text_w = Inches(8.3) if has_imgs else Inches(12.0)

        visual_lines = q['text'].count('\n') + (len(q['text']) / 32)
        q_font, q_space = (14, 1.2) if visual_lines > 15 else ((16, 1.3) if visual_lines > 10 else (18, 1.4))

        if len(q['text']) > 120:
            slide1 = create_base_slide(prs, f"习题精讲 - 第 {global_idx} 题")
            add_badge_card(slide1, Inches(0.4), Inches(1.2), text_w, Inches(5.8), "原题呈现", c_blue, q['text'], q_font,
                           q_space)
            if has_imgs:
                start_y = 1.2
                for img_path in q['matched_imgs']:
                    slide1.shapes.add_picture(img_path, Inches(8.9), Inches(start_y), width=Inches(4.0))
                    start_y += 3.2

            slide2 = create_base_slide(prs, f"习题精讲 - 第 {global_idx} 题 (解析)")
            add_badge_card(slide2, Inches(0.4), Inches(1.2), text_w, Inches(5.8), "深度解析", c_orange,
                           "待补充解析过程...", 18, 1.4)
            if has_imgs:
                start_y = 1.2
                for img_path in q['matched_imgs']:
                    slide2.shapes.add_picture(img_path, Inches(8.9), Inches(start_y), width=Inches(4.0))
                    start_y += 3.2
        else:
            slide = create_base_slide(prs, f"习题精讲 - 第 {global_idx} 题")
            add_badge_card(slide, Inches(0.4), Inches(1.2), text_w, Inches(4.0), "原题呈现", c_blue, q['text'], q_font,
                           q_space)
            add_badge_card(slide, Inches(0.4), Inches(5.4), text_w, Inches(1.8), "深度解析", c_orange,
                           "待补充解析过程...", 16, 1.3)
            if has_imgs:
                start_y = 1.2
                for img_path in q['matched_imgs']:
                    slide.shapes.add_picture(img_path, Inches(8.9), Inches(start_y), width=Inches(4.0))
                    start_y += 3.2

        global_idx += 1

    # 将生成的 PPT 存入内存 Buffer，方便 Web 前端提供下载
    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer


# ==========================================
# Streamlit 现代化 Web UI
# ==========================================
st.set_page_config(page_title="AI 物理教研课件生成器", layout="centered", page_icon="⚛️")

st.markdown("""
<div style='text-align: center; margin-bottom: 30px;'>
    <h1 style='color: #0070C0;'>🚀 AI 物理教研课件全自动工作站</h1>
    <p style='color: #666;'>支持同时上传多张 <b>教辅照片 / PDF / Word文档</b>，一键生成带有视觉配图与自动分页的巅峰排版 PPT。</p>
</div>
""", unsafe_allow_html=True)

# 文件上传组件
uploaded_files = st.file_uploader(
    "📥 拖拽或点击上传你的题库资料（支持 .jpg, .png, .pdf, .docx，可多选）",
    accept_multiple_files=True,
    type=['jpg', 'jpeg', 'png', 'pdf', 'docx']
)

if st.button("✨ 一键生成精美 PPT", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("⚠️ 老师，请先上传至少一份文件哦！")
    else:
        # 使用进度条增强用户体验
        progress_bar = st.progress(0)
        status_text = st.empty()

        # 创建临时文件夹来存储处理过程中的图片
        with tempfile.TemporaryDirectory() as temp_dir:
            status_text.info("⚙️ 正在启动 OCR 与机器视觉引擎，疯狂扫题中...")

            # 1. 路由解析所有上传的文件
            final_questions = process_uploaded_files(uploaded_files, temp_dir)
            progress_bar.progress(60)

            if not final_questions:
                st.error("❌ 抱歉，未能从上传的文件中识别到任何有效的题目结构（请检查是否带有 1. 2. 等题号）。")
            else:
                status_text.info(f"✅ 成功从文件中提取了 {len(final_questions)} 道大题！正在进行智能分屏与 UI 渲染...")

                # 2. 渲染终极 PPT
                ppt_io = make_master_ppt(final_questions)
                progress_bar.progress(100)
                status_text.success("🎉 大功告成！课件已成功生成！")

                # 3. 提供下载按钮
                st.download_button(
                    label="⬇️ 下载生成的 PPT 课件",
                    data=ppt_io,
                    file_name="核心素养习题精讲(AI生成版).pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )

                # 彩蛋：生成个气球庆祝一下
                st.balloons()
