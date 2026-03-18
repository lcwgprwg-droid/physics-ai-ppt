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
# 1. 核心常量
# ==========================================
PPT_WIDTH = Inches(13.333)
PPT_HEIGHT = Inches(7.5)
C_BLACK = RGBColor(0, 0, 0)
C_BLUE = RGBColor(0, 112, 192)
C_WHITE = RGBColor(255, 255, 255)

@st.cache_resource
def load_ocr():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：物理图示捕捉 (V25)
# ==========================================
def extract_visuals_physics(img_path, out_dir):
    """
    捕捉全图视觉块，不丢受力图和公式。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 横向连接公式
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 10))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    visuals = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.9 or h > h_img * 0.9: continue
        if w < 30 or h < 20: continue 
        
        roi = img[y:y+h, x:x+w]
        if np.mean(roi) > 252: continue
        
        f_path = os.path.join(out_dir, f"vis_{int(time.time()*1000)}_{x}.png")
        cv2.imwrite(f_path, roi)
        visuals.append({"path": f_path, "y": y + h/2, "area": w * h})
    return visuals

# ==========================================
# 3. PPT 渲染引擎：【一题一页·强制对齐】
# ==========================================
def render_one_slide_per_q(prs, q_data, idx):
    """
    严格一题一个 Slide，彻底解决重叠。
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 1. 灰色背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
    
    # 2. 标题 (左对齐)
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.8))
    p_t = title_box.text_frame.paragraphs[0]
    p_t.alignment = PP_ALIGN.LEFT
    run_t = p_t.add_run()
    run_t.text = f"习题精讲 第 {idx + 1} 题"
    run_t.font.size, run_t.font.bold = Pt(26), True
    run_t.font.color.rgb = RGBColor(20, 40, 80)
    
    has_imgs = len(q_data.get('imgs', [])) > 0
    txt_w = Inches(8.5) if has_imgs else Inches(12.3)

    # 3. 题干卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.3))
    card.fill.solid(); card.fill.fore_color.rgb = C_WHITE; card.line.color.rgb = RGBColor(220, 220, 225)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    # 4. 填充题干文字 (强制黑色左对齐)
    lines = q_data.get('text', '').split('\n')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.3
        r = p.add_run()
        r.text = line.strip()
        r.font.name = '微软雅黑'
        r.font.size = Pt(18)
        r.font.color.rgb = C_BLACK
        try:
            rPr = r._r.get_or_add_rPr()
            rPr.get_or_add_ea().set('typeface', '微软雅黑')
        except: pass

    # 5. 解析卡片
    card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.8), txt_w, Inches(1.4))
    card2.fill.solid(); card2.fill.fore_color.rgb = C_WHITE; card2.line.color.rgb = RGBColor(220, 220, 225)
    p2 = card2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run(); r2.text = "解析预留：正在整理受力分析与解题步骤..."; r2.font.color.rgb = RGBColor(100, 100, 100); r2.font.size = Pt(16)

    # 6. 图片投放
    if has_imgs:
        y_cursor = 1.3
        # 挑选面积最大的 3 张
        sorted_imgs = sorted(q_data['imgs'], key=lambda x: x['area'], reverse=True)[:3]
        # 再按原始高度 Y 轴排序
        sorted_imgs = sorted(sorted_imgs, key=lambda x: x['y'])
        for img_info in sorted_imgs:
            try:
                slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_cursor), width=Inches(3.8))
                y_cursor += 2.0 
            except: pass

# ==========================================
# 4. 业务流
# ==========================================
st.set_page_config(page_title="物理题 PPT 专家", layout="centered")
st.title("⚛️ AI 物理教研 (分页修复版)")

files = st.file_uploader("📥 上传 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 极速生成一题一页 PPT", type="primary", use_container_width=True):
    if not files:
        st.error("老师，请上传文件。")
    else:
        all_final_qs = []
        engine = load_ocr()
        
        with st.status("🔍 深度解析中，请勿刷新...", expanded=True) as status:
            with tempfile.TemporaryDirectory() as tmpdir:
                for file in files:
                    ext = file.name.split('.')[-1].lower()
                    if ext == 'docx':
                        # Word 解析：严格分页
                        doc = Document(io.BytesIO(file.read()))
                        cur_q = None
                        for p in doc.paragraphs:
                            t = p.text.strip()
                            if not t: continue
                            # 题号分页逻辑
                            if re.match(r'^(\d+|[\(（]\d+[\)）])[\.．、\s]', t) or len(t) < 8:
                                if cur_q: all_final_qs.append(cur_q)
                                cur_q = {"text": t, "imgs": []}
                            elif cur_q:
                                cur_q["text"] += "\n" + t
                            else:
                                cur_q = {"text": t, "imgs": []}
                        if cur_q: all_final_qs.append(cur_q)
                    
                    else:
                        # 视觉解析逻辑
                        imgs_to_proc = []
                        if ext == 'pdf':
                            pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
                            for page in pdf_doc:
                                pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                                p_path = os.path.join(tmpdir, f"p_{time.time()}.png")
                                pix.save(p_path); imgs_to_proc.append(p_path)
                        else:
                            p_path = os.path.join(tmpdir, file.name)
                            with open(p_path, "wb") as f: f.write(file.read())
                            imgs_to_proc.append(p_path)
                        
                        for p_path in imgs_to_proc:
                            ocr_res, _ = engine(p_path)
                            visual_elements = extract_visuals_physics(p_path, tmpdir)
                            
                            # 核心：根据题号 Y 轴切分
                            page_qs = []
                            cur_q = None
                            if ocr_res:
                                for line in ocr_res:
                                    txt = line[1].strip()
                                    y_mid = (line[0][0][1] + line[0][2][1]) / 2
                                    # 题号识别触发分页
                                    if re.match(r'^(\d+|[\(（]\d+[\)）])', txt):
                                        if cur_q: page_qs.append(cur_q)
                                        cur_q = {"text": txt, "imgs": [], "y_start": y_mid, "y_end": 9999}
                                        if page_qs: page_qs[-1]["y_end"] = y_mid
                                    elif cur_q:
                                        cur_q["text"] += "\n" + txt
                                    else: # 兜底逻辑：第一行
                                        cur_q = {"text": txt, "imgs": [], "y_start": y_mid, "y_end": 9999}
                            if cur_q: page_qs.append(cur_q)
                            
                            # 图片精准吸附到 Y 轴区间
                            for v in visual_elements:
                                for q in page_qs:
                                    if q["y_start"] <= v["y"] < q["y_end"]:
                                        q["imgs"].append(v)
                                        break
                            all_final_qs.extend(page_qs)

                if all_final_qs:
                    status.update(label="✅ 解析完成，正在渲染独立幻灯片...", state="running")
                    prs = Presentation()
                    prs.slide_width, prs.slide_height = PPT_WIDTH, PPT_HEIGHT
                    
                    for i, q in enumerate(all_final_qs):
                        render_one_slide_per_q(prs, q, i)
                    
                    ppt_buf = io.BytesIO()
                    prs.save(ppt_buf)
                    st.download_button("📥 点击下载分页纠正版 PPT", ppt_buf.getvalue(), "物理教研专家课件.pptx", use_container_width=True)
                    status.update(label="🎉 转换成功！一题一页已锁定。", state="complete")
                else:
                    st.error("未识别到题目，请确认文件是否清晰。")
