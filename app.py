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
# 1. 核心配置与状态反馈
# ==========================================
@st.cache_resource
def get_ocr_engine():
    # 第一次运行会较慢，属于正常现象
    return RapidOCR()

def log_status(text):
    st.write(f"🔔 {text}")

# ==========================================
# 2. 视觉工程：高兼容性捕获 (V24)
# ==========================================
def extract_physics_vision(img_path, ocr_result, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 建立文字屏蔽，防止文字干扰图示识别
    text_mask = np.zeros((h_img, w_img), dtype=np.uint8)
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            cv2.fillPoly(text_mask, [box], 255)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 21, 8)
    
    # 横向膨胀：连接公式
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 10))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    
    # 排除已知文字区
    clean_dilated = cv2.subtract(dilated, text_mask)
    
    contours, _ = cv2.findContours(clean_dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    visual_elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.9 or h > h_img * 0.9: continue
        if w < 30 or h < 20: continue
        
        roi = img[y:y+h, x:x+w]
        if np.mean(roi) > 252: continue
        
        f_name = f"diag_{int(time.time()*1000)}_{x}.png"
        f_path = os.path.join(out_dir, f_name)
        cv2.imwrite(f_path, roi)
        visual_elements.append({"path": f_path, "y": y + h/2})
            
    return visual_elements

# ==========================================
# 3. PPT 强力排版引擎 (锁定黑色左对齐)
# ==========================================
def render_physics_ppt(all_qs, progress_bar):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    total = len(all_qs)

    for i, q in enumerate(all_qs):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 标题 (锁定黑色)
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(11), Inches(0.8))
        p_t = title_box.text_frame.paragraphs[0]
        p_t.alignment = PP_ALIGN.LEFT
        r_t = p_t.add_run()
        r_t.text = f"习题精讲 第 {i+1} 题"
        r_t.font.size, r_t.font.bold, r_t.font.color.rgb = Pt(26), True, RGBColor(20, 40, 80)
        
        has_imgs = len(q.get('imgs', [])) > 0
        txt_w = Inches(8.5) if has_imgs else Inches(12.3)

        # 题干卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.3))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(210, 210, 215)
        
        tf = card.text_frame
        tf.word_wrap = True
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
        
        # 逐行写入 (强制黑色左对齐)
        lines = q.get('text', '').split('\n')
        for idx, line in enumerate(lines):
            p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
            p.alignment = PP_ALIGN.LEFT
            p.line_spacing = 1.3
            r = p.add_run()
            r.text = line.strip()
            r.font.name = '微软雅黑'
            r.font.size = Pt(18)
            r.font.color.rgb = RGBColor(0, 0, 0) # 绝对黑色
            try:
                rPr = r._r.get_or_add_rPr()
                ea = rPr.get_or_add_ea()
                ea.set('typeface', '微软雅黑')
            except: pass

        # 下方解析
        card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.8), txt_w, Inches(1.4))
        card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(210, 210, 215)
        p2 = card2.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.LEFT
        r2 = p2.add_run(); r2.text = "思路解析：正在整理中..."; r2.font.size = Pt(16); r2.font.color.rgb = RGBColor(100, 100, 100)

        # 投放图示
        if has_imgs:
            y_ptr = 1.3
            q['imgs'].sort(key=lambda x: x['y'])
            for img_info in q['imgs'][:3]:
                try:
                    slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                    y_ptr += 2.0 
                except: pass
        
        progress_bar.progress((i + 1) / total)
                
    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()

# ==========================================
# 4. Streamlit 交互层
# ==========================================
st.set_page_config(page_title="物理教研 AI 工具", layout="centered")
st.title("⚛️ AI 物理教研工作站 (V24 稳定版)")

files = st.file_uploader("📥 上传 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 极速生成 PPT", type="primary", use_container_width=True):
    if not files:
        st.error("请先上传文件")
    else:
        all_final_qs = []
        
        with st.status("正在启动 AI 解析引擎...", expanded=True) as status:
            engine = get_ocr_engine()
            status.update(label="AI 引擎就绪，开始解析文件...", state="running")
            
            with tempfile.TemporaryDirectory() as tmpdir:
                for file in files:
                    st.write(f"📄 正在处理: {file.name}")
                    ext = file.name.split('.')[-1].lower()
                    
                    if ext == 'docx':
                        doc = Document(io.BytesIO(file.read()))
                        cur_q = None
                        for p in doc.paragraphs:
                            txt = p.text.strip()
                            if not txt: continue
                            # 只要是以数字开头的段落，就开启新 slide
                            if re.match(r'^\d+', txt):
                                if cur_q: all_final_qs.append(cur_q)
                                cur_q = {"text": txt, "imgs": []}
                            elif cur_q:
                                cur_q["text"] += "\n" + txt
                            else:
                                cur_q = {"text": txt, "imgs": []}
                        if cur_q: all_final_qs.append(cur_q)
                        
                    else:
                        # 视觉解析 (PDF/图片)
                        imgs = []
                        if ext == 'pdf':
                            pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
                            for page in pdf_doc:
                                pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                                p = os.path.join(tmpdir, f"p_{time.time()}.png")
                                pix.save(p); imgs.append(p)
                        else:
                            p = os.path.join(tmpdir, file.name)
                            with open(p, "wb") as f: f.write(file.read())
                            imgs.append(p)
                        
                        for img_p in imgs:
                            res, _ = engine(img_p)
                            visuals = extract_physics_vision(img_p, res, tmpdir)
                            
                            full_txt = "\n".join([line[1] for line in res]) if res else ""
                            # 物理题分割逻辑：每页 PDF 至少出一个 Slide
                            if full_txt:
                                all_final_qs.append({"text": full_txt, "imgs": visuals, "y": 0})

                if all_final_qs:
                    status.update(label=f"解析完成，共 {len(all_final_qs)} 题。正在排版生成 PPT...", state="running")
                    progress = st.progress(0)
                    ppt_data = render_physics_ppt(all_final_qs, progress)
                    
                    st.download_button("📥 下载最终版 PPT 课件", ppt_data, "物理精品课件.pptx", use_container_width=True)
                    status.update(label="🎉 PPT 生成成功！", state="complete")
                else:
                    st.error("未能提取到有效内容，请检查文件。")
