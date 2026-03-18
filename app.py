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
# 1. OCR 加载
# ==========================================
@st.cache_resource
def get_rapid_ocr():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：物理图像抓取 (V26)
# ==========================================
def extract_visuals_final(img_path, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_i, w_i = img.shape[:2]
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 物理特化：横向大膨胀，把符号连成块
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (45, 10))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    res = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_i * 0.9 or h > h_i * 0.9: continue
        if w < 30 or h < 20: continue
        
        roi = img[y:y+h, x:x+w]
        if np.mean(roi) > 252: continue
        
        p = os.path.join(out_dir, f"v_{int(time.time()*1000)}_{x}.png")
        cv2.imwrite(p, roi)
        res.append({"path": p, "y": y + h/2, "area": w * h})
    return res

# ==========================================
# 3. PPT 渲染引擎：【绝对位置排版】
# ==========================================
def render_physics_slide(prs, q_data, idx):
    # 手动定义常量，杜绝 NameError
    W = Inches(13.333)
    H = Inches(7.5)
    C_BLACK = RGBColor(0, 0, 0)
    C_TITLE = RGBColor(20, 40, 80)
    C_BG = RGBColor(245, 247, 250)
    
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 底色
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    bg.fill.solid(); bg.fill.fore_color.rgb = C_BG; bg.line.fill.background()
    
    # 标题框 (位置固定)
    tb_t = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(11), Inches(0.8))
    p_t = tb_t.text_frame.paragraphs[0]
    p_t.alignment = PP_ALIGN.LEFT
    r_t = p_t.add_run()
    r_t.text = f"习题精讲 第 {idx + 1} 题"
    r_t.font.size, r_t.font.bold, r_t.font.color.rgb = Pt(28), True, C_TITLE
    
    # 判定是否有图
    has_img = len(q_data.get('imgs', [])) > 0
    txt_w = Inches(8.5) if has_img else Inches(12.3)

    # 题干卡片 (位置固定)
    card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.3))
    card1.fill.solid(); card1.fill.fore_color.rgb = RGBColor(255, 255, 255); card1.line.color.rgb = RGBColor(210, 210, 220)
    
    tf = card1.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    # 填充文字
    txt = q_data.get('text', '').strip()
    lines = txt.split('\n')
    for i, line in enumerate(lines):
        if not line.strip(): continue
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.3
        r = p.add_run()
        r.text = line.strip()
        r.font.name = '微软雅黑'
        r.font.size = Pt(18)
        r.font.color.rgb = C_BLACK
        # 强制设置中文字体节点
        try:
            rPr = r._r.get_or_add_rPr()
            ea = rPr.get_or_add_ea()
            ea.set('typeface', '微软雅黑')
        except: pass

    # 解析卡片 (位置固定在下方)
    card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.8), txt_w, Inches(1.4))
    card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(210, 210, 220)
    p2 = card2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run(); r2.text = "解析：正在整理受力分析与列式过程..."; r2.font.size = Pt(16); r2.font.color.rgb = RGBColor(120, 120, 120)

    # 投放图片
    if has_img:
        y_p = 1.3
        # 取面积最大的3张，按Y坐标排
        imgs = sorted(q_data['imgs'], key=lambda x: x['area'], reverse=True)[:3]
        imgs = sorted(imgs, key=lambda x: x['y'])
        for info in imgs:
            try:
                slide.shapes.add_picture(info['path'], Inches(9.2), Inches(y_p), width=Inches(3.8))
                y_p += 2.0
            except: pass

# ==========================================
# 4. 主程序入口
# ==========================================
st.set_page_config(page_title="物理教研工作站", layout="centered")
st.title("⚛️ AI 物理教研自动化 (最终完美版)")

files = st.file_uploader("📥 上传 Word / PDF / 图片", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 生成 PPT 课件", type="primary", use_container_width=True):
    if not files:
        st.error("老师，请先上传文件。")
    else:
        all_final_qs = []
        engine = get_rapid_ocr()
        
        with st.status("正在进行深度解析...", expanded=True) as status:
            with tempfile.TemporaryDirectory() as tmpdir:
                for file in files:
                    ext = file.name.split('.')[-1].lower()
                    if ext == 'docx':
                        doc = Document(io.BytesIO(file.read()))
                        cur_q = None
                        for p in doc.paragraphs:
                            t = p.text.strip()
                            if not t: continue
                            # 题号分页逻辑：数字开头就分页
                            if re.match(r'^(\d+|[\(（]\d+[\)）])', t):
                                if cur_q: all_final_qs.append(cur_q)
                                cur_q = {"text": t, "imgs": []}
                            elif cur_q: cur_q["text"] += "\n" + t
                            else: cur_q = {"text": t, "imgs": []}
                        if cur_q: all_final_qs.append(cur_q)
                    
                    else:
                        # 视觉解析
                        paths = []
                        if ext == 'pdf':
                            pdf = fitz.open(stream=file.read(), filetype="pdf")
                            for i in range(len(pdf)):
                                pix = pdf[i].get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                                p = os.path.join(tmpdir, f"p_{i}_{time.time()}.png")
                                pix.save(p); paths.append(p)
                        else:
                            p = os.path.join(tmpdir, file.name)
                            with open(p, "wb") as f: f.write(file.read())
                            paths.append(p)
                        
                        for p_path in paths:
                            res, _ = engine(p_path)
                            visuals = extract_visuals_final(p_path, tmpdir)
                            
                            page_qs = []
                            cur_q = None
                            if res:
                                for line in res:
                                    txt = line[1].strip()
                                    y_m = (line[0][0][1] + line[0][2][1]) / 2
                                    # 物理分页触发器
                                    if re.match(r'^(\d+|[\(（]\d+[\)）])', txt):
                                        if cur_q: page_qs.append(cur_q)
                                        cur_q = {"text": txt, "imgs": [], "y1": y_m, "y2": 9999}
                                        if page_qs: page_qs[-1]["y2"] = y_m
                                    elif cur_q: cur_q["text"] += "\n" + txt
                                    else: cur_q = {"text": txt, "imgs": [], "y1": y_m, "y2": 9999}
                            if cur_q: page_qs.append(cur_q)
                            
                            # 关联图片
                            for v in visuals:
                                for q in page_qs:
                                    if q["y1"] <= v["y"] < q["y2"]:
                                        q["imgs"].append(v); break
                            all_final_qs.extend(page_qs)

                if all_final_qs:
                    status.update(label="✅ 解析完毕，正在渲染 PPT...", state="running")
                    prs = Presentation()
                    # 强力锁定宽高，拒绝 NameError
                    prs.slide_width = Inches(13.333)
                    prs.slide_height = Inches(7.5)
                    
                    for idx, q_item in enumerate(all_final_qs):
                        render_physics_slide(prs, q_item, idx)
                    
                    out = io.BytesIO()
                    prs.save(out)
                    st.download_button("📥 下载物理精品课件", out.getvalue(), "物理教研课件.pptx", use_container_width=True)
                    status.update(label="🎉 课件生成成功！", state="complete")
                else:
                    st.error("未能提取题目内容。")
