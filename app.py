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
# 1. 核心单例模型
# ==========================================
@st.cache_resource
def get_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：物理碎片提取 (不丢符号，不丢图)
# ==========================================
def extract_physics_diagrams(img_path, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 物理特化：横向膨胀合并公式与碎片
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 10))
    dilated = cv2.dilate(binary, kernel, iterations=1)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    visuals = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w > w_img * 0.9 or h > h_img * 0.9: continue # 过滤大边框
        if w < 25 or h < 15: continue # 过滤微小噪点
        
        roi = img[y:y+h, x:x+w]
        if np.mean(roi) > 252: continue
        
        f_name = f"phys_{int(time.time()*1000)}_{x}.png"
        f_path = os.path.join(out_dir, f_name)
        cv2.imwrite(f_path, roi)
        visuals.append({"path": f_path, "y": y + h/2, "area": w * h})
    return visuals

# ==========================================
# 3. PPT 强力渲染：硬编码坐标防重叠
# ==========================================
def render_physics_problem(prs, q_data, q_idx):
    """
    通过硬编码 Inches 位置，确保版面绝对整齐。
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # --- A. 强制灰色底色 ---
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
    
    # --- B. 标题区 (高度锁定在 0.3 - 1.0) ---
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    
    tb_title = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(11), Inches(0.8))
    p_title = tb_title.text_frame.paragraphs[0]
    p_title.alignment = PP_ALIGN.LEFT
    r_title = p_title.add_run()
    r_title.text = f"习题精讲 第 {q_idx + 1} 题"
    r_title.font.size, r_title.font.bold = Pt(26), True
    r_title.font.color.rgb = RGBColor(20, 40, 80)

    # --- C. 计算左右分栏 ---
    has_imgs = len(q_data.get('imgs', [])) > 0
    txt_w = Inches(8.3) if has_imgs else Inches(12.3)

    # --- D. 题干主卡片 (位置锁定在 1.3 英寸开始) ---
    card1 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.3))
    card1.fill.solid(); card1.fill.fore_color.rgb = RGBColor(255, 255, 255); card1.line.color.rgb = RGBColor(220, 220, 225)
    
    tf = card1.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    # 文字填充 (锁定黑色 + 左对齐)
    raw_lines = q_data.get('text', '').split('\n')
    for idx, line in enumerate(raw_lines):
        if not line.strip(): continue
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.3
        r = p.add_run()
        r.text = line.strip()
        r.font.name = '微软雅黑'
        r.font.size = Pt(18)
        r.font.color.rgb = RGBColor(0, 0, 0)
        try:
            # 中文字体兼容注入
            rPr = r._r.get_or_add_rPr()
            rPr.get_or_add_ea().set('typeface', '微软雅黑')
        except: pass

    # --- E. 解析卡片 (位置锁定在 5.8 英寸开始) ---
    card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.8), txt_w, Inches(1.4))
    card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(220, 220, 225)
    p2 = card2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run(); r2.text = "解析预留：请在此处补充详细受力分析与解题步骤..."; r2.font.size = Pt(16); r2.font.color.rgb = RGBColor(120, 120, 120)

    # --- F. 右侧视觉元素展示 ---
    if has_imgs:
        y_ptr = 1.3
        # 挑选该题目区域内面积最大的3个元素
        q_imgs = sorted(q_data['imgs'], key=lambda x: x['area'], reverse=True)[:3]
        # 按原本位置重新排序
        q_imgs = sorted(q_imgs, key=lambda x: x['y'])
        for img in q_imgs:
            try:
                slide.shapes.add_picture(img['path'], Inches(9.0), Inches(y_ptr), width=Inches(4.0))
                y_ptr += 2.0 
            except: pass

# ==========================================
# 4. 主流程逻辑
# ==========================================
st.set_page_config(page_title="物理教研 AI 最终版", layout="centered")
st.title("⚛️ AI 物理教研工作站 (最终稳定修复版)")

files = st.file_uploader("📥 上传习题 (Word/PDF/图片)", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 生成最终版课件", type="primary", use_container_width=True):
    if not files:
        st.error("老师，请先上传文件。")
    else:
        all_final_qs = []
        engine = get_engine()
        with st.status("正在进行深度语义解析...", expanded=True) as status:
            with tempfile.TemporaryDirectory() as tmp:
                for file in files:
                    ext = file.name.split('.')[-1].lower()
                    if ext == 'docx':
                        doc = Document(io.BytesIO(file.read()))
                        cur_q = None
                        for p in doc.paragraphs:
                            txt = p.text.strip()
                            if not txt: continue
                            # 鲁棒分割：只要包含题号特征或短段落
                            if re.match(r'^(\d+|[\(（]\d+[\)）])[\.．、\s]', txt) or len(txt) < 10:
                                if cur_q: all_final_qs.append(cur_q)
                                cur_q = {"text": txt, "imgs": []}
                            elif cur_q: cur_q["text"] += "\n" + txt
                            else: cur_q = {"text": txt, "imgs": []}
                        if cur_q: all_final_qs.append(cur_q)
                    
                    else:
                        # 视觉处理
                        img_paths = []
                        if ext == 'pdf':
                            pdf_data = fitz.open(stream=file.read(), filetype="pdf")
                            for page in pdf_data:
                                pix = page.get_pixmap(matrix=fitz.Matrix(2.5, 2.5))
                                p_p = os.path.join(tmp, f"p_{time.time()}.png")
                                pix.save(p_p); img_paths.append(p_p)
                        else:
                            p_p = os.path.join(tmp, file.name)
                            with open(p_p, "wb") as f: f.write(file.read())
                            img_paths.append(p_p)
                        
                        for p_p in img_paths:
                            res, _ = engine(p_p)
                            visuals = extract_physics_diagrams(p_p, tmp)
                            
                            page_qs = []
                            cur_q = None
                            if res:
                                for line in res:
                                    txt = line[1].strip()
                                    y_mid = (line[0][0][1] + line[0][2][1]) / 2
                                    # 针对物理真题优化的正则：匹配 1. 或 (1) 或 3-1. 等
                                    if re.match(r'^(\d+|[\(（]\d+[\)）])', txt):
                                        if cur_q: page_qs.append(cur_q)
                                        cur_q = {"text": txt, "imgs": [], "y_start": y_mid, "y_end": 9999}
                                        if page_qs: page_qs[-1]["y_end"] = y_mid
                                    elif cur_q:
                                        cur_q["text"] += "\n" + txt
                                    else:
                                        cur_q = {"text": txt, "imgs": [], "y_start": y_mid, "y_end": 9999}
                            if cur_q: page_qs.append(cur_q)
                            
                            # 关联图片到该题目的 Y 轴区间
                            for v in visuals:
                                for q in page_qs:
                                    if q["y_start"] <= v["y"] < q["y_end"]:
                                        q["imgs"].append(v)
                                        break
                            all_final_qs.extend(page_qs)

                if all_final_qs:
                    status.update(label="✅ 题目解析完毕，正在渲染 PPT...", state="running")
                    prs = Presentation()
                    prs.slide_width, prs.slide_height = PPT_WIDTH, PPT_HEIGHT
                    for i, q in enumerate(all_final_qs):
                        render_physics_problem(prs, q, i)
                    
                    ppt_io = io.BytesIO()
                    prs.save(ppt_io)
                    st.download_button("📥 下载物理精品课件.pptx", ppt_io.getvalue(), "物理教研专家课件.pptx", use_container_width=True)
                    status.update(label="🎉 课件生成成功，祝您备课顺利！", state="complete")
