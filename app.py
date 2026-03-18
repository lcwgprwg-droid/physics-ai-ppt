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
# 1. 核心配置 (常量硬编码，杜绝属性错误)
# ==========================================
PPT_WIDTH = Inches(13.333)
PPT_HEIGHT = Inches(7.5)
C_BLACK = RGBColor(0, 0, 0)
C_BLUE = RGBColor(0, 112, 192)
C_BG = RGBColor(245, 247, 250)

@st.cache_resource
def get_engine():
    return RapidOCR()

# ==========================================
# 2. 视觉算法：物理碎片聚类提取 (不再丢图)
# ==========================================
def extract_visuals_robust(img_path, ocr_result, out_dir):
    """
    针对物理题设计：
    1. 不再扣除文字，防止弄断插图线条。
    2. 使用 (40, 10) 横向大核膨胀，专门把横排公式 T1=300K 粘成一块。
    3. 过滤全页大框，保留核心图示。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 预处理
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    # 横向膨胀：将公式字母、等号、单位连在一起
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 10))
    morphed = cv2.dilate(binary, kernel, iterations=1)
    
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    # 建立文字区域索引
    text_boxes = []
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            text_boxes.append([np.min(box[:,0]), np.min(box[:,1]), np.max(box[:,0]), np.max(box[:,1])])

    visual_elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        # 1. 排除页边大框
        if w > w_img * 0.9 or h > h_img * 0.9: continue
        # 2. 排除极小噪点
        if w < 25 or h < 15: continue 
        
        # 3. 智能判断：如果这个框内 80% 都是 OCR 认出的正文，则不切图
        is_covered_by_text = False
        for tb in text_boxes:
            # 计算重叠比例
            ix1, iy1 = max(x, tb[0]), max(y, tb[1])
            ix2, iy2 = min(x+w, tb[2]), min(y+h, tb[3])
            if ix1 < ix2 and iy1 < iy2:
                overlap_area = (ix2 - ix1) * (iy2 - iy1)
                if overlap_area / (w * h) > 0.8:
                    is_covered_by_text = True
                    break
        
        # 如果不是纯文字，或者是孤立的大视觉块（公式），就提取
        if not is_covered_by_text or (w * h > 5000):
            roi = img[y:y+h, x:x+w]
            if np.mean(roi) > 252: continue
            
            f_name = f"phys_{int(time.time()*1000)}_{x}.png"
            f_path = os.path.join(out_dir, f_name)
            cv2.imwrite(f_path, roi)
            visual_elements.append({"path": f_path, "y": y + h/2, "area": w * h})
            
    return visual_elements

# ==========================================
# 3. PPT 渲染引擎 (逐行设置，杜绝崩溃)
# ==========================================
def create_slide_safe(prs, q_data, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # A. 背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = C_BG
    bg.line.fill.background()
    
    # B. 蓝色侧边
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_BLUE
    bar.line.fill.background()
    
    # C. 标题
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(11), Inches(0.8))
    p_t = title_box.text_frame.paragraphs[0]
    p_t.alignment = PP_ALIGN.LEFT
    r_t = p_t.add_run()
    r_t.text = f"习题精讲 第 {idx + 1} 题"
    r_t.font.size, r_t.font.bold = Pt(26), True
    r_t.font.color.rgb = RGBColor(20, 40, 80)
    
    has_imgs = len(q_data.get('imgs', [])) > 0
    txt_w = Inches(8.5) if has_imgs else Inches(12.3)

    # D. 题干卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.3))
    card.fill.solid()
    card.fill.fore_color.rgb = RGBColor(255, 255, 255)
    card.line.color.rgb = RGBColor(210, 210, 215)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    # E. 填充内容 (强制黑色+左对齐)
    lines = q_data.get('text', '').split('\n')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.3
        r = p.add_run()
        r.text = line.strip()
        r.font.name = '微软雅黑'
        r.font.size = Pt(18)
        r.font.color.rgb = C_BLACK # 强制锁定黑色
        try:
            rPr = r._r.get_or_add_rPr()
            rPr.get_or_add_ea().set('typeface', '微软雅黑')
        except: pass

    # F. 解析卡片
    card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.8), txt_w, Inches(1.4))
    card2.fill.solid()
    card2.fill.fore_color.rgb = RGBColor(255, 255, 255)
    card2.line.color.rgb = RGBColor(210, 210, 215)
    p2 = card2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run()
    r2.text = "解析：正在整理受力分析与列式过程..."; r2.font.color.rgb = RGBColor(100, 100, 100); r2.font.size = Pt(16)

    # G. 图片投放
    if has_imgs:
        y_ptr = 1.3
        # 取面积最大的 3 张作为核心图，并按垂直位置重排
        best_imgs = sorted(q_data['imgs'], key=lambda x: x['area'], reverse=True)[:3]
        best_imgs = sorted(best_imgs, key=lambda x: x['y'])
        for img_info in best_imgs:
            try:
                slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                y_ptr += 2.0 
            except: pass

# ==========================================
# 4. 主业务流程
# ==========================================
st.set_page_config(page_title="高级物理教研 AI", layout="centered")
st.title("⚛️ AI 物理教研工作站 (最终工业版)")

files = st.file_uploader("📥 上传资料 (Word/PDF/图片)", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 极速一键转换", type="primary", use_container_width=True):
    if not files:
        st.error("老师，请先上传文件。")
    else:
        all_final_qs = []
        with st.status("🔍 AI 正在深度扫描内容...", expanded=True) as status:
            engine = get_engine()
            with tempfile.TemporaryDirectory() as tmpdir:
                for file in files:
                    ext = file.name.split('.')[-1].lower()
                    st.write(f"正在处理: {file.name}")
                    
                    if ext == 'docx':
                        doc = Document(io.BytesIO(file.read()))
                        # Word 简单全量捕获
                        full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                        if full_text:
                            # 鲁棒切题逻辑
                            parts = re.split(r'(\d+[\.．、\s])', full_text)
                            if len(parts) <= 1:
                                all_final_qs.append({"text": full_text, "imgs": []})
                            else:
                                cur_q = None
                                for p in parts:
                                    if re.match(r'\d+[\.．、\s]', p):
                                        if cur_q: all_final_qs.append(cur_q)
                                        cur_q = {"text": p, "imgs": []}
                                    elif cur_q: cur_q["text"] += p
                                if cur_q: all_final_qs.append(cur_q)
                    else:
                        # 视觉分支 (PDF/图片)
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
                            res, _ = engine(p_path)
                            # 核心視覺識別
                            visuals = extract_visuals_robust(p_path, res, tmpdir)
                            # 内容合并
                            page_txt = "\n".join([line[1] for line in res]) if res else ""
                            if page_txt or visuals:
                                all_final_qs.append({"text": page_txt, "imgs": visuals, "y": 0})

                if all_final_qs:
                    status.update(label="✅ 扫描完成，正在生成课件...", state="running")
                    prs = Presentation()
                    # 关键修复：确保属性名与定义一致
                    prs.slide_width, prs.slide_height = PPT_WIDTH, PPT_HEIGHT
                    
                    for i, q in enumerate(all_final_qs):
                        create_slide_safe(prs, q, i)
                    
                    ppt_buf = io.BytesIO()
                    prs.save(ppt_buf)
                    st.download_button("📥 点击下载物理精选课件", ppt_buf.getvalue(), "物理教研专家课件.pptx", use_container_width=True)
                    status.update(label="🎉 课件生成成功！", state="complete")
                else:
                    st.error("未能找到有效内容，请检查上传的文件。")
