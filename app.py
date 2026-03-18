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
# 1. 工业级配置中心
# ==========================================
class Config:
    PPT_W, PPT_H = Inches(13.333), Inches(7.5)
    BLACK = RGBColor(0, 0, 0)
    DARK_BLUE = RGBColor(20, 40, 80)
    BG_GRAY = RGBColor(245, 247, 250)
    # 物理识别参数
    MIN_FIG_AREA = 1200  # 最小图形面积
    MAX_FIG_RATIO = 0.85 # 过滤掉占页面 85% 以上的大边框

@st.cache_resource
def load_engine():
    return RapidOCR()

# ==========================================
# 2. 高级图像定位引擎 (OpenCV 重构)
# ==========================================
def physics_vision_engine(img_path, ocr_result, out_dir):
    """
    高级视觉算法：
    1. 使用 Scharr 算子强化物理线条特征。
    2. 使用 OCR 禁区避让：在文字 10px 范围内不进行图片切割，但允许包含公式。
    """
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 获取文字禁区
    forbidden_mask = np.zeros((h_img, w_img), dtype=np.uint8)
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            # 文字区域向外扩充 3 像素作为“保护区”
            cv2.fillPoly(forbidden_mask, [box], 255)

    # 预处理：灰度与边缘增强
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # 使用 Scharr 提取边缘 (比 Canny 更适合捕捉物理细线)
    grad_x = cv2.Scharr(gray, cv2.CV_32F, 1, 0)
    grad_y = cv2.Scharr(gray, cv2.CV_32F, 0, 1)
    grad = cv2.subtract(cv2.convertScaleAbs(grad_x), cv2.convertScaleAbs(grad_y))
    grad = cv2.GaussianBlur(grad, (3, 3), 0)
    
    # 二值化
    _, binary = cv2.threshold(grad, 30, 255, cv2.THRESH_BINARY)
    
    # 物理公式粘合：使用横向核连接 T_1=300K
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (45, 8))
    morphed = cv2.dilate(binary, kernel, iterations=1)
    
    # 关键：寻找轮廓
    contours, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    visual_items = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 工业级过滤逻辑
        if w > w_img * Config.MAX_FIG_RATIO: continue # 剔除大矩形外框
        if w < 30 or h < 20: continue # 剔除噪点
        
        # 判定该区域是否包含“有效视觉信息”（非文字背景）
        # 计算区域内文字占比
        roi_mask = forbidden_mask[y:y+h, x:x+w]
        text_pixel_ratio = np.sum(roi_mask == 255) / (w * h)
        
        # 如果 70% 都是文字且面积不大，则认为是普通文字行，不作为图片
        if text_pixel_ratio > 0.7 and w * h < 10000:
            continue
            
        # 截取原图
        roi = img[y:y+h, x:x+w]
        if np.mean(roi) > 251: continue # 过滤空白

        f_name = f"diag_{int(time.time()*1000)}_{x}.png"
        f_path = os.path.join(out_dir, f_name)
        cv2.imwrite(f_path, roi)
        visual_items.append({"path": f_path, "y": y + h/2, "area": w * h})
    
    # 按垂直位置排序
    return sorted(visual_items, key=lambda x: x['y'])

# ==========================================
# 3. PPT 渲染 Builder (强制黑色左对齐)
# ==========================================
def build_physics_slide(prs, q_data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = Config.BG_GRAY; bg.line.fill.background()
    
    # 标题
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(11), Inches(0.8))
    p_t = title_box.text_frame.paragraphs[0]
    p_t.alignment = PP_ALIGN.LEFT
    run_t = p_t.add_run()
    run_t.text = f"习题精讲 - 深度解析"
    run_t.font.size, run_t.font.bold = Pt(26), True
    run_t.font.color.rgb = Config.DARK_BLUE

    # 布局计算
    has_img = len(q_data.get('imgs', [])) > 0
    txt_w = Inches(8.5) if has_img else Inches(12.3)

    # 1. 题干卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.3))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(210, 210, 215)
    
    tf = card.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    # 填充正文 (锁定黑色左对齐)
    lines = q_data.get('text', '').split('\n')
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.3
        r = p.add_run()
        r.text = line.strip()
        r.font.name = '微软雅黑'
        r.font.size = Pt(18)
        r.font.color.rgb = RGBColor(0, 0, 0) # 强制黑色
        # 中文兼容
        try:
            rPr = r._r.get_or_add_rPr()
            rPr.get_or_add_ea().set('typeface', '微软雅黑')
        except: pass

    # 2. 解析卡片
    card2 = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(5.8), txt_w, Inches(1.4))
    card2.fill.solid(); card2.fill.fore_color.rgb = RGBColor(255, 255, 255); card2.line.color.rgb = RGBColor(210, 210, 215)
    p2 = card2.text_frame.paragraphs[0]
    p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run(); r2.text = "解析：待补充受力分析与列式讲解过程..."; r2.font.color.rgb = RGBColor(100, 100, 100); r2.font.size = Pt(16)

    # 3. 右侧视觉投放 (核心修正：不丢图)
    if has_img:
        y_ptr = 1.3
        # 按面积降序，取最显著的 3 张图
        imgs = sorted(q_data['imgs'], key=lambda x: x['area'], reverse=True)[:3]
        # 再按 y 轴重排，保证逻辑顺序
        imgs = sorted(imgs, key=lambda x: x['y'])
        for img_info in imgs:
            try:
                slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_ptr), width=Inches(3.8))
                y_ptr += 2.0 
            except: pass

# ==========================================
# 4. 主业务工作流
# ==========================================
st.set_page_config(page_title="高级教研 AI 工具", layout="centered")
st.title("⚛️ AI 物理教研自动化 (V25 视觉增强版)")

files = st.file_uploader("📥 上传习题 (Word/PDF/图片)", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🚀 开始极速转换", type="primary", use_container_width=True):
    if not files:
        st.error("老师，请上传文件。")
    else:
        all_final_qs = []
        with st.status("🔍 正在启动物理视觉引擎...", expanded=True) as status:
            engine = load_engine()
            
            with tempfile.TemporaryDirectory() as tmpdir:
                for file in files:
                    st.write(f"📄 正在处理: {file.name}")
                    ext = file.name.split('.')[-1].lower()
                    
                    if ext == 'docx':
                        doc = Document(io.BytesIO(file.read()))
                        full_text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                        if full_text:
                            all_final_qs.append({"text": full_text, "imgs": []})
                    else:
                        # 视觉处理路径
                        paths = []
                        if ext == 'pdf':
                            pdf_doc = fitz.open(stream=file.read(), filetype="pdf")
                            for page in pdf_doc:
                                pix = page.get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
                                p = os.path.join(tmpdir, f"p_{time.time()}.png")
                                pix.save(p); paths.append(p)
                        else:
                            p = os.path.join(tmpdir, file.name)
                            with open(p, "wb") as f: f.write(file.read())
                            paths.append(p)
                        
                        for p_path in paths:
                            ocr_res, _ = engine(p_path)
                            # 【核心改进调用】
                            visuals = physics_vision_engine(p_path, ocr_res, tmpdir)
                            
                            # 见字必录
                            full_page_text = ""
                            if ocr_res:
                                for line in ocr_res: full_page_text += line[1] + "\n"
                            
                            if full_page_text or visuals:
                                all_final_qs.append({"text": full_page_text, "imgs": visuals, "y": 0})

                if all_final_qs:
                    status.update(label="✅ 解析完成，正在渲染 PPT...", state="running")
                    prs = Presentation()
                    prs.slide_width, prs.slide_height = Config.WIDTH, Config.HEIGHT
                    for q in all_final_qs:
                        build_physics_slide(prs, q)
                    
                    ppt_buf = io.BytesIO()
                    prs.save(ppt_buf)
                    st.download_button("📥 下载生成好的 PPT", ppt_buf.getvalue(), "物理精品课件.pptx", use_container_width=True)
                    status.update(label="🎉 转换成功！", state="complete")
