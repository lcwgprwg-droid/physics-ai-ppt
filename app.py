import os
import re
import cv2
import numpy as np
import tempfile
import io
import time
import gc
import streamlit as st
import fitz  # PyMuPDF
from docx import Document # python-docx
from rapidocr_onnxruntime import RapidOCR
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn
from PIL import Image

# ==========================================
# 1. 核心引擎：全局单例 OCR
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

# ==========================================
# 2. PPT 审美还原模块 (严格遵守老师原有的排版参数)
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
    """还原老师的：灰色背景 + 蓝色侧边条设计"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 全屏底色
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
    
    # 蓝色视觉指示条
    bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.12), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    
    # 页面标题
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.32), Inches(11.5), Inches(0.8))
    p = tb.text_frame.paragraphs[0]; p.text = title_text
    p.font.bold = True; p.font.size = Pt(26); p.font.color.rgb = RGBColor(30, 40, 60)
    set_font(p.runs[0])
    return slide

def add_badge_card(slide, x, y, w, h, badge_text, badge_color, content, font_size):
    """还原老师的：白色圆角卡片 + 顶部彩色悬浮标签"""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(230, 230, 235)
    
    # 卡片上方的小标签
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.2), y - Inches(0.15), Inches(1.2), Inches(0.35))
    badge.fill.solid(); badge.fill.fore_color.rgb = badge_color; badge.line.fill.background()
    bp = badge.text_frame.paragraphs[0]; bp.text = badge_text; bp.font.bold = True; bp.font.size = Pt(12)
    bp.font.color.rgb = RGBColor(255, 255, 255); bp.alignment = PP_ALIGN.CENTER; set_font(bp.runs[0])
    
    # 卡片内的正文
    tb = slide.shapes.add_textbox(x + Inches(0.2), y + Inches(0.4), w - Inches(0.4), h - Inches(0.6))
    tf = tb.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    lines = content.split('\n')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line; p.font.size = Pt(font_size); p.font.color.rgb = RGBColor(30, 30, 30); set_font(p.runs[0])

# ==========================================
# 3. 视觉算法：排除题目边框，精准提取插图
# ==========================================
def extract_diagrams_smart(img_path, ocr_result, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    # 提取所有OCR文字的坐标，用于密度检测
    text_regions = [np.array(line[0]) for line in ocr_result] if ocr_result else []

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    thresh = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 11, 2)
    
    # 物理图线条连通性增强
    kernel = np.ones((3,3), np.uint8)
    dilated = cv2.dilate(thresh, kernel, iterations=1)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    diagrams = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        
        # 1. 尺寸过滤：排除全页边框或极小噪点
        if w > w_img * 0.9 or h > h_img * 0.9: continue
        if w < 30 or h < 30: continue
        
        # 2. 核心：密度检测。如果一个框内文字块>3个，认为它是题干容器，不是插图
        text_count = 0
        for tr in text_regions:
            tx_min, ty_min = np.min(tr, axis=0)
            tx_max, ty_max = np.max(tr, axis=0)
            if tx_min >= x and tx_max <= x+w and ty_min >= y and ty_max <= y+h:
                text_count += 1
        
        if text_count > 3: continue # 避开包含大量文字的圆角矩形框
        
        # 保存真正插图
        roi = img[y:y+h, x:x+w]
        f_path = os.path.join(out_dir, f"diag_{int(time.time()*1000)}.png")
        cv2.imwrite(f_path, roi)
        diagrams.append({"path": f_path, "y": y + h/2, "x": x + w/2})
    return diagrams

# ==========================================
# 4. Word (.docx) 解析引擎
# ==========================================
def parse_word_doc(file_bytes, temp_dir):
    doc = Document(io.BytesIO(file_bytes))
    questions = []
    current_q = None
    
    # 提取文字
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text: continue
        if re.match(r'^\d+[\.．、]', text): # 题号匹配
            if current_q: questions.append(current_q)
            current_q = {"text": text, "imgs": [], "type": "word"}
        elif current_q:
            current_q["text"] += "\n" + text
    if current_q: questions.append(current_q)

    # 提取Word内嵌入的图片
    img_idx = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            img_idx += 1
            img_p = os.path.join(temp_dir, f"word_img_{img_idx}.png")
            with open(img_p, "wb") as f:
                f.write(rel.target_part.blob)
            # 简单的Word关联策略：图片按顺序分配给最近的题
            if questions:
                target = questions[min(img_idx-1, len(questions)-1)]
                target['imgs'].append(img_p)
    return questions

# ==========================================
# 5. 主流程控制层
# ==========================================
def run_app():
    st.set_page_config(page_title="高级教研PPT生成器", layout="wide")
    st.title("⚛️ 物理教研自动化工作站")
    st.caption("支持 Word / PDF / 图片，采用精细化物理图像定位算法")

    files = st.file_uploader("📤 上传文件 (图片/PDF/Word)", accept_multiple_files=True, type=['png','jpg','pdf','docx'])
    
    if st.button("🚀 生成精美 PPT 课件", type="primary"):
        if not files:
            st.warning("老师，请先上传文件。")
            return

        engine = get_ocr_engine()
        all_questions = []
        
        with tempfile.TemporaryDirectory() as tmp:
            for file in files:
                ext = file.name.split('.')[-1].lower()
                
                # --- 分支 1: Word ---
                if ext == 'docx':
                    all_questions.extend(parse_word_doc(file.read(), tmp))
                
                # --- 分支 2: 图片/PDF ---
                else:
                    images_to_proc = []
                    if ext == 'pdf':
                        pdf = fitz.open(stream=file.read(), filetype="pdf")
                        for p_idx in range(len(pdf)):
                            pix = pdf[p_idx].get_pixmap(matrix=fitz.Matrix(2, 2))
                            p_path = os.path.join(tmp, f"pdf_{p_idx}.png")
                            pix.save(p_path)
                            images_to_proc.append(p_path)
                    else:
                        p_path = os.path.join(tmp, file.name)
                        with open(p_path, "wb") as f: f.write(file.read())
                        images_to_proc.append(p_path)

                    for img_path in images_to_proc:
                        res, _ = engine(img_path)
                        diags = extract_diagrams_smart(img_path, res, tmp)
                        
                        # 识别题干
                        page_qs = []
                        cur_q = None
                        for line in res:
                            text = line[1].strip()
                            if re.match(r'^\d+[\.．、]', text):
                                if cur_q: page_qs.append(cur_q)
                                cur_q = {"text": text, "imgs": [], "y": line[0][0][1]}
                            elif cur_q:
                                cur_q["text"] += "\n" + text
                        if cur_q: page_qs.append(cur_q)
                        
                        # 图片吸附到最近的题干
                        for d in diags:
                            if page_qs:
                                target = min(page_qs, key=lambda q: abs(q['y'] - d['y']))
                                target['imgs'].append(d['path'])
                        all_questions.extend(page_qs)

            # --- PPT 渲染 ---
            if all_questions:
                prs = Presentation()
                prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
                c_blue, c_orange = RGBColor(0, 112, 192), RGBColor(230, 90, 40)

                for i, q in enumerate(all_questions):
                    slide = create_base_slide(prs, f"习题精讲 第 {i+1} 题")
                    has_img = len(q['imgs']) > 0
                    txt_w = Inches(8.5) if has_img else Inches(12.0)
                    
                    # 绘制原题卡片
                    add_badge_card(slide, Inches(0.5), Inches(1.3), txt_w, Inches(3.8), "原题呈现", c_blue, q['text'], 18)
                    # 绘制解析预留卡片
                    add_badge_card(slide, Inches(0.5), Inches(5.4), txt_w, Inches(1.6), "思路分析", c_orange, "待补充详细解析过程...", 16)
                    
                    # 绘制插图
                    if has_img:
                        curr_y = 1.3
                        for img_p in q['imgs'][:2]:
                            pic = slide.shapes.add_picture(img_p, Inches(9.2), Inches(curr_y), width=Inches(3.8))
                            curr_y += (pic.height / 914400) + 0.2
                
                buf = io.BytesIO()
                prs.save(buf)
                st.success(f"成功处理 {len(all_questions)} 道题目！")
                st.download_button("📥 下载精美 PPT 课件", buf.getvalue(), "物理教研课件.pptx", use_container_width=True)
            else:
                st.error("未能识别出题目，请确保文档内有类似 '1.' 的题号标识。")

if __name__ == "__main__":
    run_app()
