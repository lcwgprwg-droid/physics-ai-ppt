import os
import re
import cv2
import numpy as np
import tempfile
import io
import time
import streamlit as st
import matplotlib.pyplot as plt
from docx import Document
from rapidocr_onnxruntime import RapidOCR
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

# ==========================================
# 1. 核心引擎：增加公式渲染器
# ==========================================
@st.cache_resource
def get_ocr_engine():
    return RapidOCR()

def latex_to_png(latex_str, out_path):
    """
    将 LaTeX 字符串渲染为透明背景的高清图片
    物理公式专用：支持分式、根号、希腊字母
    """
    try:
        # 去除 OCR 可能产生的非法字符
        latex_str = latex_str.replace('\n', '').strip()
        if not latex_str.startswith('$'): latex_str = f"${latex_str}$"
        
        plt.figure(figsize=(len(latex_str)*0.25, 1)) # 动态调整宽度
        plt.axis('off')
        plt.text(0.5, 0.5, latex_str, size=40, ha='center', va='center', color='#1E2850')
        # 保存为透明 PNG
        plt.savefig(out_path, format='png', transparent=True, bbox_inches='tight', pad_inches=0.1, dpi=300)
        plt.close()
        return True
    except Exception as e:
        print(f"渲染公式失败: {e}")
        return False

# ==========================================
# 2. 视觉算法：保持“工业级”重构版的稳定性
# ==========================================
def extract_physics_elements_v6(img_path, ocr_result, out_dir):
    img = cv2.imread(img_path)
    if img is None: return []
    h_img, w_img = img.shape[:2]
    
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 10)
    
    clean_binary = binary.copy()
    if ocr_result:
        for line in ocr_result:
            box = np.array(line[0]).astype(np.int32)
            cv2.rectangle(clean_binary, (box[0][0]-5, box[0][1]-5), (box[2][0]+5, box[2][1]+5), 0, -1)

    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (25, 25))
    dilated = cv2.dilate(clean_binary, kernel, iterations=1)
    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    
    elements = []
    for cnt in contours:
        x, y, w, h = cv2.boundingRect(cnt)
        if w < 20 or h < 20 or w > w_img * 0.9: continue
        
        roi = img[y:y+h, x:x+w]
        f_path = os.path.join(out_dir, f"ele_{int(time.time()*1000)}.png")
        cv2.imwrite(f_path, roi)
        elements.append({"path": f_path, "y": y + h/2, "type": "diagram"})
    return elements

# ==========================================
# 3. 语义分析：自动识别“潜伏”的公式
# ==========================================
def smart_format_text(text):
    """
    将 OCR 的烂文本清洗为带数学特征的文本
    在实际生产中，这一步建议接入 GPT-4o 或 Qwen-Max
    此处演示：识别常见的物理符号并标记
    """
    # 物理符号保护正则
    math_symbols = r'([vmaEFBtρΩθλΔπ\^/\\sqrt]+|[\d\.]+[mcm/skg]+)'
    # 模拟清洗逻辑
    formatted = text.replace("T1", "T_1").replace("300K", "300 \\text{K}")
    return formatted

# ==========================================
# 4. PPT 渲染引擎：图文+公式混合排版
# ==========================================
def apply_style(p, size=18, color=(40, 40, 40), bold=False):
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = 1.3
    run = p.add_run()
    run.font.name = '微软雅黑'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(*color)
    rPr = run._r.get_or_add_rPr()
    f = rPr.find(qn('w:rFonts'))
    if f is None:
        f = rPr.makeelement(qn('w:rFonts'))
        rPr.append(f)
    f.set(qn('w:eastAsia'), '微软雅黑')
    return run

def render_ppt_with_math(questions, out_path, tmp_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    
    for i, q in enumerate(questions):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # 灰色背景
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(245, 247, 250); bg.line.fill.background()
        
        # 标题栏
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(10), Inches(0.8))
        apply_style(title_box.text_frame.paragraphs[0], f"习题精讲 第 {i+1} 题", size=26, bold=True, color=(20, 50, 100))

        has_img = len(q['imgs']) > 0
        txt_w = Inches(8.5) if has_img else Inches(12.2)

        # 原题卡片
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.3), txt_w, Inches(4.0))
        card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 220, 225)
        
        # 识别文本中的公式并渲染 (模拟演示)
        # 如果题干中包含 T_1=300K 等字样，我们提取出来渲染
        clean_text = q['text']
        math_matches = re.findall(r'([A-Z]_\d\s*=\s*\d+\s*[K|m|s|Ω])', clean_text)
        
        tf = card.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        # 填入清洗后的题干
        apply_style(p, clean_text, size=18)
        
        # 如果有提取出的重要公式，作为“高清贴图”放在题干下方
        math_y = 4.0
        for m_str in math_matches:
            m_path = os.path.join(tmp_dir, f"math_{int(time.time()*1000)}.png")
            if latex_to_png(m_str, m_path):
                slide.shapes.add_picture(m_path, Inches(0.8), Inches(math_y), height=Inches(0.5))
                math_y += 0.6

        # 插图投放
        if has_img:
            y_offset = 1.3
            for img_info in q['imgs'][:3]:
                pic = slide.shapes.add_picture(img_info['path'], Inches(9.2), Inches(y_offset), width=Inches(3.8))
                y_offset += (pic.height / 914400) + 0.2

    prs.save(out_path)

# ==========================================
# 5. UI 与 业务流
# ==========================================
st.set_page_config(page_title="物理教研 AI 重构版", layout="centered")
st.title("⚛️ 物理习题 AI：高清公式渲染版")

uploaded_files = st.file_uploader("上传资料 (Word/PDF/图片)", accept_multiple_files=True, type=['docx', 'pdf', 'png', 'jpg'])

if st.button("🔥 生成专业级 PPT 课件", type="primary", use_container_width=True):
    if not uploaded_files:
        st.error("请先上传文件")
    else:
        all_qs = []
        engine = get_ocr_engine()
        with tempfile.TemporaryDirectory() as tmp:
            for file in uploaded_files:
                ext = file.name.split('.')[-1].lower()
                
                if ext == 'docx':
                    doc = Document(io.BytesIO(file.read()))
                    cur_q = None
                    for p in doc.paragraphs:
                        t = p.text.strip()
                        if not t: continue
                        if re.match(r'^\d+[\.．、]', t):
                            if cur_q: all_qs.append(cur_q)
                            cur_q = {"text": t, "imgs": []}
                        elif cur_q: cur_q['text'] += "\n" + t
                    if cur_q: all_qs.append(cur_q)
                
                else:
                    # PDF/图片处理逻辑 (同前一版)
                    # ... [此处保持前一版工业级视觉分割逻辑] ...
                    # 关键调用：elements = extract_physics_elements_v6(...)
                    pass # 限于篇幅不再重复粘贴视觉分割代码

            # 渲染 PPT
            output_ppt = os.path.join(tmp, "final.pptx")
            # 注意：此处 all_qs 应当由视觉/Word 解析流程填充完毕
            if all_qs:
                render_ppt_with_math(all_qs, output_ppt, tmp)
                with open(output_ppt, "rb") as f:
                    st.download_button("📥 下载专业级课件", f.read(), "物理精美课件.pptx", use_container_width=True)
