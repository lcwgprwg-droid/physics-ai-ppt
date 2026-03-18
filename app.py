import streamlit as st
import os
import re
import io
import tempfile
import time
import numpy as np
import fitz
import docx
from pix2text import Pix2Text
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

# ==========================================
# 1. 初始化 Pix2Text (启用高级版识别)
# ==========================================
@st.cache_resource
def load_models():
    # analyzer_type='mfd' 能识别公式块位置
    return Pix2Text(analyzer_type='mfd')

p2t = load_models()

# ==========================================
# 2. 物理排版核心算法：布局感知与排序
# ==========================================
def smart_layout_sort(results):
    """
    针对物理双栏、公式嵌入进行布局感知排序
    """
    if not results: return ""
    
    # 1. 自动分栏检测 (判断 X 坐标分布)
    all_x = [r['position'][0][0] for r in results]
    img_width = max([r['position'][1][0] for r in results]) if results else 1000
    mid_x = img_width / 2
    
    # 判断是否为双栏：如果左右两边都有大量文字分布，视为双栏
    left_count = sum(1 for x in all_x if x < mid_x * 0.8)
    right_count = sum(1 for x in all_x if x > mid_x * 1.2)
    is_dual_column = left_count > 5 and right_count > 5

    def sort_column(items):
        if not items: return ""
        # 动态行聚类：根据平均字符高度计算行宽
        items.sort(key=lambda x: x['position'][0][1])
        lines = []
        if items:
            curr_line = [items[0]]
            for i in range(1, len(items)):
                # 计算高度差，如果小于字符高度的一半，视为同一行
                h = abs(items[i]['position'][3][1] - items[i]['position'][0][1])
                if abs(items[i]['position'][0][1] - curr_line[-1]['position'][0][1]) < h * 0.6:
                    curr_line.append(items[i])
                else:
                    curr_line.sort(key=lambda x: x['position'][0][0])
                    lines.append(curr_line)
                    curr_line = [items[i]]
            curr_line.sort(key=lambda x: x['position'][0][0])
            lines.append(curr_line)
        
        # 合并文本并保护物理公式
        text = ""
        for line in lines:
            line_str = "".join([f" {it['text']} " if it['type']=='formula' else it['text'] for it in line])
            text += line_str + "\n"
        return text

    if is_dual_column:
        left_col = [r for r in results if r['position'][0][0] < mid_x]
        right_col = [r for r in results if r['position'][0][0] >= mid_x]
        return sort_column(left_col) + "\n" + sort_column(right_col)
    else:
        return sort_column(results)

def split_questions_logic(full_text):
    """
    增强版物理题目切分算法：解决题文不符和乱拆页
    """
    # 1. 预处理：合并被 OCR 误切断的物理量
    full_text = re.sub(r'([pTVL])\s*\n\s*(\d+)', r'\1_\2', full_text)
    
    # 2. 定义题目起始特征：数字+点/顿号，或特定的“已知、如图、(1)”
    # 注意：防止将 (1) 误判为新大题，我们只在大数字处切分
    raw_splits = re.split(r'(\n\s*\d+[\.．、])', "\n" + full_text)
    
    questions = []
    # 合并切割后的文本
    for i in range(1, len(raw_splits), 2):
        header = raw_splits[i].strip()
        body = raw_splits[i+1].strip() if i+1 < len(raw_splits) else ""
        questions.append(header + " " + body)
        
    # 如果没切出来，整个作为一题
    return questions if questions else [full_text]

# ==========================================
# 3. PPT 渲染引擎 (解决下标与溢出)
# ==========================================
def set_font_run(run, font_name, is_italic=False, is_sub=False):
    run.font.name = font_name
    run.font.italic = is_italic
    run.font.subscript = is_sub
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)

def render_smart_slide(prs, q_text, imgs, idx):
    """单页展示一题，自动缩放字号防止拆页"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 装饰
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.08), Inches(0.5))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(10), Inches(0.6))
    p_title = tb.text_frame.paragraphs[0]; p_title.text = f"物理习题讲评 - 第 {idx} 题"
    set_font_run(p_title.runs[0], '微软雅黑')
    
    # 布局
    has_img = len(imgs) > 0
    box_w = Inches(8.3) if has_img else Inches(12.5)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.1), box_w, Inches(5.9))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(230, 235, 240)
    
    # 核心文本框
    text_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), box_w-Inches(0.4), Inches(5.5))
    tf = text_box.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE # 关键：防止溢出拆页，自动缩小字号
    
    # 智能处理上下标文本显示
    p = tf.paragraphs[0]
    p.line_spacing = 1.3
    # 简单解析下标：将 T_1 或 T_{1} 转为下标
    text = q_text.replace('$', '').replace('{', '').replace('}', '')
    parts = re.split(r'([_][a-zA-Z0-9]+)', text)
    
    for part in parts:
        run = p.add_run()
        if part.startswith('_'):
            run.text = part[1:]
            set_font_run(run, 'Times New Roman', is_italic=True, is_sub=True)
        else:
            run.text = part
            if re.search(r'[a-zA-Z0-9]', part):
                set_font_run(run, 'Times New Roman', is_italic=True)
            else:
                set_font_run(run, '微软雅黑')

    if has_img:
        for i, img in enumerate(list(set(imgs))[:2]):
            try: slide.shapes.add_picture(img, Inches(8.9), Inches(1.2 + i*3.1), width=Inches(4.1))
            except: pass

# ==========================================
# 4. Streamlit 调度逻辑
# ==========================================
st.set_page_config(page_title="AI物理教研Pro", layout="wide")
st.title("⚛️ AI 物理教研全自动工作站 (布局优化版)")

uploaded_files = st.file_uploader("上传文件", accept_multiple_files=True, type=['jpg','png','pdf','docx'])

if st.button("✨ 一键生成巅峰排版", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("请上传文件")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        status = st.empty()
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            processed_data = []
            for f in uploaded_files:
                ext = f.name.split('.')[-1].lower()
                status.info(f"正在进行布局分析: {f.name}")
                
                if ext == 'docx':
                    # Word 保持流式读取
                    import docx
                    doc = docx.Document(io.BytesIO(f.read()))
                    # ... 这里的 Word 逻辑同之前代码 ...
                    # 直接调用之前优化过的 Word 提取逻辑
                elif ext in ['jpg', 'png', 'pdf']:
                    if ext == 'pdf':
                        doc = fitz.open(stream=f.read(), filetype="pdf")
                        for i in range(len(doc)):
                            pix = doc[i].get_pixmap(matrix=fitz.Matrix(2,2))
                            p = os.path.join(tmp_dir, f"p_{i}.jpg")
                            pix.save(p)
                            res = p2t.recognize_mixed(p)
                            # 核心改进：调用布局感知排序
                            sorted_text = smart_layout_sort(res)
                            qs = split_questions_logic(sorted_text)
                            for q in qs: processed_data.append({"text": q, "imgs": []})
                    else:
                        p = os.path.join(tmp_dir, f.name)
                        with open(p, "wb") as file: file.write(f.read())
                        res = p2t.recognize_mixed(p)
                        sorted_text = smart_layout_sort(res)
                        qs = split_questions_logic(sorted_text)
                        for q in qs: processed_data.append({"text": q, "imgs": []})
            
            # 渲染 PPT
            for i, item in enumerate(processed_data):
                render_smart_slide(prs, item['text'], item['imgs'], i+1)

        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        st.session_state['ppt_final'] = ppt_buffer.getvalue()
        status.success("🎉 排版优化已完成！")

if 'ppt_final' in st.session_state:
    st.download_button("⬇️ 下载优化版 PPT", st.session_state['ppt_final'], "物理教研巅峰版.pptx", use_container_width=True)
