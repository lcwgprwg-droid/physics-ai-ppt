import streamlit as st
import os
import re
import io
import tempfile
import time
import fitz  # PyMuPDF
import docx
from pix2text import Pix2Text
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml.ns import qn

# ==========================================
# 1. 初始化 Pix2Text (专业物理公式引擎)
# ==========================================
@st.cache_resource
def load_ai_engine():
    # analyzer_type='mfd' 负责定位公式在正文中的精准坐标
    return Pix2Text(analyzer_type='mfd')

p2t = load_ai_engine()

# ==========================================
# 2. 巅峰排版渲染核心 (修正命名与字体逻辑)
# ==========================================
def set_physics_font(run, font_name='微软雅黑', is_italic=False, is_sub=False, is_sup=False):
    """稳健的字体设置：修复 AttributeError 并支持上下标"""
    run.font.name = font_name
    run.font.italic = is_italic
    run.font.subscript = is_sub
    run.font.superscript = is_sup
    
    # 强制东亚字体渲染，防止中文乱码
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)

def render_rich_text(text_frame, raw_content):
    """
    智能流式渲染：解析文本中的 _(下标) 和 ^(上标)，并自动应用物理变量斜体
    """
    p = text_frame.paragraphs[0]
    p.line_spacing = 1.3
    
    # 清理并规范化标记
    text = raw_content.replace('$', '').replace('\\(', '').replace('\\)', '')
    text = text.replace('{', '').replace('}', '') 
    
    # 正则切分：识别 汉字序列 | 变量标记(_1, ^2) | 其他文本
    parts = re.split(r'([_][a-zA-Z0-9]+|[\^][a-zA-Z0-9]+)', text)
    
    for part in parts:
        if not part: continue
        is_sub = part.startswith('_')
        is_sup = part.startswith('^')
        clean_text = part[1:] if (is_sub or is_sup) else part
        
        run = p.add_run()
        run.text = clean_text
        run.font.size = Pt(20)
        
        # 物理特征识别：只要包含英文字符或上下标，统一用 Times New Roman 斜体
        if is_sub or is_sup or re.search(r'[a-zA-Z0-9\.\+\-\=\*×]', clean_text):
            set_physics_font(run, 'Times New Roman', is_italic=True, is_sub=is_sub, is_sup=is_sup)
            run.font.color.rgb = RGBColor(0, 80, 160) # 物理深蓝
        else:
            set_physics_font(run, '微软雅黑')
            run.font.color.rgb = RGBColor(40, 40, 40)

# ==========================================
# 3. 核心算法：Word 物理序解析 (解决公式乱序)
# ==========================================
def extract_docx_orderly(doc_stream, tmp_dir):
    """
    按 XML 树的物理顺序提取 Word 内容，确保图片和公式 100% 对齐
    """
    doc = docx.Document(doc_stream)
    questions = []
    current_q = {"text": "", "imgs": []}

    for para in doc.paragraphs:
        para_stream = ""
        # 深度遍历段落底层的 XML 子节点流
        for child in para._element.getchildren():
            tag = child.tag
            # 1. 文字节点
            if tag.endswith('}r'): 
                for t_node in child.iter():
                    if t_node.tag.endswith('}t'):
                        # 检查上下标格式
                        rpr = t_node.getparent().getprevious()
                        if rpr is not None and rpr.find('.//{*}vertAlign') is not None:
                            val = rpr.find('.//{*}vertAlign').get('{*}val')
                            if val == 'subscript': para_stream += "_"
                            elif val == 'superscript': para_stream += "^"
                        para_stream += t_node.text if t_node.text else ""
            
            # 2. 数学公式节点
            elif tag.endswith('}oMath'):
                for m_node in child.iter():
                    if m_node.tag.endswith('}t'):
                        is_sub = any('sSub' in p.tag for p in m_node.iterancestors())
                        is_sup = any('sSup' in p.tag for p in m_node.iterancestors())
                        t = m_node.text if m_node.text else ""
                        para_stream += f"_{t}" if is_sub else (f"^{t}" if is_sup else t)
            
            # 3. 图片节点 (解决图片消失)
            elif tag.endswith('}drawing'):
                rids = re.findall(r'r:embed="([^"]+)"', child.xml)
                for rid in rids:
                    try:
                        rel = doc.part.related_parts[rid]
                        img_name = f"w_{int(time.time()*1000)}_{rid}.png"
                        img_path = os.path.join(tmp_dir, img_name)
                        with open(img_path, "wb") as f: f.write(rel.blob)
                        current_q["imgs"].append(img_path)
                    except: pass

        txt = para_stream.strip()
        if not txt: continue
        
        # 题号切分逻辑
        if re.match(r'^\s*(\d+[\.．、]|\(\d+\))', txt):
            if current_q["text"]: questions.append(current_q)
            current_q = {"text": txt + "\n", "imgs": []}
        else:
            current_q["text"] += txt + "\n"

    if current_q["text"]: questions.append(current_q)
    return questions

# ==========================================
# 4. 核心算法：视觉投影排序 (解决 OCR 图片乱序)
# ==========================================
def visual_projection_sort(results):
    if not results: return ""
    results.sort(key=lambda x: x['position'][0][1])
    lines = []
    curr_line = [results[0]]
    for i in range(1, len(results)):
        h = abs(results[i]['position'][3][1] - results[i]['position'][0][1])
        if abs(results[i]['position'][0][1] - curr_line[-1]['position'][0][1]) < h * 0.6:
            curr_line.append(results[i])
        else:
            curr_line.sort(key=lambda x: x['position'][0][0])
            lines.append(curr_line)
            curr_line = [results[i]]
    curr_line.sort(key=lambda x: x['position'][0][0])
    lines.append(curr_line)
    
    full_text = ""
    for line in lines:
        row_str = "".join([f" {it['text']} " if it['type'] == 'formula' else it['text'] for it in line])
        full_text += row_str + "\n"
    return full_text

# ==========================================
# 5. PPT 渲染引擎 (核心调用处)
# ==========================================
def create_ppt_slide(prs, q_data, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(252, 254, 255); bg.line.fill.background()
    
    # 标题
    title_tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.6))
    p_title = title_tb.text_frame.paragraphs[0]
    p_title.text = f"物理教研 · 习题精讲 ({idx:02d})"
    set_physics_font(p_title.runs[0], font_name='微软雅黑', is_italic=False)
    p_title.runs[0].font.size, p_title.runs[0].font.bold = Pt(24), True

    has_img = len(q_data['imgs']) > 0
    box_w = Inches(8.3) if has_img else Inches(12.3)
    
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.1), box_w, Inches(6.0))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(230, 230, 235)
    
    # 渲染主体内容 (修正后的 render_rich_text)
    content_tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), box_w - Inches(0.4), Inches(5.6))
    tf = content_tb.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    render_rich_text(tf, q_data['text'])

    if has_img:
        for i, img_path in enumerate(list(set(q_data['imgs']))[:2]):
            try: slide.shapes.add_picture(img_path, Inches(8.9), Inches(1.2 + i*3.0), width=Inches(4.1))
            except: pass

# ==========================================
# 6. Streamlit UI (主逻辑)
# ==========================================
st.set_page_config(page_title="AI 物理教研全自动工作站", layout="wide")

st.markdown("<h2 style='text-align: center; color: #0070C0;'>🚀 AI 物理教研全自动工作站</h2>", unsafe_allow_html=True)

uploaded = st.file_uploader("📥 上传 物理图片/PDF/Word", accept_multiple_files=True, type=['jpg','png','jpeg','pdf','docx'])

if st.button("🚀 一键开启教研级排版生成", type="primary", use_container_width=True):
    if not uploaded:
        st.warning("请先上传文件")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        status = st.empty()
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            extracted_data = []
            for f in uploaded:
                ext = f.name.split('.')[-1].lower()
                status.info(f"正在深度解析: {f.name} ...")
                if ext == 'docx':
                    extracted_data.extend(extract_docx_orderly(io.BytesIO(f.read()), tmp_dir))
                elif ext in ['jpg', 'png', 'pdf']:
                    if ext == 'pdf':
                        doc = fitz.open(stream=f.read(), filetype="pdf")
                        for i in range(len(doc)):
                            pix = doc[i].get_pixmap(matrix=fitz.Matrix(2, 2))
                            p = os.path.join(tmp_dir, f"p_{i}.jpg")
                            pix.save(p)
                            res = p2t.recognize_mixed(p)
                            extracted_data.append({"text": visual_projection_sort(res), "imgs": []})
                    else:
                        p = os.path.join(tmp_dir, f.name)
                        with open(p, "wb") as tmp_file: tmp_file.write(f.read())
                        res = p2t.recognize_mixed(p)
                        extracted_data.append({"text": visual_projection_sort(res), "imgs": []})

            for idx, q in enumerate(extracted_data):
                create_ppt_slide(prs, q, idx + 1)
        
        buffer = io.BytesIO()
        prs.save(buffer)
        st.session_state['ready_ppt'] = buffer.getvalue()
        status.success("🎉 排版优化已完成！")

if 'ready_ppt' in st.session_state:
    st.write("---")
    st.download_button("⬇️ 下载优化版 PPT", st.session_state['ready_ppt'], "物理巅峰课件.pptx", use_container_width=True)
