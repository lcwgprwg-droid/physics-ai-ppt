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
    # analyzer_type='mfd' 负责定位公式在正文中的精准坐标，是解决“公式乱飞”的关键
    return Pix2Text(analyzer_type='mfd')

p2t = load_ai_engine()

# ==========================================
# 2. 巅峰排版渲染 (支持真·下标与专业字体)
# ==========================================
def set_physics_font(run, font_name='微软雅黑', is_italic=False, is_sub=False, is_sup=False):
    """设置中西文字体、斜体及真正的上下标属性"""
    run.font.name = font_name
    run.font.italic = is_italic
    run.font.subscript = is_sub
    run.font.superscript = is_sup
    
    # 物理公式中下标字号稍微缩小更美观
    if is_sub or is_sup:
        run.font.size = Pt(14)

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)

def render_smart_text(text_frame, raw_content):
    """
    智能流式渲染：解析文本中的 _(下标) 和 ^(上标)，并将变量应用 Times New Roman
    """
    p = text_frame.paragraphs[0]
    p.line_spacing = 1.3
    
    # 规范化 LaTeX 标记
    text = raw_content.replace('$', '').replace('\\(', '').replace('\\)', '')
    # 物理变量正则切分：汉字 | 变量标记 | 普通符号
    parts = re.split(r'([_][a-zA-Z0-9]+|[\^][a-zA-Z0-9]+)', text)
    
    for part in parts:
        if not part: continue
        is_sub = part.startswith('_')
        is_sup = part.startswith('^')
        clean_text = part[1:] if (is_sub or is_sup) else part
        
        run = p.add_run()
        run.text = clean_text
        run.font.size = Pt(20)
        
        # 物理量特征判断：字母、数字、下标标记自动转 Times New Roman 斜体
        if is_sub or is_sup or re.search(r'[a-zA-Z0-9\.\+\-\=\*×]', clean_text):
            set_physics_font(run, 'Times New Roman', is_italic=True, is_sub=is_sub, is_sup=is_sup)
            run.font.color.rgb = RGBColor(0, 80, 160) # 专业物理深蓝
        else:
            set_physics_font(run, '微软雅黑')
            run.font.color.rgb = RGBColor(40, 40, 40)

# ==========================================
# 3. 核心算法：图片/PDF 视觉重投影排序
# ==========================================
def visual_projection_sort(results):
    """
    将 OCR 散乱的文字/公式块按人眼阅读顺序缝合。
    解决公式块因为边框大小导致识别顺序错乱的问题。
    """
    if not results: return ""
    
    # 1. 按照 Y 轴坐标初步分行
    results.sort(key=lambda x: x['position'][0][1])
    lines = []
    curr_line = [results[0]]
    for i in range(1, len(results)):
        h = abs(results[i]['position'][3][1] - results[i]['position'][0][1])
        # 垂直距离在行高 60% 内的块视为同一行
        if abs(results[i]['position'][0][1] - curr_line[-1]['position'][0][1]) < h * 0.6:
            curr_line.append(results[i])
        else:
            # 行内按 X 轴从左到右排序
            curr_line.sort(key=lambda x: x['position'][0][0])
            lines.append(curr_line)
            curr_line = [results[i]]
    curr_line.sort(key=lambda x: x['position'][0][0])
    lines.append(curr_line)
    
    # 2. 缝合文本
    full_text = ""
    for line in lines:
        row_str = ""
        for item in line:
            # 公式块前后加空格防止排版粘连
            t = f" {item['text']} " if item['type'] == 'formula' else item['text']
            row_str += t
        full_text += row_str + "\n"
    return full_text

# ==========================================
# 4. 核心算法：Word 物理序节点遍历 (解决漏抓与错位)
# ==========================================
def extract_docx_orderly(doc_path, tmp_dir):
    """
    深度遍历 Word XML 骨架，按物理先后顺序提取文字、图片和公式。
    """
    doc = docx.Document(doc_path)
    questions = []
    current_q = {"text": "", "imgs": []}

    for para in doc.paragraphs:
        para_stream = ""
        # 遍历段落底层的 XML 子节点流 (w:r, m:oMath, w:drawing)
        for child in para._element.getchildren():
            tag = child.tag
            # 1. 处理文字节点 (含下标格式)
            if tag.endswith('}r'): 
                for t_node in child.iter():
                    if t_node.tag.endswith('}t'):
                        # 检查 XML 属性中的上下标
                        rpr = t_node.getparent().getprevious()
                        if rpr is not None and rpr.find('.//{*}vertAlign') is not None:
                            val = rpr.find('.//{*}vertAlign').get('{*}val')
                            if val == 'subscript': para_stream += "_"
                            elif val == 'superscript': para_stream += "^"
                        para_stream += t_node.text if t_node.text else ""
            
            # 2. 处理 Office Math 数学公式 (m:oMath)
            elif tag.endswith('}oMath'):
                for m_node in child.iter():
                    if m_node.tag.endswith('}t'):
                        is_sub = any('sSub' in p.tag for p in m_node.iterancestors())
                        is_sup = any('sSup' in p.tag for p in m_node.iterancestors())
                        t = m_node.text if m_node.text else ""
                        para_stream += f"_{t}" if is_sub else (f"^{t}" if is_sup else t)
            
            # 3. 处理插图 (w:drawing)
            elif tag.endswith('}drawing'):
                rids = re.findall(r'r:embed="([^"]+)"', child.xml)
                for rid in rids:
                    try:
                        rel = doc.part.related_parts[rid]
                        img_name = f"w_{int(time.time()*1000)}_{rid}.png"
                        img_p = os.path.join(tmp_dir, img_name)
                        with open(img_p, "wb") as f: f.write(rel.blob)
                        current_q["imgs"].append(img_p)
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
# 5. PPT 渲染引擎
# ==========================================
def create_ppt_slide(prs, q_data, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 绘制雅致背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(252, 254, 255); bg.line.fill.background()
    
    # 顶部装饰条
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.4), Inches(0.08), Inches(0.5))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    
    # 标题
    title_tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(10), Inches(0.6))
    p_title = title_tb.text_frame.paragraphs[0]
    p_title.text = f"物理教研 · 习题精讲 ({idx:02d})"
    set_physics_font(p_title.runs[0], font_name='微软雅黑', is_italic=False)
    p_title.runs[0].font.size, p_title.runs[0].font.bold = Pt(24), True

    # 布局：左文右图
    has_img = len(q_data['imgs']) > 0
    box_w = Inches(8.3) if has_img else Inches(12.3)
    
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.1), box_w, Inches(6.0))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(230, 230, 235)
    
    # 渲染主体内容
    content_tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), box_w - Inches(0.4), Inches(5.6))
    tf = content_tb.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    render_rich_text(tf, q_data['text'])

    # 图片处理
    if has_img:
        # 只取前2张图防止堆叠
        for i, img_path in enumerate(list(set(q_data['imgs']))[:2]):
            try: slide.shapes.add_picture(img_path, Inches(8.9), Inches(1.2 + i*3.0), width=Inches(4.1))
            except: pass

# ==========================================
# 6. Streamlit UI (保持 Session 状态)
# ==========================================
st.set_page_config(page_title="AI 物理教研全自动工作站", layout="wide")

st.markdown("""
    <div style='text-align: center;'>
        <h2 style='color: #0070C0;'>⚛️ AI 物理教研全自动工作站</h2>
        <p style='color: #666;'>针对 2024 高考题深度优化：解决 Word 公式错位与图片丢失</p>
    </div>
""", unsafe_allow_html=True)

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
                    # 采用物理流式节点提取逻辑
                    extracted_data.extend(extract_docx_orderly(io.BytesIO(f.read()), tmp_dir))
                elif ext in ['jpg', 'png', 'pdf']:
                    if ext == 'pdf':
                        doc = fitz.open(stream=f.read(), filetype="pdf")
                        for i in range(len(doc)):
                            pix = doc[i].get_pixmap(matrix=fitz.Matrix(2.2, 2.2))
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
