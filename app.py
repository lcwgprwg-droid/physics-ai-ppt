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
    # 结合 GitHub 智慧：analyzer_type='mfd' 负责定位公式在正文中的精准坐标
    return Pix2Text(analyzer_type='mfd')

p2t = load_ai_engine()

# ==========================================
# 2. 巅峰排版渲染 (支持真·下标与物理字体)
# ==========================================
def set_physics_font(run, font_name='微软雅黑', is_italic=False, is_sub=False, is_sup=False):
    """底层字体设置：支持中西文、斜体及真·上下标"""
    run.font.name = font_name
    run.font.italic = is_italic
    run.font.subscript = is_sub
    run.font.superscript = is_sup
    
    # 强制设置东亚字体
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)

def render_rich_text(text_frame, raw_content):
    """
    智能流式渲染：将文本中的 _1 转换为真正的 PPT 下标，将 ^2 转换为上标
    并将物理变量自动应用斜体 Times New Roman
    """
    p = text_frame.paragraphs[0]
    p.line_spacing = 1.3
    
    # 清理并规范化标记
    text = raw_content.replace('$', '').replace('\\(', '').replace('\\)', '')
    text = text.replace('{', '').replace('}', '') # 处理 T_{1} 为 T_1
    
    # 正则切分：识别 汉字 | 变量 | 下标标记 | 上标标记
    # 匹配模式：_紧跟字符 或 ^紧跟字符
    parts = re.split(r'([_][a-zA-Z0-9]+|[\^][a-zA-Z0-9]+)', text)
    
    for part in parts:
        if not part: continue
        
        is_sub = part.startswith('_')
        is_sup = part.startswith('^')
        clean_text = part[1:] if (is_sub or is_sup) else part
        
        run = p.add_run()
        run.text = clean_content = clean_text
        run.font.size = Pt(20)
        
        # 物理量特征判断：英文字母、数字、符号或已经是上下标
        if is_sub or is_sup or re.search(r'[a-zA-Z0-9\.\+\-\=\*×]', clean_content):
            set_physics_font(run, 'Times New Roman', is_italic=True, is_sub=is_sub, is_sup=is_sup)
            run.font.color.rgb = RGBColor(0, 80, 160) # 物理蓝
        else:
            set_physics_font(run, '微软雅黑')
            run.font.color.rgb = RGBColor(40, 40, 40)

# ==========================================
# 3. 核心算法 A：OCR 视觉重投影排序 (解决图片公式乱跳)
# ==========================================
def visual_projection_sort(results):
    """
    将 OCR 散乱的碎片按人眼阅读顺序重新缝合
    逻辑：先分行，行内按 X 轴排序
    """
    if not results: return ""
    
    # 1. 按照 Y 轴坐标初步分行 (容差为行高的 50%)
    results.sort(key=lambda x: x['position'][0][1])
    lines = []
    if results:
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
    
    # 2. 精准缝合文本
    full_text = ""
    for line in lines:
        row_str = ""
        for item in line:
            t = item['text']
            # 如果是公式块，给前后加微小空格防止粘连
            row_str += f" {t} " if item['type'] == 'formula' else t
        full_text += row_str + "\n"
    return full_text

# ==========================================
# 4. 核心算法 B：Word 节点物理序遍历 (解决 Word 公式乱序)
# ==========================================
def extract_docx_in_order(doc_path, tmp_dir):
    """
    按 XML 树的物理顺序提取 Word 内容，确保公式插入位置 100% 正确
    """
    doc = docx.Document(doc_path)
    questions = []
    current_q = {"text": "", "imgs": []}

    for para in doc.paragraphs:
        para_stream = ""
        # 遍历段落底层的 XML 子节点流
        for child in para._element.getchildren():
            tag = child.tag
            # 1. 处理普通文本节点 (w:r)
            if tag.endswith('}r'): 
                for t_node in child.iter():
                    if t_node.tag.endswith('}t'):
                        # 检查是否有下标属性
                        rpr = t_node.getparent().getprevious()
                        if rpr is not None and rpr.find('.//{*}vertAlign') is not None:
                            val = rpr.find('.//{*}vertAlign').get('{*}val')
                            if val == 'subscript': para_stream += "_"
                            elif val == 'superscript': para_stream += "^"
                        para_stream += t_node.text if t_node.text else ""
            
            # 2. 处理 Office Math 数学公式节点 (m:oMath)
            elif tag.endswith('}oMath'):
                for m_node in child.iter():
                    if m_node.tag.endswith('}t'):
                        # 识别公式内部的上下标容器
                        is_sub = any('sSub' in p.tag for p in m_node.iterancestors())
                        is_sup = any('sSup' in p.tag for p in m_node.iterancestors())
                        t = m_node.text if m_node.text else ""
                        if is_sub: para_stream += f"_{t}"
                        elif is_sup: para_stream += f"^{t}"
                        else: para_stream += t
            
            # 3. 处理图片 (w:drawing)
            elif tag.endswith('}drawing'):
                rids = re.findall(r'r:embed="([^"]+)"', child.xml)
                for rid in rids:
                    try:
                        rel = doc.part.related_parts[rid]
                        img_p = os.path.join(tmp_dir, f"w_{int(time.time())}_{rid}.png")
                        with open(img_p, "wb") as f: f.write(rel.blob)
                        current_q["imgs"].append(img_p)
                    except: pass

        txt = para_stream.strip()
        if not txt: continue
        
        # 题号切分逻辑 (识别 1. 或 (1) 等)
        if re.match(r'^\s*(\d+[\.．、]|\(\d+\))', txt):
            if current_q["text"]: questions.append(current_q)
            current_q = {"text": txt + "\n", "imgs": []}
        else:
            current_q["text"] += txt + "\n"

    if current_q["text"]: questions.append(current_q)
    return questions

# ==========================================
# 5. PPT 渲染控制
# ==========================================
def create_slide_page(prs, q_data, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 绘制背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(250, 252, 255); bg.line.fill.background()
    
    # 标题
    title_tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.6))
    p_title = title_tb.text_frame.paragraphs[0]
    p_title.text = f"物理核心素养·习题讲评 - 第 {idx:02d} 题"
    p_title.font.bold = True
    set_physics_font(p_title.runs[0], '微软雅黑')

    # 布局：带图左文右图，无图全屏
    has_img = len(q_data['imgs']) > 0
    box_w = Inches(8.3) if has_img else Inches(12.3)
    
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.1), box_w, Inches(5.8))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 230, 240)
    
    content_tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), box_w - Inches(0.4), Inches(5.4))
    tf = content_tb.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    # 渲染富文本内容
    render_rich_text(tf, q_data['text'])

    # 图片处理
    if has_img:
        for i, img_path in enumerate(list(set(q_data['imgs']))[:2]):
            try: slide.shapes.add_picture(img_path, Inches(8.9), Inches(1.2 + i*3.0), width=Inches(4.1))
            except: pass

# ==========================================
# 6. Streamlit 界面逻辑
# ==========================================
st.set_page_config(page_title="AI 物理教研巅峰工作站", layout="wide")

st.markdown("""
    <div style='text-align: center;'>
        <h1 style='color: #0070C0;'>🚀 AI 物理教研全自动工作站</h1>
        <p style='color: #666;'>深度解决 Word 公式乱序与图片排版混乱问题</p>
    </div>
""", unsafe_allow_html=True)

files = st.file_uploader("📥 上传 物理图片/PDF/Word", accept_multiple_files=True, type=['jpg','png','jpeg','pdf','docx'])

if st.button("✨ 一键开启巅峰排版", type="primary", use_container_width=True):
    if not files:
        st.warning("请先上传文件")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        status = st.empty()
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            all_questions = []
            for f in files:
                ext = f.name.split('.')[-1].lower()
                status.info(f"正在深度解析: {f.name} ...")
                
                if ext == 'docx':
                    # 采用 XML 流式节点顺序提取
                    all_questions.extend(extract_docx_in_order(io.BytesIO(f.read()), tmp_dir))
                elif ext in ['jpg', 'png', 'pdf']:
                    if ext == 'pdf':
                        doc = fitz.open(stream=f.read(), filetype="pdf")
                        for i in range(len(doc)):
                            pix = doc[i].get_pixmap(matrix=fitz.Matrix(2.5, 2.5)) # 提高清晰度
                            p = os.path.join(tmp_dir, f"pdf_{i}.jpg")
                            pix.save(p)
                            res = p2t.recognize_mixed(p)
                            all_questions.append({"text": visual_projection_sort(res), "imgs": []})
                    else:
                        p = os.path.join(tmp_dir, f.name)
                        with open(p, "wb") as tmp_file: tmp_file.write(f.read())
                        res = p2t.recognize_mixed(p)
                        all_questions.append({"text": visual_projection_sort(res), "imgs": []})

            # 生成 PPT 页面
            for idx, q in enumerate(all_questions):
                create_slide_page(prs, q, idx + 1)
        
        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        st.session_state['final_ppt'] = ppt_buffer.getvalue()
        status.success(f"🎉 处理完成！共提取 {len(all_questions)} 道物理题。")

if 'final_ppt' in st.session_state:
    st.write("---")
    st.download_button("⬇️ 下载优化版 PPT", st.session_state['final_ppt'], "物理巅峰课件.pptx", use_container_width=True)
