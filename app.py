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
# 1. 初始化 Pix2Text (针对物理公式深度优化)
# ==========================================
@st.cache_resource
def load_pix2text():
    # analyzer_type='mfd' 开启数学公式检测，能看懂 T_1, L_1, h 等物理变量
    return Pix2Text(analyzer_type='mfd')

p2t = load_pix2text()

# ==========================================
# 2. 巅峰排版辅助函数 (中英字体分流)
# ==========================================
def set_font_style(run, font_name, is_italic=False):
    """稳健的中西文字体设置方案，修复 AttributeError"""
    run.font.name = font_name
    run.font.italic = is_italic
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)

def add_smart_text(text_frame, raw_text):
    """
    智能排版：
    - 中文：微软雅黑
    - 物理量/数字/公式符号：Times New Roman + 斜体 + 深蓝色
    """
    p = text_frame.paragraphs[0]
    p.line_spacing = 1.3
    
    # 清理识别标记
    clean_text = raw_text.replace('$', '').replace('\\(', '').replace('\\)', '')
    
    # 按照中文和非中文块进行切割
    chunks = re.findall(r'([\u4e00-\u9fa5\s，。？！：；（）《》]+|[^\u4e00-\u9fa5，。？！：；（）《》]+)', clean_text)
    
    for chunk in chunks:
        if not chunk: continue
        run = p.add_run()
        run.text = chunk
        run.font.size = Pt(20)
        
        # 判断是否为物理变量/数字/符号块 (包含下划线用于显示下标)
        if re.match(r'^[a-zA-Z0-9\.\_\^\{\}\+\-\=\(\)\/\s]+$', chunk.strip()):
            set_font_style(run, 'Times New Roman', is_italic=True)
            run.font.color.rgb = RGBColor(0, 80, 160) # 物理专业深蓝
        else:
            set_font_style(run, '微软雅黑')
            run.font.color.rgb = RGBColor(35, 35, 35)

# ==========================================
# 3. PPT 页面构建引擎
# ==========================================
def create_physics_slide(prs, q_text, imgs, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(252, 254, 255); bg.line.fill.background()
    
    # 装饰条
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.35), Inches(0.08), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    
    tb_title = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(10), Inches(0.7))
    p_title = tb_title.text_frame.paragraphs[0]
    p_title.text = f"物理核心素养·习题讲评 第 {idx:02d} 题"
    set_font_style(p_title.runs[0], '微软雅黑')
    p_title.runs[0].font.size, p_title.runs[0].font.bold = Pt(24), True

    has_img = len(imgs) > 0
    box_w = Inches(8.5) if has_img else Inches(12.3)
    
    # 内容卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.15), box_w, Inches(5.8))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(225, 230, 235)

    content_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), box_w - Inches(0.4), Inches(5.4))
    tf = content_box.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    add_smart_text(tf, q_text)

    if has_img:
        for i, img_path in enumerate(imgs[:2]):
            try: slide.shapes.add_picture(img_path, Inches(9.2), Inches(1.15 + i*3.0), width=Inches(3.8))
            except: pass

# ==========================================
# 4. Streamlit 主程序
# ==========================================
st.set_page_config(page_title="AI 物理教研课件生成器", layout="wide", page_icon="⚛️")

st.markdown("""
<div style='text-align: center;'>
    <h1 style='color: #0070C0;'>🚀 AI 物理教研全自动工作站</h1>
    <p style='color: #666;'>针对 <b>Word 公式与下标</b> 深度优化的巅峰排版版本</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📥 上传资料 (图片/PDF/Word)", accept_multiple_files=True, type=['jpg', 'png', 'jpeg', 'pdf', 'docx'])

if st.button("✨ 一键开启 AI 深度教研排版", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("⚠️ 请先上传文件哦！")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        status_info = st.empty()
        progress_bar = st.progress(0)
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            extracted_questions = []
            
            for f_idx, uploaded_file in enumerate(uploaded_files):
                ext = uploaded_file.name.split('.')[-1].lower()
                status_info.info(f"正在深度分析: {uploaded_file.name}")
                
                # --- A. 图片与 PDF 识别 ---
                if ext in ['jpg', 'png', 'jpeg', 'pdf']:
                    if ext == 'pdf':
                        pdf_doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                        for i in range(len(pdf_doc)):
                            pix = pdf_doc[i].get_pixmap(matrix=fitz.Matrix(2, 2))
                            img_p = os.path.join(tmp_dir, f"pdf_{f_idx}_{i}.jpg")
                            pix.save(img_p)
                            outs = p2t.recognize_mixed(img_p)
                            txt = "".join([o['text'] for o in outs])
                            extracted_questions.append({"text": txt, "imgs": []})
                    else:
                        img_p = os.path.join(tmp_dir, uploaded_file.name)
                        with open(img_p, "wb") as f: f.write(uploaded_file.read())
                        outs = p2t.recognize_mixed(img_p)
                        txt = "".join([o['text'] for o in outs])
                        extracted_questions.append({"text": txt, "imgs": []})
                
                # --- B. Word 深度符号解析 (核心优化点) ---
                elif ext == 'docx':
                    doc_obj = docx.Document(io.BytesIO(uploaded_file.read()))
                    current_q = {"text": "", "imgs": []}
                    
                    for para in doc_obj.paragraphs:
                        para_rich_text = ""
                        # 穿透获取 Word 公式对象
                        math_elements = para._element.xpath('.//m:oMath')
                        if math_elements:
                            # 如果有公式节点，遍历 XML 获取所有文本
                            for node in para._element.iter():
                                if node.tag.endswith('t'): para_rich_text += node.text
                        else:
                            # 处理普通文字与下标属性
                            for run in para.runs:
                                if run.font.subscript: para_rich_text += f"_{run.text}"
                                elif run.font.superscript: para_rich_text += f"^{run.text}"
                                else: para_rich_text += run.text
                        
                        text_line = para_rich_text.strip()
                        if not text_line: continue
                        
                        # 识别题号开启新题
                        if re.match(r'^\s*\d+[\.．、]', text_line):
                            if current_q["text"]: extracted_questions.append(current_q)
                            current_q = {"text": text_line + "\n", "imgs": []}
                        elif current_q:
                            current_q["text"] += text_line + "\n"
                        
                        # 抓取图片
                        for run in para.runs:
                            if 'pic:pic' in run._element.xml:
                                rIds = re.findall(r'r:embed="([^"]+)"', run._element.xml)
                                for rId in rIds:
                                    try:
                                        img_part = doc_obj.part.related_parts[rId]
                                        img_p = os.path.join(tmp_dir, f"w_img_{rId}.png")
                                        with open(img_p, "wb") as f: f.write(img_part.blob)
                                        current_q["imgs"].append(img_p)
                                    except: pass
                    
                    if current_q["text"]: extracted_questions.append(current_q)
                
                progress_bar.progress((f_idx + 1) / len(uploaded_files))

            for idx, q in enumerate(extracted_questions):
                create_physics_slide(prs, q["text"], q["imgs"], idx + 1)

        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        st.session_state['ready_ppt'] = ppt_io.getvalue()
        status_info.success(f"🎉 识别完成！共提取 {len(extracted_questions)} 道题目。")
        st.balloons()

# 下载区域
if 'ready_ppt' in st.session_state:
    st.write("---")
    st.download_button(
        label="⬇️ 点击下载生成的专业物理课件",
        data=st.session_state['ready_ppt'],
        file_name=f"物理教研课件_{int(time.time())}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
