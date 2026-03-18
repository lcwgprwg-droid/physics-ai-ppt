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
def set_physics_font(run, cn_font='微软雅黑', en_font='Times New Roman'):
    run.font.name = en_font # 默认字体设为西文
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.append(rFonts)
    
    rFonts.set(qn('w:eastAsia'), cn_font) # 中文部分用微软雅黑
    rFonts.set(qn('w:ascii'), en_font)    # 西文部分用 Times New Roman
    rFonts.set(qn('w:hAnsi'), en_font)

def add_smart_text(text_frame, raw_text):
    """
    智能排版逻辑：
    - 中文：微软雅黑 (Yahei)
    - 物理量/数字/公式：Times New Roman + 斜体 + 深蓝色
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
        
        # 判断是否为物理变量/数字/LaTeX符号块
        if re.match(r'^[a-zA-Z0-9\.\_\^\{\}\+\-\=\(\)\/\s]+$', chunk.strip()):
            set_font_style(run, 'Times New Roman', is_italic=True)
            run.font.color.rgb = RGBColor(0, 80, 160) # 物理专业深蓝
        else:
            set_font_style(run, '微软雅黑')
            run.font.color.rgb = RGBColor(35, 35, 35)

# ==========================================
# 4. PPT 页面构建引擎
# ==========================================
def create_physics_slide(prs, q_text, imgs, idx):
    """创建专业美观的 PPT 页面"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 极简雅致背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(252, 254, 255); bg.line.fill.background()
    
    # 标题栏修饰
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.35), Inches(0.08), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    
    tb_title = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(10), Inches(0.7))
    p_title = tb_title.text_frame.paragraphs[0]
    p_title.text = f"物理核心素养·习题讲评 第 {idx:02d} 题"
    set_font_style(p_title.runs[0], '微软雅黑')
    p_title.runs[0].font.size, p_title.runs[0].font.bold = Pt(24), True

    # 布局判断：是否有图
    has_img = len(imgs) > 0
    box_w = Inches(8.5) if has_img else Inches(12.3)
    
    # 题目主体卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.15), box_w, Inches(5.8))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(225, 230, 235)

    # 写入智能内容
    content_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), box_w - Inches(0.4), Inches(5.4))
    tf = content_box.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    add_smart_text(tf, q_text)

    # 如果有配图，显示在右侧
    if has_img:
        for i, img_path in enumerate(imgs[:2]): # 每页最多显示2张图
            try: slide.shapes.add_picture(img_path, Inches(9.2), Inches(1.15 + i*3.0), width=Inches(3.8))
            except: pass

# ==========================================
# 5. Streamlit 全自动流程逻辑
# ==========================================
st.set_page_config(page_title="AI 物理教研课件生成器", layout="wide", page_icon="⚛️")

st.markdown("""
<div style='text-align: center;'>
    <h1 style='color: #0070C0;'>🚀 AI 物理教研全自动工作站</h1>
    <p style='color: #666;'>支持 <b>图片 / PDF / Word</b>，专业优化物理符号排版。</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📥 拖拽上传资料 (支持多选)", accept_multiple_files=True, type=['jpg', 'png', 'jpeg', 'pdf', 'docx'])

if st.button("✨ 一键开启 AI 深度教研排版", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("⚠️ 老师，请先上传文件哦！")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5) # 16:9
        
        status_info = st.empty()
        progress_bar = st.progress(0)
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            extracted_questions = []
            
            for f_idx, uploaded_file in enumerate(uploaded_files):
                ext = uploaded_file.name.split('.')[-1].lower()
                status_info.info(f"正在深度分析: {uploaded_file.name}")
                
                # --- 处理图片 ---
                if ext in ['jpg', 'png', 'jpeg']:
                    p = os.path.join(tmp_dir, uploaded_file.name)
                    with open(p, "wb") as f: f.write(uploaded_file.read())
                    # 使用 P2T 混合识别文字与公式
                    outs = p2t.recognize_mixed(p)
                    txt = "".join([o['text'] for o in outs])
                    extracted_questions.append({"text": txt, "imgs": []})
                
                # --- 处理 PDF ---
                elif ext == 'pdf':
                    pdf_doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                    for i in range(len(pdf_doc)):
                        page = pdf_doc.load_page(i)
                        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
                        img_path = os.path.join(tmp_dir, f"pdf_{f_idx}_p{i}.jpg")
                        pix.save(img_path)
                        outs = p2t.recognize_mixed(img_path)
                        txt = "".join([o['text'] for o in outs])
                        extracted_questions.append({"text": txt, "imgs": []})
                
                # --- 处理 Word (含图片提取) ---
                elif ext == 'docx':
                    doc_obj = docx.Document(io.BytesIO(uploaded_file.read()))
                    current_q = {"text": "", "imgs": []}
                    for para in doc_obj.paragraphs:
                        # 如果段落以数字开头，视为新题
                        if re.match(r'^\d+', para.text.strip()):
                            if current_q["text"]: extracted_questions.append(current_q)
                            current_q = {"text": para.text + "\n", "imgs": []}
                        else:
                            current_q["text"] += para.text + "\n"
                        
                        # 抓取段落中的图片
                        for run in para.runs:
                            if 'pic:pic' in run._element.xml:
                                rIds = re.findall(r'r:embed="([^"]+)"', run._element.xml)
                                for rId in rIds:
                                    img_part = doc_obj.part.related_parts[rId]
                                    img_p = os.path.join(tmp_dir, f"w_img_{rId}.png")
                                    with open(img_p, "wb") as f: f.write(img_part.blob)
                                    current_q["imgs"].append(img_p)
                    if current_q["text"]: extracted_questions.append(current_q)
                
                progress_bar.progress((f_idx + 1) / len(uploaded_files))

            # 开始渲染 PPT
            for idx, q in enumerate(extracted_questions):
                create_physics_slide(prs, q["text"], q["imgs"], idx + 1)

        # 存入 Session State 防止下载时因页面刷新导致数据丢失
        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        st.session_state['ready_ppt'] = ppt_io.getvalue()
        
        status_info.success(f"🎉 识别完成！共生成 {len(extracted_questions)} 页巅峰排版幻灯片。")
        st.balloons()

# ==========================================
# 6. 下载区域 (核心修复：锁定 Session)
# ==========================================
if 'ready_ppt' in st.session_state:
    st.write("---")
    st.download_button(
        label="⬇️ 点击下载 AI 生成的专业物理课件",
        data=st.session_state['ready_ppt'],
        file_name=f"物理习题精讲_{int(time.time())}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
