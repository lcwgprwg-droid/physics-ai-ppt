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
# 1. 初始化 Pix2Text (物理公式版)
# ==========================================
@st.cache_resource
def load_p2t():
    # 开启公式检测，识别 T_1, V_2 等
    return Pix2Text(analyzer_type='mfd')

p2t = load_p2t()

# ==========================================
# 2. 字体设置补丁 (修复 AttributeError)
# ==========================================
def apply_safe_font(run, font_name, is_italic=False, is_sub=False, is_sup=False):
    run.font.name = font_name
    run.font.italic = is_italic
    run.font.subscript = is_sub
    run.font.superscript = is_sup
    
    # 强制注入中文字体设置
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)

# ==========================================
# 3. 智能文本排版渲染 (处理下标和斜体)
# ==========================================
def render_physics_content(text_frame, raw_text):
    """
    将带有 _1(下标) 或 ^2(上标) 标记的文本渲染进 PPT
    """
    p = text_frame.paragraphs[0]
    p.line_spacing = 1.3
    
    # 清理识别标记
    text = raw_text.replace('$', '').replace('\\(', '').replace('\\)', '').replace('{', '').replace('}', '')
    
    # 使用正则切分：汉字 | 变量 | 格式标记(_1, ^2)
    parts = re.split(r'([_][a-zA-Z0-9]+|[\^][a-zA-Z0-9]+)', text)
    
    for part in parts:
        if not part: continue
        
        is_sub = part.startswith('_')
        is_sup = part.startswith('^')
        clean_val = part[1:] if (is_sub or is_sup) else part
        
        run = p.add_run()
        run.text = clean_val
        run.font.size = Pt(20)
        
        # 物理量判断 (数字、英文字母、符号块)
        if is_sub or is_sup or re.search(r'[a-zA-Z0-9\.\+\-\=\*×]', clean_val):
            apply_safe_font(run, 'Times New Roman', is_italic=True, is_sub=is_sub, is_sup=is_sup)
            run.font.color.rgb = RGBColor(0, 112, 192) # 变量用蓝色
        else:
            apply_safe_font(run, '微软雅黑')
            run.font.color.rgb = RGBColor(40, 40, 40)

# ==========================================
# 4. Word 提取引擎 (修复图片消失和变量缺失)
# ==========================================
def extract_word_physics(file_stream, tmp_dir):
    doc = docx.Document(file_stream)
    questions = []
    current_q = {"text": "", "imgs": []}

    for para in doc.paragraphs:
        para_text = ""
        # 1. 提取文字和格式 (下标转换)
        for run in para.runs:
            txt = run.text
            if run.font.subscript:
                para_text += f"_{txt}"
            elif run.font.superscript:
                para_text += f"^{txt}"
            else:
                para_text += txt
        
        # 2. 备选方案：如果 Paragraph.text 里有内容但 runs 没抓到 (针对 Math 对象)
        if not para_text.strip() and para.text.strip():
            para_text = para.text

        # 3. 提取图片 (全扫描段落 XML)
        xml_str = para._element.xml
        # 查找所有图片引用 ID
        rIds = re.findall(r'r:embed="([^"]+)"', xml_str)
        for rid in rIds:
            try:
                img_part = doc.part.related_parts[rid]
                # 使用唯一时间戳命名防止图片覆盖消失
                img_name = f"wimg_{int(time.time()*1000)}_{rid}.png"
                img_p = os.path.join(tmp_dir, img_name)
                with open(img_p, "wb") as f:
                    f.write(img_part.blob)
                current_q["imgs"].append(img_p)
            except:
                continue

        line = para_text.strip()
        if not line: continue

        # 4. 题号切分逻辑
        if re.match(r'^\s*(\d+[\.．、]|\(\d+\))', line):
            if current_q["text"]: questions.append(current_q)
            current_q = {"text": line + "\n", "imgs": []}
        else:
            current_q["text"] += line + "\n"

    if current_q["text"]: questions.append(current_q)
    return questions

# ==========================================
# 5. PPT 渲染布局
# ==========================================
def create_ppt_slide(prs, q, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # 标题设计
    title_tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.6))
    p_title = title_tb.text_frame.paragraphs[0]
    p_title.text = f"2024 高考真题精讲 - 第 {idx} 题"
    apply_safe_font(p_title.runs[0], '微软雅黑')
    p_title.runs[0].font.bold = True
    p_title.runs[0].font.size = Pt(24)

    # 布局判断：有图 6:4 分屏，无图全屏
    has_img = len(q['imgs']) > 0
    box_w = Inches(8.5) if has_img else Inches(12.3)
    
    # 内容卡片
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.1), box_w, Inches(5.9))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 230, 240)
    
    # 写入题干
    txt_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), box_w-Inches(0.4), Inches(5.5))
    tf = txt_box.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    render_physics_content(tf, q['text'])

    # 插入图片 (确保不消失)
    if has_img:
        # 只取前两张图防止堆叠混乱
        for i, img_path in enumerate(list(set(q['imgs']))[:2]):
            try:
                slide.shapes.add_picture(img_path, Inches(9.0), Inches(1.2 + i*3.0), width=Inches(4.0))
            except:
                pass

# ==========================================
# 6. Streamlit 界面
# ==========================================
st.set_page_config(page_title="AI 物理教研巅峰工作站", layout="wide")
st.markdown("<h1 style='text-align:center; color:#0070C0;'>🚀 AI 物理教研全自动工作站</h1>", unsafe_allow_html=True)

files = st.file_uploader("📥 上传 物理图片/PDF/Word", accept_multiple_files=True, type=['jpg','png','jpeg','pdf','docx'])

if st.button("✨ 一键开启全自动排版生成", type="primary", use_container_width=True):
    if not files:
        st.warning("请上传文件")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        status = st.empty()
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            all_data = []
            for f in files:
                ext = f.name.split('.')[-1].lower()
                status.info(f"正在处理: {f.name} ...")
                
                if ext == 'docx':
                    all_data.extend(extract_word_physics(io.BytesIO(f.read()), tmp_dir))
                elif ext in ['jpg', 'png', 'jpeg', 'pdf']:
                    if ext == 'pdf':
                        pdf = fitz.open(stream=f.read(), filetype="pdf")
                        for i in range(len(pdf)):
                            pix = pdf[i].get_pixmap(matrix=fitz.Matrix(2, 2))
                            p = os.path.join(tmp_dir, f"p_{i}.jpg")
                            pix.save(p)
                            res = p2t.recognize_mixed(p)
                            txt = "".join([f" {it['text']} " if it['type']=='formula' else it['text'] for it in res])
                            all_data.append({"text": txt, "imgs": []})
                    else:
                        p = os.path.join(tmp_dir, f.name)
                        with open(p, "wb") as tmp_file: tmp_file.write(f.read())
                        res = p2t.recognize_mixed(p)
                        txt = "".join([f" {it['text']} " if it['type']=='formula' else it['text'] for it in res])
                        all_data.append({"text": txt, "imgs": []})
            
            # 渲染 PPT
            for i, q in enumerate(all_data):
                create_ppt_slide(prs, q, i + 1)
        
        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        st.session_state['ppt_final'] = ppt_buffer.getvalue()
        status.success(f"🎉 处理完成！共提取 {len(all_data)} 道题。")

if 'ppt_final' in st.session_state:
    st.write("---")
    st.download_button("⬇️ 下载 AI 巅峰排版 PPT", st.session_state['ppt_final'], "物理教研课件.pptx", use_container_width=True)
