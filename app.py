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
# 1. 初始化 Pix2Text (针对图像/PDF中的公式)
# ==========================================
@st.cache_resource
def load_pix2text():
    return Pix2Text(analyzer_type='mfd')

p2t = load_pix2text()

# ==========================================
# 2. 巅峰排版辅助函数
# ==========================================
def set_font_style(run, font_name, is_italic=False):
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
    - 变量/数字/单位：Times New Roman + 斜体
    - 处理上标(^)和下标(_)
    """
    p = text_frame.paragraphs[0]
    p.line_spacing = 1.2
    
    # 格式清理
    clean_text = raw_text.replace('$', '').replace('\\(', '').replace('\\)', '')
    
    # 按中文/非中文分段
    chunks = re.findall(r'([\u4e00-\u9fa5\s，。？！：；（）《》]+|[^\u4e00-\u9fa5，。？！：；（）《》]+)', clean_text)
    
    for chunk in chunks:
        if not chunk: continue
        run = p.add_run()
        run.text = chunk
        run.font.size = Pt(20)
        
        # 物理量/公式特征判断
        if re.search(r'[a-zA-Z0-9\.\_\^\{\}\+\-\=\(\)\/\*×]', chunk):
            set_font_style(run, 'Times New Roman', is_italic=True)
            run.font.color.rgb = RGBColor(0, 80, 160)
        else:
            set_font_style(run, '微软雅黑')
            run.font.color.rgb = RGBColor(30, 30, 30)

# ==========================================
# 3. Word 深度解析引擎 (解决公式和图片漏抓)
# ==========================================
def get_docx_advanced_content(doc_obj, tmp_dir):
    """
    全方位解析 Word：
    1. 递归 XML 捕获所有文本（包括 Math 对象内的文字）
    2. 自动标记下标(_)和上标(^)
    3. 全局扫描段落内的所有图片关联(rId)
    """
    extracted = []
    current_q = {"text": "", "imgs": []}

    # Word 数学命名空间
    M_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
    W_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

    for para in doc_obj.paragraphs:
        line_text = ""
        
        # --- A. 深度解析 XML 获取文本与公式 ---
        for node in para._element.iter():
            # 1. 处理标准文本
            if node.tag == f'{W_NS}t':
                # 检查父级 rPr 是否有上下标属性
                rpr = node.getparent().getprevious() if node.getparent() is not None else None
                text_part = node.text if node.text else ""
                
                # 尝试通过属性判断上下标
                if rpr is not None:
                    va = rpr.find(f'.//{W_NS}vertAlign')
                    if va is not None:
                        val = va.get(f'{W_NS}val')
                        if val == 'subscript': text_part = f"_{text_part}"
                        elif val == 'superscript': text_part = f"^{text_part}"
                line_text += text_part

            # 2. 处理 Office Math 数学文本 (T_1, V_2 等)
            elif node.tag == f'{M_NS}t':
                # 检查是否处于下标结构中
                is_sub = any(p.tag == f'{M_NS}sSub' for p in node.iterancestors())
                is_sup = any(p.tag == f'{M_NS}sSup' for p in node.iterancestors())
                
                mt_text = node.text if node.text else ""
                if is_sub: line_text += f"_{mt_text}"
                elif is_sup: line_text += f"^{mt_text}"
                else: line_text += mt_text

        # --- B. 全局扫描图片 (解决漏抓) ---
        # 扫描段落 XML 中出现的所有 rId
        xml_str = para._element.xml
        rIds = set(re.findall(r'r:embed="([^"]+)"', xml_str) + re.findall(r'r:id="([^"]+)"', xml_str))
        for rid in rIds:
            try:
                rel = doc_obj.part.related_parts[rid]
                if "image" in rel.content_type:
                    img_path = os.path.join(tmp_dir, f"wimg_{int(time.time()*1000)}_{rid}.png")
                    with open(img_path, "wb") as f: f.write(rel.blob)
                    current_q["imgs"].append(img_path)
            except: pass

        # --- C. 题目切分 ---
        content = line_text.strip()
        if not content: continue
        
        # 识别题号（如 14. 或 (1)）
        if re.match(r'^\s*(\d+[\.．、]|\(\d+\))', content):
            if current_q["text"]: extracted.append(current_q)
            current_q = {"text": content + "\n", "imgs": current_q["imgs"] if not current_q["text"] else []}
        else:
            current_q["text"] += content + "\n"

    if current_q["text"]: extracted.append(current_q)
    return extracted

# ==========================================
# 4. PPT 页面渲染
# ==========================================
def create_physics_slide(prs, q_text, imgs, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(252, 254, 255); bg.line.fill.background()
    # 标题栏
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.35), Inches(0.1), Inches(0.55))
    bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(0, 112, 192); bar.line.fill.background()
    tb_title = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(10), Inches(0.7))
    p_title = tb_title.text_frame.paragraphs[0]
    p_title.text = f"2024 高考真题精讲 - 第 {idx} 题"
    set_font_style(p_title.runs[0], '微软雅黑')
    p_title.runs[0].font.size, p_title.runs[0].font.bold = Pt(24), True

    # 布局：左文右图
    has_img = len(imgs) > 0
    box_w = Inches(8.2) if has_img else Inches(12.3)
    
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.15), box_w, Inches(5.8))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(220, 230, 240)

    content_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), box_w - Inches(0.4), Inches(5.4))
    tf = content_box.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    add_smart_text(tf, q_text)

    if has_img:
        # 自动排版多张图片
        for i, img_path in enumerate(list(set(imgs))[:2]): 
            try: slide.shapes.add_picture(img_path, Inches(8.8), Inches(1.15 + i*3.0), width=Inches(4.2))
            except: pass

# ==========================================
# 5. Streamlit 主界面逻辑
# ==========================================
st.set_page_config(page_title="AI 物理教研全自动工作站", layout="wide")

st.markdown("""
<div style='text-align: center;'>
    <h1 style='color: #0070C0;'>🚀 AI 物理教研全自动工作站</h1>
    <p style='color: #666;'>深度解决 Word 公式 <b>p₀, 10⁵, V₁</b> 识别缺失与图片抓取问题</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📥 上传 物理图片/PDF/Word", accept_multiple_files=True, type=['jpg', 'png', 'jpeg', 'pdf', 'docx'])

if st.button("✨ 一键开启 AI 深度教研排版", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("请上传文件")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        status_info = st.empty()
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            final_extracted = []
            for f_idx, uploaded_file in enumerate(uploaded_files):
                ext = uploaded_file.name.split('.')[-1].lower()
                status_info.info(f"正在处理: {uploaded_file.name} ...")
                
                if ext == 'docx':
                    doc_obj = docx.Document(io.BytesIO(uploaded_file.read()))
                    final_extracted.extend(get_docx_advanced_content(doc_obj, tmp_dir))
                
                elif ext in ['jpg', 'png', 'jpeg', 'pdf']:
                    # 处理图片和 PDF 的逻辑 (Pix2Text)
                    if ext == 'pdf':
                        pdf_doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                        for i in range(len(pdf_doc)):
                            pix = pdf_doc[i].get_pixmap(matrix=fitz.Matrix(2, 2))
                            img_p = os.path.join(tmp_dir, f"pdf_{f_idx}_{i}.jpg")
                            pix.save(img_p)
                            outs = p2t.recognize_mixed(img_p)
                            final_extracted.append({"text": "".join([o['text'] for o in outs]), "imgs": []})
                    else:
                        img_p = os.path.join(tmp_dir, uploaded_file.name)
                        with open(img_p, "wb") as f: f.write(uploaded_file.read())
                        outs = p2t.recognize_mixed(img_p)
                        final_extracted.append({"text": "".join([o['text'] for o in outs]), "imgs": []})

            for idx, q in enumerate(final_extracted):
                create_physics_slide(prs, q["text"], q["imgs"], idx + 1)

        ppt_io = io.BytesIO()
        prs.save(ppt_io)
        st.session_state['ready_ppt'] = ppt_io.getvalue()
        status_info.success(f"🎉 处理完成！共生成 {len(final_extracted)} 页课件。")

if 'ready_ppt' in st.session_state:
    st.write("---")
    st.download_button("⬇️ 下载 AI 巅峰排版 PPT", st.session_state['ready_ppt'], "物理巅峰课件.pptx", use_container_width=True)
