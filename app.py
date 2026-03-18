import streamlit as st
import os
import re
import io
import tempfile
import time
import gc  # 垃圾回收
import fitz
import docx
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.oxml.ns import qn

# ==========================================
# 1. 内存保护方案：延迟加载 Pix2Text
# ==========================================
def get_p2t():
    """
    仅在需要时加载模型，并使用 CPU 优化配置
    """
    from pix2text import Pix2Text
    # 移除复杂的 analyzer_type='mfd'，改用默认或轻量化配置防止 1GB 内存炸裂
    # 如果必须识别复杂分式，请确保上传图片不要太大
    return Pix2Text(languages=('en', 'ch_sim'))

# ==========================================
# 2. 巅峰排版辅助函数
# ==========================================
def set_physics_font(run, font_name='微软雅黑', is_italic=False, is_sub=False):
    run.font.name = font_name
    run.font.italic = is_italic
    run.font.subscript = is_sub
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)

def render_rich_text(text_frame, raw_content):
    p = text_frame.paragraphs[0]
    p.line_spacing = 1.3
    text = raw_content.replace('$', '').replace('\\(', '').replace('\\)', '').replace('{', '').replace('}', '')
    parts = re.split(r'([_][a-zA-Z0-9]+|[\^][a-zA-Z0-9]+)', text)
    for part in parts:
        if not part: continue
        is_sub = part.startswith('_')
        is_sup = part.startswith('^')
        clean_text = part[1:] if (is_sub or is_sup) else part
        run = p.add_run()
        run.text = clean_text
        run.font.size = Pt(20)
        if is_sub or is_sup or re.search(r'[a-zA-Z0-9\.\+\-\=\*×]', clean_text):
            set_physics_font(run, 'Times New Roman', is_italic=True, is_sub=is_sub)
            run.font.color.rgb = RGBColor(0, 80, 160)
        else:
            set_physics_font(run, '微软雅黑')
            run.font.color.rgb = RGBColor(40, 40, 40)

# ==========================================
# 3. Word 解析逻辑 (内存友好)
# ==========================================
def extract_docx_orderly(doc_stream, tmp_dir):
    doc = docx.Document(doc_stream)
    questions = []
    current_q = {"text": "", "imgs": []}
    for para in doc.paragraphs:
        para_stream = ""
        for child in para._element.getchildren():
            tag = child.tag
            if tag.endswith('}r'): 
                for t_node in child.iter():
                    if t_node.tag.endswith('}t'):
                        rpr = t_node.getparent().getprevious()
                        if rpr is not None and rpr.find('.//{*}vertAlign') is not None:
                            para_stream += "_"
                        para_stream += t_node.text if t_node.text else ""
            elif tag.endswith('}oMath'):
                for m_node in child.iter():
                    if m_node.tag.endswith('}t'):
                        is_sub = any('sSub' in p.tag for p in m_node.iterancestors())
                        para_stream += f"_{m_node.text}" if is_sub else (m_node.text if m_node.text else "")
            elif tag.endswith('}drawing'):
                rids = re.findall(r'r:embed="([^"]+)"', child.xml)
                for rid in rids:
                    try:
                        rel = doc.part.related_parts[rid]
                        img_path = os.path.join(tmp_dir, f"w_{int(time.time()*1000)}_{rid}.png")
                        with open(img_path, "wb") as f: f.write(rel.blob)
                        current_q["imgs"].append(img_path)
                    except: pass
        txt = para_stream.strip()
        if not txt: continue
        if re.match(r'^\s*(\d+[\.．、]|\(\d+\))', txt):
            if current_q["text"]: questions.append(current_q)
            current_q = {"text": txt + "\n", "imgs": []}
        else:
            current_q["text"] += txt + "\n"
    if current_q["text"]: questions.append(current_q)
    return questions

# ==========================================
# 4. 图像优化：预缩放处理
# ==========================================
def optimize_image(image_path):
    """
    降低图片分辨率以减少 AI 推理时的内存占用
    """
    with Image.open(image_path) as img:
        if max(img.size) > 1500: # 如果图片太大，等比缩小
            img.thumbnail((1500, 1500))
            img.save(image_path, "JPEG", quality=85)

# ==========================================
# 5. Streamlit 主流程
# ==========================================
st.set_page_config(page_title="AI 物理教研全自动工作站", layout="wide")

st.markdown("<h2 style='text-align:center; color:#0070C0;'>🚀 AI 物理教研全自动工作站</h2>", unsafe_allow_html=True)
st.warning("⚠️ 提示：Streamlit Cloud 内存有限。若处理多图请分批上传，处理完后请点击右下角重启应用释放内存。")

uploaded = st.file_uploader("📥 上传 物理图片/PDF/Word", accept_multiple_files=True, type=['jpg','png','jpeg','pdf','docx'])

if st.button("🚀 开始生成 PPT", type="primary", use_container_width=True):
    if not uploaded:
        st.error("请先上传文件")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        status = st.empty()
        p2t = None # 延迟加载占位
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            extracted_data = []
            for f in uploaded:
                ext = f.name.split('.')[-1].lower()
                status.info(f"正在分析: {f.name}...")
                
                if ext == 'docx':
                    extracted_data.extend(extract_docx_orderly(io.BytesIO(f.read()), tmp_dir))
                else:
                    # 只有遇到图片/PDF才加载昂贵的 AI 模型
                    if p2t is None:
                        status.warning("正在加载 AI 模型（消耗约 800MB 内存）...")
                        p2t = get_p2t()
                    
                    if ext == 'pdf':
                        doc = fitz.open(stream=f.read(), filetype="pdf")
                        for i in range(len(doc)):
                            pix = doc[i].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)) # 降低采样率节省内存
                            p = os.path.join(tmp_dir, f"p_{i}.jpg")
                            pix.save(p)
                            optimize_image(p)
                            res = p2t.recognize_mixed(p)
                            extracted_data.append({"text": "".join([it['text'] for it in res]), "imgs": []})
                            gc.collect() # 强制回收内存
                    else:
                        p = os.path.join(tmp_dir, f.name)
                        with open(p, "wb") as tmp_file: tmp_file.write(f.read())
                        optimize_image(p)
                        res = p2t.recognize_mixed(p)
                        extracted_data.append({"text": "".join([it['text'] for it in res]), "imgs": []})
                        gc.collect()

            # 生成幻灯片
            for idx, q in enumerate(extracted_data):
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                # 简单渲染卡片和文本
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.1), Inches(12.5), Inches(5.9))
                card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255)
                tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), Inches(11.9), Inches(5.5))
                render_rich_text(tb.text_frame, q['text'])
                for i, img in enumerate(list(set(q['imgs']))[:2]):
                    try: slide.shapes.add_picture(img, Inches(8.9), Inches(1.2+i*3), width=Inches(3.8))
                    except: pass
        
        # 导出并清理内存
        buffer = io.BytesIO()
        prs.save(buffer)
        st.session_state['ppt_data'] = buffer.getvalue()
        # 清除模型引用，强制垃圾回收
        p2t = None
        gc.collect()
        status.success("🎉 排版完成！")

if 'ppt_data' in st.session_state:
    st.download_button("⬇️ 下载 PPT", st.session_state['ppt_data'], "课件.pptx", use_container_width=True)
