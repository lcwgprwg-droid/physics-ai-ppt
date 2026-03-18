import streamlit as st
import os
import re
import io
import tempfile
import gc
import fitz
import docx
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ==========================================
# 1. 极致内存管理：延迟加载模型
# ==========================================
def get_p2t():
    """
    延迟加载。注意：Pix2Text 1.x 版本的初始化不带 analyzer_type 参数
    使用 CPU 模式以节省空间。
    """
    from pix2text import Pix2Text
    # 如果内存还是爆，请尝试 Pix2Text() 不带参数
    return Pix2Text()

# ==========================================
# 2. 字体与排版逻辑 (保持之前的高级排版)
# ==========================================
def set_font_style(run, font_name='微软雅黑', is_italic=False, is_sub=False):
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

def render_rich_text(text_frame, raw_text):
    p = text_frame.paragraphs[0]
    p.line_spacing = 1.3
    text = raw_text.replace('$', '').replace('\\(', '').replace('\\)', '')
    # 简单的分词算法，识别物理量
    parts = re.split(r'([_][a-zA-Z0-9]+)', text)
    for part in parts:
        if not part: continue
        run = p.add_run()
        if part.startswith('_'):
            run.text = part[1:]
            set_font_style(run, font_name='Times New Roman', is_italic=True, is_sub=True)
        else:
            run.text = part
            if re.search(r'[a-zA-Z0-9]', part):
                set_font_style(run, font_name='Times New Roman', is_italic=True)
            else:
                set_font_style(run)

# ==========================================
# 3. Streamlit 主界面
# ==========================================
st.set_page_config(page_title="AI 物理教研全自动工作站", layout="wide")
st.title("🚀 AI 物理教研巅峰工作站 (内存优化版)")

uploaded = st.file_uploader("📥 上传 物理图片/PDF/Word", accept_multiple_files=True, type=['jpg','png','jpeg','pdf','docx'])

if st.button("✨ 一键生成 PPT", type="primary", use_container_width=True):
    if not uploaded:
        st.error("请先上传文件")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        status = st.empty()
        p2t_model = None
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            extracted_questions = []
            for f in uploaded:
                ext = f.name.split('.')[-1].lower()
                status.info(f"正在处理: {f.name}...")
                
                # --- A. Word 逻辑 (极省内存) ---
                if ext == 'docx':
                    from docx import Document
                    doc = Document(io.BytesIO(f.read()))
                    # 使用我们之前优化过的 Word 物理序解析逻辑
                    # (此处简写，请保持你之前 extract_docx_orderly 的代码)
                
                # --- B. 图片/PDF 逻辑 (按需加载模型) ---
                else:
                    if p2t_model is None:
                        status.warning("正在初始化 AI 引擎，请稍候...")
                        p2t_model = get_p2t()
                    
                    if ext == 'pdf':
                        doc_pdf = fitz.open(stream=f.read(), filetype="pdf")
                        for i in range(len(doc_pdf)):
                            pix = doc_pdf[i].get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                            p = os.path.join(tmp_dir, f"p_{i}.jpg")
                            pix.save(p)
                            res = p2t_model.recognize(p) # recognize_mixed 在某些版本中不可用，改用 recognize
                            txt = "".join([it['text'] for it in res])
                            extracted_questions.append({"text": txt, "imgs": []})
                    else:
                        p = os.path.join(tmp_dir, f.name)
                        with open(p, "wb") as tmp_f: tmp_f.write(f.read())
                        res = p2t_model.recognize(p)
                        txt = "".join([it['text'] for it in res])
                        extracted_questions.append({"text": txt, "imgs": []})
                
                gc.collect() # 每张图处理完，强制清理内存

            # 渲染 PPT 页面
            for idx, q in enumerate(extracted_questions):
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                tb = slide.shapes.add_textbox(Inches(0.6), Inches(1.2), Inches(12), Inches(5.5))
                render_rich_text(tb.text_frame, q['text'])
        
        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        st.session_state['ready_ppt'] = ppt_buffer.getvalue()
        status.success("🎉 PPT 生成成功！")

if 'ready_ppt' in st.session_state:
    st.download_button("⬇️ 下载 PPT", st.session_state['ready_ppt'], "物理教研.pptx", use_container_width=True)
