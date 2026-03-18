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
# 1. 初始化 Pix2Text (启用公式检测)
# ==========================================
@st.cache_resource
def load_pix2text():
    return Pix2Text(analyzer_type='mfd')

p2t = load_pix2text()

# ==========================================
# 2. 巅峰排版引擎：支持真·下标排版
# ==========================================
def set_font_run(run, font_name, is_italic=False, is_sub=False, is_sup=False):
    """底层字体设置：支持中文、西文、斜体及上下标"""
    run.font.name = font_name
    run.font.italic = is_italic
    run.font.subscript = is_sub
    run.font.superscript = is_sup
    
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = rPr.makeelement(qn('w:rFonts'))
        rPr.append(rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)

def add_smart_physics_text(text_frame, raw_text):
    """
    智能解析算法：
    将文本中的 _1 识别为下标，将 ^2 识别为上标，并嵌入题干
    """
    p = text_frame.paragraphs[0]
    p.line_spacing = 1.3
    
    # 清理 LaTeX 多余符号
    text = raw_text.replace('$', '').replace('\\(', '').replace('\\)', '')
    text = text.replace('{', '').replace('}', '') # 处理 T_{1} 为 T_1
    
    # 正则切分：识别 汉字 | 变量 | 下标标记 | 上标标记
    # 匹配模式：_紧跟数字/字母 或 ^紧跟数字/字母
    parts = re.split(r'([_][a-zA-Z0-9]+|[\^][a-zA-Z0-9]+)', text)
    
    for part in parts:
        if not part: continue
        
        is_sub = part.startswith('_')
        is_sup = part.startswith('^')
        clean_part = part[1:] if (is_sub or is_sup) else part
        
        run = p.add_run()
        run.text = clean_part
        run.font.size = Pt(20)
        
        # 物理变量判断 (英文字母、数字、符号)
        if is_sub or is_sup or re.search(r'[a-zA-Z0-9\.\+\-\=\*×]', clean_part):
            set_font_run(run, 'Times New Roman', is_italic=True, is_sub=is_sub, is_sup=is_sup)
            run.font.color.rgb = RGBColor(0, 80, 160) # 物理蓝
        else:
            set_font_run(run, '微软雅黑')
            run.font.color.rgb = RGBColor(30, 30, 30)

# ==========================================
# 3. OCR 坐标感知排序算法 (解决排版混乱关键)
# ==========================================
def sort_and_merge_p2t(results):
    """
    将 OCR 散乱的块按行重新组织，确保公式插入正确位置
    """
    if not results: return ""
    
    # 1. 按中心点 Y 坐标初筛行 (容差 15 像素)
    lines = []
    # 结果按 Y 坐标排序
    sorted_res = sorted(results, key=lambda x: x['position'][0][1])
    
    if not sorted_res: return ""
    
    current_line = [sorted_res[0]]
    for i in range(1, len(sorted_res)):
        # 如果 Y 坐标差距较小，视为同一行
        if abs(sorted_res[i]['position'][0][1] - current_line[-1]['position'][0][1]) < 20:
            current_line.append(sorted_res[i])
        else:
            # 同行内按 X 坐标排序
            current_line.sort(key=lambda x: x['position'][0][0])
            lines.append(current_line)
            current_line = [sorted_res[i]]
    current_line.sort(key=lambda x: x['position'][0][0])
    lines.append(current_line)
    
    # 2. 合并文本
    full_text = ""
    for line in lines:
        line_str = ""
        for item in line:
            # 自动在公式前后加微小空格
            if item['type'] == 'formula':
                line_str += f" {item['text']} "
            else:
                line_str += item['text']
        full_text += line_str + "\n"
    return full_text

# ==========================================
# 4. 文档深度提取逻辑
# ==========================================
def extract_docx_physics(doc_obj, tmp_dir):
    extracted = []
    current_q = {"text": "", "imgs": []}
    
    for para in doc_obj.paragraphs:
        para_text = ""
        # 流式遍历所有节点 (文字 + 公式)
        for node in para._element.iter():
            # 普通文本
            if node.tag.endswith('}t'):
                # 检查上级是否有上下标属性
                rpr = node.getparent().getprevious()
                suffix = ""
                if rpr is not None:
                    va = rpr.find('.//{*}vertAlign')
                    if va is not None:
                        v = va.get('{*}val')
                        if v == 'subscript': para_text += "_"
                        elif v == 'superscript': para_text += "^"
                para_text += node.text if node.text else ""
            # 数学公式内的文本
            elif node.tag.endswith('}mText') or node.tag.endswith('}t'):
                # 判断是否在上下标容器内
                is_sub = any('sSub' in p.tag for p in node.iterancestors())
                is_sup = any('sSup' in p.tag for p in node.iterancestors())
                t = node.text if node.text else ""
                if is_sub: para_text += f"_{t}"
                elif is_sup: para_text += f"^{t}"
                else: para_text += t
        
        # 图片提取 (全局扫描)
        xml = para._element.xml
        rids = set(re.findall(r'r:embed="([^"]+)"', xml))
        for rid in rids:
            try:
                rel = doc_obj.part.related_parts[rid]
                if "image" in rel.content_type:
                    p = os.path.join(tmp_dir, f"img_{rid}.png")
                    with open(p, "wb") as f: f.write(rel.blob)
                    current_q["imgs"].append(p)
            except: pass

        txt = para_text.strip()
        if not txt: continue
        
        # 题号切分
        if re.match(r'^\s*(\d+[\.．、]|\(\d+\))', txt):
            if current_q["text"]: extracted.append(current_q)
            current_q = {"text": txt + "\n", "imgs": []}
        else:
            current_q["text"] += txt + "\n"
            
    if current_q["text"]: extracted.append(current_q)
    return extracted

# ==========================================
# 5. PPT 渲染
# ==========================================
def create_ppt_page(prs, q_text, imgs, idx):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 背景
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid(); bg.fill.fore_color.rgb = RGBColor(250, 252, 255); bg.line.fill.background()
    # 标题栏
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(10), Inches(0.6))
    p = tb.text_frame.paragraphs[0]; p.text = f"核心素养习题精讲 - 题 {idx}"; p.font.bold = True
    set_font_run(p.runs[0], '微软雅黑')
    
    # 布局
    has_img = len(imgs) > 0
    box_w = Inches(8.3) if has_img else Inches(12.3)
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.4), Inches(1.1), box_w, Inches(5.8))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255, 255, 255); card.line.color.rgb = RGBColor(200, 210, 230)
    
    # 内容渲染
    content_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.3), box_w-Inches(0.4), Inches(5.4))
    tf = content_box.text_frame; tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    add_smart_physics_text(tf, q_text)
    
    # 图片摆放
    if has_img:
        for i, img in enumerate(list(set(imgs))[:2]):
            try: slide.shapes.add_picture(img, Inches(8.9), Inches(1.2 + i*3.0), width=Inches(4.0))
            except: pass

# ==========================================
# 6. Streamlit 界面
# ==========================================
st.set_page_config(page_title="AI 物理教研巅峰工作站", layout="wide")
st.markdown("<h1 style='text-align:center; color:#0070C0;'>🚀 AI 物理教研全自动工作站</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align:center; color:#666;'>深度解决公式乱飞、下标丢失与图片漏抓</p>", unsafe_allow_html=True)

uploaded_files = st.file_uploader("📥 上传 物理图片/PDF/Word", accept_multiple_files=True, type=['jpg', 'png', 'jpeg', 'pdf', 'docx'])

if st.button("✨ 一键开启巅峰排版生成", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("请上传文件")
    else:
        prs = Presentation()
        prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
        status = st.empty()
        
        with tempfile.TemporaryDirectory() as tmp_dir:
            all_extracted = []
            for f in uploaded_files:
                ext = f.name.split('.')[-1].lower()
                status.info(f"正在精准解析: {f.name}")
                if ext == 'docx':
                    all_extracted.extend(extract_docx_physics(docx.Document(io.BytesIO(f.read())), tmp_dir))
                elif ext in ['jpg', 'png', 'jpeg', 'pdf']:
                    # 统一转为图片处理
                    if ext == 'pdf':
                        doc = fitz.open(stream=f.read(), filetype="pdf")
                        for i in range(len(doc)):
                            pix = doc[i].get_pixmap(matrix=fitz.Matrix(2, 2))
                            p = os.path.join(tmp_dir, f"pdf_{i}.jpg")
                            pix.save(p)
                            res = p2t.recognize_mixed(p)
                            all_extracted.append({"text": sort_and_merge_p2t(res), "imgs": []})
                    else:
                        p = os.path.join(tmp_dir, f.name)
                        with open(p, "wb") as file: file.write(f.read())
                        res = p2t.recognize_mixed(p)
                        all_extracted.append({"text": sort_and_merge_p2t(res), "imgs": []})
            
            for idx, q in enumerate(all_extracted):
                create_ppt_page(prs, q['text'], q['imgs'], idx + 1)
        
        ppt_buffer = io.BytesIO()
        prs.save(ppt_buffer)
        st.session_state['final_ppt'] = ppt_buffer.getvalue()
        status.success("🎉 排版优化完成！")

if 'final_ppt' in st.session_state:
    st.download_button("⬇️ 下载优化排版 PPT", st.session_state['final_ppt'], "物理巅峰课件.pptx", use_container_width=True)
