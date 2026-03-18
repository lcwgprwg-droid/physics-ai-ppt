# ==========================================
# Streamlit 现代化 Web UI (修复下载按钮消失 Bug)
# ==========================================
st.set_page_config(page_title="AI 物理教研课件生成器", layout="centered", page_icon="⚛️")

st.markdown("""
<div style='text-align: center; margin-bottom: 30px;'>
    <h1 style='color: #0070C0;'>🚀 AI 物理教研全自动工作站</h1>
    <p style='color: #666;'>支持同时上传多张 <b>教辅照片 / PDF / Word文档</b>，一键生成带有视觉配图与自动分页的巅峰排版 PPT。</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "📥 拖拽上传题库资料（可多选）",
    accept_multiple_files=True,
    type=['jpg', 'jpeg', 'png', 'pdf', 'docx']
)

if st.button("✨ 一键生成精美 PPT", type="primary", use_container_width=True):
    if not uploaded_files:
        st.warning("⚠️ 老师，请先上传文件哦！")
    else:
        progress_bar = st.progress(0)
        status_text = st.empty()

        with tempfile.TemporaryDirectory() as temp_dir:
            status_text.info("⚙️ 正在启动 OCR 与机器视觉引擎，疯狂扫题中...（大概需要十几秒，请稍候）")
            final_questions = process_uploaded_files(uploaded_files, temp_dir)
            progress_bar.progress(60)

            if not final_questions:
                st.error("❌ 抱歉，未能识别到有效题目。")
            else:
                status_text.info(f"✅ 成功提取 {len(final_questions)} 道大题！正在渲染排版...")
                ppt_io = make_master_ppt(final_questions)
                
                # 【黑科技】：把生成好的 PPT 强行锁进浏览器的记忆（Session State）里！
                st.session_state['ready_ppt'] = ppt_io.getvalue()
                
                progress_bar.progress(100)
                status_text.success("🎉 大功告成！课件已生成，请点击下方按钮下载！")
                st.balloons()

# 【核心修复】：把下载按钮放在外面，不管怎么点，它都绝对不会消失了！
if 'ready_ppt' in st.session_state:
    st.markdown("<br>", unsafe_allow_html=True) # 加个空行更好看
    st.download_button(
        label="⬇️ 点击这里下载生成的 PPT 课件",
        data=st.session_state['ready_ppt'],
        file_name="核心素养习题精讲(AI生成版).pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
