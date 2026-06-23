import os
import sys
import io
import tempfile
import json
import streamlit as st

# 导入自定义模块
from ai_wrappers import ZhipuWrapper, AgnesWrapper
from resume_utils import (
    extract_text_from_pdf, extract_text_from_images, classify_image_type,
    fill_template, fill_word_with_images, analyze_image
)
from prompts import SYSTEM_PROMPT, SCORE_PROMPT, RISK_PROMPT

# 强制编码
sys.stdout.reconfigure(encoding='utf-8')
sys.stderr.reconfigure(encoding='utf-8')
os.environ["PYTHONIOENCODING"] = "utf-8"
os.environ["LANG"] = "zh_CN.UTF-8"

# ==================== 页面配置 ====================
st.set_page_config(page_title="智能简历制作辅助工具", page_icon="📄", layout="wide")

st.markdown("""
<style>
    .stButton button { background-color: #4CAF50; color: white; border-radius: 8px; }
    h1 { color: #2c3e50; text-align: center; }
    .score-card {
        background-color: #f0f2f6;
        border-radius: 12px;
        padding: 20px;
        margin: 10px 0;
    }
    .risk-card {
        background-color: #fff3e0;
        border-left: 5px solid #ff9800;
        border-radius: 8px;
        padding: 15px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 初始化 Session State ====================
if "pdf_content" not in st.session_state:
    st.session_state.pdf_content = None
if "excel_content" not in st.session_state:
    st.session_state.excel_content = None
if "ai_result" not in st.session_state:
    st.session_state.ai_result = None
if "pdf_text" not in st.session_state:
    st.session_state.pdf_text = None
if "last_pdf_name" not in st.session_state:
    st.session_state.last_pdf_name = None
if "last_excel_name" not in st.session_state:
    st.session_state.last_excel_name = None
if "score_result" not in st.session_state:
    st.session_state.score_result = None
if "risk_result" not in st.session_state:
    st.session_state.risk_result = None
if "tech_requirements" not in st.session_state:
    st.session_state.tech_requirements = ""
if "level" not in st.session_state:
    st.session_state.level = ""
if "ai_wrapper" not in st.session_state:
    st.session_state.ai_wrapper = None
if "model_name" not in st.session_state:
    st.session_state.model_name = "glm-4-plus"

# 图片相关
if "uploaded_images" not in st.session_state:
    st.session_state.uploaded_images = None
if "image_text" not in st.session_state:
    st.session_state.image_text = None
if "last_images_hash" not in st.session_state:
    st.session_state.last_images_hash = None

# Word 模板
if "word_template" not in st.session_state:
    st.session_state.word_template = None
if "last_word_name" not in st.session_state:
    st.session_state.last_word_name = None

# 视觉缓存
if "vision_cache" not in st.session_state:
    st.session_state.vision_cache = {}

# ==================== AI 函数（使用 wrapper）====================
def extract_resume_info(full_text: str, model: str) -> dict:
    wrapper = st.session_state.get("ai_wrapper")
    if wrapper is None:
        raise RuntimeError("AI wrapper not initialized")
    messages = [{"role": "user", "content": f"简历文本：\n{full_text}"}]
    result = wrapper.chat(messages, system_prompt=SYSTEM_PROMPT, response_format="json_object", model=model, temperature=0.1)
    return result

def call_ai_analysis(model: str, prompt: str, resume_json: str, tech_requirements: str = "", level: str = "") -> dict:
    wrapper = st.session_state.get("ai_wrapper")
    if wrapper is None:
        raise RuntimeError("AI wrapper not initialized")
    prompt = prompt.replace("{tech_requirements}", tech_requirements or "未提供")
    prompt = prompt.replace("{level}", level or "未提供")
    prompt = prompt.replace("{resume_json}", resume_json)
    messages = [{"role": "user", "content": prompt}]
    return wrapper.chat(messages, response_format="json_object", model=model, temperature=0.2)

# ==================== UI ====================
st.title("📄 智能简历制作辅助工具")
st.markdown("上传 PDF 简历、个人资料图片（身份证、毕业证、学位证、学信网截图等）和 Excel/Word 模板，AI 自动提取信息并生成标准格式文件。")

with st.sidebar:
    st.header("⚙️ 配置")
    provider = st.selectbox("选择 AI 提供商", ["智谱 AI", "Agnes AI"], index=0)
    if provider == "智谱 AI":
        api_key = st.text_input("智谱 API Key", type="password", key="zhipu_key")
        model_name = st.selectbox("模型", ["glm-4-plus", "glm-4-flash", "glm-4.7"], index=0)
    else:
        api_key = st.text_input("Agnes API Key", type="password", key="agnes_key")
        model_name = st.selectbox("模型", ["agnes-2.0-flash", "agnes-image-2.1-flash"], index=0)
    
    if api_key:
        try:
            if provider == "智谱 AI":
                st.session_state.ai_wrapper = ZhipuWrapper(api_key)
            else:
                st.session_state.ai_wrapper = AgnesWrapper(api_key)
            st.session_state.model_name = model_name
            st.success(f"✅ {provider} 已连接")
        except Exception as e:
            st.error(f"初始化失败: {e}")
    
    st.markdown("---")
    st.caption("Excel 模板需包含：本科学历、研究生学历、工作经历（...）、项目经历（...）等关键字")
    st.caption("Word 模板需包含：身份证正面照片、身份证反面照片、毕业证照片、学位证照片、学信网学历证书电子备案截图、学信网学位证书电子备案截图 等标题")

with st.expander("📌 使用说明", expanded=True):
    st.markdown("""
    1. 准备 **PDF 格式** 的原始简历文件
    2. 准备 **个人资料图片**：身份证正反面、毕业证照片、学位证照片、学信网学历/学位备案表截图等（可选，但强烈推荐）
    3. 准备 **Excel 模板**（.xlsx）或 **Word 模板**（.docx），根据需求选择
    4. 在侧边栏选择 AI 提供商并输入 **API Key**
    5. 填写 **供应商缩写、类型、岗位、级别**
    6. 点击对应按钮进行处理，下载生成的简历文件
    """)

# ---------- 文件上传 ----------
col1, col2, col3 = st.columns(3)
with col1:
    pdf_file = st.file_uploader("📂 PDF 简历", type=["pdf"])
with col2:
    excel_template = st.file_uploader("📁 Excel 模板", type=["xlsx"])
with col3:
    word_template = st.file_uploader("📄 Word 模板（用于证件照填充）", type=["docx"])

# ---------- 图片上传 ----------
st.markdown("---")
st.subheader("🖼️ 个人资料图片（可选，用于补充学历、身份等信息）")
st.caption("支持批量上传身份证正反面、毕业证照片、学位证照片、学信网学历/学位备案表截图等，AI 将自动识别图片中的文字并合并到简历中。")
image_files = st.file_uploader(
    "选择图片（可多选）", 
    type=["png", "jpg", "jpeg"], 
    accept_multiple_files=True,
    key="image_uploader"
)

def get_images_hash(images):
    if not images:
        return None
    return hash(tuple((img.name, img.size) for img in images))

current_hash = get_images_hash(image_files)
if current_hash != st.session_state.last_images_hash:
    st.session_state.image_text = None
    st.session_state.uploaded_images = image_files
    st.session_state.last_images_hash = current_hash

if image_files:
    st.write(f"已上传 {len(image_files)} 张图片")
    for img in image_files:
        st.caption(f" - {img.name}")
    
    if st.button("🔍 从图片中提取文字（使用AI视觉模型）", key="ocr_btn"):
        if st.session_state.ai_wrapper is None:
            st.error("请先在侧边栏配置 AI 提供商和 API Key")
        else:
            with st.spinner(f"正在识别 {len(image_files)} 张图片中的文字，可能需要30秒时间..."):
                try:
                    extracted = extract_text_from_images(image_files)
                    if extracted.strip():
                        st.session_state.image_text = extracted
                        st.success(f"✅ 成功提取文字，共 {len(extracted)} 字符。")
                        with st.expander("查看提取的文字摘要"):
                            st.text(extracted[:1000] + ("..." if len(extracted) > 1000 else ""))
                    else:
                        st.warning("未从图片中提取到任何文字。")
                        st.session_state.image_text = None
                except Exception as e:
                    st.error(f"❌ 图片识别失败: {e}")

# 检测文件变化，清空 AI 结果
if pdf_file is not None and st.session_state.last_pdf_name != pdf_file.name:
    st.session_state.ai_result = None
    st.session_state.pdf_text = None
    st.session_state.last_pdf_name = pdf_file.name
    st.session_state.score_result = None
    st.session_state.risk_result = None
    st.session_state.level = None
if excel_template is not None and st.session_state.last_excel_name != excel_template.name:
    st.session_state.ai_result = None
    st.session_state.last_excel_name = excel_template.name
    st.session_state.score_result = None
    st.session_state.risk_result = None
    st.session_state.level = None
if word_template is not None and st.session_state.last_word_name != word_template.name:
    st.session_state.word_template = word_template
    st.session_state.last_word_name = word_template.name

# 补充信息输入
with st.container():
    st.subheader("补充信息，用于导出文件命名及人员定级评分")
    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        supplier = st.text_input("供应商缩写 (手输)", placeholder="例如：北京南天/云南南天")
    with col_b:
        emp_type = st.selectbox("类型", ["研发", "测试"])
    with col_c:
        position = st.selectbox("岗位", ["java开发", "前端开发", "Hadoop", "质量管理", "产品分析", "系统管理", "测试实施", "功能测试", "技术测试"])
    with col_d:
        level = st.selectbox("级别", ["初级", "中级", "高级", "专家"])

# ---------- Excel 处理 ----------
if st.button("🚀 开始处理 Excel 简历填充", type="primary"):
    missing = []
    if not pdf_file:
        missing.append("PDF 简历")
    if not excel_template:
        missing.append("Excel 模板")
    if st.session_state.ai_wrapper is None:
        missing.append("AI 配置（请先在侧边栏配置）")
    if missing:
        st.error(f"❌ 缺少以下必填项：{', '.join(missing)}，请补充后重试。")
        st.stop()

    if st.session_state.pdf_text is None:
        with st.spinner("读取 PDF..."):
            st.session_state.pdf_text = extract_text_from_pdf(pdf_file)
            if not st.session_state.pdf_text.strip():
                st.error("❌ PDF 文本为空，请检查文件是否可解析（例如非扫描件）。")
                st.stop()
    
    full_text = st.session_state.pdf_text
    if st.session_state.image_text:
        full_text += "\n\n【以下为证件/图片中提取的补充信息】\n" + st.session_state.image_text
        st.info(f"已将图片中提取的文字（{len(st.session_state.image_text)} 字符）合并到简历中。")
    else:
        if image_files and st.session_state.image_text is None:
            st.warning("您上传了图片但尚未提取文字，请先点击【从图片中提取文字】按钮，或继续仅使用PDF内容。")
    
    with st.spinner(f"调用 {st.session_state.model_name} 分析简历（可能需要30秒）..."):
        try:
            ai_result = extract_resume_info(full_text, st.session_state.model_name)
            st.session_state.ai_result = ai_result
            st.success("✅ AI 提取完成")
        except Exception as e:
            st.error(f"❌ AI 调用失败: {e}")
            st.stop()

    extra_hints = []
    if not supplier:
        extra_hints.append("供应商缩写")
    if not emp_type:
        extra_hints.append("类型")
    if not position:
        extra_hints.append("岗位")
    if not level:
        extra_hints.append("级别")
    if extra_hints:
        st.info(f"ℹ️ 以下补充信息未填写，将优先使用 AI 提取结果：{', '.join(extra_hints)}")
    if level != st.session_state.level:
        st.session_state.level = level

# ---------- 评分与风险分析 ----------
if st.session_state.ai_result is not None:
    st.markdown("---")
    st.markdown("## 📊 简历智能评估（AI 评分 + 风险分析）")
    
    tech_req = st.text_area(
        "💼 客户岗位技术要求（选输，用于技术匹配度评分）",
        value=st.session_state.tech_requirements,
        placeholder="例如：Java, Spring Boot, MySQL, Redis, 微服务",
        help="输入关键词后点击下方按钮重新评分"
    )
    if tech_req != st.session_state.tech_requirements:
        st.session_state.tech_requirements = tech_req
    
    if st.button("🔍 开始AI评分与风险分析", key="eval_btn"):
        if st.session_state.ai_wrapper is None:
            st.error("请先在侧边栏配置 AI 提供商和 API Key")
        else:
            with st.spinner("AI 正在评分及分析风险，请稍候..."):
                try:
                    resume_json = json.dumps(st.session_state.ai_result, ensure_ascii=False, indent=2)
                    score_result = call_ai_analysis(
                        st.session_state.model_name, SCORE_PROMPT,
                        resume_json, st.session_state.tech_requirements, st.session_state.level
                    )
                    st.session_state.score_result = score_result
                    risk_result = call_ai_analysis(
                        st.session_state.model_name, RISK_PROMPT,
                        resume_json, "", ""
                    )
                    st.session_state.risk_result = risk_result
                    st.success("✅ 评估完成")
                except Exception as e:
                    st.error(f"❌ AI 评估失败: {e}")
    
    if st.session_state.score_result:
        score = st.session_state.score_result
        st.subheader("📈 评分详情")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("学历专业", f"{score['scores']['education']}/10")
            st.metric("跳槽频率", f"{score['scores']['job_hop']}/10")
            st.metric("项目清晰度", f"{score['scores']['project_clarity']}/20")
        with col2:
            st.metric("IT工作年限", f"{score['scores']['work_years']}/10")
            st.metric("技术栈匹配", f"{score['scores']['tech_match']}/25")
            st.metric("项目角色成果", f"{score['scores']['project_role']}/15")
        with col3:
            st.metric("稳定性/抗压", f"{score['scores']['stability']}/10")
            st.metric("总分", f"{score['total']}/100", delta=None)
        st.info(f"**推荐建议**：{score['suggestion']}  \n**评分理由**：{score['reason']}")
    
    if st.session_state.risk_result:
        st.subheader("⚠️ 潜在风险提示")
        risks = st.session_state.risk_result.get("risks", [])
        if risks:
            for risk in risks:
                level_color = {"高": "🔴", "中": "🟠", "低": "🟡"}.get(risk.get("level", "低"), "⚪")
                st.markdown(f"""<div class="risk-card">
                <strong>{level_color} {risk['category']}</strong>（{risk.get('level', '中')}风险）<br>
                📝 {risk['description']}
                </div>""", unsafe_allow_html=True)
            st.caption(f"📌 总体评价：{st.session_state.risk_result.get('summary', '')}")
        else:
            st.success("✅ 未发现明显风险点")
    
    with st.expander("✏️ 可手动补充风险标注（仅供参考）"):
        manual_risk = st.text_area("可输入额外风险备注，如：协商离职、征信问题等", placeholder="例如：候选人上一家公司协商一致离职，需核实背景")
        if manual_risk:
            st.info(f"📝 已记录额外风险：{manual_risk}")

# 显示 AI 提取结果
if st.session_state.ai_result is not None:
    with st.expander("查看 AI 提取结果（结构化数据）", expanded=False):
        st.json(st.session_state.ai_result)

# ---------- 下载 Excel ----------
if st.session_state.ai_result is not None and excel_template is not None:
    extra = st.session_state.ai_result.get("extra", {})
    if supplier:
        extra["供应商缩写"] = supplier
    if emp_type:
        extra["类型"] = emp_type
    if position:
        extra["岗位"] = position
    if level:
        extra["级别"] = level
    st.session_state.ai_result["extra"] = extra

    sup = extra.get("供应商缩写", "未知")[:4]
    name = st.session_state.ai_result.get("basic", {}).get("姓名", "未知")
    typ = extra.get("类型", "研发")
    pos = extra.get("岗位", "java开发")
    lvl = extra.get("级别", "")
    filename = f"{sup}-{name}-{typ}-{pos}-{lvl}-简历.xlsx"
    filename = "".join(c for c in filename if c not in r'\/:*?"<>|')

    if st.button("📥 下载生成的 Excel 简历", type="secondary"):
        with st.spinner("填充 Excel（完全保留原格式）..."):
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(excel_template.getbuffer())
                tmp_path = tmp.name
            out_path = os.path.join(tempfile.gettempdir(), filename)
            try:
                fill_template(tmp_path, out_path, st.session_state.ai_result)
                st.success("✅ 填充成功！")
            except Exception as e:
                st.error(f"❌ Excel 填充失败: {e}")
                st.stop()
            finally:
                os.unlink(tmp_path)

        with open(out_path, "rb") as f:
            st.download_button("📥 点击下载文件", f, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        os.unlink(out_path)

# ---------- Word 填充 ----------
if word_template is not None and image_files:
    st.markdown("---")
    st.subheader("📄 Word 证件照自动填充（AI 自动识别图片类型，带缓存）")
    st.caption("系统将自动识别您上传的个人资料图片的证件类型，并填充到 Word 模板对应标题下方。")
    
    if st.button("✨ 开始填充 Word 模板", key="fill_word_btn"):
        if st.session_state.ai_wrapper is None:
            st.error("请先在侧边栏配置 AI 提供商和 API Key")
        else:
            if st.session_state.ai_result is not None:
                extra = st.session_state.ai_result.get("extra", {})
                basic = st.session_state.ai_result.get("basic", {})
                sup = extra.get("供应商缩写", supplier)[:4] if not supplier else supplier[:4]
                name = basic.get("姓名", "未知")
                pos = extra.get("岗位", position) if not position else position
                lvl = extra.get("级别", level) if not level else level
            else:
                sup = supplier[:4] if supplier else "未知"
                name = "未知"
                pos = position if position else "java开发"
                lvl = level if level else ""
            
            classified = {}
            progress_bar = st.progress(0)
            total = len(image_files)
            for i, img_file in enumerate(image_files):
                img_bytes = img_file.getvalue()
                img_type = classify_image_type(img_bytes, img_file.name)
                st.write(f"图片 {img_file.name} -> 识别为: {img_type}")
                if img_type:
                    classified[img_type] = (img_bytes, img_file.name)
                else:
                    st.warning(f"无法识别图片 {img_file.name} 的证件类型，已跳过")
                progress_bar.progress((i+1)/total)
            progress_bar.empty()
            
            if not classified:
                st.error("未能识别任何有效证件图片，请确保图片清晰且包含所需的证件类型。")
            else:
                with st.spinner("正在填充 Word 模板并修改标题..."):
                    try:
                        word_bytes = word_template.getvalue()
                        new_title = f"{sup}-{name}-{pos}-{lvl}-个人资料"
                        output_bytes = fill_word_with_images(word_bytes, classified, new_title=new_title)
                        word_filename = f"{sup}-{name}-{pos}-{lvl}-资料.docx"
                        word_filename = "".join(c for c in word_filename if c not in r'\/:*?"<>|')
                        st.success("✅ Word 填充成功！")
                        st.download_button(
                            label="📥 下载填充后的 Word 文件",
                            data=output_bytes,
                            file_name=word_filename,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    except Exception as e:
                        st.error(f"❌ Word 处理失败: {e}")
else:
    if word_template is None and image_files:
        st.info("💡 若需使用 Word 填充功能，请先上传 Word 模板。")
    elif word_template and not image_files:
        st.info("💡 若需使用 Word 填充功能，请先在上方“个人资料图片”区域上传证件照片。")
