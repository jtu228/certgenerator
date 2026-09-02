import io
import json
import os
import re
import zipfile

import pandas as pd
import streamlit as st
from docx import Document
from docxcompose.composer import Composer
from docxtpl import DocxTemplate
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter


DEFAULT_TEMPLATE = "内审员证书.docx"
STANDARDS_FILE = "standards.json"

FALLBACK_STANDARD_PRESETS = {
    "质量管理体系（ISO 9001:2015）": "ISO 9001:2015",
    "质量管理体系新版（ISO 9001:2026）": "ISO 9001:2026",
    "环境管理体系（ISO 14001:2015）": "ISO 14001:2015",
    "环境管理体系（ISO 14001:2026）": "ISO 14001:2026",
    "职业健康安全管理体系（ISO 45001:2018）": "ISO 45001:2018",
    "食品安全管理体系（ISO 22000:2018）": "ISO 22000:2018",
    "食品安全体系认证（FSSC 22000 V6.0）": "FSSC 22000 V6.0",
    "食品安全体系认证（FSSC 22000 V7）": "FSSC 22000 V7",
    "危害分析与关键控制点（HACCP V1.0）": "HACCP V1.0",
    "信息技术服务管理体系（ISO/IEC 20000-1:2018）": "ISO/IEC 20000-1:2018",
    "信息安全管理体系（ISO/IEC 27001:2022）": "ISO/IEC 27001:2022",
    "合规管理体系（GB/T 35770—2022/ISO 37301:2021）": "GB/T 35770—2022/ISO 37301:2021",
    "能源管理体系（ISO 50001:2018）": "ISO 50001:2018",
    "汽车行业质量管理体系（IATF 16949:2016）": "IATF 16949:2016",
    "医疗器械质量管理体系（ISO 13485:2016）": "ISO 13485:2016",
    "设施管理体系（ISO 41001:2018）": "ISO 41001:2018",
    "人工智能管理体系（ISO/IEC 42001:2023）": "ISO/IEC 42001:2023",
    "企业诚信管理体系（GB/T 31950-2023）": "GB/T 31950-2023",
    "碳管理体系（T/CCAA 39—2022）": "T/CCAA 39—2022",
    "商品售后服务评价体系（GB/T 27922-2011）": "GB/T 27922-2011",
    "测量管理体系（GB/T 19022-2003/ISO 10012:2003）": "GB/T 19022-2003/ISO 10012:2003",
    "测量管理体系新版（ISO 10012:2026）": "ISO 10012:2026",
    "物业服务（GB/T 20647.9-2006）": "GB/T 20647.9-2006",
    "BRCGS 食品安全（Food Safety Issue 9）": "BRCGS Food Safety Issue 9",
    "社会责任（SA8000:2026）": "SA8000:2026",
    "创新管理体系（ISO 56001:2024）": "ISO 56001:2024",
    "供应链安全管理体系（ISO 28000:2022）": "ISO 28000:2022",
    "水效率管理体系（ISO 46001:2019）": "ISO 46001:2019",
    "全球良好农业规范（GLOBALG.A.P. IFA v6）": "GLOBALG.A.P. IFA v6",
    "中国良好农业规范（ChinaGAP / GB/T 20014系列）": "ChinaGAP / GB/T 20014系列",
    "反贿赂管理体系（ISO 37001:2025）": "ISO 37001:2025",
    "社会责任管理体系（GB/T 39604-2020）": "GB/T 39604-2020",
    "有机产品（GB/T 19630-2019）": "GB/T 19630-2019",
    "企业知识产权合规管理体系（GB/T 29490-2023）": "GB/T 29490-2023",
    "业务连续性管理体系（GB/T 30146-2023/ISO 22301:2019）": "GB/T 30146-2023/ISO 22301:2019",
}


def load_standard_presets():
    """Load editable standard presets, falling back to built-in values."""
    try:
        with open(STANDARDS_FILE, "r", encoding="utf-8") as standards_file:
            presets = json.load(standards_file)
        if isinstance(presets, dict) and presets:
            return {str(label): str(value) for label, value in presets.items()}
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        pass
    return FALLBACK_STANDARD_PRESETS


def clean_value(value):
    """Convert spreadsheet values to clean strings without leaking NaN."""
    if value is None or pd.isna(value):
        return ""
    cleaned = str(value).strip()
    return "" if cleaned.lower() == "nan" else cleaned


def mask_id_card(id_str):
    """Mask the middle eight digits of an 18-character mainland ID number."""
    id_str = clean_value(id_str)
    if len(id_str) == 18:
        return f"{id_str[:6]}********{id_str[14:]}"
    return id_str


def parse_names(raw_names):
    """Accept names pasted by line, comma, semicolon, or tab and de-duplicate."""
    names = []
    seen = set()
    for item in re.split(r"[\r\n,，;；\t]+", raw_names or ""):
        name = item.strip()
        if name and name not in seen:
            names.append(name)
            seen.add(name)
    return names


def format_training_date(start_date, end_date=None):
    """Format one date or an inclusive date range for the certificate."""
    if not start_date:
        return ""
    if not end_date or end_date == start_date:
        return f"{start_date.year}年{start_date.month}月{start_date.day}日"
    if end_date < start_date:
        start_date, end_date = end_date, start_date
    if start_date.year == end_date.year and start_date.month == end_date.month:
        return (
            f"{start_date.year}年{start_date.month}月"
            f"{start_date.day}—{end_date.day}日"
        )
    if start_date.year == end_date.year:
        return (
            f"{start_date.year}年{start_date.month}月{start_date.day}日—"
            f"{end_date.month}月{end_date.day}日"
        )
    return (
        f"{start_date.year}年{start_date.month}月{start_date.day}日—"
        f"{end_date.year}年{end_date.month}月{end_date.day}日"
    )


def merge_unique(values):
    """Remove blanks and duplicates while preserving selection order."""
    result = []
    seen = set()
    for value in values:
        cleaned = clean_value(value)
        if cleaned and cleaned not in seen:
            result.append(cleaned)
            seen.add(cleaned)
    return result


def safe_filename(value):
    """Make a short Windows-compatible filename component."""
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", clean_value(value))
    cleaned = cleaned.rstrip(". ")
    return cleaned[:80] or "未命名"


def read_template_bytes(template_source):
    if isinstance(template_source, (str, os.PathLike)):
        with open(template_source, "rb") as template_file:
            return template_file.read()
    if hasattr(template_source, "getvalue"):
        return template_source.getvalue()
    template_source.seek(0)
    return template_source.read()


def generate_documents(template_bytes, records, progress_callback=None):
    """Render one certificate per valid row and return merged DOCX + ZIP."""
    master_doc = None
    composer = None
    valid_count = 0
    zip_buffer = io.BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for index, row in enumerate(records):
            name = clean_value(row.get("姓名"))
            if not name:
                continue

            doc = DocxTemplate(io.BytesIO(template_bytes))
            doc.render(
                {
                    "number": clean_value(row.get("证书编号")),
                    "name": name,
                    "id_card": mask_id_card(row.get("身份证号")),
                    "date": clean_value(row.get("培训日期")),
                    "standards": clean_value(row.get("标准号")),
                }
            )

            certificate_buffer = io.BytesIO()
            doc.save(certificate_buffer)
            certificate_bytes = certificate_buffer.getvalue()
            valid_count += 1

            archive.writestr(
                f"{valid_count:03d}_{safe_filename(name)}.docx",
                certificate_bytes,
            )

            current_doc = Document(io.BytesIO(certificate_bytes))
            if master_doc is None:
                master_doc = current_doc
                composer = Composer(master_doc)
            else:
                master_doc.add_page_break()
                composer.append(current_doc)

            if progress_callback:
                progress_callback((index + 1) / len(records))

    if not composer or valid_count == 0:
        return None

    merged_buffer = io.BytesIO()
    composer.save(merged_buffer)
    return {
        "count": valid_count,
        "merged": merged_buffer.getvalue(),
        "zip": zip_buffer.getvalue(),
    }


def make_excel_template():
    example_data = {
        "证书编号": ["T-2026-001（示例）"],
        "姓名": ["张三（示例）"],
        "身份证号": ["440683199001010001"],
        "培训日期": ["2026年9月1—3日"],
        "标准号": ["ISO 9001:2015、ISO 14001:2015"],
    }
    example_df = pd.DataFrame(example_data)
    template_buffer = io.BytesIO()

    with pd.ExcelWriter(template_buffer, engine="openpyxl") as writer:
        example_df.to_excel(writer, index=False, sheet_name="Sheet1")
        worksheet = writer.sheets["Sheet1"]
        for column_index, column in enumerate(example_df.columns, start=1):
            column_letter = get_column_letter(column_index)
            max_length = max(example_df[column].astype(str).map(len).max(), len(column)) + 5
            worksheet.column_dimensions[column_letter].width = max_length

        yellow_fill = PatternFill(
            start_color="FFFF00", end_color="FFFF00", fill_type="solid"
        )
        for cell in worksheet[2]:
            cell.fill = yellow_fill

    return template_buffer.getvalue()


def records_from_editor(standard_presets):
    st.info("点击左上角第一个单元格，然后按 Ctrl+V，可直接粘贴 Excel 数据。")
    initial_df = pd.DataFrame(
        {
            "序号": range(1, 101),
            "证书编号": [""] * 100,
            "姓名": [""] * 100,
            "身份证号": [""] * 100,
            "培训日期": [""] * 100,
            "标准号": [""] * 100,
        }
    )
    edited_df = st.data_editor(
        initial_df,
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        height=380,
        column_config={
            "序号": st.column_config.NumberColumn("序号", width=40, disabled=True),
            "证书编号": st.column_config.TextColumn("证书编号", width="small"),
            "姓名": st.column_config.TextColumn("姓名", width="small"),
            "身份证号": st.column_config.TextColumn("身份证号", width="medium"),
            "培训日期": st.column_config.TextColumn("培训日期", width="medium"),
            "标准号": st.column_config.TextColumn(
                "标准号",
                width="large",
                help="可填写：" + "、".join(standard_presets.values()),
            ),
        },
    )

    records = []
    for row in edited_df.drop(columns=["序号"]).to_dict("records"):
        cleaned = {key: clean_value(value) for key, value in row.items()}
        if any(cleaned.values()):
            records.append(cleaned)
    return records


def records_from_upload():
    left_column, right_column = st.columns([2, 3])
    with left_column:
        st.download_button(
            label="📥 下载标准模板",
            data=make_excel_template(),
            file_name="学员信息上传模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption("系统会自动跳过标有“示例”的行。")

    with right_column:
        uploaded_data = st.file_uploader(
            "上传学员信息文件",
            type=["xlsx", "csv"],
            label_visibility="collapsed",
        )

    if not uploaded_data:
        return []

    try:
        if uploaded_data.name.lower().endswith(".csv"):
            data_frame = pd.read_csv(uploaded_data, dtype=str).fillna("")
        else:
            data_frame = pd.read_excel(uploaded_data, dtype=str).fillna("")
    except Exception as error:
        st.error(f"无法读取文件：{error}")
        return []

    if "姓名" not in data_frame.columns:
        st.error("文件中缺少“姓名”列，请使用标准模板。")
        return []

    records = []
    for row in data_frame.to_dict("records"):
        cleaned = {key: clean_value(value) for key, value in row.items()}
        if "示例" in cleaned.get("姓名", "") or "示例" in cleaned.get("证书编号", ""):
            continue
        if cleaned.get("姓名"):
            records.append(cleaned)

    if records:
        st.success(f"已成功加载 {len(records)} 条有效数据。")
    else:
        st.warning("文件中没有可生成的有效姓名。")
    return records


def quick_records(standard_presets):
    selected_labels = []
    standard_labels = list(standard_presets.keys())
    with st.expander("选择标准（支持多选）", expanded=True):
        for row_start in range(0, len(standard_labels), 2):
            standard_columns = st.columns(2)
            for column_index, label in enumerate(
                standard_labels[row_start : row_start + 2]
            ):
                with standard_columns[column_index]:
                    if st.checkbox(label, key=f"standard_option_{label}"):
                        selected_labels.append(label)

    if selected_labels:
        st.caption(f"已选择 {len(selected_labels)} 个标准")
    custom_standard = st.text_input(
        "自定义标准（可选）",
        placeholder="例如：企业内部标准 Q/ABC 001-2026",
    )

    standard_values = [standard_presets[label] for label in selected_labels]
    if custom_standard:
        standard_values.extend(re.split(r"[\r\n,，;；]+", custom_standard))
    standards = "、".join(merge_unique(standard_values))

    selected_dates = st.date_input(
        "培训日期",
        value=[],
        format="YYYY/MM/DD",
        help="选择一个日期为单日培训；继续选择结束日期即为连续培训。",
    )
    if len(selected_dates) == 1:
        training_date = format_training_date(selected_dates[0])
    elif len(selected_dates) == 2:
        training_date = format_training_date(selected_dates[0], selected_dates[1])
    else:
        training_date = ""

    if training_date:
        st.caption(f"证书显示：{training_date}")
    raw_names = st.text_area(
        "粘贴姓名",
        height=240,
        placeholder="每行一个姓名，例如：\n张三\n李四\n王五",
        help="也支持用逗号、分号或 Tab 分隔；重复姓名会自动去除。",
    )
    names = parse_names(raw_names)

    with st.expander("可选：自动生成证书编号"):
        number_prefix = st.text_input(
            "证书编号前缀",
            placeholder="例如：T-2026-（不填写则证书编号留空）",
        )
        number_start = st.number_input("起始序号", min_value=0, value=1, step=1)
        number_digits = st.number_input("序号位数", min_value=1, max_value=8, value=3, step=1)

    if names:
        st.success(f"已识别 {len(names)} 个不重复姓名。")
        with st.expander("预览姓名"):
            st.write("、".join(names))

    records = []
    for index, name in enumerate(names):
        certificate_number = ""
        if number_prefix.strip():
            sequence = int(number_start) + index
            certificate_number = f"{number_prefix.strip()}{sequence:0{int(number_digits)}d}"
        records.append(
            {
                "证书编号": certificate_number,
                "姓名": name,
                "身份证号": "",
                "培训日期": training_date,
                "标准号": standards,
            }
        )

    return records, bool(standards), bool(training_date)


st.set_page_config(page_title="证书智能制作工具", page_icon="🎓", layout="centered")
st.title("🎓 内审员证书智能制作工具")
st.caption("选择标准、填写日期并粘贴姓名，即可一次生成全部证书。")

standard_presets = load_standard_presets()
data_to_process = []
inputs_valid = True

st.markdown("### 第一步：选择录入方式")
mode = st.radio(
    "录入方式",
    ["快速生成", "完整信息录入"],
    horizontal=True,
    label_visibility="collapsed",
)

st.divider()
st.markdown("### 第二步：填写证书信息")

if mode == "快速生成":
    data_to_process, has_standard, has_date = quick_records(standard_presets)
    inputs_valid = has_standard and has_date and bool(data_to_process)
    if data_to_process and not has_standard:
        st.warning("请至少选择或填写一个标准。")
    if data_to_process and not has_date:
        st.warning("请填写培训日期。")
else:
    full_mode = st.radio(
        "完整信息录入方式",
        ["Excel 文件上传", "网页表格填写（支持粘贴）"],
        horizontal=True,
    )
    if full_mode == "Excel 文件上传":
        data_to_process = records_from_upload()
    else:
        data_to_process = records_from_editor(standard_presets)
    inputs_valid = bool(data_to_process)

st.divider()
st.markdown("### 第三步：确认模板并生成")

if os.path.exists(DEFAULT_TEMPLATE):
    template_option = st.radio(
        "证书 Word 模板",
        ["使用内置模板", "上传本地新模板"],
        horizontal=True,
    )
    if template_option == "使用内置模板":
        template_source = DEFAULT_TEMPLATE
    else:
        template_source = st.file_uploader("上传自定义 Word 模板", type=["docx"])
else:
    st.warning("仓库中未发现默认模板，请上传 Word 模板。")
    template_source = st.file_uploader("上传 Word 模板", type=["docx"])

can_generate = bool(template_source) and inputs_valid
if st.button(
    "🚀 开始批量生成",
    type="primary",
    use_container_width=True,
    disabled=not can_generate,
):
    try:
        progress_bar = st.progress(0, text="正在生成证书……")
        template_bytes = read_template_bytes(template_source)
        result = generate_documents(
            template_bytes,
            data_to_process,
            progress_callback=lambda value: progress_bar.progress(
                value, text="正在生成证书……"
            ),
        )
        progress_bar.empty()
        if not result:
            st.error("没有找到可生成证书的有效姓名。")
        else:
            st.session_state["generation_result"] = result
            st.balloons()
    except Exception as error:
        st.error(f"制作失败：{error}")

if not can_generate:
    st.info("填写完整信息并确认模板后，即可开始生成。")

result = st.session_state.get("generation_result")
if result:
    st.success(f"制作完成，共 {result['count']} 份证书。")
    download_left, download_right = st.columns(2)
    with download_left:
        st.download_button(
            "📄 下载合并 Word",
            data=result["merged"],
            file_name="证书汇总导出.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with download_right:
        st.download_button(
            "🗂️ 下载单独证书 ZIP",
            data=result["zip"],
            file_name="单独证书.zip",
            mime="application/zip",
            use_container_width=True,
        )

