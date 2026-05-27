from contextlib import redirect_stdout
from io import BytesIO, StringIO
from pathlib import Path
import zipfile

import pandas as pd
import streamlit as st

from src.loaders import has_exam_inputs
from src.score_preprocessor import (
    SPLIT_MODE_RANDOM,
    SPLIT_MODE_REAL,
    SPLIT_MODE_WEIGHT,
    WEIGHT_SPLIT_RISK_NOTICE,
    generate_standard_score_table,
)


PROJECT_ROOT = Path(__file__).resolve().parent
INPUT_ROOT = PROJECT_ROOT / "input_courses"
OUTPUT_ROOT = PROJECT_ROOT / "output"


def _course_dirs():
    if not INPUT_ROOT.exists():
        return []
    return sorted([p for p in INPUT_ROOT.iterdir() if p.is_dir()], key=lambda p: p.name)


def _default_course_index(courses):
    for idx, course in enumerate(courses):
        if course.name == "材料成型技术_2025-2026-1":
            return idx
    return 0


def _find_first(course_path: Path, patterns, exclude=None):
    exclude = exclude or []
    for pattern in patterns:
        matches = sorted(
            p for p in course_path.glob(pattern)
            if not any(term in p.name for term in exclude)
        )
        if matches:
            return matches[0]
    return None


def _course_type(course_path: Path):
    try:
        return ("试卷型" if has_exam_inputs(course_path) else "非试卷型"), ""
    except Exception as exc:
        return "输入文件不完整", str(exc)


def _input_status(course_path: Path):
    rows = []
    specs = [
        ("01_课程主数据表.xlsx", ["01_课程主数据表.xlsx", "01_*.xlsx"]),
        ("02_学生成绩表.xlsx", ["02_学生成绩表.xlsx", "02_学生成绩表*.xlsx"], ["核查报告"]),
        ("03_试卷分值对应表.xlsx", ["03_试卷分值对应表.xlsx", "03_*.xlsx"]),
        ("04_试卷题目得分长表.xlsx", ["04_试卷题目得分长表.xlsx", "04_试卷题目得分表_长表.xlsx", "04_*.xlsx"]),
        ("08_历年课程目标达成度数据.xlsx", ["08_*.xlsx", "07_*.xlsx"]),
    ]
    for item in specs:
        label, patterns = item[0], item[1]
        exclude = item[2] if len(item) > 2 else []
        found = _find_first(course_path, patterns, exclude=exclude)
        note = ""
        if label.startswith("08_") and found and found.name.startswith("07_"):
            note = "当前项目兼容使用 07_历年课程目标达成值.xlsx"
        rows.append(
            {
                "输入文件": label,
                "状态": "已存在" if found else "缺失",
                "实际文件": found.name if found else "",
                "说明": note,
            }
        )
    return pd.DataFrame(rows)


def _output_status(output_dir: Path):
    figures_dir = output_dir / "figures"
    rows = [
        ("05_课程目标结果表.xlsx", output_dir / "05_课程目标结果表.xlsx"),
        ("06_达成度报告分析底表.xlsx", output_dir / "06_达成度报告分析底表.xlsx"),
        ("07_达成度报告草稿.docx", output_dir / "07_达成度报告草稿.docx"),
    ]
    data = [
        {"输出项": label, "状态": "已生成" if path.exists() else "未生成", "路径": str(path) if path.exists() else ""}
        for label, path in rows
    ]
    figure_count = len(list(figures_dir.glob("*.png"))) if figures_dir.exists() else 0
    data.append(
        {
            "输出项": "figures/",
            "状态": f"已生成 {figure_count} 张图" if figure_count else "未生成",
            "路径": str(figures_dir) if figures_dir.exists() else "",
        }
    )
    return pd.DataFrame(data)


def _read_audit_summary(course_path: Path):
    report = course_path / "02_学生成绩表_生成核查报告.xlsx"
    if not report.exists():
        return None, {}
    try:
        df = pd.read_excel(report, sheet_name="核查汇总")
    except Exception as exc:
        df = pd.DataFrame([{"核查项": "核查报告读取失败", "结果": str(exc)}])
        return df, {
            "数据来源说明": "",
            "成绩拆分模式": "",
            "是否通过核查": "无法读取",
            "风险提示": f"核查报告已存在，但当前环境无法读取：{exc}",
        }
    summary = dict(zip(df["核查项"].astype(str), df["结果"].astype(str)))
    return df, summary


def _download_button(path: Path, label: str):
    if not path.exists():
        return
    st.download_button(
        label=label,
        data=path.read_bytes(),
        file_name=path.name,
        mime="application/octet-stream",
        use_container_width=True,
    )


def _figures_zip(figures_dir: Path):
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(figures_dir.glob("*.png")):
            zf.write(path, arcname=path.name)
    buffer.seek(0)
    return buffer


st.set_page_config(page_title="课程数据准备工作台", layout="wide")
st.title("课程数据准备工作台")
st.caption("第一版仅用于本地数据准备、核查和报告生成，不包含用户登录、数据库或云部署。")

courses = _course_dirs()
if not courses:
    st.error("未找到 input_courses/ 下的课程目录。")
    st.stop()

course_name = st.selectbox(
    "选择课程目录",
    [p.name for p in courses],
    index=_default_course_index(courses),
)
course_path = INPUT_ROOT / course_name
output_dir = OUTPUT_ROOT / course_name

course_type, course_type_error = _course_type(course_path)
col_a, col_b = st.columns([1, 2])
with col_a:
    st.metric("课程类型", course_type)
with col_b:
    if course_type_error:
        st.warning(course_type_error)

st.subheader("输入文件状态")
st.dataframe(_input_status(course_path), use_container_width=True, hide_index=True)

st.subheader("生成标准 02_学生成绩表")
uploaded = st.file_uploader("上传原始成绩表", type=["xlsx", "xls"])
if uploaded is not None and uploaded.name.lower().endswith(".xls"):
    st.info("上传的是旧版 .xls 文件；如果本机环境未安装 xlrd，建议先用 Excel 另存为 .xlsx 后再上传。")
split_mode = st.radio(
    "成绩拆分模式",
    [SPLIT_MODE_REAL, SPLIT_MODE_WEIGHT, SPLIT_MODE_RANDOM],
    horizontal=True,
)
if split_mode == SPLIT_MODE_WEIGHT:
    st.warning(WEIGHT_SPLIT_RISK_NOTICE)
elif split_mode == SPLIT_MODE_RANDOM:
    st.error("当前选择为随机扰动测试，该数据不是原始真实分项成绩，仅适合功能验证，不应用于正式归档。")

if st.button("生成标准 02_学生成绩表.xlsx", type="primary", use_container_width=True):
    if uploaded is None:
        st.error("请先上传原始成绩表。")
    else:
        upload_dir = course_path / "_uploaded_raw"
        upload_dir.mkdir(parents=True, exist_ok=True)
        raw_path = upload_dir / uploaded.name
        raw_path.write_bytes(uploaded.getbuffer())
        try:
            result = generate_standard_score_table(
                PROJECT_ROOT,
                course_path,
                raw_file=raw_path,
                split_mode=split_mode,
            )
            st.success("标准 02 和核查报告已生成。")
            st.write(f"原始成绩表：{result['raw_file']}")
            st.write(f"标准 02：{result['score_table']}")
            st.write(f"核查报告：{result['review_report']}")
        except Exception as exc:
            st.exception(exc)

audit_df, audit_summary = _read_audit_summary(course_path)
st.subheader("核查报告摘要")
if audit_df is None:
    st.info("尚未生成 02_学生成绩表_生成核查报告.xlsx。")
else:
    left, right = st.columns(2)
    with left:
        st.write(f"数据来源说明：{audit_summary.get('数据来源说明', '')}")
        st.write(f"成绩拆分模式：{audit_summary.get('成绩拆分模式', '')}")
        st.write(f"是否通过核查：{audit_summary.get('是否通过核查', '')}")
    with right:
        st.warning(audit_summary.get("风险提示", ""))
    st.dataframe(audit_df, use_container_width=True, hide_index=True)

st.subheader("运行达成度报告生成")
if st.button("运行达成度报告生成", use_container_width=True):
    buffer = StringIO()
    try:
        from src.main import main as run_course_pipeline

        with redirect_stdout(buffer):
            run_course_pipeline(["--course", course_name])
        st.success("达成度报告生成完成。")
        st.code(buffer.getvalue())
    except Exception as exc:
        output = buffer.getvalue()
        if output:
            st.code(output)
        st.exception(exc)

st.subheader("输出文件状态")
st.dataframe(_output_status(output_dir), use_container_width=True, hide_index=True)

st.subheader("下载")
download_cols = st.columns(4)
with download_cols[0]:
    _download_button(course_path / "02_学生成绩表.xlsx", "下载 02")
with download_cols[1]:
    _download_button(course_path / "02_学生成绩表_生成核查报告.xlsx", "下载核查报告")
with download_cols[2]:
    _download_button(output_dir / "05_课程目标结果表.xlsx", "下载 05")
with download_cols[3]:
    _download_button(output_dir / "06_达成度报告分析底表.xlsx", "下载 06")

download_cols_2 = st.columns(2)
with download_cols_2[0]:
    _download_button(output_dir / "07_达成度报告草稿.docx", "下载 07 Word")
with download_cols_2[1]:
    figures_dir = output_dir / "figures"
    if figures_dir.exists() and list(figures_dir.glob("*.png")):
        st.download_button(
            label="下载 figures.zip",
            data=_figures_zip(figures_dir),
            file_name="figures.zip",
            mime="application/zip",
            use_container_width=True,
        )
