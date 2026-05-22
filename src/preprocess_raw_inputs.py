from pathlib import Path
import re

import pandas as pd

from .config import debug_print


RAW_SCORE_FILE = "副本材料成型技术基础-课程目标达成情况评价表.xlsx"
RAW_LONG_SCORE_FILE = "04_试卷题目得分表_长表.xlsx"


def _normalize_text(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).strip().replace("\u3000", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _normalize_target_id(v) -> str:
    s = _normalize_text(v)
    if not s:
        return ""
    if s.startswith("课程目标"):
        tail = s.replace("课程目标", "", 1)
        return f"课程目标{tail}" if tail else s
    m = re.fullmatch(r"(?:目标|CO|co|Co|cO)(\d+)", s)
    if m:
        return f"课程目标{m.group(1)}"
    if s.isdigit():
        return f"课程目标{s}"
    return s


def _read_course_info(course_path: Path) -> dict:
    course_file = course_path / "01_课程主数据表.xlsx"
    df = pd.read_excel(course_file)
    first = df.iloc[0]
    return {
        "所属课程": first.get("课程名称", ""),
        "所属学期": first.get("开课学期", ""),
    }


def _is_summary_or_empty(student_id: str, name: str) -> bool:
    if not student_id or not name:
        return True
    summary_words = ["平均", "平均值", "合计", "汇总", "总计"]
    return any(word in student_id or word in name for word in summary_words)


def preprocess_student_scores(raw_inputs_dir: Path, course_path: Path) -> Path:
    raw_file = raw_inputs_dir / RAW_SCORE_FILE
    if not raw_file.exists():
        raise FileNotFoundError(f"未找到原始成绩表: {raw_file}")

    raw = pd.read_excel(raw_file, sheet_name="达成情况评价", header=None)
    course_info = _read_course_info(course_path)

    rows = []
    target_specs = [
        ("课程目标1", 3),
        ("课程目标2", 9),
    ]

    for _, row in raw.iloc[6:].iterrows():
        student_id = _normalize_text(row.iloc[1] if len(row) > 1 else "")
        name = _normalize_text(row.iloc[2] if len(row) > 2 else "")
        if _is_summary_or_empty(student_id, name):
            continue

        for target_id, start_col in target_specs:
            rows.append(
                {
                    "学号": student_id,
                    "姓名": name,
                    "课程目标编号": target_id,
                    "课程作业": row.iloc[start_col],
                    "实验操作": row.iloc[start_col + 1],
                    "实验报告": row.iloc[start_col + 2],
                    "课堂表现": row.iloc[start_col + 3],
                    "期末考试成绩": row.iloc[start_col + 4],
                    "原表达成度": row.iloc[start_col + 5],
                    "所属课程": course_info["所属课程"],
                    "所属学期": course_info["所属学期"],
                }
            )

    result = pd.DataFrame(rows)
    output_path = course_path / "02_学生成绩表.xlsx"
    result.to_excel(output_path, index=False)
    debug_print(f"预处理完成: {output_path} shape={result.shape}")
    return output_path


def preprocess_long_scores(raw_inputs_dir: Path, course_path: Path) -> Path:
    raw_file = raw_inputs_dir / RAW_LONG_SCORE_FILE
    if not raw_file.exists():
        raise FileNotFoundError(f"未找到原始试卷题目得分长表: {raw_file}")

    raw = pd.read_excel(raw_file, sheet_name="长表")
    required = ["学号", "姓名", "课程目标", "大题", "小题号", "学生得分"]
    missing = [col for col in required if col not in raw.columns]
    if missing:
        raise ValueError(f"原始试卷题目得分长表缺少字段: {missing}")

    result = raw[required].copy()
    result = result.dropna(subset=["学号", "姓名", "课程目标", "大题", "小题号", "学生得分"], how="all")
    result["学号"] = result["学号"].apply(_normalize_text)
    result["姓名"] = result["姓名"].apply(_normalize_text)
    result["课程目标编号"] = result["课程目标"].apply(_normalize_target_id)
    result = result[
        (result["学号"] != "")
        & (result["姓名"] != "")
        & (result["课程目标编号"] != "")
    ].copy()

    result = result[
        [
            "学号",
            "姓名",
            "课程目标编号",
            "大题",
            "小题号",
            "学生得分",
        ]
    ]

    output_path = course_path / "04_试卷题目得分长表.xlsx"
    result.to_excel(output_path, index=False)
    debug_print(f"预处理完成: {output_path} shape={result.shape}")
    return output_path


def preprocess_raw_inputs(project_root: Path, course_path: Path) -> dict:
    raw_inputs_dir = project_root / "raw_inputs"
    if not raw_inputs_dir.exists():
        raise FileNotFoundError(f"未找到 raw_inputs 目录: {raw_inputs_dir}")

    # 03_试卷分值对应表.xlsx 是题号-课程目标-满分配置表，应人工维护。
    # 不从学生最高分或学生得分明细自动推断题目满分，避免把样本最高分误当作题目满分。
    score_path = preprocess_student_scores(raw_inputs_dir, course_path)
    long_score_path = preprocess_long_scores(raw_inputs_dir, course_path)
    return {
        "scores": score_path,
        "long_scores": long_score_path,
    }
