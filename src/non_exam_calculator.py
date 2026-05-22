import re

import numpy as np
import pandas as pd

from .config import ATTAINMENT_THRESHOLD, debug_print


PASS_THRESHOLD = ATTAINMENT_THRESHOLD

DEFAULT_FULL_MARKS = {
    "课程目标1": {
        "过程检查": 30,
        "阶段答辩": 20,
        "课程报告": 60,
    },
    "课程目标2": {
        "过程检查": 10,
        "阶段答辩": 20,
        "课程报告": 20,
    },
    "课程目标3": {
        "过程检查": 10,
        "阶段答辩": 10,
        "课程报告": 20,
    },
}


def _normalize_key(v):
    if pd.isna(v):
        return ""
    s = str(v).strip().replace("\u3000", "").replace(" ", "")
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _normalize_target_id(v):
    s = _normalize_key(v)
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


def _prepare_course_df(course_df: pd.DataFrame) -> pd.DataFrame:
    df = course_df.copy()
    required = ["课程名称", "课程代码", "开课学期", "课程目标编号", "课程目标描述"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"课程主数据表缺少必要字段: {col}")

    df["课程名称"] = df["课程名称"].astype(str).str.strip()
    df["课程代码"] = df["课程代码"].astype(str).str.strip()
    df["开课学期"] = df["开课学期"].astype(str).str.strip()
    df["课程目标编号"] = df["课程目标编号"].apply(_normalize_target_id)
    df["课程目标描述"] = df["课程目标描述"].astype(str).str.strip()

    weight_cols = [col for col in df.columns if str(col).endswith("权重")]
    if not weight_cols:
        raise ValueError("非试卷型课程主数据表中未找到以“权重”结尾的考核项目权重列。")
    for col in weight_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    keep_cols = required + weight_cols
    return df[keep_cols].copy(), weight_cols


def _prepare_score_df(score_df: pd.DataFrame, score_cols: list) -> pd.DataFrame:
    df = score_df.copy()
    required = ["学号", "姓名", "课程目标编号"] + score_cols
    for col in required:
        if col not in df.columns:
            raise ValueError(f"非试卷型学生成绩表缺少必要字段: {col}")

    df = df.dropna(subset=["学号", "姓名", "课程目标编号"]).copy()
    df["学号"] = df["学号"].apply(_normalize_key)
    df["姓名"] = df["姓名"].apply(_normalize_key)
    df["课程目标编号"] = df["课程目标编号"].apply(_normalize_target_id)
    df = df[(df["学号"] != "") & (df["姓名"] != "") & (df["课程目标编号"] != "")].copy()
    for col in score_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def _score_col_from_weight_col(weight_col: str) -> str:
    return str(weight_col).replace("权重", "成绩")


def _full_mark(target_id: str, item_name: str) -> float:
    return float(DEFAULT_FULL_MARKS.get(target_id, {}).get(item_name, 100))


def _normalized_score(series: pd.Series, full_mark: float) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce").fillna(0)
    max_score = numeric.max() if not numeric.empty else 0
    if max_score > full_mark + 1e-9:
        return numeric / 100
    return np.where(full_mark > 0, numeric / full_mark, 0)


def build_non_exam_student_target_result(course_df: pd.DataFrame, score_df: pd.DataFrame) -> pd.DataFrame:
    course, weight_cols = _prepare_course_df(course_df)
    score_cols = [_score_col_from_weight_col(col) for col in weight_cols]
    scores = _prepare_score_df(score_df, score_cols)

    base = scores.merge(
        course.drop_duplicates(subset=["课程目标编号"]),
        on="课程目标编号",
        how="left",
    )

    contribution_cols = []
    normalized_cols = []
    for weight_col in weight_cols:
        item_name = weight_col.replace("权重", "")
        score_col = _score_col_from_weight_col(weight_col)
        normalized_col = f"{item_name}达成值"
        contribution_col = f"{item_name}贡献"
        normalized_cols.append(normalized_col)
        contribution_cols.append(contribution_col)

        base[normalized_col] = 0.0
        for target_id, idx in base.groupby("课程目标编号").groups.items():
            full_mark = _full_mark(target_id, item_name)
            base.loc[idx, normalized_col] = _normalized_score(base.loc[idx, score_col], full_mark)
        base[contribution_col] = base[normalized_col] * pd.to_numeric(base[weight_col], errors="coerce").fillna(0)

    base["目标权重合计"] = base[weight_cols].sum(axis=1)
    base["过程性贡献值"] = np.where(
        base["目标权重合计"] > 0,
        base[contribution_cols].sum(axis=1) / base["目标权重合计"],
        np.nan,
    )

    base["试卷得分"] = 0.0
    base["试卷满分"] = 0.0
    base["试卷达成值"] = np.nan
    base["期末贡献值"] = 0.0
    base["综合达成值"] = base["过程性贡献值"]
    base["达标阈值"] = PASS_THRESHOLD
    base["是否达标"] = np.where(base["综合达成值"] >= PASS_THRESHOLD, "达标", "未达标")
    base["是否缺少过程性数据"] = np.where(
        base[score_cols].isna().all(axis=1),
        "是",
        "否",
    )
    base["是否缺少试卷数据"] = "不适用"

    for col in ["课堂表现", "平时作业", "实验", "实验报告", "课堂表现贡献", "课程作业贡献", "实验操作贡献", "实验报告贡献"]:
        if col not in base.columns:
            base[col] = 0.0

    output_cols = [
        "课程名称",
        "课程代码",
        "开课学期",
        "学号",
        "姓名",
        "课程目标编号",
        "课程目标描述",
        "课堂表现",
        "平时作业",
        "实验",
        "实验报告",
        "课堂表现贡献",
        "课程作业贡献",
        "实验操作贡献",
        "实验报告贡献",
        "过程性贡献值",
        "试卷得分",
        "试卷满分",
        "试卷达成值",
        "期末贡献值",
        "综合达成值",
        "达标阈值",
        "是否达标",
        "是否缺少过程性数据",
        "是否缺少试卷数据",
    ]
    extra_cols = score_cols + weight_cols + normalized_cols + contribution_cols + ["目标权重合计"]
    result = base.sort_values(["课程目标编号", "学号"]).reset_index(drop=True)
    result = result[output_cols + [col for col in extra_cols if col in result.columns]].copy()
    debug_print("\n===== 非试卷型 student_target_df head =====")
    debug_print(result.head())
    return result
