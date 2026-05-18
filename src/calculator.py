import pandas as pd
import numpy as np
import re

from .config import debug_print


PASS_THRESHOLD = 0.70


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


def _normalize_question_part(v):
    s = _normalize_key(v)
    chinese_numbers = {
        "一": "1",
        "二": "2",
        "三": "3",
        "四": "4",
        "五": "5",
        "六": "6",
        "七": "7",
        "八": "8",
        "九": "9",
        "十": "10",
    }
    return chinese_numbers.get(s, s)


def _build_question_id(df: pd.DataFrame, big_col: str, small_col: str) -> pd.Series:
    return df[big_col].apply(_normalize_question_part) + "-" + df[small_col].apply(_normalize_key)


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

    default_weights = {
        "课堂表现权重": 0.04,
        "课程作业权重": 0.20,
        "实验报告权重": 0.08,
        "实验操作权重": 0.08,
        "期末权重": 0.60,
    }

    for col, default in default_weights.items():
        if col not in df.columns:
            df[col] = default
        else:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default)

    keep_cols = [
        "课程名称",
        "课程代码",
        "开课学期",
        "课程目标编号",
        "课程目标描述",
        "课堂表现权重",
        "课程作业权重",
        "实验报告权重",
        "实验操作权重",
        "期末权重",
    ]
    df = df[keep_cols].copy()

    debug_print("course课程目标编号样例:", df["课程目标编号"].drop_duplicates().tolist()[:10])
    return df


def _prepare_score_df(score_df: pd.DataFrame) -> pd.DataFrame:
    df = score_df.copy()

    required = ["学号", "姓名", "课程目标编号", "平时作业", "实验", "实验报告", "课堂表现"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"学生成绩表缺少必要字段: {col}")

    df = df.dropna(subset=["学号", "姓名", "课程目标编号"]).copy()

    df["学号"] = df["学号"].apply(_normalize_key)
    df["姓名"] = df["姓名"].apply(_normalize_key)
    df["课程目标编号"] = df["课程目标编号"].apply(_normalize_target_id)

    df = df[
        (df["学号"] != "") &
        (df["姓名"] != "") &
        (df["课程目标编号"] != "")
    ].copy()

    numeric_cols = ["平时作业", "实验", "实验报告", "课堂表现"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    keep_cols = [
        "学号", "姓名", "课程目标编号",
        "平时作业", "实验", "实验报告", "课堂表现"
    ]
    df = df[keep_cols].copy()
    debug_print("\n===== 学生成绩表中课程目标编号唯一值（标准化后）=====")
    debug_print(df["课程目标编号"].drop_duplicates().tolist())
    return df


def _prepare_mapping_df(mapping_df: pd.DataFrame) -> pd.DataFrame:
    df = mapping_df.copy()

    required_base = ["课程目标编号", "满分"]
    for col in required_base:
        if col not in df.columns:
            raise ValueError(f"试卷分值对应表缺少必要字段: {col}")

    if "题号" not in df.columns:
        if "大题" in df.columns and "小题号" in df.columns:
            df["题号"] = _build_question_id(df, "大题", "小题号")
        else:
            raise ValueError("试卷分值对应表缺少题号字段，也无法由大题/小题号生成题号。")

    df["课程目标编号"] = df["课程目标编号"].apply(_normalize_target_id)
    df["题号"] = df["题号"].apply(_normalize_key)
    df["满分"] = pd.to_numeric(df["满分"], errors="coerce").fillna(0)

    keep_cols = ["所属课程", "所属学期", "题号", "题型", "满分", "课程目标编号"]
    for col in keep_cols:
        if col not in df.columns:
            raise ValueError(f"试卷分值对应表缺少必要字段: {col}")

    debug_print("mapping课程目标编号样例:", df["课程目标编号"].drop_duplicates().tolist()[:10])
    debug_print("\n===== 试卷分值对应表中课程目标编号唯一值（标准化后）=====")
    debug_print(df["课程目标编号"].drop_duplicates().tolist())
    debug_print("mapping满分汇总:", df.groupby("课程目标编号")["满分"].sum().to_dict())

    return df[keep_cols].copy()


def _prepare_long_score_df(long_score_df: pd.DataFrame) -> pd.DataFrame:
    df = long_score_df.copy()

    if "课程目标" in df.columns and "课程目标编号" not in df.columns:
        df["课程目标编号"] = df["课程目标"]

    required = ["学号", "姓名", "课程目标编号", "大题", "小题号", "学生得分"]
    for col in required:
        if col not in df.columns:
            raise ValueError(f"试卷题目得分长表缺少必要字段: {col}")

    df["学号"] = df["学号"].apply(_normalize_key)
    df["姓名"] = df["姓名"].apply(_normalize_key)
    df["课程目标编号"] = df["课程目标编号"].apply(_normalize_target_id)
    df["题号"] = _build_question_id(df, "大题", "小题号").apply(_normalize_key)
    df["学生得分"] = pd.to_numeric(df["学生得分"], errors="coerce").fillna(0)

    keep_cols = ["学号", "姓名", "课程目标编号", "题号", "学生得分"]
    df = df[keep_cols].copy()

    debug_print("long_scores课程目标编号样例:", df["课程目标编号"].drop_duplicates().tolist()[:10])
    debug_print("\n===== 题目得分长表中课程目标编号唯一值（标准化后）=====")
    debug_print(df["课程目标编号"].drop_duplicates().tolist())
    debug_print("long_scores学生得分汇总:", df.groupby("课程目标编号")["学生得分"].sum().to_dict())

    return df


def build_student_target_result(
    course_df: pd.DataFrame,
    score_df: pd.DataFrame,
    mapping_df: pd.DataFrame,
    long_score_df: pd.DataFrame,
) -> pd.DataFrame:
    course = _prepare_course_df(course_df)
    scores = _prepare_score_df(score_df)
    mapping = _prepare_mapping_df(mapping_df)
    long_scores = _prepare_long_score_df(long_score_df)

    # 1. 每个课程目标对应的试卷满分
    target_full_score = (
        mapping.groupby("课程目标编号", as_index=False)["满分"]
        .sum()
        .rename(columns={"满分": "试卷满分"})
    )

    mapping_question_targets = mapping[["题号", "课程目标编号"]].drop_duplicates()
    duplicated_question_targets = mapping_question_targets.groupby("题号")["课程目标编号"].nunique()
    duplicated_question_targets = duplicated_question_targets[duplicated_question_targets > 1]
    if not duplicated_question_targets.empty:
        raise ValueError(f"试卷分值对应表中同一题号对应多个课程目标: {duplicated_question_targets.to_dict()}")

    exam_detail = long_scores.rename(columns={"课程目标编号": "04课程目标编号"}).merge(
        mapping_question_targets.rename(columns={"课程目标编号": "03课程目标编号"}),
        on="题号",
        how="left",
    )
    missing_mapping_count = int(exam_detail["03课程目标编号"].isna().sum())
    mismatch_count = int(
        (
            exam_detail["03课程目标编号"].notna()
            & (exam_detail["04课程目标编号"] != exam_detail["03课程目标编号"])
        ).sum()
    )
    debug_print("\n===== 期末试卷明细与 03 分值对应表按题号 merge 后 head =====")
    debug_print(exam_detail.head())
    debug_print("期末试卷明细未匹配到 03 题号数量:", missing_mapping_count)
    debug_print("04 长表目标编号与 03 映射目标编号不一致数量:", mismatch_count)

    exam_detail["课程目标编号"] = exam_detail["03课程目标编号"].fillna(exam_detail["04课程目标编号"])

    # 2. 每个学生在每个课程目标上的试卷得分，以 03 的题号-课程目标映射为准
    student_exam_score = (
        exam_detail.groupby(["学号", "姓名", "课程目标编号"], as_index=False)["学生得分"]
        .sum()
        .rename(columns={"学生得分": "试卷得分"})
    )

    debug_print("target_full_score =")
    debug_print(target_full_score)
    debug_print("\n===== 期末成绩按课程目标汇总后的 head =====")
    debug_print(student_exam_score.head())

    # 3. 以过程性成绩表为主，再与课程目标主数据合并
    base = scores.merge(
        course[
            [
                "课程名称",
                "课程代码",
                "开课学期",
                "课程目标编号",
                "课程目标描述",
                "课堂表现权重",
                "课程作业权重",
                "实验报告权重",
                "实验操作权重",
                "期末权重",
            ]
        ].drop_duplicates(subset=["课程目标编号"]),
        on="课程目标编号",
        how="left",
    )

    # 过程性成绩按课程目标分别计算
    for col in ["平时作业", "实验", "实验报告", "课堂表现"]:
        base[col] = pd.to_numeric(base[col], errors="coerce").fillna(0)

    base["课堂表现贡献"] = base["课堂表现"] / 5 * base["课堂表现权重"]
    base["课程作业贡献"] = base["平时作业"] / 25 * base["课程作业权重"]
    base["实验操作贡献"] = base["实验"] / 10 * base["实验操作权重"]
    base["实验报告贡献"] = base["实验报告"] / 10 * base["实验报告权重"]

    base["过程性贡献值"] = (
        base["课堂表现贡献"]
        + base["课程作业贡献"]
        + base["实验操作贡献"]
        + base["实验报告贡献"]
    )

    debug_print("\n===== 过程性成绩计算结果 head =====")
    debug_print(
        base[
            [
                "学号",
                "姓名",
                "课程目标编号",
                "平时作业",
                "实验",
                "实验报告",
                "课堂表现",
                "过程性贡献值",
            ]
        ].head()
    )

    # 4. 合并试卷得分与试卷满分
    base = base.merge(
        student_exam_score,
        on=["学号", "姓名", "课程目标编号"],
        how="left",
    )
    base = base.merge(
        target_full_score,
        on="课程目标编号",
        how="left",
    )

    debug_print("\n===== 过程成绩与期末成绩 merge 后的 head =====")
    debug_print(
        base[
            [
                "学号",
                "姓名",
                "课程目标编号",
                "过程性贡献值",
                "试卷得分",
                "试卷满分",
            ]
        ].head()
    )

    # 5. 缺失标记
    base["是否缺少过程性数据"] = np.where(
        base[["平时作业", "实验", "实验报告", "课堂表现"]].isna().all(axis=1),
        "是",
        "否"
    )
    base["是否缺少试卷数据"] = np.where(base["试卷得分"].isna(), "是", "否")

    # 6. 数值清洗
    base["试卷得分"] = pd.to_numeric(base["试卷得分"], errors="coerce").fillna(0)
    base["试卷满分"] = pd.to_numeric(base["试卷满分"], errors="coerce").fillna(0)

    # 8. 期末贡献
    base["试卷达成值"] = np.where(
        base["试卷满分"] > 0,
        base["试卷得分"] / base["试卷满分"],
        np.nan,
    )
    base["期末贡献值"] = base["试卷达成值"] * base["期末权重"]

    # 9. 综合达成值
    base["综合达成值"] = base["过程性贡献值"] + base["期末贡献值"]
    base["达标阈值"] = PASS_THRESHOLD
    base["是否达标"] = np.where(base["综合达成值"] >= PASS_THRESHOLD, "达标", "未达标")

    # 排序
    base = base.sort_values(by=["课程目标编号", "学号"]).reset_index(drop=True)

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
    result = base[output_cols].copy()
    debug_print("\n===== 写入 05 前 student_target_df head =====")
    debug_print(result.head())
    return result


def _count_intervals(series: pd.Series) -> dict:
    s = pd.to_numeric(series, errors="coerce").dropna()

    return {
        "0.0-0.6人数": int(((s >= 0.0) & (s < 0.6)).sum()),
        "0.6-0.7人数": int(((s >= 0.6) & (s < 0.7)).sum()),
        "0.7-0.8人数": int(((s >= 0.7) & (s < 0.8)).sum()),
        "0.8-0.9人数": int(((s >= 0.8) & (s < 0.9)).sum()),
        "0.9-1.0人数": int(((s >= 0.9) & (s <= 1.0)).sum()),
    }


def build_course_target_result(student_target_df: pd.DataFrame) -> pd.DataFrame:
    df = student_target_df.copy()

    group_cols = ["课程名称", "课程代码", "开课学期", "课程目标编号", "课程目标描述"]

    rows = []
    for keys, grp in df.groupby(group_cols):
        course_name, course_code, term, target_id, target_desc = keys

        participant_count = grp["学号"].nunique()
        missing_process_count = int((grp["是否缺少过程性数据"] == "是").sum())
        missing_exam_count = int((grp["是否缺少试卷数据"] == "是").sum())

        exam_full_score = float(grp["试卷满分"].dropna().max()) if not grp["试卷满分"].dropna().empty else 0.0
        exam_avg_score = float(grp["试卷得分"].mean()) if not grp["试卷得分"].empty else np.nan
        process_avg = float(grp["过程性贡献值"].mean()) if not grp["过程性贡献值"].empty else np.nan
        final_avg = float(grp["期末贡献值"].mean()) if not grp["期末贡献值"].empty else np.nan
        total_avg_attainment = float(grp["综合达成值"].mean()) if not grp["综合达成值"].empty else np.nan

        interval_counts = _count_intervals(grp["综合达成值"])
        pass_count = int((grp["综合达成值"] >= PASS_THRESHOLD).sum())

        rows.append(
            {
                "课程名称": course_name,
                "课程代码": course_code,
                "开课学期": term,
                "课程目标编号": target_id,
                "课程目标描述": target_desc,
                "参与人数": participant_count,
                "缺少过程性数据人数": missing_process_count,
                "缺少试卷数据人数": missing_exam_count,
                "试卷理论总分": exam_full_score,
                "试卷平均得分": exam_avg_score,
                "过程性平均贡献值": process_avg,
                "期末平均贡献值": final_avg,
                "综合平均达成值": total_avg_attainment,
                "达标阈值": PASS_THRESHOLD,
                "达标人数": pass_count,
                "是否达标": "达标" if pd.notna(total_avg_attainment) and total_avg_attainment >= PASS_THRESHOLD else "未达标",
                **interval_counts,
            }
        )

    result = pd.DataFrame(rows)

    if not result.empty:
        result["低分段占比"] = (
            (result["0.0-0.6人数"] + result["0.6-0.7人数"]) / result["参与人数"]
        )
        result["高分段占比"] = (
            (result["0.8-0.9人数"] + result["0.9-1.0人数"]) / result["参与人数"]
        )

    result = result.sort_values(by=["课程目标编号"]).reset_index(drop=True)
    debug_print("\n===== 写入 05 前 course_target_df head =====")
    debug_print(result.head())
    return result
