import pandas as pd


def _target_label(v) -> str:
    s = str(v).strip()
    return s if s.startswith("课程目标") else f"课程目标{s}"


def _judge_result(avg: float) -> str:
    if pd.isna(avg):
        return "数据不足"
    if avg >= 0.8:
        return "达成情况良好"
    if avg >= 0.7:
        return "基本达标"
    return "达成情况偏弱"


def _judge_problem(low_ratio: float, high_ratio: float, avg: float) -> str:
    if pd.isna(avg):
        return "缺少完整数据"
    if avg < 0.7:
        return "目标达成不足，需重点改进"
    if low_ratio > 0.3:
        return "虽达到标准，但学生掌握不均衡"
    if high_ratio > 0.5:
        return "整体掌握较好"
    return "整体达标，仍需持续优化"


def _build_analysis_summary(row) -> str:
    target_label = _target_label(row["课程目标编号"])
    return (
        f"{target_label}综合平均达成值为{row['综合平均达成值']:.3f}，"
        f"{'达到' if row['是否达标'] == '达标' else '未达到'}预设达标阈值{row['达标阈值']:.2f}。"
        f"低分段占比为{row['低分段占比']:.1%}，高分段占比为{row['高分段占比']:.1%}。"
    )


def _build_suggestion(row) -> str:
    if pd.isna(row["综合平均达成值"]):
        return "建议补充完整数据后再分析。"
    if row["综合平均达成值"] < 0.7:
        return "后续教学中应强化重点难点内容讲解，并加强针对性训练与过程指导。"
    if row["低分段占比"] > 0.3:
        return "后续教学中应加强分层指导和过程反馈，缩小学生之间的达成差异。"
    return "后续教学中应继续保持教学组织与训练方式，并持续优化细节。"


def build_analysis_base(course_df: pd.DataFrame, course_target_df: pd.DataFrame) -> pd.DataFrame:
    _ = course_df[["课程名称", "课程代码", "开课学期"]].drop_duplicates().iloc[0]

    df = course_target_df.copy()
    df["结果判断"] = df["综合平均达成值"].apply(_judge_result)
    df["主要问题判断"] = df.apply(
        lambda r: _judge_problem(r["低分段占比"], r["高分段占比"], r["综合平均达成值"]),
        axis=1,
    )
    df["分析摘要"] = df.apply(_build_analysis_summary, axis=1)
    df["改进建议草稿"] = df.apply(_build_suggestion, axis=1)

    output_cols = [
        "课程名称",
        "课程代码",
        "开课学期",
        "课程目标编号",
        "课程目标描述",
        "参与人数",
        "综合平均达成值",
        "达标阈值",
        "是否达标",
        "0.0-0.6人数",
        "0.6-0.7人数",
        "0.7-0.8人数",
        "0.8-0.9人数",
        "0.9-1.0人数",
        "低分段占比",
        "高分段占比",
        "结果判断",
        "主要问题判断",
        "分析摘要",
        "改进建议草稿",
    ]
    return df[output_cols].copy()
