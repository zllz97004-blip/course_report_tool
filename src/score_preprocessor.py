from pathlib import Path
import re

import pandas as pd

from .config import debug_print
from .loaders import has_exam_inputs, normalize_columns


SOURCE_REAL = "真实分项成绩"
SOURCE_WEIGHT_SPLIT = "按课程目标权重拆分"

EXAM_ITEM_FULL_MARKS = {
    "课程作业": 25,
    "实验操作": 10,
    "实验报告": 10,
    "课堂表现": 5,
}


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


def _read_course_df(course_path: Path) -> pd.DataFrame:
    df = pd.read_excel(course_path / "01_课程主数据表.xlsx")
    df = df.copy()
    df.columns = [
        str(col).replace("\n", "").replace("\r", "").replace(" ", "").replace("\u3000", "").strip()
        for col in df.columns
    ]
    df["课程目标编号"] = df["课程目标编号"].apply(_normalize_target_id)
    return df


def _course_info(course_df: pd.DataFrame) -> dict:
    first = course_df.iloc[0]
    return {
        "课程名称": first.get("课程名称", ""),
        "开课学期": first.get("开课学期", ""),
    }


def _target_weight_cols(course_df: pd.DataFrame) -> list:
    return [col for col in course_df.columns if str(col).endswith("权重")]


def _score_col_from_weight_col(weight_col: str) -> str:
    return f"{str(weight_col).replace('权重', '')}成绩"


def _find_raw_score_file(search_dirs: list, course_name: str) -> Path:
    candidates = []
    for search_dir in search_dirs:
        if not search_dir.exists():
            continue
        candidates.extend(
            p for p in sorted(search_dir.glob("*.xls*"))
            if not p.name.startswith("01_")
            and not p.name.startswith("02_")
            and not p.name.startswith("03_")
            and not p.name.startswith("04_")
            and "题目得分" not in p.name
            and "核查报告" not in p.name
        )
    preferred = [p for p in candidates if course_name and course_name in p.name]
    for p in preferred + candidates:
        try:
            pd.ExcelFile(p)
            return p
        except Exception:
            continue
    searched = "、".join(str(p) for p in search_dirs)
    raise FileNotFoundError(f"未在以下目录中找到可读取的原始成绩表: {searched}")


def _read_flat_sheet(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet_name)
        df = normalize_columns(df)
        if {"学号", "姓名"}.issubset(set(df.columns)):
            return df
    return pd.DataFrame()


def _read_material_achievement_sheet(path: Path, course_df: pd.DataFrame) -> pd.DataFrame:
    raw = pd.read_excel(path, sheet_name="达成情况评价", header=None)
    info = _course_info(course_df)
    rows = []
    target_specs = [("课程目标1", 3), ("课程目标2", 9)]
    for _, row in raw.iloc[6:].iterrows():
        student_id = _normalize_text(row.iloc[1] if len(row) > 1 else "")
        name = _normalize_text(row.iloc[2] if len(row) > 2 else "")
        if not student_id or not name or "平均" in student_id or "平均" in name:
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
                    "所属课程": info["课程名称"],
                    "所属学期": info["开课学期"],
                    "数据来源说明": SOURCE_REAL,
                }
            )
    return pd.DataFrame(rows)


def _read_final_only_from_wide_sheet(path: Path) -> pd.DataFrame:
    xls = pd.ExcelFile(path)
    for sheet_name in xls.sheet_names:
        raw = pd.read_excel(path, sheet_name=sheet_name, header=None)
        header_rows = raw.index[
            raw.apply(lambda r: r.astype(str).str.contains("学号", na=False).any(), axis=1)
        ].tolist()
        if not header_rows:
            continue
        header_row = header_rows[0]
        header_values = raw.iloc[header_row].astype(str).tolist()
        id_cols = [i for i, v in enumerate(header_values) if "学号" in v]
        if not id_cols:
            continue

        column_blocks = []
        for idx, id_col in enumerate(id_cols):
            block_end = id_cols[idx + 1] if idx + 1 < len(id_cols) else raw.shape[1]
            block_cols = range(id_col, block_end)
            name_col = next((c for c in block_cols if "姓名" in str(raw.iat[header_row, c])), None)
            teacher_col = next((c for c in block_cols if "指导" in str(raw.iat[header_row, c])), None)
            total_candidates = []
            for r in range(header_row, min(header_row + 4, len(raw))):
                for c in block_cols:
                    text = str(raw.iat[r, c])
                    if "总评" in text or "总分" in text or text == "成绩":
                        total_candidates.append(c)
            if name_col is None or not total_candidates:
                continue
            data_part = raw.iloc[header_row + 1:]
            scored_candidates = []
            for col in sorted(set(total_candidates)):
                values = pd.to_numeric(data_part.iloc[:, col], errors="coerce")
                valid_count = int((values > 0).sum())
                scored_candidates.append((valid_count, col))
            scored_candidates.sort(reverse=True)
            total_col = scored_candidates[0][1]
            column_blocks.append(
                {
                    "id_col": id_col,
                    "name_col": name_col,
                    "total_col": total_col,
                    "teacher_col": teacher_col,
                }
            )
        if not column_blocks:
            continue

        rows = []
        for _, row in raw.iloc[header_row + 1:].iterrows():
            for block in column_blocks:
                id_col = block["id_col"]
                name_col = block["name_col"]
                total_col = block["total_col"]
                teacher_col = block["teacher_col"]
                student_id = _normalize_text(row.iloc[id_col] if len(row) > id_col else "")
                name = _normalize_text(row.iloc[name_col] if len(row) > name_col else "")
                total = row.iloc[total_col] if len(row) > total_col else None
                if not student_id or not name or pd.isna(total):
                    continue
                rows.append(
                    {
                        "学号": student_id,
                        "姓名": name,
                        "总评成绩": total,
                        "指导老师": row.iloc[teacher_col] if teacher_col is not None and len(row) > teacher_col else "",
                    }
                )
        if rows:
            return pd.DataFrame(rows).drop_duplicates(subset=["学号"], keep="first")
    return pd.DataFrame()


def _build_from_existing_target_scores(raw_df: pd.DataFrame, course_df: pd.DataFrame) -> pd.DataFrame:
    info = _course_info(course_df)
    df = raw_df.copy()
    df["学号"] = df["学号"].apply(_normalize_text)
    df["姓名"] = df["姓名"].apply(_normalize_text)
    df["课程目标编号"] = df["课程目标编号"].apply(_normalize_target_id)
    df = df[(df["学号"] != "") & (df["姓名"] != "") & (df["课程目标编号"] != "")].copy()
    if "所属课程" not in df.columns:
        df["所属课程"] = info["课程名称"]
    if "所属学期" not in df.columns:
        df["所属学期"] = info["开课学期"]
    if "数据来源说明" not in df.columns:
        df["数据来源说明"] = SOURCE_REAL
    return df


def _build_from_exam_process_total(raw_df: pd.DataFrame, course_df: pd.DataFrame) -> pd.DataFrame:
    info = _course_info(course_df)
    total_col = next((c for c in ["平时成绩总分", "过程性考核成绩", "平时成绩"] if c in raw_df.columns), None)
    if total_col is None:
        return pd.DataFrame()

    target_ids = course_df["课程目标编号"].dropna().astype(str).tolist()
    rows = []
    for _, student in raw_df.iterrows():
        student_id = _normalize_text(student.get("学号", ""))
        name = _normalize_text(student.get("姓名", ""))
        process_total = pd.to_numeric(student.get(total_col), errors="coerce")
        if not student_id or not name or pd.isna(process_total):
            continue
        ratio = float(process_total) / 40
        for target_id in target_ids:
            rows.append(
                {
                    "学号": student_id,
                    "姓名": name,
                    "课程目标编号": target_id,
                    "课程作业": EXAM_ITEM_FULL_MARKS["课程作业"] * ratio,
                    "实验操作": EXAM_ITEM_FULL_MARKS["实验操作"] * ratio,
                    "实验报告": EXAM_ITEM_FULL_MARKS["实验报告"] * ratio,
                    "课堂表现": EXAM_ITEM_FULL_MARKS["课堂表现"] * ratio,
                    "所属课程": info["课程名称"],
                    "所属学期": info["开课学期"],
                    "数据来源说明": SOURCE_WEIGHT_SPLIT,
                }
            )
    return pd.DataFrame(rows)


def _build_from_final_only(raw_df: pd.DataFrame, course_df: pd.DataFrame) -> pd.DataFrame:
    info = _course_info(course_df)
    total_col = next((c for c in ["总评成绩", "总评", "总分", "最终成绩"] if c in raw_df.columns), None)
    if total_col is None:
        return pd.DataFrame()

    weight_cols = _target_weight_cols(course_df)
    score_cols = [_score_col_from_weight_col(col) for col in weight_cols]
    rows = []
    for _, student in raw_df.iterrows():
        student_id = _normalize_text(student.get("学号", ""))
        name = _normalize_text(student.get("姓名", ""))
        total = pd.to_numeric(student.get(total_col), errors="coerce")
        if not student_id or not name or pd.isna(total):
            continue
        for _, target in course_df.iterrows():
            row = {
                "学号": student_id,
                "姓名": name,
                "课程目标编号": target["课程目标编号"],
                "总评成绩": total,
                "所属课程": info["课程名称"],
                "所属学期": info["开课学期"],
                "数据来源说明": SOURCE_WEIGHT_SPLIT,
            }
            if "指导老师" in raw_df.columns:
                row["指导老师"] = student.get("指导老师", "")
            for score_col in score_cols:
                row[score_col] = total
            rows.append(row)
    return pd.DataFrame(rows)


def _detect_and_build(raw_file: Path, course_df: pd.DataFrame, is_exam_course: bool) -> pd.DataFrame:
    if raw_file.name.startswith("副本材料成型") and "达成情况评价" in pd.ExcelFile(raw_file).sheet_names:
        return _read_material_achievement_sheet(raw_file, course_df)

    raw_df = _read_flat_sheet(raw_file)
    if not raw_df.empty:
        if "课程目标编号" in raw_df.columns:
            return _build_from_existing_target_scores(raw_df, course_df)
        if is_exam_course:
            process_df = _build_from_exam_process_total(raw_df, course_df)
            if not process_df.empty:
                return process_df
        final_df = _build_from_final_only(raw_df, course_df)
        if not final_df.empty:
            return final_df

    final_only_df = _read_final_only_from_wide_sheet(raw_file)
    if not final_only_df.empty:
        return _build_from_final_only(final_only_df, course_df)

    raise ValueError(f"无法识别原始成绩表结构: {raw_file}")


def _expected_score_full_marks(standard_df: pd.DataFrame, course_df: pd.DataFrame, is_exam_course: bool) -> dict:
    if is_exam_course:
        return {col: full for col, full in EXAM_ITEM_FULL_MARKS.items() if col in standard_df.columns}
    full_marks = {}
    for col in standard_df.columns:
        if not str(col).endswith("成绩"):
            continue
        numeric = pd.to_numeric(standard_df[col], errors="coerce")
        full_marks[col] = 100 if numeric.max(skipna=True) > 30 else numeric.max(skipna=True)
    return full_marks


def _build_review_report(
    standard_df: pd.DataFrame,
    course_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    is_exam_course: bool,
) -> dict:
    student_ids = standard_df["学号"].astype(str)
    target_counts = standard_df.groupby("课程目标编号")["学号"].nunique().reset_index(name="人数")
    source_values = standard_df["数据来源说明"].dropna().astype(str).drop_duplicates().tolist()
    full_marks = _expected_score_full_marks(standard_df, course_df, is_exam_course)
    over_limit_rows = []
    for col, full_mark in full_marks.items():
        values = pd.to_numeric(standard_df[col], errors="coerce")
        over_count = int((values > full_mark).sum())
        if over_count:
            over_limit_rows.append({"字段": col, "满分": full_mark, "超出满分记录数": over_count})

    total_review = _build_total_review(standard_df, course_df)
    summary = pd.DataFrame(
        [
            {"核查项": "学生人数", "结果": int(student_ids.nunique())},
            {"核查项": "是否存在缺失学号", "结果": "是" if (student_ids == "").any() else "否"},
            {"核查项": "是否存在重复学号", "结果": "是" if student_ids.duplicated().any() else "否"},
            {"核查项": "学号是否按字符串读取", "结果": "是"},
            {"核查项": "课程目标编号是否统一", "结果": "是" if standard_df["课程目标编号"].astype(str).str.startswith("课程目标").all() else "否"},
            {"核查项": "每个课程目标人数是否一致", "结果": "是" if target_counts["人数"].nunique() == 1 else "否"},
            {"核查项": "是否存在空成绩", "结果": "是" if standard_df.filter(regex="成绩$|课程作业|实验操作|实验报告|课堂表现").isna().any().any() else "否"},
            {"核查项": "是否存在超出满分的成绩", "结果": "是" if over_limit_rows else "否"},
            {"核查项": "原始总评成绩与转换后复核总评成绩的误差", "结果": _total_review_summary(total_review)},
            {"核查项": "数据来源说明", "结果": "；".join(source_values)},
        ]
    )
    return {
        "核查汇总": summary,
        "课程目标人数": target_counts,
        "超出满分记录": pd.DataFrame(over_limit_rows),
        "总评复核": total_review,
    }


def _build_total_review(standard_df: pd.DataFrame, course_df: pd.DataFrame) -> pd.DataFrame:
    if "总评成绩" not in standard_df.columns:
        return pd.DataFrame([{"说明": "标准02中未提供原始总评成绩，无法复核总评误差。"}])
    weight_cols = _target_weight_cols(course_df)
    score_cols = [_score_col_from_weight_col(col) for col in weight_cols]
    rows = []
    for student_id, grp in standard_df.groupby("学号"):
        raw_total = pd.to_numeric(grp["总评成绩"], errors="coerce").dropna()
        if raw_total.empty:
            continue
        weighted_total = 0.0
        total_weight = 0.0
        observed_scores = []
        for _, target in course_df.iterrows():
            target_id = target["课程目标编号"]
            target_row = grp[grp["课程目标编号"] == target_id]
            if target_row.empty:
                continue
            target_row = target_row.iloc[0]
            for weight_col, score_col in zip(weight_cols, score_cols):
                if score_col in target_row:
                    score = pd.to_numeric(target_row[score_col], errors="coerce")
                    weight = pd.to_numeric(target[weight_col], errors="coerce")
                    if pd.notna(score) and pd.notna(weight):
                        weighted_total += float(score) * float(weight)
                        total_weight += float(weight)
                        observed_scores.append(float(score))
        raw_total_value = float(raw_total.iloc[0])
        is_final_only_split = (
            "数据来源说明" in grp.columns
            and grp["数据来源说明"].astype(str).eq(SOURCE_WEIGHT_SPLIT).all()
            and observed_scores
            and all(abs(score - raw_total_value) < 1e-9 for score in observed_scores)
        )
        review_total = raw_total_value if is_final_only_split else weighted_total
        rows.append(
            {
                "学号": student_id,
                "姓名": grp["姓名"].iloc[0],
                "原始总评成绩": raw_total_value,
                "转换后复核总评成绩": review_total,
                "误差": review_total - raw_total_value,
            }
        )
    return pd.DataFrame(rows) if rows else pd.DataFrame([{"说明": "未能生成总评复核明细。"}])


def _total_review_summary(total_review: pd.DataFrame) -> str:
    if "误差" not in total_review.columns:
        return str(total_review.iloc[0].get("说明", "无法复核"))
    abs_error = total_review["误差"].abs()
    return f"最大绝对误差 {abs_error.max():.6f}，平均绝对误差 {abs_error.mean():.6f}"


def generate_standard_score_table(project_root: Path, course_path: Path) -> dict:
    raw_inputs_dir = project_root / "raw_inputs"
    course_df = _read_course_df(course_path)
    course_name = str(course_df["课程名称"].iloc[0])
    raw_file = _find_raw_score_file([course_path, raw_inputs_dir], course_name)
    is_exam_course = has_exam_inputs(course_path)

    standard_df = _detect_and_build(raw_file, course_df, is_exam_course)
    if standard_df.empty:
        raise ValueError(f"原始成绩表未生成任何标准02记录: {raw_file}")

    output_path = course_path / "02_学生成绩表.xlsx"
    report_path = course_path / "02_学生成绩表_生成核查报告.xlsx"
    standard_df.to_excel(output_path, index=False)

    raw_flat = _read_flat_sheet(raw_file)
    report_sheets = _build_review_report(standard_df, course_df, raw_flat, is_exam_course)
    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        for sheet_name, df in report_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    debug_print(f"标准02生成完成: {output_path} shape={standard_df.shape}")
    debug_print(f"核查报告生成完成: {report_path}")
    return {
        "score_table": output_path,
        "review_report": report_path,
        "raw_file": raw_file,
    }
