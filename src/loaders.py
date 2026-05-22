from pathlib import Path
import pandas as pd

from .config import debug_print


COLUMN_ALIASES = {
    "课程名称": ["课程名称", "课程"],
    "课程代码": ["课程代码", "代码"],
    "开课学期": ["开课学期", "学年学期", "学期"],

    "学号": ["学号"],
    "姓名": ["姓名"],

    "平时作业": ["平时作业", "平时作业成绩", "课程作业", "课程作业成绩", "平时作业50"],
    "实验": ["实验", "实验成绩", "实验操作", "实验操作成绩", "实验20"],
    "实验报告": ["实验报告", "实验报告成绩", "实验报告20"],
    "课堂表现": ["课堂表现", "课堂表现成绩", "考勤", "考勤成绩", "课堂表现10"],

    "所属课程": ["所属课程"],
    "所属学期": ["所属学期"],

    "课程目标编号": ["课程目标编号", "课程目标", "目标编号"],
    "课程目标描述": ["课程目标描述", "目标描述"],

    "课堂表现权重": ["课堂表现权重", "考勤权重"],
    "课程作业权重": ["课程作业权重"],
    "实验报告权重": ["实验报告权重"],
    "实验操作权重": ["实验操作权重"],
    "期末权重": ["期末权重"],

    "题号": ["题号"],
    "题型": ["题型"],
    "满分": ["满分"],

    "大题": ["大题"],
    "小题号": ["小题号"],
    "学生得分": ["学生得分", "得分"],
}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    cleaned_cols = []
    for c in df.columns:
        c = (
            str(c)
            .replace("\n", "")
            .replace("\r", "")
            .replace(" ", "")
            .replace("\u3000", "")
            .strip()
        )

        # 学生成绩表常见子表头归一
        if c.startswith("平时作业"):
            c = "平时作业"
        elif c.startswith("实验报告"):
            c = "实验报告"
        elif c.startswith("课堂表现"):
            c = "课堂表现"
        elif c.startswith("实验"):
            c = "实验"

        cleaned_cols.append(c)

    df.columns = cleaned_cols

    rename_map = {}
    used_source_cols = set()

    for standard_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            alias_clean = (
                str(alias)
                .replace("\n", "")
                .replace("\r", "")
                .replace(" ", "")
                .replace("\u3000", "")
                .strip()
            )
            for col in df.columns:
                col_clean = (
                    str(col)
                    .replace("\n", "")
                    .replace("\r", "")
                    .replace(" ", "")
                    .replace("\u3000", "")
                    .strip()
                )
                if col_clean == alias_clean and col not in used_source_cols:
                    rename_map[col] = standard_name
                    used_source_cols.add(col)
                    break
            if standard_name in rename_map.values():
                break

    return df.rename(columns=rename_map)


def load_excel(path: Path) -> pd.DataFrame:
    return normalize_columns(pd.read_excel(path))


def _find_single_file(course_path: Path, exact_name: str, pattern: str) -> Path:
    exact = course_path / exact_name
    if exact.exists():
        return exact
    candidates = sorted(course_path.glob(pattern))
    if not candidates:
        raise FileNotFoundError(f"未找到 {exact_name}，也未找到匹配 {pattern} 的文件。")
    return candidates[0]


def has_exam_inputs(course_path: Path) -> bool:
    mapping_file = course_path / "03_试卷分值对应表.xlsx"
    long_score_files = [
        course_path / "04_试卷题目得分长表.xlsx",
        course_path / "04_试卷题目得分表_长表.xlsx",
    ]
    return mapping_file.exists() and any(p.exists() for p in long_score_files)


def load_course_excel(path: Path) -> pd.DataFrame:
    df = load_excel(path)
    debug_print(">>> 使用的是新版 load_course_excel")
    required_cols = {"课程名称", "课程代码", "开课学期", "课程目标编号", "课程目标描述"}
    if not required_cols.issubset(set(df.columns)):
        raise ValueError(f"课程主数据表缺少必要字段，当前列名: {list(df.columns)}")

    exam_weight_cols = {"课堂表现权重", "课程作业权重", "实验报告权重", "实验操作权重", "期末权重"}
    if exam_weight_cols.intersection(set(df.columns)):
        if "课堂表现权重" not in df.columns:
            df["课堂表现权重"] = 0.04
        if "课程作业权重" not in df.columns:
            df["课程作业权重"] = 0.20
        if "实验报告权重" not in df.columns:
            df["实验报告权重"] = 0.08
        if "实验操作权重" not in df.columns:
            df["实验操作权重"] = 0.08
        if "期末权重" not in df.columns:
            df["期末权重"] = 0.60

    return df

def load_scores_excel(path: Path) -> pd.DataFrame:
    """
    兼容两种学生成绩表：
    1. 旧版：双层表头宽表
    2. 新版：单层表头长表（每个学生 × 每个课程目标一行）
    """

    # 先按单层表头读取，优先识别新版长表
    df1 = pd.read_excel(path, header=0)
    df1 = normalize_columns(df1)

    long_required = {"学号", "姓名", "课程目标编号", "平时作业", "实验", "实验报告", "课堂表现"}
    if long_required.issubset(set(df1.columns)):
        return df1

    non_exam_base = {"学号", "姓名", "课程目标编号"}
    non_exam_score_cols = [col for col in df1.columns if str(col).endswith("成绩")]
    if non_exam_base.issubset(set(df1.columns)) and non_exam_score_cols:
        return df1

    # 若不是新版长表，再尝试按旧版双层表头读取
    raw = pd.read_excel(path, header=[0, 1])

    if isinstance(raw.columns, pd.MultiIndex):
        new_cols = []
        for top, sub in raw.columns:
            top = "" if pd.isna(top) else str(top)
            sub = "" if pd.isna(sub) else str(sub)

            top = top.replace("\n", "").replace("\r", "").replace(" ", "").replace("\u3000", "").strip()
            sub = sub.replace("\n", "").replace("\r", "").replace(" ", "").replace("\u3000", "").strip()

            if sub and not sub.startswith("Unnamed:"):
                new_cols.append(sub)
            else:
                new_cols.append(top)

        raw.columns = new_cols
        df = normalize_columns(raw)
    else:
        df = normalize_columns(raw)

    # 旧版宽表兜底：如果表头里有“平时作业50 / 实验20 ...”这种字段，重命名
    cols = list(df.columns)
    if not {"平时作业", "实验", "实验报告", "课堂表现"}.issubset(set(cols)):
        renamed = []
        for c in cols:
            c2 = str(c).replace(" ", "").replace("\u3000", "")
            if c2.startswith("平时作业"):
                renamed.append("平时作业")
            elif c2.startswith("实验报告"):
                renamed.append("实验报告")
            elif c2.startswith("课堂表现"):
                renamed.append("课堂表现")
            elif c2.startswith("实验"):
                renamed.append("实验")
            else:
                renamed.append(c)
        df.columns = renamed
        df = normalize_columns(df)

    return df

def load_all_inputs(course_path: Path) -> dict:
    course_file = course_path / "01_课程主数据表.xlsx"
    scores_file = _find_single_file(course_path, "02_学生成绩表.xlsx", "02_*.xlsx")
    mapping_file = course_path / "03_试卷分值对应表.xlsx"
    is_exam_course = has_exam_inputs(course_path)

    candidates = [
        course_path / "04_试卷题目得分长表.xlsx",
        course_path / "04_试卷题目得分表_长表.xlsx",
    ]
    long_scores_file = next((p for p in candidates if p.exists()), None)
    if is_exam_course and long_scores_file is None:
        raise FileNotFoundError("未找到 04 长表文件，请检查文件名。")

    course = load_course_excel(course_file)
    scores = load_scores_excel(scores_file)
    mapping = load_excel(mapping_file) if is_exam_course else None
    long_scores = load_excel(long_scores_file) if is_exam_course else None

    debug_items = [
        ("课程主数据表", course_file, course),
        ("学生成绩表", scores_file, scores),
    ]
    if is_exam_course:
        debug_items.extend(
            [
                ("试卷分值对应表", mapping_file, mapping),
                ("题目得分长表", long_scores_file, long_scores),
            ]
        )
    for label, path, df in debug_items:
        debug_print(f"\n===== {label} 读取后信息 =====")
        debug_print("file =", path)
        debug_print("columns =", list(df.columns))
        debug_print("shape =", df.shape)

    if "课程目标编号" in scores.columns:
        debug_print("\n===== 学生成绩表中课程目标编号的唯一值（读取后）=====")
        debug_print(scores["课程目标编号"].dropna().astype(str).drop_duplicates().tolist())
    if mapping is not None and "课程目标编号" in mapping.columns:
        debug_print("\n===== 试卷分值对应表中课程目标编号的唯一值（读取后）=====")
        debug_print(mapping["课程目标编号"].dropna().astype(str).drop_duplicates().tolist())
    long_target_col = None
    if long_scores is not None:
        long_target_col = "课程目标编号" if "课程目标编号" in long_scores.columns else "课程目标"
    if long_scores is not None and long_target_col in long_scores.columns:
        debug_print("\n===== 题目得分长表中课程目标编号的唯一值（读取后）=====")
        debug_print(long_scores[long_target_col].dropna().astype(str).drop_duplicates().tolist())

    return {
        "course_type": "exam" if is_exam_course else "non_exam",
        "course": course,
        "scores": scores,
        "mapping": mapping,
        "long_scores": long_scores,
    }
