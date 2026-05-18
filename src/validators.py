REQUIRED_COLUMNS = {
    "course": ["课程名称", "课程代码", "开课学期", "课程目标编号", "课程目标描述"],
   "scores": ["学号", "姓名", "课程目标编号", "平时作业", "实验", "实验报告", "课堂表现", "所属课程", "所属学期"],
    "mapping": ["所属课程", "所属学期", "题号", "题型", "满分", "课程目标编号"],
    "long_scores": ["学号", "姓名", "课程目标编号", "大题", "小题号", "学生得分"],
}


def validate_required_columns(df, required, name):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{name} 缺少字段: {missing}")


def validate_inputs(data: dict) -> None:
    for key, cols in REQUIRED_COLUMNS.items():
        validate_required_columns(data[key], cols, key)

    course_name = str(data["course"]["课程名称"].iloc[0]).strip()
    score_course_names = set(data["scores"]["所属课程"].dropna().astype(str).str.strip())
    mapping_course_names = set(data["mapping"]["所属课程"].dropna().astype(str).str.strip())

    if score_course_names != {course_name}:
        raise ValueError(f"学生成绩表课程名称不一致: {score_course_names} vs {course_name}")

    if mapping_course_names != {course_name}:
        raise ValueError(f"试卷分值对应表课程名称不一致: {mapping_course_names} vs {course_name}")