REQUIRED_COLUMNS = {
    "course": ["课程名称", "课程代码", "开课学期", "课程目标编号", "课程目标描述"],
    "scores": ["学号", "姓名", "课程目标编号", "所属课程", "所属学期"],
    "mapping": ["所属课程", "所属学期", "题号", "题型", "满分", "课程目标编号"],
    "long_scores": ["学号", "姓名", "课程目标编号", "大题", "小题号", "学生得分"],
}


def validate_required_columns(df, required, name):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{name} 缺少字段: {missing}")


def validate_inputs(data: dict) -> None:
    validate_required_columns(data["course"], REQUIRED_COLUMNS["course"], "course")
    validate_required_columns(data["scores"], REQUIRED_COLUMNS["scores"], "scores")
    if data.get("course_type") == "exam":
        validate_required_columns(data["mapping"], REQUIRED_COLUMNS["mapping"], "mapping")
        validate_required_columns(data["long_scores"], REQUIRED_COLUMNS["long_scores"], "long_scores")

    course_name = str(data["course"]["课程名称"].iloc[0]).strip()
    score_course_names = set(data["scores"]["所属课程"].dropna().astype(str).str.strip())

    if score_course_names != {course_name}:
        raise ValueError(f"学生成绩表课程名称不一致: {score_course_names} vs {course_name}")

    if data.get("course_type") == "exam":
        mapping_course_names = set(data["mapping"]["所属课程"].dropna().astype(str).str.strip())
        if mapping_course_names != {course_name}:
            raise ValueError(f"试卷分值对应表课程名称不一致: {mapping_course_names} vs {course_name}")
