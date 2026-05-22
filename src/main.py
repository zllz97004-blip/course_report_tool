
import os
from pathlib import Path

from .analyzer import build_analysis_base
from .calculator import build_course_target_result, build_student_target_result
from .chart_exporter import export_charts
from .config import debug_print
from .exporter import export_analysis_workbook, export_result_workbook
from .loaders import load_all_inputs
from .non_exam_calculator import build_non_exam_student_target_result
from .preprocess_raw_inputs import preprocess_raw_inputs
from .report_exporter import export_report_docx
from .validators import validate_inputs


COURSE_DIR_NAME = None
PREPROCESS_RAW_INPUTS = False


def _find_course_folder(input_root: Path, course_dir_name: str = None) -> Path:
    selected_name = os.environ.get("COURSE_DIR_NAME") or course_dir_name
    if selected_name:
        selected = input_root / selected_name
        if not selected.is_dir():
            raise FileNotFoundError(f"未找到指定课程目录: {selected}")
        return selected

    folders = [p for p in input_root.iterdir() if p.is_dir()]
    if not folders:
        raise FileNotFoundError("input_courses 下没有课程文件夹。")
    # 取最后修改时间最新的课程文件夹
    folders.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return folders[0]


def main():
    project_root = Path(__file__).resolve().parent.parent
    input_root = project_root / "input_courses"
    course_path = _find_course_folder(input_root, COURSE_DIR_NAME)

    output_dir = project_root / "output" / course_path.name
    output_dir.mkdir(parents=True, exist_ok=True)

    if PREPROCESS_RAW_INPUTS:
        preprocess_raw_inputs(project_root, course_path)

    data = load_all_inputs(course_path)
    validate_inputs(data)

    if data["course_type"] == "exam":
        student_target_df = build_student_target_result(
            course_df=data["course"],
            score_df=data["scores"],
            mapping_df=data["mapping"],
            long_score_df=data["long_scores"],
        )
    else:
        student_target_df = build_non_exam_student_target_result(
            course_df=data["course"],
            score_df=data["scores"],
        )

    debug_print("\n===== student_target_df 关键列预览 =====")
    preview_cols = ["学号", "姓名", "课程目标编号", "过程性贡献值", "期末贡献值", "综合达成值"]
    exam_preview_cols = ["试卷得分", "试卷满分", "试卷达成值"]
    preview_cols = preview_cols[:3] + [col for col in exam_preview_cols if col in student_target_df.columns] + preview_cols[3:]
    debug_print(student_target_df[preview_cols].head(10))

    course_target_df = build_course_target_result(student_target_df)

    debug_print("\n===== course_target_df 关键列预览 =====")
    debug_print(
        course_target_df[
            ["课程目标编号", "试卷理论总分", "试卷平均得分", "期末平均贡献值", "综合平均达成值"]
        ]
    )

    analysis_df = build_analysis_base(
        course_df=data["course"],
        course_target_df=course_target_df,
    )

    debug_print("\n===== 写入 06 前 analysis_df head =====")
    debug_print(analysis_df.head())

    export_result_workbook(output_dir / "05_课程目标结果表.xlsx", student_target_df, course_target_df)
    analysis_workbook_path = output_dir / "06_达成度报告分析底表.xlsx"
    export_analysis_workbook(
        analysis_workbook_path,
        analysis_df,
        student_target_df=student_target_df,
        course_target_df=course_target_df,
    )
    chart_paths = export_charts(
        course_path=course_path,
        output_dir=output_dir,
        analysis_workbook_path=analysis_workbook_path,
        course_target_df=course_target_df,
    )
    export_report_docx(
        output_path=output_dir / "07_达成度报告草稿.docx",
        course_df=data["course"],
        course_target_df=course_target_df,
        student_target_df=student_target_df,
        chart_paths=chart_paths,
    )

    print(f"处理完成，输出目录：{output_dir}")


if __name__ == "__main__":
    main()
