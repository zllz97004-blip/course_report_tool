
from pathlib import Path

from .analyzer import build_analysis_base
from .calculator import build_course_target_result, build_student_target_result
from .config import debug_print
from .exporter import export_analysis_workbook, export_report_docx, export_result_workbook
from .loaders import load_all_inputs
from .validators import validate_inputs


def _find_course_folder(input_root: Path) -> Path:
    folders = [p for p in input_root.iterdir() if p.is_dir()]
    if not folders:
        raise FileNotFoundError("input_courses 下没有课程文件夹。")
    # 取最后修改时间最新的课程文件夹
    folders.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return folders[0]


def main():
    project_root = Path(__file__).resolve().parent.parent
    input_root = project_root / "input_courses"
    course_path = _find_course_folder(input_root)

    output_dir = project_root / "output" / course_path.name
    output_dir.mkdir(parents=True, exist_ok=True)

    data = load_all_inputs(course_path)
    validate_inputs(data)

    student_target_df = build_student_target_result(
        course_df=data["course"],
        score_df=data["scores"],
        mapping_df=data["mapping"],
        long_score_df=data["long_scores"],
    )

    debug_print("\n===== student_target_df 关键列预览 =====")
    debug_print(
        student_target_df[
            ["学号", "姓名", "课程目标编号", "试卷得分", "试卷满分", "试卷达成值", "期末贡献值", "综合达成值"]
        ].head(10)
    )

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
    export_analysis_workbook(output_dir / "06_达成度报告分析底表.xlsx", analysis_df)
    export_report_docx(
        output_path=output_dir / "07_达成度报告草稿.docx",
        course_df=data["course"],
        analysis_df=analysis_df,
        student_target_df=student_target_df,
        course_path=course_path,
    )

    print(f"处理完成，输出目录：{output_dir}")


if __name__ == "__main__":
    main()
