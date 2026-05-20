from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd


PASS_THRESHOLD = 0.60


def _prepare_matplotlib() -> None:
    plt.rcParams["font.sans-serif"] = [
        "SimHei",
        "Microsoft YaHei",
        "Arial Unicode MS",
        "DejaVu Sans",
    ]
    plt.rcParams["axes.unicode_minus"] = False


def _find_history_file(course_path: Path) -> Path:
    candidates = sorted(course_path.glob("07_*.xlsx"))
    if not candidates:
        raise FileNotFoundError(f"在 {course_path} 下未找到 07_历年课程目标达成值.xlsx")
    return candidates[0]


def _read_history_df(course_path: Path, course_target_df: pd.DataFrame) -> pd.DataFrame:
    history_file = _find_history_file(course_path)
    xls = pd.ExcelFile(history_file)
    sheet_name = "历年汇总" if "历年汇总" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(history_file, sheet_name=sheet_name)

    rename_map = {}
    if "学年" not in df.columns and "年级" in df.columns:
        rename_map["年级"] = "学年"
    if "综合平均达成值" not in df.columns and "达成值" in df.columns:
        rename_map["达成值"] = "综合平均达成值"
    df = df.rename(columns=rename_map)

    required = ["学年", "课程目标编号", "综合平均达成值"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"历年课程目标达成值文件缺少字段: {missing}")

    df = df[required].copy()
    df["综合平均达成值"] = pd.to_numeric(df["综合平均达成值"], errors="coerce")

    current_values = course_target_df.set_index("课程目标编号")["综合平均达成值"].to_dict()
    for idx, row in df[df["综合平均达成值"].isna()].iterrows():
        target_id = row["课程目标编号"]
        if target_id in current_values:
            df.at[idx, "综合平均达成值"] = current_values[target_id]

    return df.dropna(subset=["综合平均达成值"]).copy()


def _read_student_attainment_df(analysis_workbook_path: Path) -> pd.DataFrame:
    return pd.read_excel(analysis_workbook_path, sheet_name="学生个体达成度表")


def _save_three_year_compare(history_df: pd.DataFrame, output_path: Path) -> None:
    fig, ax = plt.subplots(figsize=(8.5, 5.0))
    for target_id, grp in history_df.groupby("课程目标编号", sort=True):
        grp = grp.sort_values("学年")
        ax.plot(
            grp["学年"].astype(str),
            grp["综合平均达成值"],
            marker="o",
            linewidth=2,
            label=target_id,
        )
    ax.axhline(PASS_THRESHOLD, linestyle="--", linewidth=1.2, color="#B00020", label="合格阈值 0.60")
    ax.set_title("3年课程目标达成度对比图")
    ax.set_xlabel("学年")
    ax.set_ylabel("综合平均达成值")
    ax.set_ylim(0, 1)
    ax.grid(axis="y", linestyle=":", linewidth=0.8, alpha=0.7)
    ax.legend()
    fig.tight_layout()
    fig.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(fig)


def _save_target_scatter(student_df: pd.DataFrame, target_id: str, output_path: Path) -> None:
    df = student_df[student_df["课程目标编号"] == target_id].copy()
    df = df.sort_values(["学号", "姓名"]).reset_index(drop=True)
    df["学生序号"] = range(1, len(df) + 1)

    fig, ax = plt.subplots(figsize=(8.5, 5.0))
    ax.scatter(df["学生序号"], df["综合达成度"], s=28, alpha=0.85)
    ax.axhline(PASS_THRESHOLD, linestyle="--", linewidth=1.2, color="#B00020", label="合格阈值 0.60")
    ax.set_title(f"{target_id}达成度值散点图")
    ax.set_xlabel("学生序号")
    ax.set_ylabel("综合达成度")
    ax.set_ylim(0, 1)
    ax.grid(axis="y", linestyle=":", linewidth=0.8, alpha=0.7)
    ax.legend()
    fig.tight_layout()
    fig.savefig(output_path, dpi=300, bbox_inches="tight")
    plt.close(fig)


def export_charts(
    course_path: Path,
    output_dir: Path,
    analysis_workbook_path: Path,
    course_target_df: pd.DataFrame,
) -> dict:
    figures_dir = output_dir / "figures"
    figures_dir.mkdir(parents=True, exist_ok=True)
    _prepare_matplotlib()

    paths = {
        "three_year_compare": figures_dir / "01_三年课程目标达成度对比.png",
        "target1_scatter": figures_dir / "02_课程目标1达成度散点图.png",
        "target2_scatter": figures_dir / "03_课程目标2达成度散点图.png",
    }

    history_df = _read_history_df(course_path, course_target_df)
    student_df = _read_student_attainment_df(analysis_workbook_path)

    _save_three_year_compare(history_df, paths["three_year_compare"])
    _save_target_scatter(student_df, "课程目标1", paths["target1_scatter"])
    _save_target_scatter(student_df, "课程目标2", paths["target2_scatter"])

    return paths
