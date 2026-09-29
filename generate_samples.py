"""
示例数据生成脚本
运行：python generate_samples.py
会在 ./sample_data/ 目录下生成演示用的 Excel 文件。
"""
from pathlib import Path

import pandas as pd


def main() -> None:
    out_dir = Path("sample_data")
    out_dir.mkdir(exist_ok=True)

    # ── data1.xlsx ──────────────────────────────────────────────────────
    df1 = pd.DataFrame(
        {
            "姓名": ["张三", "李四", "王五"],
            "部门": ["研发", "研发", "市场"],
            "薪资": [12000, 15000, 11000],
        }
    )
    df1.to_excel(out_dir / "data1.xlsx", index=False)

    # ── data2.xlsx ──────────────────────────────────────────────────────
    df2 = pd.DataFrame(
        {
            "姓名": ["赵六", "孙七", "周八"],
            "部门": ["市场", "运营", "运营"],
            "薪资": [10000, 9000, 9500],
        }
    )
    df2.to_excel(out_dir / "data2.xlsx", index=False)

    # ── multi_sheet.xlsx ────────────────────────────────────────────────
    with pd.ExcelWriter(out_dir / "multi_sheet.xlsx", engine="openpyxl") as writer:
        df1.to_excel(writer, sheet_name="一月", index=False)
        df2.to_excel(writer, sheet_name="二月", index=False)
        pd.concat([df1, df2], ignore_index=True).to_excel(
            writer, sheet_name="汇总", index=False
        )

    print("✅ 示例文件已生成到 ./sample_data/")
    print("   data1.xlsx       - 3 行员工数据")
    print("   data2.xlsx       - 3 行员工数据")
    print("   multi_sheet.xlsx - 含 3 个 Sheet 的文件（一月/二月/汇总）")


if __name__ == "__main__":
    main()
