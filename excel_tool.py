"""
Excel 报表合并/拆分工具核心模块
"""
import re
from pathlib import Path
from typing import Generator, List, Optional, Dict, Any

import pandas as pd


# ---------------------------------------------------------------------------
# 内部辅助函数
# ---------------------------------------------------------------------------

def _ensure_dir(path: Path) -> None:
    """确保目录存在，不存在则创建。"""
    path.mkdir(parents=True, exist_ok=True)


def _safe_filename(value: str) -> str:
    """将字符串转换为安全的文件名（替换非法字符）。"""
    return re.sub(r'[\\/:*?"<>|]', "_", value)


def _unique_sheet_name(base: str, existing: set) -> str:
    """生成不与已有 Sheet 名称重复的名称（最多 31 字符）。"""
    name = base[:31]
    if name not in existing:
        return name
    for i in range(1, 10000):
        suffix = f"_{i}"
        candidate = base[: 31 - len(suffix)] + suffix
        if candidate not in existing:
            return candidate
    raise ValueError(f"无法为 '{base}' 生成唯一 Sheet 名称")


def _iter_excel_frames(
    file_list: List[str], add_source: bool
) -> Generator[pd.DataFrame, None, None]:
    """逐文件读取 Excel，按需插入来源列，以生成器方式返回，节省内存。"""
    for file_path in file_list:
        df = pd.read_excel(file_path)
        if add_source:
            df.insert(0, "来源文件", Path(file_path).name)
        yield df


# ---------------------------------------------------------------------------
# 合并工具
# ---------------------------------------------------------------------------

class ExcelMerger:
    """Excel 文件合并工具"""

    @staticmethod
    def merge_by_row(
        file_list: List[str],
        output_file: str,
        add_source: bool = True,
        sheet_name: str = "Sheet1",
    ) -> str:
        """
        按行追加合并多个 Excel 文件（纵向拼接）。

        Args:
            file_list: 要合并的 Excel 文件路径列表
            output_file: 输出文件路径
            add_source: 是否添加来源文件标识列
            sheet_name: 目标 Sheet 名称

        Returns:
            输出文件路径
        """
        if not file_list:
            raise ValueError("文件列表为空，无法合并")

        merged = pd.concat(
            _iter_excel_frames(file_list, add_source), ignore_index=True
        )
        output_path = Path(output_file)
        _ensure_dir(output_path.parent)
        merged.to_excel(str(output_path), index=False, sheet_name=sheet_name)
        print(f"✅ 按行合并完成：{output_path}（共 {len(merged)} 行）")
        return str(output_path)

    @staticmethod
    def merge_by_sheet(
        file_list: List[str],
        output_file: str,
    ) -> str:
        """
        按 Sheet 合并多个 Excel 文件（每个文件作为独立 Sheet）。

        Args:
            file_list: 要合并的 Excel 文件路径列表
            output_file: 输出文件路径

        Returns:
            输出文件路径
        """
        if not file_list:
            raise ValueError("文件列表为空，无法合并")

        output_path = Path(output_file)
        _ensure_dir(output_path.parent)

        used_names: set = set()
        with pd.ExcelWriter(str(output_path), engine="openpyxl") as writer:
            for file_path in file_list:
                base = Path(file_path).stem
                sheet_name = _unique_sheet_name(base, used_names)
                used_names.add(sheet_name)
                df = pd.read_excel(file_path)
                df.to_excel(writer, index=False, sheet_name=sheet_name)

        print(f"✅ 按 Sheet 合并完成：{output_path}（共 {len(file_list)} 个 Sheet）")
        return str(output_path)

    @staticmethod
    def merge_folder(
        folder_path: str,
        output_file: str,
        mode: str = "row",
        add_source: bool = True,
        pattern: str = "*.xlsx",
    ) -> str:
        """
        合并文件夹中所有 Excel 文件。

        Args:
            folder_path: 文件夹路径
            output_file: 输出文件路径
            mode: 合并模式，"row" 或 "sheet"
            add_source: 按行合并时是否添加来源列
            pattern: 文件匹配模式

        Returns:
            输出文件路径
        """
        folder = Path(folder_path)
        if not folder.is_dir():
            raise ValueError(f"路径不是文件夹：{folder_path}")

        file_list = sorted(str(f) for f in folder.glob(pattern) if f.is_file())
        if not file_list:
            raise ValueError(f"文件夹中没有匹配 '{pattern}' 的文件：{folder_path}")

        print(f"📂 找到 {len(file_list)} 个文件")
        if mode == "row":
            return ExcelMerger.merge_by_row(file_list, output_file, add_source=add_source)
        elif mode == "sheet":
            return ExcelMerger.merge_by_sheet(file_list, output_file)
        else:
            raise ValueError(f"未知合并模式：{mode}，请使用 'row' 或 'sheet'")


# ---------------------------------------------------------------------------
# 拆分工具
# ---------------------------------------------------------------------------

class ExcelSplitter:
    """Excel 文件拆分工具"""

    @staticmethod
    def split_by_sheet(
        input_file: str,
        output_dir: str,
    ) -> List[str]:
        """
        按 Sheet 拆分 Excel 文件（每个 Sheet 生成一个文件）。

        Args:
            input_file: 输入 Excel 文件路径
            output_dir: 输出目录路径

        Returns:
            生成的文件路径列表
        """
        output_path = Path(output_dir)
        _ensure_dir(output_path)

        output_files: List[str] = []
        base_name = Path(input_file).stem

        with pd.ExcelFile(input_file) as xl:
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                out_file = output_path / f"{base_name}_{sheet}.xlsx"
                df.to_excel(str(out_file), index=False)
                output_files.append(str(out_file))
                print(f"  📄 {sheet} -> {out_file.name}")

        print(f"✅ 按 Sheet 拆分完成：共生成 {len(output_files)} 个文件")
        return output_files

    @staticmethod
    def split_by_column(
        input_file: str,
        column_name: str,
        output_dir: str,
        sheet_name: Optional[str] = None,
    ) -> List[str]:
        """
        按列值拆分 Excel 文件（根据指定列的值分组）。

        Args:
            input_file: 输入 Excel 文件路径
            column_name: 用于分组的列名
            output_dir: 输出目录路径
            sheet_name: 读取的 Sheet 名称，默认读取第一个 Sheet

        Returns:
            生成的文件路径列表
        """
        output_path = Path(output_dir)
        _ensure_dir(output_path)

        read_kwargs: Dict[str, Any] = {}
        if sheet_name:
            read_kwargs["sheet_name"] = sheet_name
        df = pd.read_excel(input_file, **read_kwargs)

        if column_name not in df.columns:
            raise ValueError(
                f"列 '{column_name}' 不存在。可用列：{list(df.columns)}"
            )

        output_files: List[str] = []
        base_name = Path(input_file).stem

        for value, group in df.groupby(column_name, sort=True):
            safe_value = _safe_filename(str(value))
            out_file = output_path / f"{base_name}_{column_name}_{safe_value}.xlsx"
            group.to_excel(str(out_file), index=False)
            output_files.append(str(out_file))
            print(f"  📄 {column_name}={value}（{len(group)} 行）-> {out_file.name}")

        print(f"✅ 按列值拆分完成：共生成 {len(output_files)} 个文件")
        return output_files


# ---------------------------------------------------------------------------
# 信息查看工具
# ---------------------------------------------------------------------------

class ExcelInfo:
    """Excel 文件信息查看工具"""

    @staticmethod
    def get_info(file_path: str) -> Dict[str, Any]:
        """
        获取 Excel 文件的基本信息。

        Args:
            file_path: Excel 文件路径

        Returns:
            包含文件信息的字典
        """
        path = Path(file_path)
        if not path.exists():
            raise FileNotFoundError(f"文件不存在：{file_path}")

        sheets_info: List[Dict[str, Any]] = []
        with pd.ExcelFile(file_path) as xl:
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                sheets_info.append(
                    {
                        "sheet_name": sheet,
                        "rows": len(df),
                        "columns": len(df.columns),
                        "column_names": list(df.columns),
                    }
                )

        return {
            "file_name": path.name,
            "file_size": f"{path.stat().st_size / 1024:.1f} KB",
            "sheet_count": len(sheets_info),
            "sheets": sheets_info,
        }

    @staticmethod
    def print_info(file_path: str) -> None:
        """打印 Excel 文件信息到控制台"""
        info = ExcelInfo.get_info(file_path)
        print(f"\n📊 文件：{info['file_name']}  大小：{info['file_size']}")
        print(f"   Sheet 数量：{info['sheet_count']}")
        for s in info["sheets"]:
            print(f"   └─ [{s['sheet_name']}]  {s['rows']} 行 × {s['columns']} 列")
            print(f"      列名：{s['column_names']}")
