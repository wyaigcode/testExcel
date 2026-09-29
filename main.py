"""
Excel 报表合并/拆分工具 - 命令行入口
用法：python main.py --help
"""
import argparse
import sys

from excel_tool import ExcelInfo, ExcelMerger, ExcelSplitter


def cmd_merge(args: argparse.Namespace) -> None:
    """处理 merge 子命令"""
    if args.folder:
        ExcelMerger.merge_folder(
            folder_path=args.folder,
            output_file=args.output,
            mode=args.mode,
            add_source=not args.no_source,
        )
    elif args.files:
        if args.mode == "row":
            ExcelMerger.merge_by_row(
                file_list=args.files,
                output_file=args.output,
                add_source=not args.no_source,
            )
        elif args.mode == "sheet":
            ExcelMerger.merge_by_sheet(
                file_list=args.files,
                output_file=args.output,
            )
        else:
            print(f"❌ 未知合并模式：{args.mode}", file=sys.stderr)
            sys.exit(1)
    else:
        print("❌ 请通过 --files 指定文件或通过 --folder 指定文件夹", file=sys.stderr)
        sys.exit(1)


def cmd_split(args: argparse.Namespace) -> None:
    """处理 split 子命令"""
    if args.mode == "sheet":
        ExcelSplitter.split_by_sheet(
            input_file=args.file,
            output_dir=args.output,
        )
    elif args.mode == "column":
        if not args.column:
            print("❌ 按列值拆分时需要通过 --column 指定列名", file=sys.stderr)
            sys.exit(1)
        ExcelSplitter.split_by_column(
            input_file=args.file,
            column_name=args.column,
            output_dir=args.output,
        )
    else:
        print(f"❌ 未知拆分模式：{args.mode}", file=sys.stderr)
        sys.exit(1)


def cmd_info(args: argparse.Namespace) -> None:
    """处理 info 子命令"""
    ExcelInfo.print_info(args.file)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="main.py",
        description="📊 Excel 报表合并/拆分工具",
    )
    subparsers = parser.add_subparsers(dest="command", metavar="COMMAND")
    subparsers.required = True

    # ── merge ──────────────────────────────────────────────────────────
    merge_parser = subparsers.add_parser("merge", help="合并多个 Excel 文件")
    merge_parser.add_argument(
        "--mode",
        choices=["row", "sheet"],
        default="row",
        help="合并模式：row=按行追加（默认），sheet=每个文件作为独立 Sheet",
    )
    merge_parser.add_argument(
        "--files",
        nargs="+",
        metavar="FILE",
        help="要合并的 Excel 文件路径（可多个）",
    )
    merge_parser.add_argument(
        "--folder",
        metavar="DIR",
        help="合并指定文件夹中所有 .xlsx 文件",
    )
    merge_parser.add_argument(
        "-o",
        "--output",
        required=True,
        metavar="OUTPUT",
        help="输出文件路径",
    )
    merge_parser.add_argument(
        "--no-source",
        action="store_true",
        help="按行合并时不添加来源文件列",
    )
    merge_parser.set_defaults(func=cmd_merge)

    # ── split ──────────────────────────────────────────────────────────
    split_parser = subparsers.add_parser("split", help="拆分 Excel 文件")
    split_parser.add_argument(
        "--mode",
        choices=["sheet", "column"],
        required=True,
        help="拆分模式：sheet=按 Sheet 拆分，column=按列值拆分",
    )
    split_parser.add_argument(
        "--file",
        required=True,
        metavar="FILE",
        help="要拆分的 Excel 文件路径",
    )
    split_parser.add_argument(
        "--column",
        metavar="COL",
        help="按列值拆分时指定的列名（--mode column 时必填）",
    )
    split_parser.add_argument(
        "-o",
        "--output",
        required=True,
        metavar="DIR",
        help="输出目录路径",
    )
    split_parser.set_defaults(func=cmd_split)

    # ── info ───────────────────────────────────────────────────────────
    info_parser = subparsers.add_parser("info", help="查看 Excel 文件信息")
    info_parser.add_argument(
        "--file",
        required=True,
        metavar="FILE",
        help="要查看的 Excel 文件路径",
    )
    info_parser.set_defaults(func=cmd_info)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    try:
        args.func(args)
    except (FileNotFoundError, ValueError) as exc:
        print(f"❌ 错误：{exc}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n已取消", file=sys.stderr)
        sys.exit(130)


if __name__ == "__main__":
    main()
