"""
Excel 报表合并/拆分工具 - 命令行入口

支持以下三个子命令：
  merge  - 将多个 Excel 文件合并为一个（按行追加或按 Sheet 分页）
  split  - 将一个 Excel 文件拆分为多个（按 Sheet 或按列值）
  info   - 查看 Excel 文件的基本信息（Sheet 名、行列数等）

用法：python main.py --help
"""
import argparse  # 标准库：用于解析命令行参数
import sys       # 标准库：用于退出程序及标准错误输出

# 从工具模块导入核心功能类
from excel_tool import ExcelInfo, ExcelMerger, ExcelSplitter


def cmd_merge(args: argparse.Namespace) -> None:
    """处理 merge 子命令：将多个 Excel 文件合并为一个。

    优先级：--folder > --files
      - 若指定了 --folder，则合并该文件夹下所有 .xlsx 文件。
      - 若指定了 --files，则根据 --mode 决定合并方式：
          row   : 按行依次追加所有文件的数据
          sheet : 每个文件作为输出工作簿中的独立 Sheet
    """
    if args.folder:
        # 合并整个文件夹下的所有 Excel 文件
        ExcelMerger.merge_folder(
            folder_path=args.folder,
            output_file=args.output,
            mode=args.mode,
            add_source=not args.no_source,  # --no-source 时不添加来源列
        )
    elif args.files:
        if args.mode == "row":
            # 按行追加：将所有文件的数据纵向合并到同一个 Sheet
            ExcelMerger.merge_by_row(
                file_list=args.files,
                output_file=args.output,
                add_source=not args.no_source,
            )
        elif args.mode == "sheet":
            # 按 Sheet 合并：每个文件单独占一个 Sheet
            ExcelMerger.merge_by_sheet(
                file_list=args.files,
                output_file=args.output,
            )
        else:
            # 不应出现，argparse 已通过 choices 限制合法值
            print(f"❌ 未知合并模式：{args.mode}", file=sys.stderr)
            sys.exit(1)
    else:
        # 既未指定 --files 也未指定 --folder，给出提示后退出
        print("❌ 请通过 --files 指定文件或通过 --folder 指定文件夹", file=sys.stderr)
        sys.exit(1)


def cmd_split(args: argparse.Namespace) -> None:
    """处理 split 子命令：将一个 Excel 文件拆分为多个文件。

    支持两种拆分模式（由 --mode 指定）：
      sheet  : 按 Sheet 拆分，每个 Sheet 输出为独立 Excel 文件
      column : 按指定列的不同取值拆分，每个唯一值输出一个 Excel 文件
               （需同时通过 --column 指定列名）
    """
    if args.mode == "sheet":
        # 按工作表拆分：每个 Sheet 单独保存为一个文件
        ExcelSplitter.split_by_sheet(
            input_file=args.file,
            output_dir=args.output,
        )
    elif args.mode == "column":
        # 按列值拆分前，检查是否提供了列名参数
        if not args.column:
            print("❌ 按列值拆分时需要通过 --column 指定列名", file=sys.stderr)
            sys.exit(1)
        # 按指定列的唯一值拆分，每个取值对应一个输出文件
        ExcelSplitter.split_by_column(
            input_file=args.file,
            column_name=args.column,
            output_dir=args.output,
        )
    else:
        # 不应出现，argparse 已通过 choices 限制合法值
        print(f"❌ 未知拆分模式：{args.mode}", file=sys.stderr)
        sys.exit(1)


def cmd_info(args: argparse.Namespace) -> None:
    """处理 info 子命令：打印指定 Excel 文件的基本信息。

    输出内容包括：Sheet 数量与名称、每个 Sheet 的行列数等。
    """
    ExcelInfo.print_info(args.file)


def build_parser() -> argparse.ArgumentParser:
    """构建并返回命令行参数解析器。

    顶层解析器包含三个子命令：merge、split、info。
    每个子命令拥有独立的参数组，通过 set_defaults(func=...) 绑定对应处理函数，
    从而在 main() 中可以统一调用 args.func(args)。

    Returns:
        argparse.ArgumentParser: 配置好的参数解析器实例。
    """
    # 创建顶层解析器
    parser = argparse.ArgumentParser(
        prog="main.py",
        description="📊 Excel 报表合并/拆分工具",
    )
    # 添加子命令注册器；required=True 确保用户必须提供子命令
    subparsers = parser.add_subparsers(dest="command", metavar="COMMAND")
    subparsers.required = True

    # ── merge 子命令 ────────────────────────────────────────────────────
    merge_parser = subparsers.add_parser("merge", help="合并多个 Excel 文件")
    merge_parser.add_argument(
        "--mode",
        choices=["row", "sheet"],
        default="row",
        help="合并模式：row=按行追加（默认），sheet=每个文件作为独立 Sheet",
    )
    merge_parser.add_argument(
        "--files",
        nargs="+",      # 接受一个或多个文件路径
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
        action="store_true",    # 出现该标志则为 True，不出现则为 False
        help="按行合并时不添加来源文件列",
    )
    # 将 cmd_merge 函数绑定到 merge 子命令，main() 中通过 args.func 调用
    merge_parser.set_defaults(func=cmd_merge)

    # ── split 子命令 ────────────────────────────────────────────────────
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
    # 将 cmd_split 函数绑定到 split 子命令
    split_parser.set_defaults(func=cmd_split)

    # ── info 子命令 ─────────────────────────────────────────────────────
    info_parser = subparsers.add_parser("info", help="查看 Excel 文件信息")
    info_parser.add_argument(
        "--file",
        required=True,
        metavar="FILE",
        help="要查看的 Excel 文件路径",
    )
    # 将 cmd_info 函数绑定到 info 子命令
    info_parser.set_defaults(func=cmd_info)

    return parser


def main() -> None:
    """程序主入口：解析命令行参数并分发到对应的子命令处理函数。

    错误处理策略：
      - FileNotFoundError / ValueError：由底层工具函数抛出，在此统一捕获并输出友好提示。
      - KeyboardInterrupt：用户按 Ctrl+C 中断时优雅退出，返回退出码 130（Unix 惯例）。
    """
    # 构建解析器并解析用户输入的命令行参数
    parser = build_parser()
    args = parser.parse_args()

    try:
        # 调用通过 set_defaults(func=...) 绑定的子命令处理函数
        args.func(args)
    except (FileNotFoundError, ValueError) as exc:
        # 文件不存在或参数值非法时，打印错误信息并以非零状态退出
        print(f"❌ 错误：{exc}", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        # 用户中断（Ctrl+C），换行后提示并以状态码 130 退出
        print("\n已取消", file=sys.stderr)
        sys.exit(130)


# 仅在直接运行本脚本时执行 main()，作为模块导入时不自动执行
if __name__ == "__main__":
    main()
