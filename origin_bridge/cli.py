# -*- coding: utf-8 -*-
"""
origin_bridge.cli — 命令行工具

安装后可直接使用 ``origin-cli`` 命令::

    origin-cli info                          查看 Origin 环境与语言状态
    origin-cli lang                          查看当前界面语言
    origin-cli lang zh                       切换到中文
    origin-cli lang en --restart             切换到英文并重启 Origin
    origin-cli lang --versions               列出所有已安装版本
    origin-cli import data.csv --book MyData 导入数据文件
    origin-cli export out.txt                导出当前表数据
    origin-cli stats A                       查看 A 列统计量
    origin-cli plot A B --type scatter       用 A、B 两列绘图
    origin-cli save out.opju                 保存项目

未安装时也可用模块方式调用::

    python -m origin_bridge info
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from typing import List, Optional

from .bridge import OriginBridge, OriginError, PLOT_TYPES
from .language import (
    LanguageError,
    close_origin,
    describe as describe_language,
    launch_origin,
    list_installations,
    list_processes,
    switch_language,
)


# --------------------------------------------------------------------------- #
# 输出辅助
# --------------------------------------------------------------------------- #

def _setup_console() -> None:
    """尽量把标准输出切到 UTF-8，避免 Windows 控制台中文乱码。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _ok(msg: str) -> None:
    print("  ✓ %s" % msg)


def _info(msg: str) -> None:
    print("  · %s" % msg)


def _fail(msg: str) -> None:
    print("  ✗ %s" % msg, file=sys.stderr)


def _header(title: str) -> None:
    print()
    print("=" * 60)
    print("  %s" % title)
    print("=" * 60)


# --------------------------------------------------------------------------- #
# 数据文件读写
# --------------------------------------------------------------------------- #

def read_data_file(path: str) -> List[List[Optional[float]]]:
    """
    读取 CSV / TSV / 纯文本数据文件。

    自动探测分隔符（逗号、制表符、分号），跳过空行。
    无法解析为数值的单元格记为 ``None``。

    :return: 二维列表（按行组织）。
    """
    if not os.path.isfile(path):
        raise FileNotFoundError("文件不存在：%s" % path)

    with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
        sample = fh.read(8192)
        fh.seek(0)

        delimiter = ","
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            delimiter = dialect.delimiter
        except csv.Error:
            # 探测失败时按出现频次挑选
            counts = {d: sample.count(d) for d in (",", "\t", ";")}
            delimiter = max(counts, key=counts.get) if any(counts.values()) else ","

        reader = csv.reader(fh, delimiter=delimiter)
        rows: List[List[Optional[float]]] = []
        for raw in reader:
            if not raw or all(not str(c).strip() for c in raw):
                continue
            row: List[Optional[float]] = []
            for cell in raw:
                text = str(cell).strip()
                if not text:
                    row.append(None)
                    continue
                try:
                    row.append(float(text))
                except ValueError:
                    row.append(None)
            rows.append(row)

    if not rows:
        raise ValueError("文件中没有可用数据：%s" % path)
    return rows


def _index_to_letters(idx: int) -> str:
    """0-based 下标 → 列字母（0→A, 25→Z, 26→AA）。"""
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


# --------------------------------------------------------------------------- #
# 各子命令实现
# --------------------------------------------------------------------------- #

def cmd_info(args) -> int:
    """显示 Origin 安装信息、运行状态与界面语言。"""
    _header("Origin 环境信息")

    installs = list_installations()
    if not installs:
        _fail("注册表中未找到任何 Origin 安装记录")
        _info("请确认 Origin 已安装，并至少手动运行过一次")
        return 1

    print("  注册表中共有 %d 个安装记录：" % len(installs))
    for st in installs:
        print("    · %s" % st)

    print()
    pids = list_processes()
    if pids:
        _ok("Origin 正在运行（PID: %s）" % ", ".join(str(p) for p in pids))
    else:
        _info("Origin 当前未运行")

    print()
    print("  当前语言：%s" % describe_language())

    if not args.quiet:
        print()
        try:
            with OriginBridge() as ob:
                _ok("COM 连接成功")
                print("  实例名称：%s" % ob.info().name)
                print("  当前工作表：%s" % (ob.active_book() or "(无)"))
        except OriginError as exc:
            _fail("COM 连接失败：%s" % exc)
            return 1
    return 0


def cmd_lang(args) -> int:
    """查看或切换界面语言。"""
    if args.versions:
        _header("已安装的 Origin 版本")
        for st in list_installations():
            print("  · %s" % st)
        return 0

    if not args.value:
        _header("Origin 界面语言")
        print("  %s" % describe_language())
        print()
        _info("切换用法：origin-cli lang zh   /   origin-cli lang en")
        return 0

    _header("切换 Origin 界面语言")

    if args.close or args.restart:
        pids = list_processes()
        if pids:
            _info("正在关闭 %d 个 Origin 进程..." % len(pids))
            _info("（未保存的工作会丢失，可用 --no-close 跳过此步）")
            close_origin()
            _ok("已关闭 Origin")

    try:
        code = switch_language(
            args.value,
            version_key=args.version_key,
            restart=False,
            close_first=False,       # 上面已按需处理
        )
    except LanguageError as exc:
        _fail(str(exc))
        return 1

    _ok("已写入 Language = %s" % code)
    print()
    print("  当前状态：%s" % describe_language())

    if args.restart:
        _info("正在启动 Origin...")
        exe = launch_origin(wait=6.0)
        if exe:
            _ok("已启动：%s" % exe)
        else:
            _fail("未找到 Origin 可执行文件，请手动启动")
    else:
        print()
        _info("请手动重启 Origin 使设置生效")

    if not args.quiet:
        print()
        print("  备注：Origin 的中文语言包并非 100% 覆盖，")
        print("        部分高级分析工具与第三方 App 的菜单会保持英文，属正常现象。")
    return 0


def cmd_import(args) -> int:
    """把数据文件导入 Origin 工作表。"""
    _header("导入数据 → Origin")

    try:
        rows = read_data_file(args.path)
    except (FileNotFoundError, ValueError) as exc:
        _fail(str(exc))
        return 1

    n_cols = max(len(r) for r in rows)
    n_rows = len(rows)
    _info("已解析 %s：%d 行 × %d 列" % (os.path.basename(args.path), n_rows, n_cols))

    try:
        with OriginBridge() as ob:
            book = ob.new_workbook(args.book or "ImportedData")
            _ok("已创建工作簿：%s" % book)

            for c in range(n_cols):
                letter = _index_to_letters(c)
                values = [row[c] if c < len(row) else None for row in rows]
                ob.set_column(letter, values, start_row=args.start_row)
            _ok("已写入 %d 列数据" % n_cols)

            if args.save:
                if ob.save(args.save):
                    _ok("已保存项目：%s" % os.path.abspath(args.save))
                else:
                    _fail("项目保存失败：%s" % args.save)
                    return 1
    except OriginError as exc:
        _fail("操作失败：%s" % exc)
        return 1

    print()
    _ok("导入完成")
    return 0


def cmd_export(args) -> int:
    """把当前工作表导出为文本文件。"""
    _header("导出 Origin 数据 → 文件")
    try:
        with OriginBridge() as ob:
            book = ob.active_book()
            _info("来源工作表：%s" % (book or "(当前)"))
            if ob.export_ascii(args.path, book=args.book):
                _ok("已导出：%s" % os.path.abspath(args.path))
            else:
                _fail("导出失败")
                return 1
    except OriginError as exc:
        _fail("操作失败：%s" % exc)
        return 1
    return 0


def cmd_stats(args) -> int:
    """显示指定列的统计量。"""
    _header("列统计量")
    try:
        with OriginBridge() as ob:
            for col in args.columns:
                st = ob.stats(col)
                print("  %s" % st)
                print()
    except OriginError as exc:
        _fail("操作失败：%s" % exc)
        return 1
    return 0


def cmd_plot(args) -> int:
    """用两列数据绘图。"""
    _header("绘制图形")
    try:
        with OriginBridge() as ob:
            ob.plot(args.x, args.y, kind=args.type)
            _ok("已创建 %s 图（X=%s, Y=%s）" % (args.type, args.x, args.y))
            if args.save:
                if ob.save(args.save):
                    _ok("已保存项目：%s" % os.path.abspath(args.save))
                else:
                    _fail("项目保存失败")
                    return 1
    except (OriginError, ValueError) as exc:
        _fail("操作失败：%s" % exc)
        return 1
    return 0


def cmd_save(args) -> int:
    """保存当前项目。"""
    _header("保存项目")
    try:
        with OriginBridge() as ob:
            if ob.save(args.path):
                _ok("已保存：%s" % os.path.abspath(args.path))
            else:
                _fail("保存失败（文件未生成）")
                return 1
    except OriginError as exc:
        _fail("操作失败：%s" % exc)
        return 1
    return 0


# --------------------------------------------------------------------------- #
# 参数解析
# --------------------------------------------------------------------------- #

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="origin-cli",
        description="用 Python 通过 COM 自动化控制 OriginLab Origin",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例:
  origin-cli info                          查看环境与语言状态
  origin-cli lang zh                       切换到中文界面
  origin-cli lang en --restart             切换到英文并重启
  origin-cli import data.csv --book MyData 导入数据文件
  origin-cli stats A B                     查看 A、B 两列统计量
  origin-cli plot A B --type scatter       绘制散点图
  origin-cli save result.opju              保存项目
""",
    )
    parser.add_argument("-V", "--version", action="version",
                        version="origin-bridge %s" % _get_version())
    sub = parser.add_subparsers(dest="command", metavar="<命令>")

    # info
    p = sub.add_parser("info", help="查看 Origin 安装与运行状态")
    p.add_argument("-q", "--quiet", action="store_true", help="跳过 COM 连接测试")
    p.set_defaults(func=cmd_info)

    # lang
    p = sub.add_parser("lang", help="查看或切换界面语言")
    p.add_argument("value", nargs="?", help="目标语言：zh/中文/C 或 en/英文/E")
    p.add_argument("--versions", action="store_true", help="列出所有已安装版本")
    p.add_argument("--version-key", help="指定注册表版本键（如 'Origin 10.2b'）")
    p.add_argument("--restart", action="store_true", help="切换后自动重启 Origin")
    p.add_argument("--no-close", dest="close", action="store_false",
                   help="不要关闭正在运行的 Origin")
    p.add_argument("-q", "--quiet", action="store_true", help="精简输出")
    p.set_defaults(func=cmd_lang, close=True)

    # import
    p = sub.add_parser("import", help="导入数据文件到 Origin")
    p.add_argument("path", help="数据文件路径（CSV/TSV/TXT）")
    p.add_argument("--book", help="目标工作簿名称")
    p.add_argument("--start-row", type=int, default=1, help="起始行号（默认 1）")
    p.add_argument("--save", help="导入后保存项目到该路径")
    p.set_defaults(func=cmd_import)

    # export
    p = sub.add_parser("export", help="导出当前工作表数据到文件")
    p.add_argument("path", help="输出文件路径（.txt）")
    p.add_argument("--book", help="来源工作表名称")
    p.set_defaults(func=cmd_export)

    # stats
    p = sub.add_parser("stats", help="查看列统计量")
    p.add_argument("columns", nargs="+", help="列标识，如 A B 或 1 2")
    p.set_defaults(func=cmd_stats)

    # plot
    p = sub.add_parser("plot", help="用两列数据绘图")
    p.add_argument("x", help="X 列")
    p.add_argument("y", help="Y 列")
    p.add_argument("--type", default="line", choices=sorted(PLOT_TYPES),
                   help="图形类型（默认 line）")
    p.add_argument("--save", help="绘图后保存项目到该路径")
    p.set_defaults(func=cmd_plot)

    # save
    p = sub.add_parser("save", help="保存当前项目")
    p.add_argument("path", help="目标路径（.opju）")
    p.set_defaults(func=cmd_save)

    return parser


def _get_version() -> str:
    from . import __version__
    return __version__


def main(argv: Optional[List[str]] = None) -> int:
    """CLI 入口。返回进程退出码。"""
    _setup_console()
    parser = build_parser()
    args = parser.parse_args(argv)

    if not getattr(args, "command", None):
        parser.print_help()
        return 0

    try:
        return args.func(args)
    except KeyboardInterrupt:
        print()
        _fail("已中断")
        return 130
    except OriginError as exc:
        _fail(str(exc))
        return 1


if __name__ == "__main__":
    sys.exit(main())
