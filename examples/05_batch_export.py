# -*- coding: utf-8 -*-
"""
示例 05 —— 批量处理多个工作簿
==============================

演示如何遍历 Origin 项目中的所有工作表，把数据批量导出为文本文件，
并汇总一份统计报告。

适用场景：把一次实验的多个数据表批量导出做进一步分析。
"""

import os
import sys

# 允许在未安装本包的情况下直接运行本示例
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_bridge import OriginBridge  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "batch_output")


def safe_name(text):
    """把工作表名转成合法的文件名。"""
    for ch in '\\/:*?"<>|':
        text = text.replace(ch, "_")
    return text.strip() or "unnamed"


def main():
    os.makedirs(OUT_DIR, exist_ok=True)

    with OriginBridge() as ob:
        # --- 1. 找出所有工作表 ---
        books = ob.list_books()
        if not books:
            print("当前项目中没有找到任何工作表。")
            print("提示：可以先运行 examples/01_quickstart.py 创建一些数据。")
            return

        print("=" * 64)
        print("  发现 %d 个工作表" % len(books))
        print("=" * 64)
        for name in books:
            print("  · %s" % name)
        print()

        # --- 2. 逐个导出 ---
        report = []
        for name in books:
            try:
                data = ob.get_data(name)
            except Exception as exc:
                print("  跳过 %s（读取失败：%s）" % (name, exc))
                continue

            if not data:
                print("  跳过 %s（无数据）" % name)
                continue

            n_rows = len(data)
            n_cols = len(data[0])

            # 导出文本
            txt_path = os.path.join(OUT_DIR, "%s.txt" % safe_name(name))
            with open(txt_path, "w", encoding="utf-8", newline="") as fh:
                fh.write("# 来源工作表: %s\n" % name)
                fh.write("# 尺寸: %d 行 x %d 列\n" % (n_rows, n_cols))
                for row in data:
                    fh.write("\t".join(
                        "" if v is None else ("%.10g" % v) for v in row
                    ) + "\n")

            # 汇总每列统计
            col_stats = []
            for c in range(n_cols):
                values = [row[c] for row in data if c < len(row) and row[c] is not None]
                if values:
                    col_stats.append({
                        "col": c,
                        "n": len(values),
                        "mean": sum(values) / len(values),
                        "min": min(values),
                        "max": max(values),
                    })

            report.append({
                "name": name,
                "rows": n_rows,
                "cols": n_cols,
                "path": txt_path,
                "stats": col_stats,
            })
            print("  已导出 %-20s → %s（%d 行 × %d 列）"
                  % (name, os.path.basename(txt_path), n_rows, n_cols))

        # --- 3. 写出汇总报告 ---
        report_path = os.path.join(OUT_DIR, "_summary.txt")
        with open(report_path, "w", encoding="utf-8", newline="") as fh:
            fh.write("Origin 数据批量导出汇总\n")
            fh.write("=" * 60 + "\n\n")
            for item in report:
                fh.write("工作表: %s\n" % item["name"])
                fh.write("  尺寸: %d 行 x %d 列\n" % (item["rows"], item["cols"]))
                fh.write("  文件: %s\n" % item["path"])
                for st in item["stats"]:
                    fh.write("    第%d列: n=%d  均值=%.6g  范围=[%.6g, %.6g]\n"
                             % (st["col"] + 1, st["n"], st["mean"],
                                st["min"], st["max"]))
                fh.write("\n")

        print()
        print("=" * 64)
        print("  共处理 %d 个工作表" % len(report))
        print("  输出目录：%s" % OUT_DIR)
        print("  汇总报告：%s" % report_path)
        print("=" * 64)


if __name__ == "__main__":
    main()
