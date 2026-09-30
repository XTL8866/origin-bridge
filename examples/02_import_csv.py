# -*- coding: utf-8 -*-
"""
示例 02 —— 导入 CSV 数据并分析
================================

演示如何把外部数据文件（CSV/TSV/TXT）导入 Origin，并做自动化分析。

生成一份示例 CSV，导入 Origin，计算统计量，再导出结果。
"""

import csv
import math
import os
import random
import sys

# 允许在未安装本包的情况下直接运行本示例
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_bridge import OriginBridge, index_to_letters  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))


def make_sample_csv(path, n=200, seed=42):
    """生成一份带噪声的正弦波示例数据。"""
    rng = random.Random(seed)
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["时间(s)", "电压(V)", "电流(A)"])
        for i in range(n):
            t = i * 0.05
            v = 12.0 + 3.0 * math.sin(2 * math.pi * t) + rng.gauss(0, 0.15)
            a = 2.5 + 0.4 * math.sin(2 * math.pi * t + 0.6) + rng.gauss(0, 0.03)
            writer.writerow(["%.3f" % t, "%.4f" % v, "%.4f" % a])
    return path


def read_csv(path):
    """读回 CSV，跳过表头，返回按列组织的列表。"""
    rows = []
    with open(path, "r", encoding="utf-8") as fh:
        reader = csv.reader(fh)
        next(reader, None)               # 跳过表头
        for line in reader:
            if line:
                rows.append([float(c) for c in line])
    if not rows:
        return []
    n_cols = len(rows[0])
    return [[r[c] for r in rows] for c in range(n_cols)]


def main():
    csv_path = os.path.join(HERE, "sample_data.csv")
    make_sample_csv(csv_path)
    print("已生成示例数据：%s" % csv_path)

    columns = read_csv(csv_path)
    print("共 %d 列，每列 %d 个数据点" % (len(columns), len(columns[0])))
    print()

    with OriginBridge() as ob:
        book = ob.new_workbook("SignalAnalysis")
        print("已创建工作簿：%s" % book)

        # 逐列写入（列字母：A、B、C…）
        labels = ["时间 (s)", "电压 (V)", "电流 (A)"]
        for idx, values in enumerate(columns):
            letter = index_to_letters(idx)
            ob.set_column(letter, values)
            ob.set_column_label(letter, labels[idx])
        print("已导入 %d 列数据" % len(columns))

        # 统计每一列
        print()
        print("-" * 62)
        print("  列统计摘要")
        print("-" * 62)
        for idx in range(len(columns)):
            letter = index_to_letters(idx)
            st = ob.stats(letter)
            print("  %-10s  n=%-5s 均值=%-10.4f 标准差=%-9.4f 范围=[%.4f, %.4f]"
                  % (labels[idx], _fmt(st.count), st.mean or 0,
                     st.sd or 0, st.minimum or 0, st.maximum or 0))
        print("-" * 62)

        # 导出回文本（验证往返一致）
        out_txt = os.path.join(HERE, "output_exported.txt")
        if ob.export_ascii(out_txt):
            print()
            print("已导出数据：%s" % out_txt)

        out_opju = os.path.join(HERE, "output_analysis.opju")
        if ob.save(out_opju):
            print("已保存项目：%s" % out_opju)


def _fmt(value):
    """把统计量格式化成整数串。"""
    if value is None:
        return "-"
    return "%d" % int(round(value))


if __name__ == "__main__":
    main()
