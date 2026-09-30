# -*- coding: utf-8 -*-
"""
示例 01 —— 快速开始
===================

演示最基本的完整流程：连接 → 建表 → 写数据 → 读回 → 统计 → 保存。

运行前请确保 Origin 已启动（或允许由 COM 自动启动）。
"""

import os
import sys

# 允许在未安装本包的情况下直接运行本示例
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_bridge import OriginBridge  # noqa: E402


def main():
    # 用 with 语句管理连接，退出时自动断开引用（不会关闭 Origin 进程）
    with OriginBridge() as ob:
        # --- 1. 看看连上的是哪个版本 ---
        info = ob.info()
        print("已连接：%s" % info)
        print()

        # --- 2. 新建工作簿 ---
        book = ob.new_workbook("QuickStart")
        print("工作簿：%s" % book)

        # --- 3. 写入数据 ---
        # 列可以用字母 "A"/"B"，也可以用序号 1/2
        x = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
        y = [v ** 2 for v in x]          # y = x²

        ob.set_column("A", x)
        ob.set_column("B", y)
        ob.set_column_label("A", "X 值")
        ob.set_column_label("B", "X 的平方")
        print("已写入 A、B 两列，各 %d 个数据" % len(x))

        # --- 4. 读回验证 ---
        print()
        print("读回数据：")
        print("  A 列 = %s" % ob.get_column("A", rows=10))
        print("  B 列 = %s" % ob.get_column("B", rows=10))

        # 也可以一次拿回整张表（按行组织）
        data = ob.get_data()
        print("  整表共 %d 行 × %d 列" % (len(data), len(data[0]) if data else 0))

        # --- 5. 统计 ---
        print()
        for col in ("A", "B"):
            print("  %s" % ob.stats(col))

        # --- 6. 保存项目 ---
        import os
        out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           "output_quickstart.opju")
        if ob.save(out):
            print()
            print("已保存项目：%s" % out)
        else:
            print("保存失败！")


if __name__ == "__main__":
    main()
