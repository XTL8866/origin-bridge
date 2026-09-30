# -*- coding: utf-8 -*-
"""
示例 03 —— 绘图
================

演示用两列数据绘制不同类型的图形，并导出图片。

Origin 的绘图通过 LabTalk 的 ``plotxy`` 命令完成，
本库把它封装成 :meth:`OriginBridge.plot`。
"""

import math
import os
import sys

# 允许在未安装本包的情况下直接运行本示例
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_bridge import OriginBridge  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))


def main():
    with OriginBridge() as ob:
        book = ob.new_workbook("PlotDemo")
        print("工作簿：%s" % book)

        # --- 准备数据 ---
        n = 50
        x = [i * 0.2 for i in range(n)]
        y_sin = [math.sin(v) for v in x]
        y_cos = [math.cos(v) for v in x]

        ob.set_column("A", x)
        ob.set_column("B", y_sin)
        ob.set_column("C", y_cos)
        ob.set_column_label("A", "角度 (rad)")
        ob.set_column_label("B", "sin(x)")
        ob.set_column_label("C", "cos(x)")
        print("已写入 %d 个数据点，A=角度, B=sin, C=cos" % n)
        print()

        # --- 绘制多种图形 ---
        # 支持的 kind：line / scatter / line_symbol / column / bar / area
        for kind in ("line", "scatter", "line_symbol"):
            ob.plot("A", "B", kind=kind)
            print("已绘制 %s 图（A vs B）" % kind)

        print()
        print("提示：若图片导出失败，多是因为 Origin 当前没有活动的图窗口。")
        print("      可先在 Origin 中手动激活某个图，再运行导出。")

        # --- 导出图片 ---
        for ext in ("png", "tif"):
            out = os.path.join(HERE, "output_plot.%s" % ext)
            try:
                if ob.export_image(out):
                    print("已导出图片：%s" % out)
                else:
                    print("未生成图片（可能没有活动图窗口）：%s" % out)
            except Exception as exc:
                print("导出 %s 时出错：%s" % (ext, exc))

        # --- 保存项目 ---
        out_opju = os.path.join(HERE, "output_plot.opju")
        if ob.save(out_opju):
            print()
            print("已保存项目：%s" % out_opju)


if __name__ == "__main__":
    main()
