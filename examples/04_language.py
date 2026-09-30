# -*- coding: utf-8 -*-
"""
示例 04 —— 界面语言切换
========================

Origin **没有**图形化的语言切换菜单，界面语言由注册表中的一个值控制：

    HKEY_CURRENT_USER\\Software\\OriginLab\\<版本键>   →   Language = "C" / "E"

本示例演示如何查看和切换语言。

.. warning::
    切换语言后必须**重启 Origin** 才会生效，且重启会导致未保存的工作丢失。
    本示例默认**只查看不修改**，需要真正切换时请把 ``DO_SWITCH`` 改为 True 并
    在运行时确认。
"""

import os
import sys

# 允许在未安装本包的情况下直接运行本示例
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from origin_bridge import (  # noqa: E402
    LANGUAGE_NAMES,
    OriginBridge,
    find_origin_exe,
    list_installations,
    list_processes,
    describe_language,
    get_language,
    switch_language,
)

# 安全开关：保持 False 则只查看状态，不会修改注册表
DO_SWITCH = False

# 想切换到的目标语言（"C"/"中文"/"zh" 或 "E"/"英文"/"en"）
TARGET_LANGUAGE = "中文"


def main():
    # ------------------------------------------------------------------ #
    # 1. 列出所有安装
    # ------------------------------------------------------------------ #
    print("=" * 62)
    print("  已安装的 Origin 版本")
    print("=" * 62)
    installs = list_installations()
    if not installs:
        print("  未找到任何 Origin 安装记录。")
        print("  请确认 Origin 已安装，并至少手动运行过一次。")
        return

    for st in installs:
        print("  · %s" % st)

    # ------------------------------------------------------------------ #
    # 2. 当前状态
    # ------------------------------------------------------------------ #
    print()
    print("=" * 62)
    print("  当前状态")
    print("=" * 62)
    print("  界面语言：%s" % describe_language())

    lang = get_language()
    print("  语言代码：%s" % (lang or "（未设置）"))
    if lang:
        print("  含义　　：%s" % LANGUAGE_NAMES.get(lang, "未知"))

    exe = find_origin_exe()
    print("  主程序　：%s" % (exe or "（未找到）"))

    pids = list_processes()
    if pids:
        print("  运行状态：运行中（PID: %s）" % ", ".join(str(p) for p in pids))
    else:
        print("  运行状态：未运行")

    # ------------------------------------------------------------------ #
    # 3. 可选：实际切换
    # ------------------------------------------------------------------ #
    print()
    if not DO_SWITCH:
        print("=" * 62)
        print("  切换演示（当前为只读模式）")
        print("=" * 62)
        print("  想真正切换语言，请编辑本文件把 DO_SWITCH 改为 True。")
        print()
        print("  等价命令：")
        print("    origin-cli lang zh        # 切换到中文")
        print("    origin-cli lang en        # 切换到英文")
        print("    origin-cli lang zh --restart   # 切换并自动重启 Origin")
        return

    print("=" * 62)
    print("  切换语言 → %s" % TARGET_LANGUAGE)
    print("=" * 62)

    if pids:
        print("  ⚠ 检测到 %d 个 Origin 进程正在运行。" % len(pids))
        print("    切换需要重启 Origin，未保存的工作将丢失。")
        answer = input("    确认继续？输入 yes 继续：").strip().lower()
        if answer != "yes":
            print("  已取消。")
            return

    # close_first=True 会关闭 Origin；restart=True 会自动重新启动
    code = switch_language(TARGET_LANGUAGE, restart=True)
    print()
    print("  已写入 Language = %s" % code)
    print("  当前状态：%s" % describe_language())
    print()
    print("  Origin 已重启，界面语言应已生效。")


if __name__ == "__main__":
    main()
