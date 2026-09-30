# -*- coding: utf-8 -*-
"""
origin_bridge.language — Origin 界面语言切换

Origin 的界面语言**没有图形化开关**，而是由注册表中的一个字符串值决定：

    路径：HKEY_CURRENT_USER\\Software\\OriginLab\\<版本键>
    值名：Language
    取值：``C`` = 中文（简体）    ``E`` = 英文

其中「版本键」是 Origin 的**内部版本号**，并非市场名。例如：

    市场名        内部版本键
    ----------    ----------------
    Origin 2025b  Origin 10.2b
    Origin 2024b  Origin 10.1b
    Origin 2023b  Origin 10.0b

所以 2025b 对应的键名看起来像旧版本号，这是正常的，不要误改成 ``Origin 2025b``。

本模块直接使用 ``winreg`` 操作注册表：

* ``HKCU`` 分支**无需管理员权限**；
* 不依赖外部 ``reg.exe``（某些受管环境下该程序被安全策略禁用）。

注意：修改后必须**重启 Origin** 才能生效。
"""

from __future__ import annotations

import os
import subprocess
import time
import winreg
from dataclasses import dataclass
from typing import Dict, List, Optional

from .bridge import (
    ORIGIN_EXES,
    REG_ROOT,
    OriginError,
    find_origin_exe,
    iter_origin_reg_keys,
)

#: 语言代码 → 显示名
LANGUAGE_NAMES: Dict[str, str] = {
    "C": "中文（简体）",
    "E": "英文",
}

#: 显示名 / 别名 → 语言代码
_LANGUAGE_ALIASES: Dict[str, str] = {
    "c": "C", "cn": "C", "zh": "C", "chs": "C",
    "chinese": "C", "中文": "C", "简体": "C", "简体中文": "C",
    "e": "E", "en": "E", "eng": "E",
    "english": "E", "英文": "E", "英语": "E",
}

#: Language 值在注册表中的名称
LANGUAGE_VALUE_NAME = "Language"


class LanguageError(OriginError):
    """语言切换相关错误。"""


@dataclass
class LanguageState:
    """某个 Origin 安装的语言状态。"""

    version_key: str              # 内部版本键，如 "Origin 10.2b"
    language: Optional[str]       # "C" / "E" / None（未设置）
    install_path: Optional[str] = None
    build: Optional[str] = None

    @property
    def language_name(self) -> str:
        if self.language is None:
            return "未设置（跟随默认）"
        return LANGUAGE_NAMES.get(self.language, "未知(%s)" % self.language)

    def __str__(self) -> str:
        bits = ["%s: %s" % (self.version_key, self.language_name)]
        if self.build:
            bits.append("build=%s" % self.build)
        if self.install_path:
            bits.append(self.install_path)
        return "  ".join(bits)


# --------------------------------------------------------------------------- #
# 内部辅助
# --------------------------------------------------------------------------- #

def _open_version_key(version_key: str, writable: bool = False):
    """打开指定版本键，返回 winreg 句柄（需自行关闭）。"""
    access = winreg.KEY_READ | (winreg.KEY_SET_VALUE if writable else 0)
    path = "%s\\%s" % (REG_ROOT, version_key)
    return winreg.OpenKey(winreg.HKEY_CURRENT_USER, path, 0, access)


def normalize_language(language: str) -> str:
    """
    把各种写法（``zh`` / ``中文`` / ``en`` / ``english``…）统一成 ``C`` / ``E``。

    :raises LanguageError: 无法识别。
    """
    if language is None:
        raise LanguageError("语言不能为空")
    key = str(language).strip().lower()
    if key in _LANGUAGE_ALIASES:
        return _LANGUAGE_ALIASES[key]
    if key.upper() in ("C", "E"):
        return key.upper()
    raise LanguageError(
        "无法识别的语言 %r，可用取值：C/中文/zh 或 E/英文/en" % language
    )


def _probe_version_key(version_key: str) -> Optional[LanguageState]:
    """读取某个版本键的安装路径、Build 与 Language 值。"""
    try:
        with _open_version_key(version_key) as k:
            try:
                path, _ = winreg.QueryValueEx(k, "Executable Path")
                path = os.path.normpath(path) if path else None
            except FileNotFoundError:
                path = None
            try:
                build, _ = winreg.QueryValueEx(k, "Build")
                build = str(build)
            except FileNotFoundError:
                build = None
            try:
                lang, _ = winreg.QueryValueEx(k, LANGUAGE_VALUE_NAME)
                lang = str(lang).strip().upper() or None
            except FileNotFoundError:
                lang = None
    except OSError:
        return None
    return LanguageState(version_key=version_key, language=lang,
                         install_path=path, build=build)


# --------------------------------------------------------------------------- #
# 查询
# --------------------------------------------------------------------------- #

def list_installations() -> List[LanguageState]:
    """
    列出注册表中记录的所有 Origin 安装及其语言状态。

    :return: :class:`LanguageState` 列表（按版本键降序，较新版本在前）。
    """
    states: List[LanguageState] = []
    for key in iter_origin_reg_keys():
        if not key.lower().startswith("origin"):
            continue
        state = _probe_version_key(key)
        if state is not None:
            states.append(state)
    # 版本键名降序排列（"Origin 10.2b" 排在 "Origin 10.1b" 之前）
    states.sort(key=lambda s: s.version_key, reverse=True)
    return states


def resolve_version_key(version_key: Optional[str] = None) -> str:
    """
    确定要操作的版本键。

    :param version_key: 显式指定的键名；``None`` 时自动选择。
    :raises LanguageError: 未找到任何 Origin 安装。
    """
    if version_key:
        try:
            with _open_version_key(version_key):
                return version_key
        except OSError as exc:
            available = [s.version_key for s in list_installations()]
            raise LanguageError(
                "注册表中不存在版本键 %r。可用：%s"
                % (version_key, ", ".join(available) or "（无）")
            ) from exc

    states = list_installations()
    if not states:
        raise LanguageError(
            "未在注册表 %s 下找到任何 Origin 安装记录。\n"
            "请确认 Origin 已安装并至少运行过一次。" % REG_ROOT
        )
    # 优先选择真正存在可执行文件的安装
    for st in states:
        if st.install_path and any(
            os.path.isfile(os.path.join(st.install_path, exe))
            for exe in ORIGIN_EXES
        ):
            return st.version_key
    return states[0].version_key


def get_language(version_key: Optional[str] = None) -> Optional[str]:
    """
    读取当前界面语言。

    :return: ``"C"``（中文）或 ``"E"``（英文）；未设置时返回 ``None``。
    """
    key = resolve_version_key(version_key)
    state = _probe_version_key(key)
    return state.language if state else None


def describe(version_key: Optional[str] = None) -> str:
    """返回当前语言的人类可读描述。"""
    key = resolve_version_key(version_key)
    state = _probe_version_key(key)
    if state is None:
        return "无法读取 %s 的语言设置" % key
    return "%s  →  %s" % (state.version_key, state.language_name)


# --------------------------------------------------------------------------- #
# 修改
# --------------------------------------------------------------------------- #

def set_language(language: str, version_key: Optional[str] = None) -> str:
    """
    写入界面语言设置（**不负责重启 Origin**）。

    :param language: ``C`` / ``中文`` / ``zh`` 或 ``E`` / ``英文`` / ``en``。
    :param version_key: 目标版本键，``None`` 表示自动选择。
    :return: 实际写入的语言代码（``"C"`` 或 ``"E"``）。
    :raises LanguageError: 写入失败。

    提示：``HKCU`` 分支无需管理员权限；若仍报权限错误，
    通常是安全软件拦截了注册表写入。
    """
    code = normalize_language(language)
    key = resolve_version_key(version_key)
    path = "%s\\%s" % (REG_ROOT, key)
    try:
        # CreateKeyEx：键不存在时自动创建，存在时直接打开
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0,
                                winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, LANGUAGE_VALUE_NAME, 0, winreg.REG_SZ, code)
    except OSError as exc:
        raise LanguageError(
            "写入注册表失败（HKCU\\%s）：%s\n"
            "请检查是否有安全软件拦截，或以管理员身份重试。" % (path, exc)
        ) from exc

    # 立即回读校验，避免"看似成功实则未写入"
    actual = get_language(key)
    if actual != code:
        raise LanguageError(
            "注册表写入后校验失败：期望 %r，实际读到 %r" % (code, actual)
        )
    return code


def list_processes() -> List[int]:
    """返回正在运行的 Origin 进程 PID 列表。"""
    pids: List[int] = []
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq Origin64.exe", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
            encoding="gbk", errors="replace",
        )
        for line in (out.stdout or "").splitlines():
            parts = [p.strip().strip('"') for p in line.split('","')]
            if len(parts) > 1 and parts[1].isdigit():
                pids.append(int(parts[1]))
    except Exception:
        pass
    return pids


def close_origin(reason: str = "语言切换前需要关闭 Origin") -> int:
    """
    关闭所有 Origin 进程。

    .. warning::
       **未保存的工作会丢失。** 调用前请确保 Origin 中没有未保存的内容。

    :return: 实际结束的进程数。
    """
    pids = list_processes()
    if not pids:
        return 0
    try:
        subprocess.run(
            ["taskkill", "/IM", "Origin64.exe", "/F"],
            capture_output=True, text=True, timeout=20,
            encoding="gbk", errors="replace",
        )
    except Exception as exc:
        raise LanguageError("关闭 Origin 失败：%r" % (exc,)) from exc
    time.sleep(1.5)
    return len(pids)


def launch_origin(wait: float = 0.0) -> Optional[str]:
    """
    启动 Origin。

    通过 ``explorer.exe`` 打开主程序，使新进程归属桌面 shell
    （直接 ``subprocess.Popen`` 启动的 GUI 进程在部分受管环境中会被回收）。

    :return: 启动的可执行文件路径；未找到时返回 ``None``。
    """
    exe = find_origin_exe()
    if not exe:
        return None
    try:
        subprocess.Popen(["explorer.exe", exe],
                         creationflags=getattr(subprocess, "DETACHED_PROCESS", 0))
    except Exception as exc:
        raise LanguageError("启动 Origin 失败：%r" % (exc,)) from exc
    if wait:
        time.sleep(wait)
    return exe


def switch_language(language: str, version_key: Optional[str] = None,
                    restart: bool = False, close_first: bool = True) -> str:
    """
    一步完成语言切换（推荐入口）。

    :param language: 目标语言，见 :func:`set_language`。
    :param version_key: 目标版本键，``None`` 表示自动选择。
    :param restart: 写入后是否自动重启 Origin。
    :param close_first: 是否先关闭正在运行的 Origin。
        **仅当 Origin 未运行或你确认工作已保存时才可置为 True 并配合 restart=False。**
        ``HKCU`` 写入本身不需要关闭程序，但改动要重启后才生效；
        留着旧进程会导致"改了没反应"的错觉。
    :return: 实际写入的语言代码。

    .. warning::
        ``close_first=True`` 会强杀 Origin 进程，**未保存的工作会丢失**。
    """
    code = set_language(language, version_key)

    if close_first and list_processes():
        close_origin()

    if restart:
        launch_origin(wait=6.0)

    return code
