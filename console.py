#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
端末表示の共通ヘルパー。menu.py と kintone_runner.py で共用する。

- 色は端末（TTY）のときだけ付ける。環境変数 NO_COLOR があれば付けない。
- 記号は cp932 でも表示できるものだけ使う（○ × △ ・ →）。
"""

import os
import sys
import unicodedata
from typing import Optional

RESET = "\033[0m"
BOLD = "\033[1m"
DIM = "\033[2m"
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
CYAN = "\033[96m"

_color_enabled: Optional[bool] = None


def _enable_ansi_windows() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)
        mode = ctypes.c_uint32()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        pass


def supports_color() -> bool:
    """標準出力が端末で、NO_COLOR が無いときだけ True。結果はキャッシュする。"""
    global _color_enabled
    if _color_enabled is None:
        if os.environ.get("NO_COLOR") or not sys.stdout.isatty():
            _color_enabled = False
        else:
            _enable_ansi_windows()
            _color_enabled = True
    return _color_enabled


def c(text: str, *styles: str) -> str:
    """色・装飾を付ける。端末でなければそのまま返す。"""
    if not styles or not supports_color():
        return text
    return "".join(styles) + text + RESET


def display_width(text: str) -> int:
    """全角を 2、半角を 1 として表示幅を数える。"""
    width = 0
    for ch in text:
        width += 2 if unicodedata.east_asian_width(ch) in ("F", "W", "A") else 1
    return width


def pad(text: str, width: int) -> str:
    """表示幅が width になるまで右に空白を足す（全角対応）。"""
    return text + " " * max(0, width - display_width(text))


def hr(width: int = 60, char: str = "-") -> str:
    return char * width


# ---- 1 行メッセージ -------------------------------------------------------
# 記号: ○ 成功 / × 失敗 / △ 注意 / ・ 情報 / → 出力先

def ok(message: str, indent: str = "  ") -> None:
    print(c(f"{indent}○ {message}", GREEN))


def fail(message: str, indent: str = "  ") -> None:
    print(c(f"{indent}× {message}", BOLD, RED))


def warn(message: str, indent: str = "  ") -> None:
    print(c(f"{indent}△ {message}", YELLOW))


def note(message: str, indent: str = "  ") -> None:
    print(c(f"{indent}・ {message}", DIM))


def arrow(message: str, indent: str = "    ") -> None:
    """出力先など「→ ...」の行。"""
    print(c(f"{indent}→ {message}", CYAN))


def heading(text: str, width: int = 60) -> None:
    """区切り線付きの見出し。"""
    print()
    print(c(f"  {text}", BOLD, CYAN))
    print(c("  " + hr(width - 2), DIM))
