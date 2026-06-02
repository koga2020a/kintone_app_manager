#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Kintone統合実行ツールの対話型メニュー。

python menu.py で起動し、番号選択で kintone_runner.py の各機能を実行します。
"""

import subprocess
import sys
from pathlib import Path
from typing import List, Optional, Set

try:
    import yaml
except ImportError:
    yaml = None  # type: ignore

SCRIPT_DIR = Path(__file__).resolve().parent
RUNNER = SCRIPT_DIR / "kintone_runner.py"
ENV_FILE = SCRIPT_DIR / ".kintone.env"

# コンソール色（Windows 10+ / ターミナル対応）
RESET = "\033[0m"
BOLD = "\033[1m"
CYAN = "\033[96m"
YELLOW = "\033[93m"
GREEN = "\033[92m"
DIM = "\033[2m"
MARK = f"{BOLD}{YELLOW}★{RESET}"

MENU_ITEMS = [
    {
        "label": "ユーザーとグループ情報の取得",
        "command": "users",
        "description": "全ユーザーと所属グループを Excel/CSV に出力",
        "output": "output/kintone_users_groups_[日時].xlsx",
    },
    {
        "label": "ACL情報をExcelに変換",
        "command": "acl",
        "description": "アプリの ACL を Excel レポートに変換",
        "output": "output/acl_report_[アプリID]_[日時].xlsx",
        "needs_app_id": True,
    },
    {
        "label": "アプリ設定一覧表の出力",
        "command": "summary",
        "description": "取得済みアプリ設定の全体一覧を Excel で出力",
        "output": "output/kintone_app_settings_summary_[日時].xlsx",
    },
    {
        "label": "通知設定をExcelに変換",
        "command": "notifications",
        "description": "一般・レコード・リマインダー通知を Excel に変換",
        "output": "output/[アプリID]_notifications.xlsx",
        "needs_app_id": True,
    },
    {
        "label": "プロセスワークフローをExcelに変換",
        "command": "process_workflow",
        "description": "プロセス管理設定を Excel に変換",
        "output": "output/[アプリID]_process_workflow.xlsx",
        "needs_app_id": True,
    },
    {
        "label": "グループ操作",
        "command": "group",
        "description": "一覧表示・ユーザー検索・追加・削除",
        "output": "コンソール出力",
        "is_group": True,
    },
    {
        "label": "すべての機能を順番に実行",
        "command": "all",
        "description": "users → app → acl → summary → notifications → process_workflow",
        "output": "上記すべての出力ファイル",
        "needs_app_filter": True,
    },
    {
        "label": "出力ファイル一覧の表示",
        "command": "outputs",
        "description": "生成される Excel/CSV の一覧と概要を表示",
        "output": "（表示のみ）",
    },
    {
        "label": "アプリ設定・JavaScript の一括取得",
        "command": "app",
        "highlight": True,
        "description": ".kintone.env の app_tokens から Kintone API 経由で全設定・JS を ./output へ保存",
        "detail_lines": [
            "取得: フォーム/ACL/通知/プロセス等(YAML・JSON) + カスタマイズ JS(javascript/)",
            "アプリID 未入力=全件 / 指定=1件  ※ トークンに閲覧・DL権限が必要",
        ],
        "output": "./output/{アプリID}_{アプリ名}_{日時}/",
        "needs_app_id": True,
    },
]

def enable_ansi_windows():
    """Windows コンソールで ANSI エスケープを有効化する。"""
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
    if not sys.stdout.isatty():
        return False
    if sys.platform == "win32":
        enable_ansi_windows()
    return True


def c(text: str, *styles: str) -> str:
    if not supports_color() or not styles:
        return text
    return "".join(styles) + text + RESET


def load_app_tokens_preview() -> List[str]:
    """.kintone.env から app_tokens のアプリ番号一覧を取得する。"""
    if not ENV_FILE.exists() or yaml is None:
        return []
    try:
        with open(ENV_FILE, encoding="utf-8") as f:
            config = yaml.safe_load(f) or {}
        tokens = config.get("app_tokens") or {}
        if not isinstance(tokens, dict):
            return []
        return [str(k) for k in tokens.keys()]
    except Exception:
        return []


GROUP_ACTIONS = [
    ("list", "グループ一覧を表示"),
    ("search", "ユーザーを検索"),
    ("add", "ユーザーをグループに追加"),
    ("remove", "ユーザーをグループから削除"),
]


def print_header():
    print()
    print("=" * 60)
    print("  Kintone統合実行ツール")
    print("=" * 60)
    if ENV_FILE.exists():
        print(f"  設定ファイル: {ENV_FILE.name}")
    else:
        print(f"  警告: 設定ファイル {ENV_FILE.name} が見つかりません")
    print(f"  実行スクリプト: {RUNNER.name}")
    print("=" * 60)
    print()


def print_menu_item(index: int, item: dict):
    num = f"{index:2}"
    if item.get("highlight"):
        print(c(f"  {num}. {MARK} {item['label']}", BOLD, CYAN))
        print(c(f"      {item['description']}", YELLOW))
        for line in item.get("detail_lines", []):
            print(c(f"      {line}", GREEN))
        app_ids = load_app_tokens_preview()
        if app_ids:
            print(c(f"      登録済み app_tokens: {', '.join(app_ids)}", BOLD, CYAN))
        elif ENV_FILE.exists():
            print(c("      登録済み app_tokens: （未設定）", DIM))
        else:
            print(c(f"      ※ 先に {ENV_FILE.name} を作成してください", YELLOW))
        print(c(f"      出力先: {item['output']}", CYAN))
    else:
        print(f"  {num}. {item['label']}")
        print(f"      {item['description']}")
        print(f"      出力: {item['output']}")
    print()


def print_menu():
    print(c("【実行メニュー】", BOLD))
    print()
    for i, item in enumerate(MENU_ITEMS, start=1):
        print_menu_item(i, item)
    print("   0. 終了")
    print()


def read_choice(prompt: str, valid: Set[str]) -> str:
    while True:
        choice = input(prompt).strip()
        if choice in valid:
            return choice
        print(f"  無効な入力です。次のいずれかを入力してください: {', '.join(sorted(valid, key=lambda x: (x != '0', x)))}")


def read_app_id(optional: bool = True) -> Optional[str]:
    hint = "（Enter で全アプリ）" if optional else "（必須）"
    while True:
        raw = input(f"  アプリID {hint}: ").strip()
        if not raw:
            return None if optional else ""
        if raw.isdigit():
            return raw
        print("  アプリID は数値で入力してください。")


def build_users_args() -> List[str]:
    print()
    print("  出力形式: 1=excel（既定）, 2=csv")
    fmt = input("  選択 [1]: ").strip()
    if fmt == "2":
        return ["users", "--format", "csv"]
    return ["users"]


def build_group_args() -> Optional[List[str]]:
    print()
    print("【グループ操作】")
    for i, (action, desc) in enumerate(GROUP_ACTIONS, start=1):
        print(f"  {i}. {desc} ({action})")
    print("  0. 戻る")
    print()
    choice = read_choice("  番号を選択: ", {str(i) for i in range(len(GROUP_ACTIONS) + 1)})
    if choice == "0":
        return None
    action, _ = GROUP_ACTIONS[int(choice) - 1]
    args = ["group", action]
    if action == "search":
        keyword = input("  検索キーワード: ").strip()
        if not keyword:
            print("  キーワードが未入力のためキャンセルしました。")
            return None
        args.append(keyword)
    elif action == "add":
        user = input("  ユーザーコード: ").strip()
        group = input("  グループ名またはコード: ").strip()
        if not user or not group:
            print("  入力が不足しているためキャンセルしました。")
            return None
        args.extend([user, group])
    elif action == "remove":
        user = input("  ユーザーコード: ").strip()
        if not user:
            print("  ユーザーコードが未入力のためキャンセルしました。")
            return None
        args.append(user)
    return args


def build_all_args() -> List[str]:
    args = ["all"]
    print()
    print("  アプリIDの絞り込み（任意）")
    print("  1. 全アプリ")
    print("  2. 指定したアプリIDのみ")
    print("  3. 指定したアプリIDを除外")
    filter_choice = input("  選択 [1]: ").strip() or "1"
    if filter_choice == "2":
        raw = input("  対象アプリID（スペース区切り）: ").strip()
        if raw:
            args.extend(["--id", *raw.split()])
    elif filter_choice == "3":
        raw = input("  除外するアプリID（スペース区切り）: ").strip()
        if raw:
            args.extend(["--not-id", *raw.split()])
    return args


def print_app_run_confirm(item: dict):
    """アプリ一括取得実行前に概要を表示する。"""
    print()
    print(c(f"  {MARK} {item['label']}", BOLD, CYAN))
    print(c(f"  {item['description']}", YELLOW))
    app_ids = load_app_tokens_preview()
    if app_ids:
        print(c(f"  対象 app_tokens: {', '.join(app_ids)}", CYAN))
    print()


def build_app_args(item: dict) -> List[str]:
    print_app_run_confirm(item)
    app_id = read_app_id(optional=True)
    args = ["app"]
    if app_id:
        args.extend(["--id", app_id])
    else:
        preview = load_app_tokens_preview()
        if preview:
            print(c(f"  → 全 {len(preview)} 件のアプリを処理します: {', '.join(preview)}", BOLD, YELLOW))
        else:
            print(c("  警告: app_tokens が空です。.kintone.env を確認してください。", BOLD, YELLOW))
    return args


def build_runner_args(item: dict) -> Optional[List[str]]:
    cmd = item["command"]
    if cmd == "users":
        return build_users_args()
    if item.get("is_group"):
        return build_group_args()
    if cmd == "all":
        return build_all_args()
    if cmd == "outputs":
        return ["outputs"]
    if cmd == "app":
        return build_app_args(item)

    args = [cmd]
    if item.get("needs_app_id"):
        app_id = read_app_id(optional=True)
        if app_id:
            args.extend(["--id", app_id])
    return args


def run_runner(runner_args: List[str], env_path: Optional[Path] = None, highlight: bool = False) -> int:
    cmd = [sys.executable, str(RUNNER), *runner_args]
    if env_path:
        cmd.extend(["--env", str(env_path)])
    print()
    exec_line = f"  実行: python {RUNNER.name} {' '.join(runner_args)}"
    if highlight:
        print(c(exec_line, BOLD, CYAN))
    else:
        print(exec_line)
    print("-" * 60)
    result = subprocess.run(cmd, cwd=str(SCRIPT_DIR))
    print("-" * 60)
    return result.returncode


def main():
    if not RUNNER.exists():
        print(f"エラー: {RUNNER} が見つかりません。")
        sys.exit(1)

    print_header()

    while True:
        print_menu()
        valid = {str(i) for i in range(len(MENU_ITEMS) + 1)}
        choice = read_choice("番号を選択してください: ", valid)

        if choice == "0":
            print("終了します。")
            break

        item = MENU_ITEMS[int(choice) - 1]
        runner_args = build_runner_args(item)
        if runner_args is None:
            continue

        run_runner(runner_args, highlight=item.get("highlight", False))

        print()
        again = input("メニューに戻りますか？ [Y/n]: ").strip().lower()
        if again in ("n", "no"):
            print("終了します。")
            break
        print()


if __name__ == "__main__":
    main()
