#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Kintone統合実行ツールの対話型メニュー。

python menu.py で起動し、番号選択で kintone_runner.py の各機能を実行します。
詳細表示は python menu.py --detail です。

設定はすべて同じフォルダの .kintone.env に書きます。
アプリID と管理用 API トークンの対応は app_tokens に登録します。
"""

import argparse
import getpass
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
from urllib.parse import urlparse

try:
    import yaml
except ImportError:
    yaml = None  # type: ignore

from inspect_downloaded import (
    export_search_hits_to_excel,
    extract_webhook_rows,
    fetch_and_save_webhooks,
    find_app_output_dir,
    find_webhook_file,
    is_stale_webhook_file,
    list_app_output_dirs,
    load_json_or_yaml,
    search_downloaded,
    summarize_hits,
)

SCRIPT_DIR = Path(__file__).resolve().parent
RUNNER = SCRIPT_DIR / "kintone_runner.py"
ENV_FILE = SCRIPT_DIR / ".kintone.env"
OUTPUT_DIR = SCRIPT_DIR / "output"
DETAIL_MODE = False
PREFERRED_ENV_KEYS = [
    "subdomain",
    "username",
    "password",
    "user_domain",
    "app_tokens",
    "search_keywords",
    "js_dirs",
]

# コンソール色（Windows 10+ / ターミナル対応）
RESET = "\033[0m"
BOLD = "\033[1m"
CYAN = "\033[96m"
YELLOW = "\033[93m"
GREEN = "\033[92m"
DIM = "\033[2m"
MARK = f"{BOLD}{YELLOW}★{RESET}"

# 一括取得でダウンロードする主な情報（説明用）
DOWNLOAD_ITEMS = [
    "アプリ基本設定 / フォーム（フィールド・レイアウト） / 一覧（ビュー）",
    "アプリ・レコード・フィールドの ACL（アクセス権）",
    "通知（一般・レコード・リマインダー） / プロセス管理 / アクション / グラフ / プラグイン",
    "JavaScript / CSS カスタマイズ（ファイル本体を javascript/ に保存）",
]

MENU_ITEMS = [
    {
        "label": "アプリ設定・JavaScript の一括取得",
        "command": "app",
        "highlight": True,
        "brief": "設定・JSを ./output へ保存",
        "description": ".kintone.env の app_tokens（アプリID と APIトークンの対応）を使い、取得できる設定をすべて ./output へ保存",
        "detail_lines": [
            "取得: フォーム / ACL / 通知 / プロセス / ビュー / グラフ / プラグイン / カスタマイズJS など",
            "アプリID 未入力=登録済み全件 / 指定=1件   ※ トークンに「レコード閲覧」が必要（アプリ管理だけでは 403）",
        ],
        "output": "./output/{アプリID}_{アプリ名}_{日時}/",
        "needs_app_id": True,
    },
    {
        "label": "接続情報の設定（URL / ログイン / アプリ）",
        "command": "set_connection",
        "is_set_connection": True,
        "brief": ".kintone.env を作成・更新",
        "description": "subdomain・username・password・app_tokens を入力して .kintone.env に保存（既存は上書き、Enterは維持）",
        "output": ".kintone.env",
    },
    {
        "label": "kintone アプリIDとトークンを追加",
        "command": "add_token",
        "is_add_token": True,
        "brief": "app_tokens に1件追加",
        "description": "アプリIDとAPIトークンを入力して .kintone.env の app_tokens に追加（既存IDは上書き）",
        "output": ".kintone.env",
    },
    {
        "label": "ACL情報をExcelに変換",
        "command": "acl",
        "brief": "ACLをExcelレポート化",
        "description": "先に一括取得したアプリの ACL を Excel レポートに変換",
        "output": "output/acl_report_[アプリID]_[日時].xlsx",
        "needs_app_id": True,
    },
    {
        "label": "アプリ設定一覧表の出力",
        "command": "summary",
        "brief": "設定全体をExcel一覧に",
        "description": "先に一括取得したアプリ設定の全体一覧を Excel で出力",
        "output": "output/kintone_app_settings_summary_[日時].xlsx",
    },
    {
        "label": "通知設定をExcelに変換",
        "command": "notifications",
        "brief": "通知をExcel化",
        "description": "先に一括取得した一般・レコード・リマインダー通知を Excel に変換",
        "output": "output/[アプリID]_notifications.xlsx",
        "needs_app_id": True,
    },
    {
        "label": "プロセスワークフローをExcelに変換",
        "command": "process_workflow",
        "brief": "プロセス管理をExcel化",
        "description": "先に一括取得したプロセス管理設定を Excel に変換",
        "output": "output/[アプリID]_process_workflow.xlsx",
        "needs_app_id": True,
    },
    {
        "label": "Webhook設定一覧の表示",
        "command": "webhooks",
        "is_webhooks": True,
        "brief": "Webhookを一覧表示",
        "description": "先に一括取得したアプリの Webhook 設定を一覧表示",
        "output": "コンソール出力",
    },
    {
        "label": "検索文字列の設定",
        "command": "search_keywords",
        "is_search_keywords": True,
        "brief": "search_keywords を編集",
        "description": "取得済みデータ検索で使う文字列を .kintone.env の search_keywords に追加・削除",
        "output": ".kintone.env",
    },
    {
        "label": "取得済みデータから検索（JS含む）",
        "command": "search_downloaded",
        "is_search_downloaded": True,
        "brief": "YAML/JSON/JS を検索（Excel既定オン）",
        "description": "アプリID省略=取得済み全件。検索語・アプリ毎の合致サマリとヒット一覧を Excel 出力（既定オン）",
        "output": "コンソール + output/search_hits_[日時].xlsx",
    },
    {
        "label": "ユーザーとグループ情報の取得",
        "command": "users",
        "brief": "ユーザー一覧をExcel/CSV出力",
        "description": "全ユーザーと所属グループを Excel/CSV に出力（ユーザー名・パスワード認証）",
        "output": "output/kintone_users_groups_[日時].xlsx",
    },
    {
        "label": "グループ操作",
        "command": "group",
        "brief": "一覧・検索・追加・削除",
        "description": "一覧表示・ユーザー検索・追加・削除（ユーザー名・パスワード認証）",
        "output": "コンソール出力",
        "is_group": True,
    },
    {
        "label": "すべての機能を順番に実行",
        "command": "all",
        "brief": "主要機能を一括実行",
        "description": "users → app → acl → summary → notifications → process_workflow",
        "output": "上記すべての出力ファイル",
        "needs_app_filter": True,
    },
    {
        "label": "出力ファイル一覧の表示",
        "command": "outputs",
        "brief": "生成ファイルの概要",
        "description": "生成される Excel/CSV の一覧と概要を表示",
        "output": "（表示のみ）",
    },
    {
        "label": "設定ファイルの説明・登録状況",
        "command": "config",
        "is_config_help": True,
        "brief": ".kintone.env の内容を表示",
        "description": ".kintone.env の場所・書き方・現在の登録内容を表示（トークンは伏せます）",
        "output": "（表示のみ）",
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


def load_env_preview() -> Dict[str, Any]:
    """.kintone.env の概要を返す（トークン値は含めない）。"""
    preview: Dict[str, Any] = {
        "exists": ENV_FILE.exists(),
        "path": ENV_FILE,
        "readable": False,
        "subdomain": None,
        "username": None,
        "has_password": False,
        "user_domain": None,
        "app_ids": [],
        "search_keywords": [],
        "js_dirs": [],
        "missing": [],
        "error": None,
    }
    if not preview["exists"]:
        preview["missing"] = ["ファイル自体", "subdomain", "username", "password", "app_tokens"]
        return preview
    if yaml is None:
        preview["error"] = "PyYAML が未インストールです（pip install pyyaml）"
        return preview
    try:
        with open(ENV_FILE, encoding="utf-8") as f:
            config = yaml.safe_load(f) or {}
        if not isinstance(config, dict):
            preview["error"] = "YAML の形式が不正です（辞書ではありません）"
            return preview
        preview["readable"] = True
        preview["subdomain"] = config.get("subdomain") or None
        preview["username"] = config.get("username") or None
        preview["has_password"] = bool(config.get("password"))
        preview["user_domain"] = config.get("user_domain") or None
        tokens = config.get("app_tokens") or {}
        if isinstance(tokens, dict):
            preview["app_ids"] = [str(k) for k in tokens.keys()]
        keywords = config.get("search_keywords") or []
        if isinstance(keywords, list):
            preview["search_keywords"] = [str(k) for k in keywords if str(k).strip()]
        elif keywords:
            preview["search_keywords"] = [str(keywords)]
        else:
            preview["search_keywords"] = []
        js_dirs = config.get("js_dirs") or {}
        if isinstance(js_dirs, dict):
            preview["js_dirs"] = [str(k) for k in js_dirs.keys()]
        for key, label in (
            ("subdomain", "subdomain"),
            ("username", "username"),
            ("password", "password"),
        ):
            if not config.get(key):
                preview["missing"].append(label)
        if not preview["app_ids"]:
            preview["missing"].append("app_tokens")
    except Exception as e:
        preview["error"] = str(e)
    return preview


def load_app_tokens_preview() -> List[str]:
    """.kintone.env から app_tokens のアプリ番号一覧を取得する。"""
    return list(load_env_preview().get("app_ids") or [])


GROUP_ACTIONS = [
    ("list", "グループ一覧を表示"),
    ("search", "ユーザーを検索"),
    ("add", "ユーザーをグループに追加"),
    ("remove", "ユーザーをグループから削除"),
]


def print_header():
    print()
    print("=" * 72)
    print(c("  Kintone統合実行ツール", BOLD, CYAN))
    print("  アプリID と管理用 API トークンを .kintone.env に登録し、")
    print("  そのアプリから取得できる設定・JS を可能な限りダウンロードします。")
    print("=" * 72)
    print()


def print_config_status():
    """設定ファイルの場所と現在の登録状況を毎回表示する。"""
    preview = load_env_preview()
    compact_ok = (
        not DETAIL_MODE
        and preview.get("exists")
        and not preview.get("error")
        and not preview.get("missing")
    )
    if compact_ok:
        app_ids = preview.get("app_ids") or []
        apps = f"{len(app_ids)} 件  [{', '.join(app_ids)}]" if app_ids else "0 件"
        print(c("【設定】", BOLD))
        print(c(f"  {preview['path']}  対象アプリ: {apps}", CYAN))
        print()
        return

    print(c("【設定箇所】", BOLD))
    print(c(f"  ファイル: {preview['path']}", CYAN))
    print("  書き方  : YAML。アプリID とトークンは app_tokens に対で書く")
    print()

    if not preview["exists"]:
        print(c(f"  警告: {ENV_FILE.name} がありません。この場で不足分を入力できます。", BOLD, YELLOW))
        print()
        print_env_sample(indent="  ")
        print()
        if fill_missing_fields():
            preview = load_env_preview()
            if preview.get("readable"):
                print_connection_fields(preview)
                print()
        return

    if preview["error"]:
        print(c(f"  読み込みエラー: {preview['error']}", BOLD, YELLOW))
        print()
        return

    print_connection_fields(preview)
    if preview.get("missing"):
        print()
        fill_missing_fields()
        preview = load_env_preview()
        print()
        print(c("  現在の接続情報:", BOLD, CYAN))
        print_connection_fields(preview)
        print()
        if preview.get("missing"):
            print(c("  まだ不足している項目は、上の入力または設定メニューで追加できます。", DIM))
        else:
            print(c("  接続情報は設定済みです。", GREEN))
    else:
        print()
        print(c("  接続情報は設定済みです。", GREEN))
    print(c("  変更: メニュー「接続情報の設定（URL / ログイン / アプリ）」", DIM))
    print(c("  アプリ追加: メニュー「kintone アプリIDとトークンを追加」", DIM))
    print()


def print_connection_fields(preview: Dict[str, Any]) -> None:
    """接続先・ログイン・対象アプリの現在値を表示する。"""
    subdomain = preview["subdomain"] or "（未設定）"
    username = preview["username"] or "（未設定）"
    password = "設定済み" if preview["has_password"] else "（未設定）"
    app_ids = preview.get("app_ids") or []

    print(f"  接続先 subdomain : {subdomain}")
    if preview["subdomain"]:
        print(c(f"                    → https://{preview['subdomain']}.cybozu.com/", DIM))
    print(f"  ログイン username: {username}")
    print(f"  パスワード       : {password}")
    if app_ids:
        print(c(f"  対象アプリ       : {len(app_ids)} 件  [{', '.join(app_ids)}]", BOLD, CYAN))
    else:
        print(c("  対象アプリ       : 0 件  ※ app_tokens が空です", BOLD, YELLOW))

    missing = preview.get("missing") or []
    if missing:
        print(c(f"  不足している項目 : {', '.join(missing)}", YELLOW))


def print_env_sample(indent: str = "  "):
    """設定ファイルの記入例を表示する。"""
    sample = [
        "# 同じフォルダの .kintone.env に、次の形式で書いてください",
        "subdomain: \"your_subdomain\"    # example.cybozu.com の example 部分",
        "username: \"admin@example.com\"  # 管理者のログイン名",
        "password: \"your_password\"",
        "user_domain: \"example.com\"     # 任意。ユーザー一覧のメインドメイン",
        "",
        "# ここが対象アプリの登録箇所（アプリID: APIトークン）",
        "app_tokens:",
        "  123: \"xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\"",
        "  456: \"yyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy\"",
    ]
    print(c(f"{indent}----- {ENV_FILE.name} の記入例 -----", YELLOW))
    for line in sample:
        print(c(f"{indent}{line}", GREEN))
    print(c(f"{indent}--------------------------------", YELLOW))


def print_config_help():
    """設定ファイルの場所・書き方・現在値を詳しく表示する。"""
    preview = load_env_preview()
    print()
    print("=" * 72)
    print(c("  設定ファイルの説明", BOLD, CYAN))
    print("=" * 72)
    print()
    print(c("■ 設定箇所", BOLD))
    print(f"  パス     : {preview['path']}")
    print(f"  存在     : {'あり' if preview['exists'] else 'なし（要作成）'}")
    print("  編集方法 : テキストエディタで開いて保存（YAML形式）")
    print("  注意     : パスワードとトークンが入るため、Git にコミットしないこと")
    print()
    print(c("■ 必須項目", BOLD))
    print("  subdomain   … 接続先（xxx.cybozu.com の xxx）")
    print("  username    … 管理者ログイン名（ユーザー・グループ機能で使用）")
    print("  password    … 上記のパスワード")
    print("  app_tokens  … ダウンロード対象。アプリID と APIトークンの対応表")
    print()
    print(c("■ 任意項目", BOLD))
    print("  user_domain … ユーザー一覧 Excel で優先表示するメインドメイン")
    print("  js_dirs     … 追加で突き合わせるローカル JS フォルダ（任意）")
    print()
    print(c("■ APIトークンの用意", BOLD))
    print("  1. kintone で対象アプリを開く")
    print("  2. 「アプリの設定」→「設定」タブ →「APIトークン」")
    print("  3. 生成し、少なくとも「レコード閲覧」とファイルダウンロード相当の権限を付ける")
    print("  4. アプリを更新してから、.kintone.env の app_tokens に")
    print("       <アプリID>: \"<発行したトークン>\"")
    print("     の行を追加する")
    print()
    print(c("■ 一括取得で保存される情報", BOLD))
    for item in DOWNLOAD_ITEMS:
        print(f"  ・{item}")
    print("  出力先: ./output/{アプリID}_{アプリ名}_{日時}/")
    print()

    print_env_sample(indent="  ")
    print()

    print(c("■ 現在の登録状況", BOLD))
    if not preview["exists"]:
        print(c(f"  {ENV_FILE.name} が見つかりません。上記の例をコピーして作成してください。", YELLOW))
    elif preview["error"]:
        print(c(f"  読み込みエラー: {preview['error']}", YELLOW))
    else:
        print(f"  subdomain : {preview['subdomain'] or '（未設定）'}")
        print(f"  username  : {preview['username'] or '（未設定）'}")
        print(f"  password  : {'設定済み' if preview['has_password'] else '（未設定）'}")
        print(f"  user_domain: {preview['user_domain'] or '（未設定）'}")
        if preview["js_dirs"]:
            print(f"  js_dirs   : {', '.join(preview['js_dirs'])}")
        if preview["app_ids"]:
            print(c(f"  app_tokens: {len(preview['app_ids'])} 件", BOLD, CYAN))
            for app_id in preview["app_ids"]:
                print(c(f"    - アプリID {app_id}  … トークン登録済み", CYAN))
        else:
            print(c("  app_tokens: （未設定）  ← 一括取得の対象がありません", YELLOW))
        if preview["missing"]:
            print(c(f"  不足      : {', '.join(preview['missing'])}", YELLOW))
    print()
    print("=" * 72)


def print_menu_item(index: int, item: dict):
    num = f"{index:2}"
    brief = item.get("brief") or ""
    label = f"{MARK} {item['label']}" if item.get("highlight") else item["label"]
    first = f"  {num}. {label}"
    if brief:
        first = f"{first}  … {brief}"
    second = item["description"]
    if item.get("highlight"):
        app_ids = load_app_tokens_preview()
        if app_ids:
            second = f"{second}  対象: {', '.join(app_ids)}"
        elif ENV_FILE.exists():
            second = f"{second}  対象: なし → app_tokens を設定してください"
        else:
            second = f"{second}  ※ 先に {ENV_FILE.name} を作成してください"
        print(c(first, BOLD, CYAN))
        print(c(f"      {second}", YELLOW))
    else:
        print(first)
        print(f"      {second}")


def print_menu():
    print_config_status()
    print(c("【実行メニュー】", BOLD))
    print()
    for i, item in enumerate(MENU_ITEMS, start=1):
        print_menu_item(i, item)
    print()
    print("   0. 終了")
    print()


def read_yes_no(prompt: str, default: bool = True) -> bool:
    hint = "Y/n" if default else "y/N"
    raw = input(f"  {prompt} [{hint}]: ").strip().lower()
    if not raw:
        return default
    return raw in ("y", "yes")


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


def build_all_args() -> Optional[List[str]]:
    args = ["all"]
    print()
    print("  アプリIDの絞り込み（任意）")
    print("  1. 全アプリ")
    print("  2. 指定したアプリIDのみ")
    print("  3. 指定したアプリIDを除外")
    filter_choice = input("  選択 [1]: ").strip() or "1"
    if filter_choice == "2":
        raw = input("  対象アプリID（スペース区切り）: ").strip()
        ids = [x for x in raw.split() if x]
        if ids:
            for app_id in ids:
                if not app_id.isdigit() or not ensure_app_token(app_id):
                    return None
            args.extend(["--id", *ids])
    elif filter_choice == "3":
        raw = input("  除外するアプリID（スペース区切り）: ").strip()
        if raw:
            args.extend(["--not-id", *raw.split()])
    if not load_app_tokens_preview():
        print(c("  app_tokens が空です。先にアプリIDとトークンを入力してください。", BOLD, YELLOW))
        if not fill_missing_fields(required=["app_tokens"]) or not load_app_tokens_preview():
            return None
    return args


def print_app_run_confirm(item: dict):
    """アプリ一括取得実行前に概要を表示する。"""
    print()
    print(c(f"  {MARK} {item['label']}", BOLD, CYAN))
    print(c(f"  設定ファイル: {ENV_FILE}", CYAN))
    print(c("  app_tokens に書いたアプリから、取得できる設定をすべてダウンロードします。", YELLOW))
    for line in DOWNLOAD_ITEMS:
        print(c(f"    ・{line}", GREEN))
    app_ids = load_app_tokens_preview()
    if app_ids:
        print(c(f"  登録済みアプリID: {', '.join(app_ids)}", CYAN))
    else:
        print(c(f"  警告: {ENV_FILE.name} の app_tokens が空です。先にアプリIDとトークンを書いてください。", BOLD, YELLOW))
    print()


def build_app_args(item: dict) -> Optional[List[str]]:
    print_app_run_confirm(item)
    app_id = read_app_id(optional=True)
    if app_id:
        if not ensure_app_token(app_id):
            return None
        return ["app", "--id", app_id]

    preview = load_app_tokens_preview()
    if preview:
        print(c(f"  → 全 {len(preview)} 件のアプリを処理します: {', '.join(preview)}", BOLD, YELLOW))
        return ["app"]

    print(c("  app_tokens が空のため、先に1件登録してください。", BOLD, YELLOW))
    new_id = input("  アプリID: ").strip()
    if not new_id:
        print("  キャンセルしました。")
        return None
    if not new_id.isdigit():
        print("  アプリID は数値で入力してください。")
        return None
    if not ensure_app_token(new_id):
        return None
    return ["app", "--id", new_id]


def load_env_config_raw() -> Dict[str, Any]:
    """.kintone.env を辞書として読み込む（無い・空なら空辞書）。"""
    if yaml is None:
        raise RuntimeError("PyYAML が未インストールです（pip install pyyaml）")
    if not ENV_FILE.exists():
        return {}
    with open(ENV_FILE, encoding="utf-8") as f:
        content = f.read()
    if not content.strip():
        return {}
    data = yaml.safe_load(content)
    if data is None:
        return {}
    if not isinstance(data, dict):
        raise ValueError("YAML の形式が不正です（辞書ではありません）")
    return data


def save_env_config(config: Dict[str, Any]) -> None:
    """設定を .kintone.env に書き戻す。"""
    with open(ENV_FILE, "w", encoding="utf-8", newline="\n") as f:
        yaml.dump(config, f, default_flow_style=False, allow_unicode=True, sort_keys=False)


def merge_save_env(updates: Dict[str, Any]) -> None:
    """既存設定に更新をマージして保存する。"""
    config = load_env_config_raw()
    config.update(updates)
    ordered: Dict[str, Any] = {}
    for key in PREFERRED_ENV_KEYS:
        if key in config:
            ordered[key] = config[key]
    for key, value in config.items():
        if key not in ordered:
            ordered[key] = value
    save_env_config(ordered)


def _is_field_missing(preview: Dict[str, Any], key: str) -> bool:
    if key == "subdomain":
        return not preview.get("subdomain")
    if key == "username":
        return not preview.get("username")
    if key == "password":
        return not preview.get("has_password")
    if key == "app_tokens":
        return not preview.get("app_ids")
    return False


def fill_missing_fields(required: Optional[List[str]] = None) -> bool:
    """未設定の項目だけ入力して保存する。Enter はスキップ。"""
    preview = load_env_preview()
    if preview.get("error") and yaml is None:
        print(c(f"  {preview['error']}", YELLOW))
        return False

    keys = required or ["subdomain", "username", "password", "app_tokens"]
    missing = [key for key in keys if _is_field_missing(preview, key)]
    if not missing:
        return True

    print(c("  未設定の項目を入力してください（Enter でスキップ）。", BOLD, YELLOW))
    updates: Dict[str, Any] = {}

    if "subdomain" in missing:
        raw = input("  kintone URL: ").strip()
        if raw:
            parsed = parse_kintone_subdomain(raw)
            if parsed:
                updates["subdomain"] = parsed
            else:
                print("  URL の形式が不正です。https://xxxx.cybozu.com/ または サブドメインを入力してください。")

    if "username" in missing:
        raw = input("  ログイン username: ").strip()
        if raw:
            updates["username"] = raw

    if "password" in missing:
        try:
            raw = getpass.getpass("  パスワード: ")
        except Exception:
            raw = input("  パスワード: ")
        if raw and raw.strip():
            updates["password"] = raw.strip()

    if updates:
        try:
            merge_save_env(updates)
        except Exception as e:
            print(c(f"  保存に失敗しました: {e}", YELLOW))
            return False

    if "app_tokens" in missing:
        raw_id = input("  アプリID: ").strip()
        if raw_id:
            if not raw_id.isdigit():
                print("  アプリID は数値で入力してください。")
            else:
                token = input("  APIトークン: ").strip()
                if token:
                    try:
                        save_app_token(raw_id, token)
                        print(c(f"  アプリID {raw_id} のトークンを保存しました。", CYAN))
                    except Exception as e:
                        print(c(f"  保存に失敗しました: {e}", YELLOW))
                        return False
                else:
                    print("  トークンが未入力のため、アプリは登録しませんでした。")

    return True


def save_app_token(app_id: str, token: str) -> None:
    """app_tokens にアプリIDとトークンを追加／上書きして保存する。"""
    config = load_env_config_raw()
    tokens = config.get("app_tokens")
    if not isinstance(tokens, dict):
        tokens = {}
    key = resolve_app_token_key(tokens, app_id)
    tokens[key] = token
    config["app_tokens"] = tokens
    save_env_config(config)


def ensure_app_token(app_id: str) -> bool:
    """対象アプリのトークンが無ければ入力して保存する。未入力なら False。"""
    if app_id in load_app_tokens_preview():
        return True
    print(c(f"  アプリID {app_id} の APIトークンが未登録です。", BOLD, YELLOW))
    print("  この場で .kintone.env に保存できます。未入力で中止。")
    token = input("  APIトークン: ").strip()
    if not token:
        print("  トークンが未入力のため中止しました。")
        return False
    try:
        save_app_token(app_id, token)
    except Exception as e:
        print(c(f"  保存に失敗しました: {e}", YELLOW))
        return False
    print(c(f"  アプリID {app_id} のトークンを保存しました。続けて取得します。", BOLD, CYAN))
    return True


def resolve_app_token_key(tokens: dict, app_id: str):
    """既存キーの型（int / str）に合わせて app_tokens のキーを決める。"""
    if app_id in tokens:
        return app_id
    if app_id.isdigit():
        as_int = int(app_id)
        if as_int in tokens:
            return as_int
        return as_int
    return app_id


def parse_kintone_subdomain(raw: str) -> Optional[str]:
    """URL またはサブドメイン文字列から subdomain を取り出す。"""
    text = raw.strip().strip("'\"")
    if not text:
        return None
    if re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9-]*", text):
        return text
    if "://" not in text:
        text = "https://" + text
    host = (urlparse(text).hostname or "").lower()
    match = re.fullmatch(r"([a-z0-9][a-z0-9-]*)\.(cybozu\.com|kintone\.com)", host)
    if match:
        return match.group(1)
    return None


def prompt_keep(label: str, current: Optional[str] = None, placeholder: str = "未設定") -> Optional[str]:
    """値を入力する。空Enterなら None（現状維持）。"""
    shown = current if current else placeholder
    raw = input(f"  {label} [Enterで維持: {shown}]: ").strip()
    return raw or None


def set_connection_interactive() -> None:
    """URL / ログイン / アプリトークンをまとめて .kintone.env に保存する。"""
    print()
    print(c("  接続情報の設定", BOLD, CYAN))
    print(c(f"  保存先: {ENV_FILE}", CYAN))
    print("  画面上の subdomain / username / password / app_tokens をセットします。")
    print("  各項目は Enter で現状維持。値を入れると上書きします。")
    print()

    preview = load_env_preview()
    print_connection_fields(preview)
    print()

    current_url = f"https://{preview['subdomain']}.cybozu.com/" if preview.get("subdomain") else "未設定"
    raw_url = prompt_keep("kintone URL", current_url)
    subdomain = preview.get("subdomain")
    if raw_url:
        parsed = parse_kintone_subdomain(raw_url)
        if not parsed:
            print("  URL の形式が不正です。https://xxxx.cybozu.com/ または サブドメイン を入力してください。")
            return
        subdomain = parsed

    username = preview.get("username")
    raw_user = prompt_keep("ログイン username", username)
    if raw_user:
        username = raw_user

    pwd_shown = "設定済み" if preview.get("has_password") else "未設定"
    try:
        raw_password = getpass.getpass(f"  パスワード [Enterで維持: {pwd_shown}]: ")
    except Exception:
        raw_password = input(f"  パスワード [Enterで維持: {pwd_shown}]: ")
    password = raw_password.strip() if raw_password and raw_password.strip() else None

    print()
    print("  対象アプリ（app_tokens）。追加しない場合はアプリIDを空のまま Enter。")
    raw_app_id = input("  アプリID: ").strip()
    raw_token = ""
    if raw_app_id:
        if not raw_app_id.isdigit():
            print("  アプリID は数値で入力してください。")
            return
        raw_token = input("  APIトークン: ").strip()
        if not raw_token:
            print("  トークンが未入力のため、アプリの追加はスキップします。")
            raw_app_id = ""

    try:
        config = load_env_config_raw()
    except Exception as e:
        print(c(f"  読み込みエラー: {e}", YELLOW))
        return

    tokens = config.get("app_tokens")
    if not isinstance(tokens, dict):
        tokens = {}
    if raw_app_id and raw_token:
        key = resolve_app_token_key(tokens, raw_app_id)
        tokens[key] = raw_token

    new_config: Dict[str, Any] = {}
    if subdomain:
        new_config["subdomain"] = subdomain
    if username:
        new_config["username"] = username
    existing_password = config.get("password") if password is None else password
    if existing_password:
        new_config["password"] = existing_password
    if tokens:
        new_config["app_tokens"] = tokens
    for key, value in config.items():
        if key not in new_config:
            new_config[key] = value

    try:
        save_env_config(new_config)
    except Exception as e:
        print(c(f"  保存に失敗しました: {e}", YELLOW))
        return

    print()
    print(c("  保存しました。現在の接続情報:", BOLD, CYAN))
    print_connection_fields(load_env_preview())
    print()


def add_app_token_interactive() -> None:
    """アプリIDとAPIトークンを入力し、.kintone.env の app_tokens に追加／上書きする。"""
    print()
    print(c("  kintone アプリID と APIトークンの追加", BOLD, CYAN))
    print(c(f"  保存先: {ENV_FILE}", CYAN))
    print("  同じアプリIDがあれば上書きします。未入力でキャンセル。")
    print()

    app_id = input("  アプリID: ").strip()
    if not app_id:
        print("  キャンセルしました。")
        return
    if not app_id.isdigit():
        print("  アプリID は数値で入力してください。")
        return

    token = input("  APIトークン: ").strip()
    if not token:
        print("  トークンが未入力のためキャンセルしました。")
        return

    overwritten = app_id in load_app_tokens_preview()
    try:
        save_app_token(app_id, token)
    except Exception as e:
        print(c(f"  保存に失敗しました: {e}", YELLOW))
        return

    if overwritten:
        print(c(f"  アプリID {app_id} のトークンを上書きして保存しました。", BOLD, YELLOW))
    else:
        print(c(f"  アプリID {app_id} のトークンを追加して保存しました。", BOLD, CYAN))


def load_search_keywords() -> List[str]:
    return list(load_env_preview().get("search_keywords") or [])


def save_search_keywords(keywords: List[str]) -> None:
    unique: List[str] = []
    for word in keywords:
        text = str(word).strip()
        if text and text not in unique:
            unique.append(text)
    merge_save_env({"search_keywords": unique})


def print_search_keywords(keywords: Optional[List[str]] = None) -> None:
    words = keywords if keywords is not None else load_search_keywords()
    if words:
        print(c(f"  追加済み検索文字列: {len(words)} 件", CYAN))
        for i, word in enumerate(words, start=1):
            print(f"    {i}. {word}")
    else:
        print(c("  追加済み検索文字列: （なし）", DIM))


def read_keyword_lines(prompt: str) -> List[str]:
    print(f"  {prompt}")
    print("  1行に1つ。空行で終了。")
    collected: List[str] = []
    while True:
        raw = input("  検索文字列: ").strip()
        if not raw:
            break
        collected.append(raw)
    return collected


def read_required_app_id() -> Optional[str]:
    raw = input("  アプリID: ").strip()
    if not raw:
        print("  キャンセルしました。")
        return None
    if not raw.isdigit():
        print("  アプリID は数値で入力してください。")
        return None
    return raw


def resolve_app_output_dir(app_id: str) -> Optional[Path]:
    app_dir = find_app_output_dir(OUTPUT_DIR, app_id)
    if app_dir:
        return app_dir
    print(c(f"  アプリID {app_id} の取得済みフォルダが見つかりません。", BOLD, YELLOW))
    print("  先にメニュー「アプリ設定・JavaScript の一括取得」を実行してください。")
    return None


def print_webhook_rows(rows: List[dict]) -> None:
    if not rows:
        print("  Webhook は 0 件です。")
        return
    print(c(f"  {len(rows)} 件", BOLD, CYAN))
    print()
    for i, row in enumerate(rows, start=1):
        enabled = row.get("enabled")
        if enabled is True:
            status = "有効"
        elif enabled is False:
            status = "無効"
        else:
            status = str(enabled) if enabled != "" else "-"
        print(f"  {i}. {row.get('name') or '(名称なし)'}  [{status}]")
        if row.get("id"):
            print(f"     ID    : {row['id']}")
        print(f"     URL   : {row.get('url') or '-'}")
        print(f"     イベント: {row.get('events') or '-'}")
        if row.get("headers"):
            print(f"     ヘッダ: {row['headers']}")
        if row.get("creator"):
            print(f"     作成者 : {row['creator']}")
        if row.get("modifier"):
            print(f"     更新者 : {row['modifier']}")
        if row.get("modified_at"):
            print(f"     更新日時: {row['modified_at']}")
        print()


def resolve_app_token(app_id: str) -> Optional[str]:
    try:
        config = load_env_config_raw()
    except Exception:
        return None
    tokens = config.get("app_tokens") or {}
    if not isinstance(tokens, dict):
        return None
    if app_id in tokens:
        return tokens[app_id]
    if app_id.isdigit() and int(app_id) in tokens:
        return tokens[int(app_id)]
    return None


def show_webhooks_interactive() -> None:
    print()
    print(c("  Webhook設定一覧", BOLD, CYAN))
    app_id = read_required_app_id()
    if not app_id:
        return
    app_dir = resolve_app_output_dir(app_id)
    if not app_dir:
        return
    print(c(f"  参照: {app_dir}", DIM))
    webhook_file = find_webhook_file(app_dir, app_id)
    data = None
    if webhook_file:
        try:
            data = load_json_or_yaml(webhook_file)
            print(c(f"  ファイル: {webhook_file.relative_to(app_dir)}", CYAN))
        except Exception as e:
            print(c(f"  読み込みエラー: {e}", YELLOW))
            webhook_file = None
        if data is not None and is_stale_webhook_file(data):
            print(c(
                "  保存済みファイルは旧方式（管理画面HTMLの走査）で取得したもののため再取得します。",
                YELLOW,
            ))
            data = None

    if data is None:
        print(
            f"  管理画面と同じ内部API（/k/api/dev/app/{app_id}/webhook/list.json）で取得します。"
        )
        preview = load_env_preview()
        if preview.get("subdomain"):
            print(c(
                f"  https://{preview['subdomain']}.cybozu.com/k/admin/app/webhook?app={app_id}",
                DIM,
            ))
        try:
            config = load_env_config_raw()
        except Exception as e:
            print(c(f"  設定の読み込みに失敗しました: {e}", YELLOW))
            return
        saved, data, error = fetch_and_save_webhooks(
            app_dir,
            app_id,
            subdomain=config.get("subdomain") or "",
            api_token=resolve_app_token(app_id),
            username=config.get("username"),
            password=config.get("password"),
        )
        if saved and data is not None:
            print(c(f"  取得して保存しました: {saved.relative_to(app_dir)}", GREEN))
        else:
            print(c("  Webhook一覧APIからは取得できませんでした。", YELLOW))
            if error:
                print(c(f"  ({error})", DIM))
            print("  取得済みファイル内の webhook という文字列を代わりに表示します。")
            hits = search_downloaded(app_dir, ["webhook", "Webhook", "WEBHOOK"])
            if hits:
                print_search_hits(["webhook"], hits)
            else:
                print("  取得済みデータ内にも webhook の記載はありません（0件）。")
            return

    print_webhook_rows(extract_webhook_rows(data))


def manage_search_keywords_interactive() -> None:
    print()
    print(c("  検索文字列の設定", BOLD, CYAN))
    print("  取得済みデータ検索で使う文字列を追加・削除します。")
    print()
    while True:
        print_search_keywords()
        print()
        print("  1. 追加")
        print("  2. 削除")
        print("  0. 戻る")
        choice = input("  選択: ").strip() or "0"
        if choice == "0":
            return
        if choice == "1":
            added = read_keyword_lines("追加する検索文字列を入力してください。")
            if not added:
                print("  追加はありません。")
                continue
            current = load_search_keywords()
            save_search_keywords(current + added)
            print(c(f"  {len(added)} 件を追加して保存しました。", CYAN))
            continue
        if choice == "2":
            current = load_search_keywords()
            if not current:
                print("  削除する文字列がありません。")
                continue
            raw = input("  削除する番号（スペース区切り）: ").strip()
            if not raw:
                continue
            remove_idx = set()
            for part in raw.split():
                if part.isdigit() and 1 <= int(part) <= len(current):
                    remove_idx.add(int(part) - 1)
            if not remove_idx:
                print("  有効な番号がありません。")
                continue
            remain = [word for i, word in enumerate(current) if i not in remove_idx]
            save_search_keywords(remain)
            print(c(f"  {len(remove_idx)} 件を削除して保存しました。", YELLOW))
            continue
        print("  1 / 2 / 0 を入力してください。")


def print_search_hits(keywords: List[str], hits: List[dict], multi_app: bool = False) -> None:
    print()
    print(c(f"  検索語: {', '.join(keywords)}", BOLD, CYAN))
    if not hits:
        print("  該当なし")
        return
    print(c(f"  該当 {len(hits)} 行", BOLD, CYAN))
    if DETAIL_MODE:
        print("  該当したもの:")
        for kind, file_count, line_count in summarize_hits(hits):
            print(f"    ・{kind}: {file_count} ファイル / {line_count} 行")
        print()
    current_key = None
    file_keywords: Dict[Tuple[str, str], List[str]] = {}
    for hit in hits:
        key = (hit.get("app_dir") or "", hit["file"])
        seen = file_keywords.setdefault(key, [])
        for word in hit["keywords"]:
            if word not in seen:
                seen.append(word)
    for hit in hits:
        key = (hit.get("app_dir") or "", hit["file"])
        if key != current_key:
            current_key = key
            file_label = hit["file"]
            if multi_app and hit.get("app_dir"):
                file_label = f"{hit['app_dir']}/{hit['file']}"
            hit_words = ", ".join(file_keywords.get(key, hit["keywords"]))
            label = hit.get("category") or hit["kind"]
            print(c(f"  [{label}] {file_label}  ヒット: {hit_words}", YELLOW))
        preview = hit["text"]
        if len(preview) > 160:
            preview = preview[:157] + "..."
        print(f"    L{hit['line']:>5}  {preview}")


def search_downloaded_interactive() -> None:
    print()
    print(c("  取得済みデータから検索（JavaScript / YAML / JSON など）", BOLD, CYAN))
    app_id = read_app_id(optional=True)
    if app_id:
        app_dir = resolve_app_output_dir(app_id)
        if not app_dir:
            return
        targets: List[Tuple[str, Path]] = [(app_id, app_dir)]
    else:
        targets = list_app_output_dirs(OUTPUT_DIR)
        if not targets:
            print(c("  取得済みフォルダがありません。", BOLD, YELLOW))
            print("  先にメニュー「アプリ設定・JavaScript の一括取得」を実行してください。")
            return
        ids = ", ".join(aid for aid, _ in targets)
        print(c(f"  対象: {len(targets)} アプリ  [{ids}]", CYAN))
    if DETAIL_MODE:
        for _, app_dir in targets:
            print(c(f"  参照: {app_dir}", DIM))
    print()
    print_search_keywords()
    print()
    print("  1. 追加済みの検索文字列を使う")
    print("  2. 追加済みに追加して使う")
    print("  3. 今回だけ新規入力（保存しない）")
    print("  0. 戻る")
    choice = input("  選択: ").strip()
    if choice == "0" or not choice:
        print("  キャンセルしました。")
        return

    saved = load_search_keywords()
    keywords: List[str] = []
    if choice == "1":
        keywords = saved
        if not keywords:
            print(c("  追加済みがありません。先に「検索文字列の設定」か、2 / 3 を選んでください。", YELLOW))
            return
    elif choice == "2":
        added = read_keyword_lines("追加する検索文字列を入力してください。")
        if added:
            save_search_keywords(saved + added)
            print(c("  追加して保存しました。", CYAN))
        keywords = load_search_keywords()
        if not keywords:
            print("  検索する文字列がありません。")
            return
    elif choice == "3":
        keywords = read_keyword_lines("今回だけ使う検索文字列を入力してください（保存しません）。")
        if not keywords:
            print("  検索する文字列がありません。")
            return
    else:
        print("  1 / 2 / 3 / 0 を入力してください。")
        return

    write_excel = read_yes_no("Excelに出力しますか？（検索語・アプリ毎の合致サマリとヒット一覧）", default=True)

    hits: List[dict] = []
    for aid, app_dir in targets:
        found = search_downloaded(app_dir, keywords)
        for hit in found:
            hit["app_id"] = aid
            hit["app_dir"] = app_dir.name
        hits.extend(found)
    print_search_hits(keywords, hits, multi_app=len(targets) > 1)
    if write_excel:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = OUTPUT_DIR / f"search_hits_{timestamp}.xlsx"
        export_search_hits_to_excel(keywords, hits, excel_path)
        print(c(f"  Excel: {excel_path}", BOLD, CYAN))
    print()


def build_runner_args(item: dict) -> Optional[List[str]]:
    cmd = item["command"]
    if item.get("is_set_connection"):
        set_connection_interactive()
        return None
    if item.get("is_add_token"):
        add_app_token_interactive()
        return None
    if item.get("is_config_help"):
        print_config_help()
        return None
    if item.get("is_webhooks"):
        show_webhooks_interactive()
        return None
    if item.get("is_search_keywords"):
        manage_search_keywords_interactive()
        return None
    if item.get("is_search_downloaded"):
        search_downloaded_interactive()
        return None
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
            if not ensure_app_token(app_id):
                return None
            args.extend(["--id", app_id])
        elif not load_app_tokens_preview():
            print(c("  app_tokens が空です。先にアプリIDとトークンを入力してください。", BOLD, YELLOW))
            if not fill_missing_fields(required=["app_tokens"]) or not load_app_tokens_preview():
                return None
    return args


def print_token_permission_guide(app_id: Optional[str] = None):
    """APIトークンの権限追加手順を表示する。"""
    preview = load_env_preview()
    subdomain = preview.get("subdomain")
    print()
    print(c("  APIトークンの権限が不足しています（403）", BOLD, YELLOW))
    print("  設定取得には「レコード閲覧」が必要です。「アプリ管理」だけでは失敗します。")
    print()
    print("  次の手順で権限を追加してください:")
    print("  1. 対象アプリ → アプリの設定 → APIトークン")
    if subdomain and app_id:
        print(c(f"     https://{subdomain}.cybozu.com/k/admin/app/apitoken?app={app_id}", CYAN))
    elif subdomain:
        print(c(f"     https://{subdomain}.cybozu.com/", CYAN))
    print("  2. 使用中のトークンにチェックを付ける")
    print("       [必須] レコード閲覧")
    print("       [推奨] アプリ管理（ACL・カスタマイズJS用）")
    print("  3. 画面右下の「保存」をクリック")
    print("  4. 画面右上の「アプリを更新」で反映する")
    print()


def run_app_with_permission_retry(runner_args: List[str], highlight: bool = False) -> int:
    """アプリ取得を実行し、権限不足なら案内のあと再試行できる。"""
    app_id = None
    if "--id" in runner_args:
        idx = runner_args.index("--id")
        if idx + 1 < len(runner_args):
            app_id = runner_args[idx + 1]
    while True:
        code = run_runner(runner_args, highlight=highlight)
        if code == 0:
            return 0
        if code != 3:
            return code
        print_token_permission_guide(app_id)
        retry = input("  権限を追加してアプリを更新したら、再試行しますか？ [Y/n]: ").strip().lower()
        if retry in ("n", "no"):
            return code


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


def parse_args(argv: Optional[List[str]] = None):
    parser = argparse.ArgumentParser(description="Kintone統合実行ツールの対話型メニュー")
    parser.add_argument(
        "--detail",
        action="store_true",
        help="詳細表示あり（設定箇所の全文、検索の参照パス・内訳など）",
    )
    return parser.parse_args(argv)


def main():
    global DETAIL_MODE
    args = parse_args()
    DETAIL_MODE = args.detail

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
        if (
            not item.get("is_set_connection")
            and not item.get("is_add_token")
            and not item.get("is_config_help")
            and not item.get("is_webhooks")
            and not item.get("is_search_keywords")
            and not item.get("is_search_downloaded")
            and item.get("command") != "outputs"
        ):
            required = ["subdomain", "username", "password"]
            if item.get("command") in ("app", "all"):
                required.append("app_tokens")
            fill_missing_fields(required=required)

        runner_args = build_runner_args(item)
        if runner_args is None:
            continue

        if item.get("command") == "app":
            run_app_with_permission_retry(runner_args, highlight=item.get("highlight", False))
        else:
            run_runner(runner_args, highlight=item.get("highlight", False))

        print()
        again = input("メニューに戻りますか？ [Y/n]: ").strip().lower()
        if again in ("n", "no"):
            print("終了します。")
            break
        print()


if __name__ == "__main__":
    main()
