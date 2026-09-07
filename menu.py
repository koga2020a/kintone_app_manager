#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
kintone アプリ管理ツールの対話型メニュー。

python menu.py で起動し、番号を選ぶと kintone_runner.py の各機能を実行します。
python menu.py --detail で各項目の説明・出力先も表示します。

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
from typing import Any, Callable, Dict, List, Optional, Set, Tuple
from urllib.parse import urlparse

try:
    import yaml
except ImportError:
    yaml = None  # type: ignore

from console import BOLD, CYAN, DIM, GREEN, YELLOW, arrow, c, fail, hr, note, ok, warn
from inspect_downloaded import (
    export_search_hits_to_excel,
    extract_webhook_rows,
    fetch_and_save_webhooks,
    find_app_output_dir,
    find_webhook_file,
    format_search_hits,
    format_webhook_rows,
    is_stale_webhook_file,
    list_app_output_dirs,
    load_json_or_yaml,
    search_all_apps,
    search_downloaded,
)

SCRIPT_DIR = Path(__file__).resolve().parent
RUNNER = SCRIPT_DIR / "kintone_runner.py"
ENV_FILE = SCRIPT_DIR / ".kintone.env"
OUTPUT_DIR = SCRIPT_DIR / "output"
DETAIL_MODE = False
WIDTH = 70
MAX_HEADER_APPS = 5

PREFERRED_ENV_KEYS = [
    "subdomain",
    "username",
    "password",
    "user_domain",
    "app_tokens",
    "search_keywords",
    "domain_map",
    "js_dirs",
]

# 一括取得でダウンロードする主な情報（ヘルプ表示用）
DOWNLOAD_ITEMS = [
    "アプリ基本設定 / フォーム（フィールド・レイアウト） / 一覧（ビュー）",
    "アプリ・レコード・フィールドの ACL（アクセス権）",
    "通知（一般・レコード・リマインダー） / プロセス管理 / アクション / グラフ / プラグイン",
    "JavaScript / CSS カスタマイズ（ファイル本体を javascript/ に保存）",
]

GROUP_ACTIONS = [
    ("list", "グループ一覧を表示"),
    ("search", "ユーザーを検索"),
    ("add", "ユーザーをグループに追加"),
    ("remove", "ユーザーをグループから削除"),
]


# ---------------------------------------------------------------------------
# 設定ファイル（.kintone.env）の読み書き
# ---------------------------------------------------------------------------

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
        for key in ("subdomain", "username", "password"):
            if not config.get(key):
                preview["missing"].append(key)
        if not preview["app_ids"]:
            preview["missing"].append("app_tokens")
    except Exception as e:
        preview["error"] = str(e)
    return preview


def load_app_tokens_preview() -> List[str]:
    """.kintone.env から app_tokens のアプリ番号一覧を取得する。"""
    return list(load_env_preview().get("app_ids") or [])


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


def resolve_app_token(app_id: str) -> Optional[str]:
    """app_tokens から対象アプリのトークンを取り出す。"""
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
        warn(str(preview["error"]))
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
            warn(f"保存に失敗しました: {e}")
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
                        ok(f"アプリID {raw_id} のトークンを保存しました。")
                    except Exception as e:
                        warn(f"保存に失敗しました: {e}")
                        return False
                else:
                    print("  トークンが未入力のため、アプリは登録しませんでした。")

    return True


def ensure_app_token(app_id: str) -> bool:
    """対象アプリのトークンが無ければ入力して保存する。未入力なら False。"""
    if app_id in load_app_tokens_preview():
        return True
    warn(f"アプリID {app_id} の APIトークンが未登録です。この場で保存できます（未入力で中止）。")
    token = input("  APIトークン: ").strip()
    if not token:
        print("  トークンが未入力のため中止しました。")
        return False
    try:
        save_app_token(app_id, token)
    except Exception as e:
        warn(f"保存に失敗しました: {e}")
        return False
    ok(f"アプリID {app_id} のトークンを保存しました。続けて実行します。")
    return True


# ---------------------------------------------------------------------------
# 入力ヘルパー
# ---------------------------------------------------------------------------

def read_yes_no(prompt: str, default: bool = True) -> bool:
    hint = "Y/n" if default else "y/N"
    raw = input(f"  {prompt} [{hint}]: ").strip().lower()
    if not raw:
        return default
    return raw in ("y", "yes")


def read_choice(prompt: str, valid: Set[str], hint: str = "") -> str:
    while True:
        choice = input(prompt).strip()
        if choice in valid:
            return choice
        print(c(f"  {hint or '一覧にある番号を入力してください。'}", YELLOW))


def read_app_id(optional: bool = True) -> Optional[str]:
    hint = "（Enter で全アプリ）" if optional else "（必須）"
    while True:
        raw = input(f"  アプリID {hint}: ").strip()
        if not raw:
            return None if optional else ""
        if raw.isdigit():
            return raw
        print("  アプリID は数値で入力してください。")


def read_required_app_id() -> Optional[str]:
    raw = input("  アプリID: ").strip()
    if not raw:
        print("  キャンセルしました。")
        return None
    if not raw.isdigit():
        print("  アプリID は数値で入力してください。")
        return None
    return raw


def prompt_keep(label: str, current: Optional[str] = None, placeholder: str = "未設定") -> Optional[str]:
    """値を入力する。空Enterなら None（現状維持）。"""
    shown = current if current else placeholder
    raw = input(f"  {label} [Enterで維持: {shown}]: ").strip()
    return raw or None


def wait_enter() -> None:
    print()
    input(c("  Enter でメニューに戻る", DIM))


# ---------------------------------------------------------------------------
# 画面表示
# ---------------------------------------------------------------------------

def format_target_apps(preview: Dict[str, Any]) -> str:
    """ヘッダーの「対象アプリ」行の本文を組み立てる。"""
    app_ids = [str(app_id) for app_id in (preview.get("app_ids") or [])]
    if not app_ids:
        return f"なし → {menu_number('settings')} で登録"
    try:
        dirs = {app_id: path for app_id, path in list_app_output_dirs(OUTPUT_DIR)}
    except Exception:
        dirs = {}
    parts: List[str] = []
    for app_id in app_ids[:MAX_HEADER_APPS]:
        path = dirs.get(app_id)
        if path is None:
            parts.append(f"{app_id} (未取得)")
            continue
        name = path.name.split("_", 1)[1] if "_" in path.name else ""
        try:
            # Excel 変換でフォルダの更新日時が変わるため、取得時にだけ書かれる
            # settings.yaml の日時を優先する
            marker = path / f"{app_id}_settings.yaml"
            stat_target = marker if marker.exists() else path
            stamp = datetime.fromtimestamp(stat_target.stat().st_mtime).strftime("%m/%d %H:%M")
            fetched = f" (取得 {stamp})"
        except Exception:
            fetched = ""
        parts.append(f"{app_id} {name}{fetched}".replace("  ", " ").strip())
    text = f"{len(app_ids)} 件: " + ", ".join(parts)
    if len(app_ids) > MAX_HEADER_APPS:
        text += f" … 他 {len(app_ids) - MAX_HEADER_APPS} 件"
    return text


def print_header() -> None:
    """タイトルと現在の接続先・対象アプリ・設定ファイルを表示する。"""
    preview = load_env_preview()
    settings_no = menu_number("settings")
    print()
    print("=" * WIDTH)
    print(c("  kintone アプリ管理ツール", BOLD, CYAN))
    print("=" * WIDTH)

    if preview.get("error"):
        print(c(f"  設定の読み込みエラー: {preview['error']}", YELLOW))
    elif preview.get("subdomain"):
        user = f"  ({preview['username']})" if preview.get("username") else ""
        print(f"  接続先     https://{preview['subdomain']}.cybozu.com/{user}")
    else:
        print(c(f"  接続先     未設定 → {settings_no} で設定", YELLOW))

    apps_text = format_target_apps(preview)
    if preview.get("app_ids"):
        print(f"  対象アプリ {apps_text}")
    else:
        print(c(f"  対象アプリ {apps_text}", YELLOW))

    env_label = str(ENV_FILE) if DETAIL_MODE else ENV_FILE.name
    print(f"  設定ファイル {env_label}  " + c(f"… 変更は {settings_no}", DIM))


def build_menu() -> List[Tuple[str, Dict[str, Any]]]:
    """セクション定義から「番号, 項目」の一覧を作る（番号は通し）。"""
    entries: List[Tuple[str, Dict[str, Any]]] = []
    for section in MENU_SECTIONS:
        for item in section["items"]:
            entries.append((str(len(entries) + 1), item))
    return entries


def menu_number(key: str) -> str:
    """項目キーからメニュー番号を引く（見つからなければ '-'）。"""
    for number, item in build_menu():
        if item.get("key") == key:
            return number
    return "-"


def print_menu() -> List[str]:
    """セクション見出し付きでメニューを表示し、有効な番号を返す。"""
    entries = build_menu()
    numbers = iter(entries)
    print()
    for section in MENU_SECTIONS:
        print(c(f"  {section['title']}", BOLD))
        for _ in section["items"]:
            number, item = next(numbers)
            line = f"  {number:>2}. {item['label']}"
            print(c(line, BOLD, CYAN) if item.get("highlight") else line)
            if DETAIL_MODE:
                detail = item.get("description") or ""
                output = item.get("output")
                if output:
                    detail = f"{detail}  → {output}" if detail else f"→ {output}"
                if detail:
                    print(c(f"      {detail}", DIM))
    print("   0. 終了")
    print()
    return [number for number, _ in entries]


def print_env_sample(indent: str = "  ") -> None:
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
        "",
        "# 任意。ドメイン変更の as-is / to-be 確認用（変更前ホスト: 変更後ホスト）",
        "domain_map:",
        "  old.example.com: \"new.example.com\"",
    ]
    print(c(f"{indent}----- {ENV_FILE.name} の記入例 -----", YELLOW))
    for line in sample:
        print(c(f"{indent}{line}", GREEN))
    print(c(f"{indent}--------------------------------", YELLOW))


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


def print_current_settings() -> None:
    """現在の設定内容を表示する（トークンの値は表示しない）。"""
    preview = load_env_preview()
    print()
    print(c("  現在の設定", BOLD, CYAN))
    print(f"  設定ファイル     : {ENV_FILE}")
    if preview.get("error"):
        warn(f"読み込みエラー: {preview['error']}")
        return
    if not preview.get("exists"):
        warn(f"{ENV_FILE.name} がありません。")
        return
    print_connection_fields(preview)
    keywords = preview.get("search_keywords") or []
    print(f"  検索文字列       : {', '.join(keywords) if keywords else '（なし）'}")
    domain_map = load_domain_map()
    if domain_map:
        pairs = ", ".join(f"{old} → {new}" for old, new in domain_map.items())
        print(f"  ドメイン対応表   : {len(domain_map)} 件  {pairs}")
    else:
        print("  ドメイン対応表   : （なし）")
    if preview.get("user_domain"):
        print(f"  user_domain      : {preview['user_domain']}")
    if preview.get("js_dirs"):
        print(f"  js_dirs          : {', '.join(preview['js_dirs'])}")
    note("APIトークンの値は表示しません。")


def print_token_permission_guide(app_id: Optional[str] = None) -> None:
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


# ---------------------------------------------------------------------------
# 設定の対話
# ---------------------------------------------------------------------------

def set_connection_interactive() -> None:
    """接続先 URL / username / password を .kintone.env に保存する。"""
    print()
    print(c("  接続先とログインの設定", BOLD, CYAN))
    print(f"  保存先: {ENV_FILE}")
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

    updates: Dict[str, Any] = {}
    if subdomain:
        updates["subdomain"] = subdomain
    if username:
        updates["username"] = username
    if password:
        updates["password"] = password

    try:
        merge_save_env(updates)
    except Exception as e:
        warn(f"保存に失敗しました: {e}")
        return

    print()
    ok("保存しました。現在の接続情報:")
    print_connection_fields(load_env_preview())


def add_app_token_interactive() -> None:
    """アプリIDとAPIトークンを入力し、.kintone.env の app_tokens に追加／上書きする。"""
    print()
    print(c("  アプリID と APIトークンの追加・更新", BOLD, CYAN))
    print(f"  保存先: {ENV_FILE}")
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
        warn(f"保存に失敗しました: {e}")
        return

    if overwritten:
        ok(f"アプリID {app_id} のトークンを上書きして保存しました。")
    else:
        ok(f"アプリID {app_id} のトークンを追加して保存しました。")


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
        print(c(f"  保存済みの検索文字列: {len(words)} 件", CYAN))
        for i, word in enumerate(words, start=1):
            print(f"    {i}. {word}")
    else:
        print(c("  保存済みの検索文字列: （なし）", DIM))


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


def manage_search_keywords_interactive() -> None:
    print()
    print(c("  検索文字列の設定", BOLD, CYAN))
    print("  取得済みデータの検索で使う文字列を追加・削除します。")
    while True:
        print()
        print_search_keywords()
        print()
        print("   1. 追加")
        print("   2. 削除")
        print("   0. 戻る")
        choice = read_choice("  番号を選択 [0-2]: ", {"0", "1", "2"}, "0 から 2 の番号を入力してください。")
        if choice == "0":
            return
        if choice == "1":
            added = read_keyword_lines("追加する検索文字列を入力してください。")
            if not added:
                print("  追加はありません。")
                continue
            save_search_keywords(load_search_keywords() + added)
            ok(f"{len(added)} 件を追加して保存しました。")
            continue
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
        save_search_keywords([w for i, w in enumerate(current) if i not in remove_idx])
        ok(f"{len(remove_idx)} 件を削除して保存しました。")


def _normalize_host(value: Any) -> str:
    """"https://old.example.com/path" のような入力からホスト名だけを取り出す。"""
    text = str(value or "").strip()
    if not text:
        return ""
    text = re.sub(r"^[A-Za-z][A-Za-z0-9+.\-]*:", "", text).lstrip("/")
    text = re.split(r"[/?#]", text, 1)[0].split("@")[-1].split(":", 1)[0]
    return text.strip().strip(".").lower()


def _normalize_domain_map(raw: Any) -> Dict[str, str]:
    """domain_map（変更前ホスト: 変更後ホスト）を {小文字ホスト: 小文字ホスト} に整える。

    url_inventory.normalize_domain_map と同じ結果になるようにした簡易版
    （menu.py を url_inventory.py に依存させないため、ここで自前実装する）。
    """
    if not isinstance(raw, dict):
        return {}
    result: Dict[str, str] = {}
    for key, value in raw.items():
        old, new = _normalize_host(key), _normalize_host(value)
        if old and new:
            result[old] = new
    return result


def load_domain_map() -> Dict[str, str]:
    """.kintone.env の domain_map を正規化して返す（無ければ空辞書）。"""
    try:
        return _normalize_domain_map(load_env_config_raw().get("domain_map"))
    except Exception:
        return {}


def save_domain_map(mapping: Dict[str, str]) -> None:
    merge_save_env({"domain_map": dict(mapping)})


def format_domain_map(mapping: Dict[str, str], limit: int = 3) -> str:
    """「old → new」を limit 件まで並べた 1 行の文字列を返す。"""
    items = list(mapping.items())
    text = ", ".join(f"{old} → {new}" for old, new in items[:limit])
    if len(items) > limit:
        text += f" … 他 {len(items) - limit} 件"
    return text


def print_domain_map(mapping: Optional[Dict[str, str]] = None) -> None:
    items = list((load_domain_map() if mapping is None else mapping).items())
    if items:
        print(c(f"  登録済みのドメイン対応: {len(items)} 件", CYAN))
        for i, (old, new) in enumerate(items, start=1):
            print(f"    {i}. {old} → {new}")
    else:
        print(c("  登録済みのドメイン対応: （なし）", DIM))


def manage_domain_map_interactive() -> None:
    print()
    print(c("  ドメイン対応表の設定", BOLD, CYAN))
    print("  URL 一覧の「変更後 URL」「状態」に使う 変更前ホスト → 変更後ホスト を登録します。")
    while True:
        print()
        print_domain_map()
        print()
        print("   1. 追加")
        print("   2. 削除")
        print("   0. 戻る")
        choice = read_choice("  番号を選択 [0-2]: ", {"0", "1", "2"}, "0 から 2 の番号を入力してください。")
        if choice == "0":
            return
        if choice == "1":
            old = input("  変更前ホスト（例: old.example.com）: ").strip()
            if not old:
                print("  キャンセルしました。")
                continue
            new = input("  変更後ホスト（例: new.example.com）: ").strip()
            if not new:
                print("  キャンセルしました。")
                continue
            pair = _normalize_domain_map({old: new})
            if not pair:
                print("  ホスト名を正しく入力してください。")
                continue
            mapping = load_domain_map()
            mapping.update(pair)
            save_domain_map(mapping)
            for old_host, new_host in pair.items():
                ok(f"{old_host} → {new_host} を保存しました。")
            continue
        current = list(load_domain_map().items())
        if not current:
            print("  削除する対応がありません。")
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
        save_domain_map({old: new for i, (old, new) in enumerate(current) if i not in remove_idx})
        ok(f"{len(remove_idx)} 件を削除して保存しました。")


# ---------------------------------------------------------------------------
# 取得済みデータの閲覧・検索
# ---------------------------------------------------------------------------

def resolve_app_output_dir(app_id: str) -> Optional[Path]:
    app_dir = find_app_output_dir(OUTPUT_DIR, app_id)
    if app_dir:
        return app_dir
    warn(f"アプリID {app_id} の取得済みフォルダが見つかりません。")
    print(f"  先に {menu_number('download')} でダウンロードしてください。")
    return None


def show_webhooks_interactive() -> None:
    print()
    print(c("  Webhook 一覧", BOLD, CYAN))
    app_id = read_required_app_id()
    if not app_id:
        return
    app_dir = resolve_app_output_dir(app_id)
    if not app_dir:
        return
    if DETAIL_MODE:
        note(f"参照: {app_dir}")
    webhook_file = find_webhook_file(app_dir, app_id)
    data = None
    if webhook_file:
        try:
            data = load_json_or_yaml(webhook_file)
            note(f"ファイル: {webhook_file.relative_to(app_dir)}")
        except Exception as e:
            warn(f"読み込みエラー: {e}")
            webhook_file = None
        if data is not None and is_stale_webhook_file(data):
            warn("保存済みファイルは旧方式で取得したものです。取得し直します。")
            data = None

    if data is None:
        print(f"  管理画面と同じ内部API（/k/api/dev/app/{app_id}/webhook/list.json）で取得します。")
        preview = load_env_preview()
        if preview.get("subdomain"):
            note(f"https://{preview['subdomain']}.cybozu.com/k/admin/app/webhook?app={app_id}")
        try:
            config = load_env_config_raw()
        except Exception as e:
            warn(f"設定の読み込みに失敗しました: {e}")
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
            ok(f"取得して保存しました: {saved.relative_to(app_dir)}")
        else:
            warn("Webhook一覧APIからは取得できませんでした。")
            if error:
                note(str(error))
            print("  取得済みファイル内の webhook という文字列を代わりに表示します。")
            hits = search_downloaded(app_dir, ["webhook", "Webhook", "WEBHOOK"])
            if hits:
                for line in format_search_hits(["webhook"], hits, detail=DETAIL_MODE):
                    print(line)
            else:
                print("  取得済みデータ内にも webhook の記載はありません（0件）。")
            return

    print()
    for line in format_webhook_rows(extract_webhook_rows(data)):
        print(line)


def search_downloaded_interactive() -> None:
    print()
    print(c("  取得済みデータを検索（JavaScript / YAML / JSON など）", BOLD, CYAN))
    app_id = read_app_id(optional=True)
    if app_id:
        app_dir = resolve_app_output_dir(app_id)
        if not app_dir:
            return
        targets: List[Tuple[str, Path]] = [(app_id, app_dir)]
    else:
        targets = list_app_output_dirs(OUTPUT_DIR)
        if not targets:
            warn("取得済みフォルダがありません。")
            print(f"  先に {menu_number('download')} でダウンロードしてください。")
            return
        ids = ", ".join(aid for aid, _ in targets)
        print(c(f"  対象: {len(targets)} アプリ  [{ids}]", CYAN))
    if DETAIL_MODE:
        for _, target_dir in targets:
            note(f"参照: {target_dir}")

    saved = load_search_keywords()
    print()
    print(f"  保存済みの検索文字列: {', '.join(saved) if saved else '（なし）'}")
    options: List[str] = []
    valid = {"0", "2", "3"}
    if saved:
        options.append("1. 保存済みの文字列で検索")
        valid.add("1")
    options.append("2. 文字列を入力して検索（保存しない）")
    options.append("3. 保存済みに追加してから検索")
    options.append("0. 戻る")
    print("  " + "   ".join(options))
    choice = read_choice("  選択: ", valid, "表示されている番号を入力してください。")

    if choice == "0":
        print("  キャンセルしました。")
        return
    if choice == "1":
        keywords = saved
    elif choice == "2":
        keywords = read_keyword_lines("今回だけ使う検索文字列を入力してください（保存しません）。")
    else:
        added = read_keyword_lines("保存済みに追加する検索文字列を入力してください。")
        if added:
            save_search_keywords(saved + added)
            ok(f"{len(added)} 件を追加して保存しました。")
        keywords = load_search_keywords()
    if not keywords:
        print("  検索する文字列がありません。")
        return

    write_excel = read_yes_no("Excel に出力しますか？", default=True)
    hits = search_all_apps(targets, keywords)
    print()
    for line in format_search_hits(keywords, hits, multi_app=len(targets) > 1, detail=DETAIL_MODE):
        print(line)
    if write_excel:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = OUTPUT_DIR / f"search_hits_{timestamp}.xlsx"
        try:
            export_search_hits_to_excel(keywords, hits, excel_path)
            arrow(f"Excel: {excel_path}")
        except Exception as e:
            fail(f"Excel の出力に失敗しました: {e}")


# ---------------------------------------------------------------------------
# kintone_runner.py の実行
# ---------------------------------------------------------------------------

def run_runner(runner_args: List[str], env_path: Optional[Path] = None) -> int:
    cmd = [sys.executable, str(RUNNER)]
    if env_path:
        cmd.extend(["--env", str(env_path)])
    if DETAIL_MODE:
        cmd.append("--verbose")
    cmd.extend(runner_args)
    print()
    print(c(f"  実行: python {RUNNER.name} {' '.join(cmd[2:])}", DIM))
    print(hr(WIDTH))
    sys.stdout.flush()
    result = subprocess.run(cmd, cwd=str(SCRIPT_DIR))
    print(hr(WIDTH))
    return result.returncode


def run_app_with_permission_retry(runner_args: List[str]) -> int:
    """アプリ取得を実行し、権限不足なら案内のあと再試行できる。"""
    app_id = None
    if "--id" in runner_args:
        idx = runner_args.index("--id")
        if idx + 1 < len(runner_args):
            app_id = runner_args[idx + 1]
    while True:
        code = run_runner(runner_args)
        if code == 0:
            return 0
        if code != 3:
            return code
        print_token_permission_guide(app_id)
        if not read_yes_no("権限を追加してアプリを更新したら、再試行しますか？", default=True):
            return code


# ---------------------------------------------------------------------------
# 各メニュー項目の処理
# ---------------------------------------------------------------------------

def print_app_run_confirm() -> None:
    """アプリ一括取得の実行前に概要を 3 行で表示する。"""
    app_ids = load_app_tokens_preview()
    print()
    print(c("  アプリ設定・フォーム・ACL・通知・プロセス・ビュー・カスタマイズJS を保存します。", CYAN))
    if app_ids:
        print(f"  対象アプリ: {', '.join(app_ids)}")
    else:
        warn(f"{ENV_FILE.name} の app_tokens が空です。アプリIDとトークンを登録してください。")
    arrow("output/[アプリID]_[アプリ名]/")


def build_app_args() -> Optional[List[str]]:
    print_app_run_confirm()
    app_id = read_app_id(optional=True)
    if app_id:
        if not ensure_app_token(app_id):
            return None
        return ["app", "--id", app_id]

    registered = load_app_tokens_preview()
    if registered:
        print(c(f"  全 {len(registered)} 件のアプリを処理します: {', '.join(registered)}", CYAN))
        return ["app"]

    warn("app_tokens が空のため、先に 1 件登録してください。")
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


def build_app_id_args(command: str) -> Optional[List[str]]:
    """アプリID を尋ねて runner 引数を組み立てる（Enter で全件）。"""
    print()
    app_id = read_app_id(optional=True)
    args = [command]
    if app_id:
        if not ensure_app_token(app_id):
            return None
        args.extend(["--id", app_id])
    elif not load_app_tokens_preview():
        warn("app_tokens が空です。先にアプリIDとトークンを登録してください。")
        if not fill_missing_fields(required=["app_tokens"]) or not load_app_tokens_preview():
            return None
    return args


def build_group_args() -> Optional[List[str]]:
    print()
    print(c("  グループ操作", BOLD, CYAN))
    for i, (action, desc) in enumerate(GROUP_ACTIONS, start=1):
        print(f"  {i:>2}. {desc} ({action})")
    print("   0. 戻る")
    valid = {str(i) for i in range(len(GROUP_ACTIONS) + 1)}
    choice = read_choice(
        f"  番号を選択 [0-{len(GROUP_ACTIONS)}]: ",
        valid,
        f"0 から {len(GROUP_ACTIONS)} の番号を入力してください。",
    )
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
    print("   1. 全アプリ")
    print("   2. 指定したアプリIDのみ")
    print("   3. 指定したアプリIDを除外")
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
        warn("app_tokens が空です。先にアプリIDとトークンを登録してください。")
        if not fill_missing_fields(required=["app_tokens"]) or not load_app_tokens_preview():
            return None
    return args


def handle_download() -> Optional[int]:
    args = build_app_args()
    if args is None:
        return None
    return run_app_with_permission_retry(args)


def make_app_id_handler(command: str) -> Callable[[], Optional[int]]:
    def handler() -> Optional[int]:
        args = build_app_id_args(command)
        if args is None:
            return None
        return run_runner(args)

    return handler


def handle_summary() -> Optional[int]:
    return run_runner(["summary"])


def handle_search() -> Optional[int]:
    search_downloaded_interactive()
    return None


def handle_webhooks() -> Optional[int]:
    show_webhooks_interactive()
    return None


def handle_urls() -> Optional[int]:
    print()
    domain_map = load_domain_map()
    if domain_map:
        print(c(f"  ドメイン対応表: {len(domain_map)} 件  {format_domain_map(domain_map)}", CYAN))
    else:
        warn(
            f"ドメイン対応表は未設定。{menu_number('settings')} 設定 → ドメイン対応表 で登録すると"
            "変更後 URL と状態が出ます（登録しなくても一覧は出せます）"
        )
    app_id = read_app_id(optional=True)
    return run_runner(["urls"] + (["--id", app_id] if app_id else []))


def handle_users() -> Optional[int]:
    print()
    raw = input("  出力形式 [Enter=Excel / c=CSV]: ").strip().lower()
    if raw in ("c", "csv"):
        return run_runner(["users", "--format", "csv"])
    return run_runner(["users"])


def handle_group() -> Optional[int]:
    args = build_group_args()
    if args is None:
        return None
    return run_runner(args)


def handle_all() -> Optional[int]:
    args = build_all_args()
    if args is None:
        return None
    warn("output/ の内容を previous_output/ に退避してから実行します。")
    return run_app_with_permission_retry(args)


SETTINGS_ACTIONS: List[Tuple[str, Callable[[], None]]] = []


def handle_settings() -> Optional[int]:
    valid = {str(i) for i in range(len(SETTINGS_ACTIONS) + 1)}
    while True:
        print()
        print(c("  設定", BOLD))
        for i, (label, _) in enumerate(SETTINGS_ACTIONS, start=1):
            print(f"  {i:>2}. {label}")
        print("   0. 戻る")
        choice = read_choice(
            f"  番号を選択 [0-{len(SETTINGS_ACTIONS)}]: ",
            valid,
            f"0 から {len(SETTINGS_ACTIONS)} の番号を入力してください。",
        )
        if choice == "0":
            return None
        SETTINGS_ACTIONS[int(choice) - 1][1]()


def handle_help() -> Optional[int]:
    print()
    print("=" * WIDTH)
    print(c("  ヘルプ", BOLD, CYAN))
    print("=" * WIDTH)

    print()
    print(c("  基本の流れ", BOLD))
    print(
        f"    {menu_number('settings')} で接続先とアプリを登録"
        f" → {menu_number('download')} でダウンロード"
        f" → {menu_number('acl')}〜{menu_number('summary')} で Excel 化"
        f" → {menu_number('search')}・{menu_number('webhooks')} で確認"
    )
    print(
        f"    ドメイン変更の確認: 設定 → ドメイン対応表 に登録"
        f" → {menu_number('urls')} で一覧（要変更を確認）"
        f" → 変更後に {menu_number('download')} で再取得"
        f" → {menu_number('urls')} で「要変更 0 件」を確認"
    )
    print("  一括取得で保存される情報:")
    for item in DOWNLOAD_ITEMS:
        print(f"    ・{item}")

    print()
    print(c(f"  設定ファイル（{ENV_FILE}）", BOLD))
    print_env_sample(indent="    ")
    print("    subdomain / username / password / app_tokens が必須です。")
    print("    パスワードとトークンが入るため、Git にコミットしないでください。")

    print()
    print(c("  APIトークンの用意", BOLD))
    print("    1. kintone で対象アプリを開く")
    print("    2. 「アプリの設定」→「設定」タブ →「APIトークン」")
    print("    3. 生成し、[必須] レコード閲覧 / [推奨] アプリ管理 の権限を付ける")
    print("    4. 画面右上の「アプリを更新」で反映する")
    print("    5. .kintone.env の app_tokens に <アプリID>: \"<発行したトークン>\" を追加")

    print()
    print(c("  出力ファイル", BOLD))
    try:
        sys.stdout.flush()
        subprocess.run([sys.executable, str(RUNNER), "outputs"], cwd=str(SCRIPT_DIR))
    except Exception as e:
        warn(f"出力ファイル一覧を取得できませんでした: {e}")

    print()
    print(c("  CLI から直接使う場合", BOLD))
    print(f"    python {RUNNER.name} <command> ...")
    print(f"    python {RUNNER.name} --help")
    return None


MENU_SECTIONS: List[Dict[str, Any]] = [
    {
        "title": "取得",
        "items": [
            {
                "key": "download",
                "label": "アプリ設定・JS をダウンロード",
                "description": "app_tokens のアプリから設定・フォーム・ACL・通知・カスタマイズJS を取得",
                "output": "output/[アプリID]_[アプリ名]/",
                "handler": handle_download,
                "highlight": True,
            },
        ],
    },
    {
        "title": "Excel 変換（1 で取得したデータから作成）",
        "items": [
            {
                "key": "acl",
                "label": "アクセス権 (ACL)",
                "description": "アプリ・レコード・フィールドのアクセス権を Excel に変換",
                "output": "output/[アプリID]_[アプリ名]/[アプリID]_acl_report.xlsx",
                "handler": make_app_id_handler("acl"),
            },
            {
                "key": "notifications",
                "label": "通知設定",
                "description": "一般・レコード・リマインダー通知を Excel に変換",
                "output": "output/[アプリID]_[アプリ名]/[アプリID]_notifications.xlsx",
                "handler": make_app_id_handler("notifications"),
            },
            {
                "key": "process_workflow",
                "label": "プロセス管理",
                "description": "プロセス管理（ステータスと作業者）を Excel に変換",
                "output": "output/[アプリID]_[アプリ名]/[アプリID]_process_workflow.xlsx",
                "handler": make_app_id_handler("process_workflow"),
            },
            {
                "key": "summary",
                "label": "アプリ設定一覧表（全アプリ）",
                "description": "取得済みアプリ設定の全体一覧を 1 つの Excel に出力",
                "output": "output/kintone_app_settings_summary_[日時].xlsx",
                "handler": handle_summary,
            },
        ],
    },
    {
        "title": "閲覧・検索",
        "items": [
            {
                "key": "search",
                "label": "取得済みデータを検索（JS 含む）",
                "description": "YAML / JSON / JavaScript を文字列検索し、ヒット一覧を表示",
                "output": "コンソール + output/search_hits_[日時].xlsx",
                "handler": handle_search,
            },
            {
                "key": "webhooks",
                "label": "Webhook 一覧",
                "description": "取得済みアプリの Webhook 設定を一覧表示（無ければ API で取得）",
                "output": "コンソール",
                "handler": handle_webhooks,
            },
            {
                "key": "urls",
                "label": "URL 一覧を Excel 出力（Webhook / JS 呼び出し / 外部参照）",
                "description": "全アプリの Webhook・JS の HTTP 呼び出し・外部参照 URL を横断して一覧化。domain_map があれば変更後 URL と状態も出力",
                "output": "output/url_inventory_[日時].xlsx",
                "handler": handle_urls,
            },
        ],
    },
    {
        "title": "ユーザー・グループ",
        "items": [
            {
                "key": "users",
                "label": "ユーザー一覧を Excel 出力",
                "description": "全ユーザーと所属グループを出力（username / password で認証）",
                "output": "output/kintone_users_groups_[日時].xlsx",
                "handler": handle_users,
            },
            {
                "key": "group",
                "label": "グループ操作（一覧・検索・追加・削除）",
                "description": "グループ一覧・ユーザー検索・グループへの追加と削除",
                "output": "コンソール",
                "handler": handle_group,
            },
        ],
    },
    {
        "title": "その他",
        "items": [
            {
                "key": "all",
                "label": "まとめて実行（ユーザー一覧 → ダウンロード → Excel 変換）",
                "description": "users → app → acl → summary → notifications → process_workflow",
                "output": "上記すべての出力ファイル",
                "handler": handle_all,
            },
            {
                "key": "settings",
                "label": "設定（接続先 / アプリとトークン / 検索文字列）",
                "description": ".kintone.env の接続先・アプリとトークン・検索文字列を編集",
                "output": ".kintone.env",
                "handler": handle_settings,
            },
            {
                "key": "help",
                "label": "ヘルプ（使い方・設定ファイル・出力ファイル）",
                "description": "基本の流れ・設定ファイルの書き方・出力ファイル・CLI の使い方",
                "output": "コンソール",
                "handler": handle_help,
            },
        ],
    },
]

SETTINGS_ACTIONS.extend([
    ("接続先とログイン（URL / username / password）", set_connection_interactive),
    ("アプリとトークンを追加・更新", add_app_token_interactive),
    ("検索文字列", manage_search_keywords_interactive),
    ("ドメイン対応表（変更前ホスト → 変更後ホスト）", manage_domain_map_interactive),
    ("現在の設定を表示（トークンは伏せる）", print_current_settings),
])


# ---------------------------------------------------------------------------
# 起動
# ---------------------------------------------------------------------------

def initial_setup() -> None:
    """接続情報が足りないときだけ、起動時に 1 回だけ入力を促す。"""
    preview = load_env_preview()
    if preview.get("error"):
        print_header()
        warn(str(preview["error"]))
        return
    missing = [key for key in ("subdomain", "username", "password") if _is_field_missing(preview, key)]
    if not missing:
        return
    print_header()
    print()
    warn("接続情報が未設定です。")
    note(f"Enter でスキップできます。あとから {menu_number('settings')} の設定でも変更できます。")
    fill_missing_fields(required=["subdomain", "username", "password"])


def parse_args(argv: Optional[List[str]] = None):
    parser = argparse.ArgumentParser(description="kintone アプリ管理ツールの対話型メニュー")
    parser.add_argument(
        "--detail",
        action="store_true",
        help="各項目の説明と出力先も表示し、runner を --verbose で実行する",
    )
    return parser.parse_args(argv)


def main() -> None:
    global DETAIL_MODE
    args = parse_args()
    DETAIL_MODE = args.detail

    if not RUNNER.exists():
        print(f"エラー: {RUNNER} が見つかりません。")
        sys.exit(1)

    initial_setup()

    while True:
        print_header()
        numbers = print_menu()
        last = numbers[-1] if numbers else "0"
        choice = read_choice(
            f"番号を選択 [0-{last}]: ",
            set(numbers) | {"0"},
            f"0 から {last} の番号を入力してください。",
        )
        if choice == "0":
            print("終了します。")
            return

        item = dict(build_menu())[choice]
        code = item["handler"]()
        if code not in (None, 0):
            fail(f"失敗（終了コード {code}）。詳細は logs/ の最新ログを確認")
        wait_enter()


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):
        print()
        print("終了します。")
