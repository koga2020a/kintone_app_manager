#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""取得済み output ディレクトリから Webhook と全文検索を行う。"""

from __future__ import annotations

import base64
import json
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests
import yaml

SEARCH_SUFFIXES = {
    ".js",
    ".css",
    ".json",
    ".yaml",
    ".yml",
    ".txt",
    ".tsv",
    ".csv",
    ".html",
    ".md",
}

KIND_ORDER = ["JavaScript", "CSS", "設定JSON", "設定YAML", "その他テキスト"]


def find_app_output_dir(output_dir: Path, app_id: str) -> Optional[Path]:
    """output 内の {app_id}_* ディレクトリのうち、更新が新しいものを返す。"""
    if not output_dir.exists():
        return None
    matches = [
        path
        for path in output_dir.iterdir()
        if path.is_dir() and path.name.startswith(f"{app_id}_")
    ]
    if not matches:
        return None
    return max(matches, key=lambda path: path.stat().st_mtime)


def classify_file(app_dir: Path, file_path: Path) -> str:
    rel = file_path.relative_to(app_dir).as_posix()
    suffix = file_path.suffix.lower()
    if "javascript/" in rel or suffix == ".js":
        return "JavaScript"
    if suffix == ".css":
        return "CSS"
    if suffix in {".json"}:
        return "設定JSON"
    if suffix in {".yaml", ".yml"}:
        return "設定YAML"
    return "その他テキスト"


def load_json_or_yaml(path: Path) -> Any:
    text = path.read_text(encoding="utf-8")
    if path.suffix.lower() in {".yaml", ".yml"}:
        import yaml

        return yaml.safe_load(text)
    return json.loads(text)


def find_webhook_file(app_dir: Path, app_id: str) -> Optional[Path]:
    names = [
        f"{app_id}_webhooks.json",
        f"{app_id}_webhooks.yaml",
        f"{app_id}_webhook.json",
        f"{app_id}_webhook.yaml",
        "webhooks.json",
        "webhooks.yaml",
    ]
    folders = [app_dir / "json", app_dir]
    for folder in folders:
        if not folder.exists():
            continue
        for name in names:
            path = folder / name
            if path.exists():
                return path
    return None


def save_webhook_data(app_dir: Path, app_id: str, data: Any) -> Path:
    json_dir = app_dir / "json"
    json_dir.mkdir(parents=True, exist_ok=True)
    json_path = json_dir / f"{app_id}_webhooks.json"
    yaml_path = app_dir / f"{app_id}_webhooks.yaml"
    json_path.write_text(json.dumps(data, ensure_ascii=False, indent=4), encoding="utf-8")
    yaml_path.write_text(yaml.dump(data, allow_unicode=True), encoding="utf-8")
    return json_path


def fetch_webhooks_via_admin(
    subdomain: str,
    app_id: str,
    username: str,
    password: str,
) -> Tuple[Optional[Any], Optional[str]]:
    """管理画面と同じ内部APIで Webhook 一覧を取る。

    公開 REST API には Webhook 一覧を取得する API が無いため、
    管理画面 (/k/admin/app/webhook?app=N) が内部で使っている
    /k/api/dev/app/{app_id}/webhook/list.json を直接叩く。
    パスワード認証ヘッダのみで動作し、ログインセッションや
    リクエストトークンは不要。
    """
    path = f"/k/api/dev/app/{app_id}/webhook/list.json"
    url = f"https://{subdomain}.cybozu.com{path}"
    encoded = base64.b64encode(f"{username}:{password}".encode()).decode()
    headers = {
        "X-Cybozu-Authorization": encoded,
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/json",
    }
    try:
        response = requests.post(url, headers=headers, json={}, timeout=30)
        if response.status_code != 200:
            return None, f"{path} -> {response.status_code}"
        data = response.json()
        if isinstance(data, dict) and (data.get("success") or data.get("result") is not None):
            result = data.get("result")
            if result is not None:
                return result, None
        return data, None
    except Exception as e:
        return None, str(e)


def fetch_and_save_webhooks(
    app_dir: Path,
    app_id: str,
    subdomain: str,
    api_token: Optional[str] = None,
    username: Optional[str] = None,
    password: Optional[str] = None,
) -> Tuple[Optional[Path], Optional[Any], Optional[str]]:
    """Webhook一覧を取得して output に保存する。

    api_token は互換のため残しているが、この内部APIでは使用しない
    (API トークン認証ではアクセスできない)。
    """
    if not username or not password:
        return (
            None,
            None,
            "Webhook一覧の取得には .kintone.env の username / password が必要です",
        )

    data, error = fetch_webhooks_via_admin(subdomain, app_id, username, password)
    if data is not None:
        saved = save_webhook_data(app_dir, app_id, data)
        return saved, data, None
    return None, None, error or "Webhook一覧を取得できませんでした"


def _person_name(person: Any) -> str:
    if isinstance(person, dict):
        return str(person.get("name") or person.get("code") or "")
    if person is None:
        return ""
    return str(person)


def extract_webhook_rows(data: Any) -> List[Dict[str, Any]]:
    """内部API のレスポンス（result 部）から表示用の行を作る。"""
    if data is None:
        return []
    if isinstance(data, dict):
        result = data.get("result")
        if isinstance(result, (dict, list)) and not data.get("webhooks"):
            return extract_webhook_rows(result)
        items = data.get("webhooks") or data.get("webhook") or []
        if isinstance(items, dict):
            items = [items]
        if not items and any(key in data for key in ("url", "localId", "webhookId", "id")):
            items = [data]
    elif isinstance(data, list):
        items = data
    else:
        return []

    rows = []
    for item in items:
        if not isinstance(item, dict):
            continue
        events = item.get("types") or item.get("events") or item.get("event") or []
        if isinstance(events, list):
            events_text = ", ".join(str(ev) for ev in events)
        else:
            events_text = str(events)
        headers = item.get("headers") or []
        if isinstance(headers, list):
            header_text = ", ".join(
                f"{h.get('name', '')}={h.get('value', '')}" if isinstance(h, dict) else str(h)
                for h in headers
            )
        else:
            header_text = str(headers)
        rows.append(
            {
                "id": item.get("localId") or item.get("webhookId") or item.get("id") or "",
                "name": item.get("description") or item.get("name") or "",
                "url": item.get("url") or "",
                "events": events_text,
                "enabled": item.get("enabled", ""),
                "creator": _person_name(item.get("creator")),
                "modifier": _person_name(item.get("modifier")),
                "created_at": item.get("createdAt") or "",
                "modified_at": item.get("modifiedAt") or "",
                "headers": header_text,
            }
        )
    return rows


def is_stale_webhook_file(data: Any) -> bool:
    """旧方式（管理画面HTMLの走査）で保存されたデータかどうか。"""
    if isinstance(data, dict):
        source = data.get("source")
        if isinstance(source, str) and "/k/admin/app/webhook" in source:
            return True
        candidates = [data]
        result = data.get("result")
        if isinstance(result, dict):
            candidates.append(result)
        for candidate in candidates:
            items = candidate.get("webhooks") or candidate.get("webhook") or []
            if isinstance(items, dict):
                items = [items]
            if isinstance(items, list):
                for item in items:
                    if isinstance(item, dict) and item.get("source") == "admin_html":
                        return True
    elif isinstance(data, list):
        for item in data:
            if isinstance(item, dict) and item.get("source") == "admin_html":
                return True
    return False


def search_downloaded(
    app_dir: Path, keywords: Iterable[str]
) -> List[Dict[str, Any]]:
    """キーワードのいずれかにヒットした行を返す。"""
    cleaned = [kw for kw in (k.strip() for k in keywords) if kw]
    if not cleaned:
        return []

    hits: List[Dict[str, Any]] = []
    for file_path in sorted(app_dir.rglob("*")):
        if not file_path.is_file():
            continue
        if file_path.suffix.lower() not in SEARCH_SUFFIXES:
            continue
        try:
            lines = file_path.read_text(encoding="utf-8").splitlines()
        except (UnicodeDecodeError, OSError):
            continue
        kind = classify_file(app_dir, file_path)
        rel = file_path.relative_to(app_dir).as_posix()
        for line_no, line in enumerate(lines, start=1):
            matched = [kw for kw in cleaned if kw.lower() in line.lower()]
            if not matched:
                continue
            hits.append(
                {
                    "kind": kind,
                    "file": rel,
                    "line": line_no,
                    "text": line.strip(),
                    "keywords": matched,
                }
            )
    return hits


def summarize_hits(hits: List[Dict[str, Any]]) -> List[Tuple[str, int, int]]:
    """種別ごとのファイル数とヒット行数。"""
    files_by_kind: Dict[str, set] = {kind: set() for kind in KIND_ORDER}
    lines_by_kind: Dict[str, int] = {kind: 0 for kind in KIND_ORDER}
    for hit in hits:
        kind = hit["kind"]
        files_by_kind.setdefault(kind, set()).add(hit["file"])
        lines_by_kind[kind] = lines_by_kind.get(kind, 0) + 1
    summary = []
    for kind in KIND_ORDER:
        if lines_by_kind.get(kind):
            summary.append((kind, len(files_by_kind[kind]), lines_by_kind[kind]))
    return summary
