#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""kintone 管理画面が使う内部API（/k/api/dev/app/{id}/...）の共通取得。"""

from __future__ import annotations

import base64
from typing import Any, Optional, Tuple

import requests

# 管理画面と同じパス。公開 REST に無い／不足しがちなもの。
# (resource, 保存名)
DEV_APP_RESOURCES = [
    ("webhook/list.json", "webhooks"),
    ("plugin/list.json", "dev_plugins"),
    ("customize/get.json", "dev_customize"),
    ("action/list.json", "dev_actions"),
    ("view/list.json", "dev_views"),
    ("graph/list.json", "dev_graphs"),
    ("notification/general/get.json", "dev_app_notifications"),
    ("notification/record/get.json", "dev_record_notifications"),
    ("notification/reminder/get.json", "dev_reminder_notifications"),
    ("acl/app/get.json", "dev_app_acl"),
    ("acl/record/get.json", "dev_record_acl"),
    ("acl/field/get.json", "dev_field_acl"),
    ("status/get.json", "dev_process_management"),
    ("form/get.json", "dev_form"),
    ("settings/get.json", "dev_settings"),
]


def password_headers(username: str, password: str) -> dict:
    encoded = base64.b64encode(f"{username}:{password}".encode()).decode()
    return {
        "X-Cybozu-Authorization": encoded,
        "X-Requested-With": "XMLHttpRequest",
        "Content-Type": "application/json",
    }


def fetch_dev_app_json(
    subdomain: str,
    username: str,
    password: str,
    app_id: str,
    resource: str,
    timeout: int = 30,
) -> Tuple[Optional[Any], Optional[str]]:
    """POST /k/api/dev/app/{app_id}/{resource} を空JSONで呼び、result を返す。

    管理画面と同じ内部API。パスワード認証ヘッダのみで動作し、
    ログインセッションやリクエストトークンは不要。APIトークンでは不可。
    """
    path = f"/k/api/dev/app/{app_id}/{resource.lstrip('/')}"
    url = f"https://{subdomain}.cybozu.com{path}"
    try:
        response = requests.post(
            url,
            headers=password_headers(username, password),
            json={},
            timeout=timeout,
        )
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
