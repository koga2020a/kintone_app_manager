#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
KintoneのAPIクライアント
"""

import os
import json
import requests
from pathlib import Path

class KintoneClient:
    def __init__(self, domain, api_token):
        self.domain = domain
        self.api_token = api_token
        self.base_url = f"https://{domain}.cybozu.com/k/v1"
        self.headers = {
            "X-Cybozu-API-Token": api_token,
            "Content-Type": "application/json"
        }

    def get_app_settings(self, app_id):
        """アプリの設定を取得する"""
        url = f"{self.base_url}/app/settings.json"
        params = {"app": app_id}
        response = requests.get(url, headers=self.headers, params=params)
        response.raise_for_status()
        return response.json()

    def get_app_form_fields(self, app_id):
        """アプリのフィールド設定を取得する"""
        url = f"{self.base_url}/app/form/fields.json"
        params = {"app": app_id}
        response = requests.get(url, headers=self.headers, params=params)
        response.raise_for_status()
        return response.json()

    def get_app_views(self, app_id):
        """アプリのビュー設定を取得する"""
        url = f"{self.base_url}/app/views.json"
        params = {"app": app_id}
        response = requests.get(url, headers=self.headers, params=params)
        response.raise_for_status()
        return response.json()

    def get_app_acl(self, app_id):
        """アプリのアクセス権限設定を取得する"""
        url = f"{self.base_url}/app/acl.json"
        params = {"app": app_id}
        response = requests.get(url, headers=self.headers, params=params)
        response.raise_for_status()
        return response.json()

    def get_app_notifications(self, app_id):
        """アプリの通知設定を取得する"""
        url = f"{self.base_url}/app/notifications.json"
        params = {"app": app_id}
        response = requests.get(url, headers=self.headers, params=params)
        response.raise_for_status()
        return response.json()

    def get_app_status(self, app_id):
        """アプリのステータス設定を取得する"""
        url = f"{self.base_url}/app/status.json"
        params = {"app": app_id}
        response = requests.get(url, headers=self.headers, params=params)
        response.raise_for_status()
        return response.json()

    def get_app_customize(self, app_id):
        """アプリのカスタマイズ設定を取得する"""
        url = f"{self.base_url}/app/customize.json"
        params = {"app": app_id}
        response = requests.get(url, headers=self.headers, params=params)
        response.raise_for_status()
        return response.json() 