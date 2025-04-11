import base64
import logging
import requests
from typing import List, Dict, Any
from datetime import datetime

from .user import User
from .group import Group

class KintoneClient:
    def __init__(self, subdomain: str, username: str, password: str, logger: logging.Logger):
        self.subdomain = subdomain
        self.base_url = f"https://{subdomain}.cybozu.com"
        self.headers = self._get_auth_header(username, password)
        self.logger = logger

    @staticmethod
    def _get_auth_header(username: str, password: str) -> Dict[str, str]:
        credentials = f"{username}:{password}"
        base64_credentials = base64.b64encode(credentials.encode('utf-8')).decode('utf-8')
        return {
            'X-Cybozu-Authorization': base64_credentials
        }

    def _fetch_data(self, endpoint: str, params: Dict[str, Any], key: str) -> List[Dict[str, Any]]:
        url = f"{self.base_url}/v1/{endpoint}.json"
        data = []
        size = 100
        offset = 0

        while True:
            current_params = params.copy()
            current_params.update({'size': size, 'offset': offset})
            response = requests.get(url, headers=self.headers, params=current_params)
            if response.status_code != 200:
                self.logger.error(f"{endpoint.capitalize()}の取得に失敗しました: {response.status_code} {response.text}")
                raise Exception(f"{endpoint.capitalize()}の取得に失敗しました: {response.status_code} {response.text}")
            batch = response.json().get(key, [])
            if not batch:
                break
            data.extend(batch)
            if len(batch) < size:
                break
            offset += size
            self.logger.debug(f"Fetched {len(batch)} items from {endpoint} (offset: {offset})")
        self.logger.info(f"全{endpoint}を取得しました。総数: {len(data)}")
        return data

    def get_all_users(self) -> List[User]:
        users_data = self._fetch_data('users', {}, 'users')
        users = []
        for user_data in users_data:
            user = User(
                username=user_data.get('code', ''),
                email=user_data.get('email', ''),
                name=user_data.get('name', ''),
                is_active=user_data.get('valid', True)
            )
            user.id = user_data.get('id')
            if 'lastLoginTime' in user_data:
                try:
                    user.last_login = datetime.fromisoformat(user_data['lastLoginTime'].replace('Z', '+00:00'))
                except ValueError:
                    self.logger.warning(f"Invalid last login time format for user {user.username}: {user_data['lastLoginTime']}")
            users.append(user)
        return users

    def get_all_groups(self) -> List[Group]:
        groups_data = self._fetch_data('groups', {}, 'groups')
        groups = []
        for group_data in groups_data:
            group = Group(
                code=group_data.get('code', ''),
                name=group_data.get('name', '')
            )
            groups.append(group)
        return groups

    def get_users_in_group(self, group_code: str) -> List[User]:
        params = {'code': group_code}
        users_data = self._fetch_data('group/users', params, 'users')
        users = []
        for user_data in users_data:
            user = User(
                username=user_data.get('code', ''),
                email=user_data.get('email', ''),
                name=user_data.get('name', ''),
                is_active=user_data.get('valid', True)
            )
            user.id = user_data.get('id')
            if 'lastLoginTime' in user_data:
                try:
                    user.last_login = datetime.fromisoformat(user_data['lastLoginTime'].replace('Z', '+00:00'))
                except ValueError:
                    self.logger.warning(f"Invalid last login time format for user {user.username}: {user_data['lastLoginTime']}")
            users.append(user)
        return users

    def get_user_groups(self, user_code: str) -> List[Group]:
        params = {'code': user_code}
        groups_data = self._fetch_data('user/groups', params, 'groups')
        groups = []
        for group_data in groups_data:
            group = Group(
                code=group_data.get('code', ''),
                name=group_data.get('name', '')
            )
            groups.append(group)
        return groups 