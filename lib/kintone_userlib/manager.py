import pickle
from typing import Dict, Optional, List, Set
from .user import User
from .group import Group

class UserManager:
    def __init__(self):
        self.users: Dict[str, User] = {}
        self.groups: Dict[str, Group] = {}
        self.primary_domain: Optional[str] = None
        self.user_groups: Dict[str, List[Group]] = {}
        self.group_users: Dict[str, List[User]] = {}

    def set_primary_domain(self, domain: str):
        """主要ドメインを設定"""
        self.primary_domain = domain

    def add_user(self, user: User):
        """ユーザーを追加し、関連するグループ情報も更新"""
        self.users[user.username] = user
        self.user_groups[user.username] = user.groups
        for group in user.groups:
            if group.code not in self.group_users:
                self.group_users[group.code] = []
            if user not in self.group_users[group.code]:
                self.group_users[group.code].append(user)

    def add_group(self, group: Group):
        """グループを追加し、関連するユーザー情報も更新"""
        self.groups[group.code] = group
        self.group_users[group.code] = group.users
        for user in group.users:
            if user.username not in self.user_groups:
                self.user_groups[user.username] = []
            if group not in self.user_groups[user.username]:
                self.user_groups[user.username].append(group)

    def get_user(self, username: str) -> Optional[User]:
        return self.users.get(username)

    def get_group(self, group_name: str) -> Optional[Group]:
        for group in self.groups.values():
            if group.name == group_name:
                return group
        return None

    def get_users_in_group(self, group_name: str) -> List[User]:
        group = self.get_group(group_name)
        return group.get_sorted_members(self.primary_domain) if group else []

    def get_all_users(self) -> List[User]:
        return list(self.users.values())

    def get_all_groups(self) -> List[Group]:
        return list(self.groups.values())

    def to_pickle(self, filepath: str):
        with open(filepath, 'wb') as f:
            pickle.dump(self, f)

    @staticmethod
    def from_pickle(filepath: str) -> 'UserManager':
        with open(filepath, 'rb') as f:
            return pickle.load(f)

    def debug_display(self):
        """保存されたデータのデバッグ表示"""
        print("\n=== ユーザー情報 ===")
        for user_code, user in self.users.items():
            print(f"\nユーザーコード: {user_code}")
            print(f"名前: {user.name}")
            print(f"メールアドレス: {user.email}")
            print(f"所属グループ: {', '.join([g.name for g in user.groups])}")

        print("\n=== グループ情報 ===")
        for group_code, group in self.groups.items():
            print(f"\nグループコード: {group_code}")
            print(f"グループ名: {group.name}")
            print(f"所属ユーザー: {', '.join([u.name for u in group.users])}")

        print("\n=== ユーザー-グループ関連 ===")
        for user_code, groups in self.user_groups.items():
            print(f"\nユーザーコード: {user_code}")
            print(f"所属グループ: {', '.join([g.name for g in groups])}")

        print("\n=== グループ-ユーザー関連 ===")
        for group_code, users in self.group_users.items():
            print(f"\nグループコード: {group_code}")
            print(f"所属ユーザー: {', '.join([u.name for u in users])}")

    def get_user_domain(self, user: User) -> str:
        """ユーザーのドメインを取得"""
        if user.email is None:
            return ""
        if '@' in user.email:
            return user.email.split('@')[1]
        return ""
        
    def sort_users_by_domain(self, users: List[User]) -> Dict[str, List[User]]:
        """ユーザーをドメインごとに分類し、各ドメイン内でさらにアクティブステータスでグループ分け"""
        domain_users: Dict[str, List[User]] = {}
        
        # 主要ドメインのユーザーを第一グループとする
        if self.primary_domain:
            primary_users = [u for u in users if self.get_user_domain(u) == self.primary_domain]
            
            # アクティブステータスでグループ分け
            active_users = [u for u in primary_users if hasattr(u, 'is_active') and u.is_active]
            inactive_users = [u for u in primary_users if hasattr(u, 'is_active') and not u.is_active]
            unknown_status_users = [u for u in primary_users if not hasattr(u, 'is_active')]
            
            # 各グループ内で名前でソート
            active_users = sorted(active_users, key=lambda u: u.name)
            inactive_users = sorted(inactive_users, key=lambda u: u.name)
            unknown_status_users = sorted(unknown_status_users, key=lambda u: u.name)
            
            # 有効ユーザー→ステータス不明ユーザー→無効ユーザーの順に結合
            domain_users[self.primary_domain] = active_users + unknown_status_users + inactive_users
            
        # その他のドメインのユーザーを第二グループとする
        other_domains: Set[str] = set()
        for user in users:
            domain = self.get_user_domain(user)
            if domain != self.primary_domain and domain:
                other_domains.add(domain)
                
        # 各ドメインのユーザーをアクティブステータスでグループ分け
        for domain in sorted(other_domains):
            domain_user_list = [u for u in users if self.get_user_domain(u) == domain]
            
            # アクティブステータスでグループ分け
            active_users = [u for u in domain_user_list if hasattr(u, 'is_active') and u.is_active]
            inactive_users = [u for u in domain_user_list if hasattr(u, 'is_active') and not u.is_active]
            unknown_status_users = [u for u in domain_user_list if not hasattr(u, 'is_active')]
            
            # 各グループ内で名前でソート
            active_users = sorted(active_users, key=lambda u: u.name)
            inactive_users = sorted(inactive_users, key=lambda u: u.name)
            unknown_status_users = sorted(unknown_status_users, key=lambda u: u.name)
            
            # 有効ユーザー→ステータス不明ユーザー→無効ユーザーの順に結合
            domain_users[domain] = active_users + unknown_status_users + inactive_users
            
        return domain_users
        
    def display_group_users_detail(self):
        """グループ-ユーザー関連の詳細表示"""
        print("\n=== グループ-ユーザー関連 （ユーザ情報詳細）===")
        
        # グループコードでソートしたグループ一覧
        for group_code in sorted(self.group_users.keys()):
            group = self.groups.get(group_code)
            if not group:
                continue
                
            users = self.group_users[group_code]
            print(f"\n■ グループ: {group.name} (コード: {group_code})")
            
            # ユーザーをドメイン別に分類（表示ではなくソート用）
            domain_users = self.sort_users_by_domain(users)
            
            # 主要ドメインのユーザーを表示
            if self.primary_domain and self.primary_domain in domain_users:
                for user in domain_users[self.primary_domain]:
                    self._print_user_detail(user)
            
            # その他のドメインのユーザーを表示
            for domain, domain_users_list in domain_users.items():
                if domain != self.primary_domain:
                    for user in domain_users_list:
                        self._print_user_detail(user)

    def _print_user_detail(self, user: User):
        """ユーザー詳細情報の表示"""
        print(f"\nユーザー名: {user.name}")
        if hasattr(user, 'id'):
            print(f"ID: {user.id}")
        print(f"ユーザーコード: {user.username}")
        print(f"メールアドレス: {user.email}")
        if hasattr(user, 'is_active'):
            status = "有効" if user.is_active else "無効"
            print(f"ステータス: {status}")
        print(f"所属グループ: {', '.join([g.name for g in user.groups])}")
        print(f"ドメイン: {self.get_user_domain(user)}")
        # その他の属性があれば表示
        if hasattr(user, 'primary_organization'):
            print(f"主組織: {user.primary_organization}")
        if hasattr(user, 'title'):
            print(f"役職: {user.title}")
        # 区切り線
        print("-" * 40)
