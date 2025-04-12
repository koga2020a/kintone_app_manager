from .user import User
from .group import Group
from .manager import UserManager

# UserManagerのインスタンスを作成
user_manager = UserManager()

def get_priority_domain():
    """優先ドメインを取得する"""
    return user_manager.primary_domain

__all__ = ["User", "Group", "UserManager", "user_manager", "get_priority_domain"]
