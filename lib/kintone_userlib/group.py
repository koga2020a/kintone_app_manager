from typing import List, Optional
from .user import User

class Group:
    def __init__(self, code: str, name: str):
        self.code = code
        self.name = name
        self.users: List[User] = []

    def add_user(self, user: User):
        if user not in self.users:
            self.users.append(user)
            user.groups.append(self)

    def remove_user(self, user: User):
        if user in self.users:
            self.users.remove(user)
            user.groups.remove(self)

    def __eq__(self, other):
        if not isinstance(other, Group):
            return False
        return self.code == other.code

    def __hash__(self):
        return hash(self.code)

    def get_sorted_members(self, primary_domain: str) -> List[User]:
        primary = []
        secondary = {}

        for user in self.users:
            domain = user.email.split('@')[-1] if user.email else ''
            if domain == primary_domain:
                primary.append(user)
            else:
                secondary.setdefault(domain, []).append(user)

        primary.sort(key=lambda u: u.email)
        sorted_secondary = []
        for domain in sorted(secondary.keys()):
            sorted_secondary.extend(sorted(secondary[domain], key=lambda u: u.email))

        return primary + sorted_secondary
