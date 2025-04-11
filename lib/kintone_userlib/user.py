from typing import List, Optional
import datetime

class User:
    def __init__(self, username: str, email: str, name: str, is_active: bool):
        self.username = username
        self.email = email
        self.name = name
        self.is_active = is_active
        self.last_login: Optional[datetime.datetime] = None
        self.groups: List['Group'] = []

    def days_since_last_login(self, reference: Optional[datetime.datetime] = None) -> Optional[int]:
        if self.last_login is None:
            return None
        if reference is None:
            reference = datetime.datetime.now()
        return (reference - self.last_login).days
