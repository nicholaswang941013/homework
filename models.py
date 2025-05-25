from dataclasses import dataclass
from datetime import datetime

@dataclass
class User:
    id: int
    username: str
    name: str
    email: str
    role: str

@dataclass
class Requirement:
    id: int
    title: str
    description: str
    assigner_id: int
    assignee_id: int
    status: str
    priority: str
    created_at: datetime
    scheduled_time: datetime = None  # 預約發派時間，None表示立即發派 
    comment: str = None  # 員工提交的完成情況說明
    completed_at: datetime = None  # 員工提交完成的時間 