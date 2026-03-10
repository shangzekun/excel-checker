from typing import Any, Dict, List, Optional

from pydantic import BaseModel, Field


class Issue(BaseModel):
    level: str
    message: str
    sheet: Optional[str] = None
    row: Optional[int] = None
    column: Optional[str] = None
    rule: Optional[str] = None

    # 用于前端聚合显示
    group_key: Optional[str] = None
    group_title: Optional[str] = None
    entity_type: Optional[str] = None
    entity_value: Optional[str] = None

    # 用于前端展开时显示更多对比信息
    details: Dict[str, Any] = Field(default_factory=dict)


class CheckResult(BaseModel):
    ok: bool
    summary: Dict[str, Any] = Field(default_factory=dict)
    issues: List[Issue] = Field(default_factory=list)
