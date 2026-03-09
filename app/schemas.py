from pydantic import BaseModel, Field
from typing import List, Optional, Literal, Dict, Any

Level = Literal["error", "warning", "info"]

class RuleMeta(BaseModel):
    id: str
    name: str
    level: Level
    description: str
    enabled: bool = True

class Issue(BaseModel):
    level: Level
    message: str
    sheet: Optional[str] = None
    row: Optional[int] = None
    column: Optional[str] = None
    rule: Optional[str] = None

class CheckResult(BaseModel):
    ok: bool
    summary: Dict[str, Any] = Field(default_factory=dict)
    issues: List[Issue] = Field(default_factory=list)
