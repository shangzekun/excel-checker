from typing import List

from ..schemas import Issue

MR_REVIEW_RULE = {
    "id": "mr_review_check",
    "name": "MR检查",
    "level": "info",
    "description": "MR检查项预留占位，后续补充具体业务逻辑。",
    "enabled": False,
    "details": ["当前仅保留占位，不执行具体检查逻辑。"],
}


def run_mr_review_check(
    file_bytes: bytes,
    filename: str,
    ebom_bytes: bytes | None = None,
    ebom_filename: str | None = None,
) -> List[Issue]:
    return []
