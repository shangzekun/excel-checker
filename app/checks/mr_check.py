from typing import List

from ..schemas import Issue


def run_mr_check(
    file_bytes: bytes,
    filename: str,
    ebom_bytes: bytes | None = None,
    ebom_filename: str | None = None,
) -> List[Issue]:
    """MR 检查占位实现，后续补充具体业务逻辑。"""
    return [
        Issue(
            level="info",
            message="MR检查规则已预留，当前为占位实现，暂未启用具体检查逻辑。",
            sheet=None,
            row=None,
            column=None,
            rule="MR检查",
            group_key="mr_check_placeholder:not_implemented",
            group_title="MR检查占位项（待实现）",
            entity_type="mr_check",
        )
    ]
