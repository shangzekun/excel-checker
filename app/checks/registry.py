import logging
from typing import List, Dict, Any, Callable

from ..schemas import Issue
from .data_maturity import run_data_maturity_check
from .mr_check import run_mr_check

logger = logging.getLogger(__name__)

RuleHandler = Callable[[bytes, str, bytes | None, str | None], List[Issue]]

RULES: List[Dict[str, Any]] = [
    {
        "id": "data_maturity_check",
        "name": "数据成熟度检查",
        "level": "warning",
        "description": "",
        "enabled": True,
        "details": [
            "连接点号规范性检查：连接点号为“xxx123456_7890”，其中，“xxx”为对应工艺的三位字母，123456代表Fixing Part的后六位数字，最后五位为“下划线”+四位数字。标准点号规范如：“RSW349075_0104”。",
            "材料牌号规范性检查：建立了标准材料牌号规范库，对点表中的材料牌号进行一一比对，判断点表中材料牌号是否规范。",
            "连接点号重复性检查：依次遍历点表中的连接点号，判断连接点号是否唯一。",
            "连接点号位置重复性检查：按照空间位置10mm距离进行检查，两连接点的空间位置相距≤10mm报错。",
            "钣金零件：同一零件版本号一致性检查：依次遍历点表中的零件及其版本号，判断同一零件在该点表中的版本号是否唯一。",
            "钣金零件：同一零件材料牌号一致性检查：依次遍历点表中的零件及其材料牌号，判断同一零件在该点表中的材料牌号是否唯一。",
            "钣金零件：同一零件材料厚度一致性检查：依次遍历点表中的零件及其厚度，判断同一零件在该点表中的厚度是否唯一。由于铸件的厚度不唯一，该检查项并没有对铸件的厚度进行检查。",
            "Fixing-Part版本号与BOM中版本号一致性检查：将点表中Fixing-Part的版本号与选中的BOM清单中的版本号进行比对，判断两者是否相同。",
            "零件版本号：点表与BOM一致性检查：将点表中每个零件的版本号与选中的BOM清单中的该零件版本号进行比对，判断两者是否相同。",
            "零件材料牌号：点表与BOM一致性检查：将点表中每个零件的材料牌号与选中的BOM清单中的该零件材料牌号进行比对，判断两者是否相同。",
            "零件厚度：点表与BOM一致性检查：将点表中每个零件的厚度与选中的BOM清单中的该零件厚度进行比对，判断两者是否相同。由于铸件的厚度不唯一，该检查项并没有对铸件的厚度进行检查。",
            "工艺点属性的检查：如果该连接点属于单边点焊或四层搭接，判断该连接点是否被标注为工艺点。如果该连接点被标注为工艺点，判断该连接点是否属于单边点焊或四层搭接，如若不是给出提示，需人工确认该点是否为工艺点。",
        ],
    },
    {
        "id": "mr_check",
        "name": "MR检查",
        "level": "info",
        "description": "MR检查预留项，当前为占位实现，后续补充具体规则。",
        "enabled": False,
        "details": ["当前为占位规则，后续将按业务需求补充具体检查逻辑。"],
    },
]

RULE_HANDLERS: Dict[str, RuleHandler] = {
    "data_maturity_check": run_data_maturity_check,
    "mr_check": run_mr_check,
}


def get_rules() -> List[Dict[str, Any]]:
    return RULES


def get_default_enabled_rule_ids() -> List[str]:
    return [r["id"] for r in RULES if r.get("enabled")]


def sanitize_selected_rule_ids(selected_rule_ids: List[str]) -> List[str]:
    valid = set(RULE_HANDLERS.keys())
    return [rule_id for rule_id in selected_rule_ids if rule_id in valid]


def run_checks(
    file_bytes: bytes,
    filename: str,
    selected_rule_ids: List[str] | None = None,
    ebom_bytes: bytes | None = None,
    ebom_filename: str | None = None,
) -> List[Issue]:
    if not selected_rule_ids:
        selected_rule_ids = get_default_enabled_rule_ids()

    selected_rule_ids = sanitize_selected_rule_ids(selected_rule_ids)

    issues: List[Issue] = []

    for rule_id in selected_rule_ids:
        handler = RULE_HANDLERS.get(rule_id)
        if not handler:
            continue

        try:
            issues.extend(handler(file_bytes, filename, ebom_bytes, ebom_filename))
        except Exception:
            logger.exception("Rule execution failed: %s", rule_id)
            issues.append(
                Issue(
                    level="error",
                    message=f"规则执行异常：{rule_id}",
                    sheet=None,
                    row=None,
                    column=None,
                    rule=rule_id,
                    group_key=f"{rule_id}:execution_failed",
                    group_title=f"规则执行异常：{rule_id}",
                )
            )

    return issues
