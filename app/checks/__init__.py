import logging
from typing import List, Dict, Any, Callable

from ..schemas import Issue
from .data_maturity import DATA_MATURITY_RULE, run_data_maturity_check
from .mr_review import MR_REVIEW_RULE, run_mr_review_check

logger = logging.getLogger(__name__)

RuleHandler = Callable[[bytes, str, bytes | None, str | None], List[Issue]]

RULE_DEFINITIONS: List[Dict[str, Any]] = [
    DATA_MATURITY_RULE,
    MR_REVIEW_RULE,
]

RULE_HANDLERS: Dict[str, RuleHandler] = {
    "data_maturity_check": run_data_maturity_check,
    "mr_review_check": run_mr_review_check,
}


def get_rules() -> List[Dict[str, Any]]:
    return RULE_DEFINITIONS


def get_default_enabled_rule_ids() -> List[str]:
    return [r["id"] for r in get_rules() if r.get("enabled")]


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
