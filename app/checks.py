from typing import List, Dict, Any, Callable
import logging
from .schemas import Issue

logger = logging.getLogger(__name__)

def get_rules() -> List[Dict[str, Any]]:
    return [
        {
            "id": "required_fields_demo",
            "name": "必填字段检查（Demo）",
            "level": "error",
            "description": "模拟检查关键字段是否为空，返回若干 error 示例。",
            "enabled": True,
            "details": [
                "用于模拟关键主字段不能为空的情况，例如客户名称、订单编号、主键编号等。",
                "当前 demo 会返回 1 到 2 条 error 级别示例，方便联调前端样式。",
                "后续替换时，可在这里接入表头识别、必填列配置与逐行校验逻辑。"
            ],
        },
        {
            "id": "date_format_demo",
            "name": "日期格式检查（Demo）",
            "level": "warning",
            "description": "模拟检查日期格式是否标准，返回若干 warning 示例。",
            "enabled": True,
            "details": [
                "用于模拟日期类字段格式不统一的问题。",
                "可扩展为文本日期、Excel 日期序列值、非法日期字符串等多种判断。",
                "正式接入后可支持自动识别列名与格式模板。"
            ],
        },
        {
            "id": "enum_value_demo",
            "name": "枚举值检查（Demo）",
            "level": "info",
            "description": "模拟检查枚举值是否在允许范围内，返回若干 info 示例。",
            "enabled": True,
            "details": [
                "用于模拟状态、类型、分类等字段不在允许枚举中的场景。",
                "当前返回 info 示例，后续可按业务要求升级为 warning 或 error。",
                "适合与字典表或配置中心联动。"
            ],
        },
        {
            "id": "cross_sheet_demo",
            "name": "跨表关联检查（Demo）",
            "level": "error",
            "description": "模拟主子表关联关系检查，例如主键缺失、引用不存在等。",
            "enabled": True,
            "details": [
                "用于模拟多 Sheet 之间的主外键关联检查。",
                "例如子表引用主表不存在、发票引用订单缺失等。",
                "正式逻辑可以扩展为多表映射、唯一键校验与缺失明细定位。"
            ],
        },
        {
            "id": "range_check_demo",
            "name": "数值范围检查（Demo）",
            "level": "warning",
            "description": "模拟检查金额、数量、比例等是否落在合理区间。",
            "enabled": True,
            "details": [
                "用于模拟数值超过上下限、比例越界、金额异常等场景。",
                "支持后续扩展为按列配置阈值、按币种区分范围。",
                "适用于库存数量、折扣率、金额、税率等字段。"
            ],
        },
    ]

def _required_fields_demo(_: bytes, filename: str) -> List[Issue]:
    return [
        Issue(level="error", message="客户名称不能为空", sheet="客户清单", row=3, column="B", rule="required_fields_demo"),
        Issue(level="error", message="订单编号不能为空", sheet="订单明细", row=8, column="A", rule="required_fields_demo"),
    ]

def _date_format_demo(_: bytes, filename: str) -> List[Issue]:
    return [
        Issue(level="warning", message="日期格式建议统一为 YYYY-MM-DD", sheet="订单明细", row=5, column="D", rule="date_format_demo"),
        Issue(level="warning", message="检测到文本日期，建议转换为标准日期单元格", sheet="发票记录", row=12, column="C", rule="date_format_demo"),
    ]

def _enum_value_demo(_: bytes, filename: str) -> List[Issue]:
    return [
        Issue(level="info", message="状态值“已完成-手工”不在推荐枚举内，建议映射为“已完成”", sheet="任务列表", row=6, column="F", rule="enum_value_demo")
    ]

def _cross_sheet_demo(_: bytes, filename: str) -> List[Issue]:
    return [
        Issue(level="error", message="子表中的客户ID在主表中不存在", sheet="订单明细", row=9, column="B", rule="cross_sheet_demo"),
        Issue(level="error", message="发票记录引用的订单号未在订单明细中找到", sheet="发票记录", row=4, column="A", rule="cross_sheet_demo"),
    ]

def _range_check_demo(_: bytes, filename: str) -> List[Issue]:
    return [
        Issue(level="warning", message="折扣率超过 100%，请确认数据是否正确", sheet="促销策略", row=7, column="E", rule="range_check_demo"),
        Issue(level="warning", message="数量为负数，建议复核", sheet="库存流水", row=15, column="D", rule="range_check_demo"),
    ]

RULE_HANDLERS: Dict[str, Callable[[bytes, str], List[Issue]]] = {
    "required_fields_demo": _required_fields_demo,
    "date_format_demo": _date_format_demo,
    "enum_value_demo": _enum_value_demo,
    "cross_sheet_demo": _cross_sheet_demo,
    "range_check_demo": _range_check_demo,
}

def get_default_enabled_rule_ids() -> List[str]:
    return [r["id"] for r in get_rules() if r.get("enabled")]

def sanitize_selected_rule_ids(selected_rule_ids: List[str]) -> List[str]:
    valid = set(RULE_HANDLERS.keys())
    return [rule_id for rule_id in selected_rule_ids if rule_id in valid]

def run_checks(file_bytes: bytes, filename: str, selected_rule_ids: List[str] | None = None) -> List[Issue]:
    if not selected_rule_ids:
        selected_rule_ids = get_default_enabled_rule_ids()

    selected_rule_ids = sanitize_selected_rule_ids(selected_rule_ids)
    issues: List[Issue] = []

    for rule_id in selected_rule_ids:
        handler = RULE_HANDLERS.get(rule_id)
        if not handler:
            continue
        try:
            issues.extend(handler(file_bytes, filename))
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
                )
            )
    return issues
