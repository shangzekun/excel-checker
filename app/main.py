import io
import logging
import time
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook

from .checks import run_checks, get_rules, get_default_enabled_rule_ids, sanitize_selected_rule_ids
from .config import settings
from .excel_utils import validate_xlsx_bytes, ExcelValidationError
from .logging_conf import setup_logging
from .schemas import CheckResult

setup_logging()
logger = logging.getLogger(__name__)

app = FastAPI(title=settings.APP_NAME, version=settings.APP_VERSION)

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.ALLOW_ORIGINS,
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

def _parse_selected_rules(selected_rules: str) -> list[str]:
    ids = [x.strip() for x in selected_rules.split(",") if x.strip()]
    ids = sanitize_selected_rule_ids(ids)
    if not ids:
        ids = get_default_enabled_rule_ids()
    return ids

def _validate_upload(file: UploadFile, content: bytes) -> None:
    if not file.filename:
        raise HTTPException(status_code=400, detail="No filename")
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx is supported for now")
    if len(content) > settings.MAX_MB * 1024 * 1024:
        raise HTTPException(status_code=413, detail=f"File too large (>{settings.MAX_MB}MB)")
    try:
        validate_xlsx_bytes(content)
    except ExcelValidationError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

@app.get("/health")
def health():
    return {
        "status": "ok",
        "app": settings.APP_NAME,
        "version": settings.APP_VERSION,
        "env": settings.ENV,
    }

@app.get("/api/rules")
def rules():
    return {"rules": get_rules()}

@app.post("/api/check", response_model=CheckResult)
async def check_excel(file: UploadFile = File(...), selected_rules: str = Form(default="")):
    started = time.perf_counter()
    content = await file.read()
    _validate_upload(file, content)
    selected_rule_ids = _parse_selected_rules(selected_rules)
    logger.info("check start filename=%s size=%s rules=%s", file.filename, len(content), ",".join(selected_rule_ids))
    try:
        issues = run_checks(content, file.filename, selected_rule_ids)
    except Exception as exc:
        logger.exception("check failed filename=%s", file.filename)
        raise HTTPException(status_code=500, detail="Check execution failed") from exc

    summary = {
        "filename": file.filename,
        "size_bytes": len(content),
        "selected_rules": selected_rule_ids,
        "errors": sum(1 for x in issues if x.level == "error"),
        "warnings": sum(1 for x in issues if x.level == "warning"),
        "infos": sum(1 for x in issues if x.level == "info"),
        "total": len(issues),
        "duration_ms": int((time.perf_counter() - started) * 1000),
    }
    ok = summary["errors"] == 0
    logger.info(
        "check done filename=%s total=%s errors=%s warnings=%s infos=%s duration_ms=%s",
        file.filename, summary["total"], summary["errors"], summary["warnings"], summary["infos"], summary["duration_ms"]
    )
    return CheckResult(ok=ok, summary=summary, issues=issues)

@app.post("/api/report")
async def export_report(file: UploadFile = File(...), selected_rules: str = Form(default="")):
    started = time.perf_counter()
    content = await file.read()
    _validate_upload(file, content)
    selected_rule_ids = _parse_selected_rules(selected_rules)
    logger.info("report start filename=%s size=%s rules=%s", file.filename, len(content), ",".join(selected_rule_ids))
    try:
        issues = run_checks(content, file.filename, selected_rule_ids)
    except Exception as exc:
        logger.exception("report failed filename=%s", file.filename)
        raise HTTPException(status_code=500, detail="Report generation failed") from exc

    summary = {
        "filename": file.filename,
        "size_bytes": len(content),
        "selected_rules": ", ".join(selected_rule_ids),
        "errors": sum(1 for x in issues if x.level == "error"),
        "warnings": sum(1 for x in issues if x.level == "warning"),
        "infos": sum(1 for x in issues if x.level == "info"),
        "total": len(issues),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "duration_ms": int((time.perf_counter() - started) * 1000),
    }

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Summary"
    ws1.append(["key", "value"])
    for k, v in summary.items():
        ws1.append([k, v])

    ws2 = wb.create_sheet("Issues")
    ws2.append(["level", "message", "sheet", "row", "column", "rule"])
    for it in issues:
        ws2.append([it.level, it.message, it.sheet or "", it.row or "", it.column or "", it.rule or ""])

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)

    logger.info(
        "report done filename=%s total=%s errors=%s warnings=%s infos=%s duration_ms=%s",
        file.filename, summary["total"], summary["errors"], summary["warnings"], summary["infos"], summary["duration_ms"]
    )

    headers = {"Content-Disposition": 'attachment; filename="check_report.xlsx"'}
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )

app.mount("/", StaticFiles(directory="web", html=True), name="web")
