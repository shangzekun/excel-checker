from io import BytesIO
from fastapi.testclient import TestClient
from openpyxl import Workbook

from app.main import app

client = TestClient(app)

def make_xlsx_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "demo"
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def test_health():
    resp = client.get("/health")
    assert resp.status_code == 200
    assert resp.json()["status"] == "ok"

def test_rules():
    resp = client.get("/api/rules")
    assert resp.status_code == 200
    assert "rules" in resp.json()
    rule_ids = {item["id"] for item in resp.json()["rules"]}
    assert rule_ids == {"data_maturity_check", "mr_check"}

def test_check_valid_xlsx():
    data = make_xlsx_bytes()
    resp = client.post(
        "/api/check",
        files={"file": ("demo.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"selected_rules": "data_maturity_check"},
    )
    assert resp.status_code == 200
    payload = resp.json()
    assert "summary" in payload
    assert payload["summary"]["total"] >= 1

def test_check_invalid_file_extension():
    resp = client.post(
        "/api/check",
        files={"file": ("demo.txt", b"abc", "text/plain")},
        data={"selected_rules": "data_maturity_check"},
    )
    assert resp.status_code == 400

def test_check_invalid_xlsx_content():
    resp = client.post(
        "/api/check",
        files={"file": ("fake.xlsx", b"not real xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"selected_rules": "data_maturity_check"},
    )
    assert resp.status_code == 400

def test_report_valid_xlsx():
    data = make_xlsx_bytes()
    resp = client.post(
        "/api/report",
        files={"file": ("demo.xlsx", data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
        data={"selected_rules": "data_maturity_check"},
    )
    assert resp.status_code == 200
    assert resp.headers["content-type"].startswith("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
