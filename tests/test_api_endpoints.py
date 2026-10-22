from fastapi.testclient import TestClient
from main import app

client = TestClient(app)

def test_health_endpoint():
    response = client.get("/api/v1/health")
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "online"
    assert "libreoffice_available" in data

def test_list_templates():
    response = client.get("/api/v1/templates")
    assert response.status_code == 200
    data = response.json()
    assert data["count"] >= 2
    codes = [t["template_code"] for t in data["templates"]]
    assert "INVOICE_REPORT" in codes
    assert "DEPT_EMPLOYEES" in codes

def test_render_invoice_html_binary():
    payload = {
        "template_code": "INVOICE_REPORT",
        "format": "html",
        "response_mode": "binary",
        "data": {
            "company": {"name": "Apex Enterprise Solutions Inc."},
            "invoice_no": "INV-2026-8801",
            "invoice_date": "2026-10-22",
            "customer": {"name": "Oracle Corporation"},
            "items": [
                {"item_code": "LIC-01", "description": "APEX Extension License", "qty": 10, "unit_price": 450.00},
                {"item_code": "SUP-02", "description": "Enterprise 24/7 Support", "qty": 1, "unit_price": 1200.00}
            ]
        }
    }
    response = client.post("/api/v1/reports/render", json=payload)
    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]
    assert "Apex Enterprise Solutions Inc." in response.text
    assert "INV-2026-8801" in response.text
    assert "APEX Extension License" in response.text
    assert "Enterprise 24/7 Support" in response.text

def test_render_invoice_html_json():
    payload = {
        "template_code": "INVOICE_REPORT",
        "format": "html",
        "response_mode": "json",
        "data": {
            "company": {"name": "Apex Enterprise Solutions Inc."},
            "invoice_no": "INV-2026-8802",
            "invoice_date": "2026-10-22",
            "customer": {"name": "Global Logistics LLC"},
            "items": [
                {"item_code": "ITM-99", "description": "Consulting Hour", "qty": 5, "unit_price": 120.00}
            ]
        }
    }
    response = client.post("/api/v1/reports/render", json=payload)
    assert response.status_code == 200
    data = response.json()
    assert data["status"] == "success"
    assert data["template_code"] == "INVOICE_REPORT"
    assert data["format"] == "html"
    assert "data_base64" in data
    assert data["size_bytes"] > 0
    assert data["execution_time_ms"] > 0

def test_render_dept_employees_xlsx():
    payload = {
        "template_code": "DEPT_EMPLOYEES",
        "format": "xlsx",
        "response_mode": "binary",
        "data": {
            "departments": [
                {
                    "dept_name": "ACCOUNTING",
                    "loc": "NEW YORK",
                    "employees": [
                        {"empno": 7782, "ename": "CLARK", "job": "MANAGER", "sal": 2450},
                        {"empno": 7839, "ename": "KING", "job": "PRESIDENT", "sal": 5000}
                    ]
                }
            ]
        }
    }
    response = client.post("/api/v1/reports/render", json=payload)
    assert response.status_code == 200
    assert "spreadsheetml" in response.headers["content-type"]
    assert len(response.content) > 1000
