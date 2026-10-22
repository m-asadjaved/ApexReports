import pytest
from fastapi.testclient import TestClient
from main import app

client = TestClient(app)

def test_studio_page_served_at_root():
    response = client.get("/studio")
    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]
    assert "ApexReports" in response.text
    assert "Studio" in response.text

def test_playground_alias_served():
    response = client.get("/playground")
    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]

def test_studio_page_served_under_api():
    response = client.get("/api/v1/studio")
    assert response.status_code == 200
    assert "text/html" in response.headers["content-type"]

def test_inspect_invoice_report():
    response = client.get("/api/v1/templates/INVOICE_REPORT/inspect")
    assert response.status_code == 200
    data = response.json()
    
    assert data["template_code"] == "INVOICE_REPORT"
    assert data["filename"] == "INVOICE_REPORT.xlsx"
    assert "variables" in data
    assert "blocks" in data
    assert "formulas" in data
    assert "media_tags" in data
    assert "sample_data" in data
    
    # Verify blocks
    targets = [b["target"] for b in data["blocks"] if b.get("type") == "loop"]
    assert "invoices" in targets
    assert "items" in targets
    
    # Verify formulas
    formulas = [f["formula"] for f in data["formulas"]]
    assert any("SUM" in f for f in formulas)
    
    # Verify media
    media_types = [m["type"] for m in data["media_tags"]]
    assert "barcode" in media_types
    assert "qr" in media_types

def test_inspect_dept_employees():
    response = client.get("/api/v1/templates/DEPT_EMPLOYEES/inspect")
    assert response.status_code == 200
    data = response.json()
    assert data["template_code"] == "DEPT_EMPLOYEES"
    
    # Check sample data exists
    assert "departments" in data["sample_data"]
    assert len(data["sample_data"]["departments"]) >= 2

def test_inspect_nonexistent_template():
    response = client.get("/api/v1/templates/DOES_NOT_EXIST/inspect")
    assert response.status_code == 404

def test_download_template():
    response = client.get("/api/v1/templates/INVOICE_REPORT/download")
    assert response.status_code == 200
    assert "spreadsheetml" in response.headers["content-type"]
    assert len(response.content) > 1000

def test_download_nonexistent_template():
    response = client.get("/api/v1/templates/NON_EXISTENT_FILE/download")
    assert response.status_code == 404
