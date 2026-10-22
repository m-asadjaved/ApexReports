# ApexReports: High-Performance Reporting Engine for Oracle APEX

[![Python](https://img.shields.io/badge/Python-3.12-blue.svg)](https://www.python.org/)
[![FastAPI](https://img.shields.io/badge/FastAPI-0.110+-green.svg)](https://fastapi.tiangolo.com/)
[![Docker](https://img.shields.io/badge/Docker-Ready-2496ED.svg)](https://www.docker.com/)

**ApexReports** is a specialized, open-source reporting microservice designed specifically for **Oracle APEX** developers. It combines the visual layout and formula power of **Microsoft Excel (.xlsx)** templates with classical **Oracle Reports 6i** features (repeating frames, master-detail groups, matrix cross-tabs, group summary auto-adjustments, format triggers, dynamic barcodes, and QR codes).

---

## Documentation Hub

Comprehensive, production-ready documentation is available in the [`docs/`](file:///d:/Code/ApexReports/docs/README.md) directory:

| Guide | Description |
| :--- | :--- |
| 📖 **[Template Design Guide](file:///d:/Code/ApexReports/docs/template-design-guide.md)** | Code-first tutorial for creating Master-Detail, Matrix, Barcode, QR code, and auto-expanding summary templates in Excel. |
| 🛡️ **[Troubleshooting & Pitfalls](file:///d:/Code/ApexReports/docs/troubleshooting-and-pitfalls.md)** | Senior engineer guide: Top 10 mistakes, merged cell traps, formula edge cases, barcode sizing, and logging. |
| ⚡ **[Oracle APEX Integration Guide](file:///d:/Code/ApexReports/docs/apex-integration-guide.md)** | PL/SQL setup (`apex_reports_pkg`), direct downloads, in-modal HTML previews, and automated email attachments. |
| 🔌 **[REST API Reference](file:///d:/Code/ApexReports/docs/api-reference.md)** | Full OpenAPI schema, endpoints, binary vs JSON response modes, health checks, and Postman collection. |
| 🏛️ **[Architecture & Overview](file:///d:/Code/ApexReports/docs/README.md)** | System components, dual-engine rendering pipeline, and technical specs. |

---

## Key Features

- **Designed in Microsoft Excel**: Create and style templates directly in Microsoft Excel (`.xlsx`) using familiar fonts, borders, alignments, and native Excel formulas (`=SUM()`, `=C5*D5`, `=AVERAGE()`).
- **Oracle Reports 6i Parity**:
  - **Repeating Frames (Down)**: `{{#each items}}` ... `{{/each}}` with support for arbitrary nested Master-Detail loops.
  - **Matrix / Cross-Tabs (Across)**: `{{#each-col quarters}}` ... `{{/each-col}}` for dynamic horizontal column expansion.
  - **Automatic Formula Range Expansion**: Summary formulas like `=SUM(E7:E7)` automatically expand to `=SUM(E7:E25)` when repeating rows grow.
  - **Group Summary Resets**: Formulas inside an outer loop only calculate subtotals for that specific group (e.g. Department Salary Subtotal).
  - **Format Triggers**: `{{#if condition}}` ... `{{/if}}` conditionally suppresses or displays rows.
  - **Group Page Breaks**: `{{#page-break}}` or `{{#each depts page_break=true}}` inserts native Excel print page breaks.
  - **Dynamic Media**: Dynamic QR codes (`{{@qr url}}`), 1D barcodes (`{{@barcode id format=code128}}`), and dynamic images (`{{@image url_or_base64}}`).
- **Dual-Engine Rendering**:
  - **Fast HTML Preview**: Sub-25ms in-memory conversion for instant feedback inside APEX modals/regions.
  - **Print-Ready PDF**: Headless LibreOffice Calc recalculates all formulas and generates crisp vector PDFs.
- **Ready for Oracle APEX**: Comes with a drop-in PL/SQL client package ([`apex_reports_pkg`](file:///d:/Code/ApexReports/apex/apex_reports_pkg.pks)).

---

## Quickstart

### 1. Run with Docker (Recommended for PDF Generation)
```bash
docker compose up --build -d
```
The service will start on `http://localhost:8000`.

### 2. Run Locally (Python 3.12+)
```bash
# Create and activate virtual environment
python -m venv .venv
.\.venv\Scripts\activate   # Windows
# source .venv/bin/activate # Linux / macOS

# Install dependencies
pip install -r requirements.txt

# Run the server
uvicorn main:app --reload --port 8000
```

Interactive OpenAPI Swagger documentation will be available at:
`http://localhost:8000/docs`

---

## API Reference

### Render Report
`POST /api/v1/reports/render`

#### JSON Request:
```json
{
  "template_code": "INVOICE_REPORT",
  "format": "pdf",
  "response_mode": "binary",
  "filename": "Invoice_1001.pdf",
  "data": {
    "company": { "name": "Apex Enterprise Solutions Inc." },
    "invoice_no": "INV-2026-8801",
    "invoice_date": "2026-10-22",
    "customer": { "name": "Oracle Corporation" },
    "items": [
      { "item_code": "LIC-01", "description": "APEX Extension License", "qty": 10, "unit_price": 450.00 },
      { "item_code": "SUP-02", "description": "Enterprise 24/7 Support", "qty": 1, "unit_price": 1200.00 }
    ]
  }
}
```

#### Parameters:
| Field | Type | Description |
| :--- | :--- | :--- |
| `template_code` | string | Name of the `.xlsx` template located in `templates/`. |
| `format` | string | `pdf`, `html`, or `xlsx`. |
| `response_mode` | string | `binary` (stream file directly) or `json` (returns Base64 envelope). |
| `filename` | string | Optional attachment filename. |
| `data` | object | The hierarchical JSON payload. |

---

## Template Syntax Guide

Place tags directly in your Excel cells:

| Directive | Syntax in Cell | Description |
| :--- | :--- | :--- |
| **Field** | `{{customer.name}}` | Injects scalar values preserving native Excel cell formatting. |
| **Repeating Rows** | `{{#each items}}` ... `{{/each}}` | Repeats detail rows for each item in the array. |
| **Matrix Columns** | `{{#each-col quarters}} {{name}} {{/each-col}}` | Horizontally expands columns for cross-tab reports. |
| **Conditional Row** | `{{#if discount > 0}}` ... `{{/if}}` | Shows or hides rows based on expression evaluation. |
| **Page Break** | `{{#page-break}}` | Inserts horizontal page break into print setup. |
| **QR Code** | `{{@qr 'https://apex.corp/inv/' + id width=80 height=80}}` | Generates embedded QR code image. |
| **Barcode** | `{{@barcode item_code format=code128 width=140 height=35}}` | Generates embedded 1D barcode image. |
| **Image** | `{{@image logo_url width=120 height=40}}` | Embeds image from URL or Base64 string. |

---

## Oracle APEX Integration

Compile the PL/SQL wrapper in your APEX schema:
- Specification: [`apex/apex_reports_pkg.pks`](file:///d:/Code/ApexReports/apex/apex_reports_pkg.pks)
- Body: [`apex/apex_reports_pkg.pkb`](file:///d:/Code/ApexReports/apex/apex_reports_pkg.pkb)

### Single-Line PDF Download from APEX:
```sql
apex_reports_pkg.download_pdf(
    p_template_code => 'INVOICE_REPORT',
    p_json_data     => l_json_data,
    p_filename      => 'Invoice_' || :P10_ORDER_ID || '.pdf'
);
```

For complete recipes (AJAX modal previews, email attachments via `APEX_MAIL`), see the [Oracle APEX Integration Guide](file:///d:/Code/ApexReports/apex/apex_page_setup_guide.md).

---

## Running Automated Tests
```bash
pytest tests
```
