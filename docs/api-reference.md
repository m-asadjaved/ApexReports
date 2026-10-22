# REST API Reference

The ApexReports microservice provides a RESTful interface for rendering report templates and managing the template repository.

---

## 1. Gateway Overview

- **Base URL:** `http://<host>:<port>/api/v1`
- **Default Port:** `8000`
- **Interactive Documentation:** `http://<host>:8000/docs` (Swagger UI) or `/redoc`
- **Authentication:** Optional. Configurable via `API_KEY` environment variable. When set, pass header:
  `Authorization: Bearer <API_KEY>` or `X-API-Key: <API_KEY>`.

---

## 2. Endpoints

### 2.1 Render Report
`POST /api/v1/reports/render`

Populates a pre-deployed Excel template with JSON data and outputs the result as PDF, HTML, or XLSX.

#### Request Headers:
| Header | Value | Description |
| :--- | :--- | :--- |
| `Content-Type` | `application/json` | Required |
| `Accept` | `*/*` | Optional |

#### Request Body (JSON):
```json
{
  "template_code": "INVOICE_REPORT",
  "format": "pdf",
  "response_mode": "binary",
  "filename": "Invoice_10024.pdf",
  "data": {
    "company": { "name": "Acme Global Solutions" },
    "invoice_no": "INV-10024",
    "invoice_date": "2026-10-22",
    "customer": { "name": "Oracle Corporation" },
    "items": [
      { "item_code": "LIC-01", "description": "APEX Extension License", "qty": 10, "unit_price": 450.00 }
    ]
  }
}
```

#### Request Parameters:
| Property | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `template_code` | string | *required* | The filename of the `.xlsx` template in `templates/` (without extension). |
| `format` | string | `"pdf"` | Output format: `"pdf"`, `"html"`, or `"xlsx"`. |
| `response_mode` | string | `"binary"` | Delivery mode: `"binary"` (stream raw bytes) or `"json"` (Base64 payload). |
| `filename` | string | `null` | Suggested filename for browser download / `Content-Disposition`. |
| `data` | object | `{}` | The hierarchical dataset matching template tags. |

#### Delivery Modes:

##### Mode A: Binary Stream (`response_mode: "binary"`)
Returns raw file bytes with appropriate MIME headers and timing metadata.
- **Status:** `200 OK`
- **Headers:**
  - `Content-Type: application/pdf` *(or `text/html; charset=utf-8`, or `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`)*
  - `Content-Disposition: inline; filename="Invoice_10024.pdf"`
  - `X-Execution-Time-Ms: 184.2`

##### Mode B: JSON Envelope (`response_mode: "json"`)
Returns metadata, page count, and Base64-encoded payload for easy parsing in PL/SQL or microservices.
- **Status:** `200 OK`
- **Response Payload:**
```json
{
  "status": "success",
  "template_code": "INVOICE_REPORT",
  "format": "pdf",
  "filename": "Invoice_10024.pdf",
  "mime_type": "application/pdf",
  "size_bytes": 95830,
  "execution_time_ms": 184.2,
  "data_base64": "JVBERi0xLjQKJcTl8uXrp..."
}
```

---

### 2.2 Template Management

#### List Templates
`GET /api/v1/templates`

Returns an inventory of all templates available in the repository.

**Response:**
```json
{
  "count": 2,
  "templates": [
    {
      "template_code": "INVOICE_REPORT",
      "filename": "INVOICE_REPORT.xlsx",
      "size_bytes": 10278,
      "modified_at": "2026-10-22 21:59:30"
    },
    {
      "template_code": "DEPT_EMPLOYEES",
      "filename": "DEPT_EMPLOYEES.xlsx",
      "size_bytes": 8945,
      "modified_at": "2026-10-22 21:40:15"
    }
  ]
}
```

#### Upload Template
`POST /api/v1/templates/upload`

Uploads or hot-reloads an `.xlsx` template file without restarting the service.

- **Content-Type:** `multipart/form-data`
- **Form Field:** `file` (the `.xlsx` binary)

**Response:**
```json
{
  "status": "success",
  "template_code": "SALES_SUMMARY",
  "filename": "SALES_SUMMARY.xlsx",
  "size_bytes": 14210
}
```

---

### 2.3 System Diagnostics
`GET /api/v1/health`

Reports system status, directory paths, and active PDF rendering backend.

**Response:**
```json
{
  "status": "online",
  "service": "ApexReports Engine",
  "pdf_backend": "Headless Browser (C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe)",
  "libreoffice_available": false,
  "templates_directory": "D:\\Code\\ApexReports\\templates"
}
```

---

## 3. Performance & Rendering Engines

| Output Format | Rendering Backend | Typical Latency | Key Strengths |
| :--- | :--- | :--- | :--- |
| **HTML Preview** | In-Memory `FastHtmlRenderer` | **10ms – 25ms** | Instant feedback in APEX modal dialogs; computes formulas and embeds PNG vectors. |
| **PDF (Docker/Server)** | Pre-warmed LibreOffice Calc Pool | **150ms – 300ms** | Recalculates full Excel formula tree; handles complex print setups. |
| **PDF (Local Windows)** | Headless Chrome / Edge Fallback | **200ms – 400ms** | Zero-dependency local developer testing on any Windows machine. |
| **XLSX Export** | Native `openpyxl` stream | **15ms – 30ms** | Retains all native Excel formulas, data validations, and cell styles. |

---

## 4. Error Handling & HTTP Status Codes

| HTTP Status | Error Type | Description / Resolution |
| :--- | :--- | :--- |
| `400 Bad Request` | `TemplateSyntaxException` | Unclosed directive (e.g. missing `{{/each}}` or invalid attribute). |
| `401 Unauthorized` | `HTTPException` | Missing or invalid API Key in request header. |
| `404 Not Found` | `TemplateNotFoundException` | Template specified in `template_code` does not exist in `templates/`. |
| `500 Server Error` | `RenderingException` | Failed to launch conversion engine or process timed out. |
