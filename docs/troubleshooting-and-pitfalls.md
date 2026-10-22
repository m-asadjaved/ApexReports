# Troubleshooting & Common Pitfalls

A senior engineering guide to diagnosing, resolving, and preventing errors in report templates, payloads, and rendering pipelines.

---

## 1. Deep-Dive Issue Analysis

### 🐛 Issue 1: "Swallowed Rows" (Previous Groups Disappear in Master-Detail)

#### Symptom:
When running a Master-Detail report (e.g. Departments $\rightarrow$ Employees), only the last department appears in the report. Previous departments show empty cells or missing rows.

#### Root Cause:
In Microsoft Excel, merging cells vertically across multiple rows that contain loop directives (e.g. merging `A1:A5` where Row 1 is `{{#each departments}}` and Row 4 is `{{#each employees}}`) causes openpyxl to treat the entire output range as one giant merged cell. In spreadsheet engines, merging cells erases all content except the top-left cell.

#### Before vs After:

❌ **Broken Template:**
```text
Row 1: [ Merged A1:D3 Banner containing {{#each departments}} ]
Row 4: {{#each employees}}
Row 5: | {{empno}} | {{ename}} | {{job}} | {{sal}} |
Row 6: {{/each}}
```

✅ **Fixed Template:**
```text
Row 1: {{#each departments page_break=true}}
Row 2: | Merged A2:D2 Banner: DEPARTMENT: {{dept_name}} |
Row 3: | Emp No | Employee Name | Job Title | Salary |
Row 4: {{#each employees}}
Row 5: | {{empno}} | {{ename}} | {{job}} | {{sal}} |
Row 6: {{/each}}
Row 7: | Department Total: | =SUM(D5:D5) |
Row 8: {{/each}}
```
> [!IMPORTANT]
> Never merge cells vertically across different rows that contain `{{#each}}` or `{{#if}}` directives. Only merge horizontally within the **same single row** (e.g. `A2:D2`).

---

### 🐛 Issue 2: "Empty Report" (HTTP 200 OK but Table Has No Rows)

#### Symptom:
The API returns status `200 OK`, but the PDF or HTML output displays an empty table with no data.

#### Root Cause:
Payload nesting mismatch. The Excel template has an outer loop directive like `{{#each invoices}}` on Row 1, but the JSON payload passed from Oracle APEX contains a flat object without the `invoices` array wrapper.

#### Before vs After:

❌ **Mismatched Payload:**
```json
{
  "company": { "name": "Acme Inc." },
  "invoice_no": "INV-10024",
  "items": [ ... ]
}
```
*(If template starts with `{{#each invoices}}`, `invoices` evaluates to `null`/`[]` and skips the report).*

✅ **Matching Payload:**
```json
{
  "invoices": [
    {
      "company": { "name": "Acme Inc." },
      "invoice_no": "INV-10024",
      "items": [ ... ]
    }
  ]
}
```
*Alternatively, if your report is always a single invoice, remove `{{#each invoices}}` from Row 1 of your Excel template.*

---

### 🐛 Issue 3: Missing Invoice Number in QR Codes or Barcodes

#### Symptom:
The QR code generates and scans, but the invoice number or ID is missing from the URL (e.g. scans as `https://verify.corp.com/inv/` instead of `https://verify.corp.com/inv/10024`).

#### Root Cause:
In Python, concatenating a string with a number using `+` (`'https://...' + 10024`) raises a `TypeError`. In earlier versions, this resulted in an empty string fallback.

#### Before vs After:

❌ **Fragile Syntax:**
```text
{{@qr 'https://verify.corp.com/inv/' + invoice_no}}
```
*(If `invoice_no` is an Oracle DB sequence number like `10024`, standard Python string concatenation fails).*

✅ **Robust Approaches:**
```text
1. Automatic Coercion (Engine converts to string):
{{@qr 'https://verify.corp.com/inv/' + invoice_no width=90 height=90}}

2. Direct Token Interpolation (Recommended):
{{@qr https://verify.corp.com/inv/{{invoice_no}} width=90 height=90}}
```

---

### 🐛 Issue 4: Formulas Outputted as Literal Strings (`=SUM(D5:D15)`)

#### Symptom:
The report displays `=SUM(D5:D15)` or `=COUNTA(A5:A12)` in plain text rather than the calculated sum or count.

#### Root Cause:
1. In `openpyxl`, formula results are not pre-calculated unless opened by Excel.
2. In the Fast HTML preview engine, only supported functions (`SUM`, `COUNTA`, `COUNT`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `ROUND`, `IF`) are evaluated. If an unsupported or misspelled function (e.g. `=SMM(...)` or custom VBA macro) is used, it falls back to the formula string.

#### Prevention:
- Ensure aggregation formulas use standard supported functions: `=SUM(...)`, `=COUNTA(...)`, `=AVERAGE(...)`.
- To evaluate exotic Excel formulas (e.g. `XLOOKUP`), export to `pdf` via the LibreOffice engine or download as `.xlsx`.

---

### 🐛 Issue 5: Unclosed Directive Blocks (`TemplateSyntaxException`)

#### Symptom:
API returns `HTTP 400 Bad Request` with message:
`TemplateSyntaxException: Unclosed 'each_start' block starting around line 4`.

#### Root Cause:
Every `{{#each list}}` must have a matching `{{/each}}`, and every `{{#if cond}}` must have a matching `{{/if}}`. If a closing tag is missing or misspelled (e.g. `{/each}` with only single curly braces), parsing halts.

#### Checklist:
- Verify curly brace pairs: `{{` and `}}` (always two braces, never one or three).
- Ensure closing directives match opening directives (`{{#each}}` pairs with `{{/each}}`, NOT `{{/end}}`).

---

### 🐛 Issue 6: Excel Temporary Lock Files (`~$...`) in Templates Directory

#### Symptom:
When editing a template in Microsoft Excel while the engine is running, a temporary lock file (e.g. `~$INVOICE_REPORT.xlsx`) appears in `templates/`, causing ZIP extraction errors or duplicate listings in `GET /api/v1/templates`.

#### Solution:
The repository `.gitignore` ignores `~$*`. In production, never upload lock files to the templates repository. Always close Microsoft Excel before deploying or committing templates.

---

## 2. Senior Engineer "Do's and Don'ts" Matrix

| Category | ✅ What to Do | ❌ What NOT to Do |
| :--- | :--- | :--- |
| **Cell Formatting** | Format cells using Excel's native Number, Currency, and Date formats. | Do not use string format masks inside tags (e.g. avoid `{{format(val)}}`). |
| **Formula Ranges** | Reference the single template detail row: `=SUM(D6:D6)` or `=COUNTA(A6:A6)`. | Do not guess future row ranges like `=SUM(D6:D9999)` (creates empty rows). |
| **Merged Cells** | Merge cells horizontally within the same row (e.g. `A2:D2`). | Never merge across rows that contain repeating directives. |
| **Directives** | Place `{{#each}}` and `{{/each}}` on their own distinct rows. | Do not mix `{{#each}}` on the same row with data columns. |
| **Page Breaks** | Use `{{#each depts page_break=true}}` or `{{#page-break}}`. | Do not rely on manual empty rows to push content to next page. |
| **Column Titles** | Use Excel's native **Print Titles** (`Repeat rows at top: $4:$4`). | Do not duplicate header tables manually inside repeating blocks. |
| **Media Sizing** | Specify exact width and height: `{{@qr url width=90 height=90}}`. | Do not omit dimensions if your layout relies on strict column bounding boxes. |
| **API Delivery** | Use `response_mode: "binary"` for instant streaming; use `"json"` for PL/SQL storage. | Do not parse binary streams as JSON strings in PL/SQL. |
