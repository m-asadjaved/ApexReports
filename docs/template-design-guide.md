# Template Design Guide

This guide details how to author high-performance report templates in **Microsoft Excel (.xlsx)** for the ApexReports engine. It follows a **code-first approach**, focusing on data structures, cell tag grammar, formula expansion mechanics, and Oracle Reports 6i parity.

---

## 1. Golden Principles for Template Designers

1. **Native Excel Formatting First**: Never format values using template string masks (e.g. do not write `{{format(amount, '$#,##0.00')}}`). Instead, format the Excel cell using Excel's native Number/Currency dialog. The engine injects raw numbers, preserving formatting across Excel, PDF, and HTML.
2. **First Worksheet Rule**: The reporting engine processes the **first worksheet** of the template workbook (`active_sheet`).
3. **Keep Directives on Dedicated Rows**: For readability and maintainability, place block controls (`{{#each items}}` and `{{/each}}`) on their own separate rows. The engine removes directive rows during rendering without leaving empty gaps.
4. **Formula Self-Referencing**: When writing a summary formula for a repeating block (such as `=SUM(...)` or `=COUNTA(...)`), write the formula referencing the single template row (e.g. `=SUM(D6:D6)`). The engine auto-expands the range across all generated rows upon rendering.

---

## 2. Complete Tag Syntax & Grammar Reference

| Directive / Tag | Cell Syntax Example | Description |
| :--- | :--- | :--- |
| **Scalar Field** | `{{customer.name}}` | Injects string, number, or boolean from the JSON context. |
| **Fallback Value** | `{{notes \|\| 'N/A'}}` | Returns default value if field is null, empty, or undefined. |
| **Dotted Path** | `{{order.shipping.address.city}}` | Accesses nested JSON objects of arbitrary depth. |
| **Repeating Rows (Down)** | `{{#each items}}` ... `{{/each}}` | Repeats enclosed rows once for each array element. |
| **Page-Break Repeat** | `{{#each depts page_break=true}}` | Triggers a fresh print page break for each master record. |
| **Matrix Columns (Across)** | `{{#each-col quarters}} {{name}} {{/each-col}}` | Horizontally expands columns for 2D cross-tab reports. |
| **Format Trigger (If)** | `{{#if discount > 0}}` ... `{{/if}}` | Conditionally displays or suppresses enclosed rows. |
| **Manual Page Break** | `{{#page-break}}` | Explicitly forces a print page break after this row. |
| **QR Code (URL/ID)** | `{{@qr 'https://corp.com/' + id width=90 height=90}}` | Generates embedded vector QR code resized to dimensions. |
| **1D Barcode** | `{{@barcode sku format=code128 width=180 height=45}}` | Generates Code128, EAN13, or Code39 barcode. |
| **Barcode with Text** | `{{@barcode sku text=true width=180 height=50}}` | Generates barcode with human-readable text beneath bars. |
| **Dynamic Image / Logo** | `{{@image company.logo_url width=120 height=40}}` | Downloads URL or decodes Base64 into floating image. |

---

## 3. JSON Payload Structures & Mapping

### 3.1 Flat Document (Invoice / Receipt)
Use when the report has top-level header fields and a single repeating table:

```json
{
  "company": { "name": "Acme Global Solutions" },
  "invoice_no": "INV-10024",
  "invoice_date": "2026-10-22",
  "customer": { "name": "Oracle Corporation" },
  "items": [
    { "item_code": "LIC-01", "description": "APEX Extension", "qty": 10, "unit_price": 450.00 },
    { "item_code": "SUP-02", "description": "24/7 SLA Support", "qty": 1, "unit_price": 1500.00 }
  ]
}
```

**Template Layout:**
```text
Row 1: | TAX INVOICE | | | {{@barcode invoice_no text=true width=160 height=40}} |
Row 2: | Company: {{company.name}} | | | Date: {{invoice_date}} |
Row 3: | Bill To: {{customer.name}} | | | Invoice No: {{invoice_no}} |
Row 4: | Item Code | Description | Quantity | Unit Price | Line Total |
Row 5: {{#each items}}
Row 6: | {{item_code}} | {{description}} | {{qty}} | {{unit_price}} | =C6*D6 |
Row 7: {{/each}}
Row 8: | | | | Subtotal: | =SUM(E6:E6) |
Row 9: | Scan to Verify: | {{@qr 'https://verify.corp.com/inv/' + invoice_no width=80 height=80}} | | | |
```

---

### 3.2 Nested Master-Detail (Departments $\rightarrow$ Employees)
Use when a master record contains an array of detail records:

```json
{
  "departments": [
    {
      "dept_name": "ACCOUNTING",
      "loc": "NEW YORK",
      "employees": [
        { "empno": 7782, "ename": "CLARK", "job": "MANAGER", "sal": 2450.00 },
        { "empno": 7839, "ename": "KING", "job": "PRESIDENT", "sal": 5000.00 }
      ]
    },
    {
      "dept_name": "RESEARCH",
      "loc": "DALLAS",
      "employees": [
        { "empno": 7369, "ename": "SMITH", "job": "CLERK", "sal": 800.00 },
        { "empno": 7566, "ename": "JONES", "job": "MANAGER", "sal": 2975.00 }
      ]
    }
  ]
}
```

**Template Layout:**
```text
Row 1: {{#each departments page_break=true}}
Row 2: | DEPARTMENT: {{dept_name}} (Location: {{loc}}) | | | |
Row 3: | Emp No | Employee Name | Job Title | Salary |
Row 4: {{#each employees}}
Row 5: | {{empno}} | {{ename}} | {{job}} | {{sal}} |
Row 6: {{/each}}
Row 7: | Total Staff: =COUNTA(A5:A5) | | Department Total: | =SUM(D5:D5) |
Row 8: {{/each}}
```

---

### 3.3 Matrix / 2D Cross-Tab (Horizontal Column Expansion)
Use when report columns dynamically generate from data (e.g. Quarters or Regions across the top):

```json
{
  "quarters": [
    { "name": "Q1 2026" },
    { "name": "Q2 2026" },
    { "name": "Q3 2026" },
    { "name": "Q4 2026" }
  ],
  "products": [
    { "pname": "Enterprise ERP", "sales": 125000 },
    { "pname": "Cloud Storage", "sales": 45000 }
  ]
}
```

**Template Layout:**
```text
Row 1: | Product | {{#each-col quarters}} {{name}} {{/each-col}} | Total Sales |
Row 2: {{#each products}}
Row 3: | {{pname}} | {{sales}} | =SUM(B3:B3) |
Row 4: {{/each}}
```
*Result:* Column B expands into columns B, C, D, and E (one per quarter), sliding the "Total Sales" column to Column F.

---

## 4. Formula Design & Auto-Expansion Rules

### How Formula Auto-Expansion Operates
When detail rows repeat $N$ times:
1. **Self-Row Formulas (Calculated Columns / `CF_`)**:
   - Template: `=C6*D6` on Row 6.
   - Output: In row 6 $\rightarrow$ `=C6*D6`, in row 7 $\rightarrow$ `=C7*D7`, in row 8 $\rightarrow$ `=C8*D8`.
2. **Aggregations & Summaries (`CS_`)**:
   - Write the formula referencing the single template detail row:
     `=SUM(E6:E6)` or `=COUNTA(A6:A6)` or `=AVERAGE(D6:D6)`.
   - When 10 rows are generated (rows 6 to 15), the summary row shifts to row 16 and updates to:
     `=SUM(E6:E15)` or `=COUNTA(A6:A15)` or `=AVERAGE(D6:D15)`.
3. **Group Summary Resets (`RESET AT GROUP`)**:
   - In nested Master-Detail loops, any summary formula placed inside an outer loop iteration calculates **only for that specific group**.
   - Department 1's total calculates `=SUM(D3:D5)`, while Department 2's total calculates `=SUM(D9:D11)`.

### Supported Formula Functions
- **Aggregation:** `SUM`, `COUNTA`, `COUNT`, `COUNTBLANK`, `AVERAGE`, `MIN`, `MAX`, `PRODUCT`, `SUBTOTAL`
- **Arithmetic:** `+`, `-`, `*`, `/`, `%`, `^`, parentheses `(...)`
- **Rounding:** `ROUND(val, digits)`, `ROUNDUP(...)`, `ROUNDDOWN(...)`, `ABS(...)`, `INT(...)`, `MOD(...)`
- **Text & Strings:** `CONCAT(...)`, `CONCATENATE(...)`
- **Dates:** `TODAY()`, `NOW()`

---

## 5. Visual Layout & Print Setup in Excel

To ensure professional printable PDF output, configure these native settings in your template before saving:

### 1. Repeat Column Headers across Page Breaks
In Excel:
- Go to **Page Layout** tab $\rightarrow$ **Print Titles** $\rightarrow$ **Rows to repeat at top**.
- Select the table header row (e.g. `$4:$4`).
- *Result:* When a table breaks across multiple pages, the column headers automatically print at the top of every page in PDF.

### 2. Paper Size and Orientation
In Excel:
- Go to **Page Layout** tab $\rightarrow$ **Orientation** $\rightarrow$ Select `Portrait` or `Landscape`.
- **Size** $\rightarrow$ Select `A4` or `Letter`.
- **Margins** $\rightarrow$ Select `Narrow` (recommended for financial tables).

### 3. Merged Cells Best Practice
- **Allowed:** Merging cells within the same row horizontally (e.g. `A2:D2` for a department header banner).
- **Prohibited:** Merging across rows that contain repeating directives (e.g. do not merge `A4:A6` over a loop block).
