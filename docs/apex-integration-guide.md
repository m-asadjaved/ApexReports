# Oracle APEX Integration Guide

A complete, production-ready implementation manual for integrating ApexReports into **Oracle APEX 20.x, 21.x, 22.x, 23.x, and 24.x**.

---

## 1. Architecture: The APEX-to-Engine Bridge

```
┌─────────────────────────────────┐
│     Oracle Database (APEX)      │
│                                 │
│  1. SQL Data Query              │
│     JSON_OBJECT / JSON_ARRAYAGG │
│               │                 │
│               ▼                 │
│  2. PL/SQL Package              │
│     apex_reports_pkg            │
│               │                 │
└───────────────┼─────────────────┘
                │
                │ HTTP POST (JSON Payload)
                ▼
┌─────────────────────────────────┐
│    ApexReports Microservice     │
│                                 │
│  3. Template Processor          │
│     openpyxl + Formula Engine   │
│               │                 │
│               ▼                 │
│  4. Rendering Engine            │
│     LibreOffice / Chrome        │
└───────────────┼─────────────────┘
                │
                │ Binary PDF / HTML Stream
                ▼
┌─────────────────────────────────┐
│       End-User Browser          │
│  • Instant Download (PDF)       │
│  • Inline Modal Preview (HTML)  │
└─────────────────────────────────┘
```

---

## 2. Step 1: Package Installation & Network ACL

### 2.1 Compile the PL/SQL Wrapper Package
Compile the package into your APEX workspace parsing schema:
1. Run [`apex/apex_reports_pkg.pks`](file:///d:/Code/ApexReports/apex/apex_reports_pkg.pks) (Specification)
2. Run [`apex/apex_reports_pkg.pkb`](file:///d:/Code/ApexReports/apex/apex_reports_pkg.pkb) (Body)

### 2.2 Configure Oracle Network Access Control (ACL)
If Oracle Database calls the reporting microservice over HTTP/HTTPS, grant network access:

```sql
BEGIN
    DBMS_NETWORK_ACL_ADMIN.append_host_ace(
        host       => 'localhost',       -- or your Docker container / VM hostname
        lower_port => 8000,
        upper_port => 8000,
        ace        => xs$ace_type(
            privilege_list => xs$name_list('http'),
            principal_name => 'APEX_230200', -- replace with your DB APEX schema or parsing schema
            principal_type => xs_acl.ptype_db
        )
    );
END;
/
```

---

## 3. SQL Data Construction with Native JSON

Oracle Database 12c+ provides powerful, native SQL JSON generation functions. Use `JSON_OBJECT` and `JSON_ARRAYAGG` to serialize relational data directly into the exact hierarchical format expected by your Excel template.

### Example: Department & Employees Hierarchy
```sql
SELECT JSON_OBJECT(
    'departments' VALUE JSON_ARRAYAGG(
        JSON_OBJECT(
            'dept_name'  VALUE d.dname,
            'loc'        VALUE d.loc,
            'employees'  VALUE (
                SELECT JSON_ARRAYAGG(
                    JSON_OBJECT(
                        'empno' VALUE e.empno,
                        'ename' VALUE e.ename,
                        'job'   VALUE e.job,
                        'sal'   VALUE e.sal
                    )
                )
                FROM emp e WHERE e.deptno = d.deptno
            )
        )
    )
)
FROM dept d;
```

---

## 4. APEX Integration Recipes

### Recipe 1: Direct PDF Download on Button Click

#### Page Designer Setup:
1. **Create Button:**
   - Name: `BTN_DOWNLOAD_PDF`
   - Label: `Download Invoice (PDF)`
   - Action: `Submit Page`
2. **Create Page Process (After Submit):**
   - Point: `Processing` (After Submit)
   - Condition: When Button Pressed = `BTN_DOWNLOAD_PDF`
   - Type: `Execute Code` (PL/SQL)

```sql
DECLARE
    l_json_data CLOB;
BEGIN
    -- 1. Query relational data into JSON
    SELECT JSON_OBJECT(
        'company' VALUE JSON_OBJECT('name' VALUE 'Acme Corporation'),
        'invoice_no' VALUE 'INV-' || TO_CHAR(o.order_id),
        'invoice_date' VALUE TO_CHAR(o.order_date, 'YYYY-MM-DD'),
        'customer' VALUE JSON_OBJECT('name' VALUE c.cust_first_name || ' ' || c.cust_last_name),
        'items' VALUE (
            SELECT JSON_ARRAYAGG(
                JSON_OBJECT(
                    'item_code'   VALUE p.product_id,
                    'description' VALUE p.product_name,
                    'qty'         VALUE oi.quantity,
                    'unit_price'  VALUE oi.unit_price
                )
            )
            FROM demo_order_items oi
            JOIN demo_product_info p ON p.product_id = oi.product_id
            WHERE oi.order_id = o.order_id
        )
    )
    INTO l_json_data
    FROM demo_orders o
    JOIN demo_customers c ON c.customer_id = o.customer_id
    WHERE o.order_id = :P10_ORDER_ID;

    -- 2. Stream PDF directly to browser (triggers download prompt)
    apex_reports_pkg.download_pdf(
        p_template_code => 'INVOICE_REPORT',
        p_json_data     => l_json_data,
        p_filename      => 'Invoice_' || :P10_ORDER_ID || '.pdf'
    );
END;
```

---

### Recipe 2: Instant HTML Preview in an APEX Modal Dialog

Allows users to preview reports in $< 50\text{ms}$ inside a modal region before printing.

#### Step A: Create an AJAX Callback (Page Processing)
- Name: `GET_REPORT_HTML`
- Process Type: `Execute Code` (PL/SQL)

```sql
DECLARE
    l_json_data CLOB;
    l_html      CLOB;
BEGIN
    SELECT JSON_OBJECT(
        'company' VALUE JSON_OBJECT('name' VALUE 'Acme Corporation'),
        'invoice_no' VALUE 'INV-' || TO_CHAR(o.order_id),
        'invoice_date' VALUE TO_CHAR(o.order_date, 'YYYY-MM-DD'),
        'customer' VALUE JSON_OBJECT('name' VALUE c.cust_first_name || ' ' || c.cust_last_name),
        'items' VALUE (
            SELECT JSON_ARRAYAGG(
                JSON_OBJECT(
                    'item_code'   VALUE p.product_id,
                    'description' VALUE p.product_name,
                    'qty'         VALUE oi.quantity,
                    'unit_price'  VALUE oi.unit_price
                )
            )
            FROM demo_order_items oi
            JOIN demo_product_info p ON p.product_id = oi.product_id
            WHERE oi.order_id = o.order_id
        )
    )
    INTO l_json_data
    FROM demo_orders o
    JOIN demo_customers c ON c.customer_id = o.customer_id
    WHERE o.order_id = :P10_ORDER_ID;

    l_html := apex_reports_pkg.get_html_preview(
        p_template_code => 'INVOICE_REPORT',
        p_json_data     => l_json_data
    );

    sys.htp.p(l_html);
END;
```

#### Step B: Create an Inline Drawer / Modal Dialog
- Region Name: `PREVIEW_MODAL`
- Template: `Inline Drawer` or `Modal Dialog`
- Static ID: `previewModal`
- Content: Add a Static Content Region with Static ID: `previewContainer`

#### Step C: Dynamic Action on Preview Button
- Event: Click on `BTN_PREVIEW`
- Action: `Execute JavaScript Code`

```javascript
apex.server.process('GET_REPORT_HTML', {
    pageItems: '#P10_ORDER_ID'
}, {
    dataType: 'html',
    success: function(htmlContent) {
        $('#previewContainer').html(htmlContent);
        apex.theme.openRegion('previewModal');
    }
});
```

---

### Recipe 3: Automated Email Attachment with `APEX_MAIL`

Generates the vector PDF in the background and attaches it to an email without browser interaction:

```sql
DECLARE
    l_json_data CLOB;
    l_pdf_blob  BLOB;
    l_mail_id   NUMBER;
BEGIN
    -- 1. Construct JSON
    l_json_data := '{"invoice_no": "INV-10024", "company": {"name": "Acme Inc."}, "items": []}';

    -- 2. Fetch binary PDF BLOB from microservice
    l_pdf_blob := apex_reports_pkg.get_report_blob(
        p_template_code => 'INVOICE_REPORT',
        p_format        => 'pdf',
        p_json_data     => l_json_data,
        p_filename      => 'Invoice_10024.pdf'
    );

    -- 3. Queue Email with attachment
    l_mail_id := apex_mail.send(
        p_to        => 'finance@customer.com',
        p_from      => 'billing@acme.corp',
        p_subj      => 'Your Invoice INV-10024',
        p_body      => 'Please find attached your invoice PDF.'
    );

    apex_mail.add_attachment(
        p_mail_id    => l_mail_id,
        p_attachment => l_pdf_blob,
        p_filename   => 'Invoice_10024.pdf',
        p_mime_type  => 'application/pdf'
    );

    apex_mail.push_queue;
END;
```
