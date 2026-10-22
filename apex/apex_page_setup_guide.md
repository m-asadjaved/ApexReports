# Integrating ApexReports with Oracle APEX: Step-by-Step Guide

This guide shows how to integrate your custom **ApexReports Engine** into Oracle APEX pages.

---

## 1. Prerequisites
1. Compile the PL/SQL package in your APEX parsing schema:
   - Run [apex_reports_pkg.pks](file:///d:/Code/ApexReports/apex/apex_reports_pkg.pks)
   - Run [apex_reports_pkg.pkb](file:///d:/Code/ApexReports/apex/apex_reports_pkg.pkb)
2. Ensure network access (ACL) is granted if Oracle DB is calling an external URL:
   ```sql
   BEGIN
       DBMS_NETWORK_ACL_ADMIN.append_host_ace(
           host => 'reports-engine', -- or hostname/IP
           lower_port => 8000,
           upper_port => 8000,
           ace => xs$ace_type(privilege_list => xs$name_list('http'),
                              principal_name => 'APEX_230200', -- your APEX user
                              principal_type => xs_acl.ptype_db)
       );
   END;
   /
   ```

---

## 2. Recipe 1: Direct PDF Download on Button Click

### Step A: Create a Button
- Button Name: `BTN_PRINT_INVOICE`
- Label: `Download Invoice (PDF)`
- Action: `Submit Page`

### Step B: Create an APEX Page Process
- Event: `After Submit`
- Condition: When Button Pressed = `BTN_PRINT_INVOICE`
- Type: `Execute Code` (PL/SQL)
- Code:
```sql
DECLARE
    l_json_data CLOB;
BEGIN
    -- 1. Build JSON data from SQL using Oracle's JSON functions
    SELECT JSON_OBJECT(
        'company' VALUE JSON_OBJECT('name' VALUE 'Apex Solutions Corp.'),
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

    -- 2. Stream PDF directly to browser
    apex_reports_pkg.download_pdf(
        p_template_code => 'INVOICE_REPORT',
        p_json_data     => l_json_data,
        p_filename      => 'Invoice_' || :P10_ORDER_ID || '.pdf'
    );
END;
```

---

## 3. Recipe 2: Instant HTML Preview Modal

### Step A: Create an AJAX Callback (Page Processing)
- Name: `GET_REPORT_PREVIEW`
- PL/SQL Code:
```sql
DECLARE
    l_json_data CLOB;
    l_html      CLOB;
BEGIN
    SELECT JSON_OBJECT(
        'company' VALUE JSON_OBJECT('name' VALUE 'Apex Solutions Corp.'),
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

### Step B: Create a Dynamic Action to Display Preview
- Event: Click on `BTN_PREVIEW`
- Action: `Execute JavaScript Code`
```javascript
apex.server.process('GET_REPORT_PREVIEW', {
    pageItems: '#P10_ORDER_ID'
}, {
    dataType: 'html',
    success: function(data) {
        // Open APEX modal or populate a Static Content region
        $('#report_preview_container').html(data);
        apex.theme.openRegion('REPORT_PREVIEW_MODAL');
    }
});
```

---

## 4. Recipe 3: Emailing PDF via `APEX_MAIL`

```sql
DECLARE
    l_json_data CLOB;
    l_pdf_blob  BLOB;
    l_mail_id   NUMBER;
BEGIN
    -- 1. Fetch JSON data
    l_json_data := '{"invoice_no": "INV-100", "company": {"name": "Acme Inc."}, "items": []}';

    -- 2. Generate PDF BLOB
    l_pdf_blob := apex_reports_pkg.get_report_blob(
        p_template_code => 'INVOICE_REPORT',
        p_format        => 'pdf',
        p_json_data     => l_json_data
    );

    -- 3. Queue Email with attachment
    l_mail_id := apex_mail.send(
        p_to        => 'customer@example.com',
        p_from      => 'billing@apex.corp',
        p_subj      => 'Your Invoice INV-100',
        p_body      => 'Please find attached your invoice.'
    );

    apex_mail.add_attachment(
        p_mail_id    => l_mail_id,
        p_attachment => l_pdf_blob,
        p_filename   => 'Invoice_100.pdf',
        p_mime_type  => 'application/pdf'
    );

    apex_mail.push_queue;
END;
```
