CREATE OR REPLACE PACKAGE BODY apex_reports_pkg AS

    FUNCTION build_request_payload (
        p_template_code IN VARCHAR2,
        p_format        IN VARCHAR2,
        p_json_data     IN CLOB,
        p_filename      IN VARCHAR2
    ) RETURN CLOB IS
        l_payload CLOB;
    BEGIN
        -- Format outer envelope
        l_payload := '{'
            || '"template_code": "' || apex_escape.json_literal(p_template_code) || '",'
            || '"format": "'        || apex_escape.json_literal(p_format) || '",'
            || '"response_mode": "binary",'
            || '"filename": "'      || apex_escape.json_literal(p_filename) || '",'
            || '"data": '           || p_json_data
            || '}';
        RETURN l_payload;
    END build_request_payload;

    FUNCTION get_report_blob (
        p_template_code IN VARCHAR2,
        p_format        IN VARCHAR2 DEFAULT 'pdf',
        p_json_data     IN CLOB,
        p_filename      IN VARCHAR2 DEFAULT NULL,
        p_engine_url    IN VARCHAR2 DEFAULT NULL
    ) RETURN BLOB IS
        l_url      VARCHAR2(500) := NVL(p_engine_url, g_default_engine_url);
        l_body     CLOB;
        l_response BLOB;
        l_fname    VARCHAR2(255) := NVL(p_filename, p_template_code || '.' || p_format);
    BEGIN
        l_body := build_request_payload(p_template_code, p_format, p_json_data, l_fname);

        apex_web_service.g_request_headers.delete();
        apex_web_service.g_request_headers(1).name  := 'Content-Type';
        apex_web_service.g_request_headers(1).value := 'application/json';

        l_response := apex_web_service.make_rest_request_b(
            p_url         => l_url,
            p_http_method => 'POST',
            p_body        => l_body
        );

        IF apex_web_service.g_status_code NOT BETWEEN 200 AND 299 THEN
            raise_application_error(-20001, 'Reporting Engine Error (HTTP ' || apex_web_service.g_status_code || ')');
        END IF;

        RETURN l_response;
    END get_report_blob;

    FUNCTION get_html_preview (
        p_template_code IN VARCHAR2,
        p_json_data     IN CLOB,
        p_engine_url    IN VARCHAR2 DEFAULT NULL
    ) RETURN CLOB IS
        l_url      VARCHAR2(500) := NVL(p_engine_url, g_default_engine_url);
        l_body     CLOB;
        l_response CLOB;
    BEGIN
        l_body := build_request_payload(p_template_code, 'html', p_json_data, p_template_code || '.html');

        apex_web_service.g_request_headers.delete();
        apex_web_service.g_request_headers(1).name  := 'Content-Type';
        apex_web_service.g_request_headers(1).value := 'application/json';

        l_response := apex_web_service.make_rest_request(
            p_url         => l_url,
            p_http_method => 'POST',
            p_body        => l_body
        );

        IF apex_web_service.g_status_code NOT BETWEEN 200 AND 299 THEN
            raise_application_error(-20002, 'Reporting Engine Error (HTTP ' || apex_web_service.g_status_code || '): ' || l_response);
        END IF;

        RETURN l_response;
    END get_html_preview;

    PROCEDURE download_pdf (
        p_template_code IN VARCHAR2,
        p_json_data     IN CLOB,
        p_filename      IN VARCHAR2 DEFAULT 'report.pdf',
        p_engine_url    IN VARCHAR2 DEFAULT NULL
    ) IS
        l_pdf_blob BLOB;
    BEGIN
        l_pdf_blob := get_report_blob(
            p_template_code => p_template_code,
            p_format        => 'pdf',
            p_json_data     => p_json_data,
            p_filename      => p_filename,
            p_engine_url    => p_engine_url
        );

        -- Send binary HTTP response to browser
        sys.owa_util.mime_header('application/pdf', FALSE);
        sys.htp.p('Content-Length: ' || sys.dbms_lob.getlength(l_pdf_blob));
        sys.htp.p('Content-Disposition: inline; filename="' || p_filename || '"');
        sys.owa_util.http_header_close;
        sys.wpg_docload.download_file(l_pdf_blob);

        apex_application.stop_apex_engine;
    END download_pdf;

END apex_reports_pkg;
/
