CREATE OR REPLACE PACKAGE apex_reports_pkg AS
    /**
     * ApexReports PL/SQL Client for Oracle APEX
     * Connects Oracle APEX to the ApexReports Engine microservice.
     */

    -- Default base URL for the reporting microservice
    g_default_engine_url VARCHAR2(500) := 'http://localhost:8000/api/v1/reports/render';

    /**
     * Streams the generated PDF directly to the browser for download or inline viewing.
     * Use inside an APEX "Execute Code" process or button dynamic action.
     */
    PROCEDURE download_pdf (
        p_template_code IN VARCHAR2,
        p_json_data     IN CLOB,
        p_filename      IN VARCHAR2 DEFAULT 'report.pdf',
        p_engine_url    IN VARCHAR2 DEFAULT NULL
    );

    /**
     * Fetches instant HTML preview as a CLOB.
     * Ideal for AJAX callbacks, dynamic regions, or modal dialogs.
     */
    FUNCTION get_html_preview (
        p_template_code IN VARCHAR2,
        p_json_data     IN CLOB,
        p_engine_url    IN VARCHAR2 DEFAULT NULL
    ) RETURN CLOB;

    /**
     * Returns generated report as a BLOB (for PDF, XLSX, or HTML).
     * Useful for storing in database tables or attaching to emails via APEX_MAIL.
     */
    FUNCTION get_report_blob (
        p_template_code IN VARCHAR2,
        p_format        IN VARCHAR2 DEFAULT 'pdf',
        p_json_data     IN CLOB,
        p_filename      IN VARCHAR2 DEFAULT NULL,
        p_engine_url    IN VARCHAR2 DEFAULT NULL
    ) RETURN BLOB;

END apex_reports_pkg;
/
