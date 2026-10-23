import datetime
import openpyxl
import pytest

from app.engine.template_processor import TemplateProcessor
from app.renderer.fast_html_renderer import FastHtmlRenderer
from app.engine.compiled_template import TemplateCompiler, CompiledRenderer
from app.engine.template_linter import TemplateLinter


def test_page_header_footer_directives_in_template_processor():
    """Verify {{#page-header}} and {{#page-footer}} generate thead, tfoot, and print_title_rows."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "HeaderFooterTest"

    # Row 1-3: Page Header
    ws["A1"] = "{{#page-header}}"
    ws["A2"] = "ACME SALES INVOICE"
    ws["A3"] = "{{/page-header}}"

    # Row 4-6: Detail items loop
    ws["A4"] = "{{#each items}}"
    ws["A5"] = "{{name}}"
    ws["B5"] = "{{price}}"
    ws["A6"] = "{{/each}}"

    # Row 7-9: Page Footer
    ws["A7"] = "{{#page-footer}}"
    ws["A8"] = "Report Footer: Page {{@page}} of {{@pages}}"
    ws["A9"] = "{{/page-footer}}"

    proc = TemplateProcessor(wb)
    out_wb = proc.process({
        "items": [
            {"name": "Widget A", "price": 10},
            {"name": "Widget B", "price": 20},
        ]
    })
    out_ws = out_wb.active

    # Check print_title_rows was automatically set
    assert out_ws.print_title_rows is not None
    assert "$1:$1" in out_ws.print_title_rows or "1:1" in out_ws.print_title_rows.replace("$", "")

    # Render via FastHtmlRenderer
    html = FastHtmlRenderer.render(out_ws, "Header Footer Test")

    # Verify thead, tbody, and tfoot tags
    assert "<thead>" in html
    assert "</thead>" in html
    assert "ACME SALES INVOICE" in html

    assert "<tbody>" in html
    assert "</tbody>" in html
    assert "Widget A" in html
    assert "Widget B" in html

    assert "<tfoot>" in html
    assert "</tfoot>" in html
    assert "Report Footer: Page" in html

    # Verify CSS print repetition rules
    assert "thead { display: table-header-group; }" in html
    assert "tfoot { display: table-footer-group; }" in html
    assert "tr { break-inside: avoid; page-break-inside: avoid; }" in html


def test_excel_print_title_rows_auto_converted_to_thead():
    """Verify Excel native print_title_rows ($1:$2) is automatically rendered as <thead>."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PrintTitlesTest"

    ws.print_title_rows = "$1:$2"
    ws["A1"] = "Global Header Title"
    ws["A2"] = "Col 1"
    ws["B2"] = "Col 2"
    ws["A3"] = "Row 1 Data"
    ws["B3"] = "Row 1 Value"

    html = FastHtmlRenderer.render(ws, "Print Titles Test")

    assert "<thead>" in html
    assert "Global Header Title" in html
    assert "Col 1" in html
    assert "<tbody>" in html
    assert "Row 1 Data" in html


def test_compiled_template_header_footer_and_system_vars():
    """Verify compiled template cache handles header/footer elements and system variables."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CompiledHF"

    ws["A1"] = "{{#page-header}}"
    ws["A2"] = "Header: {{company.name}}"
    ws["A3"] = "{{/page-header}}"

    ws["A4"] = "{{#each items page_break=true}}"
    ws["A5"] = "Item: {{title}} (Page: {{@page}})"
    ws["A6"] = "{{/each}}"

    ws["A7"] = "{{#page-footer}}"
    ws["A8"] = "Date: {{@date}} | Time: {{@time}} | Total: {{@pages}} pages"
    ws["A9"] = "{{/page-footer}}"

    compiled = TemplateCompiler.compile(wb, "TEST_COMPILED_HF")
    assert len(compiled.header_elements) > 0
    assert len(compiled.footer_elements) > 0

    today_str = datetime.date.today().strftime("%Y-%m-%d")
    data = {
        "company": {"name": "Apex Enterprise"},
        "items": [
            {"title": "Invoice 1"},
            {"title": "Invoice 2"},
        ]
    }

    html = CompiledRenderer.render(compiled, data)

    # Header check
    assert "<thead>" in html
    assert "Header: Apex Enterprise" in html

    # Body check
    assert "<tbody>" in html
    assert "Item: Invoice 1" in html
    assert "Item: Invoice 2" in html

    # Footer check
    assert "<tfoot>" in html
    assert f"Date: {today_str}" in html

    # Verify CSS print counter classes
    assert "report-page-num" in html
    assert "print-page-num" in html
    assert "screen-page-num" in html


def test_universal_system_variables_resolution():
    """Verify @page, @date, @time, page_no, current_date resolve automatically."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Date: {{current_date}}"
    ws["A2"] = "Time: {{current_time}}"
    ws["A3"] = "Page: {{page_no}}"

    proc = TemplateProcessor(wb)
    out_wb = proc.process({})
    out_ws = out_wb.active

    today_str = datetime.date.today().strftime("%Y-%m-%d")
    assert today_str in str(out_ws["A1"].value)
    assert "Page: 1" in str(out_ws["A3"].value)


def test_template_linter_header_footer_validation():
    """Verify TemplateLinter catches unclosed header/footer tags and approves system variables."""
    # 1. Unclosed header block
    wb_unclosed = openpyxl.Workbook()
    ws1 = wb_unclosed.active
    ws1["A1"] = "{{#page-header}}"
    ws1["A2"] = "Title"
    # Missing {{/page-header}}
    res1 = TemplateLinter.lint_workbook(wb_unclosed, "BROKEN_HEADER")
    assert res1["valid"] is False
    assert any(e["code"] == "UNCLOSED_BLOCK" for e in res1["errors"])

    # 2. Mismatched close block
    wb_mismatch = openpyxl.Workbook()
    ws2 = wb_mismatch.active
    ws2["A1"] = "{{#page-header}}"
    ws2["A2"] = "Title"
    ws2["A3"] = "{{/page-footer}}"  # Mismatched close!
    res2 = TemplateLinter.lint_workbook(wb_mismatch, "MISMATCHED_HEADER")
    assert res2["valid"] is False
    assert any(e["code"] == "MISMATCHED_BLOCK_CLOSE" for e in res2["errors"])

    # 3. Valid template with system variables does not trigger UNRESOLVED_VARIABLE
    wb_valid = openpyxl.Workbook()
    ws3 = wb_valid.active
    ws3["A1"] = "{{#page-header}}"
    ws3["A2"] = "Invoice"
    ws3["A3"] = "{{/page-header}}"
    ws3["A4"] = "Page {{@page}} of {{@pages}} on {{@date}} at {{@time}}"
    res3 = TemplateLinter.lint_workbook(wb_valid, "VALID_SYS_VARS", data={})
    assert res3["valid"] is True
    assert not any(w["code"] == "UNRESOLVED_VARIABLE" for w in res3["warnings"])
