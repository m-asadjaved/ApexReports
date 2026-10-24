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


def test_no_trailing_ghost_rows_and_page_number_print_values():
    """Verify that trailing empty rows in Excel templates are not emitted and page numbers never show 0 0 on print."""
    import re
    from app.engine.compiled_template import TemplateCompiler, CompiledRenderer
    from app.engine.template_processor import TemplateProcessor
    from app.renderer.fast_html_renderer import FastHtmlRenderer

    wb = openpyxl.Workbook()
    ws = wb.active
    
    ws["A1"] = "{{#page-header}}"
    ws["A2"] = "Header Title"
    ws["A3"] = "{{/page-header}}"
    
    ws["A4"] = "{{#each items}}"
    ws["A5"] = "{{name}}: {{amount}}"
    ws["A6"] = "{{/each}}"
    
    ws["A7"] = "{{#page-footer}}"
    ws["A8"] = "Footer: Page {{@page}} of {{@pages}}"
    ws["A9"] = "{{/page-footer}}"
    
    # Touch row 50 so ws.max_row becomes 50
    _ = ws.cell(row=50, column=1)
    
    data = {"items": [{"name": "Item 1", "amount": 100}]}
    
    # 1. Fast Mode
    ast = TemplateCompiler.compile(wb, "TEST_NO_GHOST")
    fast_html = CompiledRenderer.render(ast, data)
    
    fast_trs = re.findall(r'<tr[^>]*>', fast_html)
    assert len(fast_trs) < 15, f"Expected < 15 rows, got {len(fast_trs)}"
    assert 'content: counter(page)' not in fast_html
    assert 'class="print-page-num">1<' in fast_html
    assert 'class="print-total-pages">1<' in fast_html
    
    # 2. Full Mode
    proc = TemplateProcessor(wb)
    out_wb = proc.process(data)
    out_ws = out_wb.active
    full_html = FastHtmlRenderer.render(out_ws)
    
    full_trs = re.findall(r'<tr[^>]*>', full_html)
    assert len(full_trs) < 15, f"Expected < 15 rows, got {len(full_trs)}"
    assert 'content: counter(page)' not in full_html
    assert 'Footer: Page 1 of 1' in full_html


def test_repeat_on_page_outside_each_loop():
    """Verify {{#repeat-on-page}} placed before multi-page each loop repeats table headers on every page."""
    wb = openpyxl.Workbook()
    ws = wb.active

    ws["A1"] = "{{#page-header}}"
    ws["A2"] = "Invoice Report: {{customer}}"
    ws["A3"] = "{{/page-header}}"

    ws["A4"] = "{{#repeat-on-page}}"
    ws["A5"] = "Item Description"
    ws["B5"] = "Amount"
    ws["A6"] = "{{/repeat-on-page}}"

    ws["A7"] = "{{#each items page_break=true}}"
    ws["A8"] = "{{desc}}"
    ws["B8"] = "{{amount}}"
    ws["A9"] = "{{/each}}"

    data = {
        "customer": "Apex Corp",
        "items": [
            {"desc": "Alpha", "amount": 100},
            {"desc": "Beta", "amount": 200},
            {"desc": "Gamma", "amount": 300},
        ]
    }

    # 1. Fast compiler mode
    ast = TemplateCompiler.compile(wb, "TEST_REPEAT_OUTSIDE")
    fast_html = CompiledRenderer.render(ast, data)

    # "Item Description" should repeat on every page (3 times total)
    assert fast_html.count("Item Description") == 3
    assert "Alpha" in fast_html
    assert "Beta" in fast_html
    assert "Gamma" in fast_html

    # Check that pages 2 and 3 have page-break tr class attached to the repeated table header
    assert fast_html.count('class="page-break"') == 2

    # 2. Full compiler mode
    proc = TemplateProcessor(wb)
    out_wb = proc.process(data)
    out_ws = out_wb.active

    full_html = FastHtmlRenderer.render(out_ws)
    assert full_html.count("Item Description") == 3
    assert full_html.count('class="page-break"') == 2
    assert "Alpha" in full_html
    assert "Beta" in full_html
    assert "Gamma" in full_html


def test_repeat_header_inside_each_loop():
    """Verify {{#repeat-header}} inside an each loop does NOT repeat when page_break=false, but repeats when page_break=true."""
    wb = openpyxl.Workbook()
    ws = wb.active

    # Without page_break: header should appear only once
    ws["A1"] = "{{#each items}}"
    ws["A2"] = "{{#repeat-header}}"
    ws["A3"] = "Product Name"
    ws["B3"] = "Qty"
    ws["A4"] = "{{/repeat-header}}"
    ws["A5"] = "{{name}}"
    ws["B5"] = "{{qty}}"
    ws["A6"] = "{{/each}}"

    data = {
        "items": [
            {"name": "Apple", "qty": 10},
            {"name": "Banana", "qty": 20},
            {"name": "Cherry", "qty": 30},
        ]
    }

    ast = TemplateCompiler.compile(wb, "TEST_REPEAT_INSIDE_NO_PB")
    html_no_pb = CompiledRenderer.render(ast, data)
    assert html_no_pb.count("Product Name") == 1

    proc = TemplateProcessor(wb)
    out_wb = proc.process(data)
    full_html_no_pb = FastHtmlRenderer.render(out_wb.active)
    assert full_html_no_pb.count("Product Name") == 1

    # With page_break=true: header should repeat on every item page
    ws["A1"] = "{{#each items page_break=true}}"
    ast_pb = TemplateCompiler.compile(wb, "TEST_REPEAT_INSIDE_WITH_PB")
    html_pb = CompiledRenderer.render(ast_pb, data)
    assert html_pb.count("Product Name") == 3

    proc_pb = TemplateProcessor(wb)
    out_wb_pb = proc_pb.process(data)
    full_html_pb = FastHtmlRenderer.render(out_wb_pb.active)
    assert full_html_pb.count("Product Name") == 3


def test_repeat_on_page_with_explicit_page_break():
    """Verify repeating rows re-emit after an explicit {{page_break}} directive."""
    wb = openpyxl.Workbook()
    ws = wb.active

    ws["A1"] = "{{#repeat}}"
    ws["A2"] = "Col Header 1"
    ws["B2"] = "Col Header 2"
    ws["A3"] = "{{/repeat}}"
    ws["A4"] = "Section 1 Data"
    ws["A5"] = "{{page_break}}"
    ws["A6"] = "Section 2 Data"

    data = {}

    # Fast Mode
    ast = TemplateCompiler.compile(wb, "TEST_REPEAT_EXPLICIT_PB")
    fast_html = CompiledRenderer.render(ast, data)
    assert fast_html.count("Col Header 1") == 2
    assert "Section 1 Data" in fast_html
    assert "Section 2 Data" in fast_html

    # Full Mode
    proc = TemplateProcessor(wb)
    out_wb = proc.process(data)
    full_html = FastHtmlRenderer.render(out_wb.active)
    assert full_html.count("Col Header 1") == 2
    assert "Section 1 Data" in full_html
    assert "Section 2 Data" in full_html


def test_repeat_on_page_linter_validation():
    """Verify TemplateLinter validates matching {{#repeat-on-page}} and flags unmatched/unclosed blocks."""
    # Valid
    wb_valid = openpyxl.Workbook()
    ws_v = wb_valid.active
    ws_v["A1"] = "{{#repeat-on-page}}"
    ws_v["A2"] = "Header"
    ws_v["A3"] = "{{/repeat-on-page}}"
    res_valid = TemplateLinter.lint_workbook(wb_valid, "TEST_LINT_VALID")
    assert res_valid["valid"] is True
    assert len(res_valid["errors"]) == 0

    # Unclosed
    wb_unclosed = openpyxl.Workbook()
    ws_u = wb_unclosed.active
    ws_u["A1"] = "{{#repeat-on-page}}"
    ws_u["A2"] = "Header"
    res_unclosed = TemplateLinter.lint_workbook(wb_unclosed, "TEST_LINT_UNCLOSED")
    assert res_unclosed["valid"] is False
    assert any(e["code"] == "UNCLOSED_BLOCK" for e in res_unclosed["errors"])

    # Unmatched close
    wb_unmatched = openpyxl.Workbook()
    ws_m = wb_unmatched.active
    ws_m["A1"] = "{{/repeat-on-page}}"
    res_unmatched = TemplateLinter.lint_workbook(wb_unmatched, "TEST_LINT_UNMATCHED")
    assert res_unmatched["valid"] is False
    assert any(e["code"] == "UNMATCHED_CLOSE_REPEAT" for e in res_unmatched["errors"])


