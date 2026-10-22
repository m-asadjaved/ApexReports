import openpyxl
from openpyxl.styles import Alignment
from app.engine.template_processor import TemplateProcessor
from app.renderer.fast_html_renderer import FastHtmlRenderer
from app.renderer.browser_pdf_renderer import BrowserPdfRenderer


def test_column_widths_to_colgroup():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "TestDimensions"

    # Set column widths: Col A = 16.0, Col B = 34.0, Col C = 10.0
    ws.column_dimensions["A"].width = 16.0
    ws.column_dimensions["B"].width = 34.0
    ws.column_dimensions["C"].width = 10.0

    ws["A1"] = "Col A"
    ws["B1"] = "Col B"
    ws["C1"] = "Col C"

    html = FastHtmlRenderer.render(ws, "Dimensions Test")

    # Formula: round(w * 7 + 5)
    # A (16.0): 16 * 7 + 5 = 117px
    # B (34.0): 34 * 7 + 5 = 243px
    # C (10.0): 10 * 7 + 5 = 75px
    assert "<colgroup>" in html
    assert '<col style="width: 117px;" />' in html
    assert '<col style="width: 243px;" />' in html
    assert '<col style="width: 75px;" />' in html
    assert "table-layout: fixed;" in html
    assert "min-width: 435px;" in html  # 117 + 243 + 75 = 435px


def test_row_heights_and_default_height():
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.sheet_format.defaultRowHeight = 18.0
    ws.row_dimensions[1].height = 36.5
    ws.row_dimensions[2].height = 22.0

    ws["A1"] = "Header"
    ws["A2"] = "Subheader"
    ws["A3"] = "Normal Row"

    html = FastHtmlRenderer.render(ws, "Row Heights Test")

    # Verify default row height in CSS
    assert "table.report-table tr { height: 18.0pt; }" in html
    # Verify explicit row heights inline on <tr>
    assert '<tr style="height: 36.5pt;">' in html
    assert '<tr style="height: 22.0pt;">' in html
    # Row 3 uses default height (no custom inline height on <tr>)
    assert '<tr>' in html


def test_hidden_rows_and_hidden_columns():
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.column_dimensions["A"].width = 20.0
    ws.column_dimensions["B"].hidden = True
    ws.row_dimensions[2].hidden = True

    ws["A1"] = "Visible 1"
    ws["B1"] = "Hidden Col"
    ws["A2"] = "Hidden Row"
    ws["B2"] = "Hidden Both"

    html = FastHtmlRenderer.render(ws, "Hidden Test")

    # Hidden column in colgroup
    assert '<col style="display: none; width: 0px;" />' in html
    # Cell in hidden column
    assert "display: none;" in html
    # Hidden row
    assert '<tr style="display: none;">' in html


def test_cloned_rows_preserve_height():
    wb = openpyxl.Workbook()
    ws = wb.active

    ws["A1"] = "{{#each items}}"
    ws["A2"] = "{{name}}"
    ws["A3"] = "{{/each}}"

    # Give row 2 in template explicit height 28.5pt
    ws.row_dimensions[2].height = 28.5

    proc = TemplateProcessor(wb)
    out_wb = proc.process({
        "items": [
            {"name": "Item 1"},
            {"name": "Item 2"},
            {"name": "Item 3"}
        ]
    })
    out_ws = out_wb.active

    assert out_ws.row_dimensions[1].height == 28.5
    assert out_ws.row_dimensions[2].height == 28.5
    assert out_ws.row_dimensions[3].height == 28.5

    html = FastHtmlRenderer.render(out_ws)
    assert html.count('style="height: 28.5pt;"') == 3


def test_cell_vertical_alignment_and_wrapping():
    wb = openpyxl.Workbook()
    ws = wb.active

    ws["A1"] = "Centered Text"
    ws["A1"].alignment = Alignment(vertical="center", horizontal="center")

    ws["B1"] = "Line 1\nLine 2"
    ws["B1"].alignment = Alignment(wrap_text=True)

    html = FastHtmlRenderer.render(ws)

    # In CSS, Excel 'center' vertical-alignment must map to 'middle'
    assert "vertical-align: middle" in html
    assert "text-align: center" in html
    # Wrap text
    assert "white-space: pre-wrap" in html
    # Newline in cell content converted to <br/>
    assert "Line 1<br/>Line 2" in html


def test_pdf_generation_with_dimensions():
    wb = openpyxl.Workbook()
    ws = wb.active

    ws.column_dimensions["A"].width = 25.0
    ws.column_dimensions["B"].width = 30.0
    ws.row_dimensions[1].height = 40.0

    ws["A1"] = "Report Title"
    ws["B1"] = "Status"
    ws["A2"] = "Widget 1"
    ws["B2"] = "Active"

    pdf_bytes = BrowserPdfRenderer.convert_worksheet_to_pdf(ws, "Test Dimension PDF")
    assert pdf_bytes.startswith(b"%PDF")
    assert len(pdf_bytes) > 1000
