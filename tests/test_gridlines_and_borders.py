import openpyxl
from openpyxl.styles import Border, Side
from app.renderer.fast_html_renderer import FastHtmlRenderer

def test_default_rendering_has_no_gridlines():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Invoice Title"
    ws["A2"] = "Customer Name"

    html_output = FastHtmlRenderer.render(ws)
    # Ensure the table does not have the show-gridlines class
    assert 'class="report-table"' in html_output
    assert 'class="report-table show-gridlines"' not in html_output
    # Ensure cells do not have an artificial inline border in body
    body_content = html_output.split('<body>')[1]
    assert 'style="border' not in body_content
    assert 'border-top:' not in body_content

def test_explicit_cell_borders_are_rendered():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Header"
    
    thin_side = Side(style="thin", color="000000")
    ws["A1"].border = Border(top=thin_side, bottom=thin_side, left=thin_side, right=thin_side)
    
    html_output = FastHtmlRenderer.render(ws)
    assert 'border-top: 1px solid #000000' in html_output
    assert 'border-bottom: 1px solid #000000' in html_output
    assert 'border-left: 1px solid #000000' in html_output
    assert 'border-right: 1px solid #000000' in html_output

def test_print_options_gridlines_enabled():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Data"
    ws.print_options.gridLines = True
    
    html_output = FastHtmlRenderer.render(ws)
    assert 'class="report-table show-gridlines"' in html_output
