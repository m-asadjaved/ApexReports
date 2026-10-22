import openpyxl
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import PatternFill, Font, Border, Side
from app.renderer.fast_html_renderer import FastHtmlRenderer
from app.engine.template_processor import TemplateProcessor
from app.renderer.chrome_cdp_pool import ChromeCdpPool


def test_cell_is_greater_than_and_less_than():
    """Verify CellIsRule greaterThan and lessThan apply background and font color overlays."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 50
    ws["A2"] = 150
    ws["A3"] = 250

    # Rule: If > 100, light red fill (#FFC7CE) with dark red bold font (#9C0006)
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006", bold=True)
    rule_gt = CellIsRule(operator="greaterThan", formula=["100"], fill=red_fill, font=red_font)

    # Rule: If < 60, light green fill (#C6EFCE) with dark green font (#006100)
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    green_font = Font(color="006100")
    rule_lt = CellIsRule(operator="lessThan", formula=["60"], fill=green_fill, font=green_font)

    ws.conditional_formatting.add("A1:A3", rule_gt)
    ws.conditional_formatting.add("A1:A3", rule_lt)

    html = FastHtmlRenderer.render(ws, "Conditional Formatting Test")

    # A1 (50): matches lessThan 60 -> green
    assert "background-color: #c6efce" in html.lower()
    assert "color: #006100" in html.lower()

    # A2 (150) and A3 (250): match greaterThan 100 -> red + bold
    assert "background-color: #ffc7ce" in html.lower()
    assert "color: #9c0006" in html.lower()
    assert "font-weight: bold" in html


def test_cell_is_between():
    """Verify CellIsRule between correctly matches numbers in range."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["B1"] = 10
    ws["B2"] = 50
    ws["B3"] = 100

    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    yellow_font = Font(color="9C5700", italic=True)
    rule_between = CellIsRule(operator="between", formula=["20", "80"], fill=yellow_fill, font=yellow_font)
    ws.conditional_formatting.add("B1:B3", rule_between)

    html = FastHtmlRenderer.render(ws)

    # B2 (50) is between 20 and 80 -> yellow fill + italic
    assert "background-color: #ffeb9c" in html.lower()
    assert "font-style: italic" in html


def test_text_contains_rule():
    """Verify text match rules like containsText."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "STATUS: OVERDUE"
    ws["A2"] = "STATUS: PAID"
    ws["A3"] = "STATUS: PENDING"

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006", bold=True)
    rule = CellIsRule(operator="equal", formula=['"STATUS: OVERDUE"'], fill=red_fill, font=red_font)
    ws.conditional_formatting.add("A1:A3", rule)

    html = FastHtmlRenderer.render(ws)
    assert "background-color: #ffc7ce" in html.lower()


def test_color_scale_heatmap():
    """Verify 2-color scale interpolates background colors between min and max."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 0
    ws["A2"] = 50
    ws["A3"] = 100

    # 2-color scale from white (#FFFFFF) to green (#00FF00)
    rule = ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF", end_type="num", end_value=100, end_color="00FF00")
    ws.conditional_formatting.add("A1:A3", rule)

    html = FastHtmlRenderer.render(ws)

    # A1 (0) should be #ffffff
    assert "background-color: #ffffff" in html.lower()
    # A3 (100) should be #00ff00
    assert "background-color: #00ff00" in html.lower()
    # A2 (50) should be intermediate green ~ #80ff80
    assert "background-color: #7fff7f" in html.lower() or "background-color: #80ff80" in html.lower()


def test_template_processor_expands_conditional_formatting_in_loops():
    """Verify that conditional formatting defined on a repeating row covers all generated rows."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Item"
    ws["B1"] = "Qty"
    ws["A2"] = "{{#each items}}"
    ws["A3"] = "{{name}}"
    ws["B3"] = "{{qty}}"
    ws["A4"] = "{{/each}}"

    # Rule on B3 (detail row): if Qty < 10, mark red
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    red_font = Font(color="9C0006", bold=True)
    rule = CellIsRule(operator="lessThan", formula=["10"], fill=red_fill, font=red_font)
    ws.conditional_formatting.add("B3:B3", rule)

    proc = TemplateProcessor(wb)
    data = {
        "items": [
            {"name": "Apples", "qty": 100},
            {"name": "Bananas", "qty": 4},   # Low qty!
            {"name": "Cherries", "qty": 50},
            {"name": "Dates", "qty": 2}      # Low qty!
        ]
    }
    out_wb = proc.process(data)
    out_ws = out_wb.active

    # Check that conditional formatting was copied and expanded
    assert len(out_ws.conditional_formatting) > 0

    html = FastHtmlRenderer.render(out_ws)
    # The two low qty items (Bananas=4, Dates=2) must receive the red background and bold text
    assert "background-color: #ffc7ce" in html.lower()
    assert "font-weight: bold" in html


def test_conditional_formatting_pdf_rendering():
    """Verify that Chrome CDP successfully renders conditional formatting into vector PDF."""
    if not ChromeCdpPool.is_available():
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "Metric"
    ws["B1"] = "Score"
    ws["A2"] = "Performance"
    ws["B2"] = 95
    ws["A3"] = "Reliability"
    ws["B3"] = 40

    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    rule = CellIsRule(operator="greaterThan", formula=["80"], fill=green_fill)
    ws.conditional_formatting.add("B2:B3", rule)

    html = FastHtmlRenderer.render(ws)
    assert "background-color: #c6efce" in html.lower()

    pdf_bytes = ChromeCdpPool.render_html_to_pdf(html)
    assert pdf_bytes.startswith(b"%PDF")
    assert len(pdf_bytes) > 1000
