from openpyxl import Workbook
from app.engine.template_parser import parse_media_tag
from app.engine.template_processor import TemplateProcessor
from app.renderer.fast_html_renderer import FastHtmlRenderer

def test_parse_media_tag_with_spaces_and_quotes():
    tag_str = "'https://verify.corp.com/id=' + invoice_id width=140 height=140"
    expr, attrs = parse_media_tag(tag_str)
    assert expr == "'https://verify.corp.com/id=' + invoice_id"
    assert attrs["width"] == 140
    assert attrs["height"] == 140

def test_media_custom_width_height():
    wb = Workbook()
    ws = wb.active

    ws["A1"] = "{{@qr 'ABC-123' width=180 height=180}}"
    ws["B1"] = "{{@barcode 'SKU-999' format=code128 width=260 height=55}}"

    processor = TemplateProcessor(wb)
    out_wb = processor.process({})
    out_ws = out_wb.active

    # Check openpyxl images dimensions
    images = out_ws._images
    assert len(images) == 2

    qr_img = images[0]
    bc_img = images[1]

    assert qr_img.width == 180
    assert qr_img.height == 180

    assert bc_img.width == 260
    assert bc_img.height == 55

    # Check HTML output contains exact width and height styling
    html = FastHtmlRenderer.render(out_ws)
    assert "width: 180px; height: 180px;" in html
    assert "width: 260px; height: 55px;" in html
