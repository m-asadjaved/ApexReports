import io
import base64
import html
from typing import Any, Dict, Set, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from app.engine.formula_evaluator import FormulaEvaluator

class FastHtmlRenderer:
    """Renders an openpyxl Worksheet into responsive, stylized HTML with calculated formulas in < 25ms."""

    @classmethod
    def render(cls, ws: Worksheet, title: str = "Report Preview") -> str:
        """Converts openpyxl worksheet into HTML with inline CSS styling and calculated values."""
        max_row = ws.max_row or 1
        max_col = ws.max_column or 1

        # 1. Calculate all formulas in-memory (SUM, arithmetic, etc.)
        calc_cache = FormulaEvaluator.evaluate_worksheet(ws)

        # 2. Track merged cells: (row, col) -> (rowspan, colspan)
        merged_spans: Dict[Tuple[int, int], Tuple[int, int]] = {}
        skip_cells: Set[Tuple[int, int]] = set()

        for rng in ws.merged_cells.ranges:
            min_r, min_c, max_r, max_c = rng.min_row, rng.min_col, rng.max_row, rng.max_col
            rowspan = max_r - min_r + 1
            colspan = max_c - min_c + 1
            merged_spans[(min_r, min_c)] = (rowspan, colspan)

            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    if (r, c) != (min_r, min_c):
                        skip_cells.add((r, c))

        # 3. Map embedded images (barcodes, QR codes, logos) to cell coordinates
        cell_images: Dict[str, str] = {}
        for img in getattr(ws, "_images", []):
            try:
                # Resolve coordinate from anchor
                if isinstance(img.anchor, str):
                    coord = img.anchor.upper()
                elif hasattr(img, "anchor") and hasattr(img.anchor, "_from"):
                    from_col = getattr(img.anchor._from, "col", 0) + 1
                    from_row = getattr(img.anchor._from, "row", 0) + 1
                    coord = f"{get_column_letter(from_col)}{from_row}"
                else:
                    continue

                # Extract PNG bytes from PIL image
                raw_bytes = b""
                if hasattr(img.ref, "save"):
                    buf = io.BytesIO()
                    img.ref.save(buf, format="PNG")
                    raw_bytes = buf.getvalue()
                elif isinstance(img.ref, io.BytesIO):
                    raw_bytes = img.ref.getvalue()

                if raw_bytes:
                    b64 = base64.b64encode(raw_bytes).decode("ascii")
                    cell_images[coord] = f'<img src="data:image/png;base64,{b64}" style="max-height: 70px; max-width: 100%; object-fit: contain; display: block;" />'
            except Exception:
                pass

        html_out = [
            '<!DOCTYPE html>',
            '<html><head><meta charset="utf-8">',
            f'<title>{html.escape(title)}</title>',
            '<style>',
            'body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; margin: 20px; background-color: #f8fafc; color: #1e293b; }',
            '.report-container { max-width: 1100px; margin: 0 auto; background: #ffffff; padding: 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -2px rgba(0,0,0,0.1); border-radius: 8px; overflow-x: auto; }',
            'table.report-table { border-collapse: collapse; width: 100%; font-size: 13px; line-height: 1.4; }',
            'table.report-table td { padding: 6px 10px; border: 1px solid #e2e8f0; vertical-align: middle; }',
            '</style>',
            '</head><body>',
            '<div class="report-container">',
            '<table class="report-table">'
        ]

        for r in range(1, max_row + 1):
            html_out.append('<tr>')
            for c in range(1, max_col + 1):
                if (r, c) in skip_cells:
                    continue

                cell = ws.cell(row=r, column=c)
                coord = f"{get_column_letter(c)}{r}"

                # Attributes for merged cells
                span_attrs = ""
                if (r, c) in merged_spans:
                    rowspan, colspan = merged_spans[(r, c)]
                    if rowspan > 1:
                        span_attrs += f' rowspan="{rowspan}"'
                    if colspan > 1:
                        span_attrs += f' colspan="{colspan}"'

                # Extract styling
                style_parts = []
                if cell.has_style:
                    if cell.font:
                        if cell.font.bold:
                            style_parts.append("font-weight: bold")
                        if cell.font.italic:
                            style_parts.append("font-style: italic")
                        if cell.font.size:
                            style_parts.append(f"font-size: {int(cell.font.size)}pt")
                        if cell.font.color and getattr(cell.font.color, "rgb", None):
                            rgb = str(cell.font.color.rgb)
                            if len(rgb) == 8:
                                rgb = rgb[2:]
                            style_parts.append(f"color: #{rgb}")

                    if cell.alignment:
                        if cell.alignment.horizontal:
                            style_parts.append(f"text-align: {cell.alignment.horizontal}")
                        if cell.alignment.vertical:
                            style_parts.append(f"vertical-align: {cell.alignment.vertical}")

                    if cell.fill and getattr(cell.fill, "start_color", None):
                        fill_rgb = getattr(cell.fill.start_color, "rgb", None)
                        if fill_rgb and str(fill_rgb) != "00000000":
                            s_rgb = str(fill_rgb)
                            if len(s_rgb) == 8:
                                s_rgb = s_rgb[2:]
                            style_parts.append(f"background-color: #{s_rgb}")

                style_attr = f' style="{"; ".join(style_parts)}"' if style_parts else ""

                # Determine cell display value
                # If it was a formula, use calculated value from calc_cache!
                if (r, c) in calc_cache:
                    raw_val = calc_cache[(r, c)]
                else:
                    raw_val = cell.value

                # Format according to cell's number format
                num_format = getattr(cell, "number_format", None)
                formatted_val = FormulaEvaluator.format_cell_value(raw_val, num_format)

                # Check if there is an embedded image for this cell
                img_html = cell_images.get(coord, "")
                content_html = img_html if img_html else html.escape(formatted_val)

                html_out.append(f'<td{span_attrs}{style_attr}>{content_html}</td>')

            html_out.append('</tr>')

        html_out.append('</table>')
        html_out.append('</div>')
        html_out.append('</body></html>')

        return "\n".join(html_out)
