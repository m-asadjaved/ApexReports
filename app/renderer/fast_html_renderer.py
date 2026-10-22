import io
import re
import base64
import html
from typing import Any, Dict, Optional, Set, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from app.core.config import settings
from app.engine.formula_evaluator import FormulaEvaluator

class FastHtmlRenderer:
    """Renders an openpyxl Worksheet into responsive, stylized HTML with calculated formulas in < 25ms."""

    BORDER_STYLE_MAP = {
        "thin": "1px solid",
        "medium": "2px solid",
        "thick": "3px solid",
        "double": "3px double",
        "dashed": "1px dashed",
        "dotted": "1px dotted",
        "hair": "1px dotted",
        "dashDot": "1px dashed",
        "mediumDashed": "2px dashed",
    }

    FONT_FALLBACK_MAP = {
        "calibri": "'Calibri', 'Carlito', 'Segoe UI', Arial, sans-serif",
        "aptos": "'Aptos', 'Calibri', 'Segoe UI', Arial, sans-serif",
        "arial": "'Arial', 'Liberation Sans', 'Helvetica Neue', Helvetica, sans-serif",
        "times new roman": "'Times New Roman', 'Liberation Serif', Times, serif",
        "segoe ui": "'Segoe UI', -apple-system, Roboto, Arial, sans-serif",
        "tahoma": "'Tahoma', 'Geneva', Verdana, sans-serif",
        "verdana": "'Verdana', 'Geneva', sans-serif",
        "consolas": "'Consolas', 'Courier New', monospace",
        "courier new": "'Courier New', Courier, monospace",
        "georgia": "'Georgia', 'Bitstream Charter', serif",
        "trebuchet ms": "'Trebuchet MS', 'Lucida Sans', sans-serif",
        "cambria": "'Cambria', 'Caladea', serif",
        "comic sans ms": "'Comic Sans MS', cursive",
    }

    @classmethod
    def _format_border_side(cls, side: Any) -> Optional[str]:
        """Maps an openpyxl Side border to CSS style string."""
        if not side or not getattr(side, "style", None) or side.style == "none":
            return None
        css_style = cls.BORDER_STYLE_MAP.get(side.style, "1px solid")
        color = "#cbd5e1"
        if getattr(side, "color", None) and getattr(side.color, "rgb", None):
            raw_rgb = getattr(side.color, "rgb")
            if isinstance(raw_rgb, str) and re.match(r"^[0-9a-fA-F]{6,8}$", raw_rgb):
                color = f"#{raw_rgb[2:] if len(raw_rgb) == 8 else raw_rgb}"
        return f"{css_style} {color}"

    @classmethod
    def _get_font_family_css(cls, font_name: Optional[str]) -> str:
        """Returns CSS font-family string with sensible web and cross-platform fallbacks."""
        if not font_name:
            return "'Calibri', 'Segoe UI', Arial, sans-serif"
        clean = font_name.strip().strip("'\"")
        fn_lower = clean.lower()
        if fn_lower in cls.FONT_FALLBACK_MAP:
            return cls.FONT_FALLBACK_MAP[fn_lower]
        return f"'{clean}', 'Segoe UI', Arial, sans-serif"

    @classmethod
    def render(
        cls, 
        ws: Worksheet, 
        title: str = "Report Preview", 
        show_gridlines: Optional[bool] = None,
        font_family: Optional[str] = None
    ) -> str:
        """Converts openpyxl worksheet into HTML with inline CSS styling, page breaks, and calculated values."""
        max_row = ws.max_row or 1
        max_col = ws.max_column or 1

        # Determine base document font
        if font_family:
            base_font_css = cls._get_font_family_css(font_family)
            primary_font_name = font_family
        elif settings.DEFAULT_FONT_FAMILY:
            base_font_css = cls._get_font_family_css(settings.DEFAULT_FONT_FAMILY)
            primary_font_name = settings.DEFAULT_FONT_FAMILY
        else:
            # Auto-detect primary font used in the Excel template
            detected_font = "Calibri"
            for row in ws.iter_rows():
                for c in row:
                    if c.font and c.font.name:
                        detected_font = c.font.name
                        break
                if detected_font != "Calibri":
                    break
            base_font_css = cls._get_font_family_css(detected_font)
            primary_font_name = detected_font

        # 1. Calculate all formulas in-memory (SUM, arithmetic, etc.)
        calc_cache = FormulaEvaluator.evaluate_worksheet(ws)

        # 2. Extract page breaks (1-based row index that starts a new page)
        page_break_rows: Set[int] = set()
        if hasattr(ws, "row_breaks") and hasattr(ws.row_breaks, "brk"):
            for b in ws.row_breaks.brk:
                if getattr(b, "id", None) is not None:
                    page_break_rows.add(b.id + 1)

        # 3. Track merged cells: (row, col) -> (rowspan, colspan)
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

        # 4. Map embedded images (barcodes, QR codes, logos) to cell coordinates
        cell_images: Dict[str, str] = {}
        for img in getattr(ws, "_images", []):
            try:
                if isinstance(img.anchor, str):
                    coord = img.anchor.upper()
                elif hasattr(img, "anchor") and hasattr(img.anchor, "_from"):
                    from_col = getattr(img.anchor._from, "col", 0) + 1
                    from_row = getattr(img.anchor._from, "row", 0) + 1
                    coord = f"{get_column_letter(from_col)}{from_row}"
                else:
                    continue

                raw_bytes = b""
                if hasattr(img.ref, "save"):
                    buf = io.BytesIO()
                    img.ref.save(buf, format="PNG")
                    raw_bytes = buf.getvalue()
                elif isinstance(img.ref, io.BytesIO):
                    raw_bytes = img.ref.getvalue()

                if raw_bytes:
                    b64 = base64.b64encode(raw_bytes).decode("ascii")
                    img_w = int(img.width) if getattr(img, "width", None) else None
                    img_h = int(img.height) if getattr(img, "height", None) else None

                    css_dims = []
                    if img_w:
                        css_dims.append(f"width: {img_w}px;")
                    if img_h:
                        css_dims.append(f"height: {img_h}px;")
                    if not img_w and not img_h:
                        css_dims.append("max-height: 80px; max-width: 100%;")
                    css_dims.append("object-fit: contain; display: block;")

                    style_str = " ".join(css_dims)
                    cell_images[coord] = f'<img src="data:image/png;base64,{b64}" style="{style_str}" />'
            except Exception:
                pass

        if show_gridlines is None:
            display_gridlines = getattr(ws.print_options, "gridLines", False) is True
        else:
            display_gridlines = bool(show_gridlines)

        table_class = "report-table show-gridlines" if display_gridlines else "report-table"

        html_out = [
            '<!DOCTYPE html>',
            '<html><head><meta charset="utf-8">',
            f'<title>{html.escape(title)}</title>',
            '<style>',
            f'body {{ font-family: {base_font_css}; margin: 20px; background-color: #f8fafc; color: #1e293b; }}',
            '.report-container { max-width: 1100px; margin: 0 auto; background: #ffffff; padding: 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -2px rgba(0,0,0,0.1); border-radius: 8px; overflow-x: auto; }',
            'table.report-table { border-collapse: collapse; width: 100%; font-size: 13px; line-height: 1.4; }',
            'table.report-table td { padding: 6px 10px; border: none; vertical-align: middle; }',
            'table.report-table.show-gridlines td { border: 1px solid #e2e8f0; }',
            '@media print {',
            '  body { margin: 0; background: none; }',
            '  .report-container { box-shadow: none; padding: 0; max-width: 100%; border-radius: 0; }',
            '  .page-break { page-break-before: always !important; break-before: page !important; }',
            '}',
            '@media screen {',
            '  .page-break td { border-top: 4px dashed #6366f1 !important; }',
            '  .page-break-indicator { background: #eef2ff; color: #4f46e5; text-align: center; font-size: 11px; font-weight: 700; letter-spacing: 2px; padding: 8px 0; border: 1px dashed #6366f1; border-left: none; border-right: none; }',
            '}',
            '</style>',
            '</head><body>',
            '<div class="report-container">',
            f'<table class="{table_class}">'
        ]

        for r in range(1, max_row + 1):
            is_pb = (r in page_break_rows)

            # In screen view, show an explicit visual page-break divider row
            if is_pb:
                html_out.append(f'<tr class="page-break" style="page-break-before: always !important; break-before: page !important;">')
            else:
                html_out.append('<tr>')

            for c in range(1, max_col + 1):
                if (r, c) in skip_cells:
                    continue

                cell = ws.cell(row=r, column=c)
                coord = f"{get_column_letter(c)}{r}"

                span_attrs = ""
                if (r, c) in merged_spans:
                    rowspan, colspan = merged_spans[(r, c)]
                    if rowspan > 1:
                        span_attrs += f' rowspan="{rowspan}"'
                    if colspan > 1:
                        span_attrs += f' colspan="{colspan}"'

                style_parts = []
                if cell.has_style:
                    if cell.font:
                        if not settings.FORCE_GLOBAL_FONT and cell.font.name:
                            cell_font_clean = cell.font.name.strip().strip("'\"")
                            if cell_font_clean.lower() != primary_font_name.lower():
                                style_parts.append(f"font-family: {cls._get_font_family_css(cell_font_clean)}")

                        if cell.font.bold:
                            style_parts.append("font-weight: bold")
                        if cell.font.italic:
                            style_parts.append("font-style: italic")
                        if cell.font.size:
                            style_parts.append(f"font-size: {int(cell.font.size)}pt")
                        if cell.font.color and getattr(cell.font.color, "rgb", None):
                            raw_rgb = getattr(cell.font.color, "rgb", None)
                            if isinstance(raw_rgb, str) and re.match(r"^[0-9a-fA-F]{6,8}$", raw_rgb):
                                rgb = raw_rgb[2:] if len(raw_rgb) == 8 else raw_rgb
                                style_parts.append(f"color: #{rgb}")

                    if cell.alignment:
                        if cell.alignment.horizontal:
                            style_parts.append(f"text-align: {cell.alignment.horizontal}")
                        if cell.alignment.vertical:
                            style_parts.append(f"vertical-align: {cell.alignment.vertical}")

                    if cell.fill and getattr(cell.fill, "start_color", None):
                        raw_fill = getattr(cell.fill.start_color, "rgb", None)
                        if isinstance(raw_fill, str) and re.match(r"^[0-9a-fA-F]{6,8}$", raw_fill) and raw_fill != "00000000":
                            s_rgb = raw_fill[2:] if len(raw_fill) == 8 else raw_fill
                            style_parts.append(f"background-color: #{s_rgb}")

                    if cell.border:
                        top_b = cls._format_border_side(cell.border.top)
                        if top_b:
                            style_parts.append(f"border-top: {top_b}")
                        bot_b = cls._format_border_side(cell.border.bottom)
                        if bot_b:
                            style_parts.append(f"border-bottom: {bot_b}")
                        left_b = cls._format_border_side(cell.border.left)
                        if left_b:
                            style_parts.append(f"border-left: {left_b}")
                        right_b = cls._format_border_side(cell.border.right)
                        if right_b:
                            style_parts.append(f"border-right: {right_b}")

                style_attr = f' style="{"; ".join(style_parts)}"' if style_parts else ""

                if (r, c) in calc_cache:
                    raw_val = calc_cache[(r, c)]
                else:
                    raw_val = cell.value

                num_format = getattr(cell, "number_format", None)
                formatted_val = FormulaEvaluator.format_cell_value(raw_val, num_format)

                img_html = cell_images.get(coord, "")
                content_html = img_html if img_html else html.escape(formatted_val)

                html_out.append(f'<td{span_attrs}{style_attr}>{content_html}</td>')

            html_out.append('</tr>')

        html_out.append('</table>')
        html_out.append('</div>')
        html_out.append('</body></html>')

        return "\n".join(html_out)
