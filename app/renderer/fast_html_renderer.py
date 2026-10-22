import io
import re
import base64
import html
from typing import Any, Dict, List, Optional, Set, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, column_index_from_string
from app.core.config import settings
from app.engine.formula_evaluator import FormulaEvaluator
from app.renderer.color_resolver import ExcelColorResolver

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
    def _format_border_side(cls, side: Any, theme_colors: Optional[List[str]] = None) -> Optional[str]:
        """Maps an openpyxl Side border to CSS style string with exact color."""
        if not side or not getattr(side, "style", None) or side.style == "none":
            return None
        css_style = cls.BORDER_STYLE_MAP.get(side.style, "1px solid")
        color = ExcelColorResolver.resolve_color(getattr(side, "color", None), theme_colors, default="#cbd5e1")
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
    def _get_column_dimension(cls, ws: Worksheet, col_idx: int) -> Tuple[float, bool]:
        """
        Retrieves the effective width (in character units) and hidden status for a 1-based column index.
        Checks single column letter, column ranges (e.g. 'A:C'), and sheet defaults.
        """
        col_letter = get_column_letter(col_idx)
        dim = None

        if col_letter in ws.column_dimensions:
            dim = ws.column_dimensions[col_letter]
        else:
            # Check for range dimensions e.g. 'A:E'
            for key, candidate in ws.column_dimensions.items():
                if ":" in key:
                    parts = key.split(":")
                    if len(parts) == 2:
                        try:
                            start_i = column_index_from_string(parts[0])
                            end_i = column_index_from_string(parts[1])
                            if start_i <= col_idx <= end_i:
                                dim = candidate
                                break
                        except Exception:
                            pass

        hidden = bool(dim.hidden) if (dim and getattr(dim, "hidden", None) is not None) else False
        width = None
        if dim and getattr(dim, "width", None) is not None:
            try:
                width = float(dim.width)
            except (ValueError, TypeError):
                width = None

        if width is None:
            if hasattr(ws, "sheet_format") and getattr(ws.sheet_format, "defaultColWidth", None) is not None:
                try:
                    width = float(ws.sheet_format.defaultColWidth)
                except (ValueError, TypeError):
                    width = 8.43
            else:
                width = 8.43

        return width, hidden

    @classmethod
    def _excel_col_width_to_pixels(cls, char_width: float) -> int:
        """Converts Excel character column width to CSS screen/print pixels at 96 DPI."""
        if char_width <= 0:
            return 0
        if char_width < 1.0:
            return max(1, int(round(char_width * 12.0)))
        return int(round(char_width * 7.0 + 5.0))

    @classmethod
    def _get_row_dimension(cls, ws: Worksheet, row_num: int) -> Tuple[Optional[float], bool]:
        """
        Retrieves the height (in points) and hidden status for a 1-based row index.
        Returns (height_pt, is_hidden).
        """
        row_dim = ws.row_dimensions.get(row_num)
        if not row_dim:
            return None, False

        hidden = bool(row_dim.hidden) if getattr(row_dim, "hidden", None) is not None else False
        height = None
        if getattr(row_dim, "height", None) is not None:
            try:
                height = float(row_dim.height)
            except (ValueError, TypeError):
                height = None

        return height, hidden

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

        # Extract theme color palette from workbook
        theme_colors = ExcelColorResolver.extract_theme_colors(getattr(ws, "parent", None))

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

        # Determine default row height (in pt) from sheet format or default to 15.0pt
        default_row_height = 15.0
        if hasattr(ws, "sheet_format") and getattr(ws.sheet_format, "defaultRowHeight", None) is not None:
            try:
                default_row_height = float(ws.sheet_format.defaultRowHeight)
            except (ValueError, TypeError):
                default_row_height = 15.0

        # Build column dimensions and colgroup elements
        colgroup_tags: List[str] = []
        hidden_cols: Set[int] = set()
        total_visible_table_width_px = 0

        for c in range(1, max_col + 1):
            w_chars, col_hidden = cls._get_column_dimension(ws, c)
            col_px = cls._excel_col_width_to_pixels(w_chars)
            if col_hidden:
                hidden_cols.add(c)
                colgroup_tags.append('<col style="display: none; width: 0px;" />')
            else:
                total_visible_table_width_px += col_px
                colgroup_tags.append(f'<col style="width: {col_px}px;" />')

        if show_gridlines is None:
            display_gridlines = getattr(ws.print_options, "gridLines", False) is True
        else:
            display_gridlines = bool(show_gridlines)

        table_class = "report-table show-gridlines" if display_gridlines else "report-table"

        # Check page setup for landscape / portrait / margins
        page_setup_css = []
        orientation = getattr(getattr(ws, "page_setup", None), "orientation", None)
        if orientation == "landscape":
            page_setup_css.append("size: landscape;")
        elif orientation == "portrait":
            page_setup_css.append("size: portrait;")

        margins = getattr(ws, "page_margins", None)
        if margins:
            top_in = getattr(margins, "top", 0.75) or 0.75
            bottom_in = getattr(margins, "bottom", 0.75) or 0.75
            left_in = getattr(margins, "left", 0.7) or 0.7
            right_in = getattr(margins, "right", 0.7) or 0.7
            page_setup_css.append(f"margin: {top_in}in {right_in}in {bottom_in}in {left_in}in;")

        page_rule = f"@page {{ {' '.join(page_setup_css)} }}" if page_setup_css else ""

        html_out = [
            '<!DOCTYPE html>',
            '<html><head><meta charset="utf-8">',
            f'<title>{html.escape(title)}</title>',
            '<style>',
            f'body {{ font-family: {base_font_css}; margin: 20px; background-color: #f8fafc; color: #1e293b; }}',
            '.report-container { max-width: 1100px; margin: 0 auto; background: #ffffff; padding: 24px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -2px rgba(0,0,0,0.1); border-radius: 8px; overflow-x: auto; }',
            'table.report-table { border-collapse: collapse; width: 100%; table-layout: fixed; font-size: 13px; line-height: 1.2; }',
            f'table.report-table tr {{ height: {default_row_height}pt; }}',
            'table.report-table td { padding: 3px 6px; border: none; vertical-align: middle; box-sizing: border-box; }',
            'table.report-table.show-gridlines td { border: 1px solid #e2e8f0; }',
            f'{page_rule}',
            '@media print {',
            '  body { margin: 0; background: none; }',
            '  .report-container { box-shadow: none; padding: 0; max-width: 100%; border-radius: 0; }',
            '  .page-break { page-break-before: always !important; break-before: page !important; }',
            '  table.report-table { width: 100% !important; min-width: 0 !important; }',
            '}',
            '@media screen {',
            '  .page-break td { border-top: 4px dashed #6366f1 !important; }',
            '  .page-break-indicator { background: #eef2ff; color: #4f46e5; text-align: center; font-size: 11px; font-weight: 700; letter-spacing: 2px; padding: 8px 0; border: 1px dashed #6366f1; border-left: none; border-right: none; }',
            '}',
            '</style>',
            '</head><body>',
            '<div class="report-container">',
            f'<table class="{table_class}" style="min-width: {total_visible_table_width_px}px;">',
            '<colgroup>',
            *colgroup_tags,
            '</colgroup>'
        ]

        for r in range(1, max_row + 1):
            is_pb = (r in page_break_rows)
            r_height, r_hidden = cls._get_row_dimension(ws, r)

            tr_styles = []
            tr_classes = []

            if is_pb:
                tr_classes.append("page-break")
                tr_styles.append("page-break-before: always !important; break-before: page !important;")

            if r_hidden:
                tr_styles.append("display: none;")
            elif r_height is not None and r_height > 0:
                tr_styles.append(f"height: {r_height}pt;")

            class_attr = f' class="{" ".join(tr_classes)}"' if tr_classes else ""
            style_attr = f' style="{"; ".join(tr_styles)}"' if tr_styles else ""

            html_out.append(f'<tr{class_attr}{style_attr}>')

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
                if c in hidden_cols:
                    style_parts.append("display: none;")

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

                        f_color = ExcelColorResolver.resolve_font_color(cell.font, theme_colors)
                        if f_color:
                            style_parts.append(f"color: {f_color}")

                    if cell.alignment:
                        if cell.alignment.horizontal:
                            h_align = cell.alignment.horizontal
                            if h_align in ("left", "center", "right", "justify"):
                                style_parts.append(f"text-align: {h_align}")
                        if cell.alignment.vertical:
                            v_align = cell.alignment.vertical
                            if v_align == "center":
                                style_parts.append("vertical-align: middle")
                            elif v_align in ("top", "bottom", "baseline"):
                                style_parts.append(f"vertical-align: {v_align}")
                            elif v_align in ("justify", "distributed"):
                                style_parts.append("vertical-align: middle")
                        if getattr(cell.alignment, "wrap_text", None):
                            style_parts.append("white-space: pre-wrap; word-break: break-word")

                    bg_color = ExcelColorResolver.resolve_fill(cell.fill, theme_colors)
                    if bg_color:
                        style_parts.append(f"background-color: {bg_color}")

                    if cell.border:
                        top_b = cls._format_border_side(cell.border.top, theme_colors)
                        if top_b:
                            style_parts.append(f"border-top: {top_b}")
                        bot_b = cls._format_border_side(cell.border.bottom, theme_colors)
                        if bot_b:
                            style_parts.append(f"border-bottom: {bot_b}")
                        left_b = cls._format_border_side(cell.border.left, theme_colors)
                        if left_b:
                            style_parts.append(f"border-left: {left_b}")
                        right_b = cls._format_border_side(cell.border.right, theme_colors)
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
                if img_html:
                    content_html = img_html
                else:
                    content_html = html.escape(formatted_val).replace("\n", "<br/>")

                html_out.append(f'<td{span_attrs}{style_attr}>{content_html}</td>')

            html_out.append('</tr>')

        html_out.append('</table>')
        html_out.append('</div>')
        html_out.append('</body></html>')

        return "\n".join(html_out)
