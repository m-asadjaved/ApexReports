import io
import re
import base64
import html
from typing import Any, Dict, List, Optional, Set, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import range_boundaries
from app.core.config import settings
from app.engine.formula_evaluator import FormulaEvaluator
from app.renderer.color_resolver import ExcelColorResolver
from app.engine.conditional_formatting_evaluator import ConditionalFormattingEvaluator

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
    def _apply_style_override(cls, style_parts: List[str], property_prefix: str, new_style: str) -> None:
        """Replaces an existing CSS property in style_parts or appends new_style."""
        for i, part in enumerate(style_parts):
            if part.strip().startswith(property_prefix):
                style_parts[i] = new_style
                return
        style_parts.append(new_style)

    @classmethod
    def _get_printable_boundaries(cls, ws: Worksheet) -> Tuple[int, int, int, int]:
        """
        Extracts (min_col, min_row, max_col, max_row) representing the printable area.
        Respects ws.print_area if defined; otherwise detects effective content bounds.
        """
        pa = getattr(ws, "print_area", None)
        if pa:
            pa_str = pa[0] if isinstance(pa, (list, tuple)) else str(pa)
            if pa_str.strip():
                clean = pa_str.split("!")[-1].replace("$", "").strip()
                try:
                    min_c, min_r, max_c, max_r = range_boundaries(clean)
                    if min_c and max_c:
                        eff_max_r = max(max_r or ws.max_row or 1, ws.max_row or 1)
                        eff_min_r = min_r or 1
                        return (min_c or 1, eff_min_r, max_c, eff_max_r)
                except Exception:
                    pass

        # Fallback: find highest column with actual values, styles, or explicit column dimension
        max_col_detected = 1
        max_row_detected = ws.max_row or 1

        for r in range(1, max_row_detected + 1):
            for c in range(1, (ws.max_column or 1) + 1):
                cell = ws.cell(row=r, column=c)
                if cell.value is not None or (cell.has_style and (getattr(cell.fill, "fill_type", None) or (cell.border and getattr(cell.border.top, "style", None)))):
                    max_col_detected = max(max_col_detected, c)

        for rng in getattr(ws.merged_cells, "ranges", []):
            if getattr(rng, "max_col", 0) > max_col_detected:
                max_col_detected = rng.max_col

        # Also check columns with explicit width that contain content
        for k, col_dim in ws.column_dimensions.items():
            if col_dim and getattr(col_dim, "width", None) is not None:
                if ":" not in k:
                    try:
                        ci = column_index_from_string(k)
                        if ci <= (ws.max_column or 1):
                            max_col_detected = max(max_col_detected, ci)
                    except Exception:
                        pass

        return (1, 1, max(1, min(max_col_detected, ws.max_column or 1)), max_row_detected)

    @classmethod
    def render(
        cls, 
        ws: Worksheet, 
        title: str = "Report Preview", 
        show_gridlines: Optional[bool] = None,
        font_family: Optional[str] = None,
        fit_to_page: bool = True
    ) -> str:
        """Converts openpyxl worksheet into HTML with inline CSS styling, page breaks, and calculated values."""
        min_col, min_row, max_col, max_row = cls._get_printable_boundaries(ws)

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

        # 1b. Evaluate conditional formatting rules across worksheet cells
        cf_overlays = ConditionalFormattingEvaluator.evaluate_worksheet(ws, calc_cache, theme_colors)

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
        col_dims_raw: List[Tuple[int, int, bool]] = []
        hidden_cols: Set[int] = set()
        total_visible_table_width_px = 0

        for c in range(min_col, max_col + 1):
            w_chars, col_hidden = cls._get_column_dimension(ws, c)
            col_px = cls._excel_col_width_to_pixels(w_chars)
            if col_hidden:
                hidden_cols.add(c)
                col_dims_raw.append((c, col_px, True))
            else:
                total_visible_table_width_px += col_px
                col_dims_raw.append((c, col_px, False))

        # Calculate proportional percentages so total equals exactly 100.00%
        colgroup_tags: List[str] = []
        if fit_to_page and total_visible_table_width_px > 0:
            visible_items = [item for item in col_dims_raw if not item[2]]
            pct_accum = 0.0
            col_pct_map: Dict[int, float] = {}
            for idx, (c, col_px, _) in enumerate(visible_items):
                if idx == len(visible_items) - 1:
                    pct = round(100.0 - pct_accum, 2)
                else:
                    pct = round((col_px / total_visible_table_width_px) * 100.0, 2)
                    pct_accum += pct
                col_pct_map[c] = max(0.1, pct)

            for c, col_px, col_hidden in col_dims_raw:
                if col_hidden:
                    colgroup_tags.append('<col style="display: none; width: 0px;" />')
                else:
                    pct = col_pct_map.get(c, 10.0)
                    colgroup_tags.append(f'<col style="width: {pct:.2f}%;" />')
        else:
            for c, col_px, col_hidden in col_dims_raw:
                if col_hidden:
                    colgroup_tags.append('<col style="display: none; width: 0px;" />')
                else:
                    colgroup_tags.append(f'<col style="width: {col_px}px;" />')

        if show_gridlines is None:
            display_gridlines = getattr(ws.print_options, "gridLines", False) is True
        else:
            display_gridlines = bool(show_gridlines)

        table_class = "report-table show-gridlines" if display_gridlines else "report-table"

        # Determine base table font size and cell padding for optimal fit across columns
        if fit_to_page:
            if max_col >= 12 or total_visible_table_width_px > 1100:
                base_font_size = "9.5px"
                cell_padding = "2px 3px"
            elif max_col >= 8 or total_visible_table_width_px > 800:
                base_font_size = "11px"
                cell_padding = "2px 4px"
            elif max_col >= 6:
                base_font_size = "12px"
                cell_padding = "3px 5px"
            else:
                base_font_size = "13px"
                cell_padding = "4px 8px"
        else:
            base_font_size = "13px"
            cell_padding = "3px 6px"

        # Paper size detection (9=A4, 1=Letter, 8=A3, 5=Legal)
        paper_size_code = getattr(getattr(ws, "page_setup", None), "paperSize", None)
        paper_map = {
            1: ("letter", 8.5, 11.0),
            5: ("legal", 8.5, 14.0),
            8: ("a3", 11.69, 16.54),
            9: ("a4", 8.27, 11.69),
        }
        paper_name, paper_w_in, paper_h_in = paper_map.get(paper_size_code, ("a4", 8.27, 11.69))

        orientation = getattr(getattr(ws, "page_setup", None), "orientation", None)
        num_cols = max_col - min_col + 1
        is_wide = (total_visible_table_width_px > 780 or num_cols >= 8)
        if orientation == "landscape" or (not orientation and is_wide):
            page_w_in, page_h_in = max(paper_w_in, paper_h_in), min(paper_w_in, paper_h_in)
            orient_str = "landscape"
        else:
            page_w_in, page_h_in = min(paper_w_in, paper_h_in), max(paper_w_in, paper_h_in)
            orient_str = "portrait"

        margins = getattr(ws, "page_margins", None)
        top_in = round(min(float(getattr(margins, "top", 0.5) or 0.5), 0.75), 3) if margins else 0.5
        bottom_in = round(min(float(getattr(margins, "bottom", 0.5) or 0.5), 0.75), 3) if margins else 0.5

        # Horizontally center all reports on the page and match margins
        table_w_in = total_visible_table_width_px / 96.0
        avail_h_margin = max(0.0, page_w_in - table_w_in)
        h_margin_in = round(max(0.15, avail_h_margin / 2.0), 3)

        page_rule = f"@page {{ size: {paper_name} {orient_str}; margin: {top_in}in {h_margin_in}in {bottom_in}in {h_margin_in}in; }}"

        table_width_style = "width: 100%;" if fit_to_page else f"min-width: {total_visible_table_width_px}px;"

        html_out = [
            '<!DOCTYPE html>',
            '<html><head><meta charset="utf-8">',
            f'<title>{html.escape(title)}</title>',
            '<style>',
            f'body {{ font-family: {base_font_css}; margin: 20px; background-color: #f8fafc; color: #1e293b; }}',
            '.report-container { width: 100%; max-width: 100%; margin: 0 auto; background: #ffffff; padding: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -2px rgba(0,0,0,0.1); border-radius: 8px; box-sizing: border-box; overflow-x: auto; display: flex; flex-direction: column; align-items: center; }',
            f'table.report-table {{ border-collapse: collapse; width: 100%; max-width: 100%; table-layout: fixed; font-size: {base_font_size}; line-height: 1.25; margin-left: auto !important; margin-right: auto !important; }}',
            f'table.report-table tr {{ height: {default_row_height}pt; }}',
            f'table.report-table td {{ padding: {cell_padding}; border: none; vertical-align: middle; box-sizing: border-box; word-wrap: break-word; overflow-wrap: break-word; }}',
            'table.report-table.show-gridlines td {{ border: 1px solid #e2e8f0; }}',
            f'{page_rule}',
            '@media print {',
            '  body { margin: 0; background: none; -webkit-print-color-adjust: exact; print-color-adjust: exact; }',
            '  .report-container { box-shadow: none; padding: 0; width: 100% !important; max-width: 100% !important; border-radius: 0; margin: 0 auto !important; display: block !important; }',
            '  .page-break { page-break-before: always !important; break-before: page !important; }',
            '  table.report-table { width: 100% !important; max-width: 100% !important; table-layout: fixed !important; margin-left: auto !important; margin-right: auto !important; }',
            '}',
            '@media screen {',
            '  .page-break td { border-top: 4px dashed #6366f1 !important; }',
            '  .page-break-indicator { background: #eef2ff; color: #4f46e5; text-align: center; font-size: 11px; font-weight: 700; letter-spacing: 2px; padding: 8px 0; border: 1px dashed #6366f1; border-left: none; border-right: none; }',
            '}',
            '</style>',
            '</head><body>',
            '<div class="report-container">',
            f'<table class="{table_class}" style="{table_width_style}">',
            '<colgroup>',
            *colgroup_tags,
            '</colgroup>'
        ]

        for r in range(min_row, max_row + 1):
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

            for c in range(min_col, max_col + 1):
                if (r, c) in skip_cells:
                    continue

                cell = ws.cell(row=r, column=c)
                coord = f"{get_column_letter(c)}{r}"

                span_attrs = ""
                if (r, c) in merged_spans:
                    rowspan, colspan = merged_spans[(r, c)]
                    colspan = min(colspan, max_col - c + 1)
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

                # For merged cells, ensure outer edge borders are captured from boundary cells
                if (r, c) in merged_spans:
                    r_span, c_span = merged_spans[(r, c)]
                    end_r = r + r_span - 1
                    end_c = c + c_span - 1
                    if not any(sp.startswith("border-right:") for sp in style_parts):
                        for rk in range(r, end_r + 1):
                            sc = ws.cell(row=rk, column=end_c)
                            if sc.border and getattr(sc.border.right, "style", None):
                                rb = cls._format_border_side(sc.border.right, theme_colors)
                                if rb:
                                    style_parts.append(f"border-right: {rb}")
                                break
                    if not any(sp.startswith("border-bottom:") for sp in style_parts):
                        for ck in range(c, end_c + 1):
                            sc = ws.cell(row=end_r, column=ck)
                            if sc.border and getattr(sc.border.bottom, "style", None):
                                bb = cls._format_border_side(sc.border.bottom, theme_colors)
                                if bb:
                                    style_parts.append(f"border-bottom: {bb}")
                                break
                    if not any(sp.startswith("border-left:") for sp in style_parts):
                        for rk in range(r, end_r + 1):
                            sc = ws.cell(row=rk, column=c)
                            if sc.border and getattr(sc.border.left, "style", None):
                                lb = cls._format_border_side(sc.border.left, theme_colors)
                                if lb:
                                    style_parts.append(f"border-left: {lb}")
                                break
                    if not any(sp.startswith("border-top:") for sp in style_parts):
                        for ck in range(c, end_c + 1):
                            sc = ws.cell(row=r, column=ck)
                            if sc.border and getattr(sc.border.top, "style", None):
                                tb = cls._format_border_side(sc.border.top, theme_colors)
                                if tb:
                                    style_parts.append(f"border-top: {tb}")
                                break

                cf_overlay = cf_overlays.get((r, c))
                if cf_overlay:
                    if cf_overlay.bg_color:
                        cls._apply_style_override(style_parts, "background-color:", f"background-color: {cf_overlay.bg_color}")
                    if cf_overlay.font_color:
                        cls._apply_style_override(style_parts, "color:", f"color: {cf_overlay.font_color}")
                    if cf_overlay.bold:
                        cls._apply_style_override(style_parts, "font-weight:", "font-weight: bold")
                    if cf_overlay.italic:
                        cls._apply_style_override(style_parts, "font-style:", "font-style: italic")
                    if cf_overlay.border_top:
                        cls._apply_style_override(style_parts, "border-top:", f"border-top: {cf_overlay.border_top}")
                    if cf_overlay.border_bottom:
                        cls._apply_style_override(style_parts, "border-bottom:", f"border-bottom: {cf_overlay.border_bottom}")
                    if cf_overlay.border_left:
                        cls._apply_style_override(style_parts, "border-left:", f"border-left: {cf_overlay.border_left}")
                    if cf_overlay.border_right:
                        cls._apply_style_override(style_parts, "border-right:", f"border-right: {cf_overlay.border_right}")

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
