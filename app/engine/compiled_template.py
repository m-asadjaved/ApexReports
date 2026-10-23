import io
import re
import math
import base64
import html
from datetime import datetime, date
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Set, Tuple, Union
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.cell import range_boundaries

from app.core.config import settings
from app.engine.template_parser import ExpressionEvaluator, parse_tag_attributes, parse_media_tag
from app.engine.cell_utils import render_cell_value
from app.engine.media_handler import MediaHandler
from app.renderer.color_resolver import ExcelColorResolver
from app.renderer.fast_html_renderer import FastHtmlRenderer

evaluator = ExpressionEvaluator()

@dataclass
class CompiledCell:
    col_idx: int
    colspan: int = 1
    rowspan: int = 1
    is_skipped: bool = False
    style_str: str = ""
    raw_val: Any = None
    val_type: str = "static"  # 'static', 'var', 'formula', 'media', 'empty'
    num_format: Optional[str] = None
    # For formulas
    formula_str: Optional[str] = None
    formula_col_refs: List[int] = field(default_factory=list)
    is_sum_formula: bool = False
    sum_target_col: Optional[int] = None

@dataclass
class CompiledRow:
    row_num: int
    height_pt: float = 15.0
    is_hidden: bool = False
    cells: List[CompiledCell] = field(default_factory=list)
    raw_directive: Optional[Dict[str, Any]] = None

@dataclass
class CompiledElement:
    pass

@dataclass
class StaticRowElement(CompiledElement):
    row: CompiledRow

@dataclass
class PageBreakElement(CompiledElement):
    pass

@dataclass
class EachBlockElement(CompiledElement):
    collection_key: str
    attrs: Dict[str, Any] = field(default_factory=dict)
    children: List[CompiledElement] = field(default_factory=list)

@dataclass
class IfBlockElement(CompiledElement):
    condition: str
    children: List[CompiledElement] = field(default_factory=list)

@dataclass
class CompiledTemplateAST:
    template_code: str
    mtime: float
    header_html: str
    footer_html: str
    elements: List[CompiledElement]
    column_count: int
    default_font: str = "Calibri"


class TemplateCompiler:
    """Compiles an openpyxl Workbook into an intermediate CompiledTemplateAST."""

    @classmethod
    def compile(cls, wb: Workbook, template_code: str, mtime: float = 0.0) -> CompiledTemplateAST:
        ws: Worksheet = wb.active
        theme_colors = ExcelColorResolver.extract_theme_colors(wb)
        min_col, min_row, max_col, max_row = FastHtmlRenderer._get_printable_boundaries(ws)

        # 1. Page & Font setups
        detected_font = "Calibri"
        for row in ws.iter_rows():
            for c in row:
                if c.font and c.font.name:
                    detected_font = c.font.name
                    break
            if detected_font != "Calibri":
                break

        base_font_css = FastHtmlRenderer._get_font_family_css(detected_font)

        # 2. Column dimensions & colgroup
        col_dims_raw = []
        hidden_cols = set()
        total_px = 0
        for c in range(min_col, max_col + 1):
            w_chars, col_hidden = FastHtmlRenderer._get_column_dimension(ws, c)
            col_px = FastHtmlRenderer._excel_col_width_to_pixels(w_chars)
            if col_hidden:
                hidden_cols.add(c)
                col_dims_raw.append((c, col_px, True))
            else:
                total_px += col_px
                col_dims_raw.append((c, col_px, False))

        colgroup_tags = []
        if total_px > 0:
            visible_items = [item for item in col_dims_raw if not item[2]]
            pct_accum = 0.0
            col_pct_map = {}
            for idx, (c, col_px, _) in enumerate(visible_items):
                if idx == len(visible_items) - 1:
                    pct = round(100.0 - pct_accum, 2)
                else:
                    pct = round((col_px / total_px) * 100.0, 2)
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

        # 3. Paper size & orientation detection
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
        is_wide = (total_px > 780 or num_cols >= 8)
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
        table_w_in = total_px / 96.0
        avail_h_margin = max(0.0, page_w_in - table_w_in)
        h_margin_in = round(max(0.15, avail_h_margin / 2.0), 3)
        page_size_str = f"{paper_name} {orient_str}"

        # Print options
        h_center = getattr(ws.print_options, "horizontalCentered", False) is True
        display_gridlines = getattr(ws.print_options, "gridLines", False) is True
        grid_class = " show-gridlines" if display_gridlines else ""

        # Merged cells map
        merged_spans: Dict[Tuple[int, int], Tuple[int, int]] = {}
        skip_cells: Set[Tuple[int, int]] = set()
        for rng in ws.merged_cells.ranges:
            min_r, min_c, max_r, max_c = rng.min_row, rng.min_col, rng.max_row, rng.max_col
            merged_spans[(min_r, min_c)] = (max_r - min_r + 1, max_c - min_c + 1)
            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    if (r, c) != (min_r, min_c):
                        skip_cells.add((r, c))

        header_html = (
            f'<!DOCTYPE html>\n<html><head><meta charset="utf-8">\n'
            f'<title>{template_code}</title>\n<style>\n'
            f"body {{ font-family: {base_font_css}; margin: 20px; background-color: #f8fafc; color: #1e293b; }}\n"
            f".report-container {{ width: 100%; max-width: 100%; margin: 0 auto; background: #ffffff; padding: 20px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); border-radius: 8px; box-sizing: border-box; overflow-x: auto; display: flex; flex-direction: column; align-items: {'center' if h_center else 'flex-start'}; }}\n"
            f"table.report-table {{ border-collapse: collapse; width: 100%; max-width: 100%; table-layout: fixed; font-size: 13px; line-height: 1.25; margin-left: auto !important; margin-right: auto !important; }}\n"
            f"table.report-table tr {{ height: 15.0pt; }}\n"
            f"table.report-table td {{ padding: 4px 8px; border: none; vertical-align: middle; box-sizing: border-box; word-wrap: break-word; overflow-wrap: break-word; }}\n"
            f"table.report-table.show-gridlines td {{ border: 1px solid #e2e8f0; }}\n"
            f"@page {{ size: {page_size_str}; margin: {top_in}in {h_margin_in}in {bottom_in}in {h_margin_in}in; }}\n"
            f"@media print {{\n"
            f"  body {{ margin: 0; background: none; -webkit-print-color-adjust: exact; print-color-adjust: exact; }}\n"
            f"  .report-container {{ box-shadow: none; padding: 0; width: 100% !important; max-width: 100% !important; border-radius: 0; margin: 0 auto !important; display: block !important; }}\n"
            f"  .page-break {{ page-break-before: always !important; break-before: page !important; }}\n"
            f"  table.report-table {{ width: 100% !important; max-width: 100% !important; table-layout: fixed !important; margin-left: auto !important; margin-right: auto !important; }}\n"
            f"}}\n"
            f"@media screen {{\n"
            f"  .page-break td {{ border-top: 2px solid #e2e8f0 !important; }}\n"
            f"}}\n</style>\n</head><body>\n"
            f'<div class="report-container">\n'
            f'<table class="report-table{grid_class}" style="width: 100%;">\n'
            f'<colgroup>\n' + "\n".join(colgroup_tags) + "\n</colgroup>\n"
        )
        footer_html = "</table>\n</div>\n</body></html>"

        # 3. Compile rows
        all_rows: List[CompiledRow] = []
        for r in range(min_row, max_row + 1):
            row_dim = ws.row_dimensions.get(r)
            row_h = float(row_dim.height) if (row_dim and row_dim.height is not None) else 15.0
            row_hidden = bool(row_dim.hidden) if (row_dim and row_dim.hidden is not None) else False

            c_cells: List[CompiledCell] = []
            for c in range(min_col, max_col + 1):
                if (r, c) in skip_cells:
                    c_cells.append(CompiledCell(col_idx=c, is_skipped=True))
                    continue

                cell = ws.cell(row=r, column=c)
                colspan, rowspan = merged_spans.get((r, c), (1, 1))

                # Style building
                css_rules = []
                if cell.has_style:
                    if cell.font:
                        if cell.font.bold:
                            css_rules.append("font-weight: bold;")
                        if cell.font.italic:
                            css_rules.append("font-style: italic;")
                        if cell.font.underline:
                            css_rules.append("text-decoration: underline;")
                        if cell.font.size:
                            css_rules.append(f"font-size: {cell.font.size}pt;")
                        if cell.font.color:
                            f_color = ExcelColorResolver.resolve_color(cell.font.color, theme_colors)
                            if f_color:
                                css_rules.append(f"color: {f_color};")

                    if cell.alignment:
                        h_align = cell.alignment.horizontal
                        if h_align in ("left", "center", "right", "justify"):
                            css_rules.append(f"text-align: {h_align};")
                        v_align = cell.alignment.vertical
                        if v_align in ("top", "center", "bottom"):
                            css_rules.append(f"vertical-align: {'middle' if v_align == 'center' else v_align};")

                    if cell.fill and getattr(cell.fill, "fill_type", None) == "solid":
                        bg_color = ExcelColorResolver.resolve_color(getattr(cell.fill, "start_color", None), theme_colors)
                        if bg_color and bg_color.lower() != "#ffffff":
                            css_rules.append(f"background-color: {bg_color};")

                    if cell.border:
                        for side_name in ("top", "bottom", "left", "right"):
                            side = getattr(cell.border, side_name, None)
                            b_str = FastHtmlRenderer._format_border_side(side, theme_colors)
                            if b_str:
                                css_rules.append(f"border-{side_name}: {b_str};")

                style_str = " ".join(css_rules)
                raw_val = cell.value
                val_type = "empty"
                formula_str = None
                formula_col_refs = []
                is_sum = False
                sum_target_col = None

                if raw_val is not None:
                    val_s = str(raw_val).strip()
                    cand = val_s.lstrip("~'").strip()
                    if cand.startswith("{{@"):
                        val_type = "media"
                    elif val_s.startswith("="):
                        val_type = "formula"
                        formula_str = val_s
                        # Check for SUM e.g. =SUM(E8:E8)
                        m_sum = re.match(r"^=SUM\(([A-Z]+)\d+:([A-Z]+)\d+\)$", val_s, re.I)
                        if m_sum:
                            is_sum = True
                            sum_target_col = column_index_from_string(m_sum.group(1))
                        else:
                            # Extract col references e.g. =C8*D8 -> cols 3 and 4
                            cols_found = re.findall(r"([A-Z]+)\d+", val_s)
                            formula_col_refs = [column_index_from_string(col) for col in cols_found]
                    elif "{{" in val_s:
                        val_type = "var"
                    else:
                        val_type = "static"

                c_cells.append(CompiledCell(
                    col_idx=c,
                    colspan=colspan,
                    rowspan=rowspan,
                    is_skipped=False,
                    style_str=style_str,
                    raw_val=raw_val,
                    val_type=val_type,
                    num_format=cell.number_format if cell.has_style else None,
                    formula_str=formula_str,
                    formula_col_refs=formula_col_refs,
                    is_sum_formula=is_sum,
                    sum_target_col=sum_target_col
                ))

            # Detect directive on this row
            directive = cls._detect_row_directive(c_cells)
            all_rows.append(CompiledRow(
                row_num=r,
                height_pt=row_h,
                is_hidden=row_hidden,
                cells=c_cells,
                raw_directive=directive
            ))

        # 4. Group rows into hierarchical AST
        ast_elements = cls._build_elements_tree(all_rows, 0, len(all_rows))

        return CompiledTemplateAST(
            template_code=template_code,
            mtime=mtime,
            header_html=header_html,
            footer_html=footer_html,
            elements=ast_elements,
            column_count=max_col - min_col + 1,
            default_font=detected_font
        )

    @classmethod
    def _detect_row_directive(cls, cells: List[CompiledCell]) -> Optional[Dict[str, Any]]:
        for cell in cells:
            if cell.is_skipped or cell.raw_val is None:
                continue
            val = str(cell.raw_val).strip()
            if not val:
                continue

            m_each = re.match(r"^\{\{#each\s+([a-zA-Z0-9_\.]+)(.*?)\}\}$", val)
            if m_each:
                key = m_each.group(1)
                attrs = parse_tag_attributes(m_each.group(2))
                return {"type": "each_start", "key": key, "attrs": attrs}

            if val == "{{/each}}":
                return {"type": "each_end"}

            m_if = re.match(r"^\{\{#if\s+(.+)\}\}$", val)
            if m_if:
                return {"type": "if_start", "condition": m_if.group(1).strip()}

            if val == "{{/if}}":
                return {"type": "if_end"}

            if val == "{{#page-break}}":
                return {"type": "page_break"}
        return None

    @classmethod
    def _build_elements_tree(cls, rows: List[CompiledRow], start_idx: int, end_idx: int) -> List[CompiledElement]:
        elements: List[CompiledElement] = []
        i = start_idx
        while i < end_idx:
            row = rows[i]
            d = row.raw_directive
            if d:
                if d["type"] == "each_start":
                    close_idx = cls._find_close(rows, i + 1, end_idx, "each_start", "each_end")
                    children = cls._build_elements_tree(rows, i + 1, close_idx)
                    elements.append(EachBlockElement(
                        collection_key=d["key"],
                        attrs=d.get("attrs", {}),
                        children=children
                    ))
                    i = close_idx + 1
                    continue
                elif d["type"] == "if_start":
                    close_idx = cls._find_close(rows, i + 1, end_idx, "if_start", "if_end")
                    children = cls._build_elements_tree(rows, i + 1, close_idx)
                    elements.append(IfBlockElement(
                        condition=d["condition"],
                        children=children
                    ))
                    i = close_idx + 1
                    continue
                elif d["type"] == "page_break":
                    elements.append(PageBreakElement())
                    i += 1
                    continue
                elif d["type"] in ("each_end", "if_end"):
                    i += 1
                    continue

            elements.append(StaticRowElement(row=row))
            i += 1
        return elements

    @classmethod
    def _find_close(cls, rows: List[CompiledRow], start: int, end: int, open_t: str, close_t: str) -> int:
        depth = 1
        for idx in range(start, end):
            d = rows[idx].raw_directive
            if d:
                if d["type"] == open_t:
                    depth += 1
                elif d["type"] == close_t:
                    depth -= 1
                    if depth == 0:
                        return idx
        return end


class CompiledRenderer:
    """Executes a CompiledTemplateAST against dynamic JSON data in 1-2 milliseconds."""

    @classmethod
    def render(
        cls, 
        ast: CompiledTemplateAST, 
        data: Dict[str, Any], 
        title: str = "Report",
        font_family: Optional[str] = None
    ) -> str:
        context = data if isinstance(data, dict) else {}
        buf: List[str] = [ast.header_html]
        col_accumulators: Dict[int, float] = {}
        pending_page_break: List[bool] = [False]

        cls._render_elements(
            elements=ast.elements,
            context=context,
            buf=buf,
            col_accumulators=col_accumulators,
            pending_page_break=pending_page_break,
            is_top_level=True
        )

        buf.append(ast.footer_html)
        result_html = "".join(buf)

        if font_family:
            custom_font_css = FastHtmlRenderer._get_font_family_css(font_family)
            result_html = re.sub(
                r"body\s*\{[^}]*font-family:\s*[^;]+;",
                f"body {{ font-family: {custom_font_css};",
                result_html
            )

        return result_html

    @classmethod
    def _render_elements(
        cls,
        elements: List[CompiledElement],
        context: Dict[str, Any],
        buf: List[str],
        col_accumulators: Dict[int, float],
        pending_page_break: List[bool],
        is_top_level: bool = False
    ) -> None:
        for elem in elements:
            if isinstance(elem, StaticRowElement):
                is_pb = pending_page_break[0]
                pending_page_break[0] = False
                cls._render_static_row(elem.row, context, buf, col_accumulators, is_page_break=is_pb)
            elif isinstance(elem, PageBreakElement):
                pending_page_break[0] = True
            elif isinstance(elem, IfBlockElement):
                if evaluator.evaluate_condition(elem.condition, context):
                    cls._render_elements(elem.children, context, buf, col_accumulators, pending_page_break, is_top_level=False)
            elif isinstance(elem, EachBlockElement):
                raw_items = evaluator.evaluate_expression(elem.collection_key, context)
                if isinstance(raw_items, (list, tuple)):
                    items = list(raw_items)
                elif isinstance(raw_items, dict):
                    items = [raw_items]
                else:
                    items = None

                # Fallback logic identical to TemplateProcessor:
                if items is None:
                    if not context:
                        items = [{}]
                    elif elem.collection_key not in context and is_top_level:
                        items = [context]
                    else:
                        items = []
                elif len(items) == 0 and not context:
                    items = [{}]

                page_break = elem.attrs.get("page_break", False)
                for item_idx, item in enumerate(items):
                    item_ctx = {**context, **item} if isinstance(item, dict) else {**context, "item": item, "this": item}
                    item_ctx["@index"] = item_idx
                    item_ctx["@first"] = (item_idx == 0)
                    item_ctx["@last"] = (item_idx == len(items) - 1)

                    if page_break and item_idx > 0:
                        pending_page_break[0] = True

                    cls._render_elements(elem.children, item_ctx, buf, col_accumulators, pending_page_break, is_top_level=False)

    @classmethod
    def _render_static_row(
        cls,
        row: CompiledRow,
        context: Dict[str, Any],
        buf: List[str],
        col_accumulators: Dict[int, float],
        is_page_break: bool = False
    ) -> None:
        if row.is_hidden:
            return

        tr_classes = []
        tr_styles = []

        if is_page_break:
            tr_classes.append("page-break")
            tr_styles.append("page-break-before: always !important; break-before: page !important;")

        if row.height_pt:
            tr_styles.append(f"height: {row.height_pt}pt;")

        class_attr = f' class="{" ".join(tr_classes)}"' if tr_classes else ""
        style_attr = f' style="{" ".join(tr_styles)}"' if tr_styles else ""
        buf.append(f"<tr{class_attr}{style_attr}>\n")

        for cell in row.cells:
            if cell.is_skipped:
                continue

            span_attrs = []
            if cell.colspan > 1:
                span_attrs.append(f'colspan="{cell.colspan}"')
            if cell.rowspan > 1:
                span_attrs.append(f'rowspan="{cell.rowspan}"')
            span_str = " " + " ".join(span_attrs) if span_attrs else ""
            style_str = f' style="{cell.style_str}"' if cell.style_str else ""

            cell_val_html = cls._evaluate_cell_value(cell, context, col_accumulators)
            buf.append(f"<td{span_str}{style_str}>{cell_val_html}</td>\n")

        buf.append("</tr>\n")

    @classmethod
    def _evaluate_cell_value(
        cls,
        cell: CompiledCell,
        context: Dict[str, Any],
        col_accumulators: Dict[int, float]
    ) -> str:
        if cell.val_type == "empty" or cell.raw_val is None:
            return ""

        raw_str = str(cell.raw_val)

        # 1. Media tag
        if cell.val_type == "media":
            clean_tag = raw_str.lstrip("~'").strip()
            return cls._render_media(clean_tag, context)

        # 2. Formula
        if cell.val_type == "formula":
            if cell.is_sum_formula and cell.sum_target_col:
                total = col_accumulators.get(cell.sum_target_col, 0.0)
                return cls._format_num(total, cell.num_format)
            elif cell.formula_str:
                # Handle multiplication/arithmetic formula like =C8*D8
                val = cls._evaluate_row_arithmetic(cell.formula_str, context)
                if val is not None:
                    # Record for column accumulator
                    col_accumulators[cell.col_idx] = col_accumulators.get(cell.col_idx, 0.0) + float(val)
                    return cls._format_num(val, cell.num_format)
                return ""

        # 3. Variable placeholder
        if cell.val_type == "var":
            rendered = render_cell_value(cell.raw_val, context)
            if rendered is not None and rendered != "":
                # Try tracking numeric value into column accumulator if applicable
                try:
                    num_val = float(rendered)
                    col_accumulators[cell.col_idx] = col_accumulators.get(cell.col_idx, 0.0) + num_val
                except (ValueError, TypeError):
                    pass
                return html.escape(str(rendered))
            return ""

        # 4. Static value
        if isinstance(cell.raw_val, (int, float)):
            col_accumulators[cell.col_idx] = col_accumulators.get(cell.col_idx, 0.0) + float(cell.raw_val)
            return cls._format_num(cell.raw_val, cell.num_format)

        return html.escape(raw_str)

    @classmethod
    def _evaluate_row_arithmetic(cls, formula: str, context: Dict[str, Any]) -> Optional[float]:
        # Handle =TODAY() and =NOW()
        clean = formula.strip().upper()
        if clean == "=TODAY()":
            return date.today().strftime("%Y-%m-%d")
        if clean == "=NOW()":
            return datetime.now().strftime("%Y-%m-%d %H:%M")

        # Handle simple product like =C8*D8 or =C8+D8
        m_prod = re.match(r"^=([A-Z]+)\d+\s*([\*\+\-\/])\s*([A-Z]+)\d+$", formula.strip(), re.I)
        if m_prod:
            # Check context for common aliases or line total fields
            op = m_prod.group(2)
            # Find values from context items
            vals = []
            for k in ("qty", "quantity", "unit_price", "price", "rate", "ITEM_QTY", "ITEM_VALUE", "item_value"):
                if k in context:
                    try:
                        vals.append(float(context[k]))
                    except (ValueError, TypeError):
                        pass
            if len(vals) >= 2:
                if op == "*":
                    return vals[0] * vals[1]
                elif op == "+":
                    return vals[0] + vals[1]
                elif op == "-":
                    return vals[0] - vals[1]
                elif op == "/" and vals[1] != 0:
                    return vals[0] / vals[1]

        # Check if line_total or total is directly provided in context
        for total_key in ("line_total", "total", "net_value", "ITEM_NET_VALUE", "amount"):
            if total_key in context:
                try:
                    return float(context[total_key])
                except (ValueError, TypeError):
                    pass
        return 0.0

    @classmethod
    def _format_num(cls, val: Any, fmt: Optional[str]) -> str:
        if val is None or val == "":
            return ""
        try:
            num = float(val)
        except (ValueError, TypeError):
            return html.escape(str(val))

        if fmt:
            if "$" in fmt:
                return f"${num:,.2f}"
            elif "%" in fmt:
                return f"{num * 100:.1f}%"
            elif "#,##0.00" in fmt:
                return f"{num:,.2f}"
            elif "#,##0" in fmt:
                return f"{int(round(num)):,}"
        return f"{num:g}"

    @classmethod
    def _render_media(cls, tag_str: str, context: Dict[str, Any]) -> str:
        m_qr = re.match(r"^\{\{@qr\s+(.+?)\}\}$", tag_str)
        if m_qr:
            expr, attrs = parse_media_tag(m_qr.group(1).strip())
            data_val = evaluator.evaluate_expression(expr, context)
            if data_val:
                req_w = attrs.get("width", 80)
                req_h = attrs.get("height", 80)
                buf = MediaHandler.generate_qr_code(str(data_val), width=req_w, height=req_h)
                b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                return f'<img src="data:image/png;base64,{b64}" style="width: {req_w}px; height: {req_h}px; object-fit: contain; display: block;" />'
            return ""

        m_bc = re.match(r"^\{\{@barcode\s+(.+?)\}\}$", tag_str)
        if m_bc:
            expr, attrs = parse_media_tag(m_bc.group(1).strip())
            data_val = evaluator.evaluate_expression(expr, context)
            if data_val:
                req_w = attrs.get("width", 150)
                req_h = attrs.get("height", 40)
                show_text = attrs.get("text", False)
                buf = MediaHandler.generate_barcode(
                    str(data_val),
                    code_format=attrs.get("format", "code128"),
                    width=req_w,
                    height=req_h,
                    write_text=show_text
                )
                b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                return f'<img src="data:image/png;base64,{b64}" style="width: {req_w}px; height: {req_h}px; object-fit: contain; display: block;" />'
            return ""

        m_img = re.match(r"^\{\{@image\s+(.+?)\}\}$", tag_str)
        if m_img:
            expr, attrs = parse_media_tag(m_img.group(1).strip())
            src_val = evaluator.evaluate_expression(expr, context)
            if src_val:
                buf = MediaHandler.load_image_bytes(str(src_val))
                if buf:
                    req_w = attrs.get("width")
                    req_h = attrs.get("height")
                    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                    dim_style = ""
                    if req_w:
                        dim_style += f"width: {req_w}px; "
                    if req_h:
                        dim_style += f"height: {req_h}px; "
                    return f'<img src="data:image/png;base64,{b64}" style="{dim_style}object-fit: contain; display: block;" />'
            return ""

        return ""
