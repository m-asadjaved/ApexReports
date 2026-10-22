import io
import re
from typing import Any, Dict, List, Optional, Tuple, Set
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils import get_column_letter

from app.engine.template_parser import ExpressionEvaluator, parse_tag_attributes, TAG_REGEX
from app.engine.cell_utils import copy_cell_style, render_cell_value
from app.engine.formula_adjuster import FormulaAdjuster
from app.engine.media_handler import MediaHandler
from app.engine.matrix_engine import MatrixEngine
from app.core.exceptions import TemplateSyntaxException

evaluator = ExpressionEvaluator()

class TemplateProcessor:
    """Core processor for Excel report templates with Oracle Reports 6i capabilities."""

    def __init__(self, workbook: Workbook):
        self.wb = workbook
        self.src_sheet: Worksheet = workbook.active
        self.target_wb = Workbook()
        self.target_sheet: Worksheet = self.target_wb.active
        self.target_sheet.title = self.src_sheet.title

        # Global row mapping: orig_row -> all target rows
        self.global_row_mapping: Dict[int, List[int]] = {}
        # Scoped row mapping: tracks the most recent rows generated in current group/iteration
        self.recent_row_mapping: Dict[int, List[int]] = {}
        # Pending formulas: (target_row, target_col, formula_str, orig_row, scoped_mapping)
        self.pending_formulas: List[Tuple[int, int, str, int, Dict[int, List[int]]]] = []
        # Page break row indices
        self.page_breaks: List[int] = []

    def process(self, data: Dict[str, Any]) -> Workbook:
        """Processes the template using provided data dictionary and returns populated Workbook."""
        # 1. Expand Matrix columns if {{#each-col}} tags are present
        MatrixEngine.expand_columns(self.src_sheet, data)

        # 2. Copy sheet properties and column dimensions
        self._copy_sheet_setup()

        template_rows = list(self.src_sheet.iter_rows(values_only=False))
        total_rows = len(template_rows)

        self._execute_block(
            rows=template_rows,
            start_idx=0,
            end_idx=total_rows,
            context=data,
            target_row_start=1
        )

        self._resolve_formulas()

        for pb_row in self.page_breaks:
            self.target_sheet.row_breaks.append(Break(id=pb_row))

        self._copy_merged_cells()

        return self.target_wb

    def _copy_sheet_setup(self) -> None:
        """Copies column widths, print setups, and margins from source to target."""
        for col_letter, col_dim in self.src_sheet.column_dimensions.items():
            self.target_sheet.column_dimensions[col_letter].width = col_dim.width
            self.target_sheet.column_dimensions[col_letter].hidden = col_dim.hidden

        if self.src_sheet.print_title_rows:
            self.target_sheet.print_title_rows = self.src_sheet.print_title_rows
        if self.src_sheet.print_title_cols:
            self.target_sheet.print_title_cols = self.src_sheet.print_title_cols

        try:
            if self.src_sheet.page_setup.orientation:
                self.target_sheet.page_setup.orientation = self.src_sheet.page_setup.orientation
            if self.src_sheet.page_setup.paperSize:
                self.target_sheet.page_setup.paperSize = self.src_sheet.page_setup.paperSize
        except Exception:
            pass

        try:
            if self.src_sheet.views.sheetView:
                self.target_sheet.views.sheetView[0].showGridLines = self.src_sheet.views.sheetView[0].showGridLines
        except Exception:
            pass

    def _execute_block(
        self,
        rows: List[Tuple],
        start_idx: int,
        end_idx: int,
        context: Dict[str, Any],
        target_row_start: int
    ) -> int:
        """Recursively processes a block of rows."""
        target_r = target_row_start
        i = start_idx

        while i < end_idx:
            row_cells = rows[i]
            orig_row_num = i + 1

            directive = self._detect_row_directive(row_cells)

            if directive and directive["type"] == "each_start":
                close_idx = self._find_matching_close(rows, i + 1, end_idx, "each_start", "each_end")
                list_key = directive["key"]
                items = evaluator.evaluate_expression(list_key, context)
                if not isinstance(items, (list, tuple)):
                    items = []

                page_break = directive.get("attrs", {}).get("page_break", False)

                # Reset recent mapping for lines inside this loop before it runs
                for r_idx in range(i + 1, close_idx):
                    self.recent_row_mapping[r_idx + 1] = []

                for item_idx, item in enumerate(items):
                    item_context = {**context, **item} if isinstance(item, dict) else {**context, "item": item, "this": item}
                    item_context["@index"] = item_idx
                    item_context["@first"] = (item_idx == 0)
                    item_context["@last"] = (item_idx == len(items) - 1)

                    if page_break and item_idx > 0:
                        self.page_breaks.append(target_r - 1)

                    target_r = self._execute_block(
                        rows=rows,
                        start_idx=i + 1,
                        end_idx=close_idx,
                        context=item_context,
                        target_row_start=target_r
                    )

                i = close_idx + 1
                continue

            elif directive and directive["type"] == "if_start":
                close_idx = self._find_matching_close(rows, i + 1, end_idx, "if_start", "if_end")
                condition = directive["condition"]
                is_true = evaluator.evaluate_condition(condition, context)

                if is_true:
                    target_r = self._execute_block(
                        rows=rows,
                        start_idx=i + 1,
                        end_idx=close_idx,
                        context=context,
                        target_row_start=target_r
                    )
                i = close_idx + 1
                continue

            elif directive and directive["type"] == "page_break":
                self.page_breaks.append(target_r - 1)
                i += 1
                continue

            elif directive and directive["type"] in ("each_end", "if_end"):
                i += 1
                continue

            else:
                self._emit_row(row_cells, orig_row_num, target_r, context)
                
                # Update global and recent mappings
                if orig_row_num not in self.global_row_mapping:
                    self.global_row_mapping[orig_row_num] = []
                self.global_row_mapping[orig_row_num].append(target_r)

                if orig_row_num not in self.recent_row_mapping:
                    self.recent_row_mapping[orig_row_num] = []
                self.recent_row_mapping[orig_row_num].append(target_r)

                target_r += 1
                i += 1

        return target_r

    def _detect_row_directive(self, row_cells: Tuple) -> Optional[Dict[str, Any]]:
        """Inspects cells in a row to determine if the entire row is a directive."""
        for c in row_cells:
            val = str(c.value or "").strip()
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

    def _find_matching_close(
        self, 
        rows: List[Tuple], 
        start: int, 
        end: int, 
        open_type: str, 
        close_type: str
    ) -> int:
        depth = 1
        for idx in range(start, end):
            d = self._detect_row_directive(rows[idx])
            if d:
                if d["type"] == open_type:
                    depth += 1
                elif d["type"] == close_type:
                    depth -= 1
                    if depth == 0:
                        return idx
        raise TemplateSyntaxException(f"Unclosed '{open_type}' block starting around line {start}")

    def _emit_row(
        self, 
        src_cells: Tuple, 
        orig_row_num: int, 
        target_row_num: int, 
        context: Dict[str, Any]
    ) -> None:
        """Copies and renders a single row into the target sheet."""
        src_row_dim = self.src_sheet.row_dimensions.get(orig_row_num)
        if src_row_dim and src_row_dim.height is not None:
            self.target_sheet.row_dimensions[target_row_num].height = src_row_dim.height

        # Snapshot recent mapping for this row
        scoped_mapping = {k: list(v) for k, v in self.recent_row_mapping.items()}

        for col_idx, src_cell in enumerate(src_cells, start=1):
            dst_cell = self.target_sheet.cell(row=target_row_num, column=col_idx)
            copy_cell_style(src_cell, dst_cell)

            raw_val = src_cell.value
            if raw_val is None:
                continue

            val_str = str(raw_val).strip()

            if val_str.startswith("="):
                self.pending_formulas.append((target_row_num, col_idx, val_str, orig_row_num, scoped_mapping))
                dst_cell.value = val_str
                continue

            if val_str.startswith("{{@"):
                self._handle_media_tag(val_str, dst_cell, target_row_num, col_idx, context)
                continue

            dst_cell.value = render_cell_value(raw_val, context)

    def _handle_media_tag(
        self, 
        tag_str: str, 
        cell: Any, 
        row: int, 
        col: int, 
        context: Dict[str, Any]
    ) -> None:
        cell_coord = f"{get_column_letter(col)}{row}"
        cell.value = ""

        m_qr = re.match(r"^\{\{@qr\s+(.+?)\}\}$", tag_str)
        if m_qr:
            parts = m_qr.group(1).split()
            expr = parts[0]
            attrs = parse_tag_attributes(" ".join(parts[1:]))
            data_val = evaluator.evaluate_expression(expr, context)
            if data_val:
                buf = MediaHandler.generate_qr_code(str(data_val))
                MediaHandler.embed_image_in_cell(
                    self.target_sheet,
                    cell_coord,
                    buf,
                    width=attrs.get("width", 80),
                    height=attrs.get("height", 80)
                )
            return

        m_bc = re.match(r"^\{\{@barcode\s+(.+?)\}\}$", tag_str)
        if m_bc:
            parts = m_bc.group(1).split()
            expr = parts[0]
            attrs = parse_tag_attributes(" ".join(parts[1:]))
            data_val = evaluator.evaluate_expression(expr, context)
            if data_val:
                buf = MediaHandler.generate_barcode(str(data_val), code_format=attrs.get("format", "code128"))
                MediaHandler.embed_image_in_cell(
                    self.target_sheet,
                    cell_coord,
                    buf,
                    width=attrs.get("width", 150),
                    height=attrs.get("height", 40)
                )
            return

        m_img = re.match(r"^\{\{@image\s+(.+?)\}\}$", tag_str)
        if m_img:
            parts = m_img.group(1).split()
            expr = parts[0]
            attrs = parse_tag_attributes(" ".join(parts[1:]))
            src_val = evaluator.evaluate_expression(expr, context)
            if src_val:
                buf = MediaHandler.load_image_bytes(str(src_val))
                if buf:
                    MediaHandler.embed_image_in_cell(
                        self.target_sheet,
                        cell_coord,
                        buf,
                        width=attrs.get("width"),
                        height=attrs.get("height")
                    )
            return

    def _resolve_formulas(self) -> None:
        """Adapts formulas for repeating rows and expands summary ranges."""
        for target_row, target_col, formula_str, orig_row, scoped_mapping in self.pending_formulas:
            cell = self.target_sheet.cell(row=target_row, column=target_col)
            
            # Prefer scoped mapping if it exists and has the referenced row(s), else global
            effective_mapping = {**self.global_row_mapping}
            for k, v in scoped_mapping.items():
                if v:
                    effective_mapping[k] = v

            adjusted = FormulaAdjuster.adjust_formula(
                formula=formula_str,
                orig_row=orig_row,
                target_row=target_row,
                row_mapping=effective_mapping
            )
            cell.value = adjusted

    def _copy_merged_cells(self) -> None:
        for rng in self.src_sheet.merged_cells.ranges:
            min_r, min_c, max_r, max_c = rng.min_row, rng.min_col, rng.max_row, rng.max_col
            if min_r in self.global_row_mapping and max_r in self.global_row_mapping:
                mapped_start = self.global_row_mapping[min_r]
                mapped_end = self.global_row_mapping[max_r]
                if mapped_start and mapped_end:
                    target_min_r = mapped_start[0]
                    target_max_r = mapped_end[-1]
                    try:
                        self.target_sheet.merge_cells(
                            start_row=target_min_r,
                            start_column=min_c,
                            end_row=target_max_r,
                            end_column=max_c
                        )
                    except Exception:
                        pass
