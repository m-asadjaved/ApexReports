import re
import math
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Set, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import Rule
from openpyxl.utils.cell import range_boundaries, coordinate_to_tuple
from openpyxl.utils import get_column_letter
from app.renderer.color_resolver import ExcelColorResolver


@dataclass
class ConditionalStyleOverlay:
    """Style properties applied to a cell when a conditional formatting rule matches."""
    bg_color: Optional[str] = None
    font_color: Optional[str] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    border_top: Optional[str] = None
    border_bottom: Optional[str] = None
    border_left: Optional[str] = None
    border_right: Optional[str] = None


class ConditionalFormattingEvaluator:
    """
    Evaluates openpyxl conditional formatting rules against worksheet cells.
    Supports:
      - CellIs rules (greaterThan, lessThan, equal, notEqual, between, notBetween, greaterThanOrEqual, lessThanOrEqual)
      - Text rules (containsText, notContains, beginsWith, endsWith)
      - Expression / Formula rules
      - Color Scales (2-color and 3-color gradient heatmaps)
      - AboveAverage / BelowAverage rules
      - Blanks / Non-blanks rules
    Returns a mapping of (row, col) -> ConditionalStyleOverlay.
    """

    @classmethod
    def evaluate_worksheet(
        cls, 
        ws: Worksheet, 
        calc_cache: Optional[Dict[Tuple[int, int], Any]] = None,
        theme_colors: Optional[List[str]] = None
    ) -> Dict[Tuple[int, int], ConditionalStyleOverlay]:
        """Evaluates all conditional formatting rules in the worksheet and returns overlays for matching cells."""
        overlays: Dict[Tuple[int, int], ConditionalStyleOverlay] = {}
        calc_cache = calc_cache or {}

        if not hasattr(ws, "conditional_formatting") or not ws.conditional_formatting:
            return overlays

        # Track which cells have encountered a stopIfTrue rule
        stopped_cells: Set[Tuple[int, int]] = set()

        for cf in ws.conditional_formatting:
            ranges = getattr(cf.sqref, "ranges", [])
            if not ranges:
                continue

            for rule in cf.rules:
                rule_type = getattr(rule, "type", None)

                # Special handling for ColorScale rules
                if rule_type == "colorScale" and getattr(rule, "colorScale", None):
                    cls._evaluate_color_scale(ws, ranges, rule, calc_cache, overlays, stopped_cells, theme_colors)
                    continue

                # For cell-level rules, iterate over cells in ranges
                for rng in ranges:
                    min_c, min_r, max_c, max_r = rng.min_col, rng.min_row, rng.max_col, rng.max_row
                    
                    # Pre-calculate average if rule is average-based
                    avg_val = None
                    if rule_type in ("aboveAverage", "belowAverage", "equalAverage"):
                        avg_val = cls._compute_range_average(ws, rng, calc_cache)

                    for r in range(min_r, max_r + 1):
                        for c in range(min_c, max_c + 1):
                            cell_coord = (r, c)
                            if cell_coord in stopped_cells:
                                continue

                            val = calc_cache.get(cell_coord, ws.cell(row=r, column=c).value)
                            if cls._match_rule(rule, val, avg_val, ws, r, c, calc_cache):
                                overlay = cls._extract_rule_style(rule, theme_colors)
                                if overlay:
                                    if cell_coord not in overlays:
                                        overlays[cell_coord] = overlay
                                    else:
                                        # Merge overlay (higher priority rule properties take precedence)
                                        cls._merge_overlay(overlays[cell_coord], overlay)

                                if getattr(rule, "stopIfTrue", False):
                                    stopped_cells.add(cell_coord)

        return overlays

    @classmethod
    def _match_rule(
        cls, 
        rule: Rule, 
        val: Any, 
        avg_val: Optional[float],
        ws: Worksheet, 
        r: int, 
        c: int, 
        calc_cache: Dict[Tuple[int, int], Any]
    ) -> bool:
        """Determines whether a cell value matches a specific conditional formatting rule."""
        rule_type = getattr(rule, "type", None)
        operator = getattr(rule, "operator", None)
        formula = getattr(rule, "formula", []) or []

        # 1. CellIs rules
        if rule_type == "cellIs" or operator in (
            "greaterThan", "lessThan", "equal", "notEqual", 
            "greaterThanOrEqual", "lessThanOrEqual", "between", "notBetween"
        ):
            return cls._evaluate_cell_is(operator, formula, val, ws, r, c, calc_cache)

        # 2. Text rules
        if rule_type in ("containsText", "notContains", "beginsWith", "endsWith") or operator in (
            "containsText", "notContains", "beginsWith", "endsWith"
        ):
            return cls._evaluate_text_rule(rule_type or operator, rule, val)

        # 3. Blanks / Errors
        if rule_type == "containsBlanks":
            return val is None or str(val).strip() == ""
        if rule_type == "notContainsBlanks":
            return val is not None and str(val).strip() != ""
        if rule_type == "containsErrors":
            return isinstance(val, str) and val.startswith("#")

        # 4. Average rules
        if rule_type == "aboveAverage":
            num_val = cls._to_number(val)
            return num_val is not None and avg_val is not None and num_val > avg_val
        if rule_type == "belowAverage":
            num_val = cls._to_number(val)
            return num_val is not None and avg_val is not None and num_val < avg_val

        # 5. Expression / Formula rule (e.g. "=B2>100")
        if rule_type == "expression" and formula:
            return cls._evaluate_expression_rule(formula[0], ws, r, c, calc_cache)

        return False

    @classmethod
    def _evaluate_cell_is(
        cls, 
        operator: Optional[str], 
        formula: List[str], 
        val: Any,
        ws: Worksheet,
        r: int,
        c: int,
        calc_cache: Dict[Tuple[int, int], Any]
    ) -> bool:
        """Evaluates comparison operators for CellIs rules."""
        if not operator or not formula:
            return False

        op1 = cls._resolve_operand(formula[0], ws, r, c, calc_cache)
        val_num = cls._to_number(val)
        op1_num = cls._to_number(op1)

        # Numeric comparison if both can be converted to numbers
        if val_num is not None and op1_num is not None:
            if operator == "greaterThan":
                return val_num > op1_num
            elif operator == "lessThan":
                return val_num < op1_num
            elif operator in ("equal", "equalTo"):
                return abs(val_num - op1_num) < 1e-9
            elif operator in ("notEqual", "notEqualTo"):
                return abs(val_num - op1_num) >= 1e-9
            elif operator == "greaterThanOrEqual":
                return val_num >= op1_num
            elif operator == "lessThanOrEqual":
                return val_num <= op1_num
            elif operator == "between" and len(formula) > 1:
                op2 = cls._resolve_operand(formula[1], ws, r, c, calc_cache)
                op2_num = cls._to_number(op2)
                if op2_num is not None:
                    low = min(op1_num, op2_num)
                    high = max(op1_num, op2_num)
                    return low <= val_num <= high
            elif operator == "notBetween" and len(formula) > 1:
                op2 = cls._resolve_operand(formula[1], ws, r, c, calc_cache)
                op2_num = cls._to_number(op2)
                if op2_num is not None:
                    low = min(op1_num, op2_num)
                    high = max(op1_num, op2_num)
                    return val_num < low or val_num > high

        # Fallback to string comparison
        val_str = str(val or "").strip().lower()
        op1_str = str(op1 or "").strip().strip('"\'').lower()

        if operator in ("equal", "equalTo"):
            return val_str == op1_str
        elif operator in ("notEqual", "notEqualTo"):
            return val_str != op1_str

        return False

    @classmethod
    def _evaluate_text_rule(cls, rule_type: str, rule: Rule, val: Any) -> bool:
        """Evaluates substring / text match rules."""
        val_str = str(val or "").strip().lower()
        target_text = getattr(rule, "text", None)
        if not target_text and getattr(rule, "formula", None):
            target_text = rule.formula[0]
        if not target_text:
            return False

        target_text = target_text.strip('"\'').lower()

        if rule_type == "containsText":
            return target_text in val_str
        elif rule_type == "notContains":
            return target_text not in val_str
        elif rule_type == "beginsWith":
            return val_str.startswith(target_text)
        elif rule_type == "endsWith":
            return val_str.endswith(target_text)

        return False

    @classmethod
    def _evaluate_expression_rule(
        cls, 
        expr_str: str, 
        ws: Worksheet, 
        r: int, 
        c: int, 
        calc_cache: Dict[Tuple[int, int], Any]
    ) -> bool:
        """Evaluates formula expression rules such as '=$B2>1000' or '=A1=\"PAID\"'."""
        clean_expr = expr_str.strip()
        if clean_expr.startswith("="):
            clean_expr = clean_expr[1:]

        # Simple comparison regex: (CELL_REF)(>=|<=|!=|=|>|<)(VALUE)
        m = re.match(r"^(\$?[A-Z]+\$?[0-9]+)\s*(>=|<=|!=|=|>|<)\s*(.+)$", clean_expr, re.IGNORECASE)
        if m:
            ref_str, op, target_raw = m.group(1), m.group(2), m.group(3)
            # Resolve cell reference in row context
            ref_val = cls._resolve_cell_ref(ref_str, r, c, ws, calc_cache)
            target_val = target_raw.strip().strip('"\'')

            ref_num = cls._to_number(ref_val)
            target_num = cls._to_number(target_val)

            if ref_num is not None and target_num is not None:
                if op == ">": return ref_num > target_num
                if op == "<": return ref_num < target_num
                if op == ">=": return ref_num >= target_num
                if op == "<=": return ref_num <= target_num
                if op == "=": return abs(ref_num - target_num) < 1e-9
                if op == "!=": return abs(ref_num - target_num) >= 1e-9

            # String comparison
            ref_s = str(ref_val or "").strip().lower()
            target_s = target_val.lower()
            if op == "=": return ref_s == target_s
            if op == "!=": return ref_s != target_s

        return False

    @classmethod
    def _evaluate_color_scale(
        cls,
        ws: Worksheet,
        ranges: List[Any],
        rule: Rule,
        calc_cache: Dict[Tuple[int, int], Any],
        overlays: Dict[Tuple[int, int], ConditionalStyleOverlay],
        stopped_cells: Set[Tuple[int, int]],
        theme_colors: Optional[List[str]]
    ) -> None:
        """Evaluates 2-color and 3-color scales (heatmaps) across cell ranges."""
        cs = rule.colorScale
        colors = getattr(cs, "color", [])
        if len(colors) < 2:
            return

        # 1. Collect all numeric values across all ranges in the rule
        cell_values: Dict[Tuple[int, int], float] = {}
        for rng in ranges:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    coord = (r, c)
                    if coord in stopped_cells:
                        continue
                    val = calc_cache.get(coord, ws.cell(row=r, column=c).value)
                    num_val = cls._to_number(val)
                    if num_val is not None:
                        cell_values[coord] = num_val

        if not cell_values:
            return

        min_v = min(cell_values.values())
        max_v = max(cell_values.values())
        if abs(max_v - min_v) < 1e-9:
            ratio_range = 1.0
        else:
            ratio_range = max_v - min_v

        # Extract RGB hex strings for colors
        hex_colors: List[Tuple[int, int, int]] = []
        for col_obj in colors:
            hex_str = ExcelColorResolver.resolve_color(col_obj, theme_colors, default="#ffffff")
            rgb_tup = cls._hex_to_rgb(hex_str)
            hex_colors.append(rgb_tup)

        for coord, v in cell_values.items():
            norm = (v - min_v) / ratio_range
            norm = max(0.0, min(1.0, norm))

            if len(hex_colors) == 2:
                interp_rgb = cls._interpolate_rgb(hex_colors[0], hex_colors[1], norm)
            elif len(hex_colors) >= 3:
                if norm < 0.5:
                    interp_rgb = cls._interpolate_rgb(hex_colors[0], hex_colors[1], norm * 2.0)
                else:
                    interp_rgb = cls._interpolate_rgb(hex_colors[1], hex_colors[2], (norm - 0.5) * 2.0)
            else:
                interp_rgb = hex_colors[0]

            bg_hex = f"#{interp_rgb[0]:02x}{interp_rgb[1]:02x}{interp_rgb[2]:02x}"
            if coord not in overlays:
                overlays[coord] = ConditionalStyleOverlay(bg_color=bg_hex)
            else:
                overlays[coord].bg_color = bg_hex

    @classmethod
    def _extract_rule_style(cls, rule: Rule, theme_colors: Optional[List[str]]) -> Optional[ConditionalStyleOverlay]:
        """Extracts fill, font, and border styles from a Rule's differential style (dxf)."""
        dxf = getattr(rule, "dxf", None)
        if not dxf:
            return None

        overlay = ConditionalStyleOverlay()

        # Fill
        fill = getattr(dxf, "fill", None)
        if fill:
            bg_hex = ExcelColorResolver.resolve_fill(fill, theme_colors)
            if bg_hex:
                overlay.bg_color = bg_hex

        # Font
        font = getattr(dxf, "font", None)
        if font:
            if getattr(font, "bold", None) is not None:
                overlay.bold = bool(font.bold)
            if getattr(font, "italic", None) is not None:
                overlay.italic = bool(font.italic)
            f_color = ExcelColorResolver.resolve_font_color(font, theme_colors)
            if f_color:
                overlay.font_color = f_color

        # Border
        border = getattr(dxf, "border", None)
        if border:
            if getattr(border, "top", None) and getattr(border.top, "style", None):
                overlay.border_top = f"1px solid {ExcelColorResolver.resolve_color(border.top.color, theme_colors, default='#000000')}"
            if getattr(border, "bottom", None) and getattr(border.bottom, "style", None):
                overlay.border_bottom = f"1px solid {ExcelColorResolver.resolve_color(border.bottom.color, theme_colors, default='#000000')}"
            if getattr(border, "left", None) and getattr(border.left, "style", None):
                overlay.border_left = f"1px solid {ExcelColorResolver.resolve_color(border.left.color, theme_colors, default='#000000')}"
            if getattr(border, "right", None) and getattr(border.right, "style", None):
                overlay.border_right = f"1px solid {ExcelColorResolver.resolve_color(border.right.color, theme_colors, default='#000000')}"

        return overlay

    @classmethod
    def _merge_overlay(cls, base: ConditionalStyleOverlay, incoming: ConditionalStyleOverlay) -> None:
        """Merges incoming styles into base, letting base (earlier rule) take precedence if already set."""
        if base.bg_color is None and incoming.bg_color:
            base.bg_color = incoming.bg_color
        if base.font_color is None and incoming.font_color:
            base.font_color = incoming.font_color
        if base.bold is None and incoming.bold is not None:
            base.bold = incoming.bold
        if base.italic is None and incoming.italic is not None:
            base.italic = incoming.italic
        if base.border_top is None and incoming.border_top:
            base.border_top = incoming.border_top
        if base.border_bottom is None and incoming.border_bottom:
            base.border_bottom = incoming.border_bottom
        if base.border_left is None and incoming.border_left:
            base.border_left = incoming.border_left
        if base.border_right is None and incoming.border_right:
            base.border_right = incoming.border_right

    @classmethod
    def _to_number(cls, val: Any) -> Optional[float]:
        """Safely converts value to float if possible."""
        if val is None or isinstance(val, bool):
            return None
        if isinstance(val, (int, float)):
            return float(val)
        try:
            # Strip currency symbols and whitespace
            clean = str(val).strip().replace("$", "").replace(",", "").replace("%", "")
            return float(clean)
        except (ValueError, TypeError):
            return None

    @classmethod
    def _compute_range_average(cls, ws: Worksheet, rng: Any, calc_cache: Dict[Tuple[int, int], Any]) -> Optional[float]:
        """Computes the arithmetic average of numeric cells in a range."""
        total = 0.0
        count = 0
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                val = calc_cache.get((r, c), ws.cell(row=r, column=c).value)
                num = cls._to_number(val)
                if num is not None:
                    total += num
                    count += 1
        return (total / count) if count > 0 else None

    @classmethod
    def _resolve_operand(
        cls, 
        operand_str: str, 
        ws: Worksheet, 
        r: int, 
        c: int, 
        calc_cache: Dict[Tuple[int, int], Any]
    ) -> Any:
        """Resolves operand string to literal or cell reference."""
        clean = operand_str.strip()
        if clean.startswith("="):
            clean = clean[1:]
        # Check if cell reference like $A$1 or B2
        m = re.match(r"^\$?([A-Z]+)\$?([0-9]+)$", clean, re.IGNORECASE)
        if m:
            return cls._resolve_cell_ref(clean, r, c, ws, calc_cache)
        return clean.strip('"\'')

    @classmethod
    def _resolve_cell_ref(
        cls, 
        ref_str: str, 
        current_r: int, 
        current_c: int, 
        ws: Worksheet, 
        calc_cache: Dict[Tuple[int, int], Any]
    ) -> Any:
        """Resolves an Excel cell reference (with absolute/relative coordinates) relative to current row/col."""
        # e.g. $B2 -> column B (fixed), row 2 (relative or absolute)
        clean = ref_str.replace("$", "").upper()
        try:
            target_r, target_c = coordinate_to_tuple(clean)
            return calc_cache.get((target_r, target_c), ws.cell(row=target_r, column=target_c).value)
        except Exception:
            return None

    @classmethod
    def _hex_to_rgb(cls, hex_str: str) -> Tuple[int, int, int]:
        """Converts '#RRGGBB' to (R, G, B) tuple."""
        clean = hex_str.lstrip("#")
        if len(clean) == 8:  # AARRGGBB
            clean = clean[2:]
        if len(clean) == 6:
            return (int(clean[0:2], 16), int(clean[2:4], 16), int(clean[4:6], 16))
        return (255, 255, 255)

    @classmethod
    def _interpolate_rgb(cls, rgb1: Tuple[int, int, int], rgb2: Tuple[int, int, int], ratio: float) -> Tuple[int, int, int]:
        """Linearly interpolates between two RGB colors by a ratio in [0, 1]."""
        r = int(round(rgb1[0] + ratio * (rgb2[0] - rgb1[0])))
        g = int(round(rgb1[1] + ratio * (rgb2[1] - rgb1[1])))
        b = int(round(rgb1[2] + ratio * (rgb2[2] - rgb1[2])))
        return (max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))
