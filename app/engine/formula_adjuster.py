import re
from typing import Dict, List, Optional, Tuple

CELL_REF_REGEX = re.compile(r"(\$?[A-Za-z]{1,3})(\$?)(\d+)")
RANGE_REF_REGEX = re.compile(r"(\$?[A-Za-z]{1,3}\$?\d+):(\$?[A-Za-z]{1,3}\$?\d+)")

class FormulaAdjuster:
    """Adjusts Excel formulas when rows or columns are inserted, expanded, or shifted."""

    @classmethod
    def adjust_formula(
        cls,
        formula: str,
        orig_row: int,
        target_row: int,
        row_mapping: Dict[int, List[int]]
    ) -> str:
        """Adjusts a formula based on the exact template-to-target row mapping.
        
        Args:
            formula: e.g. '=B3*C3' or '=SUM(D3:D3)'
            orig_row: The template row index where this formula originated
            target_row: The target row index where this formula is placed
            row_mapping: Dict mapping orig_row -> [list of target_rows]
        """
        if not formula or not str(formula).startswith("="):
            return formula

        # Step 1: Adjust ranges (e.g. D3:D3 or D1:D3)
        ranges_found: List[str] = []

        def range_replacer(match: re.Match) -> str:
            start_ref = match.group(1)
            end_ref = match.group(2)

            start_m = CELL_REF_REGEX.match(start_ref)
            end_m = CELL_REF_REGEX.match(end_ref)
            if not start_m or not end_m:
                return match.group(0)

            s_col, s_pin, s_row = start_m.group(1), start_m.group(2), int(start_m.group(3))
            e_col, e_pin, e_row = end_m.group(1), end_m.group(2), int(end_m.group(3))

            # Resolve start row
            if s_row in row_mapping and row_mapping[s_row]:
                new_s_row = row_mapping[s_row][0]
            else:
                new_s_row = s_row

            # Resolve end row
            if e_row in row_mapping and row_mapping[e_row]:
                new_e_row = row_mapping[e_row][-1]
            else:
                new_e_row = e_row

            return f"{s_col}{s_pin}{new_s_row}:{e_col}{e_pin}{new_e_row}"

        def token_sub(m: re.Match) -> str:
            adjusted = range_replacer(m)
            idx = len(ranges_found)
            ranges_found.append(adjusted)
            return f"__RANGE_TOKEN_{idx}__"

        temp_formula = RANGE_REF_REGEX.sub(token_sub, formula)

        # Step 2: Adjust individual cell references
        def single_cell_replacer(match: re.Match) -> str:
            col = match.group(1)
            row_pinned = match.group(2) == "$"
            ref_row = int(match.group(3))

            if row_pinned:
                return match.group(0)

            # If formula references the row it was written in (e.g. =B3*C3 on row 3 or =D12/B12 on row 12)
            if ref_row == orig_row:
                return f"{col}{target_row}"

            # If formula is in a repeating loop and references a relative preceding/following row (e.g. =F11+D12-E12 on row 12)
            is_repeating = (orig_row in row_mapping and len(row_mapping[orig_row]) > 1)
            offset = ref_row - orig_row
            if is_repeating and abs(offset) <= 5:
                # If ref_row is not separately expanding
                if ref_row not in row_mapping or len(row_mapping.get(ref_row, [])) <= 1:
                    new_ref_row = target_row + offset
                    return f"{col}{new_ref_row}"

            # If referencing another mapped row that expanded
            if ref_row in row_mapping and row_mapping[ref_row]:
                mapped_list = row_mapping[ref_row]
                if orig_row in row_mapping and target_row in row_mapping[orig_row]:
                    idx = row_mapping[orig_row].index(target_row)
                    if idx < len(mapped_list):
                        return f"{col}{mapped_list[idx]}"
                return f"{col}{mapped_list[0]}"

            return match.group(0)

        adjusted_formula = CELL_REF_REGEX.sub(single_cell_replacer, temp_formula)

        # Step 3: Restore range tokens
        for idx, r_str in enumerate(ranges_found):
            adjusted_formula = adjusted_formula.replace(f"__RANGE_TOKEN_{idx}__", r_str)

        return adjusted_formula
