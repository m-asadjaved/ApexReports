import math
import re
import datetime
from typing import Any, Dict, List, Optional, Tuple, Union
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter, column_index_from_string

CELL_COORD_REGEX = re.compile(r"(\$?[A-Za-z]{1,3})(\$?)(\d+)")
RANGE_REGEX = re.compile(r"(\$?[A-Za-z]{1,3}\$?\d+):(\$?[A-Za-z]{1,3}\$?\d+)")

# Matches function calls: FUNCTION_NAME(args...)
FUNC_CALL_REGEX = re.compile(
    r"\b(SUM|AVERAGE|AVERAGEA|COUNT|COUNTA|COUNTBLANK|MIN|MAX|PRODUCT|SUBTOTAL|ROUND|ROUNDUP|ROUNDDOWN|ABS|INT|MOD|IF|IFERROR|NVL|CONCAT|CONCATENATE|TODAY|NOW)\((.*?)\)",
    re.IGNORECASE
)

class FormulaEvaluator:
    """In-memory formula calculator for Excel spreadsheet reports.
    Supports arithmetic, aggregations (SUM, AVERAGE, COUNT, COUNTA, etc.),
    rounding, conditionals, strings, and dates.
    """

    @classmethod
    def evaluate_worksheet(cls, ws: Worksheet) -> Dict[Tuple[int, int], Any]:
        """Calculates values for all formula cells in the worksheet.
        Returns a dictionary mapping (row, col) -> evaluated_value.
        """
        calculated_values: Dict[Tuple[int, int], Any] = {}
        max_row = ws.max_row or 1
        max_col = ws.max_column or 1

        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                val = cell.value
                if val is not None and isinstance(val, str) and val.startswith("="):
                    res = cls.evaluate_formula(val, ws, calculated_values)
                    calculated_values[(r, c)] = res

        return calculated_values

    @classmethod
    def _resolve_val(cls, token: str, ws: Worksheet, cache: Dict[Tuple[int, int], Any]) -> Any:
        """Resolves a token which could be a cell coordinate ($B$5), number, or string."""
        token = token.strip()
        cleaned_token = token.replace("$", "")
        m = CELL_COORD_REGEX.fullmatch(cleaned_token)
        if m:
            col_str = m.group(1)
            row_num = int(m.group(3))
            col_num = column_index_from_string(col_str)
            val = cache.get((row_num, col_num))
            if val is None:
                val = ws.cell(row=row_num, column=col_num).value
            return val
        if cls._is_numeric(token):
            return float(token)
        return token.strip('"\'')

    @classmethod
    def evaluate_formula(
        cls, 
        formula: str, 
        ws: Worksheet, 
        cache: Dict[Tuple[int, int], Any]
    ) -> Any:
        """Evaluates a single formula string e.g. '=COUNTA(A9:A11)' or '=SUM(E6:E8)'."""
        if not formula or not str(formula).startswith("="):
            return formula

        expr = formula[1:].strip()

        # Step 1: Evaluate Excel functions iteratively
        max_iterations = 6
        while max_iterations > 0 and FUNC_CALL_REGEX.search(expr):
            max_iterations -= 1

            def func_replacer(match: re.Match) -> str:
                func_name = match.group(1).upper()
                args_str = match.group(2).strip()

                if func_name == "TODAY":
                    return f'"{datetime.date.today().isoformat()}"'
                if func_name == "NOW":
                    return f'"{datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}"'

                # Extract raw cell values
                values = cls._extract_values_from_args(args_str, ws, cache)

                if func_name == "COUNTA":
                    count_non_empty = sum(
                        1 for v in values if v is not None and str(v).strip() != ""
                    )
                    return str(count_non_empty)

                elif func_name == "COUNT":
                    count_num = sum(1 for v in values if cls._is_numeric(v))
                    return str(count_num)

                elif func_name == "COUNTBLANK":
                    count_blank = sum(
                        1 for v in values if v is None or str(v).strip() == ""
                    )
                    return str(count_blank)

                elif func_name == "SUM":
                    nums = [float(v) for v in values if cls._is_numeric(v)]
                    return str(sum(nums))

                elif func_name in ("AVERAGE", "AVERAGEA"):
                    nums = [float(v) for v in values if cls._is_numeric(v)]
                    return str(sum(nums) / len(nums)) if nums else "0"

                elif func_name == "MIN":
                    nums = [float(v) for v in values if cls._is_numeric(v)]
                    return str(min(nums)) if nums else "0"

                elif func_name == "MAX":
                    nums = [float(v) for v in values if cls._is_numeric(v)]
                    return str(max(nums)) if nums else "0"

                elif func_name == "PRODUCT":
                    nums = [float(v) for v in values if cls._is_numeric(v)]
                    res = 1.0
                    for n in nums:
                        res *= n
                    return str(res if nums else 0)

                elif func_name == "SUBTOTAL":
                    parts = [p.strip() for p in args_str.split(",", 1)]
                    if len(parts) == 2:
                        sub_func = int(float(cls._resolve_val(parts[0], ws, cache) or 9))
                        sub_vals = cls._extract_values_from_args(parts[1], ws, cache)
                        if sub_func in (3, 103):  # COUNTA
                            return str(sum(1 for v in sub_vals if v is not None and str(v).strip() != ""))
                        elif sub_func in (2, 102):  # COUNT
                            return str(sum(1 for v in sub_vals if cls._is_numeric(v)))
                        else:  # SUM (9, 109)
                            nums = [float(v) for v in sub_vals if cls._is_numeric(v)]
                            return str(sum(nums))
                    return "0"

                elif func_name in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
                    parts = [p.strip() for p in args_str.split(",")]
                    v1 = cls._resolve_val(parts[0], ws, cache)
                    num_val = float(v1) if cls._is_numeric(v1) else 0.0
                    v2 = cls._resolve_val(parts[1], ws, cache) if len(parts) > 1 else 0
                    digits = int(float(v2)) if cls._is_numeric(v2) else 0

                    if func_name == "ROUND":
                        return str(round(num_val, digits))
                    elif func_name == "ROUNDUP":
                        factor = 10 ** digits
                        return str(math.ceil(num_val * factor) / factor)
                    elif func_name == "ROUNDDOWN":
                        factor = 10 ** digits
                        return str(math.floor(num_val * factor) / factor)

                elif func_name == "ABS":
                    v = cls._resolve_val(values[0] if values else args_str, ws, cache)
                    num_val = float(v) if cls._is_numeric(v) else 0.0
                    return str(abs(num_val))

                elif func_name == "INT":
                    v = cls._resolve_val(values[0] if values else args_str, ws, cache)
                    num_val = float(v) if cls._is_numeric(v) else 0.0
                    return str(int(math.floor(num_val)))

                elif func_name in ("CONCAT", "CONCATENATE"):
                    return '"' + "".join(str(v or "") for v in values) + '"'

                return "0"

            expr = FUNC_CALL_REGEX.sub(func_replacer, expr)

        # Step 2: Replace individual cell references with their current/cached numeric values
        def cell_replacer(match: re.Match) -> str:
            col_str = match.group(1).replace("$", "")
            row_num = int(match.group(3))
            col_num = column_index_from_string(col_str)

            val = cache.get((row_num, col_num))
            if val is None:
                val = ws.cell(row=row_num, column=col_num).value

            if val is None or val == "":
                return "0"
            try:
                return str(float(val))
            except (ValueError, TypeError):
                return f'"{str(val)}"'

        expr = CELL_COORD_REGEX.sub(cell_replacer, expr)

        # Step 3: Safely evaluate math expression
        try:
            cleaned_math = re.sub(r"[^0-9\+\-\*\/\%\(\)\.]", "", expr)
            if cleaned_math == expr:
                res = eval(cleaned_math, {"__builtins__": None}, {})
                if isinstance(res, float) and res.is_integer():
                    return int(res)
                return round(res, 4) if isinstance(res, float) else res
            
            clean_str = expr.strip('"\'')
            if cls._is_numeric(clean_str):
                val = float(clean_str)
                return int(val) if val.is_integer() else val
            return clean_str
        except Exception:
            pass

        return formula

    @classmethod
    def _extract_values_from_args(
        cls, 
        args_str: str, 
        ws: Worksheet, 
        cache: Dict[Tuple[int, int], Any]
    ) -> List[Any]:
        """Extracts all cell values from comma-separated ranges or single cell references."""
        values = []
        parts = [p.strip() for p in args_str.split(",") if p.strip()]

        for p in parts:
            # Range like A9:A11
            m_range = RANGE_REGEX.match(p)
            if m_range:
                start_ref = m_range.group(1).replace("$", "")
                end_ref = m_range.group(2).replace("$", "")

                start_m = CELL_COORD_REGEX.match(start_ref)
                end_m = CELL_COORD_REGEX.match(end_ref)
                if start_m and end_m:
                    s_col = column_index_from_string(start_m.group(1).replace("$", ""))
                    s_row = int(start_m.group(3))
                    e_col = column_index_from_string(end_m.group(1).replace("$", ""))
                    e_row = int(end_m.group(3))

                    for r in range(min(s_row, e_row), max(s_row, e_row) + 1):
                        for c in range(min(s_col, e_col), max(s_col, e_col) + 1):
                            val = cache.get((r, c))
                            if val is None:
                                val = ws.cell(row=r, column=c).value
                            values.append(val)
                continue

            # Single cell coordinate like A9
            m_cell = CELL_COORD_REGEX.match(p)
            if m_cell:
                col_str = m_cell.group(1).replace("$", "")
                r_num = int(m_cell.group(3))
                c_num = column_index_from_string(col_str)
                val = cache.get((r_num, c_num))
                if val is None:
                    val = ws.cell(row=r_num, column=c_num).value
                values.append(val)
                continue

            # Literal value
            p_clean = p.strip('"\'')
            if cls._is_numeric(p_clean):
                values.append(float(p_clean))
            else:
                values.append(p_clean)

        return values

    @staticmethod
    def _is_numeric(val: Any) -> bool:
        if val is None or val == "":
            return False
        try:
            float(val)
            return True
        except (ValueError, TypeError):
            return False

    @classmethod
    def format_cell_value(cls, val: Any, number_format: Optional[str]) -> str:
        """Formats a numeric value using Excel's number format mask."""
        if val is None:
            return ""

        if not isinstance(val, (int, float)):
            try:
                f_val = float(val)
                if f_val.is_integer():
                    val = int(f_val)
                else:
                    val = f_val
            except (ValueError, TypeError):
                return str(val)

        if not number_format or number_format.lower() in ("general", ""):
            if isinstance(val, float):
                return f"{val:,.2f}" if abs(val) >= 1000 else f"{val:.2f}"
            return f"{val:,}" if abs(val) >= 1000 else str(val)

        nf = number_format

        # Currency
        if "$" in nf:
            if ".00" in nf:
                return f"${val:,.2f}"
            return f"${int(round(val)):,}"

        # Percentage
        if "%" in nf:
            pct = val * 100
            if ".00" in nf:
                return f"{pct:.2f}%"
            return f"{round(pct)}%"

        # Thousands separator
        if "#,##0" in nf:
            if ".00" in nf:
                return f"{val:,.2f}"
            return f"{int(round(val)):,}"

        if "0.00" in nf:
            return f"{val:.2f}"

        return str(val)
