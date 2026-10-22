from copy import copy
from typing import Any, Dict
from openpyxl.cell.cell import Cell
from app.engine.template_parser import ExpressionEvaluator, TAG_REGEX

evaluator = ExpressionEvaluator()

def copy_cell_style(src_cell: Cell, dst_cell: Cell) -> None:
    """Clones font, border, fill, number_format, protection, and alignment from src to dst."""
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.border = copy(src_cell.border)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.number_format = copy(src_cell.number_format)
        dst_cell.protection = copy(src_cell.protection)
        dst_cell.alignment = copy(src_cell.alignment)

def render_cell_value(val: Any, context: Dict[str, Any]) -> Any:
    """Substitutes {{expressions}} inside a cell value."""
    if val is None or not isinstance(val, str):
        return val

    # If it's a formula, do not treat as tag
    if val.startswith("="):
        return val

    val_str = str(val).strip()

    # Exact single token match: e.g. '{{amount}}' -> preserves native int/float/date type
    m_exact = TAG_REGEX.fullmatch(val_str)
    if m_exact:
        tag_content = m_exact.group(1).strip()
        # Direct directives should not be rendered here
        if tag_content.startswith(("#", "/", "@")):
            return val
        evaluated = evaluator.evaluate_expression(tag_content, context)
        return evaluated if evaluated is not None else ""

    # Embedded tokens in text: e.g. 'Customer: {{cust.name}} (ID: {{cust.id}})'
    def token_replace(m):
        tag = m.group(1).strip()
        if tag.startswith(("#", "/", "@")):
            return m.group(0)
        res = evaluator.evaluate_expression(tag, context)
        return str(res) if res is not None else ""

    return TAG_REGEX.sub(token_replace, val)
