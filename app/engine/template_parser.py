import re
from typing import Any, Dict, List, Optional, Tuple, Union
import jinja2

TAG_REGEX = re.compile(r"\{\{([^{}]+)\}\}")

class ExpressionEvaluator:
    """Safe expression evaluator using Jinja2."""
    def __init__(self):
        self.env = jinja2.Environment(
            variable_start_string="{{",
            variable_end_string="}}",
            autoescape=False,
            undefined=jinja2.Undefined
        )

    def evaluate_expression(self, expr: str, context: Dict[str, Any]) -> Any:
        """Evaluates an expression like 'item.price * item.qty' or 'user.name' against context."""
        expr = expr.strip()
        if not expr:
            return ""

        # Direct key lookup optimization
        if expr in context:
            return context[expr]

        # Handle simple dotted access like a.b.c
        if re.match(r"^[a-zA-Z_][a-zA-Z0-9_]*(\.[a-zA-Z_][a-zA-Z0-9_]*)*$", expr):
            curr = context
            parts = expr.split(".")
            for part in parts:
                if isinstance(curr, dict):
                    curr = curr.get(part)
                elif hasattr(curr, part):
                    curr = getattr(curr, part)
                else:
                    return None
            return curr

        # Fallback to Jinja2 template rendering for complex expressions
        try:
            template = self.env.from_string(f"{{{{ {expr} }}}}")
            result = template.render(context)
            if result == "True":
                return True
            if result == "False":
                return False
            try:
                if "." in result:
                    return float(result)
                return int(result)
            except ValueError:
                return result
        except Exception:
            return ""

    def evaluate_condition(self, condition: str, context: Dict[str, Any]) -> bool:
        """Evaluates a conditional expression like 'discount > 0' or 'status == \"APPROVED\"'."""
        try:
            template = self.env.from_string(f"{{% if {condition} %}}TRUE{{% else %}}FALSE{{% endif %}}")
            rendered = template.render(context).strip()
            return rendered == "TRUE"
        except Exception:
            return False

def parse_tag_attributes(attr_string: str) -> Dict[str, Any]:
    """Parses key=value attributes from directives."""
    attrs: Dict[str, Any] = {}
    pattern = re.compile(r'(\b\w+)=(?:"([^"]*)"|\'([^\']*)\'|(\S+))')
    for match in pattern.finditer(attr_string):
        key = match.group(1)
        val = match.group(2) or match.group(3) or match.group(4)
        if val.lower() == "true":
            attrs[key] = True
        elif val.lower() == "false":
            attrs[key] = False
        else:
            try:
                if "." in val:
                    attrs[key] = float(val)
                else:
                    attrs[key] = int(val)
            except ValueError:
                attrs[key] = val
    return attrs

TRAILING_ATTR_REGEX = re.compile(r'\s+([a-zA-Z_]\w*)=(?:"([^"]*)"|\'([^\']*)\'|([^\s"\']+))\s*$')

def parse_media_tag(content: str) -> Tuple[str, Dict[str, Any]]:
    """Extracts expression and trailing key=value attributes from media directives.
    e.g. "'https://test.com/id=' + id width=200 height=100 format=code128"
    Returns: ("'https://test.com/id=' + id", {"width": 200, "height": 100, "format": "code128"})
    """
    text = content.strip()
    attrs: Dict[str, Any] = {}

    while True:
        m = TRAILING_ATTR_REGEX.search(text)
        if not m:
            break
        key = m.group(1)
        val = m.group(2) or m.group(3) or m.group(4)
        if val.lower() == "true":
            attrs[key] = True
        elif val.lower() == "false":
            attrs[key] = False
        else:
            try:
                attrs[key] = float(val) if "." in val else int(val)
            except ValueError:
                attrs[key] = val

        text = text[:m.start()].strip()

    return text, attrs
