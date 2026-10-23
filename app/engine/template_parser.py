import re
from typing import Any, Dict, List, Optional, Tuple, Union
import jinja2

TAG_REGEX = re.compile(r"\{\{([^{}]+)\}\}")
TRAILING_ATTR_REGEX = re.compile(r'\s+([a-zA-Z_]\w*)=(?:"([^"]*)"|\'([^\']*)\'|([^\s"\']+))\s*$')

class ExpressionEvaluator:
    """Safe expression evaluator using Jinja2 with automatic type coercion for string concatenation."""
    def __init__(self):
        self.env = jinja2.Environment(
            variable_start_string="{{",
            variable_end_string="}}",
            autoescape=False,
            undefined=jinja2.Undefined
        )

    def evaluate_expression(self, expr: str, context: Dict[str, Any]) -> Any:
        """Evaluates an expression like 'item.price * item.qty', 'user.name', or \"'https://...' + id\" against context."""
        expr = expr.strip()
        if not expr:
            return ""

        # 1. Handle embedded interpolation: e.g. "https://verify.org/id={{invoice_no}}"
        if "{{" in expr and "}}" in expr:
            def sub_token(m: re.Match) -> str:
                tok = m.group(1).strip()
                val = self.evaluate_expression(tok, context)
                return str(val) if val is not None else ""
            return TAG_REGEX.sub(sub_token, expr)

        # 2. Direct key lookup optimization
        if expr in context:
            return context[expr]

        # 2b. System variable resolution (@page, @pages, @date, @time, page_no, total_pages, etc.)
        sys_aliases = {
            "@page": ["@page", "page_no", "_page"],
            "@pages": ["@pages", "total_pages", "_pages"],
            "@date": ["@date", "current_date", "_date"],
            "@time": ["@time", "current_time", "_time"],
            "@timestamp": ["@timestamp", "current_timestamp"],
            "page_no": ["@page", "page_no", "_page"],
            "total_pages": ["@pages", "total_pages", "_pages"],
            "current_date": ["@date", "current_date", "_date"],
            "current_time": ["@time", "current_time", "_time"],
        }
        if expr in sys_aliases:
            for alias in sys_aliases[expr]:
                if alias in context:
                    return context[alias]
            import datetime
            if expr in ("@date", "current_date", "_date"):
                return datetime.date.today().strftime("%Y-%m-%d")
            if expr in ("@time", "current_time", "_time"):
                return datetime.datetime.now().strftime("%H:%M:%S")
            if expr in ("@timestamp", "current_timestamp"):
                return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if expr in ("@page", "page_no", "_page"):
                return 1
            if expr in ("@pages", "total_pages", "_pages"):
                return 1

        # Direct @ identifier lookup (e.g. @index, @first, @last)
        if expr.startswith("@") and expr in context:
            return context[expr]

        # 3. Handle simple dotted access like a.b.c
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

        # 4. If expression involves string concatenation (+ or & near quotes), convert to Jinja2's ~ operator
        # This prevents TypeError when concatenating a string with an integer/float like "'url/' + 1001"
        jinja_expr = expr
        if ("'" in expr or '"' in expr) and ("+" in expr or "&" in expr):
            jinja_expr = re.sub(r"(['\"][^'\"]*['\"])\s*[\+&]", r"\1 ~", jinja_expr)
            jinja_expr = re.sub(r"[\+&]\s*(['\"][^'\"]*['\"])", r"~ \1", jinja_expr)

        # 5. Evaluate via Jinja2 template engine
        try:
            template = self.env.from_string(f"{{{{ {jinja_expr} }}}}")
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
