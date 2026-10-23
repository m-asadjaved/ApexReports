import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet


class TemplateLinter:
    """
    Analyzes an Excel workbook (.xlsx) and an optional JSON data payload
    for syntax errors, unbalanced blocks, merged-cell collision hazards,
    formula calculation traps, and data contract mismatches.
    """

    DIRECTIVE_REGEX = re.compile(r'\{\{\s*(.*?)\s*\}\}')
    # Detects potential single-brace typos or broken syntax
    SINGLE_BRACE_TYPO = re.compile(r'(?<!\{)\{([#/@]?[a-zA-Z0-9_\-\.]+)\}(?!\})')

    @classmethod
    def lint_file(cls, file_path: Path, template_code: Optional[str] = None, data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Lints an Excel template file directly from disk."""
        try:
            wb = openpyxl.load_workbook(str(file_path), data_only=False)
        except Exception as e:
            return {
                "valid": False,
                "template_code": template_code or file_path.stem,
                "score": 0,
                "errors": [{"code": "FILE_CORRUPT", "message": f"Cannot load Excel file: {str(e)}", "cell": None, "row": None}],
                "warnings": [],
                "info": [],
                "summary": f"Failed to load workbook: {str(e)}"
            }
        return cls.lint_workbook(wb, template_code=template_code or file_path.stem, data=data)

    @classmethod
    def lint_workbook(cls, wb: openpyxl.Workbook, template_code: str, data: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """Deeply inspects workbook structure and validates rules."""
        ws: Worksheet = wb.active
        errors: List[Dict[str, Any]] = []
        warnings: List[Dict[str, Any]] = []
        info: List[Dict[str, Any]] = []

        block_stack: List[Dict[str, Any]] = []
        loop_ranges: List[Dict[str, Any]] = []
        discovered_vars: Set[Tuple[Tuple[str, ...], str]] = set()
        required_arrays: Set[str] = set()

        # 1. Scan cell by cell
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                if cell.value is None:
                    continue

                val_str = str(cell.value).strip()

                # Check for single brace syntax typos like {name} instead of {{name}}
                typos = cls.SINGLE_BRACE_TYPO.findall(val_str)
                for t in typos:
                    warnings.append({
                        "code": "SINGLE_BRACE_TYPO",
                        "message": f"Cell contains '{{{t}}}' which looks like an unescaped tag. Did you mean '{{{{{t}}}}}'?",
                        "cell": cell.coordinate,
                        "row": cell.row
                    })

                # Check formulas
                if val_str.startswith("="):
                    cls._check_formula(cell, val_str, warnings)
                    continue

                # Check template directives
                if "{{" in val_str:
                    tags = cls.DIRECTIVE_REGEX.findall(val_str)
                    if not tags:
                        warnings.append({
                            "code": "MALFORMED_TAG",
                            "message": f"Cell contains '{{' but no valid tag could be parsed: '{val_str}'",
                            "cell": cell.coordinate,
                            "row": cell.row
                        })
                        continue

                    for tag_content in tags:
                        tag = tag_content.strip()
                        cls._process_tag(tag, cell, block_stack, loop_ranges, discovered_vars, required_arrays, errors, warnings)

        # 2. Check for unclosed blocks remaining on the stack
        while block_stack:
            unclosed = block_stack.pop()
            errors.append({
                "code": "UNCLOSED_BLOCK",
                "message": f"Block '{{{{#{unclosed['kind']} {unclosed['target']}}}}}' opened at {unclosed['cell']} is never closed with '{{{{/{unclosed['kind']}}}}}'.",
                "cell": unclosed["cell"],
                "row": unclosed["row"]
            })

        # 3. Check for Merged Cell Boundary Hazards
        cls._check_merged_cells(ws, loop_ranges, warnings, errors)

        # 4. Data Contract Alignment (if sample data provided)
        if data is not None:
            cls._check_data_contract(data, required_arrays, discovered_vars, errors, warnings, info)
        else:
            info.append({
                "code": "NO_PAYLOAD_TESTED",
                "message": "Template structural linting passed. Provide a JSON payload to validate field contracts."
            })

        # 5. Compute Health Quality Score (0 - 100)
        score = 100 - (len(errors) * 30) - (len(warnings) * 10)
        score = max(0, min(100, score))
        is_valid = len(errors) == 0

        summary = f"Template is healthy (Score: {score}/100) with 0 errors and {len(warnings)} warning(s)." if is_valid \
            else f"Validation failed with {len(errors)} error(s) and {len(warnings)} warning(s)."

        return {
            "valid": is_valid,
            "template_code": template_code,
            "score": score,
            "errors": errors,
            "warnings": warnings,
            "info": info,
            "summary": summary
        }

    @classmethod
    def _process_tag(
        cls,
        tag: str,
        cell: Any,
        block_stack: List[Dict[str, Any]],
        loop_ranges: List[Dict[str, Any]],
        discovered_vars: Set[Tuple[Tuple[str, ...], str]],
        required_arrays: Set[str],
        errors: List[Dict[str, Any]],
        warnings: List[Dict[str, Any]]
    ):
        # Open Loop
        if tag.startswith("#each "):
            parts = tag.split()
            target = parts[1]
            required_arrays.add(target)
            block_stack.append({"kind": "each", "target": target, "cell": cell.coordinate, "row": cell.row})

        # Close Loop
        elif tag == "/each":
            if not block_stack:
                errors.append({
                    "code": "UNMATCHED_CLOSE_EACH",
                    "message": f"Unexpected '{{{{/each}}}}' at {cell.coordinate} without a matching '{{{{#each}}}}'.",
                    "cell": cell.coordinate,
                    "row": cell.row
                })
            else:
                top = block_stack.pop()
                if top["kind"] != "each":
                    errors.append({
                        "code": "MISMATCHED_BLOCK_CLOSE",
                        "message": f"Mismatched closing tag '{{{{/each}}}}' at {cell.coordinate}. Expected '{{{{/{top['kind']}}}}}' for block opened at {top['cell']}.",
                        "cell": cell.coordinate,
                        "row": cell.row
                    })
                else:
                    loop_ranges.append({
                        "target": top["target"],
                        "start_row": top["row"],
                        "end_row": cell.row
                    })

        # Conditionals
        elif tag.startswith("#if "):
            cond = tag[4:].strip()
            block_stack.append({"kind": "if", "target": cond, "cell": cell.coordinate, "row": cell.row})

        elif tag == "/if":
            if not block_stack:
                errors.append({
                    "code": "UNMATCHED_CLOSE_IF",
                    "message": f"Unexpected '{{{{/if}}}}' at {cell.coordinate} without a matching '{{{{#if}}}}'.",
                    "cell": cell.coordinate,
                    "row": cell.row
                })
            else:
                top = block_stack.pop()
                if top["kind"] != "if":
                    errors.append({
                        "code": "MISMATCHED_BLOCK_CLOSE",
                        "message": f"Mismatched closing tag '{{{{/if}}}}' at {cell.coordinate}. Expected '{{{{/{top['kind']}}}}}'.",
                        "cell": cell.coordinate,
                        "row": cell.row
                    })

        # Page Headers
        elif tag in ("#page-header", "#header"):
            block_stack.append({"kind": "header", "target": tag, "cell": cell.coordinate, "row": cell.row})

        elif tag in ("/page-header", "/header"):
            if not block_stack:
                errors.append({
                    "code": "UNMATCHED_CLOSE_HEADER",
                    "message": f"Unexpected '{{{{{tag}}}}}' at {cell.coordinate} without a matching header open tag.",
                    "cell": cell.coordinate,
                    "row": cell.row
                })
            else:
                top = block_stack.pop()
                if top["kind"] != "header":
                    errors.append({
                        "code": "MISMATCHED_BLOCK_CLOSE",
                        "message": f"Mismatched closing tag '{{{{{tag}}}}}' at {cell.coordinate}. Expected '{{{{/{top['kind']}}}}}'.",
                        "cell": cell.coordinate,
                        "row": cell.row
                    })

        # Page Footers
        elif tag in ("#page-footer", "#footer"):
            block_stack.append({"kind": "footer", "target": tag, "cell": cell.coordinate, "row": cell.row})

        elif tag in ("/page-footer", "/footer"):
            if not block_stack:
                errors.append({
                    "code": "UNMATCHED_CLOSE_FOOTER",
                    "message": f"Unexpected '{{{{{tag}}}}}' at {cell.coordinate} without a matching footer open tag.",
                    "cell": cell.coordinate,
                    "row": cell.row
                })
            else:
                top = block_stack.pop()
                if top["kind"] != "footer":
                    errors.append({
                        "code": "MISMATCHED_BLOCK_CLOSE",
                        "message": f"Mismatched closing tag '{{{{{tag}}}}}' at {cell.coordinate}. Expected '{{{{/{top['kind']}}}}}'.",
                        "cell": cell.coordinate,
                        "row": cell.row
                    })

        # Media Tags (@qr, @barcode, @image)
        elif tag.startswith("@qr") or tag.startswith("@barcode") or tag.startswith("@image"):
            parts = tag.split()
            mtype = parts[0][1:]
            for p in parts[1:]:
                if "=" in p:
                    k, v = p.split("=", 1)
                    if k in ["width", "height"]:
                        try:
                            val_num = int(v)
                            if val_num <= 0 or val_num > 1000:
                                warnings.append({
                                    "code": "SUSPICIOUS_MEDIA_DIMENSION",
                                    "message": f"Media tag '{tag}' specifies {k}={val_num} which is outside recommended range (10-1000px).",
                                    "cell": cell.coordinate,
                                    "row": cell.row
                                })
                        except ValueError:
                            errors.append({
                                "code": "INVALID_MEDIA_PARAM",
                                "message": f"Media dimension '{k}' must be an integer: '{p}' in cell {cell.coordinate}.",
                                "cell": cell.coordinate,
                                "row": cell.row
                            })

        # Matrix columns & Page Breaks
        elif tag.startswith("#each-col ") or tag == "#page-break":
            pass

        # Scalar Variable
        else:
            clean_name = tag.split()[0]
            current_scope = tuple(b["target"] for b in block_stack if b["kind"] == "each")
            discovered_vars.add((current_scope, clean_name))

    @classmethod
    def _check_formula(cls, cell: Any, val_str: str, warnings: List[Dict[str, Any]]):
        """Checks formula safety and validity."""
        val_upper = val_str.upper()
        if "(" in val_upper and not val_upper.endswith(")"):
            warnings.append({
                "code": "UNBALANCED_FORMULA_PAREN",
                "message": f"Formula '{val_str}' at {cell.coordinate} may be missing a closing parenthesis.",
                "cell": cell.coordinate,
                "row": cell.row
            })

    @classmethod
    def _check_merged_cells(cls, ws: Worksheet, loop_ranges: List[Dict[str, Any]], warnings: List[Dict[str, Any]], errors: List[Dict[str, Any]]):
        """
        Detects if any merged cell range spans across repeating loop boundaries.
        Example of hazard: Merging A2:D10 when row 5-8 is an inner #each loop!
        """
        for mrange in ws.merged_cells.ranges:
            min_r, min_c, max_r, max_c = mrange.min_row, mrange.min_col, mrange.max_row, mrange.max_col
            # Check if this merge spans multiple rows
            if min_r < max_r:
                for loop in loop_ranges:
                    l_start = loop["start_row"]
                    l_end = loop["end_row"]
                    # If merge starts outside the loop but ends inside, or starts inside and ends outside
                    is_crossing = (min_r < l_start and max_r >= l_start) or (min_r <= l_end and max_r > l_end)
                    if is_crossing:
                        warnings.append({
                            "code": "MERGED_CELL_LOOP_COLLISION",
                            "message": f"Merged range '{mrange.coord}' spans across repeating frame '{{{{#each {loop['target']}}}}}' (rows {l_start}-{l_end}). This can overwrite cloned data. Keep merged cells strictly horizontal on single rows!",
                            "cell": mrange.coord,
                            "row": min_r
                        })

    @classmethod
    def _check_data_contract(
        cls,
        data: Dict[str, Any],
        required_arrays: Set[str],
        discovered_vars: Set[Tuple[Tuple[str, ...], str]],
        errors: List[Dict[str, Any]],
        warnings: List[Dict[str, Any]],
        info: List[Dict[str, Any]]
    ):
        """Validates that the JSON data satisfies all required template arrays and variables."""
        if not isinstance(data, dict):
            errors.append({
                "code": "INVALID_ROOT_PAYLOAD",
                "message": f"Expected JSON payload root to be an object (dict), received {type(data).__name__}."
            })
            return

        # Check required arrays
        for arr_name in required_arrays:
            val = cls._find_key_recursive(data, arr_name)
            if val is None:
                warnings.append({
                    "code": "MISSING_ARRAY_DATA",
                    "message": f"Template contains repeating loop '{{{{#each {arr_name}}}}}', but key '{arr_name}' was not found in the JSON payload."
                })
            elif not isinstance(val, list):
                errors.append({
                    "code": "ARRAY_TYPE_MISMATCH",
                    "message": f"Loop '{{{{#each {arr_name}}}}}' expects an array (list), but received {type(val).__name__} in JSON."
                })

        # Check variables
        missing_count = 0
        found_count = 0
        for scope, var_name in discovered_vars:
            # Skip special directives or expressions
            if "'" in var_name or '"' in var_name or "+" in var_name:
                continue

            clean_key = var_name.split(".")[0]
            if clean_key.startswith("@") or clean_key in ("page_no", "total_pages", "current_date", "current_time", "current_timestamp"):
                found_count += 1
                continue

            val = cls._find_key_recursive(data, clean_key)
            if val is None:
                missing_count += 1
                warnings.append({
                    "code": "UNRESOLVED_VARIABLE",
                    "message": f"Variable '{{{{{var_name}}}}}' (scope: {scope or 'root'}) not found in provided JSON data."
                })
            else:
                found_count += 1

        if missing_count == 0:
            info.append({
                "code": "DATA_CONTRACT_SATISFIED",
                "message": f"All {found_count} template variables are satisfied by the JSON payload."
            })
        else:
            info.append({
                "code": "DATA_CONTRACT_PARTIAL",
                "message": f"{found_count} variables matched, {missing_count} variable(s) not found in JSON."
            })

    @classmethod
    def _find_key_recursive(cls, data: Any, target_key: str) -> Any:
        """Recursively checks if target_key exists in dictionary or nested dictionaries."""
        if isinstance(data, dict):
            if target_key in data:
                return data[target_key]
            for v in data.values():
                res = cls._find_key_recursive(v, target_key)
                if res is not None:
                    return res
        elif isinstance(data, list) and data:
            return cls._find_key_recursive(data[0], target_key)
        return None
