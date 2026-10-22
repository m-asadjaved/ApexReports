import re
import colorsys
from typing import Any, List, Optional
import xml.etree.ElementTree as ET
import openpyxl
import openpyxl.styles.colors as oc


class ExcelColorResolver:
    """
    Faithfully resolves Excel colors (Theme colors with tints, Indexed colors, and direct aRGB)
    into standard CSS hex strings (#RRGGBB).
    """

    # Standard Office 2013-2026 default palette (Theme 0 to 11)
    # [lt1, dk1, lt2, dk2, accent1, accent2, accent3, accent4, accent5, accent6, hlink, folHlink]
    DEFAULT_THEME_COLORS: List[str] = [
        "FFFFFF",  # 0: lt1 (Light 1 - White)
        "000000",  # 1: dk1 (Dark 1 - Black)
        "EEECE1",  # 2: lt2 (Light 2 - Cream/Gray)
        "1F497D",  # 3: dk2 (Dark 2 - Deep Navy)
        "4F81BD",  # 4: accent1 (Blue)
        "C0504D",  # 5: accent2 (Red/Coral)
        "9BBB59",  # 6: accent3 (Olive Green)
        "8064A2",  # 7: accent4 (Purple)
        "4BACC6",  # 8: accent5 (Aqua/Teal)
        "F79646",  # 9: accent6 (Orange)
        "0000FF",  # 10: hlink (Blue)
        "800080",  # 11: folHlink (Purple)
    ]

    @classmethod
    def extract_theme_colors(cls, wb: Optional[openpyxl.Workbook]) -> List[str]:
        """
        Parses custom theme colors from an openpyxl Workbook if loaded_theme is present,
        mapping OpenXML XML tag ordering to standard ECMA-376 theme indices.
        """
        if not wb or not getattr(wb, "loaded_theme", None):
            return list(cls.DEFAULT_THEME_COLORS)

        try:
            root = ET.fromstring(wb.loaded_theme)
            ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
            scheme = root.find(".//a:clrScheme", ns)
            if scheme is None:
                return list(cls.DEFAULT_THEME_COLORS)

            xml_colors: List[str] = []
            for child in scheme:
                srgb = child.find("a:srgbClr", ns)
                sys_clr = child.find("a:sysClr", ns)
                if srgb is not None and "val" in srgb.attrib:
                    xml_colors.append(srgb.attrib["val"])
                elif sys_clr is not None:
                    xml_colors.append(sys_clr.attrib.get("lastClr", "000000"))

            # ECMA-376 XML tag order: dk1, lt1, dk2, lt2, accent1...
            # Standard theme indices: 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4=accent1...
            if len(xml_colors) >= 12:
                return [
                    xml_colors[1],  # 0: lt1
                    xml_colors[0],  # 1: dk1
                    xml_colors[3],  # 2: lt2
                    xml_colors[2],  # 3: dk2
                    *xml_colors[4:12]
                ]
        except Exception:
            pass

        return list(cls.DEFAULT_THEME_COLORS)

    @classmethod
    def apply_tint(cls, hex_rgb: str, tint: float) -> str:
        """Applies Microsoft Office HLS luminance tint to an RGB hex string."""
        if not tint or tint == 0.0:
            return hex_rgb

        clean_hex = hex_rgb.strip().lstrip("#")
        if len(clean_hex) != 6:
            return hex_rgb

        try:
            r = int(clean_hex[0:2], 16) / 255.0
            g = int(clean_hex[2:4], 16) / 255.0
            b = int(clean_hex[4:6], 16) / 255.0

            # Convert to HLS (Hue, Luminance, Saturation)
            h, l, s = colorsys.rgb_to_hls(r, g, b)

            # Excel HLS tint algorithm (ECMA-376 Part 1)
            if tint < 0:
                l = l * (1.0 + tint)
            else:
                l = l * (1.0 - tint) + tint

            l = max(0.0, min(1.0, l))
            r_new, g_new, b_new = colorsys.hls_to_rgb(h, l, s)

            r_int = max(0, min(255, int(round(r_new * 255))))
            g_int = max(0, min(255, int(round(g_new * 255))))
            b_int = max(0, min(255, int(round(b_new * 255))))

            return f"{r_int:02X}{g_int:02X}{b_int:02X}"
        except Exception:
            return hex_rgb

    @classmethod
    def resolve_color(
        cls,
        color_obj: Any,
        theme_colors: Optional[List[str]] = None,
        default: Optional[str] = None
    ) -> Optional[str]:
        """
        Resolves any openpyxl Color object (theme, indexed, rgb) into a '#RRGGBB' CSS hex string.
        """
        if not color_obj:
            return default

        # Handle raw string color values (e.g. "000000" or "#FF0000")
        if isinstance(color_obj, str):
            clean_str = color_obj.strip().lstrip("#")
            if re.match(r"^[0-9a-fA-F]{6,8}$", clean_str):
                return "#" + (clean_str[2:] if len(clean_str) == 8 else clean_str)
            return default

        themes = theme_colors or cls.DEFAULT_THEME_COLORS
        tint = getattr(color_obj, "tint", 0.0) or 0.0

        # 1. Theme Color
        if getattr(color_obj, "type", None) == "theme" or getattr(color_obj, "theme", None) is not None:
            theme_idx = getattr(color_obj, "theme", None)
            if isinstance(theme_idx, int) and 0 <= theme_idx < len(themes):
                base_hex = themes[theme_idx]
                tinted = cls.apply_tint(base_hex, tint)
                return f"#{tinted}"

        # 2. Indexed Color
        if getattr(color_obj, "type", None) == "indexed" or getattr(color_obj, "indexed", None) is not None:
            idx = getattr(color_obj, "indexed", None)
            if isinstance(idx, int) and 0 <= idx < len(oc.COLOR_INDEX):
                raw = str(oc.COLOR_INDEX[idx])
                clean_raw = raw[2:] if len(raw) == 8 else raw
                if re.match(r"^[0-9a-fA-F]{6}$", clean_raw):
                    tinted = cls.apply_tint(clean_raw, tint)
                    return f"#{tinted}"

        # 3. Direct RGB
        raw_rgb = getattr(color_obj, "rgb", None)
        if isinstance(raw_rgb, str) and re.match(r"^[0-9a-fA-F]{6,8}$", raw_rgb):
            clean_rgb = raw_rgb[2:] if len(raw_rgb) == 8 else raw_rgb
            tinted = cls.apply_tint(clean_rgb, tint)
            return f"#{tinted}"

        return default

    @classmethod
    def resolve_fill(cls, fill_obj: Any, theme_colors: Optional[List[str]] = None) -> Optional[str]:
        """Resolves background fill color for cell."""
        if not fill_obj:
            return None

        # Only solid fills have a definitive single background color in CSS
        fill_type = getattr(fill_obj, "fill_type", None)
        if fill_type not in ("solid",):
            return None

        color_source = getattr(fill_obj, "start_color", None) or getattr(fill_obj, "fgColor", None)
        if not color_source:
            return None

        return cls.resolve_color(color_source, theme_colors)

    @classmethod
    def resolve_font_color(cls, font_obj: Any, theme_colors: Optional[List[str]] = None) -> Optional[str]:
        """Resolves font text color."""
        if not font_obj or not getattr(font_obj, "color", None):
            return None

        return cls.resolve_color(font_obj.color, theme_colors)
