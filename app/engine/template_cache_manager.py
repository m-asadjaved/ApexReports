import os
import pickle
import logging
from pathlib import Path
from typing import Dict, Optional, Tuple, Any
from openpyxl import load_workbook

from app.core.config import settings
from app.engine.compiled_template import CompiledTemplateAST, TemplateCompiler

logger = logging.getLogger("apex_reports.cache")

class TemplateCacheManager:
    """
    Two-Tier Caching System for ApexReports:
      - Tier 1: In-Memory RAM Cache for ultra-fast (~0.05ms) access.
      - Tier 2: Persistent Disk Cache (.cache/templates/) to survive server restarts.
      - Automatic Invalidation: Compares .xlsx file mtime on every request.
    """

    _RAM_CACHE: Dict[str, Tuple[float, CompiledTemplateAST]] = {}
    _CACHE_DIR: Path = settings.BASE_DIR / ".cache" / "templates"

    @classmethod
    def get_or_compile(cls, template_code: str, template_path: Path) -> Tuple[CompiledTemplateAST, bool]:
        """
        Retrieves the compiled template AST from RAM or Disk cache, or compiles if missing/outdated.
        Returns: (ast, is_cache_hit)
        """
        try:
            current_mtime = template_path.stat().st_mtime
        except Exception:
            current_mtime = 0.0

        # Tier 1: In-Memory RAM Cache
        if template_code in cls._RAM_CACHE:
            cached_mtime, ast = cls._RAM_CACHE[template_code]
            if cached_mtime == current_mtime:
                return ast, True

        # Tier 2: Persistent Disk Cache
        cls._CACHE_DIR.mkdir(parents=True, exist_ok=True)
        cache_file = cls._CACHE_DIR / f"{template_code}.cache"
        if cache_file.is_file():
            try:
                with open(cache_file, "rb") as f:
                    ast: CompiledTemplateAST = pickle.load(f)
                if getattr(ast, "mtime", None) == current_mtime:
                    cls._RAM_CACHE[template_code] = (current_mtime, ast)
                    return ast, True
            except Exception as e:
                logger.warning(f"Failed to read disk cache for '{template_code}': {e}")

        # Cache Miss or File Modified: Compile from Excel
        wb = load_workbook(str(template_path), data_only=False)
        ast = TemplateCompiler.compile(wb, template_code, mtime=current_mtime)

        # Store in Tier 1 (RAM)
        cls._RAM_CACHE[template_code] = (current_mtime, ast)

        # Store in Tier 2 (Disk)
        try:
            with open(cache_file, "wb") as f:
                pickle.dump(ast, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception as e:
            logger.warning(f"Failed to write disk cache for '{template_code}': {e}")

        return ast, False

    @classmethod
    def invalidate(cls, template_code: Optional[str] = None) -> None:
        """Invalidates cache for a specific template or all templates."""
        if template_code:
            cls._RAM_CACHE.pop(template_code, None)
            cache_file = cls._CACHE_DIR / f"{template_code}.cache"
            if cache_file.is_file():
                try:
                    cache_file.unlink()
                except Exception:
                    pass
        else:
            cls._RAM_CACHE.clear()
            if cls._CACHE_DIR.is_dir():
                for f in cls._CACHE_DIR.glob("*.cache"):
                    try:
                        f.unlink()
                    except Exception:
                        pass

    @classmethod
    def get_stats(cls) -> Dict[str, Any]:
        """Diagnostic stats about the two-tier cache."""
        disk_files = list(cls._CACHE_DIR.glob("*.cache")) if cls._CACHE_DIR.is_dir() else []
        return {
            "ram_cached_count": len(cls._RAM_CACHE),
            "ram_templates": list(cls._RAM_CACHE.keys()),
            "disk_cached_count": len(disk_files),
            "disk_cache_dir": str(cls._CACHE_DIR),
        }
