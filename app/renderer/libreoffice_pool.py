import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional
from openpyxl import Workbook
from app.core.config import settings
from app.core.exceptions import RenderingException

class LibreOfficeRenderer:
    """Manages document conversion to PDF via headless LibreOffice Calc."""

    @classmethod
    def is_available(cls) -> bool:
        """Returns True if LibreOffice soffice binary is installed and detected."""
        return settings.get_soffice_executable() is not None

    @classmethod
    def convert_workbook_to_pdf(cls, wb: Workbook) -> bytes:
        """Saves an openpyxl Workbook to temporary storage, converts to PDF using LibreOffice, and returns PDF bytes."""
        soffice_bin = settings.get_soffice_executable()
        if not soffice_bin:
            raise RenderingException(
                "LibreOffice ('soffice') executable not found. "
                "For local testing without LibreOffice, please request format='html' or format='xlsx', "
                "or run the service via the provided Docker container where LibreOffice is pre-configured."
            )

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            xlsx_file = tmp_path / "document.xlsx"
            pdf_file = tmp_path / "document.pdf"

            # Save populated workbook to temporary file
            wb.save(str(xlsx_file))

            # Command: soffice --headless --convert-to pdf --outdir <tmpdir> <xlsx_file>
            cmd = [
                soffice_bin,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp_path),
                str(xlsx_file)
            ]

            try:
                proc = subprocess.run(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=settings.SOFFICE_TIMEOUT_SECONDS,
                    check=False
                )
            except subprocess.TimeoutExpired:
                raise RenderingException(f"LibreOffice conversion timed out after {settings.SOFFICE_TIMEOUT_SECONDS}s.")
            except Exception as e:
                raise RenderingException(f"Failed to launch LibreOffice process: {str(e)}")

            if not pdf_file.is_file():
                stderr_msg = proc.stderr.decode("utf-8", errors="ignore").strip()
                raise RenderingException(f"LibreOffice failed to generate PDF. Exit code {proc.returncode}. {stderr_msg}")

            return pdf_file.read_bytes()
