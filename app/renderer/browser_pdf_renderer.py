import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional
from openpyxl.worksheet.worksheet import Worksheet
from app.renderer.fast_html_renderer import FastHtmlRenderer
from app.core.exceptions import RenderingException

class BrowserPdfRenderer:
    """Generates high-resolution vector PDF using headless Chrome or Microsoft Edge on Windows/Linux."""

    @classmethod
    def get_browser_executable(cls) -> Optional[str]:
        """Detects Chrome, Chromium, or Microsoft Edge executable."""
        # 1. Check Chrome in standard PATH
        for binary in ["google-chrome", "chrome", "chromium", "chromium-browser", "msedge"]:
            found = shutil.which(binary)
            if found:
                return found

        # 2. Check Windows standard installations
        windows_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
        ]
        for p in windows_paths:
            if Path(p).is_file():
                return p

        return None

    @classmethod
    def is_available(cls) -> bool:
        return cls.get_browser_executable() is not None

    @classmethod
    def convert_html_to_pdf(cls, html_content: str) -> bytes:
        """Converts HTML string to vector PDF using headless browser."""
        browser_bin = cls.get_browser_executable()
        if not browser_bin:
            raise RenderingException(
                "Neither LibreOffice nor Chrome/Edge browser was found for PDF generation."
            )

        with tempfile.TemporaryDirectory() as tmpdir:
            tmp_path = Path(tmpdir)
            html_file = tmp_path / "page.html"
            pdf_file = tmp_path / "output.pdf"

            html_file.write_text(html_content, encoding="utf-8")

            cmd = [
                browser_bin,
                "--headless=new",
                "--disable-gpu",
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--no-pdf-header-footer",
                f"--print-to-pdf={str(pdf_file)}",
                str(html_file)
            ]

            try:
                proc = subprocess.run(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    timeout=20,
                    check=False
                )
            except subprocess.TimeoutExpired:
                raise RenderingException("Browser PDF generation timed out.")
            except Exception as e:
                raise RenderingException(f"Failed to launch browser: {str(e)}")

            if not pdf_file.is_file():
                err = proc.stderr.decode("utf-8", errors="ignore").strip()
                raise RenderingException(f"Browser failed to generate PDF. {err}")

            return pdf_file.read_bytes()

    @classmethod
    def convert_worksheet_to_pdf(cls, ws: Worksheet, title: str = "Report") -> bytes:
        """Renders worksheet to HTML and then converts to vector PDF."""
        html_str = FastHtmlRenderer.render(ws, title=title)
        return cls.convert_html_to_pdf(html_str)
