from typing import Optional
from openpyxl import Workbook
from app.renderer.libreoffice_pool import LibreOfficeRenderer
from app.renderer.browser_pdf_renderer import BrowserPdfRenderer
from app.renderer.chrome_cdp_pool import ChromeCdpPool
from app.renderer.fast_html_renderer import FastHtmlRenderer


class PdfService:
    """Unified PDF generation service with persistent Chrome CDP pool and automatic fallbacks."""

    @classmethod
    def get_active_backend(cls) -> str:
        if LibreOfficeRenderer.is_available():
            return "LibreOffice"
        elif ChromeCdpPool.is_available():
            browser = BrowserPdfRenderer.get_browser_executable()
            mode = "Persistent Chrome CDP Pool" if ChromeCdpPool.is_running() else "Chrome CDP Pool (On-Demand)"
            return f"{mode} ({browser})"
        elif BrowserPdfRenderer.is_available():
            browser = BrowserPdfRenderer.get_browser_executable()
            return f"Headless Browser CLI ({browser})"
        return "None"

    @classmethod
    def convert_to_pdf(cls, wb: Workbook, filename: str = "report.pdf", font_family: Optional[str] = None) -> bytes:
        # 1. Prefer LibreOffice container if available
        if LibreOfficeRenderer.is_available():
            return LibreOfficeRenderer.convert_workbook_to_pdf(wb)

        # 2. Prefer Persistent Chrome CDP Pool (~120ms latency)
        if ChromeCdpPool.is_available():
            html_str = FastHtmlRenderer.render(wb.active, title=filename, font_family=font_family)
            return ChromeCdpPool.render_html_to_pdf(html_str)

        # 3. Fallback to CLI browser renderer
        if BrowserPdfRenderer.is_available():
            html_str = FastHtmlRenderer.render(wb.active, title=filename, font_family=font_family)
            return BrowserPdfRenderer.convert_html_to_pdf(html_str)

        raise RuntimeError(
            "No PDF conversion engine found. Please install LibreOffice or ensure Google Chrome / Microsoft Edge is installed."
        )
