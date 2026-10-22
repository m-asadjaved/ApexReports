from openpyxl import Workbook
from app.renderer.libreoffice_pool import LibreOfficeRenderer
from app.renderer.browser_pdf_renderer import BrowserPdfRenderer

class PdfService:
    """Unified PDF generation service with automatic fallback."""

    @classmethod
    def get_active_backend(cls) -> str:
        if LibreOfficeRenderer.is_available():
            return "LibreOffice"
        elif BrowserPdfRenderer.is_available():
            browser = BrowserPdfRenderer.get_browser_executable()
            return f"Headless Browser ({browser})"
        return "None"

    @classmethod
    def convert_to_pdf(cls, wb: Workbook, filename: str = "report.pdf") -> bytes:
        # Prefer LibreOffice if available
        if LibreOfficeRenderer.is_available():
            return LibreOfficeRenderer.convert_workbook_to_pdf(wb)

        # Fallback to headless Chrome / Edge
        if BrowserPdfRenderer.is_available():
            return BrowserPdfRenderer.convert_worksheet_to_pdf(wb.active, title=filename)

        raise RuntimeError(
            "No PDF conversion engine found. Please install LibreOffice or ensure Google Chrome / Microsoft Edge is installed."
        )
