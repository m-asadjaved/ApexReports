import io
import base64
import re
from typing import Optional, Tuple
import httpx
from PIL import Image as PILImage
import qrcode
import barcode
from barcode.writer import ImageWriter
from openpyxl.drawing.image import Image as OpenPyXlImage
from openpyxl.worksheet.worksheet import Worksheet

class MediaHandler:
    """Handles generating and embedding dynamic QR codes, barcodes, and images into openpyxl worksheets."""

    @staticmethod
    def generate_qr_code(data: str, box_size: int = 4, border: int = 1) -> io.BytesIO:
        """Generates a QR Code image as an in-memory PNG buffer."""
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=box_size,
            border=border,
        )
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf

    @staticmethod
    def generate_barcode(data: str, code_format: str = "code128") -> io.BytesIO:
        """Generates a 1D barcode (Code128, EAN13, Code39) as an in-memory PNG buffer."""
        code_format_clean = code_format.lower().replace("-", "")
        writer = ImageWriter()
        try:
            barcode_cls = barcode.get_barcode_class(code_format_clean)
        except barcode.errors.BarcodeNotFoundError:
            barcode_cls = barcode.get_barcode_class("code128")
        
        buf = io.BytesIO()
        try:
            bc = barcode_cls(str(data), writer=writer)
            bc.write(buf, options={"write_text": False, "quiet_zone": 1.0})
        except Exception:
            bc = barcode.get_barcode_class("code128")(str(data), writer=writer)
            bc.write(buf, options={"write_text": False, "quiet_zone": 1.0})
            
        buf.seek(0)
        return buf

    @classmethod
    def load_image_bytes(cls, source: str) -> Optional[io.BytesIO]:
        """Loads image bytes from either a base64 string, HTTP URL, or local path."""
        if not source:
            return None
            
        source = str(source).strip()
        
        # Base64 data URI or raw base64
        if source.startswith("data:image/") or ";base64," in source:
            base64_data = source.split(";base64,")[-1]
            return io.BytesIO(base64.b64decode(base64_data))
        
        # HTTP / HTTPS URL
        if source.startswith("http://") or source.startswith("https://"):
            try:
                resp = httpx.get(source, timeout=10.0, follow_redirects=True)
                if resp.status_code == 200:
                    return io.BytesIO(resp.content)
            except Exception:
                return None

        # Check if it looks like raw base64 (long alphanumeric string)
        if len(source) > 100 and not any(c in source for c in ["\n", "\r", " "]):
            try:
                decoded = base64.b64decode(source)
                return io.BytesIO(decoded)
            except Exception:
                pass

        return None

    @classmethod
    def embed_image_in_cell(
        cls, 
        ws: Worksheet, 
        cell_coord: str, 
        img_buffer: io.BytesIO, 
        width: Optional[int] = None, 
        height: Optional[int] = None
    ) -> None:
        """Places an openpyxl Image into the worksheet anchored at cell_coord."""
        img = OpenPyXlImage(img_buffer)
        if width and height:
            img.width = width
            img.height = height
        elif width and not height:
            aspect = img.height / max(img.width, 1)
            img.width = width
            img.height = int(width * aspect)
        elif height and not width:
            aspect = img.width / max(img.height, 1)
            img.height = height
            img.width = int(height * aspect)
            
        ws.add_image(img, cell_coord)
