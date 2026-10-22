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
    def generate_qr_code(data: str, width: Optional[int] = None, height: Optional[int] = None) -> io.BytesIO:
        """Generates a QR Code image as an in-memory PNG buffer with specified dimensions."""
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=10,
            border=1,
        )
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

        if width and height:
            img = img.resize((int(width), int(height)), PILImage.Resampling.LANCZOS)
        elif width:
            aspect = img.height / max(img.width, 1)
            img = img.resize((int(width), int(width * aspect)), PILImage.Resampling.LANCZOS)
        elif height:
            aspect = img.width / max(img.height, 1)
            img = img.resize((int(height * aspect), int(height)), PILImage.Resampling.LANCZOS)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        buf.seek(0)
        return buf

    @staticmethod
    def generate_barcode(
        data: str, 
        code_format: str = "code128", 
        width: Optional[int] = None, 
        height: Optional[int] = None
    ) -> io.BytesIO:
        """Generates a 1D barcode (Code128, EAN13, Code39) resized to requested dimensions."""
        code_format_clean = code_format.lower().replace("-", "")
        writer = ImageWriter()
        try:
            barcode_cls = barcode.get_barcode_class(code_format_clean)
        except barcode.errors.BarcodeNotFoundError:
            barcode_cls = barcode.get_barcode_class("code128")

        temp_buf = io.BytesIO()
        try:
            bc = barcode_cls(str(data), writer=writer)
            bc.write(temp_buf, options={"write_text": False, "quiet_zone": 1.0})
        except Exception:
            bc = barcode.get_barcode_class("code128")(str(data), writer=writer)
            bc.write(temp_buf, options={"write_text": False, "quiet_zone": 1.0})

        temp_buf.seek(0)

        # Resize barcode image to requested dimensions
        pil_img = PILImage.open(temp_buf)
        if width and height:
            pil_img = pil_img.resize((int(width), int(height)), PILImage.Resampling.LANCZOS)
        elif width:
            aspect = pil_img.height / max(pil_img.width, 1)
            pil_img = pil_img.resize((int(width), int(width * aspect)), PILImage.Resampling.LANCZOS)
        elif height:
            aspect = pil_img.width / max(pil_img.height, 1)
            pil_img = pil_img.resize((int(height * aspect), int(height)), PILImage.Resampling.LANCZOS)

        out_buf = io.BytesIO()
        pil_img.save(out_buf, format="PNG")
        out_buf.seek(0)
        return out_buf

    @classmethod
    def load_image_bytes(cls, source: str) -> Optional[io.BytesIO]:
        """Loads image bytes from either a base64 string, HTTP URL, or local path."""
        if not source:
            return None

        source = str(source).strip()

        # Base64 data URI
        if source.startswith("data:image/") or ";base64," in source:
            base64_data = source.split(";base64,")[-1]
            return io.BytesIO(base64.b64decode(base64_data))

        # HTTP URL
        if source.startswith("http://") or source.startswith("https://"):
            try:
                resp = httpx.get(source, timeout=10.0, follow_redirects=True)
                if resp.status_code == 200:
                    return io.BytesIO(resp.content)
            except Exception:
                return None

        # Raw base64 string
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
        """Places an openpyxl Image into the worksheet anchored at cell_coord with specified dimensions."""
        img = OpenPyXlImage(img_buffer)
        if width:
            img.width = int(width)
        if height:
            img.height = int(height)

        ws.add_image(img, cell_coord)
