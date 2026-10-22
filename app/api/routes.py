import io
import time
import base64
from pathlib import Path
from typing import Optional
from fastapi import APIRouter, File, Header, HTTPException, Response, UploadFile, status
from fastapi.responses import HTMLResponse, FileResponse
from openpyxl import load_workbook

from app.core.config import settings
from app.core.exceptions import TemplateNotFoundException, RenderingException, TemplateSyntaxException
from app.api.schemas import RenderRequest, RenderResponseJson, TemplateInfo, TemplateListResponse
from app.engine.template_processor import TemplateProcessor
from app.engine.template_inspector import TemplateInspector
from app.renderer.fast_html_renderer import FastHtmlRenderer
from app.renderer.libreoffice_pool import LibreOfficeRenderer
from app.renderer.pdf_service import PdfService

router = APIRouter()

MIME_TYPES = {
    "pdf": "application/pdf",
    "html": "text/html; charset=utf-8",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
}

def verify_auth(authorization: Optional[str] = Header(None), x_api_key: Optional[str] = Header(None)) -> None:
    """Verifies optional API key if configured."""
    if not settings.API_KEY:
        return
    token = x_api_key or (authorization.replace("Bearer ", "") if authorization else None)
    if token != settings.API_KEY:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Invalid API Key.")

@router.get("/health")
def health_check():
    """Health check endpoint providing runtime and renderer diagnostic info."""
    return {
        "status": "online",
        "service": settings.PROJECT_NAME,
        "pdf_backend": PdfService.get_active_backend(),
        "libreoffice_available": LibreOfficeRenderer.is_available(),
        "templates_directory": str(settings.TEMPLATES_DIR),
    }

@router.get("/templates", response_model=TemplateListResponse)
def list_templates():
    """Lists all available report templates in the repository."""
    templates = []
    for f in settings.TEMPLATES_DIR.glob("*.xlsx"):
        stat = f.stat()
        templates.append(
            TemplateInfo(
                template_code=f.stem,
                filename=f.name,
                size_bytes=stat.st_size,
                modified_at=time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat.st_mtime))
            )
        )
    return TemplateListResponse(count=len(templates), templates=templates)

@router.post("/templates/upload")
async def upload_template(file: UploadFile = File(...)):
    """Uploads a new or updated Excel template file."""
    if not file.filename.endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported as templates.")

    dest = settings.TEMPLATES_DIR / file.filename
    content = await file.read()
    dest.write_bytes(content)

    return {
        "status": "success",
        "template_code": dest.stem,
        "filename": file.filename,
        "size_bytes": len(content)
    }

@router.get("/templates/{template_code}/download")
def download_template(template_code: str):
    """Downloads the raw .xlsx template file."""
    template_path = settings.TEMPLATES_DIR / f"{template_code}.xlsx"
    if not template_path.is_file():
        alt_path = settings.TEMPLATES_DIR / template_code
        if alt_path.is_file():
            template_path = alt_path
        else:
            raise TemplateNotFoundException(template_code)

    return FileResponse(
        path=str(template_path),
        filename=template_path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@router.get("/templates/{template_code}/inspect")
def inspect_template(template_code: str):
    """Inspects an Excel template to discover tags, repeating blocks, formulas, and schema."""
    template_path = settings.TEMPLATES_DIR / f"{template_code}.xlsx"
    if not template_path.is_file():
        alt_path = settings.TEMPLATES_DIR / template_code
        if alt_path.is_file():
            template_path = alt_path
        else:
            raise TemplateNotFoundException(template_code)

    return TemplateInspector.inspect_file(template_path, template_code)

@router.get("/studio", response_class=HTMLResponse)
@router.get("/playground", response_class=HTMLResponse)
def serve_studio():
    """Serves the interactive in-browser ApexReports Studio."""
    studio_path = Path(__file__).parent.parent / "static" / "studio.html"
    if not studio_path.is_file():
        raise HTTPException(status_code=404, detail="Studio frontend asset not found.")
    return HTMLResponse(content=studio_path.read_text(encoding="utf-8"))

@router.post("/reports/render")
def render_report(req: RenderRequest):
    """Core rendering endpoint: populates an Excel template with JSON data and outputs PDF, HTML, or XLSX."""
    start_time = time.perf_counter()

    # 1. Locate template
    template_path = settings.TEMPLATES_DIR / f"{req.template_code}.xlsx"
    if not template_path.is_file():
        # Try direct match
        alt_path = settings.TEMPLATES_DIR / req.template_code
        if alt_path.is_file():
            template_path = alt_path
        else:
            raise TemplateNotFoundException(req.template_code)

    # 2. Load workbook preserving formulas
    try:
        wb = load_workbook(str(template_path), data_only=False)
    except Exception as e:
        raise TemplateSyntaxException(f"Failed to read Excel workbook: {str(e)}")

    # 3. Process template directives & formulas
    processor = TemplateProcessor(wb)
    populated_wb = processor.process(req.data)

    # 4. Generate output bytes/string based on format
    filename = req.filename or f"{req.template_code}.{req.format}"
    mime_type = MIME_TYPES.get(req.format, "application/octet-stream")

    if req.format == "html":
        content_str = FastHtmlRenderer.render(populated_wb.active, title=filename)
        output_bytes = content_str.encode("utf-8")

    elif req.format == "xlsx":
        buf = io.BytesIO()
        populated_wb.save(buf)
        buf.seek(0)
        output_bytes = buf.getvalue()

    elif req.format == "pdf":
        output_bytes = PdfService.convert_to_pdf(populated_wb, filename=filename)

    else:
        raise HTTPException(status_code=400, detail=f"Unsupported format '{req.format}'")

    execution_time_ms = round((time.perf_counter() - start_time) * 1000, 2)

    # 5. Return based on response_mode
    if req.response_mode == "json":
        b64_data = base64.b64encode(output_bytes).decode("ascii")
        return RenderResponseJson(
            status="success",
            template_code=req.template_code,
            format=req.format,
            filename=filename,
            mime_type=mime_type,
            size_bytes=len(output_bytes),
            execution_time_ms=execution_time_ms,
            data_base64=b64_data
        )

    # Default: binary stream
    headers = {
        "Content-Disposition": f'inline; filename="{filename}"',
        "X-Execution-Time-Ms": str(execution_time_ms)
    }

    return Response(
        content=output_bytes,
        media_type=mime_type,
        headers=headers
    )
