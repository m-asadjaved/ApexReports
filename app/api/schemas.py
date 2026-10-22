from typing import Any, Dict, List, Literal, Optional
from pydantic import BaseModel, Field

class RenderRequest(BaseModel):
    template_code: str = Field(
        ..., 
        description="The template code corresponding to an .xlsx file in the template repository (e.g. 'INVOICE_V1')"
    )
    format: Literal["pdf", "html", "xlsx"] = Field(
        "pdf", 
        description="Requested output format"
    )
    response_mode: Literal["binary", "json"] = Field(
        "binary", 
        description="Delivery mode: 'binary' streams raw bytes, 'json' returns base64 envelope"
    )
    filename: Optional[str] = Field(
        None, 
        description="Optional custom filename for Content-Disposition header"
    )
    data: Dict[str, Any] = Field(
        default_factory=dict, 
        description="The hierarchical JSON dataset passed from Oracle APEX"
    )

class RenderResponseJson(BaseModel):
    status: str = "success"
    template_code: str
    format: str
    filename: str
    mime_type: str
    size_bytes: int
    execution_time_ms: float
    data_base64: str

class TemplateInfo(BaseModel):
    template_code: str
    filename: str
    size_bytes: int
    modified_at: str

class TemplateListResponse(BaseModel):
    count: int
    templates: List[TemplateInfo]

class LintDiagnostic(BaseModel):
    code: str
    message: str
    cell: Optional[str] = None
    row: Optional[int] = None

class ValidateRequest(BaseModel):
    template_code: str = Field(..., description="Template code to validate")
    data: Optional[Dict[str, Any]] = Field(None, description="Optional JSON payload to validate against template placeholders")

class ValidateResponse(BaseModel):
    valid: bool
    template_code: str
    score: int
    errors: List[LintDiagnostic]
    warnings: List[LintDiagnostic]
    info: List[LintDiagnostic]
    summary: str
