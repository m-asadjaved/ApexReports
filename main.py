from pathlib import Path
from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware

from app.core.config import settings
from app.core.exceptions import ReportEngineException
from app.api.routes import router as api_router

app = FastAPI(
    title=settings.PROJECT_NAME,
    description="High-performance reporting engine for Oracle APEX powered by Excel templates and LibreOffice.",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

# Configure CORS for Oracle APEX integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Exception handler for ReportEngineException
@app.exception_handler(ReportEngineException)
async def report_engine_exception_handler(request: Request, exc: ReportEngineException):
    return JSONResponse(
        status_code=exc.status_code,
        content={
            "status": "error",
            "error_type": exc.__class__.__name__,
            "message": exc.message
        }
    )

# Register routes
app.include_router(api_router, prefix=settings.API_V1_STR)

@app.get("/studio", response_class=HTMLResponse, include_in_schema=False)
@app.get("/playground", response_class=HTMLResponse, include_in_schema=False)
def serve_root_studio():
    studio_path = Path(__file__).parent / "app" / "static" / "studio.html"
    return HTMLResponse(content=studio_path.read_text(encoding="utf-8"))

@app.get("/")
def root():
    return {
        "service": settings.PROJECT_NAME,
        "studio": "/studio",
        "docs": "/docs",
        "api_v1": settings.API_V1_STR,
        "health": f"{settings.API_V1_STR}/health"
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
