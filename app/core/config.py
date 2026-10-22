import os
import shutil
from pathlib import Path
from typing import Optional
from pydantic_settings import BaseSettings, SettingsConfigDict

class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", case_sensitive=True, extra="allow")
    
    PROJECT_NAME: str = "ApexReports Engine"
    API_V1_STR: str = "/api/v1"
    
    # Directory paths
    BASE_DIR: Path = Path(__file__).resolve().parent.parent.parent
    TEMPLATES_DIR: Path = BASE_DIR / "templates"
    OUTPUT_DIR: Path = BASE_DIR / "output"
    
    # LibreOffice configuration
    # Can be overridden by SOFFICE_PATH env var
    SOFFICE_PATH: str = os.getenv("SOFFICE_PATH", "")
    SOFFICE_PORT: int = int(os.getenv("SOFFICE_PORT", "2002"))
    SOFFICE_TIMEOUT_SECONDS: int = 30
    
    # Optional API Key (default empty = no auth required for internal network)
    API_KEY: Optional[str] = os.getenv("API_KEY", None)
    
    # Typography configuration
    DEFAULT_FONT_FAMILY: Optional[str] = os.getenv("DEFAULT_FONT_FAMILY", None)
    FORCE_GLOBAL_FONT: bool = os.getenv("FORCE_GLOBAL_FONT", "false").lower() == "true"
    
    # Execution options
    DEBUG: bool = False

    def get_soffice_executable(self) -> Optional[str]:
        """Detect or validate LibreOffice soffice binary path."""
        if self.SOFFICE_PATH and Path(self.SOFFICE_PATH).is_file():
            return self.SOFFICE_PATH
        
        # Check standard PATH
        found = shutil.which("soffice")
        if found:
            return found
            
        # Common Windows locations
        windows_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for p in windows_paths:
            if Path(p).is_file():
                return p
                
        # Common Linux locations
        linux_paths = [
            "/usr/bin/soffice",
            "/usr/local/bin/soffice",
            "/opt/libreoffice/program/soffice"
        ]
        for p in linux_paths:
            if Path(p).is_file():
                return p
                
        return None

settings = Settings()

# Ensure directories exist
settings.TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
settings.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
