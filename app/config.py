import os
from dataclasses import dataclass
from pathlib import Path

APP_ROOT = Path(__file__).resolve().parent
# In a Python Worker, pywrangler makes the entry-point directory the runtime
# root. Locally, the same files remain under the repository's app/ directory.
PROJECT_ROOT = APP_ROOT.parent if (APP_ROOT.parent / "templates").is_dir() else APP_ROOT
RESOURCE_ROOT = APP_ROOT / "resources"


@dataclass(frozen=True)
class Settings:
    host: str = os.getenv("APP_HOST", "127.0.0.1")
    port: int = int(os.getenv("APP_PORT", "8000"))
    max_upload_mb: int = int(os.getenv("MAX_UPLOAD_MB", "20"))
    template_path: Path = Path(os.getenv("TEMPLATE_PATH", str(RESOURCE_ROOT / "po_template.xlsx")))
    storage_root: Path = Path(os.getenv("STORAGE_ROOT", str(PROJECT_ROOT / "storage")))
    watch_enabled: bool = os.getenv("WATCH_ENABLED", "false").lower() == "true"
    watch_interval_seconds: int = int(os.getenv("WATCH_INTERVAL_SECONDS", "10"))
    watch_stability_seconds: int = int(os.getenv("WATCH_STABILITY_SECONDS", "5"))


settings = Settings()
