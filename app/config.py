import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class Settings:
    host: str = os.getenv("APP_HOST", "127.0.0.1")
    port: int = int(os.getenv("APP_PORT", "8000"))
    max_upload_mb: int = int(os.getenv("MAX_UPLOAD_MB", "20"))
    template_path: Path = Path(os.getenv("TEMPLATE_PATH", "templates/po_template.xlsx"))
    storage_root: Path = Path(os.getenv("STORAGE_ROOT", "storage"))
    watch_enabled: bool = os.getenv("WATCH_ENABLED", "false").lower() == "true"
    watch_interval_seconds: int = int(os.getenv("WATCH_INTERVAL_SECONDS", "10"))
    watch_stability_seconds: int = int(os.getenv("WATCH_STABILITY_SECONDS", "5"))


settings = Settings()
