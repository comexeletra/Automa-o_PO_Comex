"""Optional polling watcher. Enable it only with WATCH_ENABLED=true."""
from __future__ import annotations

import logging
import shutil
import time
from pathlib import Path

from app.config import settings
from app.exceptions import PurchaseOrderError
from app.services.processor import process_purchase_order

logger = logging.getLogger(__name__)


def _stable(path: Path) -> bool:
    first = path.stat()
    time.sleep(settings.watch_stability_seconds)
    second = path.stat()
    return first.st_size == second.st_size and first.st_mtime_ns == second.st_mtime_ns


def run() -> None:
    if not settings.watch_enabled:
        logger.info("Watcher desabilitado; defina WATCH_ENABLED=true para iniciá-lo.")
        return
    root = settings.storage_root
    incoming, processing, completed, errors = (root / name for name in ("incoming", "processing", "completed", "error"))
    for folder in (incoming, processing, completed, errors):
        folder.mkdir(parents=True, exist_ok=True)
    while True:
        for source in incoming.glob("*.pdf"):
            if not _stable(source):
                continue
            active = processing / source.name
            shutil.move(source, active)
            try:
                result = process_purchase_order(active, completed / f"PO_{active.stem}.xlsx", settings.template_path)
                logger.info("PO %s processada com %d itens", result.po.po_number, len(result.po.items))
                shutil.move(active, completed / active.name)
            except PurchaseOrderError as exc:
                logger.exception("Falha ao processar %s", active.name)
                shutil.move(active, errors / active.name)
                (errors / f"{active.stem}.txt").write_text(str(exc), encoding="utf-8")
        time.sleep(settings.watch_interval_seconds)


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    run()
