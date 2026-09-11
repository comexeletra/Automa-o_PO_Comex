from __future__ import annotations

from pathlib import Path

try:  # Local package execution
    from app.models import ProcessingResult
    from app.services.excel_writer import write_purchase_order
    from app.services.pdf_parser import parse_purchase_order
    from app.services.validator import validate_purchase_order
except ModuleNotFoundError:  # pywrangler exposes app/ as the Worker root
    from models import ProcessingResult
    from services.excel_writer import write_purchase_order
    from services.pdf_parser import parse_purchase_order
    from services.validator import validate_purchase_order


def process_purchase_order(pdf_path: Path, output_path: Path, template_path: Path) -> ProcessingResult:
    po = parse_purchase_order(pdf_path)
    warnings = validate_purchase_order(po)
    write_purchase_order(po, template_path, output_path)
    return ProcessingResult(job_id=None, po=po, output_path=str(output_path), warnings=warnings)
