"""Parser dedicated to the compact TOTVS purchase-order layout.

This format uses four-digit item groups and (usually) nine-digit product
codes. Keeping it separate from ``pdf_parser`` prevents changes to the
well-established large-PO conversion path.
"""
from __future__ import annotations

import re
from decimal import Decimal
from pathlib import Path

from pypdf import PdfReader
from pypdf.errors import PdfReadError

try:  # Local package execution
    from app.exceptions import InvalidPdfError, PdfTextLayerNotFoundError, UnsupportedLayoutError
    from app.models import PurchaseOrder, PurchaseOrderItem
    from app.services.pdf_parser import _date, _lines, parse_ptbr_decimal
except ModuleNotFoundError:  # pywrangler exposes app/ as the Worker root
    from exceptions import InvalidPdfError, PdfTextLayerNotFoundError, UnsupportedLayoutError
    from models import PurchaseOrder, PurchaseOrderItem
    from services.pdf_parser import _date, _lines, parse_ptbr_decimal


_ITEM = re.compile(r"^\d{4}$")
_PRODUCT = re.compile(r"^\d{8,12}$")
_DECIMAL = re.compile(r"-?[\d.]+,\d+")
_DATE = re.compile(r"\d{2}/\d{2}/\d{4}")


def _row_text(row) -> str:
    return " ".join(word for _x, _y, _right, word in row)


def _column_text(row, start: float, end: float) -> str:
    return " ".join(word for x, _y, _right, word in row if start <= x < end).strip()


def _decimal_in_column(row, start: float, end: float) -> Decimal:
    value = _DECIMAL.search(_column_text(row, start, end).replace(" ", ""))
    if value is None:
        raise ValueError("Valor numérico não identificado.")
    return parse_ptbr_decimal(value.group(0))


def _first_match(rows, pattern: str) -> re.Match[str] | None:
    for row in rows:
        match = re.search(pattern, _row_text(row), flags=re.IGNORECASE)
        if match:
            return match
    return None


def _metadata(rows) -> dict[str, object | None]:
    text = "\n".join(_row_text(row) for row in rows)
    po_match = re.search(r"\b(\d{6,})\s*/\s*\d+\b", text)
    company_match = _first_match(rows, r"Empresa\s*:\s*(.*?)\s+Raz\S*\s+Social")
    supplier_match = _first_match(rows, r"Raz\S*\s+Social\s*:\s*(.*?)\s+Codigo\s*:")
    supplier_code_match = _first_match(rows, r"Codigo\s*:\s*([^\s]+)")
    total_match = _first_match(rows, r"Total\s+das\s+Mercadorias\s*:\s*([\d.]+,\d+)")

    company_cnpj = None
    for row in rows:
        value = _column_text(row, 0, 230)
        match = re.search(r"CNPJ/CPF\s*:\s*([\d./-]+)", value, flags=re.IGNORECASE)
        if match:
            company_cnpj = match.group(1)
            break

    issue_date = None
    payment_terms = None
    for index, row in enumerate(rows):
        row_text = _row_text(row)
        if "Data de Emissao" in row_text and index + 1 < len(rows):
            match = _DATE.search(_column_text(rows[index + 1], 250, 390))
            if match:
                issue_date = _date(match.group(0))
        if "Condicao de Pagto" in row_text and index + 1 < len(rows):
            match = re.search(r"(\d+\s+DIAS)\b", _column_text(rows[index + 1], 0, 180), flags=re.IGNORECASE)
            if match:
                payment_terms = match.group(1).upper()

    return {
        "po_number": po_match.group(1) if po_match else None,
        "company_name": company_match.group(1).strip() if company_match else None,
        "company_cnpj": company_cnpj,
        "supplier_name": supplier_match.group(1).strip() if supplier_match else None,
        "supplier_code": supplier_code_match.group(1).strip() if supplier_code_match else None,
        "issue_date": issue_date,
        "payment_terms": payment_terms,
        "merchandise_total": parse_ptbr_decimal(total_match.group(1)) if total_match else None,
    }


def parse_small_purchase_order(pdf_path: Path) -> PurchaseOrder:
    """Parse a compact TOTVS PO without falling back to the large-PO rules."""
    if not pdf_path.is_file() or pdf_path.suffix.lower() != ".pdf":
        raise InvalidPdfError("O arquivo enviado não é um PDF válido.")
    try:
        document = PdfReader(str(pdf_path))
    except PdfReadError as exc:
        raise InvalidPdfError("O arquivo enviado não é um PDF válido.") from exc

    try:
        rows = [row for page in document.pages for row in _lines(page)]
        if not rows:
            raise PdfTextLayerNotFoundError("PDF_TEXT_LAYER_NOT_FOUND")
        data = _metadata(rows)
        items: list[PurchaseOrderItem] = []
        current: dict[str, object] | None = None

        def append_current() -> None:
            if current is None:
                return
            items.append(PurchaseOrderItem(
                source_item_group=str(current["group"]), sequence=len(items) + 1,
                product_code=str(current["code"]),
                description=re.sub(r"\s+", " ", " ".join(current["description"])).strip(),
                unit=current["unit"], quantity=current["quantity"], unit_price=current["unit_price"],
                total_value=current["total"], delivery_date=_date(current["delivery"]),
                pdf_cost_center=current["cc"], sc_number=current["sc"],
            ))

        for row in rows:
            tokens = [word for _x, _y, _right, word in row]
            item_pos = next((i for i, token in enumerate(tokens) if _ITEM.fullmatch(token)), None)
            code_pos = item_pos + 1 if item_pos is not None and item_pos + 1 < len(tokens) else None
            if code_pos is not None and _PRODUCT.fullmatch(tokens[code_pos]):
                append_current()
                try:
                    delivery_match = _DATE.search(_column_text(row, 460, 525))
                    current = {
                        "group": tokens[item_pos], "code": tokens[code_pos],
                        "description": [_column_text(row, 85, 170)],
                        "unit": _column_text(row, 170, 230) or None,
                        "quantity": _decimal_in_column(row, 230, 330),
                        "unit_price": _decimal_in_column(row, 360, 425),
                        "total": _decimal_in_column(row, 425, 465),
                        "delivery": delivery_match.group(0) if delivery_match else None,
                        "cc": _column_text(row, 520, 573) or None,
                        "sc": _column_text(row, 573, 650).replace(" ", "") or None,
                        "last_row_y": row[0][1],
                    }
                except ValueError as exc:
                    raise UnsupportedLayoutError("UNSUPPORTED_COMPACT_TOTVS_PDF_LAYOUT") from exc
            elif current is not None:
                # Description continuations remain in the description column;
                # footer and approval text starts in the left margin instead.
                continuation = _column_text(row, 85, 230)
                distance = float(current["last_row_y"]) - row[0][1]
                if continuation and 0 < distance <= 16:
                    current["description"].append(continuation)
                    current["last_row_y"] = row[0][1]
        append_current()

        if not data["po_number"] or not items:
            raise UnsupportedLayoutError("O arquivo não corresponde ao layout de PO menor do TOTVS.")
        return PurchaseOrder(
            po_number=str(data["po_number"]), company_name=data["company_name"], company_cnpj=data["company_cnpj"],
            company_address=None, supplier_name=data["supplier_name"], supplier_code=data["supplier_code"],
            supplier_address=None, issue_date=data["issue_date"], payment_terms=data["payment_terms"],
            merchandise_total=data["merchandise_total"], items=tuple(items),
        )
    finally:
        document.close()
