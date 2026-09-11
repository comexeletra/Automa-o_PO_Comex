"""Layout-aware parser for textual TOTVS purchase-order PDFs."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import fitz

from app.exceptions import InvalidPdfError, PdfTextLayerNotFoundError, UnsupportedLayoutError
from app.models import PurchaseOrder, PurchaseOrderItem


def parse_ptbr_decimal(value: str) -> Decimal:
    normalized = value.strip().replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return Decimal(normalized)
    except InvalidOperation as exc:
        raise ValueError(f"Número brasileiro inválido: {value!r}") from exc


def _normalized(value: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", value.lower()) if unicodedata.category(c) != "Mn")


@dataclass(frozen=True)
class TotvsPdfLayout:
    product_x0: float = 32
    description_x0: float = 90
    description_x1: float = 192
    unit_x0: float = 192
    quantity_x0: float = 208
    quantity_x1: float = 300
    unit_price_x0: float = 343
    ip_x0: float = 400
    total_x0: float = 400
    delivery_x0: float = 445
    cc_x0: float = 492
    sc_x0: float = 528


_ITEM = re.compile(r"^\d{3}$")
_CODE = re.compile(r"^\d{10}$")
_NUMBER = re.compile(r"^-?[\d.]+,\d+$")


def _lines(page: fitz.Page) -> list[list[tuple[float, float, float, str]]]:
    rows: list[list[tuple[float, float, float, str]]] = []
    for x0, y0, x1, _y1, text, *_ in sorted(page.get_text("words"), key=lambda w: (round(w[1], 1), w[0])):
        if not rows or abs(rows[-1][0][1] - y0) > 2.2:
            rows.append([])
        rows[-1].append((x0, y0, x1, text))
    return rows


def _text_between(row: list[tuple[float, float, float, str]], x0: float, x1: float) -> str:
    return " ".join(word for left, _y, _right, word in row if left >= x0 and left < x1).strip()


def _first_number(row: list[tuple[float, float, float, str]], x0: float, x1: float) -> Decimal:
    text = _text_between(row, x0, x1).replace(" ", "")
    candidates = _NUMBER.findall(text)
    if not candidates:
        raise ValueError(f"Número ausente no intervalo {x0}-{x1}: {text!r}")
    # A seleção por coordenada impede que preço e IPI concatenados virem um só valor.
    return parse_ptbr_decimal(candidates[0])


def _unit_price(row: list[tuple[float, float, float, str]], layout: TotvsPdfLayout) -> Decimal:
    """Split a price glued to the IPI percentage in the textual PDF layer.

    TOTVS emits e.g. ``0,01041601,30`` in one word. Unit prices in this
    layout have seven decimal places, followed by the IPI such as ``1,30``.
    """
    text = _text_between(row, layout.unit_price_x0, layout.ip_x0).replace(" ", "")
    match = re.match(r"^(-?[\d.]+,\d{7})(?:\d+,\d+)?$", text)
    if not match:
        raise ValueError(f"Preço unitário ausente ou inválido: {text!r}")
    return parse_ptbr_decimal(match.group(1))


def _find_label_value(text: str, labels: tuple[str, ...], pattern: str) -> str | None:
    for label in labels:
        match = re.search(label + pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def _date(value: str | None):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%d/%m/%Y").date()
    except ValueError:
        return None


def _metadata(text: str) -> dict[str, object | None]:
    po_match = re.search(r"(?m)^\s*(\d{6,})\s*/\d+", text)
    po = po_match.group(1) if po_match else _find_label_value(text, (r"Pedido(?:\s+de\s+Compra)?",), r"\s*(?:N[ºo.]*)?\s*[:#]?\s*(\d{4,})")
    issue_match = re.search(r"Data\s+de\s+Emiss[aã]o(?:[^\n]*\n){0,4}?\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    issue = issue_match.group(1) if issue_match else _find_label_value(text, (r"Data\s+de\s+Emiss[aã]o",), r"\s*[:#]?\s*(\d{2}/\d{2}/\d{4})")
    cnpj = _find_label_value(text, (r"CNPJ(?:/CPF)?",), r"\s*[:#]?\s*([\d./-]+)")
    payment_match = re.search(r"Condi[cç][aã]o\s+de\s+Pagto\s+\d+(?:[^\n]*\n){0,5}?\s*(\d+\s+DIAS)\b", text, re.IGNORECASE)
    payment = payment_match.group(1).strip() if payment_match else _find_label_value(text, (r"Condi[cç][aã]o\s+de\s+Pagamento",), r"\s*[:#]?\s*([^\n]+)")
    total = _find_label_value(text, (r"Total\s+das\s+Mercadorias",), r"\s*[:#]?\s*([\d.]+,\d+)")
    supplier_code_match = re.search(r"Codigo:\s*([^\s]+)", text, re.IGNORECASE)
    supplier_code = supplier_code_match.group(1) if supplier_code_match else _find_label_value(text, (r"C[oó]digo\s+(?:do\s+)?Fornecedor",), r"\s*[:#]?\s*([^\n]+)")
    supplier_match = re.search(r"Raz[aã]o\s+Social:\s*(.*?)\s{2,}Codigo:", text, re.IGNORECASE)
    supplier = supplier_match.group(1).strip() if supplier_match else _find_label_value(text, (r"Fornecedor",), r"\s*[:#]?\s*([^\n]+)")
    company_match = re.search(r"Empresa:\s*(.*?)\s{2,}Raz[aã]o\s+Social:", text, re.IGNORECASE)
    company = company_match.group(1).strip() if company_match else _find_label_value(text, (r"Empresa",), r"\s*[:#]?\s*([^\n]+)")
    return {"po_number": po, "issue_date": _date(issue), "company_cnpj": cnpj,
            "payment_terms": payment, "merchandise_total": parse_ptbr_decimal(total) if total else None,
            "supplier_code": supplier_code, "supplier_name": supplier, "company_name": company}


def parse_purchase_order(pdf_path: Path, layout: TotvsPdfLayout = TotvsPdfLayout()) -> PurchaseOrder:
    if not pdf_path.is_file() or pdf_path.suffix.lower() != ".pdf":
        raise InvalidPdfError("O arquivo enviado não é um PDF válido.")
    try:
        document = fitz.open(pdf_path)
    except fitz.FileDataError as exc:
        raise InvalidPdfError("O arquivo enviado não é um PDF válido.") from exc
    try:
        all_text = "\n".join(page.get_text() for page in document)
        if len(all_text.strip()) < 100:
            raise PdfTextLayerNotFoundError("PDF_TEXT_LAYER_NOT_FOUND")
        normalized_text = _normalized(all_text)
        if "produto" not in normalized_text or "quantidade" not in normalized_text:
            raise UnsupportedLayoutError("UNSUPPORTED_TOTVS_PDF_LAYOUT")
        data = _metadata(all_text)
        items: list[PurchaseOrderItem] = []
        for page in document:
            current: dict[str, object] | None = None
            for row in _lines(page):
                tokens = [word for _x, _y, _r, word in row]
                item_pos = next((i for i, token in enumerate(tokens) if _ITEM.match(token)), None)
                code_pos = item_pos + 1 if item_pos is not None and item_pos + 1 < len(tokens) else None
                is_start = code_pos is not None and _CODE.match(tokens[code_pos]) is not None
                if is_start:
                    if current:
                        items.append(_to_item(current))
                    product_code = tokens[code_pos]
                    try:
                        quantity = _first_number(row, layout.quantity_x0, layout.quantity_x1)
                        unit_price = _unit_price(row, layout)
                        total_value = _first_number(row, layout.total_x0, layout.delivery_x0)
                    except ValueError as exc:
                        raise UnsupportedLayoutError("UNSUPPORTED_TOTVS_PDF_LAYOUT") from exc
                    current = {"group": tokens[item_pos], "sequence": len(items) + 1, "code": product_code,
                               "description": [_text_between(row, layout.description_x0, layout.description_x1)],
                               "unit": _text_between(row, layout.unit_x0, layout.quantity_x0) or None,
                               "quantity": quantity, "unit_price": unit_price, "total": total_value,
                               "delivery": _text_between(row, layout.delivery_x0, layout.cc_x0),
                               "cc": _text_between(row, layout.cc_x0, layout.sc_x0) or None,
                               "sc": _text_between(row, layout.sc_x0, 1000).replace(" ", "") or None}
                elif current:
                    description = _text_between(row, layout.description_x0, layout.description_x1)
                    # Continuations have no next item and stay strictly inside the description column.
                    if description and not re.search(r"continua|continuacao|pagina\s*\.*", _normalized(description)):
                        current["description"].append(description)  # type: ignore[index]
                    tail = _text_between(row, layout.sc_x0, 1000).replace(" ", "")
                    if tail.isdigit() and len(tail) <= 2 and current.get("sc"):
                        current["sc"] = str(current["sc"]) + tail
            if current:
                items.append(_to_item(current))
        if not items:
            raise UnsupportedLayoutError("O arquivo foi lido, mas nenhum item de pedido foi identificado.")
        if not data["po_number"]:
            raise UnsupportedLayoutError("Número do pedido não identificado.")
        return PurchaseOrder(po_number=str(data["po_number"]), company_name=data["company_name"], company_cnpj=data["company_cnpj"],
            company_address=None, supplier_name=data["supplier_name"], supplier_code=data["supplier_code"], supplier_address=None,
            issue_date=data["issue_date"], payment_terms=data["payment_terms"], merchandise_total=data["merchandise_total"], items=tuple(items))
    finally:
        document.close()


def _to_item(value: dict[str, object]) -> PurchaseOrderItem:
    description = re.sub(r"\s+", " ", " ".join(value["description"])) .strip()  # type: ignore[arg-type]
    return PurchaseOrderItem(source_item_group=str(value["group"]), sequence=int(value["sequence"]), product_code=str(value["code"]),
        description=description, unit=value["unit"] if isinstance(value["unit"], str) else None,
        quantity=value["quantity"], unit_price=value["unit_price"], total_value=value["total"],
        delivery_date=_date(str(value["delivery"]) if value["delivery"] else None), pdf_cost_center=value["cc"] if isinstance(value["cc"], str) else None,
        sc_number=value["sc"] if isinstance(value["sc"], str) else None)
