"""Read the TOTVS 'Listagem do Browse' export without consulting lookup sheets."""
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import unicodedata
from zipfile import BadZipFile

from openpyxl import load_workbook

try:
    from app.exceptions import UnsupportedLayoutError, ValidationError
    from app.models import PurchaseOrder, PurchaseOrderItem
except ModuleNotFoundError:
    from exceptions import UnsupportedLayoutError, ValidationError
    from models import PurchaseOrder, PurchaseOrderItem


def _label(value) -> str:
    return " ".join(unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode().lower().split())


def _text(value) -> str:
    return str(value).strip() if value is not None else ""


def _number(value, row: int, field: str) -> Decimal:
    try:
        # Excel numbers already have their decimal separator resolved. Text
        # numbers follow the locale of this Brazilian TOTVS export.
        text = _text(value)
        if "," in text:
            text = text.replace(".", "").replace(",", ".")
        number = Decimal(text)
        if not number.is_finite():
            raise InvalidOperation
        return number
    except (InvalidOperation, ValueError):
        raise ValidationError(f"Linha {row}: {field} inválido no Excel.") from None


def _date(value, row: int) -> date | None:
    if value is None or not _text(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%Y%m%d"):
        try:
            return datetime.strptime(str(value).strip(), fmt).date()
        except ValueError:
            pass
    raise ValidationError(f"Linha {row}: data inválida no Excel.")


def parse_excel_purchase_order(path: Path) -> PurchaseOrder:
    try:
        wb = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except (BadZipFile, OSError, ValueError, KeyError):
        raise UnsupportedLayoutError("Arquivo Excel inválido ou corrompido.") from None
    try:
        required = {"numero", "item", "produto", "descricao", "quantidade", "prc unitario", "vlr.total"}
        for ws in wb:
            rows = ws.iter_rows(values_only=True)
            for header_row, values in enumerate(rows, 1):
                columns = {_label(value): col for col, value in enumerate(values) if value is not None}
                if required <= columns.keys():
                    break
                if header_row >= 10:
                    break
            else:
                continue
            if required <= columns.keys():
                break
        else:
            raise UnsupportedLayoutError("Cabeçalho da exportação TOTVS não encontrado no Excel.")

        items = []
        metadata: dict[str, str | date | None] = {}
        sequences = set()
        for row_number, values in enumerate(rows, header_row + 1):
            def get(name):
                index = columns.get(name)
                return values[index] if index is not None and index < len(values) else None

            if not any(_text(value) for value in values):
                continue
            for field in ("numero", "item", "produto"):
                if not _text(get(field)) or _text(get(field)).startswith("="):
                    raise ValidationError(f"Linha {row_number}: {field} ausente ou inválido no Excel.")
            current = {
                "po_number": _text(get("numero")),
                "company_name": re.sub(r"^\d+\s*-\s*", "", _text(get("filial"))) or None,
                "supplier_name": _text(get("desc fornec")) or None,
                "supplier_code": _text(get("fornecedor")) or None,
                "issue_date": _date(get("data emissao"), row_number),
            }
            if metadata and current != metadata:
                raise ValidationError(f"Linha {row_number}: o Excel deve conter uma única PO, filial e fornecedor.")
            metadata = current
            item_text = _text(get("item"))
            if not item_text.isdigit() or int(item_text) < 1 or int(item_text) in sequences:
                raise ValidationError(f"Linha {row_number}: número de item inválido ou duplicado.")
            sequence = int(item_text)
            sequences.add(sequence)
            items.append(PurchaseOrderItem(
                source_item_group=item_text, sequence=sequence,
                product_code=_text(get("produto")), description=_text(get("descricao")),
                unit=_text(get("unidade")) or None,
                quantity=_number(get("quantidade"), row_number, "quantidade"),
                unit_price=_number(get("prc unitario"), row_number, "preço unitário"),
                total_value=_number(get("vlr.total"), row_number, "valor total"),
                delivery_date=_date(get("dt. entrega"), row_number),
                pdf_cost_center=_text(get("centro custo")) or None,
                sc_number=_text(get("numero da sc")) or None,
            ))
        if not items:
            raise ValidationError("Nenhum item encontrado na planilha Excel.")
        return PurchaseOrder(
            **metadata, company_cnpj=None, company_address=None, supplier_address=None,
            payment_terms=None,
            merchandise_total=sum((item.total_value for item in items), Decimal("0")),
            items=tuple(items),
        )
    finally:
        wb.close()
