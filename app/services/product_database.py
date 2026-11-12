"""Enrich PO items by an exact Supply code / Product Code match."""
from dataclasses import replace
from pathlib import Path

from openpyxl import load_workbook

try:
    from app.exceptions import ValidationError
    from app.models import PurchaseOrder
except ModuleNotFoundError:
    from exceptions import ValidationError
    from models import PurchaseOrder


def _text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def enrich_purchase_order(po: PurchaseOrder, database_path: Path) -> tuple[PurchaseOrder, tuple[str, ...]]:
    if not database_path.is_file():
        raise ValidationError("Database de produtos não encontrado. Verifique PRODUCT_DATABASE_PATH.")
    try:
        wb = load_workbook(database_path, read_only=True, data_only=True, keep_links=False)
    except Exception as exc:
        raise ValidationError("Não foi possível abrir o database de produtos.") from exc
    products = {}
    ambiguous = set()
    try:
        rows = wb.worksheets[0].iter_rows(values_only=True)
        required = ("Product Code", "Product Description", "NCM DATABASE")
        for row in rows:
            labels = [_text(value) for value in row]
            if all(label in labels for label in required):
                columns = [labels.index(label) for label in required]
                break
        else:
            raise ValidationError("Primeira aba do database sem Product Code, Product Description ou NCM DATABASE.")
        for row in rows:
            code, description, ncm = (_text(row[col]) for col in columns)
            if not code:
                continue
            ncm = ncm.replace(".", "").replace(" ", "")
            if not (ncm.isascii() and ncm.isdigit() and len(ncm) == 8):
                ncm = ""
            product = (description, ncm)
            if code in products and products[code] != product:
                ambiguous.add(code)
            products[code] = product
    finally:
        wb.close()
    items, warnings = [], []
    for item in po.items:
        code = _text(item.product_code)
        if code not in products or code in ambiguous:
            reason = "duplicado com dados divergentes" if code in ambiguous else "não encontrado"
            warnings.append(f"Supply code {code}: {reason} no database; descrição original mantida, NCM não preenchida.")
            items.append(replace(item, ncm=None, ncm_needs_review=True, description_needs_review=True))
            continue
        description, ncm = products[code]
        if not description or not ncm:
            warnings.append(f"Supply code {code}: cadastro incompleto (descrição ou NCM inválida); revise o item.")
        items.append(replace(item, description=description or item.description, ncm=ncm or None,
                             ncm_needs_review=not bool(ncm), description_needs_review=not bool(description)))
    return replace(po, items=tuple(items)), tuple(dict.fromkeys(warnings))
