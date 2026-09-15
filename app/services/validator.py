from decimal import Decimal, ROUND_HALF_UP

try:  # Local package execution
    from app.exceptions import ValidationError
    from app.models import PurchaseOrder
except ModuleNotFoundError:  # pywrangler exposes app/ as the Worker root
    from exceptions import ValidationError
    from models import PurchaseOrder


def validate_purchase_order(po: PurchaseOrder) -> tuple[str, ...]:
    if not po.po_number:
        raise ValidationError("Número da PO não identificado.")
    if not po.items:
        raise ValidationError("O arquivo foi lido, mas nenhum item da PO foi identificado.")
    warnings: list[str] = []
    for item in po.items:
        if not item.product_code:
            raise ValidationError("Item sem código de produto.")
        if item.quantity is None or item.unit_price is None or item.total_value is None:
            raise ValidationError("Item com quantidade, preço ou total inválido.")
        expected = (item.quantity * item.unit_price).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        if abs(expected - item.total_value) > Decimal("0.01"):
            warnings.append(f"ITEM_TOTAL_MISMATCH:{item.product_code}")
        if not item.description:
            warnings.append(f"DESCRIPTION_EMPTY:{item.product_code}")
        if item.sc_number and len(item.sc_number) < 6:
            warnings.append(f"SC_INCOMPLETE:{item.product_code}")
    item_total = sum((item.total_value for item in po.items), Decimal("0"))
    if po.merchandise_total is not None and abs(item_total - po.merchandise_total) > Decimal("0.01"):
        warnings.append("MERCHANDISE_TOTAL_MISMATCH")
    return tuple(warnings)


def unique_preserving_order(values: list[str]) -> list[str]:
    return list(dict.fromkeys(value for value in values if value))
