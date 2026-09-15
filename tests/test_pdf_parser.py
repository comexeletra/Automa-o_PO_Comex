from decimal import Decimal
from pathlib import Path

import pytest

from app.services.pdf_parser import _item_values_by_token_order, parse_purchase_order


SAMPLE = Path("samples/sample_po_027956.pdf")
SCALED_TEXT_SAMPLE = Path("assets/PO 028084 - AIR SHIPMENT AUG-2026 - SC 019418 - SI 000826 - ZEUS NG CLASS D PILOT.pdf")


@pytest.mark.skipif(not SAMPLE.exists(), reason="PDF de referência ainda não foi fornecido ao workspace")
def test_reference_pdf_regression():
    po = parse_purchase_order(SAMPLE)
    assert po.po_number == "027956"
    assert len(po.items) == 193
    assert po.items[0].product_code == "2010100040"
    assert po.items[0].quantity == Decimal("75000.000")
    assert po.items[-1].product_code == "2150184130"
    assert po.items[-1].quantity == Decimal("652.000")
    assert po.merchandise_total == Decimal("9361869.50")
    assert po.payment_terms == "240 DIAS"
    assert sum(item.quantity for item in po.items) == Decimal("53389684.000")
    assert sum(item.total_value for item in po.items) == Decimal("9361869.50")
    quantities = {item.product_code: item.quantity for item in po.items}
    for original_code in ("2010101620", "2010101960", "2020101980", "2020700071", "2060500112"):
        assert original_code in quantities
    assert {"2020302563": Decimal("263800"), "2020303211": Decimal("77000"), "2020302542": Decimal("25766"), "2052700170": Decimal("232000"), "2054300520": Decimal("95497"), "2061500150": Decimal("55000"), "2150160560": Decimal("519"), "3030502830": Decimal("74350"), "3030503490": Decimal("134800"), "3030702050": Decimal("150"), "3030702090": Decimal("50"), "3030900500": Decimal("105600")} == {code: quantities[code] for code in ("2020302563", "2020303211", "2020302542", "2052700170", "2054300520", "2061500150", "2150160560", "3030502830", "3030503490", "3030702050", "3030702090", "3030900500")}
    assert tuple(dict.fromkeys(item.sc_number for item in po.items if item.sc_number)) == ("019347", "019458")


@pytest.mark.skipif(not SCALED_TEXT_SAMPLE.exists(), reason="PDF com escala de texto alternativa nÃ£o estÃ¡ disponÃ­vel")
def test_scaled_text_pdf_regression():
    po = parse_purchase_order(SCALED_TEXT_SAMPLE)
    assert po.po_number == "028084"
    assert len(po.items) == 108
    assert po.items[0].product_code == "2010100040"
    assert po.items[0].quantity == Decimal("5000.000")
    assert po.merchandise_total == Decimal("77631.78")
    assert sum(item.total_value for item in po.items) == Decimal("77631.78")
    assert tuple(dict.fromkeys(item.sc_number for item in po.items if item.sc_number)) == ("019418",)


def test_token_order_fallback_accepts_variable_unit_and_second_unit_column():
    row = [
        (0, 0, 0, "001"), (0, 0, 0, "1234567890"), (0, 0, 0, "Material"),
        (0, 0, 0, "KG"), (0, 0, 0, "12,500"), (0, 0, 0, "0,000"),
        (0, 0, 0, "1,23456701,30"), (0, 0, 0, "15,4326/09/2026"),
        (0, 0, 0, "1202027"), (0, 0, 0, "019418"),
    ]

    values = _item_values_by_token_order(row, 1)

    assert values["description"] == "Material"
    assert values["unit"] == "KG"
    assert values["quantity"] == Decimal("12.500")
    assert values["unit_price"] == Decimal("1.2345670")
    assert values["total"] == Decimal("15.43")
    assert values["delivery"] == "26/09/2026"
