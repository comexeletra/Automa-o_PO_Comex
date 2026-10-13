from decimal import Decimal
from pathlib import Path

from app.exceptions import UnsupportedLayoutError
from app.models import PurchaseOrder, PurchaseOrderItem
from app.services import processor


def test_main_processor_selects_compact_layout(monkeypatch):
    item = PurchaseOrderItem(
        source_item_group="0001", sequence=1, product_code="ABC123.X45",
        description="Synthetic material", unit="UN", quantity=Decimal("2"),
        unit_price=Decimal("10.00"), total_value=Decimal("20.00"),
    )
    po = PurchaseOrder(
        po_number="000001", company_name=None, company_cnpj=None, company_address=None,
        supplier_name=None, supplier_code=None, supplier_address=None, issue_date=None,
        payment_terms="AVISTA", merchandise_total=Decimal("20.00"), items=(item,),
    )
    written = []

    def unsupported(_path):
        raise UnsupportedLayoutError("UNSUPPORTED_TOTVS_PDF_LAYOUT")

    def write(_po, _template, _output, compact_layout=False):
        written.append(compact_layout)

    monkeypatch.setattr(processor, "parse_purchase_order", unsupported)
    monkeypatch.setattr(processor, "parse_small_purchase_order", lambda _path: po)
    monkeypatch.setattr(processor, "write_purchase_order", write)
    result = processor.process_purchase_order(Path("input.pdf"), Path("output.xlsx"), Path("template.xlsx"))

    assert result.po is po
    assert result.warnings == ()
    assert written == [True]
