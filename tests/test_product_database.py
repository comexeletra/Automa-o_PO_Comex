from dataclasses import replace
from decimal import Decimal

import pytest
from openpyxl import Workbook

from app.exceptions import ValidationError
from app.models import PurchaseOrder, PurchaseOrderItem
from app.services.product_database import enrich_purchase_order


def order(*codes):
    item = PurchaseOrderItem("1", 1, "", "Descrição reduzida", "UN", Decimal(1), Decimal(2), Decimal(2))
    return PurchaseOrder("1", None, None, None, None, None, None, None, None,
                         Decimal(2 * len(codes)), tuple(replace(item, product_code=code) for code in codes))


def database(path, rows):
    wb = Workbook()
    wb.active.append(["Product Code", "Product Description", "NCM DATABASE"])
    for row in rows:
        wb.active.append(row)
    other = wb.create_sheet("Outra aba")
    other.append(["MISSING", "Não deve ser consultada", "12345678"])
    wb.save(path)
    wb.close()


def test_exact_match_preserves_identifiers_and_uses_only_first_sheet(tmp_path):
    path = tmp_path / "db.xlsx"
    database(path, [(123, "Descrição completa", 85332120), ("00123", "Outro produto", "0101.21.00")])
    po, warnings = enrich_purchase_order(order("123", "00123", "MISSING"), path)
    assert [(i.product_code, i.description, i.ncm) for i in po.items] == [
        ("123", "Descrição completa", "85332120"), ("00123", "Outro produto", "01012100"),
        ("MISSING", "Descrição reduzida", None),
    ]
    assert len(warnings) == 1
    assert "MISSING" in warnings[0]


def test_conflicting_duplicates_and_incomplete_records_are_flagged(tmp_path):
    path = tmp_path / "db.xlsx"
    database(path, [("A", "Produto A", "12345678"), ("A", "Produto B", "87654321"),
                    ("B", "Produto completo", "invalid"), ("C", None, "12345678")])
    po, warnings = enrich_purchase_order(order("A", "B", "C"), path)
    assert (po.items[0].description, po.items[0].ncm) == ("Descrição reduzida", None)
    assert (po.items[1].description, po.items[1].ncm) == ("Produto completo", None)
    assert (po.items[2].description, po.items[2].ncm) == ("Descrição reduzida", "12345678")
    assert len(warnings) == 3


def test_database_must_be_available_and_have_required_columns(tmp_path):
    path = tmp_path / "db.xlsx"
    with pytest.raises(ValidationError, match="não encontrado"):
        enrich_purchase_order(order("A"), path)
    Workbook().save(path)
    with pytest.raises(ValidationError, match="Primeira aba"):
        enrich_purchase_order(order("A"), path)
