from datetime import date
from decimal import Decimal
from pathlib import Path
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook

from app.models import PurchaseOrder, PurchaseOrderItem
from app.services.excel_writer import write_purchase_order


def _template(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Material for production"
    headers = ["PO #", "#", "Supply code", "NCM", "Eletra code", "Product description", "", "", "", "", "", "", "Total QTY", "Unit price", "Total (CNY)", "Finalidade", "Centro de custos", "Requer LI?", "Projeto"]
    for column, value in enumerate(headers, 1): ws.cell(18, column).value = value
    ws.merge_cells("F18:L18"); ws["F18"] = "Product description"
    ws.merge_cells("A19:A20"); ws.merge_cells("A21:L21"); ws.merge_cells("N21:O21")
    ws["A21"] = "TOTAL"; ws["A23"] = "APPROVAL"; ws["A25"] = "REMARKS"; ws["A27"] = "PO"; ws["A28"] = "SC"; ws["A29"] = "Payment Terms"
    wb.save(path)


def _po(count: int = 3) -> PurchaseOrder:
    items = tuple(PurchaseOrderItem("001", n, f"2010100{n:03d}", f"Description {n}", "UN", Decimal("2.000"), Decimal("1.2345000"), Decimal("2.47"), date(2026, 6, 16), None, "019347") for n in range(1, count + 1))
    return PurchaseOrder("000123", "Buyer", "12.115.480/0001-15", None, "Seller", "EX000135", None, date(2026, 6, 16), "240 DIAS", Decimal("7.41"), items)


def test_writer_resizes_and_keeps_lower_sections(tmp_path):
    template, output = tmp_path / "template.xlsx", tmp_path / "result.xlsx"
    _template(template)
    write_purchase_order(_po(), template, output)
    ws = load_workbook(output, data_only=False)["Material for production"]
    assert [ws.cell(row, 3).value for row in range(19, 22)] == ["2010100001", "2010100002", "2010100003"]
    assert ws["A24"].value == "APPROVAL"
    assert ws["A26"].value == "REMARKS"
    assert ws["D19"].value is None and ws["P19"].value is None
    with ZipFile(output) as archive:
        assert not any(name.startswith("xl/externalLinks/") for name in archive.namelist())


def test_writer_shrinks_item_area(tmp_path):
    template, output = tmp_path / "template.xlsx", tmp_path / "result.xlsx"
    _template(template)
    write_purchase_order(_po(1), template, output)
    ws = load_workbook(output)["Material for production"]
    assert ws["A22"].value == "APPROVAL"


def test_writer_extends_item_area_before_total_without_currency_format(tmp_path):
    output = tmp_path / "result.xlsx"
    write_purchase_order(_po(206), Path("app/resources/po_template.xlsx"), output)
    ws = load_workbook(output, data_only=False)["Material for production"]
    # The original template supports 193 rows. Extra rows must use the item-row
    # style rather than the total/approval rows' currency format.
    assert ws["B212"].value == 194
    assert ws["B212"].number_format == "0"
    assert ws["B224"].value == 206
    assert Decimal(str(ws["N225"].value)) == Decimal("7.41")
    assert ws["A227"].value == "APPROVAL"
