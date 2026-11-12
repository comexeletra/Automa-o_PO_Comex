from decimal import Decimal
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pytest
from openpyxl import Workbook, load_workbook

from app.exceptions import UnsupportedLayoutError, ValidationError
from app.services.excel_parser import parse_excel_purchase_order
from app.services.processor import process_excel_purchase_order


TEMPLATE = Path("app/resources/po_draft_template.xlsx")
SOURCE = Path("templates/new_po_model/PO 033193.xlsx")
HEADERS = ["Filial", "Numero", "Item", "Produto", "Descricao", "Quantidade", "Prc Unitario", "Vlr.Total", "Numero da SC", "Desc Fornec"]


def export_file(path, count=1):
    wb = Workbook()
    ws = wb.active
    ws.append(["Listagem do Browse"])
    ws.append(HEADERS)
    for index in range(1, count + 1):
        ws.append(["010101-ELETRA", "003193", f"{index:04}", f"00ABC-{index}", "Material", "1.250,5", "0,0023", "2,88", "000655", "Supplier"])
    wb.save(path)
    wb.close()


@pytest.mark.parametrize("count", [1, 3, 30])
def test_draft_resizes_and_preserves_layout(tmp_path, count):
    source, output = tmp_path / "source.xlsx", tmp_path / "result.xlsx"
    export_file(source, count)
    result = process_excel_purchase_order(source, output, TEMPLATE)
    assert len(result.warnings) == count
    assert "não encontrado" in result.warnings[0]
    assert result.po.items[0].quantity == Decimal("1250.5")
    assert result.po.items[0].unit_price == Decimal("0.0023")
    wb = load_workbook(output)
    ws = wb["Material for production"]
    total = 19 + count
    assert ws["A19"].value == "003193"
    assert ws["C19"].value == "00ABC-1"
    assert ws.cell(total - 1, 3).value == f"00ABC-{count}"
    assert ws.cell(total, 12).value == float(Decimal("1250.5") * count)
    assert Decimal(str(ws.cell(total, 13).value)) == Decimal("2.88") * count
    assert ws.cell(total + 2, 1).value == "APPROVAL"
    assert ws.cell(total + 17, 1).value == "REMARKS"
    assert ws.cell(total + 20, 1).value == "PO003193"
    assert ws.cell(total + 21, 1).value == "SC - 000655"
    assert ws.auto_filter.ref == f"A18:N{total - 1}"
    assert ws.print_area.endswith(f"$A$1:$N${66 + count - 22}")
    assert all(ws[address].value is None for address in ("D19", "I4", "I6", "I8", "I10", "I12"))
    template = load_workbook(TEMPLATE)
    try:
        fixed_header = ("A4", "A5", "A6", "A7", "A9", "A10", "A11", "A12")
        for address in fixed_header:
            assert ws[address].value == template[ws.title][address].value
            assert ws[address].value
        assert ws["A3"].value == "Supplier"
    finally:
        template.close()
    assert all(ws.cell(row, col).data_type != "f" for row in range(19, total + 1) for col in range(1, 15))
    assert f"E19:K19" in {str(merge) for merge in ws.merged_cells.ranges}
    assert wb["Old descriptions"].sheet_state == "hidden"
    wb.close()
    # Native Excel logo/signature drawings must survive openpyxl, including
    # their anchors when the approval block moves up or down.
    ns = {"d": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
    with ZipFile(TEMPLATE) as original, ZipFile(output) as converted:
        before = ET.fromstring(original.read("xl/drawings/drawing1.xml"))
        after = ET.fromstring(converted.read("xl/drawings/drawing1.xml"))
        assert len(before) == len(after)
        for old, new in zip(before, after):
            for point in ("from", "to"):
                a, b = old.find(f"d:{point}/d:row", ns), new.find(f"d:{point}/d:row", ns)
                if a is not None:
                    origin = int(old.find("d:from/d:row", ns).text)
                    assert int(b.text) == int(a.text) + (count - 22 if origin >= 42 else 0)
        assert not any(name.startswith("xl/externalLinks/") for name in converted.namelist())


@pytest.mark.parametrize("column,value", [(2, "999999"), (3, "0001"), (6, "=SUM(A1:A2)"), (6, "NaN")])
def test_invalid_or_mixed_orders_are_rejected(tmp_path, column, value):
    source = tmp_path / "source.xlsx"
    export_file(source, 2)
    wb = load_workbook(source)
    wb.active.cell(4, column).value = value
    wb.save(source)
    wb.close()
    with pytest.raises(ValidationError):
        parse_excel_purchase_order(source)


def test_unrecognized_workbook_is_rejected(tmp_path):
    source = tmp_path / "source.xlsx"
    Workbook().save(source)
    with pytest.raises(UnsupportedLayoutError):
        parse_excel_purchase_order(source)


@pytest.mark.skipif(not SOURCE.exists(), reason="Excel local de referência indisponível")
def test_real_po_and_web_download(tmp_path):
    from fastapi.testclient import TestClient
    from app.main import app

    result = process_excel_purchase_order(SOURCE, tmp_path / "result.xlsx", TEMPLATE)
    assert result.po.po_number == "033193"
    assert len(result.po.items) == 120
    assert result.po.merchandise_total == Decimal("2089539.05")
    assert result.po.items[-1].product_code == "HXAP1000R31043-CEMIG"
    assert len(result.warnings) == 23
    assert result.po.items[0].ncm == "85332120"
    assert result.po.items[0].description.startswith("RESISTOR (RESISTENCIA ELETRICA FIXA)")
    client = TestClient(app)
    response = client.post("/process", files={"file": (SOURCE.name, SOURCE.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
    assert response.status_code == 200
    assert response.headers["X-Item-Count"] == "120"
    assert response.headers["X-PO-Number"] == "033193"
    wb = load_workbook(BytesIO(response.content))
    assert wb["Material for production"]["D19"].value == "85332120"
    assert wb["Material for production"]["E19"].value == result.po.items[0].description
    assert wb["Material for production"]["M139"].value == 2089539.05
    wb.close()


def test_corrupt_excel_returns_readable_error():
    from fastapi.testclient import TestClient
    from app.main import app

    client = TestClient(app)
    response = client.post("/process", files={"file": ("bad.xlsx", b"PK\x03\x04broken", "application/octet-stream")})
    assert response.status_code == 422
