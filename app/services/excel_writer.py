"""Write only direct PDF fields into a clean Excel template."""
from __future__ import annotations

from copy import copy
from decimal import Decimal
from pathlib import Path
import re
import posixpath
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter

from app.exceptions import ExcelTemplateError
from app.models import PurchaseOrder


SHEET_NAME = "Material for production"
ITEM_COLUMNS = range(1, 20)


def _norm(value: object) -> str:
    return " ".join(str(value or "").lower().split())


def find_row_containing(ws, text: str) -> int | None:
    target = _norm(text)
    for row in ws.iter_rows():
        if any(target in _norm(cell.value) for cell in row):
            return row[0].row
    return None


def _header_row(ws) -> int:
    for row in ws.iter_rows():
        values = {_norm(cell.value) for cell in row}
        if "po #" in values and "supply code" in values and "total qty" in values and "total (cny)" in values:
            return row[0].row
    raise ExcelTemplateError("Cabeçalho da tabela de materiais não encontrado no template.")


def _copy_row_style(ws, source: int, destination: int) -> None:
    ws.row_dimensions[destination].height = ws.row_dimensions[source].height
    for col in ITEM_COLUMNS:
        source_cell, destination_cell = ws.cell(source, col), ws.cell(destination, col)
        destination_cell._style = copy(source_cell._style)
        destination_cell.number_format = source_cell.number_format
        destination_cell.protection = copy(source_cell.protection)
        destination_cell.alignment = copy(source_cell.alignment)


def _translate_merges(ws, item_start: int, lower_start: int, total_row: int, new_count: int, delta: int) -> None:
    original = list(ws.merged_cells.ranges)
    for merge in original:
        ws.unmerge_cells(str(merge))
    for merge in original:
        min_col, min_row, max_col, max_row = merge.bounds
        if min_row >= lower_start:
            min_row += delta
            max_row += delta
        elif min_row >= item_start:
            if min_row == item_start and max_row == total_row - 1:
                min_row, max_row = item_start, item_start + new_count - 1
            elif min_row == total_row and max_row == total_row:
                min_row = max_row = item_start + new_count
            else:
                # Item-area vertical merges other than the PO merge are not valid output data.
                continue
        ws.merge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)


def _clear_row(ws, row: int) -> None:
    for col in ITEM_COLUMNS:
        cell = ws.cell(row, col)
        if not isinstance(cell, MergedCell):
            cell.value = None


def _value_cell(ws, row: int, col: int):
    """Return the writable top-left cell when a total row has a merged field."""
    cell = ws.cell(row, col)
    if not isinstance(cell, MergedCell):
        return cell
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col <= col <= merged.max_col:
            return ws.cell(merged.min_row, merged.min_col)
    raise ExcelTemplateError(f"Célula mesclada sem origem: {get_column_letter(col)}{row}")


def _write_header_values(ws, po: PurchaseOrder) -> None:
    # Only values next to labels with an unambiguous PDF equivalent are written.
    mappings = {"seller name": po.supplier_name, "buyer name": po.company_name, "buyer cnpj": po.company_cnpj,
                "date": po.issue_date, "payment terms": po.payment_terms}
    for row in ws.iter_rows():
        for cell in row:
            label = _norm(cell.value).rstrip(":")
            if label in mappings and mappings[label] is not None:
                adjacent = ws.cell(cell.row, cell.column + 1)
                if not isinstance(adjacent, MergedCell):
                    adjacent.value = mappings[label]


def _set_lower_direct_fields(ws, lower_start: int, po: PurchaseOrder) -> None:
    scs = list(dict.fromkeys(item.sc_number for item in po.items if item.sc_number))
    for row in ws.iter_rows(min_row=lower_start):
        for cell in row:
            label = _norm(cell.value).rstrip(":")
            value = None
            if label in {"po", "po #", "purchase order"}:
                value = f"PO{po.po_number}" if label == "po" else po.po_number
            elif label in {"sc", "scs", "purchase requisition"} and scs:
                value = "SC - " + " - ".join(scs)
            elif label == "payment terms" and po.payment_terms:
                value = po.payment_terms
            if value is not None:
                adjacent = ws.cell(cell.row, cell.column + 1)
                if not isinstance(adjacent, MergedCell):
                    adjacent.value = value
    remarks_row = find_row_containing(ws, "remarks")
    if remarks_row is not None:
        # The source workbook reserves fixed visual lines below REMARKS. They
        # have no labels once the manual example is cleaned, so populate them
        # relative to the section title rather than a hard-coded worksheet row.
        ws.cell(remarks_row + 4, 1).value = f"PO{po.po_number}"
        if scs:
            ws.cell(remarks_row + 5, 1).value = "SC - " + " - ".join(scs)
        if po.payment_terms:
            ws.cell(remarks_row + 7, 1).value = f"Payment Terms: {po.payment_terms}"


def _validate_output(path: Path, po: PurchaseOrder) -> None:
    wb = load_workbook(path, data_only=False, keep_links=False)
    try:
        if SHEET_NAME not in wb.sheetnames:
            raise ExcelTemplateError("Aba Material for production ausente no arquivo gerado.")
        ws = wb[SHEET_NAME]
        start = _header_row(ws) + 1
        for offset, item in enumerate(po.items):
            row = start + offset
            actual = (ws.cell(row, 2).value, ws.cell(row, 3).value, ws.cell(row, 6).value)
            expected = (item.sequence, item.product_code, item.description)
            numbers_match = all(
                Decimal(str(ws.cell(row, col).value)) == value
                for col, value in ((13, item.quantity), (14, item.unit_price), (15, item.total_value))
            )
            if actual != expected or not numbers_match or any(ws.cell(row, col).value is not None for col in (4, 5, 16, 17, 18, 19)):
                raise ExcelTemplateError(f"Validação da linha de material falhou: {row}.")
            if any(isinstance(ws.cell(row, col).value, str) and ws.cell(row, col).value.startswith("=") for col in ITEM_COLUMNS):
                raise ExcelTemplateError("Fórmula encontrada no bloco de materiais.")
    finally:
        wb.close()


def write_purchase_order(po: PurchaseOrder, template_path: Path, output_path: Path) -> None:
    if not template_path.is_file():
        raise ExcelTemplateError(f"Template não encontrado: {template_path}")
    wb = load_workbook(template_path, keep_links=False)
    if SHEET_NAME not in wb.sheetnames:
        raise ExcelTemplateError("Aba Material for production ausente no template.")
    ws = wb[SHEET_NAME]
    header_row = _header_row(ws)
    item_start = header_row + 1
    approval_row = find_row_containing(ws, "approval")
    if approval_row is None:
        raise ExcelTemplateError("Bloco APPROVAL não encontrado no template.")
    total_row = approval_row - 2
    lower_start = approval_row
    capacity = total_row - item_start
    if capacity < 1:
        raise ExcelTemplateError("Bloco de itens inválido no template.")
    delta = len(po.items) - capacity
    if delta > 0:
        ws.insert_rows(lower_start, delta)
        for row in range(lower_start, lower_start + delta):
            _copy_row_style(ws, total_row - 1, row)
    elif delta < 0:
        ws.delete_rows(item_start + len(po.items), -delta)
    _translate_merges(ws, item_start, lower_start, total_row, len(po.items), delta)
    total_row = item_start + len(po.items)
    lower_start = total_row + 2
    for row in range(item_start, total_row):
        _clear_row(ws, row)
    for offset, item in enumerate(po.items):
        row = item_start + offset
        ws.cell(row, 2).value = item.sequence
        ws.cell(row, 3).value = item.product_code  # identifier, intentionally text
        ws.cell(row, 3).number_format = "@"
        ws.cell(row, 6).value = item.description
        ws.cell(row, 13).value = item.quantity
        ws.cell(row, 13).number_format = "#,##0.###"
        ws.cell(row, 14).value = item.unit_price
        ws.cell(row, 14).number_format = "0.0000000"
        ws.cell(row, 15).value = item.total_value
        ws.cell(row, 15).number_format = "#,##0.00"
    # Dynamic PO merge may have been recreated from the original; write top-left only.
    ws.cell(item_start, 1).value = po.po_number
    total_qty_cell = _value_cell(ws, total_row, 13)
    total_qty_cell.value = sum((item.quantity for item in po.items), Decimal("0"))
    total_qty_cell.number_format = "#,##0.###"
    total_money_cell = _value_cell(ws, total_row, 15)
    total_money_cell.value = po.merchandise_total if po.merchandise_total is not None else sum((i.total_value for i in po.items), Decimal("0"))
    total_money_cell.number_format = "#,##0.00"
    _write_header_values(ws, po)
    _set_lower_direct_fields(ws, lower_start, po)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()
    remove_external_links(output_path)
    preserve_sheet_drawings(template_path, output_path, lower_block_row=approval_row, row_delta=delta)
    _validate_output(output_path, po)


def remove_external_links(path: Path) -> None:
    """Remove residual OOXML external-link parts left by third-party workbooks."""
    temporary = path.with_suffix(".cleaning.xlsx")
    with ZipFile(path, "r") as source, ZipFile(temporary, "w", ZIP_DEFLATED) as target:
        for info in source.infolist():
            if info.filename.startswith("xl/externalLinks/"):
                continue
            data = source.read(info.filename)
            if info.filename in {"xl/_rels/workbook.xml.rels", "xl/workbook.xml"}:
                text = data.decode("utf-8")
                if info.filename.endswith(".rels"):
                    text = re.sub(r"<Relationship\b[^>]*externalLink[^>]*/>", "", text)
                else:
                    text = re.sub(r"<externalReferences>.*?</externalReferences>", "", text, flags=re.DOTALL)
                data = text.encode("utf-8")
            target.writestr(info, data)
    temporary.replace(path)


def preserve_sheet_drawings(source_path: Path, output_path: Path, *, lower_block_row: int | None = None, row_delta: int = 0) -> None:
    """Preserve DrawingML shapes that openpyxl drops when saving a workbook.

    The template's APPROVAL/signature block is made of native Excel drawing
    objects, not cells. Their anchors are shifted whenever the lower block is.
    """
    sheet_name = "xl/worksheets/sheet1.xml"
    rels_name = "xl/worksheets/_rels/sheet1.xml.rels"
    with ZipFile(source_path, "r") as source:
        names = set(source.namelist())
        if sheet_name not in names or rels_name not in names:
            return
        sheet_xml = source.read(sheet_name).decode("utf-8")
        drawing = re.search(r'<drawing\s+r:id="([^"]+)"\s*/>', sheet_xml)
        if drawing is None:
            return
        source_rid = drawing.group(1)
        rels_xml = source.read(rels_name).decode("utf-8")
        relation = re.search(rf'<Relationship\b(?=[^>]*\bId="{re.escape(source_rid)}")[^>]*/>', rels_xml)
        if relation is None:
            return
        relation_xml = relation.group(0)
        target_match = re.search(r'Target="([^"]+)"', relation_xml)
        if target_match is None:
            return
        drawing_name = posixpath.normpath(posixpath.join("xl/worksheets", target_match.group(1)))
        parts = {name: source.read(name) for name in names if name.startswith("xl/drawings/") or name.startswith("xl/media/")}
        if drawing_name in parts and lower_block_row is not None and row_delta:
            from xml.etree import ElementTree as ET
            root = ET.fromstring(parts[drawing_name])
            ns = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
            for anchor in root:
                origin = anchor.find(ns + "from")
                row = origin.find(ns + "row") if origin is not None else None
                if row is None or int(row.text or "0") < lower_block_row - 1:
                    continue
                for point in (origin, anchor.find(ns + "to")):
                    row_node = point.find(ns + "row") if point is not None else None
                    if row_node is not None:
                        row_node.text = str(int(row_node.text or "0") + row_delta)
            parts[drawing_name] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
        source_types = source.read("[Content_Types].xml").decode("utf-8")

    temporary = output_path.with_suffix(".drawings.xlsx")
    with ZipFile(output_path, "r") as current, ZipFile(temporary, "w", ZIP_DEFLATED) as target:
        current_names = set(current.namelist())
        output_sheet = current.read(sheet_name).decode("utf-8")
        output_rels = current.read(rels_name).decode("utf-8") if rels_name in current_names else '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
        used_ids = set(re.findall(r'Id="(rId\d+)"', output_rels))
        output_rid = source_rid if source_rid not in used_ids else f"rId{max([int(value[3:]) for value in used_ids] or [0]) + 1}"
        if "<drawing " not in output_sheet:
            if "xmlns:r=" not in output_sheet.split(">", 1)[0]:
                output_sheet = output_sheet.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
            # In worksheet XML, DrawingML is a trailing element (after page
            # setup/breaks). Placing it after sheetData corrupts Excel's schema.
            output_sheet = output_sheet.replace("</worksheet>", f'<drawing r:id="{output_rid}"/></worksheet>')
        if not re.search(rf'<Relationship\b(?=[^>]*\bId="{re.escape(output_rid)}")', output_rels):
            output_rels = output_rels.replace("</Relationships>", relation_xml.replace(source_rid, output_rid) + "</Relationships>")
        output_types = current.read("[Content_Types].xml").decode("utf-8")
        for tag in re.findall(r'<(?:Default|Override)\b[^>]*/>', source_types):
            if ("/drawings/" in tag or 'Extension="jpeg"' in tag or 'Extension="png"' in tag) and tag not in output_types:
                output_types = output_types.replace("</Types>", tag + "</Types>")
        for info in current.infolist():
            if info.filename in parts:
                continue
            data = current.read(info.filename)
            if info.filename == sheet_name:
                data = output_sheet.encode("utf-8")
            elif info.filename == rels_name:
                data = output_rels.encode("utf-8")
            elif info.filename == "[Content_Types].xml":
                data = output_types.encode("utf-8")
            target.writestr(info, data)
        if rels_name not in current_names:
            target.writestr(rels_name, output_rels.encode("utf-8"))
        for name, data in parts.items():
            target.writestr(name, data)
    temporary.replace(output_path)


def replace_workbook_logo(workbook_path: Path, logo_path: Path) -> None:
    """Replace the header image without losing its original DrawingML anchor."""
    if not logo_path.is_file():
        raise ExcelTemplateError(f"Logo não encontrada: {logo_path}")
    drawing_rels = "xl/drawings/_rels/drawing1.xml.rels"
    with ZipFile(workbook_path, "r") as source:
        if drawing_rels not in source.namelist():
            raise ExcelTemplateError("Desenho do logo não encontrado no template.")
        rels = source.read(drawing_rels).decode("utf-8")
        image_target = re.search(r'Target="([^\"]+\.(?:jpeg|jpg|png))"', rels, re.IGNORECASE)
        if image_target is None:
            raise ExcelTemplateError("Imagem do logo não encontrada no template.")
        old_name = posixpath.normpath(posixpath.join("xl/drawings", image_target.group(1)))
        new_name = posixpath.splitext(old_name)[0] + ".png"
        new_target = posixpath.relpath(new_name, "xl/drawings")
        new_rels = rels.replace(image_target.group(1), new_target)
        logo = logo_path.read_bytes()
    temporary = workbook_path.with_suffix(".logo.xlsx")
    with ZipFile(workbook_path, "r") as source, ZipFile(temporary, "w", ZIP_DEFLATED) as target:
        types = source.read("[Content_Types].xml").decode("utf-8")
        if 'Extension="png"' not in types:
            types = types.replace("</Types>", '<Default Extension="png" ContentType="image/png"/></Types>')
        for info in source.infolist():
            if info.filename in {old_name, new_name}:
                continue
            data = source.read(info.filename)
            if info.filename == drawing_rels:
                data = new_rels.encode("utf-8")
            elif info.filename == "[Content_Types].xml":
                data = types.encode("utf-8")
            target.writestr(info, data)
        target.writestr(new_name, logo)
    temporary.replace(workbook_path)
