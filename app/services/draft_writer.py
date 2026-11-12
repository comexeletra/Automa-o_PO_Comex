"""Populate the 14-column PO DRAFT layout, preserving its lower sections."""
from copy import copy
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.print_settings import PrintArea

try:
    from app.exceptions import ExcelTemplateError
    from app.models import PurchaseOrder
    from app.services.excel_writer import (
        SHEET_NAME, _conversion_date, _insert_row_dimensions, _delete_row_dimensions,
        _restore_signature_lines, find_row_containing, preserve_sheet_drawings, remove_external_links,
    )
except ModuleNotFoundError:
    from exceptions import ExcelTemplateError
    from models import PurchaseOrder
    from services.excel_writer import (
        SHEET_NAME, _conversion_date, _insert_row_dimensions, _delete_row_dimensions,
        _restore_signature_lines, find_row_containing, preserve_sheet_drawings, remove_external_links,
    )


def write_draft_purchase_order(po: PurchaseOrder, template_path: Path, output_path: Path) -> None:
    if not template_path.is_file():
        raise ExcelTemplateError(f"Template não encontrado: {template_path}")
    wb = load_workbook(template_path, keep_links=False)
    try:
        if SHEET_NAME not in wb.sheetnames:
            raise ExcelTemplateError("Aba Material for production ausente no DRAFT.")
        ws = wb[SHEET_NAME]
        header = next((row[0].row for row in ws if row[0].value == "PO #" and row[4].value == "Product description"), None)
        approval = find_row_containing(ws, "approval")
        if header is None or approval is None or approval <= header + 3:
            raise ExcelTemplateError("Layout PO DRAFT não reconhecido.")
        start, old_total = header + 1, approval - 2
        delta = len(po.items) - (old_total - start)
        total = start + len(po.items)
        styles = [copy(ws.cell(start, col)._style) for col in range(1, 15)]
        height = ws.row_dimensions[start].height
        merges = [merge.bounds for merge in ws.merged_cells.ranges]
        for merge in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merge))
        if delta > 0:
            _insert_row_dimensions(ws, old_total, delta)
            ws.insert_rows(old_total, delta)
        elif delta < 0:
            _delete_row_dimensions(ws, total, -delta)
            ws.delete_rows(total, -delta)
        for c1, r1, c2, r2 in merges:
            if r1 >= old_total:
                r1, r2 = r1 + delta, r2 + delta
            elif r1 >= start:
                # Rebuild the PO merge independently; no other vertical item
                # merges are carried into the populated materials table.
                continue
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        for row, item in enumerate(po.items, start):
            ws.row_dimensions[row].height = height
            for col, style in enumerate(styles, 1):
                cell = ws.cell(row, col)
                cell._style = copy(style)
                cell.value = None
            fields = {2: item.sequence, 3: item.product_code, 5: item.description,
                      12: item.quantity, 13: item.unit_price, 14: item.total_value}
            for col, value in fields.items():
                cell = ws.cell(row, col)
                cell.value = value
                if isinstance(value, str):
                    cell.data_type = "s"
            for col, fmt in {2: "0", 3: "@", 12: "#,##0.###", 13: "0.0000000", 14: "#,##0.00"}.items():
                ws.cell(row, col).number_format = fmt
            ws.merge_cells(start_row=row, start_column=5, end_row=row, end_column=11)
        if len(po.items) > 1:
            ws.merge_cells(start_row=start, start_column=1, end_row=total - 1, end_column=1)
        ws.cell(start, 1, po.po_number).data_type = "s"
        ws.cell(total, 12).value = sum((item.quantity for item in po.items), Decimal("0"))
        ws.cell(total, 12).number_format = "#,##0.###"
        ws.cell(total, 13).value = po.merchandise_total
        ws.cell(total, 13).number_format = "#,##0.00"

        # The DRAFT supplies the fixed registration details. The TOTVS branch
        # label is not a replacement for the buyer's full legal name.
        if po.supplier_name:
            ws["A3"] = po.supplier_name
            ws["A3"].data_type = "s"
        if not ws["A9"].value and po.company_name:
            ws["A9"] = po.company_name
            ws["A9"].data_type = "s"
        # Commercial example values must not become facts about the new PO.
        ws["I2"] = _conversion_date()
        for address in ("I4", "I6", "I8", "I10", "I12"):
            ws[address] = None
        ws["I8"] = po.payment_terms
        remarks = find_row_containing(ws, "remarks")
        if remarks:
            ws.cell(remarks + 2, 1).value = "Payment Terms:" + (f" {po.payment_terms}" if po.payment_terms else "")
            ws.cell(remarks + 3, 1).value = f"PO{po.po_number}"
            scs = list(dict.fromkeys(item.sc_number for item in po.items if item.sc_number))
            ws.cell(remarks + 4, 1).value = "SC - " + " - ".join(scs) if scs else None
        _restore_signature_lines(ws, approval + delta)
        ws.auto_filter.ref = f"A{header}:N{total - 1}"
        ws.auto_filter.filterColumn = []
        ws.auto_filter.sortState = None
        if ws.print_area:
            areas = []
            for area in PrintArea.from_string(ws.print_area).ranges:
                c1, r1, c2, r2 = range_boundaries(str(area))
                areas.append(f"A{r1}:N{r2 + delta}")
            ws.print_area = areas
        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)
    finally:
        wb.close()
    remove_external_links(output_path)
    preserve_sheet_drawings(template_path, output_path, lower_block_row=approval, row_delta=delta)
    _validate_draft(output_path, po, start)


def _validate_draft(path: Path, po: PurchaseOrder, start: int) -> None:
    wb = load_workbook(path, data_only=False, keep_links=False)
    try:
        ws = wb[SHEET_NAME]
        for row, item in enumerate(po.items, start):
            if (ws.cell(row, 2).value, ws.cell(row, 3).value, ws.cell(row, 5).value) != (item.sequence, item.product_code, item.description):
                raise ExcelTemplateError(f"Validação da linha de material falhou: {row}.")
            for col, expected in ((12, item.quantity), (13, item.unit_price), (14, item.total_value)):
                if Decimal(str(ws.cell(row, col).value)) != expected:
                    raise ExcelTemplateError(f"Valor divergente na linha {row}.")
            if ws.cell(row, 4).value is not None:
                raise ExcelTemplateError(f"NCM não informado na origem: linha {row}.")
    finally:
        wb.close()
