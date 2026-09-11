from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal


@dataclass(frozen=True)
class PurchaseOrderItem:
    source_item_group: str
    sequence: int
    product_code: str
    description: str
    unit: str | None
    quantity: Decimal
    unit_price: Decimal
    total_value: Decimal
    delivery_date: date | None = None
    pdf_cost_center: str | None = None
    sc_number: str | None = None


@dataclass(frozen=True)
class PurchaseOrder:
    po_number: str
    company_name: str | None
    company_cnpj: str | None
    company_address: str | None
    supplier_name: str | None
    supplier_code: str | None
    supplier_address: str | None
    issue_date: date | None
    payment_terms: str | None
    merchandise_total: Decimal | None
    items: tuple[PurchaseOrderItem, ...]


@dataclass(frozen=True)
class ProcessingResult:
    job_id: str | None
    po: PurchaseOrder
    output_path: str
    warnings: tuple[str, ...] = field(default_factory=tuple)

    @property
    def sc_numbers(self) -> tuple[str, ...]:
        return tuple(dict.fromkeys(i.sc_number for i in self.po.items if i.sc_number))
