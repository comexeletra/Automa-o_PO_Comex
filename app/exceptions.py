class PurchaseOrderError(Exception):
    """Base error exposed as a safe Portuguese message by the web layer."""


class InvalidPdfError(PurchaseOrderError):
    pass


class UnsupportedLayoutError(PurchaseOrderError):
    pass


class PdfTextLayerNotFoundError(PurchaseOrderError):
    pass


class ExcelTemplateError(PurchaseOrderError):
    pass


class ValidationError(PurchaseOrderError):
    pass
