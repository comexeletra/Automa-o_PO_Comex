from __future__ import annotations

import logging
import json
from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter
from urllib.parse import quote

from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

try:  # Available in the Cloudflare Python Workers runtime.
    from workers import asgi
except ModuleNotFoundError:  # Keep the regular local FastAPI workflow intact.
    asgi = None

try:  # Local package execution
    from app.config import APP_ROOT, settings
    from app.exceptions import PurchaseOrderError
    from app.services.processor import process_purchase_order
except ModuleNotFoundError:  # pywrangler exposes app/ as the Worker root
    from config import APP_ROOT, settings
    from exceptions import PurchaseOrderError
    from services.processor import process_purchase_order

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
app = FastAPI(title="Conversor Pedido TOTVS")
app.mount("/static", StaticFiles(directory=APP_ROOT / "web" / "static"), name="static")
app.mount("/assets", StaticFiles(directory=APP_ROOT / "resources"), name="assets")
templates = Jinja2Templates(directory=APP_ROOT / "web" / "templates")


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse(request, "index.html", {"max_upload_mb": settings.max_upload_mb})


@app.get("/health")
def health():
    return {"status": "ok"}


@app.get("/result", response_class=HTMLResponse)
def result_screen(request: Request):
    return templates.TemplateResponse(request, "result.html")


@app.post("/process", response_class=HTMLResponse)
async def process(request: Request, file: UploadFile = File(...)):
    if file.content_type not in {"application/pdf", "application/x-pdf"} or not (file.filename or "").lower().endswith(".pdf"):
        return templates.TemplateResponse(request, "error.html", {"message": "Envie apenas um arquivo PDF válido."}, status_code=400)
    content = await file.read()
    if len(content) > settings.max_upload_mb * 1024 * 1024:
        return templates.TemplateResponse(request, "error.html", {"message": f"O PDF excede o limite de {settings.max_upload_mb} MB."}, status_code=413)
    if not content.startswith(b"%PDF-"):
        return templates.TemplateResponse(request, "error.html", {"message": "O PDF possui assinatura inválida."}, status_code=400)
    # Workers expose only an ephemeral filesystem. Generate the workbook in a
    # request-scoped directory and return it immediately, rather than keeping a
    # job under storage/ for a later download request.
    with TemporaryDirectory(prefix="totvs-po-") as directory:
        job_dir = Path(directory)
        input_path, output_path = job_dir / "input.pdf", job_dir / "output.xlsx"
        input_path.write_bytes(content)
        started_at = perf_counter()
        try:
            result = process_purchase_order(input_path, output_path, settings.template_path)
        except PurchaseOrderError as exc:
            logger.exception("Falha ao processar o pedido")
            return templates.TemplateResponse(request, "error.html", {"message": str(exc)}, status_code=422)
        processing_seconds = perf_counter() - started_at
        workbook = output_path.read_bytes()
        if len(workbook) > settings.max_output_mb * 1024 * 1024:
            return templates.TemplateResponse(
                request,
                "error.html",
                {"message": f"A planilha gerada excede o limite de {settings.max_output_mb} MB para download."},
                status_code=413,
            )

    filename = f"pedido_totvs_{result.po.po_number}.xlsx"
    logger.info("Pedido %s concluído em %.3f segundos", result.po.po_number, processing_seconds)
    return Response(
        content=workbook,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-PO-Number": quote(result.po.po_number, safe=""),
            "X-Supplier": quote(result.po.supplier_name or "—", safe=""),
            "X-Item-Count": str(len(result.po.items)),
            "X-Merchandise-Total": quote(str(result.po.merchandise_total or "—"), safe=""),
            "X-SCs": quote(", ".join(result.sc_numbers) or "—", safe=""),
            "X-Warnings": quote(json.dumps(result.warnings, ensure_ascii=False), safe=""),
            "X-Processing-Seconds": f"{processing_seconds:.2f}",
        },
    )


# Python Workers expects the entrypoint to be registered by the module declared
# in wrangler.jsonc. Keeping it here avoids a second module during startup.
if asgi is not None:
    Default = asgi.entrypoint(app)
