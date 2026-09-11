from __future__ import annotations

import logging
from pathlib import Path
from tempfile import TemporaryDirectory
from time import perf_counter

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
    return templates.TemplateResponse(request, "index.html")


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/process", response_class=HTMLResponse)
async def process(request: Request, file: UploadFile = File(...)):
    if file.content_type not in {"application/pdf", "application/x-pdf"} or not (file.filename or "").lower().endswith(".pdf"):
        return templates.TemplateResponse(request, "error.html", {"message": "Envie apenas um arquivo PDF válido."}, status_code=400)
    content = await file.read()
    if len(content) > settings.max_upload_mb * 1024 * 1024 or not content.startswith(b"%PDF-"):
        return templates.TemplateResponse(request, "error.html", {"message": "O PDF excede o tamanho permitido ou possui assinatura inválida."}, status_code=400)
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

    filename = f"pedido_totvs_{result.po.po_number}.xlsx"
    logger.info("Pedido %s concluído em %.3f segundos", result.po.po_number, processing_seconds)
    return Response(
        content=workbook,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# Python Workers expects the entrypoint to be registered by the module declared
# in wrangler.jsonc. Keeping it here avoids a second module during startup.
if asgi is not None:
    Default = asgi.entrypoint(app)
