from __future__ import annotations

import logging
from time import perf_counter
from uuid import UUID, uuid4

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.config import settings
from app.exceptions import PurchaseOrderError
from app.services.processor import process_purchase_order

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
app = FastAPI(title="Conversor Pedido TOTVS")
app.mount("/static", StaticFiles(directory="app/web/static"), name="static")
app.mount("/assets", StaticFiles(directory="assets"), name="assets")
templates = Jinja2Templates(directory="app/web/templates")


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
    job_id = str(uuid4())
    job_dir = settings.storage_root / "jobs" / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    input_path, output_path = job_dir / "input.pdf", job_dir / "output.xlsx"
    input_path.write_bytes(content)
    started_at = perf_counter()
    try:
        result = process_purchase_order(input_path, output_path, settings.template_path)
    except PurchaseOrderError as exc:
        logger.exception("Falha no job %s", job_id)
        return templates.TemplateResponse(request, "error.html", {"message": str(exc)}, status_code=422)
    processing_seconds = perf_counter() - started_at
    logger.info("Job %s concluido em %.3f segundos", job_id, processing_seconds)
    return templates.TemplateResponse(
        request,
        "result.html",
        {
            "result": result,
            "job_id": job_id,
            "scs": result.sc_numbers,
            "processing_seconds": processing_seconds,
        },
    )


@app.get("/download/{job_id}")
def download(job_id: str):
    try:
        UUID(job_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="Arquivo não encontrado") from exc
    output = settings.storage_root / "jobs" / job_id / "output.xlsx"
    if not output.is_file():
        raise HTTPException(status_code=404, detail="Arquivo não encontrado")
    return FileResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename="pedido_totvs.xlsx")
