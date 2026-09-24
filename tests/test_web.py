from pathlib import Path

import pytest

pytest.importorskip("fastapi")
from fastapi.testclient import TestClient

from app.main import app


def test_health_and_index():
    client = TestClient(app)
    index = client.get("/")
    assert index.status_code == 200
    assert "Arquivos PDF de até 4 MB" in index.text
    assert client.get("/health").json() == {"status": "ok"}


def test_non_pdf_is_rejected():
    client = TestClient(app)
    response = client.post("/process", files={"file": ("not-a-pdf.txt", b"text", "text/plain")})
    assert response.status_code == 400


@pytest.mark.skipif(not Path("20188 (1).pdf").exists(), reason="PDF compacto de referência não disponível")
def test_compact_pdf_download_uses_the_separate_route():
    client = TestClient(app)
    pdf = Path("20188 (1).pdf").read_bytes()
    response = client.post("/process-small", files={"file": ("20188 (1).pdf", pdf, "application/pdf")})
    assert response.status_code == 200
    assert response.headers["content-disposition"] == 'attachment; filename="20188 (1).xlsx"'


def test_pdf_above_vercel_safe_limit_is_rejected():
    client = TestClient(app)
    large_pdf = b"%PDF-" + b"x" * (4 * 1024 * 1024)
    response = client.post("/process", files={"file": ("large.pdf", large_pdf, "application/pdf")})
    assert response.status_code == 413
    assert "excede o limite de 4 MB" in response.text


def test_generated_workbook_above_vercel_safe_limit_is_rejected(monkeypatch):
    def fake_process(_input_path, output_path, _template_path):
        output_path.write_bytes(b"x" * (4 * 1024 * 1024 + 1))

    monkeypatch.setattr("app.main.process_purchase_order", fake_process)
    client = TestClient(app)
    response = client.post("/process", files={"file": ("small.pdf", b"%PDF-1.4", "application/pdf")})
    assert response.status_code == 413
    assert "planilha gerada excede o limite de 4 MB" in response.text


@pytest.mark.skipif(not Path("samples/sample_po_027956.pdf").exists(), reason="PDF de referência não disponível")
def test_reference_pdf_download_fits_vercel_safe_limit():
    client = TestClient(app)
    pdf = Path("samples/sample_po_027956.pdf").read_bytes()
    pdf_name = "PO 027956 - SEA OCT.2026.pdf"
    response = client.post("/process", files={"file": (pdf_name, pdf, "application/pdf")})
    assert response.status_code == 200
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert response.headers["content-disposition"] == 'attachment; filename="PO 027956 - SEA OCT.2026.xlsx"'
    assert len(response.content) <= 4 * 1024 * 1024
