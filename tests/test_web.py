import pytest

pytest.importorskip("fastapi")
from fastapi.testclient import TestClient

from app.main import app


def test_health_and_index():
    client = TestClient(app)
    assert client.get("/").status_code == 200
    assert client.get("/health").json() == {"status": "ok"}


def test_non_pdf_is_rejected():
    client = TestClient(app)
    response = client.post("/process", files={"file": ("not-a-pdf.txt", b"text", "text/plain")})
    assert response.status_code == 400
