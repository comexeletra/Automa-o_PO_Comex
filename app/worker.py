"""Cloudflare Workers entry point for the FastAPI application."""

from workers import asgi

from app.main import app


# Cloudflare's ASGI adapter runs FastAPI without Uvicorn.
Default = asgi.entrypoint(app)
