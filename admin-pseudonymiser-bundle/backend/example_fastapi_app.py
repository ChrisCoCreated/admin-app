from __future__ import annotations

from fastapi import FastAPI

from pseudonymiser_core.router import router as pseudonymiser_router

app = FastAPI(title="Admin API")
app.include_router(pseudonymiser_router, prefix="/api/pseudonymiser")
