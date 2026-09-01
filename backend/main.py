"""
EHCP Document Processor — Backend API (FastAPI)
"""

import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.settings import BACKEND_HOST, BACKEND_PORT, AUTH_ENABLED, BACKEND_WORKERS
from app.dependencies import *  # noqa: ensure temp/output dirs exist
from app.routers.pipeline import router as pipeline_router

app = FastAPI(
    title="EHCP Document Processor API",
    version="1.0.0",
)

# Allow frontend to call the API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(pipeline_router)


@app.get("/")
async def root():
    return {"message": "EHCP Document Processor API", "docs": "/docs"}


if __name__ == "__main__":
    uvicorn.run("main:app", host=BACKEND_HOST,
                port=BACKEND_PORT, workers=BACKEND_WORKERS)
