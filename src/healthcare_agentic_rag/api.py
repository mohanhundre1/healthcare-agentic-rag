from __future__ import annotations

import os
from functools import lru_cache

from fastapi import FastAPI

from .models import AskRequest, AskResponse
from .workflow import HealthcareRagWorkflow


app = FastAPI(
    title="Healthcare Agentic RAG Workflow",
    version="0.1.0",
    description="Agentic retrieval workflow for healthcare education with citations and guardrails.",
)


@lru_cache(maxsize=1)
def get_workflow() -> HealthcareRagWorkflow:
    return HealthcareRagWorkflow(knowledge_dir=os.getenv("RAG_KNOWLEDGE_DIR"))


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/ask", response_model=AskResponse)
def ask(request: AskRequest) -> AskResponse:
    return get_workflow().run(question=request.question, top_k=request.top_k)

