from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

from pydantic import BaseModel, Field


class ClinicalIntent(str, Enum):
    GENERAL_EDUCATION = "general_education"
    MEDICATION = "medication"
    CHRONIC_CARE = "chronic_care"
    URGENT_RISK = "urgent_risk"


@dataclass(frozen=True)
class Document:
    doc_id: str
    source: str
    title: str
    text: str


@dataclass(frozen=True)
class Passage:
    doc_id: str
    source: str
    title: str
    text: str
    score: float


@dataclass(frozen=True)
class RetrievalPlan:
    intent: ClinicalIntent
    query: str
    top_k: int
    must_include_disclaimer: bool


class AskRequest(BaseModel):
    question: str = Field(..., min_length=5, max_length=1000)
    top_k: int = Field(default=4, ge=1, le=8)


class Citation(BaseModel):
    source: str
    title: str
    score: float


class AskResponse(BaseModel):
    question: str
    intent: ClinicalIntent
    answer: str
    citations: list[Citation]
    safety_flags: list[str]
    grounded: bool

