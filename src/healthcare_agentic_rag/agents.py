from __future__ import annotations

import re

from .models import ClinicalIntent, Passage, RetrievalPlan


URGENT_TERMS = {
    "chest pain",
    "stroke",
    "suicide",
    "overdose",
    "trouble breathing",
    "severe bleeding",
}
MEDICATION_TERMS = {"dose", "dosage", "medication", "medicine", "drug", "insulin", "statin"}
CHRONIC_TERMS = {"diabetes", "hypertension", "blood pressure", "asthma", "copd"}


class IntentAgent:
    def classify(self, question: str) -> ClinicalIntent:
        normalized = question.lower()
        if any(term in normalized for term in URGENT_TERMS):
            return ClinicalIntent.URGENT_RISK
        if any(term in normalized for term in MEDICATION_TERMS):
            return ClinicalIntent.MEDICATION
        if any(term in normalized for term in CHRONIC_TERMS):
            return ClinicalIntent.CHRONIC_CARE
        return ClinicalIntent.GENERAL_EDUCATION


class PlannerAgent:
    def plan(self, question: str, intent: ClinicalIntent, top_k: int) -> RetrievalPlan:
        expanded_query = question
        if intent == ClinicalIntent.CHRONIC_CARE:
            expanded_query = f"{question} monitoring follow up lifestyle risk factors"
            if "diabetes" in question.lower():
                expanded_query = f"{expanded_query} glucose A1C kidney foot eye"
            if "hypertension" in question.lower() or "blood pressure" in question.lower():
                expanded_query = f"{expanded_query} blood pressure cardiovascular sodium"
        elif intent == ClinicalIntent.MEDICATION:
            expanded_query = f"{question} medication safety contraindications clinician"

        return RetrievalPlan(
            intent=intent,
            query=expanded_query,
            top_k=top_k,
            must_include_disclaimer=True,
        )


class SafetyAgent:
    def review(self, question: str, passages: list[Passage], intent: ClinicalIntent) -> list[str]:
        flags: list[str] = []
        normalized = question.lower()
        if intent == ClinicalIntent.URGENT_RISK:
            flags.append("urgent_or_emergency_risk")
        if re.search(r"\b\d+\s?(mg|mcg|units|iu)\b", normalized):
            flags.append("specific_dose_request")
        if not passages:
            flags.append("insufficient_retrieval_evidence")
        return flags


class SynthesisAgent:
    def synthesize(
        self,
        question: str,
        passages: list[Passage],
        intent: ClinicalIntent,
        safety_flags: list[str],
    ) -> str:
        if "urgent_or_emergency_risk" in safety_flags:
            return (
                "This may describe an urgent medical situation. Seek emergency care now "
                "or contact local emergency services. I can provide general information, "
                "but I cannot assess emergencies remotely."
            )

        if not passages:
            return (
                "I could not find enough supporting material in the configured healthcare "
                "knowledge base to answer this safely. Please consult a qualified clinician."
            )

        key_points = []
        for passage in passages[:3]:
            cleaned = " ".join(passage.text.split())
            key_points.append(f"- {cleaned}")

        answer = [
            "Based on the retrieved healthcare guidance:",
            *key_points,
            "",
            "This is general educational information, not a diagnosis or treatment plan. "
            "A licensed clinician should confirm what applies to an individual patient.",
        ]
        if intent == ClinicalIntent.MEDICATION:
            answer.append("Do not start, stop, or change medication doses without clinician guidance.")
        return "\n".join(answer)


class GroundingVerifierAgent:
    def verify(self, answer: str, passages: list[Passage]) -> bool:
        if not passages:
            return False
        meaningful_terms = {
            token.lower()
            for passage in passages
            for token in re.findall(r"[a-zA-Z]{5,}", passage.text)
        }
        answer_terms = set(re.findall(r"[a-zA-Z]{5,}", answer.lower()))
        return bool(meaningful_terms.intersection(answer_terms))
