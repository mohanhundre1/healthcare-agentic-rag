from __future__ import annotations

from .agents import GroundingVerifierAgent, IntentAgent, PlannerAgent, SafetyAgent, SynthesisAgent
from .knowledge_base import load_documents
from .models import AskResponse, Citation
from .retriever import TfidfRetriever


class HealthcareRagWorkflow:
    def __init__(self, knowledge_dir: str | None = None) -> None:
        documents = load_documents(knowledge_dir)
        self.retriever = TfidfRetriever(documents)
        self.intent_agent = IntentAgent()
        self.planner_agent = PlannerAgent()
        self.safety_agent = SafetyAgent()
        self.synthesis_agent = SynthesisAgent()
        self.verifier_agent = GroundingVerifierAgent()

    def run(self, question: str, top_k: int = 4) -> AskResponse:
        intent = self.intent_agent.classify(question)
        plan = self.planner_agent.plan(question=question, intent=intent, top_k=top_k)
        passages = self.retriever.search(plan.query, top_k=plan.top_k)
        safety_flags = self.safety_agent.review(question, passages, intent)
        answer = self.synthesis_agent.synthesize(question, passages, intent, safety_flags)
        grounded = self.verifier_agent.verify(answer, passages)

        return AskResponse(
            question=question,
            intent=intent,
            answer=answer,
            citations=[
                Citation(source=passage.source, title=passage.title, score=passage.score)
                for passage in passages
            ],
            safety_flags=safety_flags,
            grounded=grounded,
        )

