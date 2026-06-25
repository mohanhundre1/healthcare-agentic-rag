from healthcare_agentic_rag.agents import IntentAgent
from healthcare_agentic_rag.models import ClinicalIntent
from healthcare_agentic_rag.workflow import HealthcareRagWorkflow


def test_intent_agent_detects_urgent_risk() -> None:
    assert IntentAgent().classify("I have chest pain and trouble breathing") == ClinicalIntent.URGENT_RISK


def test_workflow_returns_cited_diabetes_answer() -> None:
    workflow = HealthcareRagWorkflow()

    response = workflow.run("What should a patient with diabetes monitor regularly?")

    assert response.intent == ClinicalIntent.CHRONIC_CARE
    assert response.citations
    assert response.grounded is True
    assert "A1C" in response.answer or "blood glucose" in response.answer


def test_workflow_flags_specific_dose_request() -> None:
    workflow = HealthcareRagWorkflow()

    response = workflow.run("Should I take 20mg of a medication for blood pressure?")

    assert "specific_dose_request" in response.safety_flags
    assert "Do not start, stop, or change medication doses" in response.answer

