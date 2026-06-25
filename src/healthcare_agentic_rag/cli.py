from __future__ import annotations

import argparse
import json
import os

from .workflow import HealthcareRagWorkflow


def main() -> None:
    parser = argparse.ArgumentParser(description="Run the healthcare agentic RAG workflow.")
    parser.add_argument("question", help="Healthcare question to answer")
    parser.add_argument("--top-k", type=int, default=4, help="Number of passages to retrieve")
    args = parser.parse_args()

    workflow = HealthcareRagWorkflow(knowledge_dir=os.getenv("RAG_KNOWLEDGE_DIR"))
    response = workflow.run(args.question, top_k=args.top_k)
    print(json.dumps(response.model_dump(), indent=2))


if __name__ == "__main__":
    main()

