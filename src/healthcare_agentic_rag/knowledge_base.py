from __future__ import annotations

from pathlib import Path

from .models import Document


SUPPORTED_EXTENSIONS = {".md", ".txt"}


def default_knowledge_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "data" / "sample_guidelines"


def load_documents(path: str | Path | None = None) -> list[Document]:
    knowledge_dir = Path(path) if path else default_knowledge_dir()
    if not knowledge_dir.exists():
        raise FileNotFoundError(f"Knowledge directory does not exist: {knowledge_dir}")

    documents: list[Document] = []
    for file_path in sorted(knowledge_dir.iterdir()):
        if file_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        text = file_path.read_text(encoding="utf-8").strip()
        if not text:
            continue
        title = _extract_title(text, fallback=file_path.stem.replace("-", " ").title())
        documents.append(
            Document(
                doc_id=file_path.stem,
                source=file_path.name,
                title=title,
                text=text,
            )
        )

    if not documents:
        raise ValueError(f"No supported documents found in {knowledge_dir}")
    return documents


def _extract_title(text: str, fallback: str) -> str:
    first_line = text.splitlines()[0].strip()
    return first_line.lstrip("# ").strip() or fallback

