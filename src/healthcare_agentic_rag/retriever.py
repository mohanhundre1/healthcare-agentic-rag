from __future__ import annotations

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

from .models import Document, Passage


class TfidfRetriever:
    def __init__(self, documents: list[Document]) -> None:
        if not documents:
            raise ValueError("At least one document is required")
        self.documents = documents
        self._chunks = self._chunk_documents(documents)
        self._vectorizer = TfidfVectorizer(stop_words="english", ngram_range=(1, 2))
        self._matrix = self._vectorizer.fit_transform([chunk.text for chunk in self._chunks])

    def search(self, query: str, top_k: int = 4) -> list[Passage]:
        query_vector = self._vectorizer.transform([query])
        scores = cosine_similarity(query_vector, self._matrix).flatten()
        ranked = scores.argsort()[::-1][:top_k]
        return [
            Passage(
                doc_id=self._chunks[index].doc_id,
                source=self._chunks[index].source,
                title=self._chunks[index].title,
                text=self._chunks[index].text,
                score=round(float(scores[index]), 4),
            )
            for index in ranked
            if scores[index] > 0
        ]

    @staticmethod
    def _chunk_documents(documents: list[Document]) -> list[Document]:
        chunks: list[Document] = []
        for document in documents:
            paragraphs = [
                p.strip()
                for p in document.text.split("\n\n")
                if p.strip() and not p.strip().startswith("#")
            ]
            for offset, paragraph in enumerate(paragraphs):
                chunks.append(
                    Document(
                        doc_id=f"{document.doc_id}:{offset}",
                        source=document.source,
                        title=document.title,
                        text=paragraph,
                    )
                )
        return chunks
