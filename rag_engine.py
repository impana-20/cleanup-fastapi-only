from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


@dataclass(frozen=True)
class DocumentChunk:
    chunk_id: str
    section: str
    text: str


def _split_words(text: str, max_words: int = 220, overlap_words: int = 40) -> list[str]:
    words = text.split()
    if not words:
        return []
    if max_words <= overlap_words:
        raise ValueError("max_words must be greater than overlap_words")

    chunks: list[str] = []
    start = 0
    while start < len(words):
        end = min(start + max_words, len(words))
        chunks.append(" ".join(words[start:end]))
        if end == len(words):
            break
        start = end - overlap_words
    return chunks


def load_docx_chunks(
    path: str | Path,
    max_words: int = 220,
    overlap_words: int = 40,
) -> list[DocumentChunk]:
    """Read a DOCX and create chunks while preserving the nearest heading."""
    document = Document(str(path))
    sections: list[tuple[str, str]] = []
    current_heading = "Document overview"
    current_lines: list[str] = []

    def flush() -> None:
        nonlocal current_lines
        text = "\n".join(current_lines).strip()
        if text:
            sections.append((current_heading, text))
        current_lines = []

    for block in document.element.body.iterchildren():
        if block.tag == qn("w:p"):
            paragraph = Paragraph(block, document)
            text = paragraph.text.strip()
            if not text:
                continue
            style_name = paragraph.style.name if paragraph.style else ""
            if style_name.startswith("Heading") or style_name == "Title":
                flush()
                current_heading = text
            else:
                current_lines.append(text)
        elif block.tag == qn("w:tbl"):
            table = Table(block, document)
            for row in table.rows:
                values = [cell.text.strip().replace("\n", " ") for cell in row.cells]
                if any(values):
                    current_lines.append(" | ".join(values))
    flush()

    chunks: list[DocumentChunk] = []
    for section_index, (heading, text) in enumerate(sections, start=1):
        for part_index, part in enumerate(
            _split_words(text, max_words=max_words, overlap_words=overlap_words),
            start=1,
        ):
            chunks.append(
                DocumentChunk(
                    chunk_id=f"S{section_index:02d}-C{part_index:02d}",
                    section=heading,
                    text=part,
                )
            )
    return chunks


def normalize_rows(vectors: np.ndarray) -> np.ndarray:
    vectors = np.asarray(vectors, dtype=np.float32)
    if vectors.ndim == 1:
        vectors = vectors.reshape(1, -1)
    norms = np.linalg.norm(vectors, axis=1, keepdims=True)
    norms[norms == 0] = 1.0
    return vectors / norms


def embed_texts(client, texts: Iterable[str], model: str) -> np.ndarray:
    items = list(texts)
    response = client.embeddings.create(model=model, input=items)
    vectors = np.asarray([item.embedding for item in response.data], dtype=np.float32)
    return normalize_rows(vectors)


def retrieve(
    query_vector: np.ndarray,
    index_vectors: np.ndarray,
    chunks: list[DocumentChunk],
    top_k: int = 4,
) -> list[tuple[DocumentChunk, float]]:
    if len(chunks) != len(index_vectors):
        raise ValueError("chunks and index_vectors must have the same length")
    if not chunks:
        return []
    query = normalize_rows(query_vector)[0]
    scores = normalize_rows(index_vectors) @ query
    best = np.argsort(scores)[::-1][:top_k]
    return [(chunks[int(index)], float(scores[int(index)])) for index in best]


def build_context(results: list[tuple[DocumentChunk, float]]) -> str:
    blocks = []
    for chunk, _score in results:
        blocks.append(
            f'<policy_section name="{chunk.section}" id="{chunk.chunk_id}">\n'
            f"{chunk.text}\n"
            f"</policy_section>"
        )
    return "\n\n".join(blocks)
