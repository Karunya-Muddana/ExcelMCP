"""Semantic search over sheet descriptions.

Ranking is a plain NumPy cosine scan rather than an ANN index. The corpus here is
one short description per sheet — hundreds of vectors for a large workspace, not
millions — so an exact scan is already sub-millisecond, and it drops `hnswlib`,
which ships no wheels and therefore required a C++ toolchain to `pip install`.
That made the package uninstallable on a stock Windows or slim Linux machine.

Vectors are L2-normalised at index time, so cosine similarity is a dot product.
"""

import io
import json
import re
from pathlib import Path
from typing import Any, Optional

import numpy as np

from excelmcp.storage import (
    atomic_write_bytes,
    atomic_write_text,
    get_config_dir,
    log,
    read_json,
)

EXCELMCP_DIR = Path.home() / ".excelmcp"
VECTORS_PATH = EXCELMCP_DIR / "vectors.npy"
METADATA_PATH = EXCELMCP_DIR / "metadata.json"
_LEGACY_INDEX_PATH = EXCELMCP_DIR / "index.bin"

DIM = 384
MODEL_NAME = "BAAI/bge-small-en-v1.5"

_model: Any = None
_vectors: Optional[np.ndarray] = None
_metadata: Optional[list[dict]] = None


def get_model():
    """Loads the embedding model, downloading it on first use (~130 MB)."""
    global _model
    if _model is None:
        from fastembed import TextEmbedding

        _model = TextEmbedding(model_name=MODEL_NAME)
    return _model


def _normalize(matrix: np.ndarray) -> np.ndarray:
    """L2-normalises rows, leaving all-zero rows as zero rather than NaN."""
    norms = np.linalg.norm(matrix, axis=1, keepdims=True)
    norms[norms == 0] = 1.0
    return matrix / norms


def _load() -> tuple[Optional[np.ndarray], list[dict]]:
    """Loads vectors and metadata from disk, tolerating a missing or corrupt index."""
    global _vectors, _metadata
    if _vectors is not None and _metadata is not None:
        return _vectors, _metadata

    metadata = read_json(METADATA_PATH, [])
    if not isinstance(metadata, list) or not metadata:
        _vectors, _metadata = None, []
        if _LEGACY_INDEX_PATH.exists():
            log(
                "[Embed] Found an index from an older ExcelMCP version — "
                "re-run excelmcp-setup to rebuild it."
            )
        else:
            log("[Embed] No index found — run excelmcp-setup to build.")
        return _vectors, _metadata

    if not VECTORS_PATH.exists():
        log("[Embed] Metadata present but vectors missing — run excelmcp-setup.")
        _vectors, _metadata = None, []
        return _vectors, _metadata

    try:
        vectors = np.load(VECTORS_PATH)
    except (OSError, ValueError) as exc:
        log(f"[Embed] Could not read vectors ({exc}) — run excelmcp-setup to rebuild.")
        _vectors, _metadata = None, []
        return _vectors, _metadata

    # A crash between the two writes would previously leave them out of sync and
    # silently mis-attribute search hits to the wrong sheet.
    if vectors.ndim != 2 or vectors.shape[0] != len(metadata):
        log(
            f"[Embed] Index is inconsistent (vectors={getattr(vectors, 'shape', None)}, "
            f"metadata={len(metadata)}) — run excelmcp-setup to rebuild."
        )
        _vectors, _metadata = None, []
        return _vectors, _metadata

    _vectors = vectors.astype(np.float32, copy=False)
    _metadata = metadata
    log(f"[Embed] Loaded {len(_metadata)} embeddings from disk.")
    return _vectors, _metadata


def embed(texts: list[str]) -> list[list[float]]:
    model = get_model()
    return [[float(x) for x in vec] for vec in model.embed(texts)]


def update_embeddings(workspace_path: str, files_dict: dict[str, Any]) -> None:
    """Rebuilds the index for one workspace, preserving entries for all others.

    Previously this replaced the whole index with only the workspace being
    scanned, so scanning a second folder silently wiped the first one's
    embeddings and its `query` tool stopped returning anything.
    """
    global _vectors, _metadata
    workspace_path = workspace_path.rstrip("/")
    get_config_dir()

    documents: list[str] = []
    new_metadata: list[dict] = []

    for filename, file_data in files_dict.items():
        for sheet_name, sheet_data in file_data.get("sheets", {}).items():
            description = sheet_data.get("description", "")
            if not description:
                continue
            documents.append(description)
            # Lexical terms feed the rerank stage: content words from column
            # names and sampled values. Near-duplicate schemas embed almost
            # identically, but only the sheet that actually contains "BESTEX"
            # carries it as a term.
            terms = _tokens(" ".join(sheet_data.get("columns", [])))
            for values in sheet_data.get("sampled_values", {}).values():
                terms |= _tokens(" ".join(values))
            new_metadata.append(
                {
                    "id": f"{workspace_path}::{filename}::{sheet_name}",
                    "workspace": workspace_path,
                    "file": filename,
                    "sheet": sheet_name,
                    "description": description,
                    "lexical_terms": sorted(terms),
                }
            )

    if not documents:
        raise RuntimeError(
            "No sheet descriptions found to embed. "
            "Check that structure discovery ran correctly."
        )

    new_vectors = _normalize(np.asarray(embed(documents), dtype=np.float32))

    existing_vectors, existing_metadata = _load()
    keep_idx = [
        i
        for i, m in enumerate(existing_metadata)
        if m.get("workspace") != workspace_path
    ]

    if keep_idx and existing_vectors is not None:
        merged_vectors = np.vstack([existing_vectors[keep_idx], new_vectors])
        merged_metadata = [existing_metadata[i] for i in keep_idx] + new_metadata
    else:
        merged_vectors = new_vectors
        merged_metadata = new_metadata

    # Metadata is written first: if the process dies between the two writes, the
    # consistency check in _load() sees the mismatch and asks for a rebuild
    # rather than serving wrong results.
    atomic_write_text(METADATA_PATH, json.dumps(merged_metadata, indent=2))
    buf = io.BytesIO()
    np.save(buf, merged_vectors, allow_pickle=False)
    atomic_write_bytes(VECTORS_PATH, buf.getvalue())

    _vectors = merged_vectors
    _metadata = merged_metadata

    log(
        f"[Embed] Indexed {len(new_metadata)} sheet descriptions for "
        f"'{workspace_path}' ({len(merged_metadata)} total) → {VECTORS_PATH}"
    )


_WORD_RE = re.compile(r"[a-z0-9]+")

# How many vector hits enter the lexical rerank, and how much a full lexical
# overlap can add to a cosine score. Reranking is what separates twenty
# near-identical schemas: their vectors collapse to cosine ~1.0 of each other,
# but only the right sheet shares the question's literal content words.
_RERANK_POOL = 20
_LEXICAL_WEIGHT = 0.3


def _tokens(text: str) -> set[str]:
    return set(_WORD_RE.findall(str(text).lower()))


def search(workspace_path: str, query: str, n_results: int = 5) -> list[dict]:
    """Two-stage retrieval: top _RERANK_POOL by cosine, reranked lexically.

    Each result carries `score` (combined), `vector_score`, and
    `lexical_overlap` (fraction of the query's content words found in the
    sheet's column names and sampled values).

    The workspace filter is applied *before* ranking. It used to run afterwards,
    so with more than one workspace indexed the global top-k could be entirely
    other workspaces' sheets and this returned an empty list for a workspace that
    was correctly indexed.
    """
    workspace_path = workspace_path.rstrip("/")
    vectors, metadata = _load()

    if vectors is None or not metadata:
        log("[Embed] Index is empty — run excelmcp-setup first.")
        return []

    candidate_idx = [
        i for i, m in enumerate(metadata) if m.get("workspace") == workspace_path
    ]
    if not candidate_idx:
        log(f"[Embed] No sheets indexed for workspace '{workspace_path}'.")
        return []

    query_vec = _normalize(np.asarray(embed([query]), dtype=np.float32))[0]
    scores = vectors[candidate_idx] @ query_vec

    pool = min(max(_RERANK_POOL, n_results), len(candidate_idx))
    top = np.argsort(-scores)[:pool]

    query_tokens = _tokens(query)
    reranked = []
    for rank in top:
        entry = metadata[candidate_idx[int(rank)]].copy()
        vector_score = float(scores[int(rank)])
        terms = set(entry.get("lexical_terms", []))
        overlap = (
            len(query_tokens & terms) / len(query_tokens) if query_tokens else 0.0
        )
        entry["vector_score"] = vector_score
        entry["lexical_overlap"] = round(overlap, 4)
        entry["score"] = vector_score + _LEXICAL_WEIGHT * overlap
        reranked.append(entry)

    reranked.sort(key=lambda e: -e["score"])
    return reranked[:n_results]


def collection_count(workspace_path: str) -> int:
    workspace_path = workspace_path.rstrip("/")
    _, metadata = _load()
    return sum(1 for m in metadata if m.get("workspace") == workspace_path)
