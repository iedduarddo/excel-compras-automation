"""Normalização e comparação conservadora de textos."""

from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher


_NON_ALPHANUMERIC = re.compile(r"[^a-z0-9]+")


def normalize_text(value: object) -> str:
    """Remove acentos, pontuação e diferenças entre maiúsculas e minúsculas."""

    if value is None:
        return ""

    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(
        character
        for character in text
        if not unicodedata.combining(character)
    )
    text = text.casefold().strip()

    return _NON_ALPHANUMERIC.sub(" ", text).strip()


def text_similarity(left: object, right: object) -> float:
    """Retorna uma similaridade entre zero e um."""

    normalized_left = normalize_text(left)
    normalized_right = normalize_text(right)

    if not normalized_left or not normalized_right:
        return 0.0

    if normalized_left == normalized_right:
        return 1.0

    left_tokens = set(normalized_left.split())
    right_tokens = set(normalized_right.split())

    intersection = left_tokens & right_tokens
    union = left_tokens | right_tokens

    jaccard = len(intersection) / len(union) if union else 0.0

    containment = 0.0
    shorter = min(len(left_tokens), len(right_tokens))

    if shorter >= 2 and len(intersection) == shorter:
        containment = 0.94

    sequence = SequenceMatcher(
        None,
        normalized_left,
        normalized_right,
    ).ratio()

    return max(jaccard, containment, sequence)


def sanitize_filename(value: str) -> str:
    """Transforma um nome em um trecho seguro para arquivos."""

    normalized = unicodedata.normalize("NFKD", value.strip())
    without_accents = "".join(
        character
        for character in normalized
        if not unicodedata.combining(character)
    )

    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "", without_accents)
    safe = re.sub(r"\s+", "_", safe).strip("._ ")

    return safe or "Candidato"