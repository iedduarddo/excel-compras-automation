"""Testes das funções de tratamento de textos."""

from src.services.text import normalize_text, sanitize_filename, text_similarity


def test_normalize_text_removes_accents_and_punctuation() -> None:
    result = normalize_text("Políticas_Fornecedores")

    assert result == "politicas fornecedores"


def test_similar_headers_are_recognized() -> None:
    similarity = text_similarity(
        "Dias de Antecedência",
        "Dias Antecedencia",
    )

    assert similarity >= 0.90


def test_filename_is_safe() -> None:
    result = sanitize_filename("Alex da Silva")

    assert result == "Alex_da_Silva"
