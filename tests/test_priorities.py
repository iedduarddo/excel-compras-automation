from datetime import date

from src.business.priorities import (
    _currency,
    _days,
    _deduplicate,
    apply_priority_scores,
    rank_immediate_requests,
)
from src.core.models import TravelResult
from src.settings import load_rules


def make_travel(
    request_id: str,
    *,
    card: str = "Conferido",
    criticality: str = "Normal",
    status: str = "OK",
    difference: float = -100.0,
    lead_days: int = 20,
    min_days: int = 10,
    total: float = 1000.0,
) -> TravelResult:
    return TravelResult(
        worksheet_row=2,
        request_id=request_id,
        request_date=date(2026, 1, 1),
        travel_date=date(2026, 1, 21),
        service_type="Aéreo Nacional",
        supplier="Fornecedor",
        cost_center="CC1001",
        booking_status="Confirmada",
        criticality=criticality,
        card_status=card,
        total_value=total,
        lead_days=lead_days,
        policy_limit=1800.0,
        min_lead_days=min_days,
        policy_status=status,
        limit_difference=difference,
    )


def test_combined_risks_create_higher_priority() -> None:
    normal = make_travel("NORMAL")
    critical = make_travel(
        "CRITICO",
        card="Divergente",
        criticality="Emergencial",
        status="Fora",
        difference=900.0,
        lead_days=1,
        min_days=10,
        total=2700.0,
    )

    apply_priority_scores([normal, critical], load_rules())

    assert critical.priority == "Crítica"
    assert critical.score > normal.score
    assert any("cartão" in reason for reason in critical.reasons)
    assert any("antecedência" in reason for reason in critical.reasons)


def test_ranking_is_deterministic() -> None:
    first = make_travel("VIA-002", total=2000.0)
    second = make_travel("VIA-001", total=2000.0)
    apply_priority_scores([first, second], load_rules())

    ranked = rank_immediate_requests([first, second], quantity=2)

    assert [item.request_id for item in ranked] == ["VIA-001", "VIA-002"]


def test_pending_executive_and_missing_policy_are_explained() -> None:
    travel = make_travel(
        "PENDENTE",
        card="Pendente",
        criticality="Executivo",
        status="Revisar",
    )
    travel.booking_status = "Pendente"

    apply_priority_scores([travel], load_rules())

    assert travel.priority == "Crítica"
    assert travel.score > 75
    assert "cartão corporativo pendente" in travel.reasons
    assert "viagem executiva" in travel.reasons
    assert "política do serviço não localizada" in travel.reasons
    assert "reserva pendente" in travel.reasons
    assert any(
        "regularizar o cartão" in action for action in travel.recommended_actions
    )
    assert any(
        "cadastro da política" in action for action in travel.recommended_actions
    )


def test_other_card_issue_and_reschedule_create_high_priority() -> None:
    travel = make_travel("REMARCACAO", card="Bloqueado")
    travel.booking_status = "Remarcação"

    apply_priority_scores([travel], load_rules())

    assert travel.priority == "Alta"
    assert travel.score == 40.0
    assert "cartão com status Bloqueado" in travel.reasons
    assert "reserva em remarcação" in travel.reasons


def test_outside_policy_by_lead_time_uses_singular_day() -> None:
    travel = make_travel(
        "PRAZO",
        status="Fora",
        difference=0,
        lead_days=1,
        min_days=2,
    )

    apply_priority_scores([travel], load_rules())

    assert travel.priority == "Alta"
    assert travel.score == 50.0
    assert travel.reasons == ["antecedência de 1 dia (mínimo 2 dias)"]
    assert any("emitir com urgência" in action for action in travel.recommended_actions)


def test_formatting_and_deduplication_helpers() -> None:
    assert _currency(1234.5) == "R$ 1.234,50"
    assert _days(1) == "1 dia"
    assert _days(0) == "0 dias"
    assert _deduplicate(["Ação", " acao ", "", "Outra", "outra"]) == [
        "Ação",
        "Outra",
    ]
