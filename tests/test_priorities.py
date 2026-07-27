from datetime import date

from src.business.priorities import apply_priority_scores, rank_immediate_requests
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
