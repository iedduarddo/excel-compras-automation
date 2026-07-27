"""Motor de pontuação e justificativa das prioridades."""

from __future__ import annotations

from collections.abc import Iterable

from src.core.models import TravelResult
from src.services.text import normalize_text


def _currency(value: float) -> str:
    text = f"{value:,.2f}"
    text = text.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {text}"


def _days(value: int) -> str:
    return f"{value} dia" if value == 1 else f"{value} dias"


def apply_priority_scores(
    travels: list[TravelResult],
    rules: dict[str, object],
) -> list[TravelResult]:
    """Pontua todos os fatores solicitados e classifica as viagens."""

    weights = rules["priority_weights"]
    thresholds = rules["priority_thresholds"]
    max_total = max(item.total_value for item in travels) or 1.0

    for travel in travels:
        score = 0.0
        reasons: list[str] = []
        actions: list[str] = []

        card = normalize_text(travel.card_status)
        if card == "divergente":
            score += weights["card_divergent"]
            reasons.append("cartão corporativo divergente")
            actions.append("corrigir a divergência do cartão antes do Financeiro")
        elif card == "pendente":
            score += weights["card_pending"]
            reasons.append("cartão corporativo pendente")
            actions.append("conferir e regularizar o cartão")
        elif card and card != "conferido":
            score += weights["card_other_issue"]
            reasons.append(f"cartão com status {travel.card_status}")
            actions.append("validar o cartão corporativo")

        criticality = normalize_text(travel.criticality)
        if criticality == "emergencial":
            score += weights["criticality_emergency"]
            reasons.append("viagem emergencial")
            actions.append("obter autorização e tratar a reserva imediatamente")
        elif criticality == "executivo":
            score += weights["criticality_executive"]
            reasons.append("viagem executiva")
            actions.append("confirmar a aprovação e o atendimento prioritário")

        if travel.policy_status == "Revisar":
            score += weights["policy_not_found"]
            reasons.append("política do serviço não localizada")
            actions.append("validar o cadastro da política antes da emissão")
        elif travel.policy_status == "Fora":
            score += weights["outside_policy"]
            if travel.limit_difference > 0:
                reasons.append(
                    f"custo {_currency(travel.limit_difference)} acima do limite"
                )
                actions.append("renegociar a tarifa ou solicitar exceção aprovada")
            if travel.lead_days < travel.min_lead_days:
                reasons.append(
                    f"antecedência de {_days(travel.lead_days)} "
                    f"(mínimo {_days(travel.min_lead_days)})"
                )
                actions.append(
                    "registrar a justificativa do prazo e emitir com urgência"
                )

        booking = normalize_text(travel.booking_status)
        if booking == "pendente":
            score += weights["booking_pending"]
            reasons.append("reserva pendente")
            actions.append("confirmar disponibilidade com o fornecedor")
        elif booking == "remarcacao":
            score += weights["booking_reschedule"]
            reasons.append("reserva em remarcação")
            actions.append("concluir a remarcação e validar eventuais taxas")

        if travel.policy_limit > 0 and travel.limit_difference > 0:
            ratio = travel.limit_difference / travel.policy_limit
            score += min(
                weights["cost_over_limit_max"],
                ratio * weights["cost_over_limit_max"],
            )

        if travel.min_lead_days > 0 and travel.lead_days < travel.min_lead_days:
            shortfall = (travel.min_lead_days - travel.lead_days) / travel.min_lead_days
            score += min(
                weights["lead_time_shortfall_max"],
                shortfall * weights["lead_time_shortfall_max"],
            )

        score += (travel.total_value / max_total) * weights["total_value_max"]
        travel.score = round(score, 2)
        travel.priority = (
            "Crítica"
            if score >= thresholds["critical"]
            else "Alta"
            if score >= thresholds["high"]
            else "Normal"
        )
        travel.reasons = _deduplicate(reasons)
        travel.recommended_actions = _deduplicate(actions)

    return travels


def rank_immediate_requests(
    travels: Iterable[TravelResult],
    quantity: int,
) -> list[TravelResult]:
    """Ordena de forma determinística por risco, custo e ID."""

    return sorted(
        travels,
        key=lambda item: (-item.score, -item.total_value, item.request_id),
    )[:quantity]


def _deduplicate(values: Iterable[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = normalize_text(value)
        if key and key not in seen:
            result.append(value)
            seen.add(key)
    return result
