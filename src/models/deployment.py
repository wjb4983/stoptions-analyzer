from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any


REASON_PROMOTION_GATES_PASSED = "PROMOTION_GATES_PASSED"
REASON_SHADOW_UNDERPERFORMANCE = "SHADOW_UNDERPERFORMANCE"
REASON_RISK_BREACH = "RISK_BREACH"
REASON_STABILITY_FAILURE = "STABILITY_FAILURE"
REASON_MANUAL_CANDIDATE_PROMOTION = "MANUAL_CANDIDATE_PROMOTION"
REASON_ROLLBACK_TO_PRIOR_CHAMPION = "ROLLBACK_TO_PRIOR_CHAMPION"


@dataclass(frozen=True)
class PromotionGates:
    min_risk_adjusted_return_delta: float = 0.0
    max_drawdown: float = -0.15
    max_turnover: float = 4.0
    min_stability_score: float = 0.55


@dataclass(frozen=True)
class SlotEvent:
    timestamp: str
    action: str
    reason_code: str
    from_model: str | None
    to_model: str | None
    details: dict[str, Any] = field(default_factory=dict)


@dataclass
class ModelSlots:
    champion: str | None = None
    challenger: str | None = None
    candidate: str | None = None
    prior_champion: str | None = None
    audit_log: list[SlotEvent] = field(default_factory=list)

    def assign_candidate(self, model_id: str) -> None:
        self.candidate = str(model_id)

    def promote_candidate_to_challenger(self) -> bool:
        if not self.candidate:
            return False
        previous = self.challenger
        self.challenger = self.candidate
        self.candidate = None
        self._log(
            action="promote_to_challenger",
            reason_code=REASON_MANUAL_CANDIDATE_PROMOTION,
            from_model=previous,
            to_model=self.challenger,
        )
        return True

    def evaluate_shadow_mode(
        self,
        *,
        champion_metrics: dict[str, Any],
        challenger_metrics: dict[str, Any],
        gates: PromotionGates,
    ) -> dict[str, Any]:
        if not self.champion or not self.challenger:
            return {"decision": "no_op", "reason_code": "MISSING_SLOT", "promoted": False, "rolled_back": False}

        champ_risk_adj = float(champion_metrics.get("risk_adjusted_return", champion_metrics.get("sharpe", 0.0)))
        chal_risk_adj = float(challenger_metrics.get("risk_adjusted_return", challenger_metrics.get("sharpe", 0.0)))
        chal_drawdown = float(challenger_metrics.get("max_drawdown", 0.0))
        chal_turnover = float(challenger_metrics.get("turnover_total", 0.0))
        chal_stability = float(challenger_metrics.get("stability_score", 0.0))

        if chal_drawdown < gates.max_drawdown or chal_turnover > gates.max_turnover:
            self._rollback_challenger(REASON_RISK_BREACH, challenger_metrics)
            return {"decision": "rollback", "reason_code": REASON_RISK_BREACH, "promoted": False, "rolled_back": True}

        if chal_stability < gates.min_stability_score:
            self._rollback_challenger(REASON_STABILITY_FAILURE, challenger_metrics)
            return {"decision": "rollback", "reason_code": REASON_STABILITY_FAILURE, "promoted": False, "rolled_back": True}

        if chal_risk_adj < champ_risk_adj + gates.min_risk_adjusted_return_delta:
            self._rollback_challenger(
                REASON_SHADOW_UNDERPERFORMANCE,
                {
                    "champion_risk_adjusted_return": champ_risk_adj,
                    "challenger_risk_adjusted_return": chal_risk_adj,
                    "required_delta": gates.min_risk_adjusted_return_delta,
                },
            )
            return {"decision": "rollback", "reason_code": REASON_SHADOW_UNDERPERFORMANCE, "promoted": False, "rolled_back": True}

        old_champion = self.champion
        self.prior_champion = old_champion
        self.champion = self.challenger
        self.challenger = old_champion
        self._log(
            action="promote_to_champion",
            reason_code=REASON_PROMOTION_GATES_PASSED,
            from_model=old_champion,
            to_model=self.champion,
            details={
                "champion_risk_adjusted_return": champ_risk_adj,
                "challenger_risk_adjusted_return": chal_risk_adj,
                "drawdown": chal_drawdown,
                "turnover_total": chal_turnover,
                "stability_score": chal_stability,
            },
        )
        return {"decision": "promote", "reason_code": REASON_PROMOTION_GATES_PASSED, "promoted": True, "rolled_back": False}

    def monitor_challenger_and_rollback(
        self,
        *,
        champion_metrics: dict[str, Any],
        challenger_metrics: dict[str, Any],
        gates: PromotionGates,
    ) -> bool:
        """Rollback to prior champion if promoted challenger degrades in production."""
        if not self.prior_champion:
            return False
        if not self.champion:
            return False

        champ_risk_adj = float(champion_metrics.get("risk_adjusted_return", champion_metrics.get("sharpe", 0.0)))
        chal_risk_adj = float(challenger_metrics.get("risk_adjusted_return", challenger_metrics.get("sharpe", 0.0)))
        chal_drawdown = float(challenger_metrics.get("max_drawdown", 0.0))
        chal_turnover = float(challenger_metrics.get("turnover_total", 0.0))

        underperforming = chal_risk_adj + gates.min_risk_adjusted_return_delta < champ_risk_adj
        risk_breach = chal_drawdown < gates.max_drawdown or chal_turnover > gates.max_turnover
        if not (underperforming or risk_breach):
            return False

        promoted_champion = self.champion
        self.champion = self.prior_champion
        self.challenger = promoted_champion
        self._log(
            action="rollback_to_prior_champion",
            reason_code=REASON_ROLLBACK_TO_PRIOR_CHAMPION,
            from_model=promoted_champion,
            to_model=self.champion,
            details={
                "underperforming": underperforming,
                "risk_breach": risk_breach,
                "challenger_metrics": dict(challenger_metrics),
                "champion_metrics": dict(champion_metrics),
            },
        )
        return True

    def _rollback_challenger(self, reason_code: str, details: dict[str, Any]) -> None:
        previous = self.challenger
        self.candidate = previous
        self.challenger = None
        self._log(
            action="rollback_challenger",
            reason_code=reason_code,
            from_model=previous,
            to_model=self.champion,
            details=details,
        )

    def _log(
        self,
        *,
        action: str,
        reason_code: str,
        from_model: str | None,
        to_model: str | None,
        details: dict[str, Any] | None = None,
    ) -> None:
        self.audit_log.append(
            SlotEvent(
                timestamp=datetime.now().isoformat(),
                action=action,
                reason_code=reason_code,
                from_model=from_model,
                to_model=to_model,
                details=dict(details or {}),
            )
        )
