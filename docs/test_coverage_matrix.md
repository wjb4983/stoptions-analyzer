# Test Coverage Traceability Matrix

This matrix maps major packages to current test coverage posture and release risk.

## Coverage policy used for release blocking

- **High criticality** modules require: `unit + integration` coverage at minimum.
- **Medium criticality** modules require: `unit` coverage at minimum.
- **Low criticality** modules require: at least one automated test (unit or integration).
- `e2e` and `perf` are mandatory only where noted in **Required test types**.

If a module is missing mandatory coverage, set **Release blocker** to `YES`.

## Matrix

| Area | Owner | Criticality | Existing tests | Missing tests | Required test types | Release blocker |
|---|---|---:|---|---|---|---|
| `src/data_access/` | Unassigned (`CODEOWNERS` not present) | High | `tests/test_data_access_api_client.py`, `tests/test_data_access_provider_contracts.py`, `tests/test_option_loader_contract.py`, `tests/test_storage_roundtrip.py`, `tests/test_normalization.py`, `tests/test_engine_loader.py`, `tests/test_provider_contract.py` | End-to-end path test from provider → normalization → cache/storage → consumer API; sustained-load perf test for provider/cache hot paths | unit, integration, perf | **YES** |
| `src/analysis/` | Unassigned (`CODEOWNERS` not present) | Medium | `tests/test_time_series_momentum.py`, `tests/test_factor_exposure_model.py`, `tests/test_attribution_pipeline.py`, `tests/test_explainability.py`, `tests/test_cross_asset_macro_integration.py`, `tests/test_research_lab_explainability_cards.py` | Integration tests for `analysis/reporting.py` artifact outputs and cross-module contract checks with `backtesting/` and `ui/` | unit, integration | NO |
| `src/backtesting/` | Unassigned (`CODEOWNERS` not present) | High | `tests/test_walk_forward.py`, `tests/test_backtest_signals.py`, `tests/test_chain_runner.py`, `tests/test_backtesting_perf_and_rolling.py`, `tests/test_vectorized_no_lookahead.py`, `tests/test_event_driven_position_apply_fill.py`, `tests/test_partial_fill_and_latency.py`, `tests/test_backtest_invariants_and_edge_cases.py`, `tests/test_backtest_sweep.py`, `tests/test_execution_models.py`, `tests/test_scenario_toolkit.py`, `tests/test_allocation_optimizer.py`, `tests/test_optimizer.py`, `tests/test_signal_diagnostics.py`, `tests/test_portfolio_construction.py` | Scenario-level e2e run validating full CLI/runner artifact set on representative multi-asset input; reproducible performance baseline check for vectorized vs event-driven paths | unit, integration, e2e, perf | **YES** |
| `src/modeling_nextgen/` | Unassigned (`CODEOWNERS` not present) | High | `tests/test_modeling_nextgen_inference_contract.py`, `tests/test_modeling_nextgen_interface_contracts.py`, `tests/test_modeling_nextgen_walkforward_hpo.py`, `tests/test_modeling_nextgen_purged_cv.py`, `tests/test_modeling_nextgen_adversarial_validation.py`, `tests/test_modeling_nextgen_cross_asset_graph.py`, `tests/test_modeling_nextgen_multitask.py`, `tests/test_modeling_nextgen_shadow_runner.py`, `tests/modeling_nextgen/test_nextgen_integration_regression.py`, `tests/test_meta_label_conformal.py`, `tests/test_panel_baselines.py`, `tests/test_sequence_encoder.py`, `tests/test_vol_factor_kalman.py`, `tests/test_no_arb.py`, `tests/test_vol_surface.py`, `tests/test_stress_scenarios.py`, `tests/test_probability_calibration.py`, `tests/test_bayesian_uncertainty_and_sizing_adapter.py`, `tests/test_nextgen_registry.py` | Dedicated e2e smoke covering `features -> models -> serving` with production-like schemas; targeted performance test for inference throughput/latency budgets | unit, integration, e2e, perf | **YES** |
| `src/models/` | Unassigned (`CODEOWNERS` not present) | Medium | `tests/test_model_registry_and_ensemble.py`, `tests/test_model_robustness_scorecards.py`, `tests/test_model_deployment_slots.py`, `tests/test_strategy_capacity.py` | Missing direct unit tests for `paradigms.py` and `base.py` contracts; add integration test to validate compatibility with `src/modeling_nextgen/adapters/registry_adapter.py` | unit, integration | NO |
| `src/ui/` | Unassigned (`CODEOWNERS` not present) | Medium | `tests/ui/test_helpers.py`, `tests/ui/test_state_contracts.py`, `tests/ui/test_ui_pages.py`, `tests/test_research_lab_presets.py`, `tests/test_research_lab_manifest.py` | Browser-level e2e workflow test (menu navigation + backtesting configuration + report rendering) | unit, integration, e2e | **YES** |
| `reports/` | Unassigned (`CODEOWNERS` not present) | High | `tests/test_quality_gate_artifacts.py`, `tests/test_benchmark_bundle.py`, `tests/test_governance_artifact_validator.py`, `tests/test_governance_drift_monitoring.py`, `tests/test_backtest_reporting_and_exports.py` | Performance regression check for nightly report generation time and artifact size growth; integration test for end-to-end `run_baselines.py` + validator chain | unit, integration, perf | **YES** |

## Maintenance workflow (required on every module change)

1. Update this matrix when:
   - a new module/package is introduced, or
   - required test types change, or
   - missing coverage is closed.
2. In the PR checklist, confirm this file was updated (or explicitly marked `N/A` with reason).
3. PRs adding a new module without a corresponding matrix row are considered **release blocked** until mapping is added.
