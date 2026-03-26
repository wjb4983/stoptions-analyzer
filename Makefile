.PHONY: setup lint unit-test smoke-run test-matrix

setup:
	python -m pip install --upgrade pip
	python -m pip install -r requirements.txt

lint:
	python -m compileall -q src tests

unit-test:
	pytest -q tests/snn_bench

smoke-run:
	PYTHONPATH=src python -m snn_bench.scripts.smoke_pipeline --ticker NVDA --timeframe 1D

test-matrix:
	./scripts/run_test_matrix.sh
