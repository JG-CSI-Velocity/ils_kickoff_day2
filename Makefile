.PHONY: setup run test lint fmt clean

setup:
	python -m venv .venv
	.venv/bin/pip install --upgrade pip
	.venv/bin/pip install -r requirements.txt

setup-dev:
	python -m venv .venv
	.venv/bin/pip install --upgrade pip
	.venv/bin/pip install -r requirements-dev.txt

run:
	.venv/bin/python -m ils_kickoff

test:
	.venv/bin/pytest

lint:
	.venv/bin/ruff check ils_kickoff/ tests/

fmt:
	.venv/bin/ruff format ils_kickoff/ tests/

clean:
	rm -rf __pycache__ .pytest_cache .ruff_cache
	find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
