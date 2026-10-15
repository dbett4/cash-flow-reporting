PYTHON ?= .venv/bin/python

.PHONY: build demo test fixtures
build:
	cargo build --release --locked --manifest-path tools/report-compiler/Cargo.toml

demo: build
	$(PYTHON) scripts/demo.py

test: build
	cargo test --release --locked --manifest-path tools/report-compiler/Cargo.toml
	$(PYTHON) -m unittest discover -s tests -v

fixtures:
	$(PYTHON) scripts/author_fixture.py
