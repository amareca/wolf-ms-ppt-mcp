# Contributing

Thanks for your interest in improving **wolf-ms-ppt-mcp**.

## Requirements

- Python 3.11+
- [uv](https://docs.astral.sh/uv/)

## Setup

```bash
uv sync --extra dev
```

## Tests

```bash
uv run pytest -q
```

Please add tests for new behavior and keep the existing suite green.

## Project layout

- `src/wolf_ms_ppt_mcp/` — MCP server and slide/text/image operations.
- `tests/` — automated tests.
- `docs/` — functional, technical and phase documentation.
- `scripts/` — development helpers (E2E sample generator).

## Guidelines

- Keep changes focused and consistent with the existing style.
- Add type hints to public functions.
- Update the docs and `CHANGELOG.md` when behavior changes.
- Do not commit generated artifacts (`samples/`, `.venv/`, caches).

## Reporting issues

Open an issue with:

- What you expected to happen.
- What actually happened.
- Steps to reproduce (including a minimal `.pptx` or tool arguments if relevant).
- Your OS, Python version and PowerPoint version (if applicable).
