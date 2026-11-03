# wolf-ms-ppt-mcp

MCP server that lets AI agents create and manipulate Microsoft PowerPoint
presentations (`.pptx`) locally, **without requiring Microsoft Office**.

[![CI](https://github.com/amareca/wolf-ms-ppt-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/amareca/wolf-ms-ppt-mcp/actions/workflows/ci.yml)

## Status

**Phases 1, 2 and 3** are implemented.

**Phase 1 — Foundation and reading:**

- `create_presentation` — create a new deck (16:9 by default) or from a template.
- `open_presentation` — open an existing `.pptx`.
- `save_presentation` / `save_as` — persist changes.
- `list_slides` — list slides with layout, title and text preview.
- `inspect_slide` — detailed shapes/text/images/positions of one slide.

**Phase 2 — Slide management:**

- `add_slide` — add a slide using a layout, optionally at a given position.
- `duplicate_slide` — duplicate a slide (shapes, background and images).
- `delete_slide` — delete a slide.
- `move_slide` — reorder slides.

**Phase 3 — Content (text and images):**

- `set_slide_text` — set slide titles, subtitles, bodies or named shapes, with
  optional color.
- `add_text_box` — add a text box at a position (inches), with optional color.
- `replace_text` — find/replace text across all slides or one slide.
- `add_image` — insert an image from a local path or an http(s) URL.

All planned MVP phases are implemented. Future work includes verifying
PowerPoint for the Web and the post-MVP versions (COM backend, advanced
content, cloud).

## Requirements

- Python 3.11+
- No Microsoft Office installation required

## Installation

```bash
uv sync
```

Or install from source:

```bash
pip install .
```

## Usage

Run the server over stdio:

```bash
uv run wolf-ms-ppt-mcp
```

Configure your MCP client (Claude, Cursor, etc.):

```json
{
  "mcpServers": {
    "wolf-ms-ppt": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/wolf-ms-ppt-mcp",
        "run",
        "wolf-ms-ppt-mcp"
      ]
    }
  }
}
```

Once published to PyPI, this can be simplified to `uvx wolf-ms-ppt-mcp`.

## Development

```bash
uv sync --extra dev
uv run pytest
```

## Contributing

See [`CONTRIBUTING.md`](CONTRIBUTING.md).

## Security

See [`SECURITY.md`](SECURITY.md).

## Changelog

See [`CHANGELOG.md`](CHANGELOG.md).

## License

MIT. See [`LICENSE`](LICENSE). Third-party attributions are listed in
[`THIRD_PARTY_LICENSES.md`](THIRD_PARTY_LICENSES.md).

## Disclaimer

PowerPoint and Microsoft Office are trademarks of Microsoft Corporation.
This project is not affiliated with or endorsed by Microsoft.
