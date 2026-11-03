# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.0] - 2026-11-03

### Added

- Phase 1: MCP server foundation with `create_presentation`,
  `open_presentation`, `save_presentation`, `save_as`, `list_slides` and
  `inspect_slide`.
- Phase 2: `add_slide`, `duplicate_slide`, `delete_slide` and `move_slide`.
- Phase 3: `set_slide_text`, `add_text_box`, `replace_text` and `add_image`
  (local path or http(s) URL).
- 50 automated tests.
- Functional and technical documentation, plus beginner-friendly guides for
  phases 1–3.
- E2E sample generator and results (validated on PowerPoint 2016 desktop).
