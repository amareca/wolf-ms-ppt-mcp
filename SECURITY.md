# Security Policy

## Supported versions

This project is at an early stage; security fixes are applied to the latest
release.

## Reporting a vulnerability

Please **do not** open a public issue for security problems. Instead, use
GitHub's private security advisory feature (**Security** tab → **Report a
vulnerability**) or contact the maintainer directly.

Include:

- A description of the issue.
- Steps to reproduce.
- Potential impact.

## Scope

`wolf-ms-ppt-mcp` runs locally and operates on local `.pptx` files. Its only
network access is downloading images from `http(s)` URLs passed to
`add_image`. It does not execute macros or arbitrary code.
