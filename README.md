# Outlook Skill

Python CLI for Outlook email and calendar via Microsoft Graph API.

## Features

- Read, search, send, reply, forward emails
- View, create, update, delete calendar events
- Multi-account support
- OAuth2 authentication via MSAL device code flow (no client secret needed)
- Rich terminal output with `--json` flag for programmatic use

## Quick Start

Prerequisites: Python 3.11+, [uv](https://docs.astral.sh/uv/)

```bash
# First-time setup (registers device code flow with Entra ID app)
uv run outlook.py setup

# Read your inbox
uv run outlook.py mail inbox

# Today's calendar
uv run outlook.py calendar today

# JSON output for agents
uv run outlook.py --json mail inbox
```

See [SKILL.md](./SKILL.md) for full documentation.

## Credits

This skill is a fork of [jotamed/outlook](https://clawhub.ai/jotamed/outlook) with additional improvements:

| Version | Change |
|---------|--------|
| v1.3.1 | **Security:** Fixed path traversal vulnerability in attachment download |
| v1.3.2 | **Fix:** Auto-detect system timezone (was hardcoded to Europe/Madrid) |
| v1.4.0 | **Feature:** Multi-account support (`--account` flag) |
| v2.0.0 | **Rewrite:** Complete Python CLI rewrite (replaces shell scripts) |

Thanks to [@jotamed](https://clawhub.ai/u/jotamed) for the original implementation.

## License

MIT
