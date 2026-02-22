# Outlook Skill

CLI for Outlook email and calendar via Microsoft Graph API.

## Features

- 📧 Read, search, send, reply, forward emails
- 📅 View, create, update, delete calendar events
- 👥 Multi-account support
- 🔐 OAuth2 authentication via Azure

## Credits

This skill is a fork of [jotamed/outlook](https://clawhub.ai/jotamed/outlook) with additional improvements:

| Version | Change |
|---------|--------|
| v1.3.1 | **Security:** Fixed path traversal vulnerability in attachment download |
| v1.3.2 | **Fix:** Auto-detect system timezone (was hardcoded to Europe/Madrid) |
| v1.4.0 | **Feature:** Multi-account support (`--account` flag) |

Thanks to [@jotamed](https://clawhub.ai/u/jotamed) for the original implementation.

## Usage

See [SKILL.md](./SKILL.md) for full documentation.

## License

MIT
