#!/usr/bin/env -S uv run
# /// script
# requires-python = ">=3.11"
# dependencies = [
#     "typer[all]>=0.12",
#     "msgraph-sdk>=1.50",
#     "msal>=1.28",
#     "rich>=13.0",
#     "azure-core>=1.30",
# ]
# ///
"""Outlook CLI — Python rewrite of the Outlook shell scripts.

Provides mail, calendar, and token management via Microsoft Graph API.
"""

from __future__ import annotations

import asyncio
import html
import json
import os
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, NoReturn

import msal
import typer
from azure.core.credentials import AccessToken, TokenCredential
from msgraph import GraphServiceClient
from rich.console import Console
from rich.panel import Panel
from rich.table import Table

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

BASE_DIR = Path.home() / ".outlook-mcp"
ACCOUNT_PATTERN = re.compile(r"^[a-zA-Z0-9_-]+$")
TZ_PATTERN = re.compile(r"^[a-zA-Z0-9/_+-]+$")
SCOPES = [
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "User.Read",
    "offline_access",
]

# Rich output on stderr so --json on stdout stays clean
console = Console(stderr=True)

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------


@dataclass
class State:
    account: str = "default"
    json_mode: bool = False
    debug: bool = False


state = State()

# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------


class AuthError(Exception):
    """Raised when authentication fails."""


# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------


def validate_account(name: str) -> str:
    """Validate an account name. Raises SystemExit on invalid input."""
    if not ACCOUNT_PATTERN.match(name):
        _error_exit(
            f"Invalid account name '{name}'. Use only letters, numbers, hyphens, and underscores."
        )
    return name


def _account_dir(account: str) -> Path:
    validate_account(account)
    return BASE_DIR / account


def load_config(account: str) -> dict:
    """Read ~/.outlook-mcp/<account>/config.json and return its contents."""
    config_path = _account_dir(account) / "config.json"
    if not config_path.exists():
        _error_exit(
            f"Account '{account}' not configured. Run: uv run outlook.py setup --account {account}"
        )
    try:
        return json.loads(config_path.read_text())
    except (json.JSONDecodeError, OSError) as exc:
        _error_exit(f"Failed to read config for account '{account}': {exc}")


# ---------------------------------------------------------------------------
# Error / output helpers
# ---------------------------------------------------------------------------


def _error_exit(msg: str, code: int = 1) -> NoReturn:
    """Print an error and exit."""
    if state.json_mode:
        console.print_json(json.dumps({"error": msg}))
    else:
        console.print(Panel(f"[bold red]Error:[/bold red] {msg}", border_style="red"))
    raise SystemExit(code)


def output_table(title: str, columns: list[tuple[str, str]], rows: list[dict]) -> None:
    """Render a Rich table or JSON list depending on state.json_mode."""
    if state.json_mode:
        print(json.dumps(rows, indent=2, default=str))
        return
    table = Table(title=title, show_lines=False)
    for col_key, col_label in columns:
        table.add_column(col_label)
    for row in rows:
        table.add_row(*(str(row.get(k, "")) for k, _ in columns))
    console.print(table)


def output_detail(data: dict) -> None:
    """Render a Rich panel or JSON object."""
    if state.json_mode:
        print(json.dumps(data, indent=2, default=str))
        return
    lines = [f"[bold]{k}:[/bold] {v}" for k, v in data.items()]
    console.print(Panel("\n".join(lines)))


def output_status(data: dict) -> None:
    """Render a success/error status."""
    if state.json_mode:
        print(json.dumps(data, indent=2, default=str))
        return
    status = data.get("status", "info")
    msg = data.get("message", "")
    style = "green" if status == "ok" else "red" if status == "error" else "blue"
    console.print(Panel(f"[bold {style}]{msg}[/bold {style}]", border_style=style))


# ---------------------------------------------------------------------------
# HTML stripping
# ---------------------------------------------------------------------------


def strip_html(html_str: str) -> str:
    """Convert HTML to plain text — strip tags, decode entities, collapse whitespace."""
    if not html_str:
        return ""
    # Remove style/script blocks
    text = re.sub(r"<(style|script)[^>]*>.*?</\1>", "", html_str, flags=re.DOTALL | re.IGNORECASE)
    # Replace <br> / <p> / <div> with newlines
    text = re.sub(r"<br\s*/?>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</(p|div|tr|li)>", "\n", text, flags=re.IGNORECASE)
    # Strip remaining tags
    text = re.sub(r"<[^>]+>", "", text)
    # Decode HTML entities
    text = html.unescape(text)
    # Collapse whitespace (preserve newlines)
    text = re.sub(r"[^\S\n]+", " ", text)
    # Collapse multiple blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# ---------------------------------------------------------------------------
# Timezone detection
# ---------------------------------------------------------------------------


def detect_timezone() -> str:
    """Detect the system timezone.

    Priority: OUTLOOK_TZ env var -> macOS readlink /etc/localtime -> /etc/timezone -> UTC fallback.
    """
    # 1. Environment variable
    tz = os.environ.get("OUTLOOK_TZ", "").strip()
    if tz and TZ_PATTERN.match(tz):
        return tz

    # 2. macOS: readlink /etc/localtime -> .../zoneinfo/America/New_York
    try:
        link = os.readlink("/etc/localtime")
        # e.g. /var/db/timezone/zoneinfo/America/New_York
        if "zoneinfo/" in link:
            tz = link.split("zoneinfo/", 1)[1]
            if TZ_PATTERN.match(tz):
                return tz
    except OSError:
        pass

    # 3. /etc/timezone (Linux)
    tz_file = Path("/etc/timezone")
    if tz_file.exists():
        tz = tz_file.read_text().strip()
        if TZ_PATTERN.match(tz):
            return tz

    # 4. Fallback
    return "UTC"


# ---------------------------------------------------------------------------
# MSAL credential bridge
# ---------------------------------------------------------------------------


class MsalTokenCredential(TokenCredential):
    """Bridges MSAL token cache -> azure.core.credentials for GraphServiceClient."""

    def __init__(self, client_id: str, cache_path: Path):
        self._cache = msal.SerializableTokenCache()
        if cache_path.exists():
            self._cache.deserialize(cache_path.read_text())
        self._app = msal.PublicClientApplication(
            client_id,
            authority="https://login.microsoftonline.com/common",
            token_cache=self._cache,
        )
        self._cache_path = cache_path

    def get_token(
        self,
        *scopes: str,
        claims: str | None = None,
        tenant_id: str | None = None,
        enable_cae: bool = False,
        **kwargs: Any,
    ) -> AccessToken:
        accounts = self._app.get_accounts()
        if not accounts:
            raise AuthError("No cached account. Run: uv run outlook.py setup")
        result = self._app.acquire_token_silent(list(scopes), account=accounts[0])
        if not result or "access_token" not in result:
            raise AuthError("Token refresh failed. Run: uv run outlook.py setup")
        self._save_cache()
        return AccessToken(
            result["access_token"],
            int(time.time()) + result.get("expires_in", 3600),
        )

    def _save_cache(self) -> None:
        if self._cache.has_state_changed:
            self._cache_path.write_text(self._cache.serialize())
            self._cache_path.chmod(0o600)


# ---------------------------------------------------------------------------
# Graph client factory
# ---------------------------------------------------------------------------


def get_graph_client(account: str | None = None) -> GraphServiceClient:
    """Create a GraphServiceClient using the MSAL token cache for the given account."""
    acct = account or state.account
    config = load_config(acct)
    client_id = config.get("client_id")
    if not client_id:
        _error_exit(f"No client_id in config for account '{acct}'.")

    cache_path = _account_dir(acct) / "token_cache.json"
    credential = MsalTokenCredential(client_id, cache_path)

    try:
        return GraphServiceClient(credentials=credential)
    except AuthError as exc:
        _error_exit(str(exc))


# ---------------------------------------------------------------------------
# ID resolution helpers
# ---------------------------------------------------------------------------


async def _resolve_message_id(client: GraphServiceClient, partial_id: str) -> str:
    """Fetch recent 100 messages and find one whose ID ends with partial_id."""
    from msgraph.generated.users.item.messages.messages_request_builder import (
        MessagesRequestBuilder,
    )

    query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
        top=100,
        select=["id"],
        orderby=["receivedDateTime desc"],
    )
    config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
        query_parameters=query,
    )
    result = await client.me.messages.get(request_configuration=config)
    if result and result.value:
        for msg in result.value:
            if msg.id and msg.id.endswith(partial_id):
                return msg.id
    raise ValueError(f"No message found ending with '{partial_id}'")


async def _resolve_event_id(client: GraphServiceClient, partial_id: str) -> str:
    """Fetch recent 100 events and find one whose ID ends with partial_id."""
    from msgraph.generated.users.item.events.events_request_builder import (
        EventsRequestBuilder,
    )

    query = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters(
        top=100,
        select=["id"],
        orderby=["start/dateTime desc"],
    )
    config = EventsRequestBuilder.EventsRequestBuilderGetRequestConfiguration(
        query_parameters=query,
    )
    result = await client.me.events.get(request_configuration=config)
    if result and result.value:
        for event in result.value:
            if event.id and event.id.endswith(partial_id):
                return event.id
    raise ValueError(f"No event found ending with '{partial_id}'")


def resolve_message_id(client: GraphServiceClient, partial_id: str) -> str:
    """Sync wrapper around async message ID resolution."""
    return asyncio.run(_resolve_message_id(client, partial_id))


def resolve_event_id(client: GraphServiceClient, partial_id: str) -> str:
    """Sync wrapper around async event ID resolution."""
    return asyncio.run(_resolve_event_id(client, partial_id))


# ---------------------------------------------------------------------------
# Typer app
# ---------------------------------------------------------------------------

app = typer.Typer(
    name="outlook",
    help="Outlook CLI — manage mail, calendar, and tokens via Microsoft Graph.",
    no_args_is_help=True,
)
mail_app = typer.Typer(help="Mail commands.", no_args_is_help=True)
calendar_app = typer.Typer(help="Calendar commands.", no_args_is_help=True)
token_app = typer.Typer(help="Token management commands.", no_args_is_help=True)

app.add_typer(mail_app, name="mail")
app.add_typer(calendar_app, name="calendar")
app.add_typer(token_app, name="token")


@app.callback()
def main(
    json_output: bool = typer.Option(False, "--json", help="Output JSON instead of Rich tables."),
    account: str = typer.Option(
        None,
        "--account",
        "-a",
        help="Account name (default: OUTLOOK_ACCOUNT env var or 'default').",
    ),
    debug: bool = typer.Option(False, "--debug", help="Enable debug output."),
) -> None:
    """Outlook CLI — manage mail, calendar, and tokens via Microsoft Graph."""
    state.json_mode = json_output
    state.debug = debug
    acct = account or os.environ.get("OUTLOOK_ACCOUNT", "default")
    state.account = validate_account(acct)


# ---------------------------------------------------------------------------
# setup command
# ---------------------------------------------------------------------------


@app.command()
def setup() -> None:
    """Authenticate with Microsoft Graph using device code flow."""
    acct = state.account

    console.print("\n[bold blue]=== Outlook Setup ===[/bold blue]")
    console.print(f"Account: [green]{acct}[/green]\n")

    # Prompt for client_id
    acct_dir = _account_dir(acct)
    config_path = acct_dir / "config.json"

    existing_client_id = ""
    if config_path.exists():
        try:
            existing = json.loads(config_path.read_text())
            existing_client_id = existing.get("client_id", "")
        except (json.JSONDecodeError, OSError):
            pass

    if existing_client_id:
        console.print(f"Existing client_id: [cyan]{existing_client_id}[/cyan]")
        client_id = (
            typer.prompt(
                "Client ID (press Enter to keep existing)",
                default=existing_client_id,
            )
            .strip()
        )
    else:
        console.print(
            "You need an App Registration in Azure Entra ID.\n"
            "Create one at: https://entra.microsoft.com → App registrations → New registration\n"
            "  - Supported account types: Personal Microsoft accounts only (or multi-tenant)\n"
            "  - Redirect URI: http://localhost (Mobile and desktop applications)\n"
            "  - Under API permissions, add: Mail.ReadWrite, Mail.Send, Calendars.ReadWrite, User.Read\n"
        )
        client_id = typer.prompt("Client ID (from Azure Entra ID)").strip()

    if not client_id:
        _error_exit("Client ID is required.")

    # Create config directory
    acct_dir.mkdir(parents=True, exist_ok=True)

    # Save config.json
    config_data: dict[str, Any] = {"client_id": client_id}
    # Preserve client_secret if it exists (backward compat)
    if config_path.exists():
        try:
            old = json.loads(config_path.read_text())
            if "client_secret" in old:
                config_data["client_secret"] = old["client_secret"]
        except (json.JSONDecodeError, OSError):
            pass

    config_path.write_text(json.dumps(config_data, indent=2) + "\n")
    config_path.chmod(0o600)

    # Device code flow
    console.print("\n[bold yellow]Starting device code authentication...[/bold yellow]\n")

    cache = msal.SerializableTokenCache()
    cache_path = acct_dir / "token_cache.json"
    if cache_path.exists():
        cache.deserialize(cache_path.read_text())

    msal_app = msal.PublicClientApplication(
        client_id,
        authority="https://login.microsoftonline.com/common",
        token_cache=cache,
    )

    flow = msal_app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        _error_exit(f"Device flow initiation failed: {flow.get('error_description', 'unknown error')}")

    console.print("[bold]To sign in:[/bold]")
    console.print(f"  1. Open: [cyan]{flow['verification_uri']}[/cyan]")
    console.print(f"  2. Enter code: [bold yellow]{flow['user_code']}[/bold yellow]")
    console.print("\nWaiting for authentication...")

    result = msal_app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        error_desc = result.get("error_description", result.get("error", "Unknown error"))
        _error_exit(f"Authentication failed: {error_desc}")

    # Save token cache
    cache_path.write_text(cache.serialize())
    cache_path.chmod(0o600)

    console.print("\n[bold green]=== Setup Complete ===[/bold green]")
    console.print(f"Config saved to: {config_path}")
    console.print(f"Token cache saved to: {cache_path}")
    console.print(f"\nTest with: [cyan]uv run outlook.py token test --account {acct}[/cyan]")


# ---------------------------------------------------------------------------
# token test command
# ---------------------------------------------------------------------------


@token_app.command()
def test() -> None:
    """Test the connection by calling the /me endpoint."""
    acct = state.account

    try:
        client = get_graph_client(acct)
    except (SystemExit, AuthError) as exc:
        _error_exit(f"Failed to create Graph client: {exc}")

    async def _test() -> dict:
        me = await client.me.get()
        return {
            "displayName": me.display_name if me else "Unknown",
            "mail": me.mail or me.user_principal_name if me else "Unknown",
        }

    try:
        info = asyncio.run(_test())
        output_status(
            {
                "status": "ok",
                "message": f"Connected as {info['displayName']} ({info['mail']})",
            }
        )
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Connection test failed: {exc}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app()
