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
    msg = data.get("message", "")
    if not msg:
        parts = [f"[bold]{k}:[/bold] {v}" for k, v in data.items()]
        msg = "\n".join(parts)
    status_val = str(data.get("status", "info"))
    style = "red" if "error" in status_val or "fail" in status_val else "green"
    console.print(Panel(msg, border_style=style))


# ---------------------------------------------------------------------------
# Shared message list helpers
# ---------------------------------------------------------------------------

MSG_COLUMNS = [("n", "#"), ("subject", "Subject"), ("from", "From"), ("date", "Date"), ("id", "ID")]
MSG_COLUMNS_WITH_READ = [("n", "#"), ("subject", "Subject"), ("from", "From"), ("date", "Date"), ("read", "Read"), ("id", "ID")]


def _format_message_rows(messages, include_read: bool = False) -> list[dict]:
    """Format Graph SDK message objects into output rows."""
    rows = []
    for i, msg in enumerate(messages, 1):
        row = {
            "n": i,
            "subject": msg.subject or "(no subject)",
            "from": msg.from_.email_address.address if msg.from_ and msg.from_.email_address else "",
            "date": str(msg.received_date_time)[:16] if msg.received_date_time else "",
            "id": (msg.id or "")[-20:],
        }
        if include_read:
            row["read"] = msg.is_read
        rows.append(row)
    return rows


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
        force_refresh: bool = False,
        **kwargs: Any,
    ) -> AccessToken:
        accounts = self._app.get_accounts()
        if not accounts:
            raise AuthError("No cached account. Run: uv run outlook.py setup")
        result = self._app.acquire_token_silent(list(scopes), account=accounts[0], force_refresh=force_refresh)
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
# Mail commands
# ---------------------------------------------------------------------------


@mail_app.command()
def inbox(
    count: int = typer.Option(10, "--count", "-n", help="Number of messages"),
) -> None:
    """List inbox messages."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
        query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=count,
            select=["id", "subject", "from", "receivedDateTime", "isRead"],
            orderby=["receivedDateTime desc"],
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query,
        )
        result = await client.me.messages.get(request_configuration=config)
        return result.value or []

    try:
        messages = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch inbox: {exc}")

    output_table("Inbox", MSG_COLUMNS_WITH_READ, _format_message_rows(messages, include_read=True))


@mail_app.command()
def unread(
    count: int = typer.Option(20, "--count", "-n", help="Number of messages"),
) -> None:
    """List unread messages."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
        query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=count,
            select=["id", "subject", "from", "receivedDateTime"],
            orderby=["receivedDateTime desc"],
            filter="isRead eq false",
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query,
        )
        result = await client.me.messages.get(request_configuration=config)
        return result.value or []

    try:
        messages = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch unread messages: {exc}")

    output_table("Unread", MSG_COLUMNS, _format_message_rows(messages))


@mail_app.command()
def search(
    query: str = typer.Argument(..., help="Search query"),
    count: int = typer.Option(20, "--count", "-n", help="Number of messages"),
) -> None:
    """Search messages."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
        params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=count,
            select=["id", "subject", "from", "receivedDateTime"],
            search=query,
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=params,
        )
        result = await client.me.messages.get(request_configuration=config)
        return result.value or []

    try:
        messages = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to search messages: {exc}")

    output_table("Search Results", MSG_COLUMNS, _format_message_rows(messages))


@mail_app.command("from")
def from_(
    email: str = typer.Argument(..., help="Email address to search"),
    count: int = typer.Option(20, "--count", "-n", help="Number of messages"),
) -> None:
    """List messages from a specific sender."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
        params = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=count,
            select=["id", "subject", "from", "receivedDateTime", "isRead"],
            search=f"from:{email}",
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=params,
        )
        result = await client.me.messages.get(request_configuration=config)
        return result.value or []

    try:
        messages = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch messages from {email}: {exc}")

    output_table(f"From {email}", MSG_COLUMNS_WITH_READ, _format_message_rows(messages, include_read=True))


@mail_app.command("read")
def read_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """Read a specific message."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.users.item.messages.item.message_item_request_builder import MessageItemRequestBuilder
        params = MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters(
            select=["id", "subject", "from", "receivedDateTime", "body", "toRecipients"],
        )
        config = MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration(
            query_parameters=params,
        )
        return await client.me.messages.by_message_id(full_id).get(request_configuration=config)

    try:
        msg = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to read message: {exc}")

    if not msg:
        _error_exit("Message not found.")

    from_name = ""
    from_addr = ""
    if msg.from_ and msg.from_.email_address:
        from_name = msg.from_.email_address.name or ""
        from_addr = msg.from_.email_address.address or ""

    to_list = []
    if msg.to_recipients:
        for r in msg.to_recipients:
            if r.email_address and r.email_address.address:
                to_list.append(r.email_address.address)

    body_text = ""
    if msg.body and msg.body.content:
        body_text = strip_html(msg.body.content)[:2000]

    if state.json_mode:
        data = {
            "subject": msg.subject or "(no subject)",
            "from": {"name": from_name, "address": from_addr},
            "to": to_list,
            "date": str(msg.received_date_time)[:16] if msg.received_date_time else "",
            "body": body_text,
        }
        print(json.dumps(data, indent=2, default=str))
    else:
        data = {
            "subject": msg.subject or "(no subject)",
            "from": f"{from_name} <{from_addr}>" if from_name else from_addr,
            "to": ", ".join(to_list),
            "date": str(msg.received_date_time)[:16] if msg.received_date_time else "",
            "body": body_text,
        }
        output_detail(data)


@mail_app.command()
def attachments(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """List attachments for a message."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        result = await client.me.messages.by_message_id(full_id).attachments.get()
        return result.value or []

    try:
        atts = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch attachments: {exc}")

    rows = []
    for att in atts:
        rows.append({
            "name": att.name or "",
            "size": att.size or 0,
            "contentType": att.content_type or "",
            "id": att.id or "",
        })

    output_table("Attachments", [("name", "Name"), ("size", "Size"), ("contentType", "Content Type"), ("id", "ID")], rows)


@mail_app.command()
def focused(
    count: int = typer.Option(10, "--count", "-n", help="Number of messages"),
) -> None:
    """List focused inbox messages."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
        query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=count,
            select=["id", "subject", "from", "receivedDateTime"],
            orderby=["receivedDateTime desc"],
            filter="inferenceClassification eq 'focused'",
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query,
        )
        result = await client.me.messages.get(request_configuration=config)
        return result.value or []

    try:
        messages = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch focused messages: {exc}")

    output_table("Focused", MSG_COLUMNS, _format_message_rows(messages))


@mail_app.command()
def other(
    count: int = typer.Option(10, "--count", "-n", help="Number of messages"),
) -> None:
    """List 'other' inbox messages."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
        query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=count,
            select=["id", "subject", "from", "receivedDateTime"],
            orderby=["receivedDateTime desc"],
            filter="inferenceClassification eq 'other'",
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query,
        )
        result = await client.me.messages.get(request_configuration=config)
        return result.value or []

    try:
        messages = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch other messages: {exc}")

    output_table("Other", MSG_COLUMNS, _format_message_rows(messages))


@mail_app.command()
def thread(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """List all messages in the same conversation thread."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.users.item.messages.item.message_item_request_builder import MessageItemRequestBuilder
        from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder

        # First get the message to find its conversationId
        params = MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters(
            select=["id", "conversationId"],
        )
        msg_config = MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration(
            query_parameters=params,
        )
        msg = await client.me.messages.by_message_id(full_id).get(request_configuration=msg_config)
        if not msg or not msg.conversation_id:
            raise ValueError("Could not retrieve conversation ID for message.")

        # Now fetch all messages with same conversationId
        query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=50,
            select=["id", "subject", "from", "receivedDateTime"],
            orderby=["receivedDateTime asc"],
            filter=f"conversationId eq '{msg.conversation_id}'",
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query,
        )
        result = await client.me.messages.get(request_configuration=config)
        return result.value or []

    try:
        messages = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch thread: {exc}")

    output_table("Thread", MSG_COLUMNS, _format_message_rows(messages))


@mail_app.command()
def folders() -> None:
    """List mail folders."""
    client = get_graph_client()

    async def _run():
        result = await client.me.mail_folders.get()
        return result.value or []

    try:
        folder_list = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch mail folders: {exc}")

    rows = []
    for f in folder_list:
        rows.append({
            "name": f.display_name or "",
            "total": f.total_item_count or 0,
            "unread": f.unread_item_count or 0,
        })

    output_table("Mail Folders", [("name", "Name"), ("total", "Total"), ("unread", "Unread")], rows)


@mail_app.command()
def stats() -> None:
    """Show inbox statistics."""
    client = get_graph_client()

    async def _run():
        return await client.me.mail_folders.by_mail_folder_id("inbox").get()

    try:
        folder = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch inbox stats: {exc}")

    if not folder:
        _error_exit("Could not retrieve inbox folder.")

    data = {
        "folder": folder.display_name or "Inbox",
        "total": folder.total_item_count or 0,
        "unread": folder.unread_item_count or 0,
    }
    output_detail(data)


@mail_app.command()
def categories() -> None:
    """List master categories."""
    client = get_graph_client()

    async def _run():
        result = await client.me.outlook.master_categories.get()
        return result.value or []

    try:
        cats = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch categories: {exc}")

    rows = []
    for cat in cats:
        rows.append({
            "name": cat.display_name or "",
            "color": cat.color or "",
            "id": (cat.id or "")[:8],
        })

    output_table("Categories", [("name", "Name"), ("color", "Color"), ("id", "ID")], rows)


# ---------------------------------------------------------------------------
# Mail action commands — management
# ---------------------------------------------------------------------------


@mail_app.command("mark-read")
def mark_read(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """Mark a message as read."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.models.message import Message as MsgModel
        body = MsgModel()
        body.is_read = True
        await client.me.messages.by_message_id(full_id).patch(body)
        return full_id

    try:
        full_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to mark message as read: {exc}")

    output_status({"status": "marked as read", "subject": "...", "id": full_id[-20:]})


@mail_app.command("mark-unread")
def mark_unread(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """Mark a message as unread."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.models.message import Message as MsgModel
        body = MsgModel()
        body.is_read = False
        await client.me.messages.by_message_id(full_id).patch(body)
        return full_id

    try:
        full_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to mark message as unread: {exc}")

    output_status({"status": "marked as unread", "subject": "...", "id": full_id[-20:]})


@mail_app.command("flag")
def flag_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """Flag a message."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.models.message import Message as MsgModel
        from msgraph.generated.models.followup_flag import FollowupFlag
        from msgraph.generated.models.followup_flag_status import FollowupFlagStatus
        body = MsgModel()
        body.flag = FollowupFlag(flag_status=FollowupFlagStatus.Flagged)
        await client.me.messages.by_message_id(full_id).patch(body)
        return full_id

    try:
        full_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to flag message: {exc}")

    output_status({"status": "flagged", "subject": "...", "id": full_id[-20:]})


@mail_app.command("unflag")
def unflag_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """Unflag a message."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.models.message import Message as MsgModel
        from msgraph.generated.models.followup_flag import FollowupFlag
        from msgraph.generated.models.followup_flag_status import FollowupFlagStatus
        body = MsgModel()
        body.flag = FollowupFlag(flag_status=FollowupFlagStatus.NotFlagged)
        await client.me.messages.by_message_id(full_id).patch(body)
        return full_id

    try:
        full_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to unflag message: {exc}")

    output_status({"status": "unflagged", "subject": "...", "id": full_id[-20:]})


@mail_app.command("delete")
def delete_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """Move a message to trash."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.users.item.messages.item.move.move_post_request_body import MovePostRequestBody
        body = MovePostRequestBody()
        body.destination_id = "deleteditems"
        await client.me.messages.by_message_id(full_id).move.post(body)
        return full_id

    try:
        full_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to delete message: {exc}")

    output_status({"status": "moved to trash", "subject": "...", "id": full_id[-20:]})


@mail_app.command("archive")
def archive_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """Move a message to archive."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.users.item.messages.item.move.move_post_request_body import MovePostRequestBody
        body = MovePostRequestBody()
        body.destination_id = "archive"
        await client.me.messages.by_message_id(full_id).move.post(body)
        return full_id

    try:
        full_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to archive message: {exc}")

    output_status({"status": "archived", "subject": "...", "id": full_id[-20:]})


@mail_app.command("move")
def move_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
    folder: str = typer.Argument(..., help="Destination folder name"),
) -> None:
    """Move a message to a specific folder."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        # Look up folder by name (case-insensitive)
        result = await client.me.mail_folders.get()
        folder_id = None
        folder_name = None
        if result and result.value:
            for f in result.value:
                if f.display_name and f.display_name.lower() == folder.lower():
                    folder_id = f.id
                    folder_name = f.display_name
                    break
        if not folder_id:
            raise ValueError(f"Folder '{folder}' not found.")
        from msgraph.generated.users.item.messages.item.move.move_post_request_body import MovePostRequestBody
        body = MovePostRequestBody()
        body.destination_id = folder_id
        await client.me.messages.by_message_id(full_id).move.post(body)
        return full_id, folder_name

    try:
        full_id, folder_name = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to move message: {exc}")

    output_status({"status": "moved", "folder": folder_name, "subject": "...", "id": full_id[-20:]})


# ---------------------------------------------------------------------------
# Mail action commands — sending
# ---------------------------------------------------------------------------


@mail_app.command("send")
def send_msg(
    to: str = typer.Argument(..., help="Recipient email address"),
    subject: str = typer.Argument(..., help="Email subject"),
    body: str = typer.Argument("", help="Email body text"),
) -> None:
    """Send a new email."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.send_mail.send_mail_post_request_body import SendMailPostRequestBody
        from msgraph.generated.models.message import Message as MsgModel
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.models.email_address import EmailAddress

        msg = MsgModel()
        msg.subject = subject
        msg.body = ItemBody(content_type=BodyType.Text, content=body)
        msg.to_recipients = [Recipient(email_address=EmailAddress(address=to))]

        send_body = SendMailPostRequestBody(message=msg)
        await client.me.send_mail.post(send_body)

    try:
        asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to send message: {exc}")

    output_status({"status": "sent", "to": to, "subject": subject})


@mail_app.command("reply")
def reply_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
    body: str = typer.Argument(..., help="Reply body text"),
) -> None:
    """Reply to a message."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.users.item.messages.item.reply.reply_post_request_body import ReplyPostRequestBody
        reply_body = ReplyPostRequestBody()
        reply_body.comment = body
        await client.me.messages.by_message_id(full_id).reply.post(reply_body)
        return full_id

    try:
        full_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to reply to message: {exc}")

    output_status({"status": "reply sent", "id": full_id[-20:]})


@mail_app.command("forward")
def forward_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
    to: str = typer.Argument(..., help="Recipient email address"),
    comment: str = typer.Argument("", help="Optional comment"),
) -> None:
    """Forward a message."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.users.item.messages.item.forward.forward_post_request_body import ForwardPostRequestBody
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.models.email_address import EmailAddress
        fwd_body = ForwardPostRequestBody()
        fwd_body.comment = comment
        fwd_body.to_recipients = [Recipient(email_address=EmailAddress(address=to))]
        await client.me.messages.by_message_id(full_id).forward.post(fwd_body)
        return full_id

    try:
        full_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to forward message: {exc}")

    output_status({"status": "forwarded", "to": to, "id": full_id[-20:]})


@mail_app.command("draft")
def create_draft(
    to: str = typer.Argument(..., help="Recipient email address"),
    subject: str = typer.Argument(..., help="Email subject"),
    body: str = typer.Argument("", help="Email body text"),
) -> None:
    """Create a draft message."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.models.message import Message as MsgModel
        from msgraph.generated.models.item_body import ItemBody
        from msgraph.generated.models.body_type import BodyType
        from msgraph.generated.models.recipient import Recipient
        from msgraph.generated.models.email_address import EmailAddress

        msg = MsgModel()
        msg.subject = subject
        msg.body = ItemBody(content_type=BodyType.Text, content=body)
        msg.to_recipients = [Recipient(email_address=EmailAddress(address=to))]
        result = await client.me.messages.post(msg)
        return result

    try:
        result = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to create draft: {exc}")

    draft_id = (result.id or "")[-20:] if result else ""
    output_status({"status": "draft created", "subject": subject, "to": to, "id": draft_id})


@mail_app.command("drafts")
def list_drafts(
    count: int = typer.Option(10, "--count", "-n", help="Number of drafts"),
) -> None:
    """List draft messages."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import MessagesRequestBuilder
        query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=count,
            select=["id", "subject", "toRecipients", "createdDateTime"],
            orderby=["createdDateTime desc"],
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query,
        )
        result = await client.me.mail_folders.by_mail_folder_id("drafts").messages.get(request_configuration=config)
        return result.value or []

    try:
        messages = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch drafts: {exc}")

    rows = []
    for i, msg in enumerate(messages, 1):
        to_addr = ""
        if msg.to_recipients:
            addrs = [r.email_address.address for r in msg.to_recipients if r.email_address and r.email_address.address]
            to_addr = ", ".join(addrs)
        rows.append({
            "n": i,
            "subject": msg.subject or "(no subject)",
            "to": to_addr,
            "date": str(msg.created_date_time)[:16] if msg.created_date_time else "",
            "id": (msg.id or "")[-20:],
        })

    output_table("Drafts", [("n", "#"), ("subject", "Subject"), ("to", "To"), ("date", "Date"), ("id", "ID")], rows)


@mail_app.command("send-draft")
def send_draft(
    id: str = typer.Argument(..., help="Draft message ID (or partial ID suffix)"),
) -> None:
    """Send an existing draft."""
    client = get_graph_client()

    async def _run():
        # Resolve ID from drafts folder
        from msgraph.generated.users.item.mail_folders.item.messages.messages_request_builder import MessagesRequestBuilder
        query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=100,
            select=["id"],
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query,
        )
        result = await client.me.mail_folders.by_mail_folder_id("drafts").messages.get(request_configuration=config)
        full_id = None
        if result and result.value:
            for msg in result.value:
                if msg.id and msg.id.endswith(id):
                    full_id = msg.id
                    break
        if not full_id:
            raise ValueError(f"No draft found ending with '{id}'")
        await client.me.messages.by_message_id(full_id).send.post()
        return full_id

    try:
        full_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to send draft: {exc}")

    output_status({"status": "draft sent", "id": full_id[-20:]})


# ---------------------------------------------------------------------------
# Mail action commands — categories
# ---------------------------------------------------------------------------


@mail_app.command("categorize")
def categorize_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
    category: str = typer.Argument(..., help="Category to add"),
) -> None:
    """Add a category to a message."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.users.item.messages.item.message_item_request_builder import MessageItemRequestBuilder
        from msgraph.generated.models.message import Message as MsgModel
        # Get current categories
        params = MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters(
            select=["id", "subject", "categories"],
        )
        cfg = MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration(
            query_parameters=params,
        )
        msg = await client.me.messages.by_message_id(full_id).get(request_configuration=cfg)
        current_cats = list(msg.categories) if msg and msg.categories else []
        if category not in current_cats:
            current_cats.append(category)
        body = MsgModel()
        body.categories = current_cats
        await client.me.messages.by_message_id(full_id).patch(body)
        subject = msg.subject if msg else "..."
        return full_id, subject, current_cats

    try:
        full_id, subject, cats = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to categorize message: {exc}")

    output_status({"status": "categorized", "subject": subject, "categories": cats, "id": full_id[-20:]})


@mail_app.command("uncategorize")
def uncategorize_msg(
    id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
) -> None:
    """Remove all categories from a message."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, id)
        from msgraph.generated.users.item.messages.item.message_item_request_builder import MessageItemRequestBuilder
        from msgraph.generated.models.message import Message as MsgModel
        params = MessageItemRequestBuilder.MessageItemRequestBuilderGetQueryParameters(
            select=["id", "subject"],
        )
        cfg = MessageItemRequestBuilder.MessageItemRequestBuilderGetRequestConfiguration(
            query_parameters=params,
        )
        msg = await client.me.messages.by_message_id(full_id).get(request_configuration=cfg)
        body = MsgModel()
        body.categories = []
        await client.me.messages.by_message_id(full_id).patch(body)
        subject = msg.subject if msg else "..."
        return full_id, subject

    try:
        full_id, subject = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to uncategorize message: {exc}")

    output_status({"status": "categories removed", "subject": subject, "id": full_id[-20:]})


# ---------------------------------------------------------------------------
# Mail action commands — folders
# ---------------------------------------------------------------------------


@mail_app.command("create-folder")
def create_folder(
    name: str = typer.Argument(..., help="Folder name"),
    parent: str = typer.Option(None, "--parent", help="Parent folder name"),
) -> None:
    """Create a new mail folder."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.models.mail_folder import MailFolder
        folder = MailFolder()
        folder.display_name = name

        if parent:
            # Look up parent folder by name (case-insensitive)
            result = await client.me.mail_folders.get()
            parent_id = None
            if result and result.value:
                for f in result.value:
                    if f.display_name and f.display_name.lower() == parent.lower():
                        parent_id = f.id
                        break
            if not parent_id:
                raise ValueError(f"Parent folder '{parent}' not found.")
            created = await client.me.mail_folders.by_mail_folder_id(parent_id).child_folders.post(folder)
        else:
            created = await client.me.mail_folders.post(folder)
        return created

    try:
        created = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to create folder: {exc}")

    folder_id = (created.id or "")[-20:] if created else ""
    output_status({"status": "folder created", "name": name, "id": folder_id})


@mail_app.command("delete-folder")
def delete_folder_cmd(
    name: str = typer.Argument(..., help="Folder name to delete"),
) -> None:
    """Delete a mail folder."""
    client = get_graph_client()

    async def _run():
        result = await client.me.mail_folders.get()
        folder_id = None
        if result and result.value:
            for f in result.value:
                if f.display_name and f.display_name.lower() == name.lower():
                    folder_id = f.id
                    break
        if not folder_id:
            raise ValueError(f"Folder '{name}' not found.")
        await client.me.mail_folders.by_mail_folder_id(folder_id).delete()

    try:
        asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to delete folder: {exc}")

    output_status({"status": "folder deleted", "name": name})


# ---------------------------------------------------------------------------
# Mail action commands — bulk operations
# ---------------------------------------------------------------------------


@mail_app.command("bulk-read")
def bulk_read(
    ids: list[str] = typer.Argument(..., help="Message IDs (or partial ID suffixes)"),
) -> None:
    """Mark multiple messages as read."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.models.message import Message as MsgModel
        marked = 0
        not_found = 0
        for partial_id in ids:
            try:
                full_id = await _resolve_message_id(client, partial_id)
                body = MsgModel()
                body.is_read = True
                await client.me.messages.by_message_id(full_id).patch(body)
                marked += 1
            except Exception:
                not_found += 1
        return marked, not_found

    try:
        marked, not_found = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Bulk read failed: {exc}")

    output_status({"status": "bulk operation complete", "marked_read": marked, "not_found": not_found})


@mail_app.command("bulk-delete")
def bulk_delete(
    ids: list[str] = typer.Argument(..., help="Message IDs (or partial ID suffixes)"),
) -> None:
    """Move multiple messages to trash."""
    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.messages.item.move.move_post_request_body import MovePostRequestBody
        deleted = 0
        not_found = 0
        for partial_id in ids:
            try:
                full_id = await _resolve_message_id(client, partial_id)
                body = MovePostRequestBody()
                body.destination_id = "deleteditems"
                await client.me.messages.by_message_id(full_id).move.post(body)
                deleted += 1
            except Exception:
                not_found += 1
        return deleted, not_found

    try:
        deleted, not_found = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Bulk delete failed: {exc}")

    output_status({"status": "bulk delete complete", "deleted": deleted, "not_found": not_found})


# ---------------------------------------------------------------------------
# Mail action commands — auto-categorize rules
# ---------------------------------------------------------------------------


def _rules_path() -> Path:
    """Return the path to rules.json for the current account."""
    return _account_dir(state.account) / "rules.json"


def _load_rules() -> list[dict]:
    """Load rules from rules.json."""
    rp = _rules_path()
    if not rp.exists():
        return []
    try:
        data = json.loads(rp.read_text())
        return data.get("rules", [])
    except (json.JSONDecodeError, OSError):
        return []


def _save_rules(rules: list[dict]) -> None:
    """Save rules to rules.json."""
    rp = _rules_path()
    rp.parent.mkdir(parents=True, exist_ok=True)
    rp.write_text(json.dumps({"rules": rules}, indent=2) + "\n")
    rp.chmod(0o600)


@mail_app.command("rules")
def list_rules() -> None:
    """Display auto-categorize rules."""
    rules = _load_rules()
    rows = []
    for i, rule in enumerate(rules):
        rows.append({
            "index": i,
            "match": rule.get("match", ""),
            "pattern": rule.get("pattern", ""),
            "category": rule.get("category", ""),
        })

    output_table("Rules", [("index", "Index"), ("match", "Match"), ("pattern", "Pattern"), ("category", "Category")], rows)


@mail_app.command("add-rule")
def add_rule(
    field: str = typer.Argument(..., help="Field to match: 'from' or 'subject'"),
    pattern: str = typer.Argument(..., help="Pattern to match (case-insensitive)"),
    category: str = typer.Argument(..., help="Category to apply"),
) -> None:
    """Add an auto-categorize rule."""
    if field not in ("from", "subject"):
        _error_exit("Field must be 'from' or 'subject'.")

    rules = _load_rules()
    rules.append({"match": field, "pattern": pattern, "category": category})
    _save_rules(rules)

    output_status({"status": "rule added", "match": field, "pattern": pattern, "category": category})


@mail_app.command("remove-rule")
def remove_rule(
    index: int = typer.Argument(..., help="Rule index (0-based)"),
) -> None:
    """Remove an auto-categorize rule by index."""
    rules = _load_rules()
    if index < 0 or index >= len(rules):
        _error_exit(f"Invalid rule index {index}. Valid range: 0-{len(rules) - 1}")

    removed = rules.pop(index)
    _save_rules(rules)

    output_status({"status": "rule removed", "match": removed.get("match", ""), "pattern": removed.get("pattern", ""), "category": removed.get("category", "")})


@mail_app.command("auto-categorize")
def auto_categorize(
    count: int = typer.Option(50, "--count", "-n", help="Number of messages to scan"),
) -> None:
    """Auto-categorize messages using rules."""
    rules = _load_rules()
    if not rules:
        _error_exit("No rules defined. Use 'mail add-rule' to create rules.")

    client = get_graph_client()

    async def _run():
        from msgraph.generated.users.item.messages.messages_request_builder import MessagesRequestBuilder
        from msgraph.generated.models.message import Message as MsgModel

        query = MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters(
            top=count,
            select=["id", "subject", "from", "categories"],
            orderby=["receivedDateTime desc"],
        )
        config = MessagesRequestBuilder.MessagesRequestBuilderGetRequestConfiguration(
            query_parameters=query,
        )
        result = await client.me.messages.get(request_configuration=config)
        messages = result.value or []

        scanned = 0
        categorized = 0
        no_match = 0
        already_categorized = 0

        for msg in messages:
            scanned += 1
            matched_category = None
            from_addr = ""
            if msg.from_ and msg.from_.email_address:
                from_addr = msg.from_.email_address.address or ""
            subj = msg.subject or ""
            current_cats = list(msg.categories) if msg.categories else []

            for rule in rules:
                match_field = rule.get("match", "")
                pat = rule.get("pattern", "").lower()
                cat = rule.get("category", "")
                if match_field == "from" and pat in from_addr.lower():
                    matched_category = cat
                    break
                elif match_field == "subject" and pat in subj.lower():
                    matched_category = cat
                    break

            if matched_category is None:
                no_match += 1
            elif matched_category in current_cats:
                already_categorized += 1
            else:
                current_cats.append(matched_category)
                body = MsgModel()
                body.categories = current_cats
                await client.me.messages.by_message_id(msg.id).patch(body)
                categorized += 1

        return scanned, categorized, no_match, already_categorized

    try:
        scanned, categorized, no_match, already_categorized = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Auto-categorize failed: {exc}")

    output_status({
        "status": "auto-categorize complete",
        "scanned": scanned,
        "categorized": categorized,
        "no_match": no_match,
        "already_categorized": already_categorized,
    })


# ---------------------------------------------------------------------------
# Mail action commands — attachment download
# ---------------------------------------------------------------------------


@mail_app.command("download")
def download_attachment(
    msg_id: str = typer.Argument(..., help="Message ID (or partial ID suffix)"),
    attachment_name: str = typer.Argument(..., help="Attachment filename"),
    output: str = typer.Option(None, "--output", "-o", help="Output file path"),
) -> None:
    """Download an attachment from a message."""
    import base64

    client = get_graph_client()

    async def _run():
        full_id = await _resolve_message_id(client, msg_id)
        result = await client.me.messages.by_message_id(full_id).attachments.get()
        atts = result.value or []
        target = None
        for att in atts:
            if att.name and att.name == attachment_name:
                target = att
                break
        if not target:
            raise ValueError(f"Attachment '{attachment_name}' not found.")
        return full_id, target

    try:
        full_id, target = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to download attachment: {exc}")

    # Sanitize filename — basename only, strip path traversal
    safe_name = os.path.basename(attachment_name)
    if not safe_name:
        _error_exit("Invalid attachment filename.")

    if output:
        out_path = Path(output)
    else:
        out_path = Path.cwd() / safe_name

    # Validate output directory exists
    if not out_path.parent.exists():
        _error_exit(f"Output directory '{out_path.parent}' does not exist.")

    # Decode and save content
    content_bytes = getattr(target, "content_bytes", None)
    if content_bytes is None:
        _error_exit("Attachment has no downloadable content.")

    if isinstance(content_bytes, str):
        file_data = base64.b64decode(content_bytes)
    elif isinstance(content_bytes, bytes):
        file_data = content_bytes
    else:
        file_data = base64.b64decode(str(content_bytes))

    out_path.write_bytes(file_data)
    output_status({"status": "downloaded", "file": str(out_path), "size": len(file_data)})


# ---------------------------------------------------------------------------
# Calendar commands
# ---------------------------------------------------------------------------

EVENT_COLUMNS = [
    ("n", "#"),
    ("subject", "Subject"),
    ("start", "Start"),
    ("end", "End"),
    ("location", "Location"),
    ("id", "ID"),
]


def _format_event_rows(events) -> list[dict]:
    """Format Graph SDK event objects into output rows."""
    rows = []
    for i, ev in enumerate(events, 1):
        start_str = ""
        if ev.start and ev.start.date_time:
            start_str = ev.start.date_time[:16]
        end_str = ""
        if ev.end and ev.end.date_time:
            end_str = ev.end.date_time[:16]
        loc = ""
        if ev.location and ev.location.display_name:
            loc = ev.location.display_name
        rows.append({
            "n": i,
            "subject": ev.subject or "(no subject)",
            "start": start_str,
            "end": end_str,
            "location": loc,
            "id": (ev.id or "")[-20:],
        })
    return rows


@calendar_app.command()
def events(
    count: int = typer.Option(10, "--count", "-n", help="Number of events"),
) -> None:
    """List recent calendar events."""
    client = get_graph_client()
    tz = detect_timezone()

    async def _run():
        from msgraph.generated.users.item.events.events_request_builder import EventsRequestBuilder

        query = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters(
            top=count,
            select=["id", "subject", "start", "end", "location", "isAllDay"],
            orderby=["start/dateTime desc"],
        )
        config = EventsRequestBuilder.EventsRequestBuilderGetRequestConfiguration(
            query_parameters=query,
            headers={"Prefer": f'outlook.timezone="{tz}"'},
        )
        result = await client.me.events.get(request_configuration=config)
        return result.value or []

    try:
        evts = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch events: {exc}")

    output_table("Calendar Events", EVENT_COLUMNS, _format_event_rows(evts))


@calendar_app.command()
def today() -> None:
    """List today's calendar events."""
    client = get_graph_client()
    tz = detect_timezone()
    from datetime import datetime as _dt

    now = _dt.now()
    start_iso = now.strftime("%Y-%m-%dT00:00:00")
    end_iso = now.strftime("%Y-%m-%dT23:59:59")

    async def _run():
        from msgraph.generated.users.item.calendar_view.calendar_view_request_builder import CalendarViewRequestBuilder

        query = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters(
            start_date_time=start_iso,
            end_date_time=end_iso,
            select=["id", "subject", "start", "end", "location", "isAllDay"],
            orderby=["start/dateTime"],
        )
        config = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetRequestConfiguration(
            query_parameters=query,
            headers={"Prefer": f'outlook.timezone="{tz}"'},
        )
        result = await client.me.calendar_view.get(request_configuration=config)
        return result.value or []

    try:
        evts = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch today's events: {exc}")

    output_table("Today's Events", EVENT_COLUMNS, _format_event_rows(evts))


@calendar_app.command()
def week() -> None:
    """List this week's calendar events (next 7 days)."""
    client = get_graph_client()
    tz = detect_timezone()
    from datetime import datetime as _dt, timedelta

    now = _dt.now()
    start_iso = now.strftime("%Y-%m-%dT00:00:00")
    end_dt = now + timedelta(days=7)
    end_iso = end_dt.strftime("%Y-%m-%dT23:59:59")

    async def _run():
        from msgraph.generated.users.item.calendar_view.calendar_view_request_builder import CalendarViewRequestBuilder

        query = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters(
            start_date_time=start_iso,
            end_date_time=end_iso,
            select=["id", "subject", "start", "end", "location", "isAllDay"],
            orderby=["start/dateTime"],
        )
        config = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetRequestConfiguration(
            query_parameters=query,
            headers={"Prefer": f'outlook.timezone="{tz}"'},
        )
        result = await client.me.calendar_view.get(request_configuration=config)
        return result.value or []

    try:
        evts = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to fetch week's events: {exc}")

    output_table("This Week's Events", EVENT_COLUMNS, _format_event_rows(evts))


@calendar_app.command()
def read(
    event_id: str = typer.Argument(..., help="Event ID (or partial suffix)"),
) -> None:
    """Read full details of a calendar event."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_event_id(client, event_id)
        ev = await client.me.events.by_event_id(full_id).get()
        return ev

    try:
        ev = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to read event: {exc}")

    start_str = ev.start.date_time[:16] if ev.start and ev.start.date_time else ""
    end_str = ev.end.date_time[:16] if ev.end and ev.end.date_time else ""
    loc = ev.location.display_name if ev.location and ev.location.display_name else ""
    body_text = ""
    if ev.body and ev.body.content:
        body_text = strip_html(ev.body.content)[:500]
    attendees_list = []
    if ev.attendees:
        for att in ev.attendees:
            if att.email_address and att.email_address.address:
                attendees_list.append(att.email_address.address)
    join_url = ""
    if ev.online_meeting_url:
        join_url = ev.online_meeting_url
    elif ev.online_meeting and hasattr(ev.online_meeting, "join_url") and ev.online_meeting.join_url:
        join_url = ev.online_meeting.join_url

    output_detail({
        "subject": ev.subject or "",
        "start": start_str,
        "end": end_str,
        "location": loc,
        "body": body_text,
        "attendees": ", ".join(attendees_list) if attendees_list else "",
        "isOnline": getattr(ev, "is_online_meeting", False),
        "link": join_url,
    })


@calendar_app.command()
def calendars() -> None:
    """List all calendars."""
    client = get_graph_client()

    async def _run():
        result = await client.me.calendars.get()
        return result.value or []

    try:
        cals = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to list calendars: {exc}")

    cal_columns = [("name", "Name"), ("color", "Color"), ("canEdit", "Can Edit"), ("id", "ID")]
    rows = []
    for cal in cals:
        rows.append({
            "name": cal.name or "",
            "color": str(cal.color) if cal.color else "",
            "canEdit": getattr(cal, "can_edit", ""),
            "id": (cal.id or "")[-20:],
        })
    output_table("Calendars", cal_columns, rows)


@calendar_app.command()
def free(
    start: str = typer.Argument(..., help="Start datetime (ISO format)"),
    end: str = typer.Argument(..., help="End datetime (ISO format)"),
) -> None:
    """Check free/busy status for a time range."""
    from datetime import datetime as _dt
    try:
        _dt.fromisoformat(start)
    except ValueError:
        _error_exit(f"Invalid datetime format: {start}. Use YYYY-MM-DDTHH:MM")
    try:
        _dt.fromisoformat(end)
    except ValueError:
        _error_exit(f"Invalid datetime format: {end}. Use YYYY-MM-DDTHH:MM")
    client = get_graph_client()
    tz = detect_timezone()

    async def _run():
        from msgraph.generated.users.item.calendar_view.calendar_view_request_builder import CalendarViewRequestBuilder

        query = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetQueryParameters(
            start_date_time=start,
            end_date_time=end,
            select=["id", "subject", "start", "end"],
            orderby=["start/dateTime"],
        )
        config = CalendarViewRequestBuilder.CalendarViewRequestBuilderGetRequestConfiguration(
            query_parameters=query,
            headers={"Prefer": f'outlook.timezone="{tz}"'},
        )
        result = await client.me.calendar_view.get(request_configuration=config)
        return result.value or []

    try:
        evts = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to check free/busy: {exc}")

    if not evts:
        output_status({"status": "free", "start": start, "end": end})
    else:
        subjects = [ev.subject or "(no subject)" for ev in evts]
        output_status({"status": "busy", "events": subjects})


@calendar_app.command()
def create(
    subject: str = typer.Argument(..., help="Event subject"),
    start: str = typer.Argument(..., help="Start datetime (ISO format)"),
    end: str = typer.Argument(..., help="End datetime (ISO format)"),
    location: str = typer.Option(None, "--location", "-l", help="Event location"),
) -> None:
    """Create a new calendar event."""
    from datetime import datetime as _dt
    try:
        _dt.fromisoformat(start)
    except ValueError:
        _error_exit(f"Invalid datetime format: {start}. Use YYYY-MM-DDTHH:MM")
    try:
        _dt.fromisoformat(end)
    except ValueError:
        _error_exit(f"Invalid datetime format: {end}. Use YYYY-MM-DDTHH:MM")
    client = get_graph_client()
    tz = detect_timezone()

    async def _run():
        from msgraph.generated.models.event import Event
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.models.location import Location as LocationModel

        event = Event()
        event.subject = subject
        event.start = DateTimeTimeZone(date_time=start, time_zone=tz)
        event.end = DateTimeTimeZone(date_time=end, time_zone=tz)
        if location:
            event.location = LocationModel(display_name=location)
        result = await client.me.events.post(event)
        return result

    try:
        ev = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to create event: {exc}")

    start_str = ev.start.date_time[:16] if ev.start and ev.start.date_time else start
    end_str = ev.end.date_time[:16] if ev.end and ev.end.date_time else end
    output_status({
        "status": "event created",
        "subject": ev.subject or subject,
        "start": start_str,
        "end": end_str,
        "id": ev.id or "",
    })


@calendar_app.command()
def quick(
    subject: str = typer.Argument(..., help="Event subject"),
    start_time: str = typer.Option(None, "--start", "-s", help="Start datetime (ISO format, default: now)"),
) -> None:
    """Create a quick 1-hour event."""
    client = get_graph_client()
    tz = detect_timezone()
    from datetime import datetime as _dt, timedelta

    if start_time:
        start_dt = _dt.fromisoformat(start_time)
    else:
        start_dt = _dt.now()
    end_dt = start_dt + timedelta(hours=1)
    start_str = start_dt.strftime("%Y-%m-%dT%H:%M")
    end_str = end_dt.strftime("%Y-%m-%dT%H:%M")

    async def _run():
        from msgraph.generated.models.event import Event
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone

        event = Event()
        event.subject = subject
        event.start = DateTimeTimeZone(date_time=start_str, time_zone=tz)
        event.end = DateTimeTimeZone(date_time=end_str, time_zone=tz)
        result = await client.me.events.post(event)
        return result

    try:
        ev = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to create quick event: {exc}")

    output_status({
        "status": "quick event created",
        "subject": ev.subject or subject,
        "start": start_str,
        "end": end_str,
        "id": ev.id or "",
    })


@calendar_app.command()
def update(
    event_id: str = typer.Argument(..., help="Event ID (or partial suffix)"),
    field: str = typer.Argument(..., help="Field to update (subject, location, start, end)"),
    value: str = typer.Argument(..., help="New value"),
) -> None:
    """Update a field on an existing calendar event."""
    client = get_graph_client()
    tz = detect_timezone()
    valid_fields = {"subject", "location", "start", "end"}
    if field not in valid_fields:
        _error_exit(f"Invalid field '{field}'. Valid fields: {', '.join(sorted(valid_fields))}")
    if field in ("start", "end"):
        from datetime import datetime as _dt
        try:
            _dt.fromisoformat(value)
        except ValueError:
            _error_exit(f"Invalid datetime format: {value}. Use YYYY-MM-DDTHH:MM")

    async def _run():
        from msgraph.generated.models.event import Event
        from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
        from msgraph.generated.models.location import Location as LocationModel

        full_id = await _resolve_event_id(client, event_id)
        event = Event()
        if field == "subject":
            event.subject = value
        elif field == "location":
            event.location = LocationModel(display_name=value)
        elif field == "start":
            event.start = DateTimeTimeZone(date_time=value, time_zone=tz)
        elif field == "end":
            event.end = DateTimeTimeZone(date_time=value, time_zone=tz)
        result = await client.me.events.by_event_id(full_id).patch(event)
        return result

    try:
        ev = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to update event: {exc}")

    start_str = ev.start.date_time[:16] if ev.start and ev.start.date_time else ""
    end_str = ev.end.date_time[:16] if ev.end and ev.end.date_time else ""
    output_status({
        "status": "event updated",
        "subject": ev.subject or "",
        "start": start_str,
        "end": end_str,
        "id": ev.id or "",
    })


@calendar_app.command()
def delete(
    event_id: str = typer.Argument(..., help="Event ID (or partial suffix)"),
) -> None:
    """Delete a calendar event."""
    client = get_graph_client()

    async def _run():
        full_id = await _resolve_event_id(client, event_id)
        await client.me.events.by_event_id(full_id).delete()
        return full_id

    try:
        deleted_id = asyncio.run(_run())
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to delete event: {exc}")

    output_status({"status": "event deleted", "id": deleted_id})


# ---------------------------------------------------------------------------
# Token commands (refresh, get, list)
# ---------------------------------------------------------------------------


@token_app.command()
def refresh() -> None:
    """Force a token refresh."""
    acct = state.account
    try:
        config = load_config(acct)
        client_id = config.get("client_id")
        if not client_id:
            _error_exit(f"No client_id in config for account '{acct}'.")
        cache_path = _account_dir(acct) / "token_cache.json"
        credential = MsalTokenCredential(client_id, cache_path)
        credential.get_token("https://graph.microsoft.com/.default", force_refresh=True)
        output_status({"status": "ok", "message": "Token refreshed successfully"})
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Token refresh failed: {exc}")


@token_app.command()
def get() -> None:
    """Print the current access token."""
    acct = state.account
    try:
        config = load_config(acct)
        client_id = config.get("client_id")
        if not client_id:
            _error_exit(f"No client_id in config for account '{acct}'.")
        cache_path = _account_dir(acct) / "token_cache.json"
        credential = MsalTokenCredential(client_id, cache_path)
        token = credential.get_token("https://graph.microsoft.com/.default")
        print(token.token)
    except AuthError as exc:
        _error_exit(str(exc))
    except Exception as exc:
        _error_exit(f"Failed to get token: {exc}")


@token_app.command(name="list")
def list_accounts() -> None:
    """List configured accounts."""
    if not BASE_DIR.exists():
        output_table("Accounts", [("account", "Account"), ("has_token", "Has Token")], [])
        return

    rows = []
    for entry in sorted(BASE_DIR.iterdir()):
        if entry.is_dir() and (entry / "config.json").exists():
            has_token = (entry / "token_cache.json").exists()
            rows.append({
                "account": entry.name,
                "has_token": str(has_token),
            })
    output_table("Accounts", [("account", "Account"), ("has_token", "Has Token")], rows)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app()
