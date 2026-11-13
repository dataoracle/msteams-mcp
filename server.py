"""Microsoft Teams MCP server (delegated Microsoft Graph access).

Read tools:  whoami, find_people, list_chats, get_chat, read_chat_messages,
             get_chat_message, search_messages, list_teams, list_channels,
             list_team_members, read_channel_messages, get_channel_thread.
Write tools (only when ENABLE_WRITE_TOOLS=true):
             send_chat_message, send_direct_message, send_channel_message,
             reply_to_channel_message, create_online_meeting.
"""
from __future__ import annotations

import asyncio
import hmac
import html
import logging
import os
import re
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote

import uvicorn
from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from starlette.requests import Request
from starlette.responses import JSONResponse

import formatting as fmt
from auth import GUID_RE, AuthError, TokenManager, check_client_id, check_tenant
from graph_client import GraphClient, GraphError
from vault import Vault, VaultError, load_key, vault_context

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("msteams-mcp")


# ---------------------------------------------------------------- config
def _env(name: str, default: str | None = None) -> str:
    value = os.getenv(name, default)
    if value is None or value.strip() == "":
        raise SystemExit(f"Missing required environment variable {name}")
    return value.strip()


def _secret(env_name: str, cred_name: str) -> str:
    """Load a secret, preferring systemd credentials over environment variables.

    1. $CREDENTIALS_DIRECTORY/<cred_name>  (systemd LoadCredentialEncrypted=)
    2. file named by <ENV_NAME>_FILE       (e.g. Docker secrets)
    3. <ENV_NAME> environment variable      (development only)
    """
    cred_dir = os.getenv("CREDENTIALS_DIRECTORY")
    if cred_dir:
        path = Path(cred_dir) / cred_name
        if path.is_file():
            value = path.read_text(encoding="utf-8").rstrip("\r\n")
            if value:
                return value
    file_var = os.getenv(f"{env_name}_FILE")
    if file_var:
        value = Path(file_var).read_text(encoding="utf-8").rstrip("\r\n")
        if value:
            return value
    return _env(env_name)


def _flag(name: str, default: str) -> bool:
    return os.getenv(name, default).strip().lower() in ("1", "true", "yes", "on")


TENANT = check_tenant(_env("MS_TENANT_ID"))
CLIENT_ID = check_client_id(_env("MS_CLIENT_ID"))
MCP_TOKEN = _secret("MCP_TOKEN", "mcp_token")          # shared secret for mcp-remote
CLIENT_SECRET = _secret("MS_CLIENT_SECRET", "client_secret")  # Entra app secret (Web platform)
try:
    VAULT_KEY = load_key(_secret("VAULT_KEY", "vault_key"))
except VaultError as e:
    raise SystemExit(str(e)) from None
# systemd sets STATE_DIRECTORY=/var/lib/msteams-mcp (StateDirectory=)
_state_dir = os.getenv("STATE_DIRECTORY", ".").split(":")[0]
VAULT_PATH = os.getenv("VAULT_PATH") or str(Path(_state_dir) / "vault.json.enc")
HOST = os.getenv("HOST", "127.0.0.1")
PORT = int(os.getenv("PORT", "8766"))
ENABLE_WRITE = _flag("ENABLE_WRITE_TOOLS", "true")
KEEPALIVE_HOURS = float(os.getenv("KEEPALIVE_HOURS", "12"))
MAX_LIST_TEXT_CHARS = int(os.getenv("MAX_LIST_TEXT_CHARS", "2000"))
MAX_SEND_BYTES = 28000   # Graph rejects message bodies above ~28 KB

vault = Vault(VAULT_PATH, VAULT_KEY, vault_context(TENANT, CLIENT_ID))
tokens = TokenManager(vault, TENANT, CLIENT_ID, CLIENT_SECRET)
graph = GraphClient(tokens)

READ_ONLY = {"readOnlyHint": True, "openWorldHint": True}
WRITES = {"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False,
          "openWorldHint": True}

mcp = FastMCP(
    "microsoft-teams",
    instructions=(
        "Access to Microsoft Teams chats, channels and messages as the signed-in user "
        "(delegated permissions). Use list_chats to find a chat_id, then read_chat_messages. "
        "Use search_messages to find messages by keyword across chats and channels. "
        "For channels, use list_teams, then list_channels, then read_channel_messages. "
        + ("Messages you send appear as the signed-in user and cannot be unsent from here: "
           "confirm the recipient and wording with the user before calling any send_ or "
           "reply_ tool." if ENABLE_WRITE else
           "This server is configured read-only: it cannot send messages.")
    ),
)

# ---------------------------------------------------------------- helpers
ID_RE = re.compile(r"^[A-Za-z0-9:@._=\-]{1,400}$")


def _id(value: str, label: str, source: str) -> str:
    value = (value or "").strip()
    if not ID_RE.match(value):
        raise ToolError(f"{label} {value[:80]!r} doesn't look like a Teams id. Get it from {source}.")
    return value


def _team(team_id: str) -> str:
    team_id = (team_id or "").strip()
    if not GUID_RE.match(team_id):
        raise ToolError(f"team_id {team_id[:80]!r} must be a GUID from list_teams.")
    return team_id


def _clamp(n: int, lo: int, hi: int) -> int:
    return max(lo, min(int(n), hi))


async def _call(coro):
    """Run a data function and turn backend errors into tool errors."""
    try:
        return await coro
    except (GraphError, AuthError) as e:
        raise ToolError(str(e)) from e


async def _me_id() -> str:
    uid = tokens.user.get("id")
    if not uid:
        me = await graph.get("/me", {"$select": "id"})
        uid = me["id"]
    return uid


def _body(content: str, format: str) -> dict:
    content = (content or "").strip()
    if not content:
        raise ToolError("The message is empty.")
    if format == "text":
        body = html.escape(content).replace("\n", "<br>")
    elif format == "html":
        body = content
    else:
        raise ToolError("format must be 'text' or 'html'.")
    if len(body.encode()) > MAX_SEND_BYTES:
        raise ToolError(f"The message is too long for Teams (max ~{MAX_SEND_BYTES // 1000} KB).")
    return {"contentType": "html", "content": body}


def _sent(m: dict, **ids) -> dict:
    return {"sent": True, **ids, "message_id": m.get("id"), "created": m.get("createdDateTime"),
            **({"web_url": m["webUrl"]} if m.get("webUrl") else {})}


PERSON_FIELDS = "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation"


# ---------------------------------------------------------------- data functions
async def whoami_data() -> dict:
    me = await graph.get("/me", {"$select": PERSON_FIELDS})
    return {"user": fmt.person(me), "auth": tokens.info()}


async def find_people_data(query: str, limit: int) -> list[dict]:
    q = (query or "").strip()
    if not q or len(q) > 100:
        raise ToolError("Give a name or email address (1-100 characters).")
    if "@" in q and " " not in q:
        try:
            return [fmt.person(await graph.get(f"/users/{quote(q, safe='@')}",
                                               {"$select": PERSON_FIELDS}))]
        except GraphError as e:
            if e.status != 404:
                raise
    term = re.sub(r'["\\:]', " ", q).strip()
    rows = await graph.get("/users", {
        "$search": f'"displayName:{term}" OR "mail:{term}"',
        "$select": PERSON_FIELDS,
        "$top": _clamp(limit, 1, 25),
        "$count": "true",
    }, headers={"ConsistencyLevel": "eventual"})
    return [fmt.person(u) for u in rows.get("value") or []]


async def _fetch_chats(limit: int) -> tuple[list[dict], bool]:
    params = {"$top": 50, "$expand": "members,lastMessagePreview",
              "$orderby": "lastMessagePreview/createdDateTime desc"}
    try:
        return await graph.get_all("/me/chats", params, limit)
    except GraphError as e:
        if e.status != 400:
            raise
        # Fallback if this tenant rejects the combined $expand.
        params["$expand"] = "lastMessagePreview"
        return await graph.get_all("/me/chats", params, limit)


async def list_chats_data(query: str, limit: int) -> dict:
    limit = _clamp(limit, 1, 100)
    q = (query or "").strip().lower()
    rows, more = await _fetch_chats(250 if q else limit)
    me = await _me_id()
    chats = [fmt.chat(c, me) for c in rows]
    if q:
        chats = [c for c in chats
                 if q in c["name"].lower() or any(q in m.lower() for m in c.get("members", []))]
    out = {"count": min(len(chats), limit), "chats": chats[:limit]}
    if q and more:
        out["note"] = "Only the 250 most recently active chats were searched."
    elif not q and more:
        out["note"] = f"Showing the {limit} most recently active chats; more exist."
    return out


async def get_chat_data(chat_id: str) -> dict:
    cid = _id(chat_id, "chat_id", "list_chats")
    c = await graph.get(f"/chats/{cid}", {"$expand": "members"})
    out = fmt.chat(c, await _me_id())
    out.update({
        "topic": c.get("topic"),
        "created": c.get("createdDateTime"),
        "members": [fmt.member(m) for m in c.get("members") or []],
    })
    meeting = c.get("onlineMeetingInfo") or {}
    if meeting.get("joinWebUrl"):
        out["meeting_join_url"] = meeting["joinWebUrl"]
    return out


async def read_chat_messages_data(chat_id: str, limit: int, since: str, include_system: bool) -> dict:
    cid = _id(chat_id, "chat_id", "list_chats")
    limit = _clamp(limit, 1, 200)
    params: dict = {"$top": min(limit, 50)}
    if since.strip():
        try:
            dt = fmt.parse_time(since)
        except ValueError:
            raise ToolError("since must be an ISO 8601 date/time, e.g. 2026-11-13T08:00:00Z") from None
        params["$orderby"] = "lastModifiedDateTime desc"
        params["$filter"] = f"lastModifiedDateTime gt {fmt.iso_z(dt)}"
    else:
        params["$orderby"] = "createdDateTime desc"
    rows, more = await graph.get_all(f"/chats/{cid}/messages", params, limit)
    if not include_system:
        rows = [m for m in rows if m.get("messageType") in (None, "message")]
    rows.sort(key=lambda m: m.get("createdDateTime") or "")
    out = {"chat_id": cid, "count": len(rows), "order": "oldest first",
           "messages": [fmt.message(m, MAX_LIST_TEXT_CHARS) for m in rows]}
    if more:
        out["note"] = f"Showing the latest {limit} messages; older ones exist (raise limit or use since)."
    return out


async def get_chat_message_data(chat_id: str, message_id: str) -> dict:
    cid = _id(chat_id, "chat_id", "list_chats")
    mid = _id(message_id, "message_id", "read_chat_messages")
    return fmt.message(await graph.get(f"/chats/{cid}/messages/{mid}"))


async def search_messages_data(query: str, limit: int) -> dict:
    q = (query or "").strip()
    if not q:
        raise ToolError("Give something to search for.")
    r = await graph.post("/search/query", {"requests": [{
        "entityTypes": ["chatMessage"],
        "query": {"queryString": q},
        "from": 0,
        "size": _clamp(limit, 1, 50),
    }]})
    hits = []
    for v in r.get("value") or []:
        for c in v.get("hitsContainers") or []:
            for h in c.get("hits") or []:
                res = h.get("resource") or {}
                channel = res.get("channelIdentity") or {}
                sender = ((res.get("from") or {}).get("emailAddress") or {})
                hits.append({k: val for k, val in {
                    "summary": fmt.strip_search_markup(h.get("summary") or ""),
                    "created": res.get("createdDateTime"),
                    "from": sender.get("name") or sender.get("address"),
                    "chat_id": res.get("chatId"),
                    "team_id": channel.get("teamId"),
                    "channel_id": channel.get("channelId"),
                    "message_id": res.get("id"),
                    "web_url": res.get("webUrl"),
                }.items() if val})
    return {"query": q, "count": len(hits), "hits": hits}


async def list_teams_data() -> list[dict]:
    rows, _ = await graph.get_all("/me/joinedTeams", {"$select": "id,displayName,description"}, 500)
    return [{"team_id": t.get("id"), "name": t.get("displayName"),
             "description": t.get("description")} for t in rows]


async def list_channels_data(team_id: str) -> list[dict]:
    tid = _team(team_id)
    rows, _ = await graph.get_all(
        f"/teams/{tid}/channels",
        {"$select": "id,displayName,description,membershipType,webUrl"}, 500)
    return [{"channel_id": c.get("id"), "name": c.get("displayName"),
             "description": c.get("description"), "membership": c.get("membershipType"),
             "web_url": c.get("webUrl")} for c in rows]


async def list_team_members_data(team_id: str, limit: int) -> dict:
    tid = _team(team_id)
    rows, more = await graph.get_all(f"/teams/{tid}/members", None, _clamp(limit, 1, 1000))
    out = {"count": len(rows), "members": [fmt.member(m) for m in rows]}
    if more:
        out["note"] = "More members exist; raise limit."
    return out


async def read_channel_messages_data(team_id: str, channel_id: str, limit: int,
                                     include_replies: bool) -> dict:
    tid = _team(team_id)
    chid = _id(channel_id, "channel_id", "list_channels")
    limit = _clamp(limit, 1, 100)
    params: dict = {"$top": min(limit, 50)}
    if include_replies:
        params["$expand"] = "replies"
    rows, more = await graph.get_all(f"/teams/{tid}/channels/{chid}/messages", params, limit)
    rows.sort(key=lambda m: m.get("createdDateTime") or "")
    threads = []
    for m in rows:
        t = fmt.message(m, MAX_LIST_TEXT_CHARS)
        if include_replies:
            replies = sorted(m.get("replies") or [], key=lambda r: r.get("createdDateTime") or "")
            t["reply_count"] = len(replies)
            t["replies"] = [fmt.message(r, MAX_LIST_TEXT_CHARS) for r in replies[-20:]]
            if len(replies) > 20:
                t["replies_note"] = "Only the latest 20 replies shown; use get_channel_thread."
        threads.append(t)
    out = {"team_id": tid, "channel_id": chid, "count": len(threads),
           "order": "oldest first", "threads": threads}
    if more:
        out["note"] = f"Showing the latest {limit} threads; older ones exist."
    return out


async def get_channel_thread_data(team_id: str, channel_id: str, message_id: str) -> dict:
    tid = _team(team_id)
    chid = _id(channel_id, "channel_id", "list_channels")
    mid = _id(message_id, "message_id", "read_channel_messages")
    base = f"/teams/{tid}/channels/{chid}/messages/{mid}"
    root, (replies, more) = await asyncio.gather(
        graph.get(base), graph.get_all(f"{base}/replies", {"$top": 50}, 500))
    replies.sort(key=lambda r: r.get("createdDateTime") or "")
    out = {"message": fmt.message(root), "reply_count": len(replies),
           "replies": [fmt.message(r) for r in replies]}
    if more:
        out["note"] = "Only the first 500 replies are shown."
    return out


# ---------------------------------------------------------------- write functions
async def _resolve_user(user: str) -> dict:
    u = (user or "").strip()
    if not (GUID_RE.match(u) or ("@" in u and " " not in u)):
        raise ToolError("Give the person's email address or user_id (find_people returns both). "
                        "Names are not accepted, to avoid messaging the wrong person.")
    try:
        return await graph.get(f"/users/{quote(u, safe='@')}", {"$select": PERSON_FIELDS})
    except GraphError as e:
        if e.status == 404:
            raise ToolError(f"No user {u!r} in this tenant.") from None
        raise


async def send_chat_message_data(chat_id: str, content: str, format: str, importance: str) -> dict:
    cid = _id(chat_id, "chat_id", "list_chats")
    payload: dict = {"body": _body(content, format)}
    if importance not in ("normal", "high", "urgent"):
        raise ToolError("importance must be normal, high or urgent.")
    if importance != "normal":
        payload["importance"] = importance
    m = await graph.post(f"/chats/{cid}/messages", payload)
    return _sent(m, chat_id=cid)


async def send_direct_message_data(user: str, content: str, format: str) -> dict:
    body = _body(content, format)
    target = await _resolve_user(user)
    me = await _me_id()
    if target["id"].lower() == me.lower():
        raise ToolError("That is the signed-in user; Teams can't create a 1:1 chat with yourself.")

    def member(uid: str) -> dict:
        return {"@odata.type": "#microsoft.graph.aadUserConversationMember", "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{uid}')"}

    # Returns the existing 1:1 chat if there already is one.
    c = await graph.post("/chats", {"chatType": "oneOnOne", "members": [member(me), member(target["id"])]})
    m = await graph.post(f"/chats/{c['id']}/messages", {"body": body})
    return _sent(m, chat_id=c["id"], to=fmt.person(target))


async def send_channel_message_data(team_id: str, channel_id: str, content: str,
                                    format: str, subject: str) -> dict:
    tid = _team(team_id)
    chid = _id(channel_id, "channel_id", "list_channels")
    payload: dict = {"body": _body(content, format)}
    if subject.strip():
        payload["subject"] = subject.strip()[:255]
    m = await graph.post(f"/teams/{tid}/channels/{chid}/messages", payload)
    return _sent(m, team_id=tid, channel_id=chid)


async def reply_to_channel_message_data(team_id: str, channel_id: str, message_id: str,
                                        content: str, format: str) -> dict:
    tid = _team(team_id)
    chid = _id(channel_id, "channel_id", "list_channels")
    mid = _id(message_id, "message_id", "read_channel_messages")
    m = await graph.post(f"/teams/{tid}/channels/{chid}/messages/{mid}/replies",
                         {"body": _body(content, format)})
    return _sent(m, team_id=tid, channel_id=chid, reply_to_id=mid)


async def create_online_meeting_data(subject: str, start: str, end: str) -> dict:
    try:
        start_dt = fmt.parse_time(start) if start.strip() else datetime.now(timezone.utc)
        end_dt = fmt.parse_time(end) if end.strip() else start_dt + timedelta(minutes=30)
    except ValueError:
        raise ToolError("start/end must be ISO 8601, e.g. 2026-11-14T10:00:00+04:00") from None
    if end_dt <= start_dt:
        raise ToolError("end must be after start.")
    m = await graph.post("/me/onlineMeetings", {
        "subject": (subject or "Meeting").strip()[:255],
        "startDateTime": fmt.iso_z(start_dt),
        "endDateTime": fmt.iso_z(end_dt),
    })
    return {"meeting_id": m.get("id"), "subject": m.get("subject"),
            "start": m.get("startDateTime"), "end": m.get("endDateTime"),
            "join_url": m.get("joinWebUrl"),
            "note": "This creates a Teams meeting link only; no calendar invitation is sent. "
                    "Share join_url, for example with send_chat_message."}


# ---------------------------------------------------------------- read tools
@mcp.tool(annotations=READ_ONLY)
async def whoami() -> dict:
    """The Microsoft 365 account this server acts as, and the state of its sign-in."""
    return await _call(whoami_data())


@mcp.tool(annotations=READ_ONLY)
async def find_people(query: str, limit: int = 10) -> list[dict]:
    """Look up people in the directory by name or email. Returns name, email,
    user_id, job title and department. Use the email or user_id with
    send_direct_message."""
    return await _call(find_people_data(query, limit))


@mcp.tool(annotations=READ_ONLY)
async def list_chats(query: str = "", limit: int = 25) -> dict:
    """The signed-in user's chats (1:1, group and meeting chats), most recently
    active first, with chat_id, name, members, a preview of the last message and
    an unread flag. With query, keeps only chats whose title or member names
    contain it (case-insensitive)."""
    return await _call(list_chats_data(query, limit))


@mcp.tool(annotations=READ_ONLY)
async def get_chat(chat_id: str) -> dict:
    """Details of one chat: type, topic, full member list with roles, and the
    meeting join link for meeting chats."""
    return await _call(get_chat_data(chat_id))


@mcp.tool(annotations=READ_ONLY)
async def read_chat_messages(chat_id: str, limit: int = 30, since: str = "",
                             include_system: bool = False) -> dict:
    """Recent messages in a chat, returned oldest first. limit is 1-200.
    since (ISO 8601, e.g. 2026-11-13T08:00:00Z) returns only messages created or
    edited after that time. System events (members added, calls, renames) are
    left out unless include_system is true. Long messages are cut; use
    get_chat_message for the full text."""
    return await _call(read_chat_messages_data(chat_id, limit, since, include_system))


@mcp.tool(annotations=READ_ONLY)
async def get_chat_message(chat_id: str, message_id: str) -> dict:
    """One chat message in full: text, sender, attachments, mentions, reactions."""
    return await _call(get_chat_message_data(chat_id, message_id))


@mcp.tool(annotations=READ_ONLY)
async def search_messages(query: str, limit: int = 25) -> dict:
    """Full-text search across the signed-in user's chats and channel messages
    (Microsoft Search, KQL syntax, e.g. 'budget from:alice'). Returns
    snippets with chat_id or team_id/channel_id and message_id to read more."""
    return await _call(search_messages_data(query, limit))


@mcp.tool(annotations=READ_ONLY)
async def list_teams() -> list[dict]:
    """Teams the signed-in user is a member of, with team_id."""
    return await _call(list_teams_data())


@mcp.tool(annotations=READ_ONLY)
async def list_channels(team_id: str) -> list[dict]:
    """Channels in a team that the signed-in user can see, with channel_id."""
    return await _call(list_channels_data(team_id))


@mcp.tool(annotations=READ_ONLY)
async def list_team_members(team_id: str, limit: int = 200) -> dict:
    """Members of a team with their roles (owner, member, guest)."""
    return await _call(list_team_members_data(team_id, limit))


@mcp.tool(annotations=READ_ONLY)
async def read_channel_messages(team_id: str, channel_id: str, limit: int = 20,
                                include_replies: bool = True) -> dict:
    """Recent top-level posts in a channel (limit 1-100), oldest first, each with
    its latest replies when include_replies is true."""
    return await _call(read_channel_messages_data(team_id, channel_id, limit, include_replies))


@mcp.tool(annotations=READ_ONLY)
async def get_channel_thread(team_id: str, channel_id: str, message_id: str) -> dict:
    """One channel post and all its replies, in full."""
    return await _call(get_channel_thread_data(team_id, channel_id, message_id))


# ---------------------------------------------------------------- write tools
async def send_chat_message(chat_id: str, content: str, format: str = "text",
                            importance: str = "normal") -> dict:
    """Send a message to an existing chat as the signed-in user. format 'text'
    sends the content as typed (line breaks kept); 'html' allows Teams HTML such
    as <b>, <i>, <a href>, <ul><li>, <br>. importance: normal, high or urgent.
    Confirm the chat and wording with the user first: messages can't be unsent from here."""
    return await _call(send_chat_message_data(chat_id, content, format, importance))


async def send_direct_message(user: str, content: str, format: str = "text") -> dict:
    """Send a 1:1 message to a person by email address or user_id (from
    find_people), creating the 1:1 chat if it doesn't exist yet. Confirm the
    recipient and wording with the user first."""
    return await _call(send_direct_message_data(user, content, format))


async def send_channel_message(team_id: str, channel_id: str, content: str,
                               format: str = "text", subject: str = "") -> dict:
    """Start a new post in a channel as the signed-in user, with an optional
    subject. Confirm the channel and wording with the user first."""
    return await _call(send_channel_message_data(team_id, channel_id, content, format, subject))


async def reply_to_channel_message(team_id: str, channel_id: str, message_id: str,
                                   content: str, format: str = "text") -> dict:
    """Reply in the thread of a channel post as the signed-in user. Confirm the
    wording with the user first."""
    return await _call(reply_to_channel_message_data(team_id, channel_id, message_id, content, format))


async def create_online_meeting(subject: str, start: str = "", end: str = "") -> dict:
    """Create a Teams meeting and return its join link. start/end are ISO 8601
    with an offset (e.g. 2026-11-14T10:00:00+04:00); start defaults to now and
    end to start + 30 minutes. No calendar invitation is sent."""
    return await _call(create_online_meeting_data(subject, start, end))


if ENABLE_WRITE:
    for _fn in (send_chat_message, send_direct_message, send_channel_message,
                reply_to_channel_message, create_online_meeting):
        mcp.tool(annotations=WRITES)(_fn)


# ---------------------------------------------------------------- http + auth
@mcp.custom_route("/health", methods=["GET"])
async def health(_: Request):
    return JSONResponse({"status": "ok", "auth": tokens.status()})


class BearerAuth:
    """Minimal ASGI middleware: every request except /health needs the token."""

    def __init__(self, app, token: str):
        self.app = app
        self.expected = f"Bearer {token}".encode()

    async def __call__(self, scope, receive, send):
        if scope["type"] == "http" and scope.get("path") != "/health":
            supplied = dict(scope.get("headers") or []).get(b"authorization", b"")
            if not hmac.compare_digest(supplied, self.expected):
                await send({"type": "http.response.start", "status": 401,
                            "headers": [(b"content-type", b"text/plain")]})
                await send({"type": "http.response.body", "body": b"Unauthorized"})
                return
        await self.app(scope, receive, send)


def build_app():
    return BearerAuth(mcp.http_app(path="/mcp"), MCP_TOKEN)


async def _keepalive_loop() -> None:
    """Refresh periodically so the stored refresh token keeps rotating even when
    the server is idle; otherwise it would expire after ~90 days of no use."""
    if KEEPALIVE_HOURS <= 0:
        return
    await asyncio.sleep(60)
    while True:
        await tokens.keepalive()
        await asyncio.sleep(KEEPALIVE_HOURS * 3600)


async def main() -> None:
    log.info("Starting on %s:%d (write tools %s, vault %s, sign-in %s)", HOST, PORT,
             "enabled" if ENABLE_WRITE else "disabled", VAULT_PATH, tokens.status())
    server = uvicorn.Server(uvicorn.Config(build_app(), host=HOST, port=PORT, proxy_headers=True))
    keepalive = asyncio.create_task(_keepalive_loop())
    try:
        await server.serve()
    finally:
        keepalive.cancel()
        await graph.aclose()
        await tokens.aclose()


if __name__ == "__main__":
    asyncio.run(main())
