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
import json
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
        "Read tools return compact text (times in UTC); pass output='json' for structured data. "
        + ("Messages you send appear as the signed-in user and cannot be unsent from here: "
           "confirm the recipient and wording with the user before calling any send_ or "
           "reply_ tool." if ENABLE_WRITE else
           "This server is configured read-only: it cannot send messages.")
    ),
)

# ---------------------------------------------------------------- helpers
ID_RE = re.compile(r"^[A-Za-z0-9:@._=\-]{1,400}$")
# Tools return plain text: no output schema, so FastMCP doesn't also send a structured copy.
READ = {"annotations": READ_ONLY, "output_schema": None}


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


async def _run(coro, render, output: str = "text") -> str:
    """Run a data function; return compact text (default) or compact JSON."""
    if output not in ("text", "json"):
        coro.close()
        raise ToolError("output must be 'text' or 'json'.")
    data = await _call(coro)
    if output == "json":
        return json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return render(data)


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


PERSON_FIELDS = "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation"


# ---------------------------------------------------------------- data functions
async def whoami_data() -> dict:
    me = await graph.get("/me", {"$select": PERSON_FIELDS})
    return {"user": fmt.person(me), "auth": {k: v for k, v in tokens.info().items() if v}}


def whoami_text(d: dict) -> str:
    a = d["auth"]
    return "\n".join([
        "Signed in as " + fmt.render_people([d["user"]]).removeprefix("• "),
        f"Auth: {a.get('status')}"
        + (f" · signed in {fmt.short_time(a['signed_in_at'])}" if a.get("signed_in_at") else "")
        + (f" · token rotated {fmt.short_time(a['refresh_token_rotated_at'])}"
           if a.get("refresh_token_rotated_at") else "")
        + (f" · problem: {a['problem']}" if a.get("problem") else ""),
    ])


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
    if q:
        def hit(c: dict) -> bool:
            names = [m.get("displayName") or "" for m in c.get("members") or []]
            return any(q in s.lower() for s in [c.get("topic") or ""] + names)
        rows = [c for c in rows if hit(c)]
    chats = [fmt.chat(c, me) for c in rows[:limit]]
    out: dict = {"count": len(chats), "chats": chats}
    if q and more:
        out["note"] = "only the 250 most recently active chats were searched"
    elif not q and more:
        out["note"] = f"more chats exist beyond these {limit}"
    return out


def list_chats_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} chats, most recent first (times UTC)",
                           fmt.render_chats(d["chats"]), d.get("note"))


async def get_chat_data(chat_id: str) -> dict:
    cid = _id(chat_id, "chat_id", "list_chats")
    c = await graph.get(f"/chats/{cid}", {"$expand": "members"})
    out = fmt.chat(c, await _me_id(), full=True)
    out.pop("members", None)
    out.update({
        "topic": c.get("topic"),
        "created": c.get("createdDateTime"),
        "members": [fmt.member(m) for m in c.get("members") or []],
    })
    meeting = c.get("onlineMeetingInfo") or {}
    if meeting.get("joinWebUrl"):
        out["meeting_join_url"] = meeting["joinWebUrl"]
    return {k: v for k, v in out.items() if v not in (None, "")}


def get_chat_text(d: dict) -> str:
    lines = [f"{d['name']} ({d.get('type')}) · created {fmt.short_time(d.get('created'))} · id {d['chat_id']}"]
    if d.get("last_message_at"):
        lines.append(f"Last message {fmt.short_time(d['last_message_at'])} UTC "
                     f"{d.get('last_message_from') or '?'}: {d.get('last_message_preview', '')}")
    lines.append(f"Members ({len(d['members'])}):")
    lines.append(fmt.render_members(d["members"]))
    if d.get("meeting_join_url"):
        lines.append(f"Meeting join link: {d['meeting_join_url']}")
    if d.get("web_url"):
        lines.append(f"Open in Teams: {d['web_url']}")
    return "\n".join(lines)


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
    out: dict = {"count": len(rows), "messages": [fmt.message(m, MAX_LIST_TEXT_CHARS) for m in rows]}
    if more:
        out["note"] = f"older messages exist beyond these {limit}: raise limit or use since"
    return out


def read_chat_messages_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} messages, oldest first (times UTC)",
                           fmt.render_messages(d["messages"]), d.get("note"))


async def get_chat_message_data(chat_id: str, message_id: str) -> dict:
    cid = _id(chat_id, "chat_id", "list_chats")
    mid = _id(message_id, "message_id", "read_chat_messages")
    return fmt.message(await graph.get(f"/chats/{cid}/messages/{mid}"), full=True)


def full_message_text(d: dict) -> str:
    lines = ["Message (times UTC)", fmt.render_messages([d], ids="all")]
    if isinstance(d.get("edited"), str):
        lines.append(f"Edited {fmt.short_time(d['edited'])}")
    if d.get("reply_to_id"):
        lines.append(f"Reply to message {d['reply_to_id']}")
    if d.get("web_url"):
        lines.append(f"Open in Teams: {d['web_url']}")
    return "\n".join(lines)


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
                }.items() if val})
    return {"query": q, "count": len(hits), "hits": hits}


def search_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} results for {d['query']!r} (times UTC)", fmt.render_hits(d["hits"]))


async def list_teams_data() -> list[dict]:
    rows, _ = await graph.get_all("/me/joinedTeams", {"$select": "id,displayName,description"}, 500)
    return [{k: v for k, v in {"team_id": t.get("id"), "name": t.get("displayName"),
                                "description": t.get("description")}.items() if v} for t in rows]


def list_teams_text(teams: list[dict]) -> str:
    return "\n".join(
        f"• {t.get('name')} · id {t.get('team_id')}"
        + (f" · {fmt._shorten(t['description'], 80)}" if t.get("description") else "")
        for t in teams) or "Not a member of any team."


async def list_channels_data(team_id: str) -> list[dict]:
    tid = _team(team_id)
    rows, _ = await graph.get_all(
        f"/teams/{tid}/channels", {"$select": "id,displayName,description,membershipType"}, 500)
    return [{k: v for k, v in {"channel_id": c.get("id"), "name": c.get("displayName"),
                                "description": c.get("description"),
                                "membership": c.get("membershipType")}.items() if v} for c in rows]


def list_channels_text(channels: list[dict]) -> str:
    return "\n".join(
        f"• {c.get('name')}"
        + (f" ({c['membership']})" if c.get("membership") not in (None, "standard") else "")
        + f" · id {c.get('channel_id')}"
        + (f" · {fmt._shorten(c['description'], 80)}" if c.get("description") else "")
        for c in channels) or "No visible channels."


async def list_team_members_data(team_id: str, limit: int) -> dict:
    tid = _team(team_id)
    rows, more = await graph.get_all(f"/teams/{tid}/members", None, _clamp(limit, 1, 1000))
    out: dict = {"count": len(rows), "members": [fmt.member(m) for m in rows]}
    if more:
        out["note"] = "more members exist: raise limit"
    return out


def members_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} members", fmt.render_members(d["members"]), d.get("note"))


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
            if replies:
                t["reply_count"] = len(replies)
                t["replies"] = [fmt.message(r, MAX_LIST_TEXT_CHARS) for r in replies[-20:]]
        threads.append(t)
    out: dict = {"count": len(threads), "threads": threads}
    if more:
        out["note"] = f"older posts exist beyond these {limit}"
    return out


def threads_text(d: dict) -> str:
    return fmt.with_header(
        f"{d['count']} posts, oldest first (times UTC; reply to a post with its post id; "
        f"get_channel_thread shows full text)",
        fmt.render_threads(d["threads"]), d.get("note"))


async def get_channel_thread_data(team_id: str, channel_id: str, message_id: str) -> dict:
    tid = _team(team_id)
    chid = _id(channel_id, "channel_id", "list_channels")
    mid = _id(message_id, "message_id", "read_channel_messages")
    base = f"/teams/{tid}/channels/{chid}/messages/{mid}"
    root, (replies, more) = await asyncio.gather(
        graph.get(base), graph.get_all(f"{base}/replies", {"$top": 50}, 500))
    replies.sort(key=lambda r: r.get("createdDateTime") or "")
    post = fmt.message(root, full=True)
    out: dict = {"post": post, "reply_count": len(replies),
                 "replies": [fmt.message(r, full=True) for r in replies]}
    if more:
        out["note"] = "only the first 500 replies are shown"
    return out


def thread_text(d: dict) -> str:
    post = dict(d["post"], replies=d["replies"], reply_count=d["reply_count"])
    body = fmt.render_threads([post])
    if d["post"].get("web_url"):
        body += f"\nOpen in Teams: {d['post']['web_url']}"
    return fmt.with_header("Thread (times UTC)", body, d.get("note"))


# ---------------------------------------------------------------- write functions
async def _resolve_user(user: str) -> dict:
    u = (user or "").strip()
    if not (GUID_RE.match(u) or ("@" in u and " " not in u)):
        raise ToolError("Give the person's email address (find_people returns it). "
                        "Names are not accepted, to avoid messaging the wrong person.")
    try:
        return await graph.get(f"/users/{quote(u, safe='@')}", {"$select": PERSON_FIELDS})
    except GraphError as e:
        if e.status == 404:
            raise ToolError(f"No user {u!r} in this tenant.") from None
        raise


def _when(m: dict) -> str:
    return f"{fmt.short_time(m.get('createdDateTime'))} UTC"


async def send_chat_message_data(chat_id: str, content: str, format: str, importance: str) -> str:
    cid = _id(chat_id, "chat_id", "list_chats")
    payload: dict = {"body": _body(content, format)}
    if importance not in ("normal", "high", "urgent"):
        raise ToolError("importance must be normal, high or urgent.")
    if importance != "normal":
        payload["importance"] = importance
    m = await graph.post(f"/chats/{cid}/messages", payload)
    return f"Sent to chat {cid} at {_when(m)} (message id {m.get('id')})"


async def send_direct_message_data(user: str, content: str, format: str) -> str:
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
    p = fmt.person(target)
    return (f"Sent to {p.get('name')} <{p.get('email', '?')}> at {_when(m)} "
            f"(chat {c['id']}, message id {m.get('id')})")


async def send_channel_message_data(team_id: str, channel_id: str, content: str,
                                    format: str, subject: str) -> str:
    tid = _team(team_id)
    chid = _id(channel_id, "channel_id", "list_channels")
    payload: dict = {"body": _body(content, format)}
    if subject.strip():
        payload["subject"] = subject.strip()[:255]
    m = await graph.post(f"/teams/{tid}/channels/{chid}/messages", payload)
    return f"Posted in channel {chid} at {_when(m)} (post id {m.get('id')})"


async def reply_to_channel_message_data(team_id: str, channel_id: str, message_id: str,
                                        content: str, format: str) -> str:
    tid = _team(team_id)
    chid = _id(channel_id, "channel_id", "list_channels")
    mid = _id(message_id, "message_id", "read_channel_messages")
    m = await graph.post(f"/teams/{tid}/channels/{chid}/messages/{mid}/replies",
                         {"body": _body(content, format)})
    return f"Replied to post {mid} at {_when(m)} (reply id {m.get('id')})"


async def create_online_meeting_data(subject: str, start: str, end: str) -> str:
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
    return (f"Meeting created: {m.get('subject')}, {fmt.short_time(m.get('startDateTime'))} to "
            f"{fmt.short_time(m.get('endDateTime'))} UTC\n"
            f"Join link: {m.get('joinWebUrl')}\n"
            f"No calendar invitation was sent: share the join link, e.g. with send_chat_message.")


# ---------------------------------------------------------------- read tools
@mcp.tool(**READ)
async def whoami(output: str = "text") -> str:
    """The Microsoft 365 account this server acts as, and the state of its sign-in."""
    return await _run(whoami_data(), whoami_text, output)


@mcp.tool(**READ)
async def find_people(query: str, limit: int = 10, output: str = "text") -> str:
    """Look up people in the directory by name or email: name, email, job
    title, department. Use the email with send_direct_message."""
    return await _run(find_people_data(query, limit), fmt.render_people, output)


@mcp.tool(**READ)
async def list_chats(query: str = "", limit: int = 25, output: str = "text") -> str:
    """The signed-in user's chats (1:1, group, meeting), most recently active
    first: name, last message, UNREAD flag and chat id. With query, keeps only
    chats whose title or member names contain it (case-insensitive)."""
    return await _run(list_chats_data(query, limit), list_chats_text, output)


@mcp.tool(**READ)
async def get_chat(chat_id: str, output: str = "text") -> str:
    """Details of one chat: full member list with emails and roles, meeting
    join link for meeting chats, link to open it in Teams."""
    return await _run(get_chat_data(chat_id), get_chat_text, output)


@mcp.tool(**READ)
async def read_chat_messages(chat_id: str, limit: int = 30, since: str = "",
                             include_system: bool = False, output: str = "text") -> str:
    """Recent messages in a chat as a transcript, oldest first. limit is 1-200.
    since (ISO 8601, e.g. 2026-11-13T08:00:00Z) returns only messages created or
    edited after that time. System events are left out unless include_system.
    Long messages end with [truncated] [id N]: get_chat_message gives the full text."""
    return await _run(read_chat_messages_data(chat_id, limit, since, include_system),
                      read_chat_messages_text, output)


@mcp.tool(**READ)
async def get_chat_message(chat_id: str, message_id: str, output: str = "text") -> str:
    """One chat message in full, with attachment links."""
    return await _run(get_chat_message_data(chat_id, message_id), full_message_text, output)


@mcp.tool(**READ)
async def search_messages(query: str, limit: int = 25, output: str = "text") -> str:
    """Full-text search across the signed-in user's chats and channel messages
    (Microsoft Search, KQL syntax, e.g. 'budget from:alice'). Each hit gives the
    chat id, or team and channel ids, and the message id to read more."""
    return await _run(search_messages_data(query, limit), search_text, output)


@mcp.tool(**READ)
async def list_teams(output: str = "text") -> str:
    """Teams the signed-in user is a member of, with their ids."""
    return await _run(list_teams_data(), list_teams_text, output)


@mcp.tool(**READ)
async def list_channels(team_id: str, output: str = "text") -> str:
    """Channels in a team that the signed-in user can see, with their ids."""
    return await _run(list_channels_data(team_id), list_channels_text, output)


@mcp.tool(**READ)
async def list_team_members(team_id: str, limit: int = 200, output: str = "text") -> str:
    """Members of a team with email and role (owner, guest)."""
    return await _run(list_team_members_data(team_id, limit), members_text, output)


@mcp.tool(**READ)
async def read_channel_messages(team_id: str, channel_id: str, limit: int = 20,
                                include_replies: bool = True, output: str = "text") -> str:
    """Recent posts in a channel (limit 1-100), oldest first, each with its
    post id and, if include_replies, its latest 20 replies indented below."""
    return await _run(read_channel_messages_data(team_id, channel_id, limit, include_replies),
                      threads_text, output)


@mcp.tool(**READ)
async def get_channel_thread(team_id: str, channel_id: str, message_id: str,
                             output: str = "text") -> str:
    """One channel post and all its replies, in full, with attachment links."""
    return await _run(get_channel_thread_data(team_id, channel_id, message_id), thread_text, output)


# ---------------------------------------------------------------- write tools
async def send_chat_message(chat_id: str, content: str, format: str = "text",
                            importance: str = "normal") -> str:
    """Send a message to an existing chat as the signed-in user. format 'text'
    sends the content as typed (line breaks kept); 'html' allows Teams HTML such
    as <b>, <i>, <a href>, <ul><li>, <br>. importance: normal, high or urgent.
    Confirm the chat and wording with the user first: messages can't be unsent from here."""
    return await _call(send_chat_message_data(chat_id, content, format, importance))


async def send_direct_message(user: str, content: str, format: str = "text") -> str:
    """Send a 1:1 message to a person by email address (from find_people) or
    user id, creating the 1:1 chat if it doesn't exist yet. Confirm the
    recipient and wording with the user first."""
    return await _call(send_direct_message_data(user, content, format))


async def send_channel_message(team_id: str, channel_id: str, content: str,
                               format: str = "text", subject: str = "") -> str:
    """Start a new post in a channel as the signed-in user, with an optional
    subject. Confirm the channel and wording with the user first."""
    return await _call(send_channel_message_data(team_id, channel_id, content, format, subject))


async def reply_to_channel_message(team_id: str, channel_id: str, message_id: str,
                                   content: str, format: str = "text") -> str:
    """Reply in the thread of a channel post (message_id = its post id) as the
    signed-in user. Confirm the wording with the user first."""
    return await _call(reply_to_channel_message_data(team_id, channel_id, message_id, content, format))


async def create_online_meeting(subject: str, start: str = "", end: str = "") -> str:
    """Create a Teams meeting and return its join link. start/end are ISO 8601
    with an offset (e.g. 2026-11-14T10:00:00+04:00); start defaults to now and
    end to start + 30 minutes. No calendar invitation is sent."""
    return await _call(create_online_meeting_data(subject, start, end))


if ENABLE_WRITE:
    for _fn in (send_chat_message, send_direct_message, send_channel_message,
                reply_to_channel_message, create_online_meeting):
        mcp.tool(annotations=WRITES, output_schema=None)(_fn)


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
