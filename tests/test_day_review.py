"""Offline tests for the "review my day" path: read_chat_messages since,
list_chats since, and whitespace in message bodies. Microsoft Graph and
Entra ID are mocked; nothing leaves the machine.

Run:  python tests/test_day_review.py      (or: pytest tests/)
Pass -v to print sample output.
"""
import asyncio
import os
import sys
import tempfile
from pathlib import Path
from urllib.parse import parse_qsl

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import httpx  # noqa: E402

from vault import Vault, load_key, new_key_text, vault_context  # noqa: E402

TENANT, CLIENT = "contoso.onmicrosoft.com", "11111111-2222-3333-4444-555555555555"
_work = Path(tempfile.mkdtemp())
_key = new_key_text()
os.environ.update(MS_TENANT_ID=TENANT, MS_CLIENT_ID=CLIENT, MCP_TOKEN="test-token",
                  MS_CLIENT_SECRET="test-secret", VAULT_KEY=_key, KEEPALIVE_HOURS="0",
                  VAULT_PATH=str(_work / "vault.json.enc"))
Vault(_work / "vault.json.enc", load_key(_key), vault_context(TENANT, CLIENT)).save(
    {"refresh_token": "RT", "user": {"id": "me-id", "name": "Me", "upn": "me@contoso.example"}})

import formatting as fmt  # noqa: E402
import server  # noqa: E402
from fastmcp import Client  # noqa: E402

SINCE = "2026-11-12T23:00:00Z"   # 00:00 Europe/Madrid on 2026-11-13
VERBOSE = "-v" in sys.argv
GROUP = "19:group@thread.v2"
requests: list[tuple[str, dict]] = []


def user(name: str) -> dict:
    return {"user": {"displayName": name, "id": name.lower()}}


def msg(mid, created, sender, html, modified=None, edited=None, reactions=0):
    m = {"id": mid, "createdDateTime": created, "lastModifiedDateTime": modified or created,
         "messageType": "message", "from": user(sender), "attachments": [],
         "body": {"contentType": "html", "content": html},
         "reactions": [{"reactionType": "like"}] * reactions}
    if edited:
        m["lastEditedDateTime"] = edited
    return m


MESSAGES = [
    # Two weeks old. Someone reacted today, which moved lastModifiedDateTime only.
    msg("100", "2026-10-29T09:00:00Z", "Dana", "<p>old announcement</p>",
        modified="2026-11-13T07:00:00Z", reactions=2),
    # Three days old, edited today.
    msg("200", "2026-11-10T15:30:00Z", "Carol", "<p>budget is 40k</p>",
        modified="2026-11-13T08:15:00Z", edited="2026-11-13T08:15:00Z"),
    # Today. The body starts with a blank paragraph and indented lines.
    msg("300", "2026-11-13T06:05:00Z", "Alice",
        "<p>&nbsp;</p><p>&nbsp; (Entra admin center → App registrations)</p>"
        "<p>​ then Certificates &amp; secrets&nbsp;</p>"),
    # Today, same sender two minutes later, plain text with stray whitespace.
    {**msg("301", "2026-11-13T06:07:00Z", "Alice", ""),
     "body": {"contentType": "text", "content": "\r\n   new secret is sealed\n\n    restart done  \n"}},
    # Old and untouched: Graph's own filter drops it.
    msg("050", "2026-10-01T10:00:00Z", "Dana", "<p>ancient</p>"),
]

CHATS = []
for i, (name, last) in enumerate([
        ("Bob Smith", "2026-11-13T09:40:00Z"),
        ("Ops war room", "2026-11-13T08:20:00Z"),
        ("Carol Diaz", "2026-11-13T07:05:00Z"),
        ("Budget review", "2026-11-12T23:30:00Z"),
        ("Dana Lee", "2026-11-12T22:10:00Z"),     # before the cutoff
        ("Old project", "2026-11-01T12:00:00Z"),
        ("Erin Park", "2026-10-20T12:00:00Z"),
]):
    topic = name if name in ("Ops war room", "Budget review", "Old project") else None
    others = (["Alice", "Bob", "Carol", "Dana", "Erin", "Frank"] if topic else [name])
    CHATS.append({
        "id": f"19:c{i}@thread.v2", "chatType": "group" if topic else "oneOnOne", "topic": topic,
        "members": [{"userId": "me-id", "displayName": "Me"}]
                   + [{"userId": o.lower(), "displayName": o} for o in others],
        "lastMessagePreview": {"createdDateTime": last, "from": user(others[0]),
                               "body": {"contentType": "text", "content": f"msg in {name}"}},
        "viewpoint": {"lastMessageReadDateTime": last},
    })
CHAT_PAGE = 3


def graph(req: httpx.Request) -> httpx.Response:
    p, q = req.url.path, dict(req.url.params)
    requests.append((p, q))
    if p == "/v1.0/me":
        return httpx.Response(200, json={"id": "me-id", "displayName": "Me"})
    if p == "/v1.0/me/chats":
        start = int(q.get("$skiptoken", 0))
        page = {"value": CHATS[start:start + CHAT_PAGE]}
        if start + CHAT_PAGE < len(CHATS):
            page["@odata.nextLink"] = (f"https://graph.microsoft.com/v1.0/me/chats"
                                       f"?$skiptoken={start + CHAT_PAGE}")
        return httpx.Response(200, json=page)
    if p == f"/v1.0/chats/{GROUP}/messages":
        rows = list(MESSAGES)
        flt = q.get("$filter", "")
        if flt:  # behave like Graph: only lastModifiedDateTime gt <time>
            assert flt.startswith("lastModifiedDateTime gt "), flt
            cut = flt.split(" gt ", 1)[1]  # same Z format as the data: compare as text
            rows = [m for m in rows if m["lastModifiedDateTime"] > cut]
        key = q.get("$orderby", "").split()[0]
        rows.sort(key=lambda m: m[key], reverse=True)
        return httpx.Response(200, json={"value": rows})
    return httpx.Response(404, json={"error": {"code": "NotFound", "message": p}})


def token(req: httpx.Request) -> httpx.Response:
    form = dict(parse_qsl(req.content.decode()))
    assert form["client_secret"] == "test-secret"
    return httpx.Response(200, json={"access_token": "AT", "expires_in": 3600, "refresh_token": "RT"})


server.tokens._http = httpx.AsyncClient(transport=httpx.MockTransport(token))
server.graph._http = httpx.AsyncClient(base_url="https://graph.microsoft.com/v1.0",
                                       transport=httpx.MockTransport(graph))


async def call(name: str, **args) -> str:
    async with Client(server.mcp) as c:
        res = await c.call_tool(name, args, raise_on_error=False)
    text = res.content[0].text
    if VERBOSE:
        print(f"\n--- {name}({', '.join(f'{k}={v!r}' for k, v in args.items())})\n{text}")
    return text


# ------------------------------------------------------------------ tests
def test_since_ignores_reaction_only_updates():
    out = asyncio.run(call("read_chat_messages", chat_id=GROUP, since=SINCE))
    assert "old announcement" not in out and "2026-10-29" not in out, out
    assert out.splitlines()[0] == "3 messages since 2026-11-12T23:00Z, oldest first (times UTC)", out


def test_since_keeps_old_edited_message_with_its_real_date():
    lines = asyncio.run(call("read_chat_messages", chat_id=GROUP, since=SINCE)).splitlines()
    # The date line above it is the day it was written, not the day it was edited.
    assert lines[1:3] == ["2026-11-10",
                          "15:30 Carol: budget is 40k (edited 2026-11-13 08:15)"], lines
    assert lines[3] == "2026-11-13", lines


def test_since_mode_modified_includes_reaction_only_updates():
    out = asyncio.run(call("read_chat_messages", chat_id=GROUP, since=SINCE, since_mode="modified"))
    lines = out.splitlines()
    assert lines[0].startswith("4 messages") and lines[1:3] == [
        "2026-10-29", "09:00 Dana: old announcement 👍2"], out
    bad = asyncio.run(call("read_chat_messages", chat_id=GROUP, since=SINCE, since_mode="x"))
    assert "since_mode must be" in bad, bad


def test_since_limit_counts_kept_messages_only():
    out = asyncio.run(call("read_chat_messages", chat_id=GROUP, since=SINCE, limit=2))
    # The reaction-only message must not use up one of the two slots.
    assert out.splitlines()[0].startswith("2 messages") and "old announcement" not in out, out
    # Default keep="newest": the two latest by creation time; the footer says the
    # earliest one was dropped.
    assert "budget is 40k" not in out and "06:05 Alice" in out, out
    assert out.endswith("(1 earlier message since the cutoff omitted: raise limit, "
                        "or pass keep='oldest' to read from the cutoff forward)"), out


def test_since_keep_oldest_reads_forward_from_cutoff():
    out = asyncio.run(call("read_chat_messages", chat_id=GROUP, since=SINCE, limit=2, keep="oldest"))
    lines = out.splitlines()
    assert lines[1:3] == ["2026-11-10", "15:30 Carol: budget is 40k (edited 2026-11-13 08:15)"], out
    assert "06:05 Alice" in out and "new secret is sealed" not in out, out
    assert out.endswith("(1 later message since the cutoff omitted: raise limit, "
                        "or pass keep='newest' for the latest ones)"), out
    bad = asyncio.run(call("read_chat_messages", chat_id=GROUP, since=SINCE, keep="middle"))
    assert "keep must be" in bad, bad


def test_message_starting_with_line_break_keeps_one_indent():
    out = asyncio.run(call("read_chat_messages", chat_id=GROUP, since=SINCE))
    tail = out.splitlines()[4:]
    assert tail == [
        "06:05 Alice: (Entra admin center → App registrations)",
        "      then Certificates & secrets",
        "      new secret is sealed",
        "      restart done",
    ], tail
    assert fmt.tidy("\n​  a  \n\n\t b \n") == "a\nb"


def test_list_chats_since_stops_at_cutoff():
    requests.clear()
    out = asyncio.run(call("list_chats", since=SINCE))
    lines = out.splitlines()
    assert lines[0] == ("4 chats active since 2026-11-12T23:00Z (cutoff reached), "
                        "most recent first (times UTC)"), out
    assert [ln.split(" (")[0] for ln in lines[1:]] == [
        "• Bob Smith", "• Ops war room", "• Carol Diaz", "• Budget review"], out
    # Stopped on page 2 (Dana Lee is older): page 3 was never requested.
    assert len([r for r in requests if r[0] == "/v1.0/me/chats"]) == 2, requests
    # Titled group chats show 3 member names by default, then +N.
    assert "members: Alice, Bob, Carol +3" in lines[2], lines[2]


def test_list_chats_since_limit_reached():
    out = asyncio.run(call("list_chats", since=SINCE, limit=2))
    assert out.splitlines()[0] == ("2 chats active since 2026-11-12T23:00Z "
                                   "(limit reached, more chats may exist), most recent first (times UTC)"), out
    # A limit equal to the number of active chats is still a complete list.
    out = asyncio.run(call("list_chats", since=SINCE, limit=4))
    assert "(cutoff reached)" in out.splitlines()[0], out


def test_list_chats_max_members():
    out = asyncio.run(call("list_chats", since=SINCE, max_members=0))
    assert "members: 6 others" in out, out
    out = asyncio.run(call("list_chats", since=SINCE, max_members=5))
    assert "members: Alice, Bob, Carol, Dana, Erin +1" in out, out


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_")]
    for name, fn in tests:
        fn()
        print(f"ok  {name}")
    print(f"{len(tests)} passed")
