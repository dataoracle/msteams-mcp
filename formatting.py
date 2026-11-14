"""Turn Graph chat/channel objects into compact output for the model.

Two layers:
- normalisers (message, chat, person, member, ...) turn Graph JSON into small
  dicts, keeping only what is useful; these are also the format="json" output;
- renderers (render_*) turn those dicts into a compact, transcript-style text,
  the default output. Ids and URLs only appear where a follow-up call needs them.
"""
from __future__ import annotations

import json
import re
from collections import Counter
from datetime import datetime, timezone
from html.parser import HTMLParser

_BLOCK = {"p", "div", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre", "ul", "ol", "table"}


class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out: list[str] = []

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "br":
            self.out.append("\n")
        elif tag == "li":
            self.out.append("\n- ")
        elif tag == "img":
            self.out.append(f"[{a.get('alt') or 'image'}]")
        elif tag == "emoji":
            self.out.append(a.get("alt") or "")
        elif tag in _BLOCK:
            # <attachment> placeholders are dropped: attachments are listed separately.
            self.out.append("\n")

    def handle_endtag(self, tag):
        if tag in _BLOCK:
            self.out.append("\n")

    def handle_data(self, data):
        self.out.append(data)


def html_to_text(content: str) -> str:
    p = _TextExtractor()
    p.feed(content or "")
    p.close()
    return tidy("".join(p.out))


# Whitespace Teams puts in message bodies, including the invisible kinds that
# str.strip() keeps (zero-width space, word joiner, BOM).
_WS = " \t\u00a0\u2007\u202f\u200b\u2060\ufeff"


def tidy(text: str) -> str:
    """Trim every line and drop blank lines, so a transcript's continuation
    lines all share one indent whatever spacing the message started with."""
    lines = (ln.strip(_WS) for ln in (text or "").replace("\r\n", "\n").split("\n"))
    return "\n".join(ln for ln in lines if ln)


def strip_search_markup(s: str) -> str:
    """Search summaries wrap hits in <c0>..</c0> and use <ddd/> for ellipses."""
    return html_to_text((s or "").replace("<ddd/>", "…").replace("<c0>", "").replace("</c0>", ""))


def parse_time(value: str) -> datetime:
    """Parse ISO 8601 (with Z or offset, up to 7 fractional digits). No offset = UTC."""
    s = value.strip().replace("Z", "+00:00").replace("z", "+00:00")
    m = re.match(r"^(.*T\d\d:\d\d:\d\d)(\.\d+)?(.*)$", s)
    if m and m.group(2):
        s = m.group(1) + m.group(2)[:7] + m.group(3)
    dt = datetime.fromisoformat(s)
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)


def iso_z(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def utc(value: str | None) -> datetime | None:
    """Parse a Graph timestamp to an aware UTC datetime; None if missing or invalid."""
    try:
        return parse_time(value).astimezone(timezone.utc) if value else None
    except ValueError:
        return None


_t = utc


def short_time(value: str | None) -> str:
    """'2026-11-13T10:03:21.1234567Z' -> '2026-11-13 10:03' (UTC)."""
    t = _t(value)
    return t.strftime("%Y-%m-%d %H:%M") if t else (value or "?")


def _shorten(text: str, n: int) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= n else text[: n - 1].rstrip() + "…"


def who(identity: dict | None) -> str | None:
    """Display name from a Graph identitySet (user, application or device)."""
    if not identity:
        return None
    for k in ("user", "application", "device"):
        v = identity.get(k)
        if v:
            return v.get("displayName") or v.get("id")
    return None


# ---------------------------------------------------------------- normalisers
_REACTIONS = {"like": "👍", "heart": "❤️", "laugh": "😆", "surprised": "😮", "sad": "😢", "angry": "😡"}


def _reaction(r: dict) -> str:
    kind = r.get("reactionType") or ""
    if kind in _REACTIONS:
        return _REACTIONS[kind]
    if kind and not kind.isascii():  # newer API: reactionType is the emoji itself
        return kind
    return r.get("displayName") or kind or "?"


def message(m: dict, max_chars: int | None = None, full: bool = False) -> dict:
    """Normalise a chatMessage. full=True (single-item tools) keeps URLs,
    timestamps of edits, the replied-to id and longer quote previews."""
    body = m.get("body") or {}
    content = body.get("content") or ""
    text = html_to_text(content) if (body.get("contentType") or "").lower() == "html" else tidy(content)
    out: dict = {"message_id": m.get("id"), "created": m.get("createdDateTime"),
                 "from": who(m.get("from")), "text": text}
    if max_chars and len(text) > max_chars:
        out["text"] = text[:max_chars].rstrip() + "…"
        out["truncated"] = True
    mtype = m.get("messageType")
    if mtype and mtype != "message":
        out["type"] = mtype
        event = (m.get("eventDetail") or {}).get("@odata.type")
        if event:
            out["event"] = event.rsplit(".", 1)[-1].removesuffix("EventMessageDetail")
    if m.get("subject"):
        out["subject"] = m["subject"]
    if m.get("importance") not in (None, "normal"):
        out["importance"] = m["importance"]
    if full and m.get("replyToId"):
        out["reply_to_id"] = m["replyToId"]
    if m.get("lastEditedDateTime"):
        out["edited"] = m["lastEditedDateTime"] if full else True
    if m.get("deletedDateTime"):
        out["deleted"] = True
        out["text"] = ""
    files, cards, quoted = [], [], []
    for a in m.get("attachments") or []:
        ct = a.get("contentType") or ""
        if ct == "messageReference":  # a quoted reply
            try:
                ref = json.loads(a.get("content") or "{}")
            except ValueError:
                ref = {}
            q = {"from": who(ref.get("messageSender")),
                 "preview": _shorten(ref.get("messagePreview") or "", 300 if full else 60)}
            if full:
                q["message_id"] = ref.get("messageId") or a.get("id")
            quoted.append(q)
        elif ct.startswith("application/vnd.microsoft.card"):
            cards.append(ct.rsplit(".", 1)[-1])
        else:
            f = {"name": a.get("name") or ct or "file"}
            if full and a.get("contentUrl"):
                f["url"] = a["contentUrl"]
            files.append(f)
    if quoted:
        out["quoted"] = quoted
    if files:
        out["files"] = files
    if cards:
        out["cards"] = cards
    reactions = Counter(_reaction(r) for r in m.get("reactions") or [])
    if reactions:
        out["reactions"] = " ".join(f"{e}{n}" for e, n in reactions.items())
    if full and m.get("webUrl"):
        out["web_url"] = m["webUrl"]
    return out


def person(u: dict) -> dict:
    return {k: v for k, v in {
        "name": u.get("displayName"),
        "email": u.get("mail") or u.get("userPrincipalName"),
        "job_title": u.get("jobTitle"),
        "department": u.get("department"),
        "office": u.get("officeLocation"),
    }.items() if v}


def member(m: dict) -> dict:
    roles = [r for r in m.get("roles") or [] if r]
    return {k: v for k, v in {
        "name": m.get("displayName"),
        "email": m.get("email"),
        "roles": ",".join(roles) or None,
    }.items() if v}


def _names(names: list[str], cap: int = 5) -> str:
    return ", ".join(names[:cap]) + (f" +{len(names) - cap}" if len(names) > cap else "")


CHAT_TYPES = {"oneOnOne": "1:1", "group": "group", "meeting": "meeting"}


def chat(c: dict, me_id: str | None, preview_chars: int = 80, full: bool = False,
         max_members: int = 3) -> dict:
    others = [m.get("displayName") for m in c.get("members") or []
              if m.get("userId") != me_id and m.get("displayName")]
    ctype = CHAT_TYPES.get(c.get("chatType") or "", c.get("chatType"))
    name = c.get("topic") or (_names(others) if others else "(untitled chat)")
    out: dict = {"chat_id": c.get("id"), "name": name, "type": ctype}
    # Members only add information for titled group chats (otherwise the name is the members).
    if others and c.get("topic") and ctype != "1:1":
        if full or max_members > 0:
            out["members"] = _names(others, len(others) if full else max_members)
        else:
            out["members"] = f"{len(others)} others"
    prev = c.get("lastMessagePreview") or {}
    if prev:
        body = prev.get("body") or {}
        text = body.get("content") or ""
        if (body.get("contentType") or "").lower() == "html":
            text = html_to_text(text)
        out["last_message_at"] = prev.get("createdDateTime")
        out["last_message_from"] = who(prev.get("from"))
        out["last_message_preview"] = _shorten(text, preview_chars)
        read_at = (c.get("viewpoint") or {}).get("lastMessageReadDateTime")
        t_prev, t_read = _t(prev.get("createdDateTime")), _t(read_at)
        if t_prev and t_read:
            out["unread"] = t_prev > t_read
    else:
        out["last_updated"] = c.get("lastUpdatedDateTime")
    if full and c.get("webUrl"):
        out["web_url"] = c["webUrl"]
    return out


# ---------------------------------------------------------------- renderers
_PAD = "      "  # width of "HH:MM " so continuation lines align


def _message_body(d: dict, show_id: bool) -> str:
    if d.get("type"):
        text = f"({d.get('event') or d['type']})"
    else:
        text = d.get("text") or ""
    lines = [f'↪ {q.get("from") or "?"}: "{q.get("preview", "")}"' for q in d.get("quoted", [])]
    if d.get("subject"):
        lines.append(f"Subject: {d['subject']}")
    tail = []
    for f in d.get("files", []):
        tail.append(f"[file: {f['name']}{' ' + f['url'] if f.get('url') else ''}]")
    tail += ["[card]"] * len(d.get("cards", []))
    if d.get("reactions"):
        tail.append(d["reactions"])
    if d.get("importance"):
        tail.append(f"!{d['importance']}")
    if isinstance(d.get("edited"), str):
        tail.append(f"(edited {short_time(d['edited'])})")
    elif d.get("edited"):
        tail.append("(edited)")
    if d.get("deleted"):
        tail.append("(deleted)")
    if d.get("truncated"):
        tail.append("[truncated]")
    if show_id:
        tail.append(f"[id {d.get('message_id')}]")
    lines.append(" ".join(x for x in (text, " ".join(tail)) if x))
    return "\n".join(lines)


def render_messages(msgs: list[dict], ids: str = "truncated", pad: str = "",
                    day_headers: bool = True) -> str:
    """Transcript: a date line per day, 'HH:MM Sender: text' per message.
    Consecutive messages from one sender within 5 minutes are grouped.
    ids: 'truncated' (only where get_chat_message is useful), 'all' or 'none'."""
    lines: list[str] = []
    day = last_from = last_t = None
    for d in msgs:
        t = _t(d.get("created"))
        d_day = t.strftime("%Y-%m-%d") if t else "?"
        if day_headers and d_day != day:
            lines.append(f"{pad}{d_day}")
            last_from = None
        day = d_day
        sender = d.get("from") or "system"
        show_id = ids == "all" or (ids == "truncated" and d.get("truncated"))
        body = _message_body(d, bool(show_id)).replace("\n", "\n" + pad + _PAD)
        hm = t.strftime("%H:%M") if t else "--:--"
        if not day_headers:
            hm = t.strftime("%m-%d %H:%M") if t else hm
        grouped = (sender == last_from and t and last_t and 0 <= (t - last_t).total_seconds() <= 300)
        if grouped:
            lines.append(f"{pad}{_PAD}{body}")
        else:
            lines.append(f"{pad}{hm} {sender}: {body}")
        last_from, last_t = sender, t
    return "\n".join(lines)


def render_threads(threads: list[dict]) -> str:
    """Channel posts (with [post id] for replying) and their replies, indented."""
    lines: list[str] = []
    day = None
    for p in threads:
        t = _t(p.get("created"))
        d_day = t.strftime("%Y-%m-%d") if t else "?"
        if d_day != day:
            lines.append(d_day)
            day = d_day
        hm = t.strftime("%H:%M") if t else "--:--"
        body = _message_body(p, show_id=False).replace("\n", "\n" + _PAD)
        count = p.get("reply_count")
        extra = (f" ({count} {'reply' if count == 1 else 'replies'}"
                 f"{', latest 20 shown' if count > 20 and len(p.get('replies') or []) < count else ''})"
                 if count else "")
        lines.append(f"{hm} {p.get('from') or 'system'}: {body} [post {p.get('message_id')}]{extra}")
        if p.get("replies"):
            lines.append(render_messages(p["replies"], ids="none", pad="    ", day_headers=False))
    return "\n".join(lines)


def render_chats(chats: list[dict]) -> str:
    lines = []
    for c in chats:
        parts = [f"• {c['name']} ({c.get('type')})"]
        if c.get("members"):
            parts.append(f"members: {c['members']}")
        if c.get("last_message_at"):
            parts.append(f"{short_time(c['last_message_at'])} {c.get('last_message_from') or '?'}: "
                         f"{c.get('last_message_preview', '')}")
        elif c.get("last_updated"):
            parts.append(f"updated {short_time(c['last_updated'])}")
        if c.get("unread"):
            parts.append("UNREAD")
        parts.append(f"id {c['chat_id']}")
        lines.append(" · ".join(parts))
    return "\n".join(lines)


def render_people(people: list[dict]) -> str:
    return "\n".join(
        "• " + " · ".join(x for x in (f"{p.get('name')} <{p.get('email', '?')}>", p.get("job_title"),
                                     p.get("department"), p.get("office")) if x)
        for p in people) or "No matches."


def render_members(members: list[dict]) -> str:
    return "\n".join(
        f"• {m.get('name')}{' <' + m['email'] + '>' if m.get('email') else ''}"
        f"{' (' + m['roles'] + ')' if m.get('roles') else ''}" for m in members)


def render_hits(hits: list[dict]) -> str:
    lines = []
    for h in hits:
        where = (f"chat {h['chat_id']}" if h.get("chat_id") else
                 f"team {h.get('team_id')} channel {h.get('channel_id')}" if h.get("channel_id") else "")
        lines.append(f"• {short_time(h.get('created'))} {h.get('from') or '?'}: {h.get('summary', '')}"
                     f" · {where} msg {h.get('message_id')}".rstrip())
    return "\n".join(lines) or "No matches."


def with_header(header: str, body: str, note: str | None = None) -> str:
    return "\n".join(x for x in (header, body, f"({note})" if note else None) if x)
