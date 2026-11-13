"""Turn Graph chat/channel objects into compact, readable dicts for the model."""
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
        elif tag == "attachment":
            self.out.append("[attachment]")
        elif tag in _BLOCK:
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
    text = "".join(p.out).replace("\u00a0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


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


def who(identity: dict | None) -> str | None:
    """Display name from a Graph identitySet (user, application or device)."""
    if not identity:
        return None
    for k in ("user", "application", "device"):
        v = identity.get(k)
        if v:
            return v.get("displayName") or v.get("id")
    return None


def _attachment(a: dict) -> dict:
    ct = a.get("contentType") or ""
    if ct == "messageReference":  # a quoted reply
        try:
            ref = json.loads(a.get("content") or "{}")
        except ValueError:
            ref = {}
        return {"quoted_message_id": ref.get("messageId") or a.get("id"),
                "quoted_from": who(ref.get("messageSender")),
                "quoted_preview": (ref.get("messagePreview") or "")[:300]}
    if ct.startswith("application/vnd.microsoft.card"):
        return {"type": "card", "card_type": ct.rsplit(".", 1)[-1], "name": a.get("name")}
    return {"name": a.get("name"), "type": ct or None, "url": a.get("contentUrl")}


def message(m: dict, max_chars: int | None = None) -> dict:
    body = m.get("body") or {}
    content = body.get("content") or ""
    text = html_to_text(content) if (body.get("contentType") or "").lower() == "html" else content.strip()
    out: dict = {
        "message_id": m.get("id"),
        "created": m.get("createdDateTime"),
        "from": who(m.get("from")),
        "text": text,
    }
    if max_chars and len(text) > max_chars:
        out["text"] = text[:max_chars] + " […]"
        out["truncated"] = True
    mtype = m.get("messageType")
    if mtype and mtype != "message":
        out["type"] = mtype
        event = (m.get("eventDetail") or {}).get("@odata.type")
        if event:
            out["event"] = event.rsplit(".", 1)[-1]
    if m.get("subject"):
        out["subject"] = m["subject"]
    if m.get("importance") not in (None, "normal"):
        out["importance"] = m["importance"]
    if m.get("replyToId"):
        out["reply_to_id"] = m["replyToId"]
    if m.get("lastEditedDateTime"):
        out["edited"] = m["lastEditedDateTime"]
    if m.get("deletedDateTime"):
        out["deleted"] = m["deletedDateTime"]
        out["text"] = ""
    atts = [_attachment(a) for a in m.get("attachments") or []]
    if atts:
        out["attachments"] = atts
    mentions = sorted({x.get("mentionText") for x in m.get("mentions") or [] if x.get("mentionText")})
    if mentions:
        out["mentions"] = mentions
    reactions = Counter((r.get("displayName") or r.get("reactionType") or "?")
                        for r in m.get("reactions") or [])
    if reactions:
        out["reactions"] = dict(reactions)
    if m.get("webUrl"):
        out["web_url"] = m["webUrl"]
    return out


def person(u: dict) -> dict:
    return {k: v for k, v in {
        "name": u.get("displayName"),
        "email": u.get("mail") or u.get("userPrincipalName"),
        "upn": u.get("userPrincipalName"),
        "user_id": u.get("id"),
        "job_title": u.get("jobTitle"),
        "department": u.get("department"),
        "office": u.get("officeLocation"),
    }.items() if v}


def member(m: dict) -> dict:
    return {k: v for k, v in {
        "name": m.get("displayName"),
        "email": m.get("email"),
        "user_id": m.get("userId"),
        "roles": m.get("roles") or None,
    }.items() if v}


def chat(c: dict, me_id: str | None, preview_chars: int = 200) -> dict:
    others = [m.get("displayName") for m in c.get("members") or []
              if m.get("userId") != me_id and m.get("displayName")]
    name = c.get("topic") or ", ".join(others) or "(untitled chat)"
    out: dict = {"chat_id": c.get("id"), "name": name, "type": c.get("chatType")}
    if others:
        out["members"] = others if len(others) <= 12 else others[:12] + [f"+{len(others) - 12} more"]
    prev = c.get("lastMessagePreview") or {}
    if prev:
        body = prev.get("body") or {}
        text = body.get("content") or ""
        if (body.get("contentType") or "").lower() == "html":
            text = html_to_text(text)
        out["last_message_at"] = prev.get("createdDateTime")
        out["last_message_from"] = who(prev.get("from"))
        out["last_message_preview"] = text[:preview_chars] + ("…" if len(text) > preview_chars else "")
        read_at = (c.get("viewpoint") or {}).get("lastMessageReadDateTime")
        if read_at and prev.get("createdDateTime"):
            try:
                out["unread"] = parse_time(prev["createdDateTime"]) > parse_time(read_at)
            except ValueError:
                pass
    else:
        out["last_updated"] = c.get("lastUpdatedDateTime")
    if c.get("webUrl"):
        out["web_url"] = c["webUrl"]
    return out
