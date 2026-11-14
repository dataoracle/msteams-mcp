---
name: microsoft-teams
description: Read and send Microsoft Teams chats and channel messages through the msteams-mcp server, acting as the signed-in user. Use whenever the user asks about their Teams chats, what someone said or asked them, catching up on unread conversations, finding a message, a channel discussion, a team's members, or wants to send a Teams message, reply in a channel, or create a Teams meeting link.
---

# Microsoft Teams (msteams-mcp)

The msteams-mcp server acts in Microsoft Teams **as the signed-in user**,
through delegated Microsoft Graph permissions. Anything it reads, the user
could read in Teams. Anything it sends appears in Teams as written by the
user personally.

## Tool names

Clients add their own prefix (for example `mcp__msteams_mcp__list_chats`).
This skill uses the bare names.

| Tool | Use for |
|---|---|
| `list_chats(query?, limit?, since?)` | Finding a chat id; catching up (UNREAD flag, last message); every chat active since a time |
| `read_chat_messages(chat_id, limit?, since?)` | Reading a conversation (oldest first); what was said since a time |
| `get_chat_message(chat_id, message_id)` | One message in full, when a transcript line ends with `[truncated] [id N]` |
| `get_chat(chat_id)` | Members with emails and roles, meeting link of a meeting chat |
| `search_messages(query)` | "Where did someone mention X?": keyword search across chats and channels |
| `find_people(query)` | Resolving a name to an email |
| `list_teams` → `list_channels(team_id)` → `read_channel_messages(...)` | Channel conversations |
| `get_channel_thread(team_id, channel_id, message_id)` | A whole channel thread |
| `list_team_members(team_id)` | Who is in a team |
| `whoami()` | Which account the server acts as; sign-in status |
| `send_chat_message`, `send_direct_message`, `send_channel_message`, `reply_to_channel_message`, `create_online_meeting` | Writing. May be absent if the server is read-only |

## Workflow

### Reading

1. **Find the conversation.**
   - With a person or chat title, call `list_chats(query="<name or title>")`.
     It matches chat titles and member names. A 1:1 chat is named after the
     other person.
   - With a topic or keyword ("the invoice thread"), call `search_messages`
     instead, then read around a hit with `read_chat_messages` (its `chat`
     id) or `get_channel_thread` (its `team`, `channel` and `msg` ids).
   - Several chats match: show the candidates (name, members, last activity)
     and ask which one. Don't guess.
2. **Read only what the question needs.** The default `limit` (30) is usually
   enough. For "what happened since yesterday", pass `since` (ISO 8601 with
   offset) rather than a big limit.
3. **Catch-up requests** ("anything I missed?"): call `list_chats()`, focus on
   the chats marked UNREAD, and read those with a small limit. Summarise per chat.
4. **Reviewing a period** ("what happened today?"): convert the start to UTC
   (for example 00:00 Europe/Madrid is 23:00Z the day before). Then:
   - Call `list_chats(since=...)`. If the header says "cutoff reached", the list is
     complete. If it says "limit reached", raise `limit` and call it again.
   - Call `read_chat_messages(chat_id, since=...)` for each chat.
   - An older message may appear under its own earlier date line, marked
     `(edited <time>)`. It was edited during the period, and is not new. Say so if
     it matters.
   - Messages that only received reactions are left out. Pass
     `since_mode="modified"` to include them.

### Sending (only if the write tools are present)

Messages are sent as the user and can't be unsent from here. Before **every**
send, reply, or meeting creation:

1. Resolve the exact target:
   - a chat_id from `list_chats`;
   - or the person's email from `find_people` (`send_direct_message` refuses
     plain names);
   - or a `team_id`/`channel_id`.
2. Show the user the recipient (chat name or person and email, or team and
   channel) and the exact text, then **wait for explicit confirmation**. A
   general instruction earlier in the conversation isn't confirmation for a
   specific message, unless the user said to send without asking.
3. Send once. Report the result line ("Sent to … at …"). Never claim a message
   was sent if the tool returned an error.

Formatting: use `format="text"` for normal messages (line breaks are kept).
Use `format="html"` only when rich formatting is needed (`<b>`, `<i>`,
`<a href="…">`, `<ul><li>`, `<br>`). Teams doesn't render Markdown.

`create_online_meeting` only creates a join link. It doesn't invite anyone or
put anything in a calendar. Offer to post the link in a chat.

## Reading the data

Read tools return compact text. A chat reads as a transcript:

```
30 messages, oldest first (times UTC)
2026-11-13
10:01 Alice Johnson: first message [file: Q4.xlsx] 👍2
      second message from Alice, sent within 5 minutes
10:20 Bob Smith: ↪ Alice Johnson: "first message"
      reply text (edited)
10:31 Carol Diaz: a very long message cut off here… [truncated] [id 1731492660000]
(older messages exist beyond these 30: raise limit or use since)
```

- **Lines.** A date line starts each day. An indented line without a time
  continues the previous sender's message, or is a further message from them
  sent within 5 minutes.
- **Markers:**
  - `↪ Name: "…"` is a quoted reply.
  - `[file: name]` is an attachment and `[card]` is an adaptive card; their links
    appear only in the full-message tools.
  - `👍2` are reactions, and `!high` / `!urgent` the importance.
  - `(edited)` and `(deleted)` mean what they say. `(edited 2026-11-13 08:15)`
    marks an older message returned by `since` because it was edited after the
    cutoff.
  - `[image]` marks an inline image.
- **Ids are shown only where you need them:**
  - chat lists end with `id <chat id>`;
  - channel posts end with `[post <id>]`, which you use to reply or to open the thread;
  - truncated messages end with `[truncated] [id N]`: call `get_chat_message` with that id;
  - in channels, `get_channel_thread` gives the full text.
- `system: (membersAdded)` and similar lines are system events. They are hidden
  unless `include_system=true`.
- **Times** are UTC. Convert them to the user's timezone when you answer, if you
  know it.
- **Notes.** A last line in parentheses means more data exists than was
  returned. Mention it if it matters.
- **Structured data.** `output="json"` returns the same data as JSON. Use it
  only if you need exact fields, since it costs more tokens.

## Errors

- "Not signed in" / "sign-in is no longer valid … msteams-mcp-authorize": the
  server's Microsoft sign-in needs renewing. Tell the user to run
  `sudo msteams-mcp-authorize` on the server, then stop. Retrying won't help.
- "rejected the app's client secret": the app's secret is wrong or expired.
  Tell the user to seal a new secret on the server (the message includes the
  command), then stop. No new sign-in is needed.
- 403 Forbidden: the user (or the app's permissions) can't access that
  chat/channel, for example a private channel they're not in. Say so.
- 404 Not found: wrong id, or the user isn't a member. Re-list rather than
  guessing ids.
- Search returns nothing for a recent message: the index can lag a few
  minutes. Fall back to `list_chats` + `read_chat_messages`.

## Data handling

Chats are private conversations. Quote only what the question needs, don't
copy chat content into other tools, files or messages unless the user asks,
and never forward one person's messages to someone else without the user's
explicit instruction.
