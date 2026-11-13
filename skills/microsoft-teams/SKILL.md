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
| `list_chats(query?, limit?)` | Finding a chat_id; catching up (`unread` flag, last message preview) |
| `read_chat_messages(chat_id, limit?, since?)` | Reading a conversation (oldest first) |
| `get_chat_message(chat_id, message_id)` | One message in full, when a list entry says `truncated` |
| `get_chat(chat_id)` | Members and roles, meeting link of a meeting chat |
| `search_messages(query)` | "Where did someone mention X?": keyword search across chats and channels |
| `find_people(query)` | Resolving a name to an email/user_id |
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
     instead, then read around a hit with `read_chat_messages` (via `chat_id`)
     or `get_channel_thread` (via `team_id`/`channel_id`/`message_id`).
   - Several chats match: show the candidates (name, members, last activity)
     and ask which one. Don't guess.
2. **Read only what the question needs.** The default `limit` (30) is usually
   enough. For "what happened since yesterday", pass `since` (ISO 8601 with
   offset) rather than a big limit.
3. **Catch-up requests** ("anything I missed?"): call `list_chats()`, focus on
   `unread: true`, and read those chats with a small limit. Summarise per chat.

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
3. Send once. Report the result (`sent: true`, time). Never claim a message
   was sent if the tool returned an error.

Formatting: use `format="text"` for normal messages (line breaks are kept).
Use `format="html"` only when rich formatting is needed (`<b>`, `<i>`,
`<a href="…">`, `<ul><li>`, `<br>`). Teams doesn't render Markdown.

`create_online_meeting` only creates a join link. It doesn't invite anyone or
put anything in a calendar. Offer to post the link in a chat.

## Reading the data

- Message `text` is converted from Teams HTML: `[image]` and `[attachment]`
  mark inline content. `attachments` lists files (name, SharePoint URL) and
  quoted replies (`quoted_from`, `quoted_preview`).
- `from: null` with a `type`/`event` field is a system event (member added,
  call ended...). These are hidden by default (`include_system=false`).
- `deleted` means the message was deleted; its text is empty.
- Times are UTC (`Z`). Convert them to the user's timezone when you answer if
  you know it.
- A `note` field means more data exists than was returned. Mention it if it
  matters.

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
