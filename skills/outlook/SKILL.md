---
name: outlook
description: Use the axlook MCP tools (outlook_*) whenever the user asks anything about their email or calendars — reading, finding, summarizing or triaging mail, inboxes, senders, threads, attachments, folders, drafts, replies, forwards, moving/flagging/archiving messages, or schedules, meetings, events, availability and calendars. They drive the user's running New Outlook for Mac, which holds all their accounts (Exchange / Microsoft 365, Gmail, IMAP, shared mailboxes). Use them even when the user does not say "Outlook", unless they explicitly name another mail or calendar source.
---

# Outlook mail and calendars (axlook)

The `axlook` MCP server controls the user's **New Outlook for Mac** through macOS Accessibility. It sees every account Outlook shows. Use it for every mail and calendar question unless the user names another source.

## Personal conventions come first

The user may have saved their own rules: which calendar each kind of event goes in, which account to send from, how to file mail. They appear in the server instructions, and `outlook_get_conventions` returns the current version.

- **Follow them.** In particular, pass the `calendar` (`outlook_create_event`) and the `from` (`outlook_create_draft`) they prescribe.
- **An explicit request wins for that call.** If the user says "put this one in Personal", do that and don't change the conventions.
- **No conventions yet:** if the user creates their first event or draft, or states a lasting preference ("from now on…"), offer to set them up with the `conventions` skill.

## Workflow

1. **Discover first, never guess.** Call `outlook_list_folders` to get paths such as `you@example.com/Inbox` or `All Accounts/Inbox` (the unified view), and `outlook_list_calendars` for calendar paths.
2. **Find messages.** Use `outlook_list_messages` for a folder, or `outlook_search_messages` with Outlook syntax: `from:`, `subject:`, `hasattachments:yes`. Search defaults to all accounts. Both return a `ref` for each message.
3. **Act on a message with its `ref`.**
   - `outlook_read_message` reads it and keeps it unread unless asked.
   - `outlook_get_message_source` returns the full headers and HTML.
   - `outlook_save_attachments` saves attachments.
   - `outlook_move_message`, `outlook_mark_read`, `outlook_set_flag`, `outlook_archive_message` and `outlook_delete_message` change it. Delete moves it to Deleted Items.
4. **Conversation rows.** A row with `conversationCount > 1` is a whole thread. The first call on it expands it and returns guidance. List the folder again and use the ref of the specific message.
5. **Drafts.** `outlook_create_draft`, `outlook_create_reply_draft` and `outlook_create_forward_draft` save to Drafts and **never send**. Tell the user the draft is ready for review. Pass `from` to `outlook_create_draft` to choose the sending account. Replies and forwards always use the account that received the original message: the result reports it as `from`. If that contradicts a convention, tell the user.
6. **Calendar.**
   - `outlook_list_events` takes a `start` and `end` date (`YYYY-MM-DD`, at most about two months).
   - `outlook_get_event` returns one event; `full=true` adds attendees and description.
   - `outlook_create_event`, `outlook_update_event` and `outlook_delete_event` change events. Times are local wall-clock values such as `2026-11-20T10:00`.
   - Always pass `calendar` when creating, chosen from the user's conventions. If it's left out, the server uses the conventions' `default_calendar`, or else whichever calendar Outlook used last.

## Safe operation

- Prefer read-only tools for searches, summaries and anything ambiguous.
- Confirm moves, archives, deletions, bulk flagging and calendar changes the user did not clearly ask for.
- Meetings with attendees are read-only. Ask the user to accept or decline invitations in Outlook.
- Share only the message content the request needs.

## Practicalities

- Calls take 2–15 seconds and run one at a time. Avoid needless calls.
- Some tools bring Outlook to the front for a few seconds, then return focus.
- **When a call fails, check `outlook_status` first:**
  - `screenLocked: true`: the Mac must be unlocked.
  - `trusted: false`: the app running Claude needs Accessibility permission, in System Settings → Privacy & Security → Accessibility. Restart that app afterwards.
  - `outlookRunning: false`: start Outlook.
- A stale ref means the list changed: list or search again.
- If the user is typing a draft in Outlook, tools refuse to navigate away. Ask them to finish or close it.

## Response style

Name the account when it matters, then give sender, subject, local time and a short gist. Say so when results cover only some folders or accounts.
