---
name: conventions
description: Set up or update the user's personal mail and calendar conventions for axlook — which calendar each kind of event goes in (work, personal, family, social…), which account to send from in which context, how to file mail, tone and language of drafts. Use when the user asks to configure or personalise axlook/Outlook behaviour, says things like "from now on put X in my Y calendar" or "new mail to clients goes from my work address", or when no conventions exist yet and they create their first event or draft.
---

# Personal conventions (axlook)

axlook stores each user's own rules in one Markdown file: `~/.config/axlook/conventions.md`, or `$AXLOOK_CONFIG_DIR/conventions.md` if that is set. The server shows the file to Claude in its instructions, and `outlook_get_conventions` returns the current version. Two front-matter keys are also applied by the server itself whenever a call leaves them out:

- `default_calendar`: used when `outlook_create_event` gets no `calendar`.
- `default_from`: used when `outlook_create_draft` gets no `from`. Replies and forwards always answer from the account that received the original message, so don't promise rules for them.

Any other front-matter keys, such as `timezone`, are guidance only: the server shows them to Claude but never applies them.

## Setting up (first time, or "personalise axlook")

1. **Look first.** Call `outlook_get_conventions`. If conventions exist, show a short summary and ask what to change instead of starting over.
2. **Discover, don't guess.** Call `outlook_list_calendars` and `outlook_list_folders` to get the user's real accounts and calendar paths. Use those exact paths in the file.
3. **Interview briefly.** Ask at most 3–4 compact questions, offering their real calendars and accounts as choices:
   - **Calendars:** which calendar for work or professional events, which for personal ones (for themselves only), and which for social or family events (with a partner, family, friends). Which should be the default?
   - **Sending:** which account to use in which context, for example work, a shared mailbox for a project, or private.
   - **Optional:** filing rules ("receipts → Finance/Receipts"), folders never to touch, the language and tone of drafts, their timezone.

   Skip anything the user doesn't care about. Never invent rules they didn't state.
4. **Draft the file** using the template below. Keep it short, with one line per rule, and quote calendar paths exactly as listed.
5. **Show it and get approval.** Present the complete text and ask for confirmation or edits. Only then call `outlook_set_conventions` with the full content.
6. **Confirm.** Report where it was saved. `default_calendar` and `default_from` work immediately. The rule text reaches the server's instructions at the next client restart; until then, read it with `outlook_get_conventions`.

## Updating a rule

When the user states a lasting preference ("from now on…", "always…", "social events go in Family"):

1. Read the current file with `outlook_get_conventions`.
2. Change only the affected line, keeping everything else.
3. Show the changed lines and ask for confirmation.
4. Save the whole file with `outlook_set_conventions`, unchanged except for the approved lines. The previous version is kept as `conventions.md.bak`.

A one-off instruction ("put *this* one in Personal") is not a convention. Just follow it for that call.

## Template

```markdown
---
default_calendar: <account>/<group>/<calendar used when unsure>
default_from: <address used for new drafts when unsure>
---
# Calendars
- Work and professional (meetings, students, conferences, colleagues) → `<work calendar path>`
- Personal, for me only (appointments, sport, errands) → `<personal calendar path>`
- Social, with my partner, family or friends → `<shared calendar path>`
- Never create events in: <calendars to leave alone>
- If unclear: ask, suggesting the most likely calendar first.

# Sending
- Work mail from `<work address>`; project/conference mail from `<shared mailbox>`; private mail from `<private address>`.

# Filing and drafts (optional)
- <folder rules, tone, language, signature habits>

# Time
- Times in <Region/City>.
```

## Notes

- The file belongs to the user. Store only what they asked for: no message contents, no third-party personal data beyond what a rule needs.
- Keep it under about 8,000 characters. Rules, not notes.
- Users can also edit the file by hand. It's plain Markdown.
