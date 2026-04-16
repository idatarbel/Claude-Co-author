# Task Log

<!-- Automatically maintained by Claude Code. Do not edit task numbers. -->
<!-- Use /audit purge N to remove items 1 through N after confirming they're good. -->

- [Pending QA] 1. **Restructure repo for Google Docs + online Word documents** — Moved Apps Script files (Claude.gs, Code.gs, Triggers.gs, appsscript.json) into `Google Docs/` via `git mv` (history preserved). Created `Word Online/` with a full Office Add-in: `manifest.xml`, `taskpane.html`, `taskpane.css`, `Code.js` (mirrors Code.gs + Triggers.gs), and `Claude.js` (mirrors Claude.gs, same system prompt and JSON contract, uses `anthropic-dangerous-direct-browser-access: true`). Rewrote `README.md` to document both integrations with separate setup sections and a unified troubleshooting block. The Word version reads comments via `document.getComments()`, applies edits via `body.search().insertText('Replace')`, adds anchored comments via `range.insertComment()`, and replies via `comment.reply()` — all WordApi 1.4+. Auto-polling uses `setInterval(5 min)` while the task pane is open (Office Add-ins can't run while the doc is closed — this is called out in the README).
