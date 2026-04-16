# Claude Co-author — Google Apps Script

An always-on Claude integration for Google Docs. Add `@claude <instruction>` as a
comment in any Google Doc and Claude will automatically reply, edit the document,
and add new comments — all within 5 minutes.

---

## How It Works

1. You write a comment in any Google Doc starting with `@claude`
2. A background trigger (every 5 min) scans all your recently active docs
3. Claude searches the web if needed, then acts on your instruction
4. Claude edits the document directly, adds new comments where needed, and
   replies to your comment explaining what it did

---

## Comment Syntax

`@claude` is case-insensitive and the separator after it is flexible:

```
@claude: fix the grammar in this paragraph
@Claude, make this more concise
@claude please rewrite this in a formal tone
@claudeCan you summarize this section?
```

---

## What Claude Can Do

- **Edit text** — replace, rewrite, fix, expand, or restructure any passage
- **Answer questions** — give feedback, analysis, or explanations as a comment reply
- **Backfill placeholders** — find and replace [VERIFY:...] tags with real information
  using web search; add new document comments for anything it can't verify
- **Continue a thread** — reply to Claude's reply with another @claude message and
  it will respond in context, seeing the full conversation history

---

## One-Time Setup

### Step 1 — Create the Script Project

1. Go to https://script.google.com
2. Click **New project** and name it **Claude Co-author**

### Step 2 — Add the Script Files

**appsscript.json**
- Click the gear icon in the left sidebar → check "Show appsscript.json manifest file"
- Replace its entire contents with the `appsscript.json` from this package

**Code.gs**
- Click the existing `Code.gs` file and replace its contents

**Claude.gs**
- Click **+** → Script → name it `Claude` → paste contents of `Claude.gs`

**Triggers.gs**
- Click **+** → Script → name it `Triggers` → paste contents of `Triggers.gs`

### Step 3 — Enable Advanced Services

In the left sidebar click **Services** (+):
1. Add **Drive API** — find under D, select v3, click Add
2. Add **Google Docs API** — find under G, select v1, click Add

> If saving appsscript.json gives a "duplicate service" error, remove the
> `dependencies` block from appsscript.json — the Services UI already added it.

### Step 4 — Set Your API Key

`promptApiKey` uses DocumentApp.getUi() which only works inside an open Google Doc,
not from the script editor. Use this temporary helper instead:

**4a.** Add this function temporarily to the bottom of `Triggers.gs`:

```javascript
function setApiKeyDirect() {
  PropertiesService.getUserProperties().setProperty(
    'claudeApiKey',
    'sk-ant-YOUR-KEY-HERE'  // paste your full key here
  );
  Logger.log('API key saved.');
}
```

**4b.** Replace `sk-ant-YOUR-KEY-HERE` with your key from https://console.anthropic.com

**4c.** Open `Triggers.gs` in the editor (the dropdown only shows functions from the
currently open file)

**4d.** Select `setApiKeyDirect` from the dropdown and click **Run**

**4e.** Confirm "API key saved." in the Execution log

**4f.** Delete the `setApiKeyDirect` function

### Step 5 — Start the Background Trigger

**5a.** With `Triggers.gs` open, select `setupTrigger` from the dropdown and click **Run**

**5b.** When Google warns "this app isn't verified", click **Advanced** →
**Go to [project name] (unsafe)** — expected for all personal scripts

**5c.** Confirm "Auto-polling activated" in the Execution log

The trigger now runs permanently every 5 minutes, even when your browser is closed.

---

## Running Immediately (Without Waiting 5 Minutes)

Open `Code.gs` in the script editor, select `processAllRecentDocs` from the
function dropdown, and click **Run**. This processes all pending @claude comments
in any doc modified in the last 10 minutes right now.

---

## Verifying the Trigger is Active

1. Go to https://script.google.com → open Claude Co-author
2. Click the clock icon in the left sidebar (Triggers)
3. You should see `processAllRecentDocs` listed with a time-based trigger

---

## File Reference

| File             | Purpose                                               |
|------------------|-------------------------------------------------------|
| Code.gs          | Comment scanning, placeholder extraction, edit logic  |
| Claude.gs        | Anthropic API call, web search, response parsing      |
| Triggers.gs      | Time trigger management, API key storage, status      |
| appsscript.json  | OAuth scopes and advanced service declarations        |

---

## Troubleshooting

**"Cannot call DocumentApp.getUi() from this context"**
You ran promptApiKey or a menu function directly from the script editor. Use the
temporary setApiKeyDirect workaround in Step 4. setupTrigger uses Logger.log and
runs fine from the editor.

**"Drive is not defined" or "Docs API error 403"**
One or both advanced services are not enabled. Go to Services in the left sidebar
and add Drive API v3 and Google Docs API v1.

**Function not in dropdown**
The dropdown only shows functions from the currently open file. Click the correct
.gs file in the left sidebar first.

**"Google hasn't verified this app"**
Expected for personal scripts. Click Advanced → Go to [project name] (unsafe).

**Duplicate Drive service error**
Remove the `dependencies` block from appsscript.json.

**Comment not being picked up**
- The file must be a native Google Doc — not a .docx. Use File → Save as Google Docs.
- Confirm the comment starts with @claude (any capitalization).
- Run processAllRecentDocs manually for immediate processing.

**Edits applied but wrong text replaced**
Claude uses the exact strings from the document as search keys. If the document
was edited between Claude reading it and applying the edit, the string may no
longer match. Just re-run.

**0 edits applied / string not found**
The text Claude targeted may span a paragraph break. Reply to the comment with a
more specific instruction targeting a shorter, single-paragraph passage.

---

## Security Notes

- API key stored in PropertiesService.getUserProperties() — tied to your Google
  account, not visible in code, not shared with anyone
- Script only accesses documents in your own Google Drive
- No data sent anywhere except the Anthropic API using your own key
