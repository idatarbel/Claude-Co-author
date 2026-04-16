# Claude Co-author

An always-on Claude integration for word processors. Add `@claude <instruction>`
as a comment in a document and Claude will reply, edit the document, and add
new comments — automatically.

This repo ships two implementations:

| Directory         | Platform                                       | Runtime                         |
|-------------------|------------------------------------------------|---------------------------------|
| `google-docs/`    | Google Docs                                    | Google Apps Script (cloud-side) |
| `word-online/`    | Word for the Web, Word Desktop, Word for Mac   | Office Add-in (task pane)       |

Both versions share the same **behavior**, **comment syntax**, and **response
contract** — you write `@claude <instruction>` and Claude does the work.

---

## Comment Syntax (both platforms)

`@claude` is case-insensitive; the separator after it is flexible:

```
@claude: fix the grammar in this paragraph
@Claude, make this more concise
@claude please rewrite this in a formal tone
@claudeCan you summarize this section?
```

---

## What Claude Can Do (both platforms)

- **Edit text** — replace, rewrite, fix, expand, or restructure any passage
- **Answer questions** — give feedback, analysis, or explanations as a comment reply
- **Backfill placeholders** — find and replace `[VERIFY:...]` or `[RESEARCH NEEDED:...]`
  tags with real information using web search; add new document comments for
  anything it can't verify
- **Continue a thread** — reply to Claude's reply with another `@claude` message
  and it will respond in context, seeing the full conversation history

---

## Google Docs (`google-docs/`)

Runs as a Google Apps Script bound to your Google account. A time trigger
scans all recently-modified Google Docs in your Drive every 5 minutes, even
when no document is open.

### File reference

| File             | Purpose                                               |
|------------------|-------------------------------------------------------|
| `Code.gs`        | Comment scanning, placeholder extraction, edit logic  |
| `Claude.gs`      | Anthropic API call, web search, response parsing      |
| `Triggers.gs`    | Time trigger management, API key storage, status      |
| `appsscript.json`| OAuth scopes and advanced service declarations        |

### One-Time Setup

#### Step 1 — Create the Script Project

1. Go to https://script.google.com
2. Click **New project** and name it **Claude Co-author**

#### Step 2 — Add the Script Files

**appsscript.json**
- Click the gear icon in the left sidebar → check "Show appsscript.json manifest file"
- Replace its entire contents with `google-docs/appsscript.json` from this repo

**Code.gs**
- Click the existing `Code.gs` file and replace its contents with `google-docs/Code.gs`

**Claude.gs**
- Click **+** → Script → name it `Claude` → paste contents of `google-docs/Claude.gs`

**Triggers.gs**
- Click **+** → Script → name it `Triggers` → paste contents of `google-docs/Triggers.gs`

#### Step 3 — Enable Advanced Services

In the left sidebar click **Services** (+):
1. Add **Drive API** — find under D, select v3, click Add
2. Add **Google Docs API** — find under G, select v1, click Add

> If saving `appsscript.json` gives a "duplicate service" error, remove the
> `dependencies` block from `appsscript.json` — the Services UI already added it.

#### Step 4 — Set Your API Key

`promptApiKey` uses `DocumentApp.getUi()` which only works inside an open
Google Doc, not from the script editor. Use this temporary helper instead:

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

**4c.** Open `Triggers.gs` in the editor (the dropdown only shows functions
from the currently open file)

**4d.** Select `setApiKeyDirect` from the dropdown and click **Run**

**4e.** Confirm "API key saved." in the Execution log

**4f.** Delete the `setApiKeyDirect` function

#### Step 5 — Start the Background Trigger

**5a.** With `Triggers.gs` open, select `setupTrigger` from the dropdown and click **Run**

**5b.** When Google warns "this app isn't verified", click **Advanced** →
**Go to [project name] (unsafe)** — expected for all personal scripts

**5c.** Confirm "Auto-polling activated" in the Execution log

The trigger now runs permanently every 5 minutes, even when your browser is closed.

### Running Immediately (Without Waiting 5 Minutes)

Open `Code.gs` in the script editor, select `processAllRecentDocs` from the
function dropdown, and click **Run**. This processes all pending `@claude`
comments in any doc modified in the last 10 minutes right now.

### Verifying the Trigger is Active

1. Go to https://script.google.com → open Claude Co-author
2. Click the clock icon in the left sidebar (Triggers)
3. You should see `processAllRecentDocs` listed with a time-based trigger

---

## Word Online (`word-online/`)

Runs as a Microsoft Office Add-in that lives inside Word. When the Claude
Co-author task pane is open in a document, it scans that document's comments
and can auto-process every 5 minutes.

> **Behavioral difference from the Google Docs version:** Office Add-ins can
> only operate on the currently-open document and only while the task pane is
> active. For truly background, cross-document polling (the Apps Script
> behavior), you would need an Azure Functions + Microsoft Graph deployment —
> not included here.

### File reference

| File              | Purpose                                                        |
|-------------------|----------------------------------------------------------------|
| `manifest.xml`    | Office Add-in manifest (registers the task pane and ribbon)    |
| `taskpane.html`   | Task pane UI                                                   |
| `taskpane.css`    | Task pane styling                                              |
| `Code.js`         | Comment scanning, placeholder extraction, edit logic           |
| `Claude.js`       | Anthropic API call, web search, response parsing               |

### Hosting the Add-in Files

Office Add-ins require the HTML/JS/CSS to be served over HTTPS. The simplest
way is to enable GitHub Pages on this repo:

1. Push the repo to GitHub.
2. Go to **Settings → Pages**, select branch `main`, folder `/ (root)`, **Save**.
3. Wait ~1 minute. The files will be available at
   `https://<your-username>.github.io/<repo-name>/word-online/taskpane.html`.
4. Open `word-online/manifest.xml` and replace every occurrence of
   `https://idatarbel.github.io/Claude-Co-author/word-online/` with your own
   GitHub Pages URL (unless you're forking Dan's original at
   `idatarbel/Claude-Co-author` — that URL is already correct).

> **Alternative for local dev:** sideload the add-in while serving the files
> locally over HTTPS with any static server (e.g. `npx http-server -S -p 3000`
> with a self-signed cert, then point the manifest at `https://localhost:3000`).

### Installing the Add-in (Sideload)

**Word for the Web**

1. In Word for the Web, open any document.
2. Click **Home → Add-ins → Get Add-ins → Upload My Add-in**.
3. Browse to `word-online/manifest.xml` and upload it.
4. The **Claude** group appears on the Home tab — click **Claude Co-author**
   to open the task pane.

**Word Desktop (Windows or Mac)**

1. Create a shared folder trusted for add-ins — see Microsoft's guide at
   https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins
2. Copy `manifest.xml` into that shared folder.
3. In Word, go to **File → Options → Trust Center → Trust Center Settings →
   Trusted Add-in Catalogs**, add the folder, tick **Show in Menu**, OK.
4. Restart Word. Open **Home → Add-ins → Shared Folder**, select
   **Claude Co-author**, click Add.

### Configure Your API Key

1. Open the Claude Co-author task pane in any Word document.
2. Paste your Anthropic API key (from https://console.anthropic.com) into the
   **API Key** field.
3. Click **Save key**. The key is persisted in `OfficeRuntime.storage` (per
   user, per device) — it is not sent anywhere except the Anthropic API.

### Using It

- Write a comment in a Word document starting with `@claude`.
- Open the task pane and click **Process @claude comments now**, or tick
  **Auto-process every 5 minutes while this pane is open**.
- Claude's reply appears as a reply on the comment. Edits are applied
  directly to the document. New research comments are anchored where relevant.

### CORS Note

`Claude.js` sends the Anthropic API request from the task pane using the
`anthropic-dangerous-direct-browser-access: true` header, which Anthropic
requires for direct browser calls. If you run into CORS or policy issues in
your Office environment, host a small proxy endpoint and point `CLAUDE_API_URL`
at it.

---

## Troubleshooting

### Shared between both platforms

**"Comment not being picked up"**
- Confirm the comment starts with `@claude` (any capitalization).
- Make sure the comment isn't already resolved.
- Check that the last reply in the thread isn't already a `🤖 Claude:` reply.

**"Edits applied but wrong text replaced"**
Claude uses the exact strings from the document as search keys. If the
document was edited between Claude reading it and applying the edit, the
string may no longer match. Just re-run.

**"0 edits applied / string not found"**
The text Claude targeted may span a paragraph break or formatting boundary.
Reply to the comment with a more specific instruction targeting a shorter,
single-paragraph passage.

### Google Docs only

**"Cannot call DocumentApp.getUi() from this context"**
You ran `promptApiKey` or a menu function directly from the script editor.
Use the `setApiKeyDirect` workaround in Step 4. `setupTrigger` uses
`Logger.log` and runs fine from the editor.

**"Drive is not defined" or "Docs API error 403"**
One or both advanced services are not enabled. Go to Services in the left
sidebar and add Drive API v3 and Google Docs API v1.

**Function not in dropdown**
The dropdown only shows functions from the currently open file. Click the
correct `.gs` file in the left sidebar first.

**"Google hasn't verified this app"**
Expected for personal scripts. Click Advanced → Go to [project name] (unsafe).

**Duplicate Drive service error**
Remove the `dependencies` block from `appsscript.json`.

**Comment not picked up in Google Docs**
The file must be a native Google Doc — not a .docx. Use **File → Save as
Google Docs**.

### Word Online only

**Task pane loads but shows "Office.js not available"**
Make sure you opened the add-in from inside a Word document — the task pane
cannot be loaded standalone in a browser.

**"Comments not supported in this version of Word"**
The comments API requires WordApi 1.4+. Update Word Desktop to a recent
Microsoft 365 build or use Word for the Web.

**Manifest upload rejected**
Validate the manifest with
[Office Add-in Manifest Validator](https://learn.microsoft.com/office/dev/add-ins/testing/troubleshoot-manifest)
and make sure all `SourceLocation` / `IconUrl` / `Taskpane.Url` values point
to a reachable HTTPS URL.

---

## Security Notes

- API keys are stored only on your side (Google `PropertiesService` for the
  Apps Script version, `OfficeRuntime.storage` for the Word Add-in).
- Neither version sends your documents anywhere except the Anthropic API using
  your own key.
- The Apps Script version only accesses documents in your own Google Drive.
- The Office Add-in version only accesses the document it is currently opened
  in (per Microsoft's add-in sandbox).
