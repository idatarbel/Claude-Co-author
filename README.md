# Claude Co-author — Google Apps Script

An always-on Claude integration for Google Docs. Add `@claude <instruction>` as a comment anywhere in any Google Doc, and Claude will automatically reply — and optionally edit the document — within 5 minutes.

---

## How It Works

1. You write a comment in any Google Doc starting with `@claude`
2. A background trigger (every 5 min) scans all your recently active docs
3. Claude receives the document content + the comment context
4. Claude replies to the comment and optionally edits the document directly
5. You see the reply thread on the comment, just like a human collaborator responded

---

## Comment Syntax

All of the following formats work — `@claude` is case-insensitive and the separator after it is optional:

```
@claude: fix the grammar in this paragraph
@Claude, make this more concise
@claude please rewrite this in a formal tone
@claudeCan you summarize this section?
```

---

## One-Time Setup

### Step 1 — Create the Script Project

1. Go to https://script.google.com
2. Click **New project**
3. Name it **Claude Co-author**

### Step 2 — Add the Script Files

#### appsscript.json
- Click the gear icon in the left sidebar → check **"Show appsscript.json manifest file"**
- Click `appsscript.json` in the file list
- Replace the entire contents with the `appsscript.json` from this package

#### Code.gs
- Click the existing `Code.gs` file
- Replace its contents with `Code.gs` from this package

#### Claude.gs
- Click **+** → Script → name it `Claude`
- Paste the contents of `Claude.gs`

#### Triggers.gs
- Click **+** → Script → name it `Triggers`
- Paste the contents of `Triggers.gs`

### Step 3 — Enable the Drive Advanced Service

1. In the left sidebar, click **Services** (the + icon)
2. Find **Drive API** (listed under D, not G) → select **v3** → click **Add**

### Step 4 — Set Your API Key

`promptApiKey` uses `DocumentApp.getUi()` which cannot run from the script editor
directly — it only works inside an actual open Google Doc. For initial setup,
use this temporary helper function instead:

**4a.** Add this function temporarily to the bottom of `Triggers.gs`:

```javascript
function setApiKeyDirect() {
  PropertiesService.getUserProperties().setProperty(
    'claudeApiKey',
    'sk-ant-YOUR-KEY-HERE'  // paste your full Anthropic API key here
  );
  Logger.log('API key saved.');
}
```

**4b.** Replace `sk-ant-YOUR-KEY-HERE` with your actual key from https://console.anthropic.com

**4c.** Open `Triggers.gs` in the editor — the function dropdown only shows functions
from the currently open file

**4d.** Select `setApiKeyDirect` from the function dropdown and click **Run**

**4e.** Confirm you see `API key saved.` in the Execution log

**4f.** Delete the `setApiKeyDirect` function — you don't need it anymore

### Step 5 — Start the Background Trigger

**5a.** With `Triggers.gs` still open, select `setupTrigger` from the dropdown and click **Run**

**5b.** Approve the permissions prompt. When Google warns "this app isn't verified",
click **Advanced** → **Go to [project name] (unsafe)** — this is expected for any
personal script and is safe since you are the developer.

**5c.** Confirm you see `Auto-polling activated` in the Execution log

The trigger now runs permanently every 5 minutes, even when your browser is closed.

---

## Usage

In any Google Doc, add a comment with `@claude` followed by your instruction:

```
@claude rewrite this paragraph to be more concise
@claude does this argument make logical sense?
@claude add a transition sentence before this paragraph
@claude fix grammar and punctuation throughout this section
@claude summarize the key points here as bullet points
```

Claude will reply only for questions and feedback, or edit the document directly
plus post a reply explaining what changed — depending on your instruction.

---

## Verifying the Trigger is Running

1. Go to https://script.google.com
2. Open the Claude Co-author project
3. Click the clock icon in the left sidebar (Triggers)
4. You should see `processAllRecentDocs` listed with a time-based trigger

---

## File Reference

| File         | Purpose                                              |
|--------------|------------------------------------------------------|
| Code.gs      | Comment scanning, document editing, core logic       |
| Claude.gs    | Anthropic API call and response parsing              |
| Triggers.gs  | Time trigger management, API key storage, status     |
| appsscript.json | OAuth scopes and Drive API v3 advanced service    |

---

## Troubleshooting

**"Cannot call DocumentApp.getUi() from this context"**
You ran a function directly from the script editor that requires a live Doc context.
Use the temporary `setApiKeyDirect` workaround in Step 4 above. `setupTrigger` in
this package already uses Logger.log instead of getUi() and runs fine from the editor.

**Function not showing in the dropdown**
The dropdown only shows functions from the currently open file. Click the specific
.gs file in the left sidebar first, then check the dropdown.

**"Google hasn't verified this app" warning**
Expected for all personal scripts. Click Advanced → Go to [project name] (unsafe).

**Duplicate Drive service error when saving appsscript.json**
Remove the `dependencies` block from appsscript.json. The Services UI already added
it automatically and the manifest doesn't need it too.

**Comment not being picked up**
- Confirm the comment starts with @claude (any capitalization)
- The file must be a native Google Doc, not a .docx opened in Drive.
  Go to File → Save as Google Docs to convert it first.
- Run `processAllRecentDocs` directly from the editor to trigger immediately.

**Claude replied but didn't apply the edit**
Claude notes in its reply when it couldn't locate the target text. Rephrase your
instruction quoting more specific text from the document.

---

## Security Notes

- Your API key is stored in PropertiesService.getUserProperties() — tied to your
  Google account, not visible in the script code, not shared with anyone.
- The script only accesses documents in your own Google Drive.
- No data is sent anywhere except the Anthropic API using your own key.
