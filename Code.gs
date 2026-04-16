// ============================================================
// Claude Co-author — Code.gs
// Scans Google Docs for @claude comments and processes them.
// ============================================================

const CLAUDE_API_URL   = 'https://api.anthropic.com/v1/messages';
const CLAUDE_MODEL     = 'claude-sonnet-4-20250514';
const COMMENT_TRIGGER  = '@claude';
const REPLY_MARKER     = '🤖 Claude:';
const LOOKBACK_MINUTES = 10;

// ─── Add-on Lifecycle ─────────────────────────────────────

function onOpen(e) {
  try {
    DocumentApp.getUi()
      .createAddonMenu()
      .addItem('⚡ Process @claude comments now', 'processCurrentDoc')
      .addSeparator()
      .addItem('▶ Start auto-polling (every 5 min)', 'setupTrigger')
      .addItem('⏹ Stop auto-polling', 'removeTrigger')
      .addSeparator()
      .addItem('🔑 Set API Key', 'promptApiKey')
      .addItem('📊 View polling status', 'showStatus')
      .addToUi();
  } catch(e) {
    // Not running inside a Google Doc — skip UI setup
  }
}

function onInstall(e) {
  onOpen(e);
}

// ─── Manual Trigger (Menu Item) ────────────────────────────

function processCurrentDoc() {
  const doc = DocumentApp.getActiveDocument();
  if (!doc) {
    DocumentApp.getUi().alert('❌ No active document found.');
    return;
  }
  const apiKey = getApiKey();
  if (!apiKey) {
    DocumentApp.getUi().alert('❌ No API key set. Use Extensions > Claude Co-author > Set API Key first.');
    return;
  }
  const count = processDocById(doc.getId(), apiKey);
  DocumentApp.getUi().alert(
    count > 0
      ? `✅ Done! Processed ${count} @claude comment(s).`
      : `ℹ️ No new @claude comments found in this document.`
  );
}

// ─── Automatic Time Trigger (runs every 5 min) ─────────────

function processAllRecentDocs() {
  const apiKey = getApiKey();
  if (!apiKey) return;

  const d = new Date(Date.now() - LOOKBACK_MINUTES * 60 * 1000);
  const since = Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const query = `mimeType="application/vnd.google-apps.document" and modifiedDate > "${since}"`;

  let files;
  try {
    files = DriveApp.searchFiles(query);
  } catch (e) {
    console.error('Drive search failed: ' + e.message);
    return;
  }

  while (files.hasNext()) {
    const file = files.next();
    try {
      processDocById(file.getId(), apiKey);
    } catch (e) {
      console.error(`Skipped doc ${file.getId()}: ${e.message}`);
    }
  }
}

// ─── Core Processing Logic ─────────────────────────────────

function processDocById(docId, apiKey) {
  let commentsResponse;
  try {
    commentsResponse = Drive.Comments.list(docId, {
      fields: 'comments(id,content,resolved,replies(content),quotedFileContent)',
      pageSize: 100
    });
  } catch (e) {
    return 0;
  }

  const comments = commentsResponse.comments;
  if (!comments || comments.length === 0) return 0;

  let docText = '';
  let docName = '';
  try {
    const doc = DocumentApp.openById(docId);
    docText = doc.getBody().getText();
    docName = doc.getName();
  } catch (e) {
    console.error(`Could not open doc ${docId}: ${e.message}`);
    return 0;
  }

  let processedCount = 0;

  for (const comment of comments) {
    try {
      if (shouldSkipComment(comment)) continue;

      const instruction = extractInstruction(comment.content);
      const quotedText  = (comment.quotedFileContent || {}).value || '';

      const result = callClaude(apiKey, instruction, docText, quotedText, docName);
      if (!result) continue;

      let replyText = REPLY_MARKER + ' ';
      if (result.action === 'edit' && result.edit && result.edit.original_text) {
        replyText += applyEdit(docId, result.edit, result.reply);
      } else {
        replyText += result.reply;
      }

      Drive.Replies.create(
        { content: replyText },
        docId,
        comment.id,
        { fields: 'id' }
      );

      processedCount++;
    } catch (e) {
      console.error(`Error on comment ${comment.id}: ${e.message}`);
    }
  }

  return processedCount;
}

// ─── Helpers ───────────────────────────────────────────────

function shouldSkipComment(comment) {
  if (comment.resolved) return true;

  const content = (comment.content || '').trim();
  if (!content.toLowerCase().startsWith(COMMENT_TRIGGER)) return true;

  const replies = comment.replies || [];
  if (replies.some(r => r.content && r.content.startsWith(REPLY_MARKER))) return true;

  return false;
}

function extractInstruction(commentContent) {
  // Remove @claude prefix then strip any leading non-alphanumeric characters
  // (colon, comma, space, dash, etc.) so all these work:
  // "@claude: fix this"  "@claude, fix this"  "@claudeFix this"  "@claude fix this"
  const stripped = commentContent.substring(COMMENT_TRIGGER.length);
  return stripped.replace(/^[^a-zA-Z0-9]+/, '').trim();
}

function applyEdit(docId, edit, claudeReply) {
  try {
    const doc  = DocumentApp.openById(docId);
    const body = doc.getBody();
    const orig = edit.original_text;
    const repl = edit.replacement_text;

    const escapedOrig = escapeRegex(orig);
    const escapedRepl = repl.replace(/\$/g, '$$$$');

    const found = body.findText(escapedOrig);
    if (!found) {
      return claudeReply + '\n\n⚠️ Could not locate the target text to edit — please apply manually.';
    }

    body.replaceText(escapedOrig, escapedRepl);
    return claudeReply + '\n\n✅ Edit applied: replaced the indicated text.';
  } catch (e) {
    return claudeReply + `\n\n⚠️ Edit failed (${e.message}) — please apply manually.`;
  }
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


function debugComments() {
  const apiKey = getApiKey();
  Logger.log('API key present: ' + !!apiKey);

  const d = new Date(Date.now() - 60 * 60 * 1000); // last 60 minutes
  const since = Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const query = `mimeType="application/vnd.google-apps.document" and modifiedDate > "${since}"`;
  
  const files = DriveApp.searchFiles(query);
  let fileCount = 0;
  while (files.hasNext()) {
    const file = files.next();
    fileCount++;
    Logger.log('Found doc: ' + file.getName() + ' (' + file.getId() + ')');
    
    try {
      const resp = Drive.Comments.list(file.getId(), {
        fields: 'comments(id,content,resolved,replies(content),quotedFileContent)',
        pageSize: 100
      });
      const comments = resp.comments || [];
      Logger.log('  Comments found: ' + comments.length);
      comments.forEach(c => {
        Logger.log('  Comment: ' + c.content + ' | resolved: ' + c.resolved);
      });
    } catch(e) {
      Logger.log('  ERROR fetching comments: ' + e.message);
    }
  }
  Logger.log('Total docs found: ' + fileCount);
}