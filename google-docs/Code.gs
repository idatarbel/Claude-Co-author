// ============================================================
// Claude Co-author — Code.gs
// Scans Google Docs for @claude comments and processes them.
// Responds to @claude in the original comment OR any reply.
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

// ─── Manual Trigger ────────────────────────────────────────

function processCurrentDoc() {
  const doc = DocumentApp.getActiveDocument();
  if (!doc) {
    DocumentApp.getUi().alert('❌ No active document found.');
    return;
  }
  const apiKey = getApiKey();
  if (!apiKey) {
    DocumentApp.getUi().alert('❌ No API key set.');
    return;
  }
  const count = processDocById(doc.getId(), apiKey);
  DocumentApp.getUi().alert(
    count > 0
      ? `✅ Done! Processed ${count} @claude comment(s).`
      : `ℹ️ No new @claude comments found.`
  );
}

// ─── Automatic Time Trigger ────────────────────────────────

function processAllRecentDocs() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    console.log('Another execution in progress — skipping this trigger fire.');
    return;
  }
  try {
    runScan();
  } finally {
    lock.releaseLock();
  }
}

function runScan() {
  const apiKey = getApiKey();
  if (!apiKey) return;

  const d = new Date(Date.now() - LOOKBACK_MINUTES * 60 * 1000);
  const since = Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'HH:mm:ss'Z'");
  const query = "mimeType='application/vnd.google-apps.document' and modifiedTime > '" + since + "'";

  const token    = ScriptApp.getOAuthToken();
  const url      = 'https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(query) + '&fields=files(id,name)&pageSize=50';
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    console.error('Drive REST API error: ' + response.getContentText());
    return;
  }

  const files = JSON.parse(response.getContentText()).files || [];
  files.forEach(file => {
    try {
      processDocById(file.id, apiKey);
    } catch(e) {
      console.error('Skipped ' + file.id + ': ' + e.message);
    }
  });
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
    console.error('Comments fetch failed: ' + e.message);
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

  const placeholders = extractPlaceholders(docText);

  let processedCount = 0;

  for (const comment of comments) {
    try {
      if (comment.resolved) continue;

      const thread        = buildThread(comment);
      const lastClaudeIdx = findLastClaudeMessage(thread);
      if (lastClaudeIdx === -1) continue;

      const messagesAfter = thread.slice(lastClaudeIdx + 1);
      if (messagesAfter.some(m => m.startsWith(REPLY_MARKER))) continue;

      const instruction = extractInstruction(thread[lastClaudeIdx]);
      const history     = thread.slice(0, lastClaudeIdx);
      const quotedText  = (comment.quotedFileContent || {}).value || '';

      const result = callClaude(
        apiKey, instruction, docText, quotedText, docName, history, placeholders
      );
      if (!result) continue;

      let editSummary = '';
      if (result.edits && result.edits.length > 0) {
        editSummary = '\n\n' + applyEdits(docId, result.edits);
      }

      let commentSummary = '';
      if (result.comments_to_add && result.comments_to_add.length > 0) {
        const added = addDocComments(docId, result.comments_to_add);
        commentSummary = `\n\n💬 ${added} research comment(s) added to document.`;
      }

      const replyText = REPLY_MARKER + ' ' + sanitizeReplacement(result.reply) + editSummary + commentSummary;

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

// ─── Add New Comments to Document ─────────────────────────

function addDocComments(docId, commentsToAdd) {
  let added = 0;
  for (const c of commentsToAdd) {
    if (!c.comment) continue;
    try {
      const commentBody = { content: c.comment };
      if (c.quoted_text) {
        commentBody.quotedFileContent = {
          mimeType: 'text/plain',
          value:    c.quoted_text
        };
      }
      Drive.Comments.create(commentBody, docId, { fields: 'id' });
      added++;
    } catch(e) {
      console.error('Failed to add comment: ' + e.message);
    }
  }
  return added;
}

// ─── Placeholder Extraction ────────────────────────────────

function extractPlaceholders(text) {
  const placeholders = [];
  const prefixes = ['[VERIFY:', '[RESEARCH NEEDED:'];

  for (const prefix of prefixes) {
    let i = 0;
    while (i < text.length) {
      const start = text.indexOf(prefix, i);
      if (start === -1) break;

      let depth = 1;
      let j = start + 1;
      while (j < text.length && depth > 0) {
        if (text[j] === '[') depth++;
        else if (text[j] === ']') depth--;
        j++;
      }

      if (depth === 0) {
        const placeholder = text.substring(start, j);
        if (!placeholders.includes(placeholder)) {
          placeholders.push(placeholder);
        }
      }

      i = start + 1;
    }
  }

  return placeholders;
}

// ─── Thread Helpers ────────────────────────────────────────

function buildThread(comment) {
  const thread = [comment.content || ''];
  (comment.replies || []).forEach(r => thread.push(r.content || ''));
  return thread;
}

function findLastClaudeMessage(thread) {
  for (let i = thread.length - 1; i >= 0; i--) {
    if (thread[i].trim().toLowerCase().startsWith(COMMENT_TRIGGER)) return i;
  }
  return -1;
}

function extractInstruction(message) {
  return message.substring(COMMENT_TRIGGER.length).replace(/^[^a-zA-Z0-9]+/, '').trim();
}

// ─── Edit Application via Docs REST API ───────────────────

function sanitizeReplacement(text) {
  return text
    .replace(/<cite[^>]*>/gi, '')
    .replace(/<\/cite>/gi, '')
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/\*(.*?)\*/g, '$1')
    .replace(/__(.*?)__/g, '$1')
    .replace(/_(.*?)_/g, '$1')
    .replace(/<[^>]+>/g, '')
    .trim();
}

function applyEdits(docId, edits) {
  if (!edits || edits.length === 0) return '';

  const requests = edits
    .filter(e => e.original_text && e.replacement_text)
    .map(e => ({
      replaceAllText: {
        containsText: { text: e.original_text, matchCase: true },
        replaceText:  sanitizeReplacement(e.replacement_text)
      }
    }));

  if (requests.length === 0) return '⚠️ No valid edits to apply.';

  let result;
  try {
    const response = UrlFetchApp.fetch(
      `https://docs.googleapis.com/v1/documents/${docId}:batchUpdate`,
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
          'Content-Type':  'application/json'
        },
        payload: JSON.stringify({ requests }),
        muteHttpExceptions: true
      }
    );

    const status = response.getResponseCode();
    if (status !== 200) {
      return `⚠️ Docs API error ${status}: ${response.getContentText()}`;
    }

    result = JSON.parse(response.getContentText());
  } catch(e) {
    return `⚠️ Request failed: ${e.message}`;
  }

  const replies = result.replies || [];
  let applied = 0;
  let missed  = 0;
  replies.forEach(r => {
    const count = (r.replaceAllText && r.replaceAllText.occurrencesChanged) || 0;
    if (count > 0) applied++;
    else missed++;
  });

  let summary = `✅ ${applied} edit(s) applied.`;
  if (missed > 0) summary += ` ⚠️ ${missed} string(s) not found in document.`;
  return summary;
}

// ─── Debug Helper (delete after use) ──────────────────────

function debugComments() {
  const apiKey = getApiKey();
  Logger.log('API key present: ' + !!apiKey);

  const d = new Date(Date.now() - 60 * 60 * 1000);
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
        Logger.log('  Comment: ' + c.content);
        (c.replies || []).forEach((r, i) => Logger.log('    Reply ' + i + ': ' + r.content));
      });
    } catch(e) {
      Logger.log('  ERROR: ' + e.message);
    }
  }
  Logger.log('Total docs found: ' + fileCount);
}