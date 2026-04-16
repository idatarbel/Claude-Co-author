// ============================================================
// Claude Co-author — Code.js (Word Online)
// Scans a Word document for @claude comments and processes them.
// Responds to @claude in the original comment OR any reply.
// Mirrors Google Docs/Code.gs.
// ============================================================

const COMMENT_TRIGGER      = '@claude';
const REPLY_MARKER         = '\uD83E\uDD16 Claude:'; // 🤖 Claude:
const POLL_INTERVAL_MS     = 5 * 60 * 1000;
const STORAGE_KEY_API_KEY  = 'claudeApiKey';
const STORAGE_KEY_AUTOPOLL = 'claudeAutoPoll';

let pollIntervalId = null;
let isProcessing   = false;

// ─── Office Add-in Lifecycle ──────────────────────────────

Office.onReady(async info => {
  if (info.host !== Office.HostType.Word) return;

  document.getElementById('save-key-btn').onclick      = saveApiKey;
  document.getElementById('change-key-btn').onclick    = () => showKeyEditView(true);
  document.getElementById('cancel-key-btn').onclick    = () => showKeyEditView(false);
  document.getElementById('clear-key-btn').onclick     = clearApiKey;
  document.getElementById('process-now-btn').onclick   = onProcessNowClick;
  document.getElementById('auto-poll-toggle').onchange = onAutoPollToggle;

  // Restore saved API key and auto-poll preference
  const savedKey = await getStoredValue(STORAGE_KEY_API_KEY);
  renderKeyView(savedKey);

  const autoPoll = await getStoredValue(STORAGE_KEY_AUTOPOLL);
  if (autoPoll === 'true') {
    document.getElementById('auto-poll-toggle').checked = true;
    startAutoPolling();
  }

  log('info', 'Ready. Add @claude comments in the document and click Process now.');
});

// ─── UI Handlers ──────────────────────────────────────────

async function saveApiKey() {
  const input = document.getElementById('api-key-input');
  const key   = input.value.trim();

  if (!key || !key.startsWith('sk-ant')) {
    setKeyStatus('err', 'That does not look like a valid Anthropic API key.');
    return;
  }

  await setStoredValue(STORAGE_KEY_API_KEY, key);
  input.value = '';
  renderKeyView(key);
  log('ok', 'API key saved.');
}

async function clearApiKey() {
  await setStoredValue(STORAGE_KEY_API_KEY, '');
  renderKeyView(null);
  log('info', 'API key removed.');
}

async function onProcessNowClick() {
  if (isProcessing) {
    log('info', 'Already processing — please wait.');
    return;
  }
  await processCurrentDoc(/*manual=*/true);
}

async function onAutoPollToggle(e) {
  const on = e.target.checked;
  await setStoredValue(STORAGE_KEY_AUTOPOLL, on ? 'true' : 'false');
  if (on) startAutoPolling();
  else    stopAutoPolling();
}

function startAutoPolling() {
  if (pollIntervalId) return;
  pollIntervalId = setInterval(() => {
    processCurrentDoc(/*manual=*/false).catch(err => log('err', 'Auto-poll error: ' + err.message));
  }, POLL_INTERVAL_MS);
  log('ok', 'Auto-polling started (every 5 min while this pane is open).');
}

function stopAutoPolling() {
  if (!pollIntervalId) return;
  clearInterval(pollIntervalId);
  pollIntervalId = null;
  log('info', 'Auto-polling stopped.');
}

// ─── Core Processing Logic ────────────────────────────────

async function processCurrentDoc(manual) {
  const apiKey = await getStoredValue(STORAGE_KEY_API_KEY);
  if (!apiKey) {
    log('err', 'No API key set. Enter one above and click Save key.');
    return;
  }

  isProcessing = true;
  const processBtn = document.getElementById('process-now-btn');
  processBtn.disabled = true;
  processBtn.textContent = 'Processing...';

  try {
    const { docText, docName, commentData } = await readDocState();

    if (!commentData.length) {
      if (manual) log('info', 'No comments found in document.');
      return;
    }

    const placeholders = extractPlaceholders(docText);
    const toProcess    = commentData.filter(shouldProcessComment);

    if (toProcess.length === 0) {
      if (manual) log('info', 'No unprocessed @claude comments found.');
      return;
    }

    log('info', `Processing ${toProcess.length} @claude comment(s) in "${docName}"...`);

    let processed = 0;
    for (const c of toProcess) {
      try {
        const ok = await processOneComment(c, apiKey, docText, docName, placeholders);
        if (ok) processed++;
      } catch (e) {
        log('err', `Comment error: ${e.message}`);
      }
    }

    log('ok', `Done. Processed ${processed} of ${toProcess.length} comment(s).`);
  } catch (e) {
    log('err', 'Fatal: ' + e.message);
  } finally {
    isProcessing = false;
    processBtn.disabled = false;
    processBtn.textContent = 'Process @claude comments now';
  }
}

// Read document text, name, and comment tree in one Word.run pass.
async function readDocState() {
  return Word.run(async context => {
    const body     = context.document.body;
    const comments = body.getComments();
    const props    = context.document.properties;

    body.load('text');
    props.load('title');
    comments.load('items/id,items/content,items/resolved,items/authorName,items/replies/items/id,items/replies/items/content,items/replies/items/authorName');
    await context.sync();

    const commentData = [];
    for (const c of comments.items) {
      // Pull anchored range text separately (getRange is a method, not a loaded prop)
      const range = c.getRange();
      range.load('text');
      await context.sync();

      commentData.push({
        id:        c.id,
        content:   c.content || '',
        resolved:  c.resolved || false,
        replies:   (c.replies && c.replies.items) ? c.replies.items.map(r => ({
          id:      r.id,
          content: r.content || ''
        })) : [],
        quotedText: range.text || ''
      });
    }

    return {
      docText:     body.text || '',
      docName:     props.title || 'Untitled',
      commentData: commentData
    };
  });
}

// Decide whether this comment has a @claude message awaiting a response.
function shouldProcessComment(c) {
  if (c.resolved) return false;

  const thread        = [c.content, ...c.replies.map(r => r.content)];
  const lastClaudeIdx = findLastClaudeMessage(thread);
  if (lastClaudeIdx === -1) return false;

  const messagesAfter = thread.slice(lastClaudeIdx + 1);
  if (messagesAfter.some(m => m.startsWith(REPLY_MARKER))) return false;

  c._thread        = thread;
  c._lastClaudeIdx = lastClaudeIdx;
  return true;
}

async function processOneComment(c, apiKey, docText, docName, placeholders) {
  const instruction = extractInstruction(c._thread[c._lastClaudeIdx]);
  const history     = c._thread.slice(0, c._lastClaudeIdx);

  const result = await callClaude(
    apiKey, instruction, docText, c.quotedText, docName, history, placeholders
  );
  if (!result) {
    log('err', 'No result from Claude for comment.');
    return false;
  }

  let editSummary = '';
  if (result.edits && result.edits.length > 0) {
    editSummary = '\n\n' + await applyEdits(result.edits);
  }

  let insertSummary = '';
  if (result.inserts && result.inserts.length > 0) {
    insertSummary = '\n\n' + await applyInserts(result.inserts);
  }

  let commentSummary = '';
  if (result.comments_to_add && result.comments_to_add.length > 0) {
    const added = await addDocComments(result.comments_to_add);
    commentSummary = `\n\n\uD83D\uDCAC ${added} research comment(s) added to document.`;
  }

  const replyText = REPLY_MARKER + ' ' + sanitizeReplacement(result.reply) + editSummary + insertSummary + commentSummary;
  await replyToComment(c.id, replyText);
  return true;
}

// ─── Word Document Mutations ──────────────────────────────

async function applyEdits(edits) {
  const valid = edits.filter(e => e.original_text && e.replacement_text);
  if (valid.length === 0) return 'No valid edits to apply.';

  let applied      = 0;
  let missed       = 0;
  let skipped      = 0;
  const missedTexts = [];

  await Word.run(async context => {
    for (const edit of valid) {
      // search() can't match across paragraph breaks; skip multi-line edits.
      if (/\r|\n/.test(edit.original_text)) {
        skipped++;
        missedTexts.push('(multi-line edit skipped) ' + edit.original_text.split(/\r?\n/)[0]);
        continue;
      }
      const results = context.document.body.search(edit.original_text, { matchCase: true });
      results.load('items');
      await context.sync();

      if (results.items.length === 0) {
        missed++;
        missedTexts.push(edit.original_text);
        continue;
      }
      // Replace only the first match to avoid clobbering identical text elsewhere.
      results.items[0].insertText(sanitizeReplacement(edit.replacement_text), 'Replace');
      applied++;
    }
    await context.sync();
  });

  if (missedTexts.length > 0) {
    log('err', 'Not found in document:\n  ' + missedTexts.map(t => `"${t.substring(0, 80)}"`).join('\n  '));
  }

  let summary = `\u2705 ${applied} edit(s) applied.`;
  if (missed  > 0) summary += ` \u26A0\uFE0F ${missed} string(s) not found.`;
  if (skipped > 0) summary += ` \u26A0\uFE0F ${skipped} multi-line edit(s) skipped — use "inserts" instead.`;
  return summary;
}

async function applyInserts(inserts) {
  const valid = inserts.filter(i =>
    i.after_text &&
    Array.isArray(i.new_paragraphs) &&
    i.new_paragraphs.length > 0
  );
  if (valid.length === 0) return 'No valid inserts to apply.';

  let inserted    = 0;
  let missed      = 0;
  const missedTexts = [];

  await Word.run(async context => {
    for (const ins of valid) {
      if (/\r|\n/.test(ins.after_text)) {
        missed++;
        missedTexts.push('(multi-line anchor) ' + ins.after_text.split(/\r?\n/)[0]);
        continue;
      }
      const results = context.document.body.search(ins.after_text, { matchCase: true });
      results.load('items');
      await context.sync();

      if (results.items.length === 0) {
        missed++;
        missedTexts.push(ins.after_text);
        continue;
      }

      // Walk up to the containing paragraph.
      const paragraphs = results.items[0].paragraphs;
      paragraphs.load('items');
      await context.sync();

      let anchorPara = paragraphs.items[paragraphs.items.length - 1];

      for (const newText of ins.new_paragraphs) {
        const clean = sanitizeReplacement(newText);
        if (!clean) continue;
        anchorPara = anchorPara.insertParagraph(clean, 'After');
        await context.sync();
        inserted++;
      }
    }
  });

  if (missedTexts.length > 0) {
    log('err', 'Anchor not found for inserts:\n  ' + missedTexts.map(t => `"${t.substring(0, 80)}"`).join('\n  '));
  }

  let summary = `\u2795 ${inserted} paragraph(s) inserted.`;
  if (missed > 0) summary += ` \u26A0\uFE0F ${missed} anchor(s) not found.`;
  return summary;
}

async function addDocComments(commentsToAdd) {
  let added = 0;

  await Word.run(async context => {
    for (const c of commentsToAdd) {
      if (!c.comment) continue;
      let ok = false;

      if (c.quoted_text) {
        const results = context.document.body.search(c.quoted_text, { matchCase: true });
        results.load('items');
        await context.sync();

        if (results.items.length > 0) {
          results.items[0].insertComment(c.comment);
          ok = true;
        }
      }

      if (!ok) {
        // Fallback: attach to the end of the document so the comment isn't lost.
        const end = context.document.body.getRange('End');
        end.insertComment(c.comment);
      }

      added++;
      await context.sync();
    }
  });

  return added;
}

async function replyToComment(commentId, replyText) {
  await Word.run(async context => {
    const comments = context.document.body.getComments();
    comments.load('items/id');
    await context.sync();

    const target = comments.items.find(c => c.id === commentId);
    if (!target) throw new Error(`Comment ${commentId} not found for reply.`);

    target.reply(replyText);
    await context.sync();
  });
}

// ─── Placeholder Extraction (mirrors Code.gs) ─────────────

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

// ─── Thread Helpers ───────────────────────────────────────

function findLastClaudeMessage(thread) {
  for (let i = thread.length - 1; i >= 0; i--) {
    if ((thread[i] || '').trim().toLowerCase().startsWith(COMMENT_TRIGGER)) return i;
  }
  return -1;
}

function extractInstruction(message) {
  return (message || '').substring(COMMENT_TRIGGER.length).replace(/^[^a-zA-Z0-9]+/, '').trim();
}

// ─── Replacement Text Sanitization ────────────────────────

function sanitizeReplacement(text) {
  return (text || '')
    .replace(/<cite[^>]*>/gi, '')
    .replace(/<\/cite>/gi, '')
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/\*(.*?)\*/g, '$1')
    .replace(/__(.*?)__/g, '$1')
    .replace(/_(.*?)_/g, '$1')
    .replace(/<[^>]+>/g, '')
    .trim();
}

// ─── Storage (OfficeRuntime.storage with localStorage fallback) ─

async function getStoredValue(key) {
  try {
    if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
      return await OfficeRuntime.storage.getItem(key);
    }
  } catch (e) { /* fall through */ }
  try {
    return localStorage.getItem(key);
  } catch (e) {
    return null;
  }
}

async function setStoredValue(key, value) {
  try {
    if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) {
      await OfficeRuntime.storage.setItem(key, value);
      return;
    }
  } catch (e) { /* fall through */ }
  try {
    localStorage.setItem(key, value);
  } catch (e) {
    console.error('Could not persist value: ' + e.message);
  }
}

// ─── UI Helpers ───────────────────────────────────────────

// Show the "saved key" view if a key exists, otherwise the edit view.
function renderKeyView(key) {
  const savedView = document.getElementById('key-saved-view');
  const editView  = document.getElementById('key-edit-view');

  if (key && key.startsWith('sk-ant')) {
    document.getElementById('key-masked').textContent = maskKey(key);
    savedView.classList.remove('hidden');
    editView.classList.add('hidden');
  } else {
    savedView.classList.add('hidden');
    editView.classList.remove('hidden');
    setKeyStatus('', '');
  }
}

// Explicitly toggle into the edit view (used by Change / Cancel).
async function showKeyEditView(editing) {
  if (editing) {
    document.getElementById('key-saved-view').classList.add('hidden');
    document.getElementById('key-edit-view').classList.remove('hidden');
    setKeyStatus('', '');
    document.getElementById('api-key-input').focus();
  } else {
    const savedKey = await getStoredValue(STORAGE_KEY_API_KEY);
    document.getElementById('api-key-input').value = '';
    renderKeyView(savedKey);
  }
}

function maskKey(key) {
  if (!key || key.length < 10) return '\u2022\u2022\u2022';
  return key.slice(0, 7) + '\u2026' + key.slice(-4);
}

function setKeyStatus(level, message) {
  const el = document.getElementById('key-status');
  if (!el) return;
  el.textContent = message || '';
  el.className   = 'status' + (level ? ' ' + level : '');
}

function log(level, message) {
  const logEl = document.getElementById('log');
  if (!logEl) return;

  const entry = document.createElement('div');
  entry.className = 'entry';

  const ts = document.createElement('span');
  ts.className = 'ts';
  ts.textContent = new Date().toLocaleTimeString();

  const msg = document.createElement('span');
  msg.className = level; // ok | err | info
  msg.textContent = message;

  entry.appendChild(ts);
  entry.appendChild(msg);
  logEl.appendChild(entry);
  logEl.scrollTop = logEl.scrollHeight;
}
