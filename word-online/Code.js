// ============================================================
// Claude Co-author — Code.js (Word Online)
// Scans a Word document for @claude comments and processes them.
// Responds to @claude in the original comment OR any reply.
// Mirrors Google Docs/Code.gs.
// ============================================================

const BUILD_VERSION        = 'v17';
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

  // Surface the build we're running so there's never any ambiguity.
  const buildEl = document.getElementById('build-tag');
  if (buildEl) buildEl.textContent = BUILD_VERSION;

  // Self-update: if build-version.txt on the server says we're stale,
  // reload the page once. Guarded by sessionStorage to avoid loops.
  autoUpdateCheck();

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

// ─── Self-Update ──────────────────────────────────────────

// Fetch build-version.txt from the server. If the reported version is
// different from the bundled BUILD_VERSION, force a hard reload of the
// task pane so stale cached JS stops being a problem. The sessionStorage
// guard prevents a reload loop when the server version is also stale.
async function autoUpdateCheck() {
  try {
    const url = './build-version.txt?_=' + Date.now();
    const resp = await fetch(url, { cache: 'no-store' });
    if (!resp.ok) return;
    const latest = (await resp.text()).trim();
    if (!latest || latest === BUILD_VERSION) return;

    const key = 'claudeReloadedTo:' + latest;
    if (sessionStorage.getItem(key)) {
      // Already tried to reload into this version and still don't see it —
      // probably a deeper cache. Don't loop; just log.
      console.warn(`Build ${BUILD_VERSION} but server reports ${latest}; reload did not help.`);
      return;
    }
    sessionStorage.setItem(key, '1');
    console.log(`Build ${BUILD_VERSION} is stale; server has ${latest}. Reloading.`);
    window.location.reload();
  } catch (e) {
    // Network/fetch failures are non-fatal — just skip the check.
  }
}

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
    const state = await readDocState();
    const docText      = state && state.docText      || '';
    const docAnnotated = state && state.docAnnotated || state && state.docText || '';
    const docName      = state && state.docName      || 'Untitled';
    const commentData  = state && state.commentData  || [];

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

    log('info', `Processing ${toProcess.length} @claude comment(s) in "${docName}"... [build ${BUILD_VERSION}]`);

    // Diagnostic: show Claude exactly what we are showing Claude.
    const annotatedPreview = (docAnnotated || '').substring(0, 2000);
    log('info', 'Annotated doc sent to Claude (first 2000 chars):\n' + annotatedPreview +
      (docAnnotated.length > 2000 ? '\n[...truncated]' : ''));

    const bulletCount = (docAnnotated.match(/\[Bullet L\d+\]/g) || []).length;
    if (bulletCount === 0) {
      log('err', 'No real Word bullets detected in this document. ' +
        'The lines that look like bullets are probably manual "• " characters or tab-indented text, ' +
        'not actual Word list items. Claude cannot tell them apart from normal paragraphs.');
    } else {
      log('ok', `${bulletCount} bulleted paragraph(s) detected in the annotated view.`);
    }

    let processed = 0;
    for (const c of toProcess) {
      try {
        const ok = await processOneComment(c, apiKey, docAnnotated, docName, placeholders);
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

// Read document text, comment tree, and an annotated paragraph-by-paragraph
// view (with style + list info) in one Word.run pass.
async function readDocState() {
  return Word.run(async context => {
    const body       = context.document.body;
    const comments   = body.getComments();
    const props      = context.document.properties;
    const paragraphs = body.paragraphs;

    body.load('text');
    props.load('title');
    paragraphs.load('items/text,items/styleBuiltIn');
    comments.load('items/id,items/content,items/resolved,items/authorName,items/replies/items/id,items/replies/items/content,items/replies/items/authorName');
    await context.sync();

    // List membership is a navigation property; load per paragraph.
    for (const p of paragraphs.items) {
      p.listItemOrNullObject.load('level');
    }
    await context.sync();

    const docAnnotated = paragraphs.items.map(p => annotateParagraph(p)).join('\n');

    const commentData = [];
    for (const c of comments.items) {
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
      docText:      body.text || '',
      docAnnotated: docAnnotated,
      docName:      props.title || 'Untitled',
      commentData:  commentData
    };
  });
}

// Format one paragraph as "[Style tag] text" for Claude's context window.
// Only prefixes structurally interesting paragraphs (headings, titles, list
// items); normal body paragraphs pass through unchanged.
function annotateParagraph(p) {
  const text    = p.text || '';
  const builtIn = p.styleBuiltIn || '';

  if (p.listItemOrNullObject && !p.listItemOrNullObject.isNullObject) {
    return `[Bullet L${p.listItemOrNullObject.level}] ${text}`;
  }

  const headingMatch = /^Heading(\d)$/.exec(builtIn);
  if (headingMatch) return `[Heading ${headingMatch[1]}] ${text}`;

  if (builtIn === 'Title')    return `[Title] ${text}`;
  if (builtIn === 'Subtitle') return `[Subtitle] ${text}`;
  if (builtIn === 'Quote')    return `[Quote] ${text}`;

  return text;
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

  // Log Claude's full raw response so we can see exactly what it asked for.
  if (result.edits && result.edits.length > 0) {
    log('info', 'Claude edits:\n' + result.edits.map(e =>
      `  o: ${JSON.stringify((e.original_text || '').substring(0, 100))}\n` +
      `  r: ${JSON.stringify((e.replacement_text || '').substring(0, 100))}`
    ).join('\n'));
  }
  if (result.inserts && result.inserts.length > 0) {
    log('info', 'Claude inserts:\n' + result.inserts.map(i =>
      `  after: ${JSON.stringify((i.after_text || '').substring(0, 100))}\n` +
      `  new:   ${JSON.stringify(i.new_paragraphs || [])}`
    ).join('\n'));
  }

  // Safety rails: drop edits/inserts that would damage the document.
  const safeEdits   = (result.edits   || []).filter(e => isSafeEdit(e));
  const safeInserts = (result.inserts || []).filter(i => isSafeInsert(i));

  let editSummary = '';
  if (safeEdits.length > 0) {
    editSummary = '\n\n' + await applyEdits(safeEdits);
  }

  let insertSummary = '';
  if (safeInserts.length > 0) {
    insertSummary = '\n\n' + await applyInserts(safeInserts);
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

// ─── Safety Rails ─────────────────────────────────────────

// Reject destructive edits before they touch the document.
function isSafeEdit(e) {
  if (!e || !e.original_text || !e.replacement_text) return false;
  if (/\r|\n/.test(e.original_text)) {
    log('err', `Rejected edit: multi-line original_text "${e.original_text.substring(0, 60)}..."`);
    return false;
  }
  if (/\r|\n/.test(e.replacement_text)) {
    log('err', `Rejected edit: replacement_text contains newline(s); would split paragraphs. original="${e.original_text.substring(0, 60)}"`);
    return false;
  }
  if (e.replacement_text.length > 2000) {
    log('err', `Rejected edit: replacement_text > 2000 chars; looks suspicious.`);
    return false;
  }
  return true;
}

// Reject inserts that would dump whole sections into the doc.
function isSafeInsert(i) {
  if (!i || !i.after_text || !Array.isArray(i.new_paragraphs) || i.new_paragraphs.length === 0) {
    return false;
  }
  if (/\r|\n/.test(i.after_text)) {
    log('err', `Rejected insert: multi-line after_text "${i.after_text.substring(0, 60)}..."`);
    return false;
  }
  if (i.new_paragraphs.length > 10) {
    log('err', `Rejected insert: ${i.new_paragraphs.length} new paragraphs in one insert — exceeds safety cap of 10.`);
    return false;
  }
  for (const p of i.new_paragraphs) {
    if (typeof p !== 'string') {
      log('err', `Rejected insert: non-string new_paragraph.`);
      return false;
    }
    if (/\r|\n/.test(p)) {
      log('err', `Rejected insert: new_paragraph contains newline(s). Split into separate entries instead.`);
      return false;
    }
  }
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

      // Walk up to the containing paragraph of the first match.
      const paragraphs = results.items[0].paragraphs;
      paragraphs.load('items');
      await context.sync();

      let anchorPara = paragraphs.items[paragraphs.items.length - 1];

      // Gather formatting info about the anchor so we can diagnose and
      // branch on list vs non-list insertion.
      anchorPara.load('text,styleBuiltIn');
      const anchorList = anchorPara.listOrNullObject;
      const anchorItem = anchorPara.listItemOrNullObject;
      anchorList.load('id');
      anchorItem.load('level');
      await context.sync();

      const anchorInList = !anchorList.isNullObject && !anchorItem.isNullObject;
      const anchorStyle  = anchorPara.styleBuiltIn || '(no style)';
      log('info',
        `Anchor "${(anchorPara.text || '').substring(0, 50)}" — style=${anchorStyle}, ` +
        (anchorInList ? `list id=${anchorList.id} level=${anchorItem.level}` : 'NOT in a list'));

      for (const newText of ins.new_paragraphs) {
        const clean = sanitizeReplacement(newText);
        if (!clean) continue;

        let newPara = null;

        if (anchorInList) {
          // Preferred path: insert directly into the list. In Word for the
          // Web this is the most reliable way to preserve bullet formatting.
          try {
            newPara = anchorList.insertParagraph(clean, 'End');
            await context.sync();
            log('ok', `Inserted "${clean.substring(0, 40)}" via List.insertParagraph(End).`);
          } catch (e) {
            log('err', `List.insertParagraph failed: ${e.message} — falling back.`);
            newPara = null;
          }
        }

        if (!newPara) {
          // Fallback: insert after the anchor paragraph, then try to attach.
          newPara = anchorPara.insertParagraph(clean, 'After');
          await context.sync();

          if (anchorInList) {
            const newItem = newPara.listItemOrNullObject;
            newItem.load('level');
            await context.sync();

            if (newItem.isNullObject) {
              try {
                newPara.attachToList(anchorList.id, anchorItem.level);
                await context.sync();
                log('ok', `Fallback attached "${clean.substring(0, 40)}" to list ${anchorList.id} L${anchorItem.level}.`);
              } catch (e) {
                log('err', `Fallback attachToList failed: ${e.message}`);
              }
            } else {
              log('info', `"${clean.substring(0, 40)}" auto-joined the list at L${newItem.level}.`);
            }
          }
        }

        // Apply whatever style Claude chose for this paragraph. Bullet
        // anchors are left alone — the list insert already carries the
        // correct formatting. For non-list inserts, trust Claude's
        // decision; it has seen the annotated doc and picked a style that
        // fits the context.
        if (!anchorInList && ins.style) {
          const desired = String(ins.style).trim();
          try {
            newPara.style = desired;
            await context.sync();
            log('info', `Applied style "${desired}" to new paragraph.`);
          } catch (e) {
            log('err', `Could not apply style "${desired}": ${e.message}`);
          }
        }

        // Next insert goes after the one we just placed.
        anchorPara = newPara;
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
