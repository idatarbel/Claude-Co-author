// ============================================================
// Claude Co-author — Claude.js (Word Online)
// Calls the Anthropic API with web search and returns actions.
// Mirrors Google Docs/Claude.gs — same prompt, same JSON contract.
// ============================================================

const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const CLAUDE_MODEL   = 'claude-sonnet-4-20250514';

async function callClaude(apiKey, instruction, docContent, quotedText, docName, threadHistory, placeholders) {
  docContent   = docContent   || '';
  docName      = docName      || 'Untitled';
  quotedText   = quotedText   || '';
  placeholders = placeholders || [];
  threadHistory = threadHistory || [];

  const systemPrompt = `You are Claude, an AI writing assistant embedded in a Microsoft Word document called "${docName}".

The user has left a comment with an instruction. Execute it fully and directly — do not ask permission, do not hedge, do not suggest the user do the work themselves.

You have web search available. Use it to look up any facts, names, dates, or information you are uncertain about before responding.

Respond with a single JSON object in this exact format:
{
  "action": "edit" | "reply_only",
  "edits": [
    { "original_text": "exact text from document to replace", "replacement_text": "new text" }
  ],
  "comments_to_add": [
    { "quoted_text": "exact text in document to anchor the comment to", "comment": "comment text to add" }
  ],
  "reply": "What you did, concisely."
}

Rules:
- "edits" is an array — return as many replacements as needed to fully complete the task
- Only add an edit if you found REAL, VERIFIED information to replace the placeholder with
- If you did NOT find real information for a placeholder: do NOT add an edit for it — leave it unchanged in the document
- For every placeholder you could not fill with real information: add an entry to comments_to_add with the placeholder as quoted_text and a specific comment telling the user exactly what to research
- "original_text" and "quoted_text" must be copied verbatim from the document
- When placeholders are provided below, use those exact strings verbatim
- replacement_text must be plain prose only — no markdown (**bold**, _italic_), no citation tags (<cite>), no HTML tags of any kind
- When a comment is anchored to specific text and asks to "elaborate", "expand", "add detail", or similar — always use action "edit" to replace that text with an expanded version in the document, not reply_only
- If action is "reply_only", set both "edits" and "comments_to_add" to []`;

  const maxDocChars = 12000;
  const docSnippet  = docContent.substring(0, maxDocChars) +
    (docContent.length > maxDocChars ? '\n[...document truncated...]' : '');

  let historyBlock = '';
  if (threadHistory.length > 0) {
    historyBlock = 'Comment thread so far:\n' +
      threadHistory.map((msg, i) =>
        `  [${i === 0 ? 'Original comment' : 'Reply ' + i}]: ${msg}`
      ).join('\n');
  }

  let placeholderBlock = '';
  if (placeholders.length > 0) {
    placeholderBlock = 'Placeholders found in document (use these EXACT strings as original_text or quoted_text):\n' +
      placeholders.map((p, i) => `  ${i + 1}. ${p}`).join('\n');
  }

  const userMessage = [
    `Document content:\n---\n${docSnippet}\n---`,
    quotedText       ? `Text the comment is anchored to:\n"${quotedText}"` : '',
    historyBlock     || '',
    placeholderBlock || '',
    `@claude: ${instruction}`
  ].filter(Boolean).join('\n\n');

  let messages = [{ role: 'user', content: userMessage }];
  let finalText = null;
  const maxIterations = 5;

  for (let i = 0; i < maxIterations; i++) {
    let response;
    try {
      response = await fetch(CLAUDE_API_URL, {
        method: 'POST',
        headers: {
          'Content-Type':                            'application/json',
          'x-api-key':                               apiKey,
          'anthropic-version':                       '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true'
        },
        body: JSON.stringify({
          model:      CLAUDE_MODEL,
          max_tokens: 4000,
          system:     systemPrompt,
          tools:      [{ type: 'web_search_20250305', name: 'web_search' }],
          messages:   messages
        })
      });
    } catch (e) {
      console.error('fetch error: ' + e.message);
      return null;
    }

    if (!response.ok) {
      const errText = await response.text();
      console.error(`Claude API returned ${response.status}: ${errText}`);
      return null;
    }

    const data = await response.json();
    messages.push({ role: 'assistant', content: data.content });

    if (data.stop_reason === 'end_turn') {
      const textBlocks = data.content.filter(b => b.type === 'text');
      const textBlock  = textBlocks[textBlocks.length - 1];
      if (textBlock) finalText = textBlock.text.trim();
      break;
    }

    if (data.stop_reason === 'tool_use') {
      const hasToolResults = data.content.some(b =>
        b.type === 'tool_result' || b.type === 'web_search_tool_result'
      );
      if (!hasToolResults) {
        const toolResults = data.content
          .filter(b => b.type === 'tool_use')
          .map(b => ({ type: 'tool_result', tool_use_id: b.id, content: '' }));
        if (toolResults.length > 0) {
          messages.push({ role: 'user', content: toolResults });
        }
      }
      continue;
    }

    const textBlocks = data.content.filter(b => b.type === 'text');
    const textBlock  = textBlocks[textBlocks.length - 1];
    if (textBlock) finalText = textBlock.text.trim();
    break;
  }

  if (!finalText) {
    console.error('No final text from Claude after tool loop');
    return null;
  }

  try {
    const jsonMatch = finalText.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      console.error('No JSON found in Claude response. Raw: ' + finalText);
      return null;
    }
    const parsed = JSON.parse(jsonMatch[0]);

    if (parsed.edit && !parsed.edits) parsed.edits = [parsed.edit];
    if (!parsed.edits)           parsed.edits           = [];
    if (!parsed.comments_to_add) parsed.comments_to_add = [];

    return parsed;
  } catch (e) {
    console.error('Failed to parse Claude response: ' + e.message + '\nRaw: ' + finalText);
    return null;
  }
}
