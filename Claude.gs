// ============================================================
// Claude Co-author — Claude.gs
// Calls the Anthropic API and returns a structured action.
// ============================================================

/**
 * Calls Claude with the document context and the user's instruction.
 *
 * Returns an object like:
 *   { action: "edit", edit: { original_text, replacement_text }, reply: "..." }
 *   { action: "reply_only", reply: "..." }
 * or null on error.
 */
function callClaude(apiKey, instruction, docContent, quotedText, docName) {
  const systemPrompt = `You are Claude, an AI writing assistant embedded in a Google Doc called "${docName}".
The user has written a comment in the document beginning with "@Claude:" followed by an instruction for you.

You have access to the full document text. The comment may also be anchored to a specific passage.

Your task is to fulfill the instruction. You can either:
1. Respond with information, feedback, or analysis only (reply_only).
2. Make a specific text edit AND explain what you did (edit).

When making an edit, "original_text" must be an exact verbatim substring of the document — copy it character-for-character. Keep it as short as possible while still being uniquely locatable. "replacement_text" is the full new text to replace it with.

RESPOND ONLY WITH A VALID JSON OBJECT — no preamble, no markdown fences:
{
  "action": "edit" | "reply_only",
  "edit": {
    "original_text": "exact verbatim text from document",
    "replacement_text": "the new replacement text"
  },
  "reply": "Concise reply to post on the comment. If you edited, briefly describe what changed and why."
}

If action is "reply_only", set "edit" to null.
Keep "reply" under 300 characters — clear and direct.`;

  // Trim doc content to stay within context limits
  const maxDocChars = 12000;
  const truncated   = docContent.length > maxDocChars;
  const docSnippet  = docContent.substring(0, maxDocChars) + (truncated ? '\n[...document truncated for length...]' : '');

  const userMessage = [
    `Document content:\n---\n${docSnippet}\n---`,
    quotedText ? `Text the comment is anchored to:\n"${quotedText}"` : '',
    `@Claude: ${instruction}`
  ].filter(Boolean).join('\n\n');

  let response;
  try {
    response = UrlFetchApp.fetch(CLAUDE_API_URL, {
      method: 'POST',
      headers: {
        'Content-Type':      'application/json',
        'x-api-key':         apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model:      CLAUDE_MODEL,
        max_tokens: 1500,
        system:     systemPrompt,
        messages:   [{ role: 'user', content: userMessage }]
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('UrlFetchApp error: ' + e.message);
    return null;
  }

  const status = response.getResponseCode();
  if (status !== 200) {
    console.error(`Claude API returned ${status}: ${response.getContentText()}`);
    return null;
  }

  try {
    const body = JSON.parse(response.getContentText());
    let raw = body.content[0].text.trim();

    // Strip accidental markdown code fences
    raw = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
    return JSON.parse(raw);
  } catch (e) {
    console.error('Failed to parse Claude response: ' + e.message);
    return null;
  }
}
