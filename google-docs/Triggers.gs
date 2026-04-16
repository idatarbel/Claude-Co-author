// ============================================================
// Claude Co-author — Triggers.gs
// Manages the background time trigger and user settings.
// ============================================================

const TRIGGER_FUNCTION      = 'processAllRecentDocs';
const POLL_INTERVAL_MINUTES = 5;

// ─── Trigger Setup ─────────────────────────────────────────

function setupTrigger() {
  removeTrigger(); // Remove any existing trigger first (prevent duplicates)

  ScriptApp.newTrigger(TRIGGER_FUNCTION)
    .timeBased()
    .everyMinutes(POLL_INTERVAL_MINUTES)
    .create();

  Logger.log('Auto-polling activated. Claude will check docs every ' + POLL_INTERVAL_MINUTES + ' minutes.');
}

function removeTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === TRIGGER_FUNCTION)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ─── API Key Management ────────────────────────────────────

function promptApiKey() {
  const ui     = DocumentApp.getUi();
  const result = ui.prompt(
    '🔑 Set Claude API Key',
    'Enter your Anthropic API key (starts with sk-ant-):\n\nThis is stored securely in your Google account properties.',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const key = result.getResponseText().trim();
  if (!key || !key.startsWith('sk-ant')) {
    ui.alert('⚠️ That doesn\'t look like a valid Anthropic API key. Please check and try again.');
    return;
  }

  PropertiesService.getUserProperties().setProperty('claudeApiKey', key);
  ui.alert('✅ API key saved. Claude Co-author is ready to use.');
}

function getApiKey() {
  return PropertiesService.getUserProperties().getProperty('claudeApiKey') || null;
}

// ─── Model Selection ───────────────────────────────────────

function promptModel() {
  const ui = DocumentApp.getUi();
  const props = PropertiesService.getUserProperties();
  const current = props.getProperty('claudeModel') || DEFAULT_MODEL;

  const result = ui.prompt(
    '🤖 Claude Model',
    'Currently using: ' + current + '\n\n' +
    'Enter a Claude model ID. Common options:\n' +
    '  claude-sonnet-4-5            — recommended\n' +
    '  claude-sonnet-4-6            — latest Sonnet\n' +
    '  claude-opus-4-6              — most capable\n' +
    '  claude-haiku-4-5-20251001    — fastest / cheapest\n\n' +
    'Leave blank to reset to the default (' + DEFAULT_MODEL + ').',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const entered = result.getResponseText().trim();
  if (!entered) {
    props.deleteProperty('claudeModel');
    ui.alert('✅ Reset to default model: ' + DEFAULT_MODEL);
    return;
  }

  if (!/^claude-/.test(entered)) {
    ui.alert('⚠️ That doesn\'t look like a Claude model id (should start with "claude-"). Nothing was saved.');
    return;
  }

  props.setProperty('claudeModel', entered);
  ui.alert('✅ Model set to: ' + entered);
}

// ─── Status Display ────────────────────────────────────────

function showStatus() {
  const triggers    = ScriptApp.getProjectTriggers();
  const pollTrigger = triggers.find(t => t.getHandlerFunction() === TRIGGER_FUNCTION);
  const hasKey      = !!getApiKey();
  const model       = PropertiesService.getUserProperties().getProperty('claudeModel') || (DEFAULT_MODEL + ' (default)');

  const statusLines = [
    `API Key:      ${hasKey ? '✅ Set' : '❌ Not set'}`,
    `Model:        ${model}`,
    `Auto-polling: ${pollTrigger ? `✅ Active (every ${POLL_INTERVAL_MINUTES} min)` : '⏹ Stopped'}`,
    ``,
    `Usage: In any Google Doc, add a comment starting with:`,
    `  @claude <your instruction>`,
    ``,
    `Claude will reply within ${POLL_INTERVAL_MINUTES} minutes, or immediately`,
    `if you use "Process @claude comments now".`
  ];

  DocumentApp.getUi().alert('Claude Co-author Status', statusLines.join('\n'), DocumentApp.getUi().ButtonSet.OK);
}
