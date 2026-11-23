'use strict';

const PROMPT_LIMIT = 12000;
const HISTORY_LIMIT = 5000;
const CATALOG_LIMIT = 4000;
const TOOL_TAG = 'm365_tool_call';

// The model-provider API supplies messages but no stable chat identity.
// Replay only complete recent pairs in a temporary Microsoft turn.
function assemblePrompt(messages, instructions = '') {
  const current = messages.at(-1);
  if (!current || current.role !== 'user' || !current.text.trim()) {
    throw Error('Select a text message to send to Manual M365 Copilot.');
  }
  const head = 'Recent conversation:\n\n';
  const tail = '\n\nCurrent user message:\n' + current.text;
  const omittedNotice = '[Earlier turns omitted.]\n\n';
  const suffix = instructions ? '\n\n' + instructions : '';
  if (current.text.length + suffix.length > PROMPT_LIMIT) {
    throw Error('Shorten the latest message or reduce attached context. The combined request exceeds 12,000 characters.');
  }
  if (messages.length === 1) return current.text + suffix;

  const available = Math.min(HISTORY_LIMIT,
    PROMPT_LIMIT - head.length - tail.length - omittedNotice.length - suffix.length);
  const retained = [];
  let used = 0;
  for (let i = messages.length - 2; i >= 1; i -= 2) {
    const user = messages[i - 1];
    const reply = messages[i];
    if (user.role !== 'user' || reply.role !== 'assistant') break;
    const block = 'User: ' + user.text + '\nAssistant: ' + reply.text + '\n\n';
    if (used + block.length > available) break;
    retained.unshift(block);
    used += block.length;
  }
  if (!retained.length) return current.text + suffix;
  const omitted = retained.length * 2 !== messages.length - 1 ? omittedNotice : '';
  return head + omitted + retained.join('').trimEnd() + tail + suffix;
}

// VS Code chooses which tools to offer. Translate a bounded subset to text
// for BizChat and report at most one call. VS Code runs and approves the call.
function prepareTools(tools, required) {
  const names = new Set();
  const catalog = [];
  let used = 0;
  for (const tool of tools) {
    if (names.size >= 12) break;
    if (!tool || typeof tool.name !== 'string' || !/^[\w.-]{1,100}$/.test(tool.name) || names.has(tool.name)) continue;
    const schema = JSON.stringify(tool.inputSchema ?? { type: 'object' });
    if (typeof schema !== 'string' || schema.length > 1800) continue;
    const record = JSON.stringify({
      name: tool.name, description: String(tool.description ?? '').slice(0, 220), inputSchema: JSON.parse(schema),
    });
    if (used + record.length + 1 > CATALOG_LIMIT) continue;
    names.add(tool.name);
    catalog.push(record);
    used += record.length + 1;
  }
  if (!names.size) {
    if (required) throw Error('VS Code requires a tool call, but no offered tool fits this model’s context budget.');
    return { names, instructions: '' };
  }
  const instruction = required
    ? 'VS Code requires one tool call for this turn.'
    : 'Answer normally unless a tool is necessary.';
  const instructions = [
    instruction,
    'To call an editor tool, reply with one fenced block and no other text:',
    '```' + TOOL_TAG,
    '{"name":"tool_name","input":{}}',
    '```',
    'Choose a name and input from these editor tools. The editor will run it and return the result:',
    ...catalog,
  ].join('\n');
  return { names, instructions };
}

function decodeToolCall(answer, offered, required) {
  const match = /^\s*```m365_tool_call[ \t]*\r?\n([\s\S]*?)\r?\n```\s*$/.exec(answer);
  if (!match) {
    if (required) throw Error('Microsoft returned text instead of the required tool call.');
    return null;
  }
  let call;
  try { call = JSON.parse(match[1]); } catch { throw Error('Microsoft returned invalid tool-call JSON.'); }
  if (!call || !offered.has(call.name) || !call.input || typeof call.input !== 'object' || Array.isArray(call.input)) {
    throw Error('Microsoft requested a tool or tool input that this turn does not support.');
  }
  return call;
}

module.exports = { assemblePrompt, prepareTools, decodeToolCall };
