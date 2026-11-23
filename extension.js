'use strict';

const vscode = require('vscode');
const { randomUUID } = require('node:crypto');
const { Chat, parseCredential } = require('./chat');
const { version } = require('./package.json');
const { assemblePrompt, prepareTools, decodeToolCall } = require('./prompt');
const { Continuations, MARKER_MIME, decodeMarker } = require('./continuation');

const VENDOR = 'manual-m365';
// Match the upstream provider's planning limits for the VS Code host.
// The independent cap in prompt.js still limits what is sent to Microsoft.
const HOST_INPUT_TOKENS = 128_000;
const HOST_OUTPUT_TOKENS = 16_000;
const MODELS = [
  { id: 'manual-m365-auto', name: 'M365 Copilot (Manual Token)', tone: 'magic' },
  { id: 'manual-m365-quick', name: 'M365 Copilot (Quick Response)', tone: 'Chat' },
  { id: 'manual-m365-reasoning', name: 'M365 Copilot (Think Deeper)', tone: 'Reasoning' },
];
const PASTE_TOKEN = 'Run “Manual M365 Copilot: Paste Token” in the Command Palette.';

function activate(context) {
  let credential = null;
  let route = 'included';
  let active = false;
  let retryUntil = 0;
  let last;
  let total = 0;
  let tokenInputOpen = false;
  let output;
  const changed = new vscode.EventEmitter();
  const continuations = new Continuations();
  const supportsMarkers = typeof vscode.LanguageModelDataPart === 'function';

  function freshChat(tone) {
    const chat = new Chat();
    chat.credential = credential;
    chat.configure(route, tone, false);
    return chat;
  }

  function requireIdle() {
    if (active) throw Error('Wait for the current reply or stop it first.');
  }

  async function send(chat, message, token, onText) {
    requireIdle();
    if (!credential || !chat.hasToken) throw Error(PASTE_TOKEN);
    if (Date.now() < retryUntil) throw Error('Microsoft asked you to wait until ' + new Date(retryUntil).toLocaleTimeString() + '.');
    if (token.isCancellationRequested) return { kind: 'stopped', label: 'Stopped.' };
    const stop = token.onCancellationRequested(() => chat.stop());
    active = true;
    const previous = chat.last;
    const previousCount = chat.total;
    try {
      const outcome = await chat.send(message, onText);
      return outcome;
    } finally {
      stop.dispose();
      active = false;
      if (chat.last !== previous) {
        last = chat.last;
        total += chat.total - previousCount;
        retryUntil = Math.max(retryUntil, chat.retryUntil);
      }
    }
  }

  function textParts(message) {
    return message.content
      .filter(part => part instanceof vscode.LanguageModelTextPart)
      .map(part => part.value).join('\n');
  }

  function readTranscript(messages) {
    const transcript = [];
    for (const message of messages) {
      const role = message.role === vscode.LanguageModelChatMessageRole.User ? 'user'
        : message.role === vscode.LanguageModelChatMessageRole.Assistant ? 'assistant' : null;
      if (!role) continue;
      const toolCalls = [];
      const toolResults = [];
      const text = message.content.map(part => {
        if (part instanceof vscode.LanguageModelTextPart) return part.value;
        if (part instanceof vscode.LanguageModelToolCallPart) {
          toolCalls.push(part.callId);
          const input = JSON.stringify(part.input) ?? '{}';
          return 'Tool call ' + part.name + ' (' + part.callId + '): ' + input.slice(0, 2000);
        }
        if (part instanceof vscode.LanguageModelToolResultPart) {
          toolResults.push(part.callId);
          const result = textParts(part);
          const clipped = result.length > 2500 ? result.slice(0, 2500) + '\n[Tool result shortened.]' : result;
          return 'Tool result (' + part.callId + '): ' + (clipped || '[No text result]');
        }
        return '';
      }).filter(Boolean).join('\n');
      if (text.trim()) transcript.push({ role, text, toolCalls, toolResults });
    }
    return transcript;
  }

  function findMarker(modelId, messages) {
    // The Copilot bridge prepends its authoritative prior response as an
    // assistant data part. Prefer it to any older marker in the transcript.
    const assistant = vscode.LanguageModelChatMessageRole.Assistant;
    const markerIn = message => {
      if (message?.role !== assistant) return undefined;
      for (const part of message.content) {
        if (!(part instanceof vscode.LanguageModelDataPart)) continue;
        const id = decodeMarker(part, modelId);
        if (id) return id;
      }
    };
    const first = messages[0];
    const leading = first?.content.every(part => part instanceof vscode.LanguageModelDataPart)
      ? markerIn(first) : undefined;
    if (leading) return leading;
    for (let i = messages.length - 1; i >= 0; i--) {
      const id = markerIn(messages[i]);
      if (id) return id;
    }
  }

  function currentInput(messages, expectedToolId) {
    const lastAssistant = messages.findLastIndex(message => message.role === vscode.LanguageModelChatMessageRole.Assistant);
    const tail = readTranscript(messages.slice(lastAssistant + 1));
    if (!tail.length) throw Error('There is no new user message or tool result to send.');
    const toolResults = tail.flatMap(message => message.toolResults);
    if (expectedToolId ? toolResults.length !== 1 || toolResults[0] !== expectedToolId : toolResults.length) {
      throw Error('The tool result does not match this Microsoft 365 conversation. Start a new VS Code chat.');
    }
    // Some hosts append several user messages together. The newest text
    // is the next turn. Tool results may be split across adjacent parts.
    return expectedToolId
      ? { role: 'user', text: tail.map(message => message.text).join('\n') }
      : tail.at(-1);
  }

  const provider = {
    onDidChangeLanguageModelChatInformation: changed.event,
    provideLanguageModelChatInformation() {
      return MODELS.map(({ id, name }) => ({
        id, name, family: VENDOR, version,
        maxInputTokens: HOST_INPUT_TOKENS,
        maxOutputTokens: HOST_OUTPUT_TOKENS,
        maxContextWindowTokens: HOST_INPUT_TOKENS + HOST_OUTPUT_TOKENS,
        capabilities: { imageInput: false, toolCalling: true },
        isUserSelectable: true,
      }));
    },
    provideTokenCount(_model, value) {
      return Math.ceil((typeof value === 'string' ? value : textParts(value)).length / 3.5);
    },
    async provideLanguageModelChatResponse(model, messages, options, progress, token) {
      const selected = MODELS.find(item => item.id === model.id);
      if (!selected) throw Error('Unknown Manual M365 Copilot model.');
      requireIdle();
      if (!credential) throw Error(PASTE_TOKEN);
      const id = supportsMarkers ? findMarker(selected.id, messages) : undefined;
      const previous = id ? continuations.get(id, selected.id) : undefined;
      const required = options.toolMode === vscode.LanguageModelChatToolMode.Required;
      const offered = prepareTools(options.tools ?? [], required);
      // With a verified continuation the backend already has the transcript.
      const input = previous ? [currentInput(messages, previous.toolId)] : readTranscript(messages);
      if (!previous && input.at(-1)?.toolResults.length) {
        const prior = input.at(-2);
        if (prior?.role !== 'assistant' || !input.at(-1).toolResults.every(id => prior.toolCalls.includes(id))) {
          throw Error('VS Code did not supply the conversation or tool call for this result. Start a new chat.');
        }
      }
      let instructions = offered.instructions;
      if (previous?.instructions === instructions) instructions = '';
      else if (previous?.instructions && !instructions) {
        instructions = 'No editor tools are available for this turn. Answer directly.';
      }
      const prompt = assemblePrompt(input, instructions);
      const newMessageChars = input.at(-1).text.length;
      const instructionChars = instructions ? instructions.length + 2 : 0;
      const replayedContextChars = prompt.length - newMessageChars - instructionChars;
      if (!previous && input.at(-1).toolResults.length && !replayedContextChars) {
        throw Error('The previous tool call does not fit the context budget. Start a new VS Code chat.');
      }
      const chat = previous?.chat ?? freshChat(selected.tone);
      chat.credential = credential;
      const before = chat.total;
      let answer = '';
      let callId;
      // Reserve the marker for this attempt. Reuse is safe only if nothing
      // went over the wire. In particular, a fork cannot mutate the old turn.
      if (id) continuations.remove(id);
      const priorStats = chat.last;
      try {
        const outcome = await send(chat, prompt, token, delta => {
          if (offered.names.size) answer += delta;
          else progress.report(new vscode.LanguageModelTextPart(delta));
        });
        if (outcome.kind === 'stopped') return;
        if (outcome.kind !== 'success') throw Error(outcome.label);
        if (offered.names.size) {
          const call = decodeToolCall(answer, offered.names, required);
          if (call) {
            callId = randomUUID();
            progress.report(new vscode.LanguageModelToolCallPart(callId, call.name, call.input));
          } else progress.report(new vscode.LanguageModelTextPart(answer));
        }
        if (supportsMarkers) {
          const marker = continuations.issue(selected.id, chat, offered.instructions, callId);
          progress.report(new vscode.LanguageModelDataPart(marker, MARKER_MIME));
        }
      } finally {
        chat.credential = null;
        if (id && chat.total === before) continuations.restore(id, previous);
        if (chat.last && chat.last !== priorStats) {
          Object.assign(chat.last, {
            contextMode: previous ? 'continued' : replayedContextChars ? 'bounded-replay' : 'new',
            newMessageChars, replayedContextChars, instructionChars,
          });
        }
      }
    },
  };

  const command = (id, fn) => vscode.commands.registerCommand(id, async () => {
    try { await fn(); }
    catch (error) { void vscode.window.showErrorMessage(error.message); }
  });

  context.subscriptions.push(
    changed,
    { dispose: () => continuations.clear() },
    vscode.lm.registerLanguageModelChatProvider(VENDOR, provider),
    command('manualM365.pasteToken', async () => {
      if (tokenInputOpen) return;
      requireIdle();
      tokenInputOpen = true;
      try {
        const value = await vscode.window.showInputBox({
          title: 'Paste Microsoft 365 BizChat token', password: true, ignoreFocusOut: true,
          prompt: 'The token stays in memory until the VS Code window closes.',
        });
        if (!value) return;
        requireIdle();
        const parsed = parseCredential(value);
        if (parsed.oid === credential?.oid && parsed.tid === credential?.tid) {
          credential = parsed;
        } else {
          continuations.clear();
          total = 0;
          last = undefined;
          retryUntil = 0;
          route = parsed.route ?? 'included';
          credential = parsed;
        }
        changed.fire();
        void vscode.window.showInformationMessage('M365 token ready. Select M365 Copilot (Manual Token) in the VS Code Chat model picker.');
      } finally { tokenInputOpen = false; }
    }),
    command('manualM365.forgetToken', () => {
      requireIdle();
      credential = null;
      continuations.clear();
      total = 0;
      last = undefined;
      retryUntil = 0;
      changed.fire();
      void vscode.window.showInformationMessage('M365 token forgotten.');
    }),
    command('manualM365.configure', async () => {
      requireIdle();
      const choice = await vscode.window.showQuickPick([
        { label: 'Included Copilot Chat', picked: route === 'included', value: 'included' },
        { label: 'Paid Microsoft 365 Copilot', picked: route === 'paid', value: 'paid' },
      ], { title: 'Choose your Microsoft 365 Copilot license route' });
      if (!choice || choice.value === route) return;
      requireIdle();
      route = choice.value;
      continuations.clear();
    }),
    command('manualM365.diagnostics', () => {
      if (!output) {
        output = vscode.window.createOutputChannel('Manual M365 Copilot');
        context.subscriptions.push(output);
      }
      output.clear();
      output.appendLine(JSON.stringify({ chatInvocations: total, last: last ?? null }, null, 2));
      output.show();
    }),
  );
}
module.exports = { activate };
