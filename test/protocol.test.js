'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const { createHash } = require('node:crypto');
const { Chat, parseCredential } = require('../chat');

const RS = '\x1e';
const b64 = obj => Buffer.from(JSON.stringify(obj)).toString('base64url');
const token = (extra = {}) => b64({ alg: 'none' }) + '.' + b64({
  oid: 'fake-user', tid: 'fake-tenant', aud: 'https://substrate.office.com', exp: Math.floor(Date.now() / 1000) + 7200,
  ...extra,
}) + '.fake_signature';

function frame(body, opcode = 1, fin = true) {
  const data = Buffer.from(body);
  assert.ok(data.length <= 65535);
  if (data.length < 126) return Buffer.concat([Buffer.from([(fin ? 128 : 0) | opcode, data.length]), data]);
  const header = Buffer.from([(fin ? 128 : 0) | opcode, 126, 0, 0]);
  header.writeUInt16BE(data.length, 2);
  return Buffer.concat([header, data]);
}

async function mock(t, onUpgrade) {
  const server = http.createServer();
  const sockets = new Set();
  const calls = [];
  server.on('upgrade', (req, socket, head) => {
    sockets.add(socket);
    socket.on('close', () => sockets.delete(socket));
    socket.on('error', () => {});
    onUpgrade(req, socket, head, calls);
  });
  await new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', resolve);
  });
  t.after(async () => {
    for (const socket of sockets) socket.destroy();
    await new Promise(resolve => server.close(resolve));
  });
  const chat = new Chat('http://127.0.0.1:' + server.address().port + '/m365Copilot/Chathub');
  chat.login(token());
  return { chat, calls };
}

function accept(onChat, onOther) {
  return (req, socket, head, calls) => {
    const accept = createHash('sha1').update(req.headers['sec-websocket-key'] + '258EAFA5-E914-47DA-95CA-C5AB0DC85B11').digest('base64');
    socket.write('HTTP/1.1 101 Switching Protocols\r\nConnection: Upgrade\r\nUpgrade: websocket\r\nSec-WebSocket-Accept: ' + accept + '\r\n\r\n');
    const send = obj => socket.write(frame(JSON.stringify(obj) + RS));
    let buffer = head;
    socket.on('data', data => {
      buffer = Buffer.concat([buffer, data]);
      while (buffer.length >= 6) {
        assert.ok(buffer[1] & 128, 'client frame is masked');
        const len = buffer[1] & 127;
        if (len === 126 && buffer.length < 8) return;
        const bytes = len === 126 ? buffer.readUInt16BE(2) : len;
        const offset = len === 126 ? 4 : 2;
        if (buffer.length < offset + 4 + bytes) return;
        const opcode = buffer[0] & 15;
        const mask = buffer.subarray(offset, offset + 4);
        const body = Buffer.from(buffer.subarray(offset + 4, offset + 4 + bytes));
        for (let i = 0; i < body.length; ++i) body[i] ^= mask[i % 4];
        buffer = buffer.subarray(offset + 4 + bytes);
        if (opcode !== 1) { onOther?.({ opcode, body }); continue; }
        for (const segment of body.toString().split(RS).filter(Boolean)) {
          const obj = JSON.parse(segment);
          if (obj.protocol) send({});
          else if (obj.target === 'chat') {
            const call = { req, args: obj.arguments[0] };
            calls.push(call);
            onChat(call, socket, send);
          } else onOther?.({ obj, send });
        }
      }
    });
  };
}

test('continuation uses the same context and sends only the new text', async t => {
  const { chat, calls } = await mock(t, accept((call, socket, send) => {
    const item = JSON.stringify({ type: 2, item: { writeAtCursor: 'orchid' } }) + RS;
    socket.write(Buffer.concat([frame(item.slice(0, 5), 1, false), frame('ping', 9), frame(item.slice(5), 0)]));
    send({ type: 2, item: { messages: [{ author: 'bot', messageId: 'id', text: 'orchid' }], result: { value: 'Success' } } });
    send({ type: 3, invocationId: '0' });
  }));
  const ask = async prompt => {
    let text = '';
    assert.equal((await chat.send(prompt, delta => { text += delta; })).kind, 'success');
    assert.equal(text, 'orchid');
  };
  chat.configure('paid', 'magic', true);
  await ask('Remember orchid ' + 'x'.repeat(270));
  await ask('What was the word?');
  assert.equal(chat.total, 2);
  const [first, second] = calls;
  assert.equal(first.args.isStartOfSession, true);
  assert.equal(second.args.isStartOfSession, false);
  assert.equal(first.args.conversationId, second.args.conversationId);
  assert.equal(first.req.headers['x-routingparameter-sessionkey'], second.req.headers['x-routingparameter-sessionkey']);
  assert.equal(second.args.message.text, 'What was the word?');
  assert.equal(second.args.previousMessages, undefined);
  const url = new URL(second.req.url, 'http://localhost');
  assert.equal(url.searchParams.get('scenario'), 'OfficeWebPaidCopilot');
  assert.equal(url.searchParams.get('licenseType'), 'Premium');
  assert.equal(url.searchParams.has('disableMemory'), false);
  chat.reset();
  await ask('New topic');
  assert.notEqual(calls[2].args.conversationId, second.args.conversationId);
});

test('quota refusal records returned counters and makes no retry', async t => {
  const { chat, calls } = await mock(t, accept((_call, _socket, send) => send({
    type: 2, item: {
      result: { value: 'OutOfCredits' },
      throttling: { metering: { Daily: { remainingAllowance: 0 } }, numUserMessagesInConversation: 2, maxNumUserMessagesInConversation: 35 },
    },
  })));
  const outcome = await chat.send('Hello', () => assert.fail('quota message should not be emitted as a reply'));
  assert.equal(outcome.kind, 'quota');
  assert.equal(calls.length, 1);
  assert.equal(chat.last.remainingAllowance.Daily, 0);
  assert.deepEqual(chat.last.conversationMessages, { used: 2, max: 35 });
  assert.equal(chat.conversation.turns, 0);
  assert.equal(new URL(calls[0].req.url, 'http://localhost').searchParams.get('disableMemory'), '1');
});

test('429 honors Retry-After without submitting or retrying chat', async t => {
  let upgrades = 0;
  const { chat } = await mock(t, (_req, socket) => {
    upgrades++;
    socket.end('HTTP/1.1 429 Too Many Requests\r\nRetry-After: 60\r\nConnection: close\r\nContent-Length: 0\r\n\r\n');
  });
  assert.equal((await chat.send('Hello', () => {})).kind, 'throttle');
  assert.equal(chat.total, 0);
  assert.equal(chat.last.httpStatus, 429);
  await assert.rejects(chat.send('Again', () => {}), /asked you to wait/);
  assert.equal(upgrades, 1);
});

test('stopping midway sends the stop invocation and prevents silent replay', async t => {
  let didStop;
  const stop = new Promise(resolve => { didStop = resolve; });
  const { chat } = await mock(t, accept((_call, _socket, send) =>
    send({ type: 2, item: { writeAtCursor: 'partial' } }), ({ obj, send }) => {
    if (obj?.target === 'stop') {
      send({ type: 3, invocationId: '1' });
      didStop();
    }
  }));
  assert.equal((await chat.send('Hello', () => chat.stop())).kind, 'stopped');
  await Promise.race([stop, new Promise((_, reject) => setTimeout(() => reject(Error('No stop sent')), 1000).unref())]);
  await assert.rejects(chat.send('Again', () => {}), /Start a new chat/);
});

test('rejects expired and wrong-audience tokens locally', () => {
  assert.throws(() => parseCredential(token({ exp: 1 })), /expired/);
  assert.throws(() => parseCredential(token({ aud: 'https://graph.microsoft.com' })), /audience/);
});

// Exercise the installed extension entry point and real Chat transport with
// the smallest possible VS Code API stand-in and a local SignalR server.
const { readFileSync } = require('node:fs');
const { createRequire } = require('node:module');
const { compileFunction } = require('node:vm');
const extensionPath = require.resolve('../extension');
const extensionRequire = createRequire(extensionPath);

class TextPart { constructor(value) { this.value = value; } }
class DataPart { constructor(data, mimeType) { Object.assign(this, { data, mimeType }); } }
class ToolCallPart { constructor(callId, name, input) { Object.assign(this, { callId, name, input }); } }
class ToolResultPart { constructor(callId, content) { Object.assign(this, { callId, content }); } }
const user = (...parts) => ({ role: 1, content: parts.map(p => typeof p === 'string' ? new TextPart(p) : p) });
const assistant = (...parts) => ({ role: 2, content: parts.map(p => typeof p === 'string' ? new TextPart(p) : p) });
const noopDisposable = { dispose() {} };
const uncancelled = { isCancellationRequested: false, onCancellationRequested: () => noopDisposable };
function getText(parts) { return parts.filter(p => p instanceof TextPart).map(p => p.value).join(''); }
function getMarker(parts) {
  const p = parts.find(part => part instanceof DataPart && part.mimeType === 'stateful_marker');
  assert.ok(p, 'an opaque continuation marker was emitted');
  return p;
}

async function testProvider(endpoint, options = {}) {
  let provider;
  let diagnostics;
  let pasted = token();
  const commands = new Map();
  const vscode = {
    EventEmitter: class { fire() {} },
    LanguageModelTextPart: TextPart,
    LanguageModelDataPart: DataPart,
    LanguageModelToolCallPart: ToolCallPart,
    LanguageModelToolResultPart: ToolResultPart,
    LanguageModelChatMessageRole: { User: 1, Assistant: 2 },
    LanguageModelChatToolMode: { Required: 2 },
    lm: { registerLanguageModelChatProvider(_vendor, p) { provider = p; return noopDisposable; } },
    commands: { registerCommand(name, fn) { commands.set(name, fn); return noopDisposable; } },
    window: {
      showInputBox: async () => pasted,
      showQuickPick: async choices => choices[1],
      showInformationMessage() {},
      showErrorMessage(message) { assert.fail(message); },
      createOutputChannel: () => ({
        clear() {}, show() {}, appendLine(text) { diagnostics = JSON.parse(text); },
        ...noopDisposable,
      }),
    },
  };
  const redirect = name => {
    if (name === 'vscode') return vscode;
    if (name === './chat') return { parseCredential, Chat: class extends Chat { constructor() { super(endpoint); } } };
    return extensionRequire(name);
  };
  const module = { exports: {} };
  compileFunction(readFileSync(extensionPath, 'utf8'), ['require', 'module'])(redirect, module);
  module.exports.activate({ subscriptions: [] });
  const command = name => commands.get('manualM365.' + name)();
  await command('pasteToken');
  const model = provider.provideLanguageModelChatInformation()[0];
  return {
    model, command,
    async run(messages, selectedOptions = options, cancellation = uncancelled) {
      const parts = [];
      await provider.provideLanguageModelChatResponse(model, messages, selectedOptions, {
        report(part) { parts.push(part); },
      }, cancellation);
      return parts;
    },
    async stats() { await command('diagnostics'); return diagnostics; },
    async replaceToken(claims) { pasted = token(claims); await command('pasteToken'); },
  };
}

function answer(send, text) {
  send({ type: 2, item: { writeAtCursor: text } });
  send({ type: 3, invocationId: '0' });
}

test('native provider sends only next text and isolates chats with identical first prompts', async t => {
  const memory = new Map();
  const { chat, calls } = await mock(t, accept(({ args }, _socket, send) => {
    if (args.isStartOfSession) {
      memory.set(args.conversationId, memory.size ? 'dahlia' : 'orchid');
      answer(send, 'I picked a word.');
    } else {
      answer(send, memory.get(args.conversationId));
    }
  }));
  const api = await testProvider(chat.base);
  const opener = 'Pick a word and remember it.';
  const first = getMarker(await api.run([user(opener)]));
  const second = getMarker(await api.run([user(opener)]));
  assert.notEqual(first.data.toString(), second.data.toString());
  // Agent Host supplies a leading marker; it may also include old history.
  const reply = await api.run([
    assistant(first), user(opener), assistant('I picked a word.'), user('What was the word?'),
  ]);
  const other = await api.run([assistant(second), user('What was the word?')]);
  assert.equal(getText(reply), 'orchid');
  assert.equal(getText(other), 'dahlia');
  assert.equal(calls.length, 4);
  assert.notEqual(calls[0].args.conversationId, calls[1].args.conversationId);
  assert.equal(calls[2].args.conversationId, calls[0].args.conversationId);
  assert.equal(calls[3].args.conversationId, calls[1].args.conversationId);
  for (const call of calls.slice(2)) {
    assert.equal(call.args.message.text, 'What was the word?');
    assert.equal(call.args.isStartOfSession, false);
    assert.equal(new URL(call.req.url, 'http://local').searchParams.get('disableMemory'), '1');
  }
  assert.equal(calls[2].req.headers['x-routingparameter-sessionkey'], calls[0].req.headers['x-routingparameter-sessionkey']);
  assert.equal((await api.stats()).last.contextMode, 'continued');
  assert.equal((await api.stats()).last.replayedContextChars, 0);
  await assert.rejects(api.run([assistant(first), user('Retry from old turn')]), /expired or was already continued/);
  assert.equal(calls.length, 4, 'a stale branch never reached Microsoft');
  const next = getMarker(reply);
  await api.replaceToken({ exp: Math.floor(Date.now() / 1000) + 7000 });
  assert.equal(getText(await api.run([assistant(next), user('Repeat the word.') ])), 'orchid');
});

test('tool continuation verifies call ID and drops unchanged catalog and replay', async t => {
  const { chat, calls } = await mock(t, accept(({ args }, _socket, send) => {
    if (args.isStartOfSession) answer(send, '```m365_tool_call\n{"name":"lookup","input":{"query":"orchid"}}\n```');
    else answer(send, 'The tool found orchid.');
  }));
  const tools = [{ name: 'lookup', description: 'Find a record', inputSchema: { type: 'object' } }];
  const api = await testProvider(chat.base, { tools });
  const parts = await api.run([user('Look up the word.')]);
  const marker = getMarker(parts);
  const call = parts.find(p => p instanceof ToolCallPart);
  assert.equal(call.name, 'lookup');
  await assert.rejects(api.run([assistant(marker), user(new ToolResultPart('wrong', [new TextPart('dahlia')]))]), /does not match/);
  assert.equal(calls.length, 1);
  const result = await api.run([assistant(marker), user(new ToolResultPart(call.callId, [new TextPart('orchid')]))]);
  assert.equal(getText(result), 'The tool found orchid.');
  assert.equal(calls[1].args.conversationId, calls[0].args.conversationId);
  assert.equal(calls[1].args.message.text, 'Tool result (' + call.callId + '): orchid');
  assert.equal((await api.stats()).last.instructionChars, 0);
});

test('missing marker uses bounded replay with a fresh conversation; account switches invalidate old markers', async t => {
  const { chat, calls } = await mock(t, accept((_call, _socket, send) => answer(send, 'Response')));
  const api = await testProvider(chat.base);
  const original = getMarker(await api.run([user('Remember orchid.')]));
  const replay = getMarker(await api.run([
    user('Remember orchid.'), assistant('Response'), user('What was the word?'),
  ]));
  assert.notEqual(calls[0].args.conversationId, calls[1].args.conversationId);
  assert.ok(calls[1].args.message.text.includes('Remember orchid.'));
  assert.equal((await api.stats()).last.contextMode, 'bounded-replay');
  await api.replaceToken({ oid: 'different-account' });
  await assert.rejects(api.run([assistant(original), user('What was the word?')]), /expired or was already continued/);
  await assert.rejects(api.run([assistant(replay), user('What was the word?')]), /expired or was already continued/);
  assert.equal(calls.length, 2);
});

test('a dispatched quota refusal cannot automatically replay the same marker', async t => {
  const { chat, calls } = await mock(t, accept(({ args }, _socket, send) => {
    if (args.isStartOfSession) answer(send, 'Ready');
    else send({ type: 2, item: { result: { value: 'OutOfCredits' } } });
  }));
  const api = await testProvider(chat.base);
  const marker = getMarker(await api.run([user('Hello')]));
  await assert.rejects(api.run([assistant(marker), user('Follow-up')]), /allowance is exhausted/);
  await assert.rejects(api.run([assistant(marker), user('Follow-up')]), /already continued/);
  assert.equal(calls.length, 2);
});
