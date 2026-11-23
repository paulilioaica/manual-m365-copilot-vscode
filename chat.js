'use strict';

// BizChat request and SignalR framing adapted from M365CopilotVSCode.
// See NOTICE for the source revision.
const { randomUUID } = require('node:crypto');
const { connect } = require('./socket');

const HOST = 'substrate.office.com';
const ORIGIN = 'https://m365.cloud.microsoft';
const ROUTES = { included: 'OfficeWebIncludedCopilot', paid: 'OfficeWebPaidCopilot' };
const RS = '\x1e';
const STOP = { arguments: [{}], invocationId: '1', target: 'stop', type: 1 };
const TONE = /^[a-z0-9][a-z0-9_-]{0,79}$/i;
const JWT = /^eyJ[\w-]*\.[\w-]+\.[\w-]+$/;
const ID = /^[\w-]{1,128}$/;
const SERVICE_CODE = /^[\w.-]{1,80}$/;
const personalTenants = new Set(['9188040d-6c67-4c5b-b112-36a304b66dad', '84df9e7f-e9f6-40af-b435-aaaaaaaaaaaa']);

function parseCredential(value) {
  let token = value.trim().replace(/^Bearer\s+/i, '');
  let route;
  if (token.length > 64 * 1024) throw Error('The pasted value is too long.');
  if (/^wss?:/i.test(token)) {
    let url;
    try { url = new URL(token); } catch { throw Error('The pasted Copilot URL is invalid.'); }
    if (url.protocol !== 'wss:' || url.hostname !== HOST || !/^\/m365Copilot\/Chathub\//i.test(url.pathname)) {
      throw Error('Paste a Microsoft 365 BizChat access token or its Copilot chat URL.');
    }
    route = url.searchParams.get('scenario') === ROUTES.paid ? 'paid' : 'included';
    token = url.searchParams.get('access_token') ?? '';
  }
  if (!JWT.test(token)) throw Error('Paste a Microsoft 365 business BizChat access token.');
  let claims;
  try { claims = JSON.parse(Buffer.from(token.split('.')[1], 'base64url').toString('utf8')); }
  catch { throw Error('The pasted token is invalid.'); }
  const { oid, tid, exp, aud, scp, scope } = claims ?? {};
  if (typeof oid !== 'string' || typeof tid !== 'string' || !ID.test(oid) || !ID.test(tid) ||
      personalTenants.has(tid.toLowerCase())) {
    throw Error('A readable business-account token with account and tenant IDs is required.');
  }
  const audiences = Array.isArray(aud) ? aud : [aud];
  const substrate = audiences.some(a => typeof a === 'string' &&
    /^(?:https:\/\/)?substrate\.(?:office\.com|svc\.cloud\.microsoft)\/?$/i.test(a));
  if (!substrate && !/\b(?:sydney|m365chat)(?:[._\s]|$)/i.test(String(scp ?? scope ?? ''))) {
    throw Error('The token has the wrong audience. Paste a BizChat token for Microsoft 365.');
  }
  if (!Number.isFinite(exp) || exp * 1000 <= Date.now() + 60000) {
    throw Error('The token has expired or is about to expire. Paste a fresh token.');
  }
  return { token, oid, tid, exp, route };
}

function buildRequest(credential, chat, config, text, base) {
  const url = new URL(base);
  url.pathname += '/' + credential.oid + '@' + credential.tid;
  const requestId = randomUUID();
  const scenario = ROUTES[config.route];
  url.search = new URLSearchParams({
    access_token: credential.token,
    ConversationId: chat.id,
    chatsessionid: requestId,
    clientrequestid: requestId,
    'X-SessionId': requestId,
    source: 'officeweb',
    product: 'Office',
    agentHost: 'Bizchat.FullScreen',
    scenario,
  });
  if (config.route === 'paid') url.searchParams.set('licenseType', 'Premium');
  if (!config.saveHistory) url.searchParams.set('disableMemory', '1');
  const headers = {
    Origin: ORIGIN,
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36',
    'X-RoutingParameter-SessionKey': chat.key,
  };
  const args = {
    source: 'officeweb',
    tone: config.tone,
    streamingMode: 'ConciseWithPadding',
    isStartOfSession: chat.turns === 0,
    allowedMessageTypes: ['Chat', 'Suggestion', 'Progress', 'Disengaged', 'EndOfRequest'],
    clientInfo: { clientPlatform: 'mcmcopilot-web', clientAppName: 'Office' },
    conversationId: chat.id,
    message: { author: 'user', messageType: 'Chat', locale: 'en-US', text },
  };
  return { url, headers, args, scenario };
}

function serviceError(value, code, message, refused = false) {
  const quota = value === 'OutOfCredits' || /priority access|daily chat limit|acceso prioritario/i.test(message ?? '');
  if (quota) return { kind: 'quota', label: 'Copilot reported that the selected allowance is exhausted.' };
  if (value === 'Throttled' || code === 'PerUserThrottled' || /^throttled\b/i.test(message ?? '')) {
    return { kind: 'throttle', label: 'Copilot is temporarily limiting requests.' };
  }
  if (value === 'InvalidCopilotLicense' || code === 'InvalidCopilotLicense' || /valid licen[cs]e/i.test(message ?? '')) {
    return { kind: 'license', label: 'Copilot declined this request for the selected license or model.' };
  }
  if (value === 'InvalidConversation' || /conversation.{0,35}(?:expired|limit|not found)/i.test(message ?? '')) {
    return { kind: 'conversation', label: 'Copilot ended this conversation. Start a new chat.' };
  }
  return { kind: refused ? 'rejected' : 'interrupted', label: 'Copilot declined or interrupted the request.' };
}

function runTurn(request, stats, signal, onText) {
  return new Promise(resolve => {
    let channel;
    let phase = 'connecting';
    let settled = false;
    let buffered = '';
    let currentId = '';
    const snapshots = new Map();
    let timeout;
    let stopTimer;
    let confirmed = null;
    const hardLimit = setTimeout(() => interrupt('Copilot did not finish within five minutes.'), 300000);

    function finish(outcome, stop = false) {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      clearTimeout(hardLimit);
      signal.removeEventListener('abort', abort);
      if (stop && channel?.send(JSON.stringify(STOP) + RS)) {
        stopTimer = setTimeout(() => channel.dispose(), 1500);
        stopTimer.unref();
      } else channel?.dispose();
      resolve(outcome);
    }

    function interrupt(label) {
      finish({ kind: 'interrupted', label }, phase === 'sent');
    }

    function abort() {
      finish({ kind: 'stopped', label: 'Stopped.' }, phase === 'sent');
    }

    function refreshTimeout(ms) {
      clearTimeout(timeout);
      timeout = setTimeout(() => interrupt('The Copilot connection timed out.'), ms);
    }

    function complete() {
      if (stats.outputChars) finish({ kind: 'success' });
      else finish({ kind: 'interrupted', label: 'Copilot ended the turn without an answer.' });
    }

    function emit(delta) {
      if (!delta || settled) return;
      if (stats.outputChars + delta.length > 250000) { interrupt('The answer exceeded the local display limit.'); return; }
      stats.outputChars += delta.length;
      onText(delta);
    }

    function append(delta) {
      snapshots.set(currentId, (snapshots.get(currentId) ?? '') + delta);
      emit(delta);
    }

    function snapshot(message) {
      const id = typeof message.messageId === 'string' ? message.messageId : '';
      const text = message.text;
      let known = snapshots.get(id);
      if (known === undefined && id && snapshots.size === 1) {
        const anonymous = snapshots.get('');
        if (anonymous && text.startsWith(anonymous)) { known = anonymous; snapshots.delete(''); }
      }
      currentId = id;
      snapshots.set(id, text);
      emit(known !== undefined && text.startsWith(known) ? text.slice(known.length) : text);
    }

    function observe(payload) {
      const metering = payload.throttling?.metering;
      if (metering && typeof metering === 'object') {
        for (const [name, counter] of Object.entries(metering).slice(0, 25)) {
          if (SERVICE_CODE.test(name) && Number.isFinite(counter?.remainingAllowance)) {
            stats.remainingAllowance[name] = counter.remainingAllowance;
          }
        }
      }
      const used = payload.throttling?.numUserMessagesInConversation;
      const max = payload.throttling?.maxNumUserMessagesInConversation;
      if (Number.isFinite(used) && Number.isFinite(max)) stats.conversationMessages = { used, max };
      const result = payload.result;
      if (result && typeof result.value === 'string') {
        if (SERVICE_CODE.test(result.value)) stats.result = result.value;
        if (typeof result.errorCode === 'string' && SERVICE_CODE.test(result.errorCode)) stats.errorCode = result.errorCode;
        if (result.value !== 'Success') {
          finish(serviceError(result.value, result.errorCode, result.message, true));
          return;
        }
        confirmed = 'success';
      }
      if (typeof payload.writeAtCursor === 'string') append(payload.writeAtCursor);
      if (!Array.isArray(payload.messages)) return;
      for (const item of payload.messages) {
        if (settled || item?.author !== 'bot') continue;
        if (item.messageType === 'Disengaged') {
          finish({ kind: 'rejected', label: 'Copilot declined to answer this message.' });
          return;
        }
        if (item.messageType === 'EndOfRequest') { complete(); return; }
        if (item.messageType != null && item.messageType !== 'Chat') continue;
        if (typeof item.contentOrigin === 'string' && SERVICE_CODE.test(item.contentOrigin) &&
            !stats.contentOrigins.includes(item.contentOrigin) && stats.contentOrigins.length < 10) {
          stats.contentOrigins.push(item.contentOrigin);
        }
        if (typeof item.text !== 'string' || !item.text) continue;
        if (item.contentOrigin === 'BotConnection' && !stats.outputChars) {
          const error = serviceError(null, null, item.text);
          if (error.kind !== 'interrupted') { stats.result = 'ServiceNotice'; finish(error); return; }
        }
        snapshot(item);
      }
    }

    function processFrame(frame) {
      if (settled) {
        if (stopTimer && (frame.type === 3 || frame.type === 7)) {
          clearTimeout(stopTimer);
          channel.dispose();
        }
        return;
      }
      if (phase === 'handshake') {
        if (frame.error || frame.type !== undefined) {
          finish({ kind: 'rejected', label: 'Copilot did not accept the chat handshake.' });
          return;
        }
        const now = new Date().toISOString();
        const chat = { arguments: [request.args], invocationId: '0', target: 'chat', type: 4 };
        const metrics = { arguments: [{ Timestamps: {
          ConnectionStart: now, ConnectionEstablished: now, UserInputStart: now, UserInputSubmit: now,
        } }], target: 'Metrics', type: 1 };
        if (channel.send(JSON.stringify(chat) + RS + JSON.stringify(metrics) + RS)) {
          phase = 'sent';
          stats.invocations = 1;
          refreshTimeout(90000);
        } else interrupt('The connection closed before the message was sent.');
        return;
      }
      refreshTimeout(90000);
      if (frame.type === 6) { channel.send(JSON.stringify({ type: 6 }) + RS); return; }
      if (frame.type === 7 || (frame.type === 3 && (frame.invocationId === '0' || frame.invocationId === undefined))) {
        if (frame.error) finish(serviceError(null, null, frame.error));
        else complete();
        return;
      }
      if (frame.type !== 1 && frame.type !== 2) return;
      const payloads = Array.isArray(frame.arguments) ? frame.arguments : [frame.item];
      for (const payload of payloads) if (!settled && payload && typeof payload === 'object') observe(payload);
    }

    if (signal.aborted) { abort(); return; }
    signal.addEventListener('abort', abort, { once: true });
    refreshTimeout(15000);
    channel = connect(request.url, request.headers, {
      open() {
        phase = 'handshake';
        channel.send(JSON.stringify({ protocol: 'json', version: 1 }) + RS);
      },
      message(value) {
        buffered += value;
        if (buffered.length > 2 * 1024 * 1024) { interrupt('Copilot returned an oversized reply.'); return; }
        const parts = buffered.split(RS);
        buffered = parts.pop();
        for (const chunk of parts) {
          if (!chunk) continue;
          let frame;
          try { frame = JSON.parse(chunk); } catch { continue; }
          if (frame && typeof frame === 'object') processFrame(frame);
        }
      },
      error: interrupt,
      close(code) {
        clearTimeout(stopTimer);
        if (settled) return;
        if (confirmed === 'success' || (code === 1000 && stats.outputChars)) complete();
        else interrupt('The connection closed before Copilot finished.');
      },
      httpError(status, headers) {
        stats.httpStatus = status;
        const retryAfter = headers['retry-after'];
        if (typeof retryAfter === 'string' && retryAfter.length < 100) {
          const retryTime = /^\d+$/.test(retryAfter) ? Date.now() + Number(retryAfter) * 1000 : Date.parse(retryAfter);
          if (Number.isFinite(retryTime) && retryTime > Date.now() && retryTime <= 8640000000000000) {
            stats.retryAt = new Date(retryTime).toISOString();
          }
        }
        if (status === 401) finish({ kind: 'auth', label: 'Copilot rejected the token. Paste a fresh token.' });
        else if (status === 403) finish({ kind: 'license', label: 'Copilot denied access. Check the token, license and route.' });
        else if (status === 429) finish({ kind: 'throttle', label: 'Copilot temporarily limited requests.' });
        else finish({ kind: 'rejected', label: 'Copilot refused the connection (HTTP ' + status + ').' });
      },
    });
  });
}

class Chat {
  constructor(base = 'https://' + HOST + '/m365Copilot/Chathub') {
    this.base = base;
    this.credential = null;
    this.config = { route: 'included', tone: 'magic', saveHistory: false };
    this.last = null;
    this.total = 0;
    this.retryUntil = 0;
    this.active = null;
    this.reset();
  }

  reset() {
    if (this.active) throw Error('Stop the current reply first.');
    this.conversation = { id: randomUUID(), key: randomUUID(), turns: 0, uncertain: false };
  }

  login(value) {
    if (this.active) throw Error('Stop the current reply first.');
    const credential = parseCredential(value);
    const old = this.credential;
    const sameAccount = old?.oid === credential.oid && old?.tid === credential.tid;
    this.credential = credential;
    if (!sameAccount) {
      this.config = { route: credential.route ?? 'included', tone: 'magic', saveHistory: false };
      this.last = null;
      this.total = 0;
      this.retryUntil = 0;
      this.reset();
    }
    return !sameAccount;
  }

  configure(route, tone, saveHistory) {
    if (this.active) throw Error('Stop the current reply first.');
    if (route !== 'included' && route !== 'paid') throw Error('Choose a valid Copilot license.');
    if (typeof tone !== 'string' || !TONE.test(tone)) throw Error('Enter a valid model tone, such as magic, Chat or Reasoning.');
    this.reset();
    this.config = { route, tone, saveHistory: Boolean(saveHistory) };
  }

  get hasToken() { return Boolean(this.credential && this.credential.exp * 1000 > Date.now() + 60000); }

  validate(text) {
    if (this.active) throw Error('A reply is already running.');
    if (!this.hasToken) throw Error('Paste a current Microsoft 365 Copilot token.');
    if (this.conversation.uncertain) throw Error('Start a new chat after an interrupted or incomplete turn.');
    if (this.retryUntil > Date.now()) throw Error('Copilot asked you to wait until ' + new Date(this.retryUntil).toLocaleTimeString() + '.');
    if (typeof text !== 'string' || !text.trim() || text.length > 24000) throw Error('Enter a message up to 24,000 characters.');
    if (text.includes(this.credential.token) || /access_token=|eyJ[\w-]*\.[\w-]+\.[\w-]+/.test(text)) {
      throw Error('Use Paste Token for credentials.');
    }
  }

  async send(text, onText) {
    this.validate(text);
    this.active = new AbortController();
    const request = buildRequest(this.credential, this.conversation, this.config, text, this.base);
    const stats = {
      scenario: request.scenario, tone: request.args.tone, invocations: 0, inputChars: text.length,
      outputChars: 0, remainingAllowance: Object.create(null), contentOrigins: [],
    };
    this.last = stats;
    try {
      const outcome = await runTurn(request, stats, this.active.signal, onText);
      stats.outcome = outcome.kind;
      if (stats.retryAt) this.retryUntil = Date.parse(stats.retryAt);
      if (outcome.kind === 'auth' && this.credential) this.credential.exp = 0;
      if (outcome.kind === 'success') this.conversation.turns++;
      else if (stats.invocations && ['conversation', 'stopped', 'interrupted'].includes(outcome.kind)) {
        this.conversation.uncertain = true;
      }
      return outcome;
    } finally {
      this.total += stats.invocations;
      this.active = null;
    }
  }

  stop() { this.active?.abort(); }
}

module.exports = { Chat, parseCredential };
