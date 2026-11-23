'use strict';

const http = require('node:http');
const https = require('node:https');
const { createHash, randomBytes } = require('node:crypto');
const { TextDecoder } = require('node:util');

const ACCEPT_SUFFIX = '258EAFA5-E914-47DA-95CA-C5AB0DC85B11';
const MAX_MESSAGE = 2 * 1024 * 1024;
const utf8 = new TextDecoder('utf-8', { fatal: true });

function encodeFrame(opcode, value) {
  const body = Buffer.from(value);
  const extra = body.length < 126 ? 0 : body.length <= 65535 ? 2 : 8;
  const frame = Buffer.allocUnsafe(6 + extra + body.length);
  frame[0] = 0x80 | opcode;
  frame[1] = 0x80 | (extra === 0 ? body.length : extra === 2 ? 126 : 127);
  if (extra === 2) frame.writeUInt16BE(body.length, 2);
  if (extra === 8) frame.writeBigUInt64BE(BigInt(body.length), 2);
  const mask = randomBytes(4);
  mask.copy(frame, 2 + extra);
  for (let i = 0; i < body.length; i++) frame[6 + extra + i] = body[i] ^ mask[i % 4];
  return frame;
}

function connect(url, headers, handlers) {
  let socket;
  let closed = false;
  let pending = Buffer.alloc(0);
  let fragments = [];
  let fragmentBytes = 0;
  let fragmentType = 0;
  let closing = false;

  const key = randomBytes(16).toString('base64');
  const expectedAccept = createHash('sha1').update(key + ACCEPT_SUFFIX).digest('base64');
  const network = url.protocol === 'https:' ? https : http; // HTTP is used by the local tests.
  const request = network.request(url, {
    agent: false,
    headers: {
      ...headers,
      Connection: 'Upgrade',
      Upgrade: 'websocket',
      'Sec-WebSocket-Version': '13',
      'Sec-WebSocket-Key': key,
    },
  });

  function dispose() {
    if (closed) return;
    closed = true;
    request.destroy();
    socket?.destroy();
  }

  function fail(message) {
    if (closed) return;
    dispose();
    handlers.error(message);
  }

  function write(opcode, body) {
    if (!socket || socket.destroyed || closed || closing) return false;
    socket.write(encodeFrame(opcode, body));
    return true;
  }

  function deliver(opcode, body) {
    if (opcode === 1 || opcode === 2) handlers.message(utf8.decode(body));
  }

  function receive(chunk) {
    if (closed) return;
    pending = pending.length ? Buffer.concat([pending, chunk]) : chunk;
    try {
      while (!closed && pending.length >= 2) {
        const flags = pending[0];
        const opcode = flags & 15;
        const final = Boolean(flags & 128);
        let length = pending[1] & 127;
        let offset = 2;
        if ((flags & 112) || (pending[1] & 128)) throw Error('Unsupported WebSocket flags.');
        if (length === 126) {
          if (pending.length < 4) return;
          length = pending.readUInt16BE(2);
          offset = 4;
        } else if (length === 127) {
          if (pending.length < 10) return;
          const wide = pending.readBigUInt64BE(2);
          if (wide > BigInt(MAX_MESSAGE)) throw Error('WebSocket response is too large.');
          length = Number(wide);
          offset = 10;
        }
        if (length > MAX_MESSAGE) throw Error('WebSocket response is too large.');
        if (opcode >= 8 && (!final || length > 125)) throw Error('Invalid WebSocket control frame.');
        if (pending.length < offset + length) return;
        const body = pending.subarray(offset, offset + length);
        pending = pending.subarray(offset + length);

        if (opcode === 8) {
          if (length === 1) throw Error('Invalid WebSocket close frame.');
          const code = length >= 2 ? body.readUInt16BE(0) : 1005;
          write(8, body);
          closing = true;
          closed = true;
          socket.end();
          socket.unref();
          handlers.close(code);
          return;
        }
        if (opcode === 9) { write(10, body); continue; }
        if (opcode === 10) continue;
        if (opcode === 0) {
          if (!fragmentType || fragmentBytes + length > MAX_MESSAGE) throw Error('Invalid WebSocket continuation.');
          fragments.push(body);
          fragmentBytes += length;
          if (final) {
            const joined = Buffer.concat(fragments, fragmentBytes);
            const type = fragmentType;
            fragments = [];
            fragmentType = fragmentBytes = 0;
            deliver(type, joined);
          }
        } else if (opcode === 1 || opcode === 2) {
          if (fragmentType) throw Error('Unexpected WebSocket data frame.');
          if (final) deliver(opcode, body);
          else {
            fragments = [body];
            fragmentBytes = length;
            fragmentType = opcode;
          }
        } else {
          throw Error('Unsupported WebSocket frame.');
        }
      }
    } catch {
      fail('Copilot returned an invalid or oversized WebSocket response.');
    }
  }

  request.on('error', () => fail('Could not connect to Microsoft 365 Copilot.'));
  request.on('response', response => {
    if (closed) return;
    const { statusCode, headers: responseHeaders } = response;
    response.destroy();
    dispose();
    handlers.httpError(statusCode, responseHeaders);
  });
  request.on('upgrade', (response, stream, head) => {
    if (closed) { stream.destroy(); return; }
    socket = stream;
    socket.on('error', () => fail('The Copilot connection was interrupted.'));
    socket.on('close', () => {
      if (closed) return;
      closed = true;
      handlers.close(1006);
    });
    if (response.statusCode !== 101 ||
        response.headers.upgrade?.toLowerCase() !== 'websocket' ||
        !response.headers.connection?.split(',').some(x => x.trim().toLowerCase() === 'upgrade') ||
        response.headers['sec-websocket-accept'] !== expectedAccept ||
        response.headers['sec-websocket-extensions'] || response.headers['sec-websocket-protocol']) {
      fail('Microsoft did not accept the WebSocket handshake.');
      return;
    }
    socket.setNoDelay(true);
    socket.on('data', receive);
    handlers.open();
    if (head.length) receive(head);
  });
  request.end();

  return { send: message => write(1, message), dispose };
}

module.exports = { connect };
