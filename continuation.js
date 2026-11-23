'use strict';

const { randomUUID } = require('node:crypto');

const MARKER_MIME = 'stateful_marker';
const MARKER_ID = /^resp_m365_[a-f0-9-]{36}$/;
const MAX_SESSIONS = 64;
const IDLE_LIMIT_MS = 45 * 60_000;

class Continuations {
  constructor() {
    this.entries = new Map();
  }

  issue(modelId, chat, instructions, toolId) {
    const id = 'resp_m365_' + randomUUID();
    this.entries.set(id, { modelId, chat, instructions, toolId, expires: Date.now() + IDLE_LIMIT_MS });
    if (this.entries.size > MAX_SESSIONS) this.entries.delete(this.entries.keys().next().value);
    return Buffer.from(modelId + '\\' + id);
  }

  get(id, modelId) {
    const entry = this.entries.get(id);
    if (entry?.expires <= Date.now()) this.entries.delete(id);
    if (!entry || entry.expires <= Date.now() || entry.modelId !== modelId) {
      throw Error('This Microsoft 365 conversation has expired or was already continued. Start a new VS Code chat.');
    }
    return entry;
  }

  remove(id) {
    this.entries.delete(id);
  }

  restore(id, entry) {
    if (entry.expires > Date.now()) this.entries.set(id, entry);
  }

  clear() {
    this.entries.clear();
  }
}

function decodeMarker(part, modelId) {
  if (part?.mimeType !== MARKER_MIME || !(part.data instanceof Uint8Array) || part.data.length > 180) return undefined;
  const value = Buffer.from(part.data).toString('utf8');
  const split = value.indexOf('\\');
  const id = value.slice(split + 1);
  if (split <= 0 || value.slice(0, split) !== modelId || !MARKER_ID.test(id)) return undefined;
  return id;
}

module.exports = { Continuations, MARKER_MIME, decodeMarker };
