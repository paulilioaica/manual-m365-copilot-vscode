'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const { assemblePrompt, prepareTools, decodeToolCall } = require('../prompt');

test('caps the final request and preserves only complete recent pairs', () => {
  const newest = 'keep this question ' + 'n'.repeat(4400);
  const prompt = assemblePrompt([
    { role:'user', text: 'drop-me ' + 'a'.repeat(8000) },
    { role:'assistant', text: 'drop-reply' },
    { role:'user', text: 'retain-me' },
    { role:'assistant', text: 'retain-reply' },
    { role:'user', text: newest },
  ], 'x'.repeat(3900));
  assert.ok(prompt.length <= 12000);
  assert.ok(prompt.includes(newest));
  assert.ok(prompt.includes('retain-me') && prompt.includes('retain-reply'));
  assert.ok(!prompt.includes('drop-me'));
  assert.throws(() => assemblePrompt([{ role:'user', text: 'X'.repeat(11999) }], 'instructions'), /12,000/);
  assert.equal(assemblePrompt([{ role:'user', text:'exact' }]), 'exact');
});

test('tool calls can address only tools supplied by the host in this turn', () => {
  const provided = prepareTools([{ name:'lookup', description:'Get a result', inputSchema:{ type:'object' } }], false);
  assert.deepEqual([...provided.names], ['lookup']);
  assert.deepEqual(decodeToolCall('```m365_tool_call\n{"name":"lookup","input":{"query":"orchid"}}\n```', provided.names, false), {
    name: 'lookup', input: { query:'orchid' },
  });
  assert.throws(() => decodeToolCall('```m365_tool_call\n{"name":"unknown","input":{}}\n```', provided.names, false), /does not support/);
  assert.throws(() => decodeToolCall('A normal reply', provided.names, true), /required/);
  assert.equal(decodeToolCall('A normal reply', provided.names, false), null);
});
