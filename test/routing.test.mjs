// A user with Word, Excel and PowerPoint open has one pane per application,
// and the instance delivers every tool call of that user to every pane. Each
// pane must answer the calls of its own conversation and stay silent on the
// others: the relay keeps the first answer, whichever document it comes from.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadPane, settle } from './harness.mjs';

const word = () => loadPane({ host: 'Word', word: { paragraphs: [{ text: 'A Word document.' }] } });
const excel = () =>
  loadPane({ host: 'Excel', excel: { active: 'Ventes', sheets: [{ name: 'Ventes', rows: [['Région', 'Montant'], ['Nord', 10], ['Sud', 5]] }] } });

const conversation = (pane, conversationId, reason = 'created') =>
  pane.frame({ type: 'event', name: 'chat:conversation-changed', data: { state: { conversationId }, reason } });
const call = (pane, name, contextId, args = {}, correlationId = `c-${name}`) =>
  pane.frame({ type: 'tool-exec', correlationId, name, args, contextId });

test('with three applications open, only the pane of the conversation answers', async () => {
  const panes = { word: word(), excel: excel(), powerpoint: loadPane({ host: 'PowerPoint' }) };
  conversation(panes.word, 'ctx_word');
  conversation(panes.excel, 'ctx_excel');
  conversation(panes.powerpoint, 'ctx_powerpoint');

  for (const pane of Object.values(panes)) {
    call(pane, 'excel_compute', 'ctx_excel', { op: 'sum', value: 'Montant', group_by: 'Région' });
  }
  await settle();
  assert.deepEqual(panes.word.results(), []);
  assert.deepEqual(panes.powerpoint.results(), []);
  const [answer] = panes.excel.results();
  assert.equal(answer.correlationId, 'c-excel_compute');
  assert.deepEqual(answer.result.groups.map((group) => [group.key, group.value]), [['Nord', 10], ['Sud', 5]]);
});

test('a tool every application has is answered from the right document', async () => {
  const panes = { word: word(), excel: excel() };
  conversation(panes.word, 'ctx_word');
  conversation(panes.excel, 'ctx_excel');
  for (const pane of Object.values(panes)) call(pane, 'office_read', 'ctx_word');
  await settle();
  assert.deepEqual(panes.excel.results(), [], 'Excel must not answer with its sheet');
  assert.equal(panes.word.results()[0].result.paragraphs[0].text, 'A Word document.');
});

test('a pane with no conversation yet leaves the tools of other applications alone', async () => {
  const pane = word();
  conversation(pane, null, 'mount');
  call(pane, 'excel_compute', 'ctx_someone_else', { op: 'count' });
  await settle();
  assert.deepEqual(pane.results(), []);
});

test('a chat that never gives the id of its conversation does not silence the pane', async () => {
  // The chat reports a fresh conversation at start-up and nothing after the
  // first message: the pane must still answer the calls of its own tools.
  const pane = excel();
  conversation(pane, null, 'mount');
  call(pane, 'excel_compute', 'ctx_never_reported', { op: 'sum', value: 'Montant' });
  await settle();
  assert.equal(pane.results()[0].result.result, 15);
});

test('a pane follows its chat: first message, then a new conversation', async () => {
  const pane = word();
  conversation(pane, null, 'mount');
  conversation(pane, 'ctx_first');
  call(pane, 'office_read', 'ctx_first', {}, 'one');
  await settle();
  assert.deepEqual(pane.results().map((message) => message.correlationId), ['one']);

  // "New conversation": the chat is reloaded and starts fresh.
  pane.get('loadChat')();
  pane.frame({ type: 'ready', features: ['conversationId'] });
  conversation(pane, 'ctx_second');
  call(pane, 'office_read', 'ctx_first', {}, 'stale');
  await settle();
  assert.equal(pane.results().length, 1, 'the previous conversation is no longer this pane\'s');

  call(pane, 'office_read', 'ctx_second', {}, 'two');
  await settle();
  assert.deepEqual(pane.results().map((message) => message.correlationId), ['one', 'two']);
});

test('the conversation id is matched with or without its prefix', async () => {
  const pane = word();
  conversation(pane, 'abc-123');
  call(pane, 'office_read', 'ctx_abc-123');
  await settle();
  assert.equal(pane.results().length, 1);
});

test('the pane of the right conversation still reports a tool of another application', async () => {
  const pane = word();
  conversation(pane, 'ctx_word');
  call(pane, 'excel_compute', 'ctx_word', { op: 'count' });
  await settle();
  assert.match(pane.results()[0].error.message, /excel_compute is not available in Word/);
});

test('on an instance that does not report conversations, a pane leaves the tools of other applications alone', async () => {
  const panes = { word: word(), excel: excel() };
  for (const pane of Object.values(panes)) {
    pane.frame({ type: 'ready' }); // no features announced
    call(pane, 'excel_compute', 'ctx_any', { op: 'sum', value: 'Montant' });
  }
  await settle();
  assert.deepEqual(panes.word.results(), [], 'Word stays silent instead of failing the call');
  assert.equal(panes.excel.results()[0].result.result, 15);

  // A tool they share cannot be told apart there: both answer, as before.
  for (const pane of Object.values(panes)) call(pane, 'office_read', 'ctx_any', {}, 'shared');
  await settle();
  assert.equal(panes.word.results().length, 1);
  assert.equal(panes.excel.results().length, 2);
});

test('a chat that announces conversations but never reports one does not silence the pane', async () => {
  // Seen on a live instance: the feature is announced at start-up, and no
  // conversation event follows. The pane must keep answering its calls.
  const pane = excel();
  pane.frame({ type: 'ready', features: ['conversationId'] });
  call(pane, 'excel_compute', 'ctx_unknown_to_the_pane', { op: 'sum', value: 'Montant' });
  await settle();
  assert.equal(pane.results()[0].result.result, 15);
});

test('the debug figures say which conversation the pane believes it is in', async () => {
  const silent = loadPane({ host: 'Word', word: { paragraphs: [{ text: 'x' }] }, debug: true });
  call(silent, 'office_read', 'ctx_a');
  await settle();
  assert.equal(silent.results()[0].result.diag.conversation, 'not reported');

  const told = loadPane({ host: 'Word', word: { paragraphs: [{ text: 'x' }] }, debug: true });
  conversation(told, 'ctx_a');
  call(told, 'office_read', 'ctx_b', {}, 'other');
  call(told, 'office_read', 'ctx_a', {}, 'mine');
  await settle();
  assert.deepEqual(told.results().map((message) => message.correlationId), ['mine']);
  assert.equal(told.results()[0].result.diag.conversation, 'ctx_a');
  assert.equal(told.results()[0].result.diag.ignored, 1);
});

test('a call without a conversation id is answered, as before', async () => {
  const pane = word();
  conversation(pane, 'ctx_word');
  call(pane, 'office_read', undefined);
  await settle();
  assert.equal(pane.results().length, 1);
});

test('a pane opened outside Office leaves to the Office panes the calls it cannot tie to its conversation', async () => {
  const panes = { browser: loadPane({}), word: word() };
  for (const pane of Object.values(panes)) call(pane, 'office_read', 'ctx_word');
  await settle();
  assert.deepEqual(panes.browser.results(), [], 'no document: it must not beat the Word pane');
  assert.equal(panes.word.results().length, 1);

  // Its own conversation is still its own.
  conversation(panes.browser, 'ctx_browser');
  call(panes.browser, 'office_read', 'ctx_browser', {}, 'mine');
  await settle();
  assert.deepEqual(panes.browser.results().map((message) => message.correlationId), ['mine']);
});

test('messages that do not come from the chat frame are ignored', async () => {
  const pane = word();
  conversation(pane, 'ctx_word');
  pane.frame({ type: 'tool-exec', correlationId: 'x', name: 'office_read', args: {}, contextId: 'ctx_word' }, 'https://evil.test');
  await settle();
  assert.deepEqual(pane.results(), []);
});
