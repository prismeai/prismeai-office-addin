// Single sign-on: the chat asks the pane for its bearer ('token-request'), and
// the pane signs in with Microsoft and trades the ID token at the instance.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadPane, settle } from './harness.mjs';

test('token requests that arrive during a sign-in wait for it instead of starting another', async () => {
  const pane = loadPane({ host: 'Word' });
  // A sign-in held open until the test ends it (the user is still in Microsoft's window).
  pane.get(`openSession = () => {
    globalThis.signIns = (globalThis.signIns || 0) + 1;
    return new Promise((resolve) => { globalThis.endSignIn = resolve; });
  }`);
  pane.frame({ type: 'token-request' });
  pane.frame({ type: 'token-request' });
  await settle();
  pane.get('endSignIn')({ token: 'bearer-1', expires: Number.MAX_SAFE_INTEGER, servedAt: 0 });
  await settle();
  assert.equal(pane.get('signIns'), 1, 'a second sign-in, outside the click, would be blocked on the web');
  assert.deepEqual(pane.sent.filter((message) => message.type === 'token').map((message) => message.token), ['bearer-1', 'bearer-1']);
});
