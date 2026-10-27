// The path of a relayed tool call: the chat frame hands the pane a call, the
// pane runs it against the open document and posts back exactly one answer.
// A call left without an answer shows up to the user as "no add-in attached".
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadPane, settle } from './harness.mjs';

const doc = { paragraphs: [{ text: 'Intro', style: 'Heading1' }, { text: 'Hello world.' }] };

test('an answer carries the id of its call and either a result or an error, never both', async () => {
  const pane = loadPane({ host: 'Word', word: doc });
  const ok = await pane.call('office_read', {});
  assert.equal(ok.type, 'tool-result');
  assert.equal(ok.correlationId, 'call-1');
  assert.equal(ok.result.paragraphs[1].text, 'Hello world.');
  assert.equal(ok.error, undefined);
  const ko = await pane.call('word_replace_text', { search: 'absent', replace: 'x' });
  assert.equal(ko.correlationId, 'call-2');
  assert.equal(ko.result, undefined);
  assert.equal(ko.error.code, 1);
  assert.match(ko.error.message, /search not found in the document/);
  assert.equal(pane.sent.filter((message) => message.type === 'tool-result').length, 2, 'one answer per call');
});

test('a tool the pane does not know is answered with an error', async () => {
  const pane = loadPane({ host: 'Word', word: doc });
  const reply = await pane.call('office_delete_everything', {});
  assert.match(reply.error.message, /unknown tool office_delete_everything/);
});

test('a tool of another application is refused by name', async () => {
  const cases = {
    Word: ['excel_compute', 'powerpoint_delete_slide', 'outlook_draft_new'],
    Excel: ['word_replace_text', 'powerpoint_write_slide', 'outlook_draft_reply'],
    PowerPoint: ['word_add_comment', 'excel_write_range'],
    Outlook: ['word_insert_paragraph', 'excel_add_sheet', 'powerpoint_move_slide'],
  };
  for (const [host, tools] of Object.entries(cases)) {
    const pane = loadPane({ host, mail: { subject: 's', body: 'b', to: [] } });
    for (const tool of tools) {
      const reply = await pane.call(tool, {});
      assert.match(reply.error.message, new RegExp(`${tool} is not available in ${host}`));
    }
    assert.equal(pane.stats.syncs, 0, `${host}: nothing was attempted on the document`);
  }
});

test('outside Office, a call is answered with an error instead of being left hanging', async () => {
  const pane = loadPane({ host: null });
  const reply = await pane.call('office_read', {});
  assert.match(reply.error.message, /not running inside an Office application/);
});

test('every tool of the catalogue that reads has the same three verbs in each application', async () => {
  const panes = {
    Word: loadPane({ host: 'Word', word: doc }),
    Excel: loadPane({ host: 'Excel', excel: { active: 'S', sheets: [{ name: 'S', rows: [['Hello'], ['world']] }] } }),
    PowerPoint: loadPane({
      host: 'PowerPoint',
      powerpoint: { slides: [{ shapes: [{ id: '2', name: 'Title', type: 'Placeholder', kind: 'Title', text: 'Hello' }] }] },
    }),
    Outlook: loadPane({ host: 'Outlook', mail: { subject: 'Hello', body: 'Hello there', to: [] } }),
  };
  for (const [host, pane] of Object.entries(panes)) {
    for (const [tool, args] of [['office_outline', {}], ['office_read', {}], ['office_search', { query: 'hello' }]]) {
      const reply = await pane.call(tool, args);
      assert.ok(reply.result, `${tool} in ${host}: ${reply.error && reply.error.message}`);
      assert.equal(reply.result.host, host);
    }
    const found = await pane.call('office_search', { query: 'hello' });
    assert.equal(found.result.matches, 1, host);
  }
});

test('a search query is validated before the document is touched', async () => {
  const pane = loadPane({ host: 'Word', word: doc });
  for (const [query, expected] of [
    [undefined, /query is required/],
    ['   ', /query is required/],
    [42, /query is required/],
    ['two\nlines', /single line of at most 255 characters/],
    ['x'.repeat(256), /single line of at most 255 characters/],
  ]) {
    const reply = await pane.call('office_search', { query });
    assert.match(reply.error.message, expected);
  }
  assert.equal(pane.stats.syncs, 0);
});

test('the number of search results is bounded whatever the agent asks for', async () => {
  const many = { paragraphs: Array.from({ length: 500 }, (_, i) => ({ text: `needle ${i}` })) };
  const pane = loadPane({ host: 'Word', word: many });
  assert.equal((await pane.exec('office_search', { query: 'needle' })).results.length, pane.get('SEARCH_LIMIT'));
  assert.equal((await pane.exec('office_search', { query: 'needle', limit: 100000 })).results.length, pane.get('MAX_SEARCH_LIMIT'));
  assert.equal((await pane.exec('office_search', { query: 'needle', limit: -3 })).results.length, 1);
  assert.equal((await pane.exec('office_search', { query: 'needle', limit: 'many' })).results.length, pane.get('SEARCH_LIMIT'));
});

test('a result too large for a conversation is turned into an error that says what to do', async () => {
  // One row of 100 long cells: a read keeps at least one row, so the per-read
  // limit cannot catch it. The hard stop does.
  const row = Array.from({ length: 100 }, () => 'x'.repeat(600));
  const pane = loadPane({ host: 'Excel', excel: { active: 'S', sheets: [{ name: 'S', rows: [row] }] } });
  const reply = await pane.call('office_read', {});
  assert.equal(reply.result, undefined);
  assert.match(reply.error.message, /office_read returned \d+ characters \(max 40000\) — ask for a smaller part/);
});

test('no read tool can return more than the hard limit on a very large document', async () => {
  const limit = loadPane({ host: 'Word' }).get('MAX_RESULT_CHARS');
  const long = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit. '.repeat(40);
  const word = loadPane({
    host: 'Word',
    word: { paragraphs: Array.from({ length: 5000 }, (_, i) => ({ text: `${i} ${long}`, style: i % 3 ? undefined : 'Heading1' })) },
  });
  for (const [tool, args] of [
    ['office_outline', {}],
    ['office_read', {}],
    ['office_read', { target: 'p:1-5000' }],
    ['office_search', { query: 'lorem', limit: 50 }],
  ]) {
    const reply = await word.call(tool, args);
    assert.ok(reply.result, `${tool}: ${reply.error && reply.error.message}`);
    assert.ok(JSON.stringify(reply.result).length <= limit, `${tool} ${JSON.stringify(args)}`);
  }
});

test('the selection is returned as text, cut to the read limit, and empty when it is not text', async () => {
  const limit = loadPane({ host: 'Word' }).get('MAX_OUTPUT_CHARS');
  const some = await loadPane({ host: 'Word', selection: 'picked text' }).exec('office_get_selection', {});
  assert.deepEqual(some, { host: 'Word', text: 'picked text' });
  const huge = await loadPane({ host: 'Word', selection: 'x'.repeat(limit * 4) }).exec('office_get_selection', {});
  assert.equal(huge.text.length, limit);
  // A slide thumbnail or an image is selected: the host refuses the read.
  const none = await loadPane({ host: 'PowerPoint' }).exec('office_get_selection', {});
  assert.deepEqual(none, { host: 'PowerPoint', text: '' });
});

test('a long email is read in parts and searched by position', async () => {
  const body = Array.from({ length: 3000 }, (_, i) => `Line ${i} of the thread.`).join('\n') + '\nDEADLINE is Friday.';
  const pane = loadPane({
    host: 'Outlook',
    mail: { subject: 'Contract', from: { displayName: 'Ana', emailAddress: 'ana@example.test' }, to: [{ emailAddress: 'bo@example.test' }], body },
  });
  const outline = await pane.exec('office_outline', {});
  assert.deepEqual(outline, { host: 'Outlook', subject: 'Contract', from: 'Ana <ana@example.test>', to: ['bo@example.test'], characters: body.length });

  let text = '';
  let target;
  for (let guard = 0; guard < 100; guard++) {
    const part = await pane.exec('office_read', target ? { target } : {});
    assert.ok(part.body.length <= pane.get('MAX_OUTPUT_CHARS'));
    text += part.body;
    target = part.next;
    if (!target) break;
  }
  assert.equal(text, body, 'the parts put end to end are the whole message');

  const found = await pane.exec('office_search', { query: 'deadline' });
  assert.equal(found.matches, 1);
  const around = await pane.exec('office_read', { target: found.results[0].at });
  assert.ok(around.body.startsWith('DEADLINE is Friday.'));
});

test('reading an email when none is open is an error', async () => {
  const pane = loadPane({ host: 'Outlook' });
  const reply = await pane.call('office_read', {});
  assert.match(reply.error.message, /no email is open/);
});

test('calls run one at a time, in the order they arrived', async () => {
  const pane = loadPane({ host: 'Word', word: doc });
  const exclusive = pane.get('exclusive');
  const log = [];
  const gates = [];
  const task = (name) => () =>
    new Promise((resolve) => {
      log.push(`start ${name}`);
      gates.push(() => {
        log.push(`end ${name}`);
        resolve(name);
      });
    });
  const results = [exclusive(task('a')), exclusive(task('b')), exclusive(task('c'))];
  await settle();
  assert.deepEqual(log, ['start a'], 'b and c wait for a');
  gates[0]();
  await settle();
  assert.deepEqual(log, ['start a', 'end a', 'start b']);
  gates[1]();
  await settle();
  gates[2]();
  assert.deepEqual(await Promise.all(results), ['a', 'b', 'c']);
  assert.deepEqual(log, ['start a', 'end a', 'start b', 'end b', 'start c', 'end c']);
});

test('a call that fails does not block the next one', async () => {
  const pane = loadPane({ host: 'Word', word: doc });
  const exclusive = pane.get('exclusive');
  const failed = exclusive(() => Promise.reject(new Error('boom')));
  const next = exclusive(() => 'fine');
  await assert.rejects(failed, /boom/);
  assert.equal(await next, 'fine');
});

test('a call Office never answers is reported after the time limit, with the step it was on', async () => {
  const model = { slides: [{ shapes: [{ id: '2', name: 'Title', type: 'Placeholder', kind: 'Title', text: 'Hello' }] }] };
  const pane = loadPane({ host: 'PowerPoint', powerpoint: model });
  // The host stops answering.
  pane.get('PowerPoint').run = () => new Promise(() => {});
  let reply;
  pane.call('office_read', {}).then((message) => {
    reply = message;
  });
  await settle();
  assert.equal(reply, undefined, 'no answer yet');
  pane.elapse(pane.get('TOOL_TIMEOUT_MS'));
  await settle();
  assert.match(reply.error.message, /office_read did not complete in PowerPoint within 25 s \(while listing the shapes\)/);
});

test('a stuck call is given up on, so the following one still runs', async () => {
  const pane = loadPane({ host: 'Word', word: doc });
  const run = pane.get('Word').run;
  pane.get('Word').run = () => new Promise(() => {});
  const stuck = pane.call('office_read', {});
  await settle();
  pane.get('Word').run = run;
  pane.elapse(20000);
  let second;
  pane.call('office_read', {}).then((message) => {
    second = message;
  });
  await settle();
  assert.equal(second, undefined, 'the second call waits behind the stuck one');
  pane.elapse(5000); // the stuck call reaches its time limit
  assert.match((await stuck).error.message, /did not complete/);
  await settle();
  assert.equal(second.result.paragraphs[1].text, 'Hello world.');
});

test('the context refresh and a tool call never read the document at the same time', async () => {
  const pane = loadPane({ host: 'Word', word: doc });
  let active = 0;
  let overlap = false;
  const run = pane.get('Word').run;
  pane.get('Word').run = async (task) => {
    active++;
    overlap = overlap || active > 1;
    await settle();
    try {
      return await run(task);
    } finally {
      active--;
    }
  };
  pane.get('refreshContext')();
  pane.elapse(800); // the refresh is debounced
  const reply = await pane.call('office_read', {});
  await settle();
  assert.ok(reply.result);
  assert.equal(overlap, false);
});

test('text inserted into an email being read is refused with the tool to use instead', async () => {
  const pane = loadPane({ host: 'Outlook', mail: { subject: 's', body: 'b', to: [] } });
  const reply = await pane.call('office_insert_text', { text: 'Hi' });
  assert.match(reply.error.message, /open in read mode — use outlook_draft_reply/);
  const missing = await pane.call('office_insert_text', {});
  assert.match(missing.error.message, /text is required/);
});

test('a meeting is opened as an unsent form with its attendees and its slot', async () => {
  const pane = loadPane({ host: 'Outlook' });
  const done = await pane.exec('outlook_draft_meeting', {
    attendees: ['ana@example.test'],
    optional: 'bo@example.test',
    subject: 'Point projet',
    body: 'Ordre du jour : planning.',
    start: '2026-10-14T10:00',
    end: '2026-10-14T11:00',
    location: 'Salle 2',
  });
  assert.deepEqual(done, { ok: true, sent: false });
  assert.equal(pane.forms.length, 1);
  const form = pane.forms[0];
  assert.equal(form.kind, 'meeting');
  assert.deepEqual([...form.requiredAttendees], ['ana@example.test']);
  assert.deepEqual([...form.optionalAttendees], ['bo@example.test']);
  assert.equal(form.subject, 'Point projet');
  assert.equal(form.location, 'Salle 2');
  // The slot is the user's local time, as typed.
  assert.equal(form.start.getHours(), 10);
  assert.equal(form.end.getTime() - form.start.getTime(), 3600000);
});

test('a meeting without an end lasts thirty minutes, and without a start leaves the slot to Outlook', async () => {
  const pane = loadPane({ host: 'Outlook' });
  await pane.exec('outlook_draft_meeting', { subject: 'Call', start: '2026-10-14T15:30' });
  assert.equal(pane.forms[0].end.getTime() - pane.forms[0].start.getTime(), 1800000);
  await pane.exec('outlook_draft_meeting', { subject: 'Later' });
  assert.equal('start' in pane.forms[1], false);
  assert.equal('end' in pane.forms[1], false);
});

test('a meeting with an unusable slot is refused and no form is opened', async () => {
  const pane = loadPane({ host: 'Outlook' });
  const cases = [
    [{ start: 'demain 10h' }, /start must be a date and time such as "2026-10-14T10:00"/],
    [{ start: '2026-10-14T10:00', end: 'plus tard' }, /end must be a date and time/],
    [{ start: '2026-10-14T10:00', end: '2026-10-14T09:00' }, /end must be after start/],
    [{ end: '2026-10-14T09:00' }, /end needs a start/],
    [{ start: 1760436000000 }, /start must be a date and time/],
  ];
  for (const [args, expected] of cases) await assert.rejects(pane.exec('outlook_draft_meeting', args), expected);
  assert.equal(pane.forms.length, 0);
});

test('a new email is opened unsent, with its text shown as written', async () => {
  const pane = loadPane({ host: 'Outlook' });
  await pane.exec('outlook_draft_new', { to: 'ana@example.test', subject: 'Hello', body: 'Line 1\n<b>not bold</b> & co' });
  const form = pane.forms[0];
  assert.equal(form.kind, 'message');
  assert.deepEqual([...form.toRecipients], ['ana@example.test']);
  assert.equal(form.htmlBody, 'Line 1<br>&lt;b&gt;not bold&lt;/b&gt; &amp; co');
  await assert.rejects(pane.exec('outlook_draft_new', { to: 'ana@example.test' }), /body is required/);
});

test('the tools the pane runs are exactly the ones of the catalogue it was built for', () => {
  const pane = loadPane({ host: 'Word' });
  assert.deepEqual(Object.keys(pane.get('TOOLS')).sort(), [
    'excel_add_chart',
    'excel_add_sheet',
    'excel_compute',
    'excel_write_range',
    'office_get_selection',
    'office_insert_text',
    'office_outline',
    'office_read',
    'office_search',
    'outlook_draft_meeting',
    'outlook_draft_new',
    'outlook_draft_reply',
    'powerpoint_add_slide',
    'powerpoint_delete_slide',
    'powerpoint_move_slide',
    'powerpoint_write_slide',
    'word_add_comment',
    'word_insert_paragraph',
    'word_replace_text',
  ]);
});

test('a tool name that is a property of every object is still an unknown tool', async () => {
  const pane = loadPane({ host: 'Word' });
  for (const name of ['constructor', 'toString', '__proto__', 'hasOwnProperty']) {
    const reply = await pane.call(name, {});
    assert.match(reply.error.message, /unknown tool/, name);
  }
});

test('timings are only attached to results and errors in debug mode', async () => {
  const quiet = loadPane({ host: 'Word', word: doc });
  assert.equal((await quiet.call('office_read', {})).result.diag, undefined);
  assert.doesNotMatch((await quiet.call('word_replace_text', { search: 'absent', replace: 'x' })).error.message, /diag/);
  const debug = loadPane({ host: 'Word', word: doc, debug: true });
  assert.equal(typeof (await debug.call('office_read', {})).result.diag.ms, 'number');
  assert.match((await debug.call('word_replace_text', { search: 'absent', replace: 'x' })).error.message, /\| diag \{"pane":"Word","ms":\d+,"waited":\d+,"conversation":"not reported"\}/);
});
