import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadPane } from './harness.mjs';

// A report of `sections` sections: one heading then `perSection` paragraphs,
// each carrying its own number so a test can tell exactly what came back.
function report(sections, perSection, body = (n) => `Paragraph ${n} of the report.`) {
  const paragraphs = [];
  for (let s = 1; s <= sections; s++) {
    paragraphs.push({ text: `Section ${s}`, style: s % 5 === 0 ? 'Heading2' : 'Heading1' });
    for (let i = 0; i < perSection; i++) paragraphs.push({ text: body(paragraphs.length + 1) });
  }
  return { paragraphs };
}

const word = (doc, extra = {}) => loadPane({ host: 'Word', word: doc, ...extra });

// Follows `next` until the tool says there is nothing left.
async function readAll(pane, first) {
  const pages = [];
  for (let target = first, guard = 0; guard < 1000; guard++) {
    const page = await pane.exec('office_read', target ? { target } : {});
    pages.push(page);
    if (!page.next) return pages;
    target = page.next;
  }
  throw new Error('office_read never reached the end');
}

test('reading page after page returns every paragraph once, in order', async () => {
  const doc = report(40, 30);
  const pane = word(doc);
  const pages = await readAll(pane);
  assert.ok(pages.length > 1, 'a long document takes several reads');
  const numbers = pages.flatMap((page) => page.paragraphs.map((p) => p.p));
  assert.deepEqual(numbers, doc.paragraphs.map((_, i) => i + 1));
  const texts = pages.flatMap((page) => page.paragraphs.map((p) => p.text));
  assert.deepEqual(texts, doc.paragraphs.map((p) => p.text));
});

test('a read never exceeds the size limit, however long the document', async () => {
  const pane = word(report(60, 40, (n) => `Paragraph ${n}. ` + 'Lorem ipsum dolor sit amet. '.repeat(20)));
  const limit = pane.get('MAX_OUTPUT_CHARS');
  for (const page of await readAll(pane)) {
    const size = page.paragraphs.reduce((sum, p) => sum + p.text.length, 0);
    assert.ok(size <= limit, `${page.target} returned ${size} characters`);
  }
});

test('a read only pulls the requested slice out of Word', async () => {
  const pane = word(report(100, 99)); // 10,000 paragraphs
  await pane.exec('office_read', { target: 'p:6000-6019' });
  assert.equal(pane.stats.paragraphsSent, 21, 'the slice plus one paragraph to know whether more follows');
  assert.equal(pane.stats.syncs, 1);
});

test('an explicit span is returned exactly, with nothing to continue', async () => {
  const pane = word(report(10, 9));
  const page = await pane.exec('office_read', { target: 'p:12-14' });
  assert.deepEqual(page.paragraphs.map((p) => p.p), [12, 13, 14]);
  assert.equal(page.target, 'p:12-14');
  assert.equal(page.next, undefined);
});

test('a span wider than one read continues up to its own end, not to the end of the document', async () => {
  const doc = report(10, 99); // 1,000 paragraphs
  const pane = word(doc);
  const pages = await readAll(pane, 'p:101-600');
  const numbers = pages.flatMap((page) => page.paragraphs.map((p) => p.p));
  assert.equal(numbers[0], 101);
  assert.equal(numbers.at(-1), 600);
  assert.equal(numbers.length, 500);
  assert.match(pages[0].next, /^p:\d+-600$/);
});

test('reading past the end returns nothing rather than failing', async () => {
  const pane = word(report(2, 3));
  const page = await pane.exec('office_read', { target: 'p:500' });
  assert.deepEqual(page.paragraphs, []);
  assert.equal(page.next, undefined);
});

test('headings are flagged and empty paragraphs are skipped without shifting the numbering', async () => {
  const pane = word({
    paragraphs: [{ text: 'Title', style: 'Heading1' }, { text: '' }, { text: '   ' }, { text: 'Body' }],
  });
  const page = await pane.exec('office_read', {});
  assert.deepEqual(page.paragraphs, [
    { p: 1, text: 'Title', heading: 1 },
    { p: 4, text: 'Body' },
  ]);
});

test('one paragraph longer than the limit is cut and says so, and the read moves on', async () => {
  const limit = loadPane({ host: 'Word' }).get('MAX_OUTPUT_CHARS');
  const pane = word({ paragraphs: [{ text: 'x'.repeat(limit * 3) }, { text: 'after' }] });
  const first = await pane.exec('office_read', {});
  assert.equal(first.paragraphs.length, 1);
  assert.equal(first.paragraphs[0].text.length, limit);
  assert.equal(first.paragraphs[0].cut, true);
  assert.equal(first.next, 'p:2');
  const second = await pane.exec('office_read', { target: first.next });
  assert.deepEqual(second.paragraphs, [{ p: 2, text: 'after' }]);
});

test('a host that ignores the requested window still yields the right slice', async () => {
  const doc = report(10, 9);
  const pane = word(doc, { options: { ignoreWindow: true }, debug: true });
  const reply = await pane.call('office_read', { target: 'p:42-44' });
  assert.deepEqual(reply.result.paragraphs.map((p) => p.text), doc.paragraphs.slice(41, 44).map((p) => p.text));
  assert.equal(reply.result.diag.unwindowed, true, 'the fallback is visible in the diagnostics');
});

test('the outline gives the size of the document and where each heading is', async () => {
  const doc = report(3, 4);
  const pane = word(doc);
  const outline = await pane.exec('office_outline', {});
  assert.equal(outline.host, 'Word');
  assert.equal(outline.title, 'Report.docx');
  assert.equal(outline.paragraphs, 15);
  assert.equal(outline.headingCount, 3);
  assert.deepEqual(outline.headings, [
    { at: 'p:1', level: 1, text: 'Section 1' },
    { at: 'p:6', level: 1, text: 'Section 2' },
    { at: 'p:11', level: 1, text: 'Section 3' },
  ]);
  // A heading position is directly readable.
  const page = await pane.exec('office_read', { target: 'p:6-7' });
  assert.equal(page.paragraphs[0].text, 'Section 2');
});

test('the outline only pulls the text of the headings, not of the document', async () => {
  const pane = word(report(50, 199)); // 10,000 paragraphs, 50 headings
  await pane.exec('office_outline', {});
  assert.equal(pane.stats.paragraphsSent, 50);
});

test('a document with more headings than one page of outline is listed in several calls', async () => {
  const pane = word(report(400, 1));
  const max = pane.get('MAX_HEADINGS');
  const seen = [];
  let from;
  for (let guard = 0; guard < 10; guard++) {
    const outline = await pane.exec('office_outline', from ? { from } : {});
    assert.ok(outline.headings.length <= max);
    assert.equal(outline.headingCount, 400);
    seen.push(...outline.headings.map((heading) => heading.text));
    from = outline.next;
    if (!from) break;
  }
  assert.equal(seen.length, 400);
  assert.equal(new Set(seen).size, 400, 'no heading listed twice');
  assert.equal(seen[5 - 1], 'Section 5');
});

test('a document without headings has an empty outline, not an error', async () => {
  const pane = word({ paragraphs: [{ text: 'one' }, { text: 'two' }] });
  const outline = await pane.exec('office_outline', {});
  assert.equal(outline.paragraphs, 2);
  assert.deepEqual(outline.headings, []);
});

test('search finds a passage deep in the document and gives a position that reads back', async () => {
  const doc = report(60, 30);
  doc.paragraphs[1500].text = 'The penalty clause applies after REF-1200 is signed.';
  const pane = word(doc);
  const found = await pane.exec('office_search', { query: 'ref-1200' });
  assert.equal(found.matches, 1);
  assert.equal(found.results[0].at, 'p:1501');
  assert.match(found.results[0].text, /REF-1200/);
  const page = await pane.exec('office_read', { target: found.results[0].at + '-1501' });
  assert.equal(page.paragraphs[0].text, doc.paragraphs[1500].text);
});

test('search reports every match but returns only the requested number of them', async () => {
  const pane = word(report(20, 30, (n) => (n % 10 === 0 ? `Clause ${n}: confidential.` : `Clause ${n}.`)));
  const found = await pane.exec('office_search', { query: 'Confidential', limit: 5 });
  assert.equal(found.results.length, 5);
  assert.ok(found.matches > 5, 'the agent knows there is more than what it sees');
  assert.equal(found.partial, undefined);
});

test('search can be narrowed to a span', async () => {
  const pane = word(report(10, 9, (n) => `Item ${n} marker`));
  const found = await pane.exec('office_search', { query: 'marker', in: 'p:20-29', limit: 50 });
  assert.deepEqual(found.results.map((hit) => hit.at), ['p:20', 'p:22', 'p:23', 'p:24', 'p:25', 'p:26', 'p:27', 'p:28', 'p:29']);
});

test('search with no match is an empty answer, not an error', async () => {
  const pane = word(report(3, 3));
  const found = await pane.exec('office_search', { query: 'nowhere to be found' });
  assert.equal(found.matches, 0);
  assert.deepEqual(found.results, []);
});

test('a search that runs out of time says where it stopped, and resuming there finds the rest', async () => {
  const doc = report(100, 99); // 10,000 paragraphs
  doc.paragraphs[300].text = 'early TARGET';
  doc.paragraphs[9500].text = 'late TARGET';
  // A slow host: the time budget only allows a few round trips.
  const pane = word(doc, { msPerSync: 4000 });
  const first = await pane.exec('office_search', { query: 'target' });
  assert.equal(first.partial, true);
  assert.deepEqual(first.results.map((hit) => hit.at), ['p:301']);
  assert.match(first.next, /^p:\d+$/);

  const hits = [...first.results];
  let next = first.next;
  for (let guard = 0; next && guard < 20; guard++) {
    const more = await pane.exec('office_search', { query: 'target', in: next });
    hits.push(...more.results);
    next = more.next;
  }
  assert.equal(next, undefined, 'the search ends');
  assert.deepEqual(hits.map((hit) => hit.at), ['p:301', 'p:9501']);
});

test('a replacement aimed at one paragraph leaves the same sentence untouched elsewhere', async () => {
  const doc = {
    paragraphs: [
      { text: 'The supplier shall deliver within 30 days.' },
      { text: 'Unrelated.' },
      { text: 'The supplier shall deliver within 30 days.' },
    ],
  };
  const pane = word(doc);
  const done = await pane.exec('word_replace_text', { search: '30 days', replace: '45 days', paragraph: 3 });
  assert.equal(done.replaced, 1);
  assert.equal(doc.paragraphs[0].text, 'The supplier shall deliver within 30 days.');
  assert.equal(doc.paragraphs[2].text, 'The supplier shall deliver within 45 days.');
});

test('without a paragraph the first occurrence is replaced, or all of them on request', async () => {
  const doc = { paragraphs: [{ text: 'ACME and ACME' }, { text: 'ACME again' }] };
  const pane = word(doc);
  const one = await pane.exec('word_replace_text', { search: 'ACME', replace: 'Contoso' });
  assert.deepEqual([one.replaced, one.occurrences], [1, 3]);
  assert.equal(doc.paragraphs[0].text, 'Contoso and ACME');
  const all = await pane.exec('word_replace_text', { search: 'ACME', replace: 'Contoso', all: true });
  assert.equal(all.replaced, 2);
  assert.deepEqual(doc.paragraphs.map((p) => p.text), ['Contoso and Contoso', 'Contoso again']);
});

test('a replacement that cannot be located changes nothing and explains why', async () => {
  const doc = { paragraphs: [{ text: 'Alpha' }, { text: 'Beta' }] };
  const pane = word(doc);
  await assert.rejects(pane.exec('word_replace_text', { search: 'Alpha', replace: 'x', paragraph: 2 }), /not found in paragraph 2/);
  await assert.rejects(pane.exec('word_replace_text', { search: 'alpha', replace: 'x' }), /match the text exactly, case included/);
  await assert.rejects(pane.exec('word_replace_text', { search: 'Alpha', replace: 'x', paragraph: 9 }), /past the end of the document/);
  await assert.rejects(pane.exec('word_replace_text', { search: 'Alpha', replace: 'x', paragraph: 0 }), /paragraph number, starting at 1/);
  await assert.rejects(pane.exec('word_replace_text', { search: 'two\nlines', replace: 'x' }), /single line of at most 255/);
  await assert.rejects(pane.exec('word_replace_text', { search: 'Alpha' }), /replace is required/);
  assert.deepEqual(doc.paragraphs.map((p) => p.text), ['Alpha', 'Beta']);
});

test('paragraphs are inserted after a numbered paragraph, in order, the first one as a heading', async () => {
  const doc = report(2, 2); // p1 heading, p2, p3, p4 heading, p5, p6
  const pane = word(doc);
  await pane.exec('word_insert_paragraph', { text: 'Risks\nFirst risk.\n\nSecond risk.', after_paragraph: 3, heading: 2 });
  assert.deepEqual(
    doc.paragraphs.slice(2, 7).map((p) => [p.text, p.style || 'Normal']),
    [
      ['Paragraph 3 of the report.', 'Normal'],
      ['Risks', 'Heading2'],
      ['First risk.', 'Normal'],
      ['Second risk.', 'Normal'],
      ['Section 2', 'Heading1'],
    ]
  );
});

test('without a position, paragraphs go to the end of the document', async () => {
  const doc = report(1, 1);
  const pane = word(doc);
  await pane.exec('word_insert_paragraph', { text: 'Conclusion', heading: 1 });
  assert.deepEqual(doc.paragraphs.at(-1), { text: 'Conclusion', style: 'Heading1' });
  await assert.rejects(pane.exec('word_insert_paragraph', { text: 'x', heading: 7 }), /heading must be 1, 2 or 3/);
});

test('a comment is attached to the passage and the text is left as it was', async () => {
  const doc = { paragraphs: [{ text: 'Payment is due soon.' }, { text: 'Payment is due soon.' }] };
  const pane = word(doc);
  await pane.exec('word_add_comment', { anchor: 'soon', comment: 'Give a date.', paragraph: 2 });
  assert.equal(doc.paragraphs[0].comments, undefined);
  assert.deepEqual(doc.paragraphs[1].comments, ['Give a date.']);
  assert.equal(doc.paragraphs[1].text, 'Payment is due soon.');
});

test('the context shared with the agent is the beginning of the document, flagged as partial', async () => {
  const pane = word(report(100, 99, (n) => `Paragraph ${n}. ` + 'Lorem ipsum. '.repeat(30)));
  const doc = await pane.get('readDocument')();
  assert.ok(doc.content.length <= pane.get('CONTEXT_CHARS'));
  assert.equal(doc.more, true);
  assert.ok(pane.stats.paragraphsSent <= pane.get('WORD_READ_MAX') + 1, 'the whole document is never pulled for the context');
  const note = pane.get('contextNote')(doc.more);
  for (const tool of ['office_outline', 'office_read', 'office_search']) assert.ok(note.includes(tool), tool);
  assert.ok(!pane.get('contextNote')(false).includes('office_outline'), 'no detour suggested when everything is shown');
});
