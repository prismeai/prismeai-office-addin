import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadPane, settle } from './harness.mjs';

// A deck the way templates build them: a title and a body placeholder, the
// footer placeholders, and now and then a picture or a chart (no text frame).
function deck(count, { pictures = true } = {}) {
  const slides = [];
  for (let n = 1; n <= count; n++) {
    const shapes = [
      { id: '2', name: 'Title 1', type: 'Placeholder', kind: 'Title', text: `SLIDE-${String(n).padStart(3, '0')} Title` },
      { id: '3', name: 'Content 2', type: 'Placeholder', kind: 'Body', text: `Body of slide ${n}` },
      { id: '4', name: 'Slide Number 3', type: 'Placeholder', kind: 'SlideNumber', text: String(n) },
    ];
    if (pictures && n % 3 === 0) shapes.push({ id: '5', name: 'Picture 4', type: 'Image', text: null });
    if (pictures && n % 4 === 0) {
      shapes.push({ id: '6', name: 'Chart 5', type: 'Placeholder', kind: 'Content', contained: 'Chart', text: null });
    }
    if (n % 5 === 0) shapes.push({ id: '7', name: 'TextBox 6', type: 'TextBox', text: `Note on slide ${n}` });
    slides.push({ shapes });
  }
  return { slides };
}

const powerpoint = (model, extra = {}) =>
  loadPane({ host: 'PowerPoint', powerpoint: model, url: 'https://files.test/Deck.pptx', ...extra });

async function readAll(pane, first) {
  const pages = [];
  for (let target = first, guard = 0; guard < 500; guard++) {
    const page = await pane.exec('office_read', target ? { target } : {});
    pages.push(page);
    if (!page.next) return pages;
    target = page.next;
  }
  throw new Error('office_read never reached the end');
}

const textOf = (model) =>
  model.slides.map((slide) => slide.shapes.filter((shape) => shape.text !== null).map((shape) => shape.text));

test('reading page after page returns every slide once, with the text of each text shape', async () => {
  const model = deck(150);
  const pane = powerpoint(model);
  const pages = await readAll(pane);
  assert.ok(pages.length > 1);
  const slides = pages.flatMap((page) => page.slides);
  assert.deepEqual(slides.map((slide) => slide.slide), model.slides.map((_, i) => i + 1));
  assert.deepEqual(slides.map((slide) => slide.shapes.map((shape) => shape.text)), textOf(model));
  assert.equal(pages[0].slideCount, 150);
});

test('each shape comes with the id the write tool takes', async () => {
  const model = deck(10);
  const pane = powerpoint(model);
  const page = await pane.exec('office_read', { target: 'slide:5' });
  assert.deepEqual(page.slides[0].shapes.map((shape) => [shape.id, shape.kind || null]), [
    ['2', 'Title'],
    ['3', 'Body'],
    ['4', 'SlideNumber'],
    ['7', null],
  ]);
  const body = page.slides[0].shapes[1];
  await pane.exec('powerpoint_write_slide', { slide: 5, shape: body.id, text: 'Rewritten' });
  assert.equal(model.slides[4].shapes[1].text, 'Rewritten');
});

test('a read is bounded in slides and in size, and pulls the text of those slides only', async () => {
  const pane = powerpoint(deck(150));
  const page = await pane.exec('office_read', { target: 'slide:60' });
  assert.ok(page.slides.length <= pane.get('READ_SLIDES'));
  assert.ok(JSON.stringify(page.slides).length <= pane.get('MAX_OUTPUT_CHARS') + page.slides.length);
  assert.equal(page.slides[0].slide, 60);
  assert.ok(pane.stats.textsSent <= pane.get('READ_SLIDES') * 4, `${pane.stats.textsSent} texts pulled`);
  assert.ok(pane.stats.syncs <= 5, `${pane.stats.syncs} round trips`);
});

test('an explicit span is returned exactly, and a span past the end stops at the last slide', async () => {
  const pane = powerpoint(deck(12));
  const exact = await pane.exec('office_read', { target: 'slide:3-5' });
  assert.deepEqual(exact.slides.map((slide) => slide.slide), [3, 4, 5]);
  assert.equal(exact.next, undefined);
  const tail = await pane.exec('office_read', { target: 'slide:10-40' });
  assert.deepEqual(tail.slides.map((slide) => slide.slide), [10, 11, 12]);
  assert.equal(tail.next, undefined);
  await assert.rejects(pane.exec('office_read', { target: 'slide:13' }), /the presentation has 12 slides/);
});

test('pictures and charts on the slides do not hide the text around them', async () => {
  // The host cannot say what the placeholders hold: every one is tried, and
  // each batch holding a shape without text is rejected as a whole.
  const model = deck(24);
  const pane = powerpoint(model, { options: { noPlaceholderInfo: true }, debug: true });
  const reply = await pane.call('office_read', { target: 'slide:1-20' });
  const expected = textOf(model).slice(0, 20);
  assert.deepEqual(reply.result.slides.map((slide) => slide.shapes.map((shape) => shape.text)), expected);
  assert.ok(reply.result.diag.rejected > 0, 'batches were rejected and split');
});

test('when the host describes the placeholders, shapes without text are never asked for', async () => {
  const pane = powerpoint(deck(20), { debug: true });
  const reply = await pane.call('office_read', { target: 'slide:1-20' });
  assert.equal(reply.result.diag.rejected, undefined);
  assert.equal(reply.result.diag.batches, 1);
});

test('a read that runs out of time stops at the first slide it could not read, and resumes there', async () => {
  const model = deck(20);
  const pane = powerpoint(model, { options: { noPlaceholderInfo: true }, msPerSync: 3000, debug: true });
  const reply = await pane.call('office_read', { target: 'slide:1-20' });
  assert.ok(reply.result, 'a result, not an error');
  assert.ok(reply.result.diag.skipped > 0, 'the shapes left unread are counted');
  const got = reply.result.slides;
  assert.ok(got.length >= 1 && got.length < 20);
  // No slide is handed back with part of its text missing, unless flagged.
  const expected = textOf(model);
  got.forEach((slide) => {
    if (!slide.incomplete) assert.deepEqual(slide.shapes.map((shape) => shape.text), expected[slide.slide - 1]);
  });
  assert.equal(reply.result.next, `slide:${got.length + 1}-20`);
});

test('a search never reports the deck as fully searched when some text was not reached', async () => {
  const model = deck(60);
  model.slides[55].shapes[1].text = 'the TARGET';
  const pane = powerpoint(model, { options: { noPlaceholderInfo: true }, msPerSync: 1000 });
  const hits = [];
  let next;
  let calls = 0;
  do {
    const found = await pane.exec('office_search', next ? { query: 'target', in: next } : { query: 'target' });
    hits.push(...found.results.map((hit) => hit.at));
    assert.equal(Boolean(found.partial), Boolean(found.next));
    next = found.next;
    calls++;
  } while (next && calls < 200);
  assert.equal(next, undefined, 'the search ends');
  assert.ok(calls > 1, 'the slow host needed several calls');
  assert.deepEqual(hits, ['slide:56']);
});

test('a shape that cannot be read is named in the debug figures', async () => {
  const model = deck(6);
  const pane = powerpoint(model, { options: { noPlaceholderInfo: true }, debug: true });
  const reply = await pane.call('office_read', { target: 'slide:1-6' });
  assert.equal(reply.result.diag.undescribed, true);
  assert.deepEqual(
    reply.result.diag.failed.map((entry) => [entry.slide, entry.id, entry.error]),
    // The chart placeholder of slide 4. Pictures are never asked for.
    [[4, '6', 'InvalidArgument']]
  );
});

test('a slide PowerPoint has not finished loading is flagged, never shown as empty', async () => {
  const model = deck(12, { pictures: false });
  const options = { unloaded: [5, 6] };
  const pane = powerpoint(model, { options });
  const page = await pane.exec('office_read', { target: 'slide:1-12' });
  assert.deepEqual(page.slides.filter((slide) => slide.incomplete).map((slide) => slide.slide), [5, 6]);
  const expected = textOf(model);
  page.slides.filter((slide) => !slide.incomplete).forEach((slide) => {
    assert.deepEqual(slide.shapes.map((shape) => shape.text), expected[slide.slide - 1]);
  });

  const outline = await pane.exec('office_outline', {});
  assert.deepEqual(outline.titles.filter((entry) => entry.incomplete).map((entry) => entry.at), ['slide:5', 'slide:6']);

  model.slides[4].shapes[1].text = 'hidden TARGET';
  const blind = await pane.exec('office_search', { query: 'target' });
  assert.equal(blind.matches, 0);
  assert.deepEqual(blind.incomplete, ['slide:5', 'slide:6'], 'the agent is told where the search could not look');

  // Once loaded, the same calls are complete.
  options.unloaded = [];
  const again = await pane.exec('office_read', { target: 'slide:5-6' });
  assert.equal(again.slides.some((slide) => slide.incomplete), false);
  const found = await pane.exec('office_search', { query: 'target' });
  assert.deepEqual(found.results.map((hit) => hit.at), ['slide:5']);
  assert.equal(found.incomplete, undefined);
});

// A call that waits on the pane's timers (the pauses between two reads).
async function pumped(pane, running) {
  let settled = false;
  running.finally(() => {
    settled = true;
  });
  for (let guard = 0; !settled && guard < 100; guard++) {
    await settle();
    pane.elapse(pane.get('LOAD_RETRY_MS'));
  }
  return running;
}

test('far slides PowerPoint on the web has not loaded are brought on screen, read, and the user is put back', async () => {
  const model = deck(150, { pictures: false });
  const pane = powerpoint(model, { options: { unlisted: [141, 142, 143, 144, 145, 146, 147, 148, 149, 150], onScreen: 3 }, debug: true });
  const reply = await pumped(pane, pane.call('office_read', { target: 'slide:141-150' }));
  assert.deepEqual(reply.result.slides.map((slide) => slide.shapes.map((shape) => shape.text)), textOf(model).slice(140));
  assert.equal(reply.result.slides.some((slide) => slide.incomplete), false);
  assert.deepEqual(pane.stats.selections, [[141], [3]], 'shown, then the user\'s slide again');
  assert.deepEqual(pane.onScreen(), [3]);
  assert.equal(reply.result.diag.rereads, 1);
});

test('the outline loads the far slides the same way', async () => {
  const model = deck(150, { pictures: false });
  const unlisted = Array.from({ length: 50 }, (_, i) => 101 + i);
  const pane = powerpoint(model, { options: { unlisted, onScreen: 1 } });
  const outline = await pumped(pane, pane.exec('office_outline', { from: 'slide:101' }));
  assert.equal(outline.titles.length, 50);
  assert.equal(outline.titles.some((entry) => entry.incomplete), false);
  assert.equal(outline.titles[49].title, 'SLIDE-150 Title');
  assert.deepEqual(pane.onScreen(), [1]);
});

test('on the web, a blank slide is shown once and read again twice at most', async () => {
  const model = deck(6, { pictures: false });
  model.slides[3].shapes = [];
  const pane = powerpoint(model, { options: { onScreen: 2 }, debug: true });
  const reply = await pumped(pane, pane.call('office_read', { target: 'slide:1-6' }));
  assert.equal(reply.result.slides.length, 6, 'no slide lost');
  assert.deepEqual(reply.result.slides.filter((slide) => slide.incomplete).map((slide) => slide.slide), [4]);
  assert.equal(reply.result.diag.rereads, 2);
  assert.deepEqual(pane.stats.selections, [[4], [2]]);
});

test('a reading again that runs out of time keeps the fuller first read', async () => {
  const model = deck(10, { pictures: false });
  const pane = powerpoint(model, { options: { unlisted: [8, 9], onScreen: 1 }, msPerSync: 2000, debug: true });
  const reply = await pumped(pane, pane.call('office_read', { target: 'slide:1-10' }));
  assert.equal(reply.result.diag.rereads, 1, 'read again once, out of time');
  assert.equal(reply.result.slides.length, 10);
  assert.deepEqual(reply.result.slides.filter((slide) => slide.incomplete).map((slide) => slide.slide), [8, 9]);
  const expected = textOf(model);
  reply.result.slides.filter((slide) => !slide.incomplete).forEach((slide) => {
    assert.deepEqual(slide.shapes.map((shape) => shape.text), expected[slide.slide - 1]);
  });
  assert.deepEqual(pane.onScreen(), [1]);
});

test('a slide is not brought on screen when the read has no time left for it', async () => {
  const pane = powerpoint(deck(10, { pictures: false }), { options: { unlisted: [10], onScreen: 1 }, msPerSync: 3000 });
  const page = await pumped(pane, pane.exec('office_read', { target: 'slide:1-10' }));
  assert.equal(page.slides.length, 10);
  assert.deepEqual(pane.stats.selections, []);
});

test('outside PowerPoint on the web, a slide without shapes is blank, not loading', async () => {
  const model = deck(5, { pictures: false });
  model.slides[1].shapes = [];
  const reply = await powerpoint(model, { platform: 'PC', debug: true }).call('office_read', { target: 'slide:1-5' });
  assert.deepEqual(reply.result.slides[1], { slide: 2, shapes: [] });
  assert.equal(reply.result.slides.some((slide) => slide.incomplete), false);
  assert.equal(reply.result.diag.shown, undefined, 'no slide is brought on screen');
  assert.equal(reply.result.diag.showFailed, undefined);
});

test('pictures and charts never make a slide look incomplete', async () => {
  const withInfo = await powerpoint(deck(24)).exec('office_read', { target: 'slide:1-20' });
  assert.equal(withInfo.slides.some((slide) => slide.incomplete), false);
  // When the host cannot describe the placeholders, a refusal is ambiguous: not flagged.
  const blind = await powerpoint(deck(24), { options: { noPlaceholderInfo: true } }).exec('office_read', { target: 'slide:1-20' });
  assert.equal(blind.slides.some((slide) => slide.incomplete), false);
});

test('the outline gives the number of slides and the title of each', async () => {
  const model = deck(8);
  // A slide whose title placeholder comes after its body, and one with no title at all.
  model.slides[1].shapes.reverse();
  model.slides[2].shapes = [{ id: '9', name: 'TextBox 1', type: 'TextBox', text: 'Free text only' }];
  model.slides[3].shapes = [{ id: '5', name: 'Picture 4', type: 'Image', text: null }];
  const pane = powerpoint(model);
  const outline = await pane.exec('office_outline', {});
  assert.equal(outline.slides, 8);
  assert.deepEqual(outline.titles.slice(0, 4), [
    { at: 'slide:1', title: 'SLIDE-001 Title' },
    { at: 'slide:2', title: 'SLIDE-002 Title' },
    { at: 'slide:3', title: 'Free text only' },
    { at: 'slide:4', title: '' },
  ]);
});

test('the outline of a long deck reads one text per slide and is listed in pages', async () => {
  const pane = powerpoint(deck(250));
  const first = await pane.exec('office_outline', {});
  assert.equal(first.titles.length, pane.get('OUTLINE_SLIDES'));
  assert.equal(pane.stats.textsSent, pane.get('OUTLINE_SLIDES'));
  const titles = [...first.titles];
  for (let from = first.next, guard = 0; from && guard < 10; guard++) {
    const page = await pane.exec('office_outline', { from });
    titles.push(...page.titles);
    from = page.next;
  }
  assert.deepEqual(titles.map((entry) => entry.at), Array.from({ length: 250 }, (_, i) => `slide:${i + 1}`));
  assert.equal(titles[249].title, 'SLIDE-250 Title');
});

test('search gives the slide and the shape of each hit, far into the deck', async () => {
  const model = deck(150);
  model.slides[119].shapes[1].text = 'Our churn rate dropped to 2%.';
  const pane = powerpoint(model);
  const found = await pane.exec('office_search', { query: 'CHURN' });
  assert.equal(found.matches, 1);
  assert.deepEqual(found.results, [{ at: 'slide:120', shape: '3', name: 'Content 2', text: 'Our churn rate dropped to 2%.' }]);
  await pane.exec('powerpoint_write_slide', { slide: 120, shape: found.results[0].shape, text: 'Updated' });
  assert.equal(model.slides[119].shapes[1].text, 'Updated');
});

test('search can be narrowed to a span of slides and caps what it returns', async () => {
  const pane = powerpoint(deck(100));
  const scoped = await pane.exec('office_search', { query: 'note on slide', in: 'slide:10-30' });
  assert.deepEqual(scoped.results.map((hit) => hit.at), ['slide:10', 'slide:15', 'slide:20', 'slide:25', 'slide:30']);
  const capped = await pane.exec('office_search', { query: 'body of slide', limit: 3 });
  assert.equal(capped.results.length, 3);
  assert.equal(capped.matches, 100);
});

test('a search that runs out of time says where it stopped, and resuming there finds the rest', async () => {
  const model = deck(400, { pictures: false });
  model.slides[10].shapes[1].text = 'first TARGET';
  model.slides[380].shapes[1].text = 'second TARGET';
  const pane = powerpoint(model, { msPerSync: 1500 });
  const first = await pane.exec('office_search', { query: 'target' });
  assert.equal(first.partial, true);
  assert.deepEqual(first.results.map((hit) => hit.at), ['slide:11']);
  const hits = [...first.results];
  let next = first.next;
  for (let guard = 0; next && guard < 50; guard++) {
    const more = await pane.exec('office_search', { query: 'target', in: next });
    hits.push(...more.results);
    next = more.next;
  }
  assert.equal(next, undefined, 'the search ends');
  assert.deepEqual(hits.map((hit) => hit.at), ['slide:11', 'slide:381']);
});

test('writing to a slide checks the slide and the shape before changing anything', async () => {
  const model = deck(6);
  const before = JSON.stringify(model);
  const pane = powerpoint(model);
  await assert.rejects(pane.exec('powerpoint_write_slide', { slide: 7, shape: '2', text: 'x' }), /slide 7 does not exist \(the presentation has 6 slides\)/);
  await assert.rejects(pane.exec('powerpoint_write_slide', { slide: 0, shape: '2', text: 'x' }), /does not exist/);
  await assert.rejects(pane.exec('powerpoint_write_slide', { slide: 3, shape: '99', text: 'x' }), /shape 99 not found on slide 3/);
  await assert.rejects(pane.exec('powerpoint_write_slide', { slide: 3, shape: '5', text: 'x' }), /holds no text/);
  await assert.rejects(pane.exec('powerpoint_write_slide', { slide: 3, shape: '2' }), /text is required/);
  await assert.rejects(pane.exec('powerpoint_write_slide', { shape: '2', text: 'x' }), /slide is required — ask the user for the slide number/);
  assert.equal(JSON.stringify(model), before);
});

test('a rewritten shape is set to fit its text, and the write holds when the host cannot do that', async () => {
  const fits = deck(3);
  const done = await powerpoint(fits).exec('powerpoint_write_slide', { slide: 2, shape: '2', text: 'A much longer title than before' });
  assert.deepEqual(done, { ok: true, slide: 2, shape: '2', fitted: true });
  assert.equal(fits.slides[1].shapes[0].autoSize, 'AutoSizeTextToFitShape');

  const plainHost = deck(3);
  const kept = await powerpoint(plainHost, { options: { noAutoSize: true } }).exec('powerpoint_write_slide', { slide: 2, shape: '2', text: 'New title' });
  assert.equal(kept.fitted, false);
  assert.equal(plainHost.slides[1].shapes[0].text, 'New title', 'the text is written all the same');
});

test('without a shape, the text goes into a new text box on the slide', async () => {
  const model = deck(2);
  const done = await powerpoint(model).exec('powerpoint_write_slide', { slide: 1, text: 'A side note' });
  assert.equal(done.added, true);
  assert.equal(model.slides[0].shapes.at(-1).text, 'A side note');
  assert.equal(model.slides[0].shapes.at(-1).id, done.shape);
});

test('a slide is deleted or moved by its number, within the bounds of the deck', async () => {
  const model = deck(5);
  const title = (i) => model.slides[i].shapes[0].text.slice(0, 9);
  const pane = powerpoint(model);
  await pane.exec('powerpoint_move_slide', { slide: 5, position: 2 });
  assert.deepEqual([0, 1, 2, 3, 4].map(title), ['SLIDE-001', 'SLIDE-005', 'SLIDE-002', 'SLIDE-003', 'SLIDE-004']);
  const gone = await pane.exec('powerpoint_delete_slide', { slide: 2 });
  assert.deepEqual(gone, { ok: true, deleted: 2, slides: 4 });
  assert.deepEqual([0, 1, 2, 3].map(title), ['SLIDE-001', 'SLIDE-002', 'SLIDE-003', 'SLIDE-004']);
  await assert.rejects(pane.exec('powerpoint_delete_slide', { slide: 9 }), /slide 9 does not exist/);
  await assert.rejects(pane.exec('powerpoint_move_slide', { slide: 1, position: 9 }), /position must be between 1 and 4/);
  assert.equal(model.slides.length, 4);
});

test('the last slide of a deck cannot be deleted', async () => {
  const model = deck(1);
  await assert.rejects(powerpoint(model).exec('powerpoint_delete_slide', { slide: 1 }), /only slide of a presentation cannot be deleted/);
  assert.equal(model.slides.length, 1);
});

test('the context shared with the agent covers the first slides, flagged as partial on a long deck', async () => {
  const pane = powerpoint(deck(150));
  const doc = await pane.get('readDocument')();
  assert.ok(doc.content.startsWith('--- Slide 1 ---\nSLIDE-001 Title'));
  assert.equal(doc.more, true);
  assert.ok(!doc.content.includes('SLIDE-150'));
  const short = await powerpoint(deck(4)).get('readDocument')();
  assert.equal(short.more, false);
});
