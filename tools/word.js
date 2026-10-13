// Word: outline, read, search — then the tools that change the document.

const WORD_READ_MAX = 200; // paragraphs per read
const WORD_SCAN = 400; // paragraphs per round trip while searching
const MAX_HEADINGS = 150;

// Paragraphs [skip, skip + top) of the body, as plain objects holding `fields`.
// Each window runs in its own context so only that slice crosses the wire.
async function wordWindow(skip, top, fields) {
  return Word.run(async (ctx) => {
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load({ select: fields.join(','), skip, top });
    await ctx.sync();
    let items = paragraphs.items;
    // A host that ignores the window hands back the whole collection.
    if (items.length > top) {
      toolDiag.unwindowed = true;
      items = items.slice(skip, skip + top);
    }
    return items.map((item) => Object.fromEntries(fields.map((field) => [field, item[field]])));
  });
}

// The paragraph numbered `number` (1-based), as a proxy usable in `ctx`.
async function wordParagraphAt(ctx, number, label) {
  const index = Number(number);
  if (!Number.isInteger(index) || index < 1) throw new Error(`${label} must be a paragraph number, starting at 1`);
  const paragraphs = ctx.document.body.paragraphs;
  paragraphs.load({ select: 'text', skip: index - 1, top: 1 });
  await ctx.sync();
  const item = paragraphs.items.length > 1 ? paragraphs.items[index - 1] : paragraphs.items[0];
  if (!item) throw new Error(`${label} ${index} is past the end of the document`);
  return item;
}

// 1-9 for a heading, 0 for anything else.
function headingLevel(paragraph) {
  const style = /^Heading([1-9])$/.exec(paragraph.styleBuiltIn || '');
  if (style) return +style[1];
  const level = Number(paragraph.outlineLevel);
  return level >= 1 && level <= 9 ? level : 0;
}

async function wordOutline({ from }) {
  const start = parseSpan(from, 'p').from;
  return Word.run(async (ctx) => {
    officeStep = 'listing the paragraphs';
    const paragraphs = ctx.document.body.paragraphs;
    paragraphs.load('items/outlineLevel,items/styleBuiltIn');
    await ctx.sync();
    const headings = [];
    paragraphs.items.forEach((item, i) => {
      const level = headingLevel(item);
      if (level) headings.push({ p: i + 1, level, item });
    });
    const first = headings.findIndex((heading) => heading.p >= start);
    const page = first < 0 ? [] : headings.slice(first, first + MAX_HEADINGS);
    officeStep = 'reading the headings';
    page.forEach((heading) => heading.item.load('text'));
    await ctx.sync();
    officeStep = '';
    const rest = first < 0 ? null : headings[first + MAX_HEADINGS];
    return {
      paragraphs: paragraphs.items.length,
      headingCount: headings.length,
      headings: page.map((heading) => ({
        at: `p:${heading.p}`,
        level: heading.level,
        text: heading.item.text.trim().slice(0, 200),
      })),
      next: rest ? `p:${rest.p}` : undefined,
    };
  });
}

async function wordRead({ target }, maxChars = MAX_OUTPUT_CHARS) {
  const span = parseSpan(target, 'p');
  const wanted = Math.min(span.to ? span.to - span.from + 1 : WORD_READ_MAX, WORD_READ_MAX);
  // One more than wanted, to know whether the document goes on.
  const page = await wordWindow(span.from - 1, wanted + 1, ['text', 'outlineLevel', 'styleBuiltIn']);
  const out = [];
  let chars = 0;
  let read = 0;
  for (const item of page.slice(0, wanted)) {
    const text = (item.text || '').trim();
    if (out.length && chars + text.length > maxChars) break;
    read++;
    if (!text) continue;
    const entry = { p: span.from + read - 1, text: text.slice(0, maxChars) };
    if (text.length > maxChars) entry.cut = true;
    const level = headingLevel(item);
    if (level) entry.heading = level;
    out.push(entry);
    chars += entry.text.length;
  }
  const last = span.from + read - 1;
  // A next position when the read stopped before the requested end, or — with
  // no end requested — when the document goes on.
  const more = span.to ? last < span.to && page.length > read : page.length > read;
  return {
    target: `p:${span.from}-${Math.max(last, span.from)}`,
    paragraphs: out,
    next: more ? `p:${last + 1}${span.to ? `-${span.to}` : ''}` : undefined,
  };
}

async function wordSearch({ query, scope, limit }) {
  const span = parseSpan(scope, 'p');
  const needle = query.toLowerCase();
  const deadline = Date.now() + READ_BUDGET_MS;
  const results = [];
  let matches = 0;
  let position = span.from;
  let ended = false;
  officeStep = 'searching the paragraphs';
  while (!ended && Date.now() < deadline) {
    const top = span.to ? Math.min(WORD_SCAN, span.to - position + 1) : WORD_SCAN;
    if (top < 1) break;
    const page = await wordWindow(position - 1, top, ['text']);
    page.forEach((item, i) => {
      const text = item.text || '';
      const index = text.toLowerCase().indexOf(needle);
      if (index < 0) return;
      matches++;
      if (results.length < limit) results.push({ at: `p:${position + i}`, text: snippet(text, index, needle.length) });
    });
    position += page.length;
    ended = page.length < top;
  }
  officeStep = '';
  const complete = ended || (span.to && position > span.to);
  return {
    query,
    matches,
    results,
    // Out of time before the end: the search can be resumed from there.
    partial: complete ? undefined : true,
    next: complete ? undefined : `p:${position}${span.to ? `-${span.to}` : ''}`,
  };
}

// `paragraph` narrows the search to that paragraph: exact on a long document,
// where the same sentence can come back many times.
async function wordFind(ctx, text, label, paragraph) {
  if (typeof text !== 'string' || !text.trim()) throw new Error(`${label} is required`);
  if (text.length > MAX_SEARCH || /[\r\n]/.test(text)) {
    throw new Error(`${label} must be a single line of at most ${MAX_SEARCH} characters`);
  }
  const numbered = given(paragraph);
  const scope = numbered ? await wordParagraphAt(ctx, paragraph, 'paragraph') : ctx.document.body;
  const found = scope.search(text, { matchCase: true });
  found.load('items');
  await ctx.sync();
  if (!found.items.length) {
    throw new Error(
      `${label} not found in ${numbered ? `paragraph ${paragraph}` : 'the document'} — it must match the text exactly, case included`
    );
  }
  return found.items;
}

async function wordReplaceText({ search, replace, all, paragraph }) {
  if (typeof replace !== 'string') throw new Error('replace is required');
  return Word.run(async (ctx) => {
    const found = await wordFind(ctx, search, 'search', paragraph);
    const targets = all ? found : found.slice(0, 1);
    targets.forEach((range) => range.insertText(replace, 'Replace'));
    await ctx.sync();
    return { ok: true, replaced: targets.length, occurrences: found.length };
  });
}

async function wordInsertParagraph({ text, after, after_paragraph: afterParagraph, heading }) {
  if (typeof text !== 'string' || !text.trim()) throw new Error('text is required');
  const level = given(heading) ? Number(heading) : 0;
  if (![0, 1, 2, 3].includes(level)) throw new Error('heading must be 1, 2 or 3');
  const lines = text.split(/\r?\n+/).filter((line) => line.trim());
  return Word.run(async (ctx) => {
    // After the paragraph holding the anchor, never in the middle of it.
    let previous = null;
    if (given(afterParagraph)) {
      previous = await wordParagraphAt(ctx, afterParagraph, 'after_paragraph');
    } else if (after) {
      previous = (await wordFind(ctx, after, 'after'))[0].paragraphs.getLast();
    }
    lines.forEach((line, i) => {
      const paragraph = previous
        ? previous.insertParagraph(line, 'After')
        : ctx.document.body.insertParagraph(line, 'End');
      // Set explicitly: a new paragraph otherwise inherits the style of its neighbour.
      paragraph.styleBuiltIn = level && i === 0 ? `Heading${level}` : 'Normal';
      previous = paragraph;
    });
    await ctx.sync();
    return { ok: true, paragraphs: lines.length };
  });
}

async function wordAddComment({ anchor, comment, paragraph }) {
  if (typeof comment !== 'string' || !comment.trim()) throw new Error('comment is required');
  if (!Office.context.requirements.isSetSupported('WordApi', '1.4')) {
    throw new Error('this version of Word cannot add comments from the add-in');
  }
  return Word.run(async (ctx) => {
    const found = await wordFind(ctx, anchor, 'anchor', paragraph);
    found[0].insertComment(comment);
    await ctx.sync();
    return { ok: true, occurrences: found.length };
  });
}
