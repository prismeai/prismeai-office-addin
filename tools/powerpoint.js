// PowerPoint: outline, read, search — then the tools that change the deck.

// PowerPoint on the web answers slowly (a few hundred ms per round trip), so a
// read must stay within a handful of batches. One shape without a text frame
// rejects its whole batch with InvalidArgument, hence the filtering below.
const TEXT_SHAPES = ['TextBox', 'GeometricShape'];
const TITLE_KINDS = ['Title', 'CenterTitle'];
const CONTEXT_SLIDES = 12; // slides read for the ambient context
const READ_SLIDES = 20; // slides per read
const OUTLINE_SLIDES = 100; // titles per outline page
const SEARCH_SLIDES = 40; // slides per round of a search
const SHOW_SLIDE_MS = 4000; // time given to PowerPoint to bring a slide on screen
const LOAD_RETRIES = 4; // reads after showing a slide that was not loaded
const LOAD_RETRY_MS = 700;

const lazyHost = () => Office.context.platform === Office.PlatformType.OfficeOnline;

function shapeAt(ctx, ref) {
  return ctx.presentation.slides.getItemAt(ref.slide).shapes.getItem(ref.id);
}

// Every shape of slides [from, from + count), as { slide (0-based), id, name, type }.
async function listShapes(from, count) {
  return PowerPoint.run(async (ctx) => {
    const size = ctx.presentation.slides.getCount();
    await ctx.sync();
    const total = size.value;
    const end = Math.min(total, from + count);
    const collections = [];
    for (let i = from; i < end; i++) {
      const shapes = ctx.presentation.slides.getItemAt(i).shapes;
      shapes.load('items/id,items/name,items/type');
      collections.push(shapes);
    }
    await ctx.sync();
    const refs = [];
    collections.forEach((shapes, i) =>
      shapes.items.forEach((shape) => refs.push({ slide: from + i, id: shape.id, name: shape.name, type: shape.type }))
    );
    return { total, end, refs };
  });
}

// What each placeholder holds (`contained` is null while it is empty or holds
// text). Returns false when the host can't tell — every placeholder is then tried.
async function describePlaceholders(refs) {
  if (!refs.length) return true;
  try {
    const kinds = await PowerPoint.run(async (ctx) => {
      const formats = refs.map((ref) => {
        const format = shapeAt(ctx, ref).placeholderFormat;
        format.load('type,containedType');
        return format;
      });
      await ctx.sync();
      return formats.map((format) => ({ kind: format.type, contained: format.containedType }));
    });
    refs.forEach((ref, i) => Object.assign(ref, kinds[i]));
    return true;
  } catch (err) {
    console.warn('[prisme-office] placeholders:', err);
    return false;
  }
}

// Fills `ref.text` for each shape. A rejected batch is split and retried in a
// fresh context (the proxies of a failed batch can't be reused) until the
// shape without a text frame is alone.
async function readTexts(refs, deadline) {
  if (!refs.length) return;
  if (Date.now() > deadline) {
    // Out of time: not the same as a shape without text.
    refs.forEach((ref) => {
      ref.unread = true;
    });
    toolDiag.skipped = (toolDiag.skipped || 0) + refs.length;
    return;
  }
  toolDiag.batches = (toolDiag.batches || 0) + 1;
  try {
    const texts = await PowerPoint.run(async (ctx) => {
      const ranges = refs.map((ref) => {
        const range = shapeAt(ctx, ref).textFrame.textRange;
        range.load('text');
        return range;
      });
      await ctx.sync();
      return ranges.map((range) => range.text);
    });
    refs.forEach((ref, i) => {
      ref.text = (texts[i] || '').trim();
    });
  } catch (err) {
    toolDiag.rejected = (toolDiag.rejected || 0) + 1;
    if (refs.length === 1) {
      // No text frame — or a slide PowerPoint has not finished loading: on a
      // long deck just opened, the text of the far slides is refused for a
      // while, then reads normally.
      refs[0].refused = true;
      const failed = (toolDiag.failed = toolDiag.failed || []);
      if (failed.length < 12) {
        const [ref] = refs;
        failed.push({ slide: ref.slide + 1, id: ref.id, type: ref.type, kind: ref.kind, error: errorCode(err) });
      }
      return;
    }
    const mid = refs.length >> 1;
    await readTexts(refs.slice(0, mid), deadline);
    await readTexts(refs.slice(mid), deadline);
  }
}

// Slides [from, from + count) with the text of their shapes — or, with
// `titlesOnly`, of the one shape per slide that names it.
// Returns { total, end, slides: [{ slide (1-based), shapes: [{ id, name, kind?, text }] }] },
// `end` being the first slide (0-based) left to read.
async function readSlides(from, count, deadline, titlesOnly) {
  officeStep = 'listing the shapes';
  const { total, end, refs } = await listShapes(from, count);
  officeStep = 'identifying the placeholders';
  const described = await describePlaceholders(refs.filter((ref) => ref.type === 'Placeholder'));
  if (!described) toolDiag.undescribed = true;
  let candidates = refs.filter((ref) =>
    ref.type === 'Placeholder' ? !ref.contained || TEXT_SHAPES.includes(ref.contained) : TEXT_SHAPES.includes(ref.type)
  );
  if (titlesOnly) {
    const bySlide = new Map();
    candidates.forEach((ref) => {
      const chosen = bySlide.get(ref.slide);
      if (!chosen || (!TITLE_KINDS.includes(chosen.kind) && TITLE_KINDS.includes(ref.kind))) bySlide.set(ref.slide, ref);
    });
    candidates = [...bySlide.values()];
  }
  officeStep = 'reading the text of the shapes';
  await readTexts(candidates, deadline);
  officeStep = '';
  // Out of time before the last slide: stop at the first one whose text was
  // not reached, so that it is read by the next call instead of being shown as
  // if it were empty. At least one slide is always returned, to keep moving.
  const unread = candidates.filter((ref) => ref.unread).map((ref) => ref.slide);
  const stop = unread.length ? Math.max(Math.min(...unread), from + 1) : end;
  // Once the placeholders are described, only shapes that hold text are asked
  // for: a refusal then means the slide is not ready, not that it is empty.
  const refused = described ? candidates.filter((ref) => ref.refused).map((ref) => ref.slide) : [];
  const slides = [];
  const unlisted = [];
  for (let i = from; i < stop; i++) {
    const slide = {
      slide: i + 1,
      shapes: candidates
        .filter((ref) => ref.slide === i && typeof ref.text === 'string')
        .map((ref) => ({ id: ref.id, name: ref.name, kind: ref.kind || undefined, text: ref.text })),
    };
    // A slide with no shape at all: blank — or, far into a long deck, one that
    // PowerPoint on the web has not loaded yet (it then lists nothing for it).
    // Only the web loads slides that late: anywhere else such a slide is blank.
    const notLoaded = lazyHost() && !refs.some((ref) => ref.slide === i);
    if (notLoaded) unlisted.push(i + 1);
    if (unread.includes(i) || refused.includes(i) || notLoaded) slide.incomplete = true;
    slides.push(slide);
  }
  return { total, end: stop, slides, unlisted };
}

const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Brings a slide on screen, which makes PowerPoint on the web load it and its
// neighbours. Returns the ids of the slides the user had selected, or null
// when the host can't, or doesn't answer in time.
async function showSlide(number) {
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.5')) return null;
  officeStep = `opening slide ${number}`;
  const diag = toolDiag; // the call this belongs to, even if PowerPoint answers late
  let late = false;
  const shown = PowerPoint.run(async (ctx) => {
    const selected = ctx.presentation.getSelectedSlides();
    selected.load('items/id');
    const slide = ctx.presentation.slides.getItemAt(number - 1);
    slide.load('id');
    await ctx.sync();
    // The read gave up on this slide: the view must not jump after it.
    if (late) return null;
    const before = selected.items.map((item) => item.id);
    ctx.presentation.setSelectedSlides([slide.id]);
    await ctx.sync();
    if (late && before.length) {
      // Shown after the read gave up on it: nobody else will put the user back.
      ctx.presentation.setSelectedSlides(before);
      await ctx.sync();
      return null;
    }
    return before;
  }).catch((err) => {
    diag.showFailed = errorCode(err);
    return null;
  });
  const timeout = pause(SHOW_SLIDE_MS).then(() => {
    late = true;
    return null;
  });
  const before = await Promise.race([shown, timeout]);
  if (before) diag.shown = number;
  return before;
}

// readSlides, and when PowerPoint on the web has not loaded some of these
// slides: show the first one so it loads that part of the deck, then read
// again a few times, within the budget. The user's selection is put back.
async function readLoadedSlides(from, count, deadline, titlesOnly) {
  let read = await readSlides(from, count, deadline, titlesOnly);
  if (!read.unlisted.length || Date.now() + SHOW_SLIDE_MS > deadline) return read;
  const before = await showSlide(read.unlisted[0]);
  if (!before) return read;
  try {
    for (let attempt = 0; attempt < LOAD_RETRIES && read.unlisted.length && Date.now() + LOAD_RETRY_MS < deadline; attempt++) {
      await pause(LOAD_RETRY_MS);
      toolDiag.rereads = attempt + 1;
      const again = await readSlides(from, count, deadline, titlesOnly);
      // Out of time halfway: the first read covers more.
      if (again.slides.length < read.slides.length) break;
      const progressed = again.unlisted.length < read.unlisted.length;
      read = again;
      // Nothing more loads: what is left is blank.
      if (!progressed && attempt > 0) break;
    }
  } finally {
    if (before.length) {
      officeStep = 'selecting the slide back';
      await PowerPoint.run(async (ctx) => {
        ctx.presentation.setSelectedSlides(before);
        await ctx.sync();
      }).catch((err) => {
        toolDiag.restoreFailed = errorCode(err);
      });
    }
  }
  return read;
}

async function powerPointOutline({ from }) {
  const start = parseSpan(from, 'slide').from;
  const { total, end, slides } = await readLoadedSlides(start - 1, OUTLINE_SLIDES, Date.now() + READ_BUDGET_MS, true);
  return {
    slides: total,
    titles: slides.map((slide) => ({
      at: `slide:${slide.slide}`,
      title: slide.shapes.length ? slide.shapes[0].text : '',
      incomplete: slide.incomplete,
    })),
    next: end < total ? `slide:${end + 1}` : undefined,
  };
}

async function powerPointRead({ target }, maxChars = MAX_OUTPUT_CHARS) {
  const span = parseSpan(target, 'slide');
  const wanted = Math.min(span.to ? span.to - span.from + 1 : READ_SLIDES, READ_SLIDES);
  const { total, slides } = await readLoadedSlides(span.from - 1, wanted, Date.now() + READ_BUDGET_MS);
  if (span.from > total) throw new Error(`the presentation has ${total} slides`);
  const out = [];
  let chars = 0;
  for (const slide of slides) {
    const size = JSON.stringify(slide).length;
    if (out.length && chars + size > maxChars) break;
    out.push(slide);
    chars += size;
  }
  const last = span.from + out.length - 1;
  const stop = span.to ? Math.min(span.to, total) : total;
  return {
    target: `slide:${span.from}-${last}`,
    slideCount: total,
    slides: out,
    next: last < stop ? `slide:${last + 1}${span.to ? `-${span.to}` : ''}` : undefined,
  };
}

async function powerPointSearch({ query, scope, limit }) {
  const span = parseSpan(scope, 'slide');
  const needle = query.toLowerCase();
  const deadline = Date.now() + READ_BUDGET_MS;
  const results = [];
  const incomplete = []; // slides whose text could not all be read
  let matches = 0;
  let position = span.from - 1; // 0-based
  let stop = span.to || Infinity;
  while (position < stop && Date.now() < deadline) {
    const { total, end, slides } = await readSlides(position, Math.min(SEARCH_SLIDES, stop - position), deadline);
    stop = Math.min(stop, total);
    slides.forEach((slide) => {
      if (slide.incomplete && incomplete.length < MAX_SEARCH_LIMIT) incomplete.push(`slide:${slide.slide}`);
      slide.shapes.forEach((shape) => {
        const index = shape.text.toLowerCase().indexOf(needle);
        if (index < 0) return;
        matches++;
        if (results.length < limit) {
          results.push({ at: `slide:${slide.slide}`, shape: shape.id, name: shape.name, text: snippet(shape.text, index, needle.length) });
        }
      });
    });
    if (end <= position) break;
    position = end;
  }
  const complete = position >= stop;
  return {
    query,
    matches,
    results,
    // Slides the search could not look into: a match there would be missed.
    incomplete: incomplete.length ? incomplete : undefined,
    // Out of time before the end: the search can be resumed from there.
    partial: complete ? undefined : true,
    next: complete ? undefined : `slide:${position + 1}${span.to ? `-${span.to}` : ''}`,
  };
}

async function powerPointWriteSlide({ slide, shape, text }) {
  if (typeof text !== 'string') throw new Error('text is required');
  // The slide on screen can't be told reliably in PowerPoint on the web.
  if (!given(slide)) throw new Error('slide is required — ask the user for the slide number');
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
    throw new Error('this version of PowerPoint cannot edit slides');
  }
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    const count = slides.getCount();
    await ctx.sync();
    const index = Number(slide) - 1;
    if (!Number.isInteger(index) || index < 0 || index >= count.value) {
      throw new Error(`slide ${slide} does not exist (the presentation has ${count.value} slides)`);
    }
    const target = slides.getItemAt(index);
    if (shape) {
      target.shapes.getItem(String(shape)).textFrame.textRange.text = text;
      try {
        await ctx.sync();
      } catch {
        throw new Error(
          `shape ${shape} not found on slide ${index + 1}, or it holds no text — call office_read on that slide for the shape ids`
        );
      }
      // Keep the new text inside the shape: a longer title otherwise overflows.
      // Its own batch — a host that refuses it must not undo the write.
      let fitted = true;
      try {
        await PowerPoint.run(async (fit) => {
          shapeAt(fit, { slide: index, id: String(shape) }).textFrame.autoSizeSetting = 'AutoSizeTextToFitShape';
          await fit.sync();
        });
      } catch (err) {
        console.warn('[prisme-office] autosize:', err);
        fitted = false;
      }
      return { ok: true, slide: index + 1, shape: String(shape), fitted };
    }
    // No target shape → the text goes into a new text box.
    const box = target.shapes.addTextBox(text);
    box.load('id');
    await ctx.sync();
    return { ok: true, slide: index + 1, shape: box.id, added: true };
  });
}

async function powerPointAddSlide({ like, position, duplicate }) {
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.8')) {
    throw new Error('this version of PowerPoint cannot add slides from the add-in');
  }
  const target = await PowerPoint.run(async (ctx) => {
    officeStep = 'counting the slides';
    const slides = ctx.presentation.slides;
    const count = slides.getCount();
    await ctx.sync();
    const total = count.value;
    const at = given(position) ? Number(position) : total + 1;
    if (!Number.isInteger(at) || at < 1 || at > total + 1) {
      throw new Error(`position must be between 1 and ${total + 1} (the presentation has ${total} slides)`);
    }
    // Default model: the slide just before the insertion point.
    const model = given(like) ? Number(like) : Math.max(at - 1, 1);
    if (!Number.isInteger(model) || model < 1 || model > total) {
      throw new Error(`slide ${like} does not exist (the presentation has ${total} slides)`);
    }
    const source = slides.getItemAt(model - 1);
    if (duplicate) {
      officeStep = 'copying the slide';
      const copy = source.exportAsBase64();
      const before = at > 1 ? slides.getItemAt(at - 2) : null;
      if (before) before.load('id');
      await ctx.sync();
      // Without a target the copy goes first; otherwise right after `before`.
      const options = { formatting: 'KeepSourceFormatting' };
      if (before) options.targetSlideId = before.id;
      ctx.presentation.insertSlidesFromBase64(copy.value, options);
      await ctx.sync();
    } else {
      officeStep = 'adding the slide';
      source.layout.load('id');
      source.slideMaster.load('id');
      await ctx.sync();
      slides.add({ layoutId: source.layout.id, slideMasterId: source.slideMaster.id });
      await ctx.sync();
      if (at <= total) {
        // A new slide is appended: move it into place.
        officeStep = 'moving the slide';
        slides.getItemAt(total).moveTo(at - 1);
        await ctx.sync();
      }
    }
    return at;
  });
  // The shapes of the new slide, so their text can be written right away.
  officeStep = 'listing the new slide';
  const refs = await PowerPoint.run(async (ctx) => {
    const shapes = ctx.presentation.slides.getItemAt(target - 1).shapes;
    shapes.load('items/id,items/name,items/type');
    await ctx.sync();
    return shapes.items.map((shape) => ({ slide: target - 1, id: shape.id, name: shape.name, type: shape.type }));
  });
  await describePlaceholders(refs.filter((ref) => ref.type === 'Placeholder'));
  officeStep = '';
  return {
    ok: true,
    slide: target,
    shapes: refs.map((ref) => ({ id: ref.id, name: ref.name, kind: ref.kind || ref.type })),
  };
}

async function powerPointDeleteSlide({ slide }) {
  return PowerPoint.run(async (ctx) => {
    const count = ctx.presentation.slides.getCount();
    await ctx.sync();
    const number = Number(slide);
    if (!Number.isInteger(number) || number < 1 || number > count.value) {
      throw new Error(`slide ${slide} does not exist (the presentation has ${count.value} slides)`);
    }
    if (count.value === 1) throw new Error('the only slide of a presentation cannot be deleted');
    ctx.presentation.slides.getItemAt(number - 1).delete();
    await ctx.sync();
    return { ok: true, deleted: number, slides: count.value - 1 };
  });
}

async function powerPointMoveSlide({ slide, position }) {
  if (!Office.context.requirements.isSetSupported('PowerPointApi', '1.8')) {
    throw new Error('this version of PowerPoint cannot move slides from the add-in');
  }
  return PowerPoint.run(async (ctx) => {
    const count = ctx.presentation.slides.getCount();
    await ctx.sync();
    const from = Number(slide);
    const to = Number(position);
    for (const [label, value] of [['slide', from], ['position', to]]) {
      if (!Number.isInteger(value) || value < 1 || value > count.value) {
        throw new Error(`${label} must be between 1 and ${count.value}`);
      }
    }
    ctx.presentation.slides.getItemAt(from - 1).moveTo(to - 1);
    await ctx.sync();
    return { ok: true, slide: to };
  });
}
