// Positions are the contract between two tool calls: what a read or a search
// hands back ("p:120", "'Q1 2025'!A41:L900", "slide:40") is what the agent
// passes to the next one. A position that doesn't round-trip sends the agent to
// the wrong place in the customer's document.
import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadPane, plain } from './harness.mjs';

const pane = loadPane({ host: 'Word' });
const parseSpan = pane.get('parseSpan');
const splitTarget = pane.get('splitTarget');
const parseA1 = pane.get('parseA1');
const columnName = pane.get('columnName');
const columnIndex = pane.get('columnIndex');
const sheetRef = pane.get('sheetRef');
const snippet = pane.get('snippet');

test('a missing position means the start of the document', () => {
  for (const target of [undefined, null, '']) assert.deepEqual(plain(parseSpan(target, 'p')), { from: 1, to: null });
});

test('a position is a single place or an inclusive span', () => {
  assert.deepEqual(plain(parseSpan('p:12', 'p')), { from: 12, to: null });
  assert.deepEqual(plain(parseSpan(' slide:3-9 ', 'slide')), { from: 3, to: 9 });
  assert.deepEqual(plain(parseSpan('p:7-7', 'p')), { from: 7, to: 7 });
});

test('a malformed position is refused with the expected form, never guessed', () => {
  for (const target of ['12', 'p:0', 'p:9-3', 'slide:2', 'p:1-', 'p:a', 'p:1.5', 12]) {
    assert.throws(() => parseSpan(target, 'p'), /expected a position such as "p:12" or "p:12-40"/, String(target));
  }
});

test('column letters and indexes convert both ways, up to the last column of a sheet', () => {
  const cases = { 0: 'A', 25: 'Z', 26: 'AA', 27: 'AB', 51: 'AZ', 52: 'BA', 701: 'ZZ', 702: 'AAA', 16383: 'XFD' };
  for (const [index, name] of Object.entries(cases)) {
    assert.equal(columnName(+index), name);
    assert.equal(columnIndex(name), +index);
  }
  assert.equal(columnIndex('ab'), 27, 'lower case is accepted');
});

test('an Excel target is a sheet, a range, or both', () => {
  assert.deepEqual(plain(splitTarget('Ventes!A1:L50')), { sheet: 'Ventes', address: 'A1:L50' });
  assert.deepEqual(plain(splitTarget('A1:L50')), { address: 'A1:L50' });
  assert.deepEqual(plain(splitTarget('$B$2')), { address: '$B$2' });
  assert.deepEqual(plain(splitTarget('A:C')), { address: 'A:C' });
  assert.deepEqual(plain(splitTarget('3:10')), { address: '3:10' });
  assert.deepEqual(plain(splitTarget('Ventes')), { sheet: 'Ventes' });
  assert.deepEqual(plain(splitTarget('Annexe 1')), { sheet: 'Annexe 1' });
  assert.deepEqual(plain(splitTarget(undefined)), {});
  assert.deepEqual(plain(splitTarget('  ')), {});
});

test('a sheet name with spaces, quotes or an exclamation mark survives the round trip', () => {
  for (const name of ['Ventes', 'Annexe 1', "Budget d'équipe", 'Total!', 'Synthèse', '2025']) {
    const target = `${sheetRef(name)}!B2:C3`;
    assert.deepEqual(plain(splitTarget(target)), { sheet: name, address: 'B2:C3' }, target);
  }
});

test('a cell range is read as zero-based bounds', () => {
  assert.deepEqual(plain(parseA1('C2:D5')), { r0: 1, c0: 2, r1: 4, c1: 3 });
  assert.deepEqual(plain(parseA1('$AA$10')), { r0: 9, c0: 26, r1: 9, c1: 26 });
  assert.equal(parseA1('C:C'), null, 'a whole column is not a cell range');
});

test('a search extract keeps the match and marks what was cut around it', () => {
  const text = 'a'.repeat(200) + ' NEEDLE ' + 'b'.repeat(200);
  const extract = snippet(text, 201, 6);
  assert.ok(extract.includes('NEEDLE'));
  assert.ok(extract.startsWith('…') && extract.endsWith('…'));
  assert.ok(extract.length <= 6 + 160 + 2);
  assert.equal(snippet('short\n  text', 0, 5), 'short text', 'line breaks are flattened, nothing is marked as cut');
});
