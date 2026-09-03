// the parts of a header or footer the user is allowed to change.
//
// laymirror does not build a header. it finds the two or three words inside the
// template's own header that belong to whoever holds the file — a team code, a
// year, a title, a name — and swaps them, leaving every other byte alone.
//
// two ways to find them, in order of preference:
//
//   marked    the template wraps a placeholder in a zero-width character
//             (`<zwj>File Title<zwj>` in word). explicit, so it wins.
//   inferred  no marks anywhere, so every stretch of plain text between tabs
//             and word fields is offered instead.
//
// discovery always runs against the pristine template, so a field keeps its
// identity after its value has been replaced.

import { parseXml, serializeXml } from './xml.js';

/** zero-width characters word will happily hold and never draw. */
const ZERO_WIDTH = /[​‌‍⁠﻿]/;

const HEADER_OR_FOOTER = /\/(header|footer)\d*\.xml$/;

export interface Field {
  /** stable across value changes: the part, the paragraph, the slot. */
  key: string;
  /** what the template had there, which is also what the row is called. */
  label: string;
  part: string;
}

export type Values = Record<string, string>;

interface TextRun {
  node: Element;
  /** where this node's text starts in the paragraph's flattened text. */
  at: number;
  text: string;
}

interface Segment {
  from: number;
  to: number;
}

/** what interrupted the text at an offset. the " page " and " of " around
 *  `PAGE` and `NUMPAGES` read as plain text but are not the user's to edit, and
 *  touching a field boundary is what tells them apart from real content. */
type Break = { at: number; field: boolean };

/** properties, not content. `w:pPr` holds tab stops, and reading those as tabs
 *  would split a header where nothing is written; `mc:Fallback` is the vml copy
 *  of a text box already given in `mc:Choice`, so it doubles every field. */
const SKIP = new Set(['w:pPr', 'w:rPr', 'w:instrText', 'w:delText', 'mc:Fallback']);

function tagged(parent: Element, tag: string): Element[] {
  const found = parent.getElementsByTagName(tag);
  const out: Element[] = [];
  for (let i = 0; i < found.length; i++) out.push(found.item(i)!);
  return out;
}

const inFallback = (node: Element): boolean => {
  for (let cursor: Node | null = node.parentNode; cursor; cursor = cursor.parentNode) {
    if (cursor.nodeType === 1 && (cursor as Element).tagName === 'mc:Fallback') return true;
  }
  return false;
};

/** the `w:t` nodes this paragraph owns, in document order, each with a flattened
 *  text offset — plus the offsets where the text is interrupted.
 *
 *  a nested paragraph is skipped rather than inlined: a text box holds whole
 *  paragraphs inside the run that anchors it, and each is visited on its own.
 *
 *  a tab, a line break and a word field all interrupt — the two sides of a tab
 *  are two different things to edit. everything between a field's begin and end
 *  is dropped: the "3" in "page 3 of 9" is a result word recomputes. */
function flatten(paragraph: Element): { runs: TextRun[]; breaks: Break[] } {
  const runs: TextRun[] = [];
  const breaks: Break[] = [];
  let at = 0;
  let inField = 0;

  const walk = (parent: Element): void => {
    for (let node = parent.firstChild; node; node = node.nextSibling) {
      if (node.nodeType !== 1) continue;
      const el = node as Element;
      const tag = el.tagName;
      if (tag === 'w:p' || SKIP.has(tag)) continue;

      if (tag === 'w:fldChar') {
        const type = el.getAttribute('w:fldCharType');
        if (type === 'begin') inField++;
        else if (type === 'end') inField = Math.max(0, inField - 1);
        breaks.push({ at, field: true });
      } else if (tag === 'w:tab' || tag === 'w:ptab' || tag === 'w:br') {
        breaks.push({ at, field: false });
      } else if (tag === 'w:t') {
        if (inField > 0) continue;
        const text = el.textContent ?? '';
        runs.push({ node: el, at, text });
        at += text.length;
      } else {
        walk(el);
      }
    }
  };

  walk(paragraph);
  return { runs, breaks };
}

/** spans wrapped in a zero-width character: the template author saying
 *  "this bit is mine to change". */
function markedSegments(text: string): Segment[] {
  const marks: number[] = [];
  for (let i = 0; i < text.length; i++) {
    if (ZERO_WIDTH.test(text[i]!)) marks.push(i);
  }
  const out: Segment[] = [];
  for (let i = 0; i + 1 < marks.length; i += 2) {
    const from = marks[i]! + 1;
    const to = marks[i + 1]!;
    if (to > from) out.push({ from, to });
  }
  return out;
}

/** every stretch of plain text between interruptions, minus the ones that
 *  only exist to label a word field. */
function inferredSegments(text: string, breaks: readonly Break[]): Segment[] {
  const beside = new Map<number, boolean>([
    [0, false],
    [text.length, false],
  ]);
  for (const { at, field } of breaks) beside.set(at, field || (beside.get(at) ?? false));

  const edges = [...beside.keys()].sort((a, b) => a - b);
  const out: Segment[] = [];
  for (let i = 0; i + 1 < edges.length; i++) {
    const from = edges[i]!;
    const to = edges[i + 1]!;
    if (beside.get(from) || beside.get(to)) continue;
    if (text.slice(from, to).trim().length > 0) out.push({ from, to });
  }
  return out;
}

interface Placed {
  key: string;
  label: string;
  runs: TextRun[];
  segment: Segment;
}

/** every editable span in one header or footer, against a document the caller
 *  can go on to edit: a segment only means anything beside the nodes it was
 *  measured over. */
function place(partName: string, doc: Document): Placed[] {
  const out: Placed[] = [];

  tagged(doc.documentElement, 'w:p').forEach((paragraph, index) => {
    if (inFallback(paragraph)) return;

    const { runs, breaks } = flatten(paragraph);
    if (runs.length === 0) return;

    const text = runs.map((run) => run.text).join('');
    const segments = ZERO_WIDTH.test(text)
      ? markedSegments(text)
      : inferredSegments(text, breaks);

    segments.forEach((segment, slot) => {
      out.push({
        key: `${partName}#${index}.${slot}`,
        label: text.slice(segment.from, segment.to).trim(),
        runs,
        segment,
      });
    });
  });

  return out;
}

/** headers first, then footers: the order they read on the page */
const inReadingOrder = (parts: Record<string, string>): string[] =>
  Object.keys(parts)
    .filter((name) => HEADER_OR_FOOTER.test(name))
    .sort((a, b) => Number(a.includes('/footer')) - Number(b.includes('/footer')) || a.localeCompare(b));

/** every field in a template's headers and footers, in reading order. */
export function findFields(parts: Record<string, string>): Field[] {
  const out: Field[] = [];
  for (const name of inReadingOrder(parts)) {
    const doc = parseXml(parts[name]!, name);
    for (const { key, label } of place(name, doc)) out.push({ key, label, part: name });
  }
  return out;
}

/** replace the runs a segment covers with one run's worth of text. the value
 *  lands whole in the first node and the rest keep only what falls outside — a
 *  `w:r` carries the formatting, so emptying one is safe where deleting it
 *  would take the small caps with it. */
function writeSegment(runs: readonly TextRun[], segment: Segment, value: string): void {
  let written = false;
  for (const run of runs) {
    const from = Math.max(segment.from, run.at);
    const to = Math.min(segment.to, run.at + run.text.length);
    if (to <= from) continue;

    const before = run.text.slice(0, from - run.at);
    const after = run.text.slice(to - run.at);
    run.node.textContent = written ? before + after : before + value + after;
    run.node.setAttribute('xml:space', 'preserve');
    written = true;
  }
}

/** the template's parts with every named value filled in. a field with no value
 *  keeps the template's own text rather than printing a blank. */
export function fillFields(
  parts: Record<string, string>,
  values: Values,
): Record<string, string> {
  const filled: Record<string, string> = { ...parts };

  for (const name of Object.keys(parts)) {
    if (!HEADER_OR_FOOTER.test(name)) continue;

    const doc = parseXml(parts[name]!, name);
    let touched = false;

    for (const field of place(name, doc)) {
      const value = values[field.key];
      if (typeof value !== 'string' || value === field.label) continue;
      writeSegment(field.runs, field.segment, value);
      touched = true;
    }

    if (touched) filled[name] = serializeXml(doc);
  }

  return filled;
}
