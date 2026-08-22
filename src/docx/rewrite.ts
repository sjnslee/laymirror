// the save pipeline.
//
// cardmirror's exporter rebuilds the package from scratch on every save: its
// own styles.xml, one hardcoded letter section with 1" margins, and no
// header, footer or theme at all. so this is not a formatter. it is the thing
// that puts back what the exporter has just thrown away.
//
// exactly one question decides what happens, and only word can answer it: does
// the file carry a header reference? cardmirror never writes one, so if there
// is one, word wrote this file and it is authoritative — we re-adopt it and
// touch nothing. if there is not, cardmirror just stripped it, and we restore.

import { EXPORT_STYLE_BY_TYPE, LEGACY_SENTINEL } from '../profile/mapping.js';
import type { Profile } from '../profile/profile.js';
import { injectBreaks, paragraphSpans, type PageBreak } from './breaks.js';
import { writeMarker } from './marker.js';
import {
  captureSnapshot,
  hasOwnHeader,
  restoreSnapshot,
  type Snapshot,
} from './snapshot.js';
import { parseXml, serializeXml } from './xml.js';
import { isDocx, readText, unzip, writeText, zip, type Parts } from './zip.js';

const DOCUMENT = 'word/document.xml';
const SETTINGS = 'word/settings.xml';
const SETTINGS_RELS = 'word/_rels/settings.xml.rels';

const TEMPLATE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate';
const TEMPLATE_REL_ID = 'rIdLayMirrorTemplate';

const CITE_EXPORT = EXPORT_STYLE_BY_TYPE.cite_mark!;
const UNDERLINE_EXPORT = EXPORT_STYLE_BY_TYPE.underline_mark!;

export type SaveOutcome =
  /** word wrote this file. its header is the truth; keep it. */
  | { kind: 'adopted'; snapshot: Snapshot }
  /** cardmirror wrote it, and we put the school's document back. */
  | { kind: 'restored'; bytes: Uint8Array }
  | { kind: 'skipped'; because: string };

function directChild(parent: Element, tag: string): Element | null {
  const children = parent.childNodes;
  for (let i = 0; i < children.length; i++) {
    const node = children.item(i);
    if (node?.nodeType === 1 && (node as Element).tagName === tag) return node as Element;
  }
  return null;
}

function elements(parent: Element, tag: string): Element[] {
  const found = parent.getElementsByTagName(tag);
  const out: Element[] = [];
  for (let i = 0; i < found.length; i++) out.push(found.item(i)!);
  return out;
}

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function setPStyle(doc: Document, paragraph: Element, styleId: string): void {
  let pPr = directChild(paragraph, 'w:pPr');
  if (!pPr) {
    pPr = doc.createElementNS(W, 'w:pPr');
    paragraph.insertBefore(pPr, paragraph.firstChild);
  }
  let pStyle = directChild(pPr, 'w:pStyle');
  if (!pStyle) {
    pStyle = doc.createElementNS(W, 'w:pStyle');
    // the schema wants pStyle first in pPr
    pPr.insertBefore(pStyle, pPr.firstChild);
  }
  pStyle.setAttribute('w:val', styleId);
}

/** which type a bare paragraph is, judged by the marks its runs carry.
 *
 *  a cite paragraph and a card body leave cardmirror with no style of their
 *  own, and `card` is flattened away on export, so the marks inside are the
 *  only evidence left in the file. the inference reaches exactly one
 *  paragraph past the last mark it saw: guessing further would indent
 *  ordinary prose, while leaving a paragraph bare merely renders it as the
 *  template's Normal, which is the cheaper mistake. */
function classifyBare(
  paragraph: Element,
  openCard: boolean,
): 'cite_paragraph' | 'card_body' | null {
  const marks = new Set(
    elements(paragraph, 'w:rStyle')
      .map((el) => el.getAttribute('w:val'))
      .filter((val): val is string => val !== null),
  );

  if (marks.has(CITE_EXPORT)) return 'cite_paragraph';
  if (marks.has(UNDERLINE_EXPORT)) return 'card_body';
  return openCard ? 'card_body' : null;
}

/** cardmirror's export style ids -> the template's own. */
export function applyStyles(documentXml: string, profile: Profile): string {
  const doc = parseXml(documentXml, DOCUMENT);
  let openCard = false;

  for (const paragraph of elements(doc.documentElement, 'w:p')) {
    const pPr = directChild(paragraph, 'w:pPr');
    const pStyle = pPr ? directChild(pPr, 'w:pStyle') : null;
    const exported = pStyle?.getAttribute('w:val') ?? null;

    if (exported !== null) {
      const mapped = profile.styleMap[exported];
      if (mapped) pStyle!.setAttribute('w:val', mapped);
      // a heading, a tag or an analytic closes whatever card was open
      openCard = false;
    } else {
      const type = classifyBare(paragraph, openCard);
      const target = type ? profile.bareStyles[type] : null;
      if (target) setPStyle(doc, paragraph, target);
      openCard =
        type === 'cite_paragraph' ||
        elements(paragraph, 'w:rStyle').some(
          (el) => el.getAttribute('w:val') === UNDERLINE_EXPORT,
        );
    }

    for (const rStyle of elements(paragraph, 'w:rStyle')) {
      const mapped = profile.styleMap[rStyle.getAttribute('w:val') ?? ''];
      if (mapped) rStyle.setAttribute('w:val', mapped);
    }
  }

  return serializeXml(doc);
}

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '</Relationships>';

/** word matches an attached template by basename out of the user's templates
 *  folder, so a basename is both the safe target and the working one. */
function pointAttachedTemplate(parts: Parts, template: string | null): void {
  if (!template) return;

  const rels = readText(parts, SETTINGS_RELS);
  if (rels?.includes(TEMPLATE_REL_TYPE)) {
    writeText(
      parts,
      SETTINGS_RELS,
      rels.replace(
        /(<Relationship\b[^>]*attachedTemplate"[^>]*\bTarget=")[^"]*(")/,
        `$1${template}$2`,
      ),
    );
    return;
  }

  const settings = readText(parts, SETTINGS);
  if (!settings) return;

  const relationship =
    `<Relationship Id="${TEMPLATE_REL_ID}" Type="${TEMPLATE_REL_TYPE}"` +
    ` Target="${template}" TargetMode="External"/>`;
  writeText(
    parts,
    SETTINGS_RELS,
    (rels ?? EMPTY_RELS).replace('</Relationships>', `${relationship}</Relationships>`),
  );
  writeText(
    parts,
    SETTINGS,
    settings.replace(
      /<w:settings\b([^>]*)>/,
      `<w:settings$1><w:attachedTemplate r:id="${TEMPLATE_REL_ID}"/>`,
    ),
  );
}

/** an earlier laymirror carried manual breaks as the literal text
 *  `[page break]`. breaks live outside the document now, so a paragraph that
 *  is nothing but the old sentinel is swept up rather than printed. */
export function dropLegacySentinel(documentXml: string): string {
  const doomed = paragraphSpans(documentXml).filter(
    (span) =>
      span.xml
        .replace(/<[^>]*>/g, '')
        .trim() === LEGACY_SENTINEL,
  );
  let xml = documentXml;
  for (const span of doomed.reverse()) {
    xml = xml.slice(0, span.start) + xml.slice(span.end);
  }
  return xml;
}

/** throws rather than returning something half-written: this runs against a
 *  file the user is actively saving, and a partial read must never become a
 *  partial write. */
export function applyProfile(
  bytes: Uint8Array,
  profile: Profile,
  breaks: readonly PageBreak[] = [],
): SaveOutcome {
  const parts = unzip(bytes);
  if (!isDocx(parts)) throw new Error('not a complete docx — read again in a moment');

  // only word writes a header reference. if there is one, this file is word's
  // and its header carries whatever the user typed into it — adopt, never
  // overwrite. this is what makes a hand-edited team code permanent.
  if (hasOwnHeader(parts)) {
    const snapshot = captureSnapshot(parts);
    return snapshot
      ? { kind: 'adopted', snapshot }
      : { kind: 'skipped', because: 'nothing to adopt' };
  }

  if (!profile.snapshot) return { kind: 'skipped', because: 'no template loaded' };

  const documentXml = readText(parts, DOCUMENT);
  if (!documentXml) throw new Error('document.xml is unreadable');

  writeText(parts, DOCUMENT, applyStyles(dropLegacySentinel(documentXml), profile));
  restoreSnapshot(parts, profile.snapshot);

  const restored = readText(parts, DOCUMENT)!;
  writeText(parts, DOCUMENT, injectBreaks(restored, breaks));

  pointAttachedTemplate(parts, profile.snapshot.attachedTemplate);
  writeMarker(parts, profile.id);

  return { kind: 'restored', bytes: zip(parts) };
}
