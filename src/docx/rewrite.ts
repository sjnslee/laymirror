// the save pipeline: a cardmirror export in, a lay document word will open
// looking like the school's template out.
//
// cardmirror's exporter writes its own style ids (`Heading4` for a tag, and
// nothing at all for a cite paragraph or a card body), one hardcoded letter
// section with no header or footer, and a `Debate.dotm` attached template.
// every one of those is replaced here.

import { EXPORT_STYLE_BY_TYPE, TYPE_BY_EXPORT_STYLE } from '../profile/mapping.js';
import type { BlockType, Profile, RunType } from '../profile/profile.js';
import { buildSectPr, replaceSectPr } from './sect.js';
import { buildStylesXml } from './styles.js';
import { writeHeaderFooter, type DocMeta } from './headers.js';
import { writeMarker } from './marker.js';
import { parseXml, serializeXml } from './xml.js';
import { isDocx, readText, unzip, writeText, zip, type Parts } from './zip.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const DOCUMENT = 'word/document.xml';
const SETTINGS = 'word/settings.xml';
const SETTINGS_RELS = 'word/_rels/settings.xml.rels';

const TEMPLATE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate';
const TEMPLATE_REL_ID = 'rIdLayMirrorTemplate';

const CITE_EXPORT = EXPORT_STYLE_BY_TYPE.cite_mark!;
const UNDERLINE_EXPORT = EXPORT_STYLE_BY_TYPE.underline_mark!;

export type { DocMeta };

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

const styleIdFor = (profile: Profile, type: BlockType | RunType): string =>
  profile.types[type].styleId;

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
 *  own, so the marks inside them are the only evidence in the file. a cite is
 *  the paragraph carrying cite marks; a body is the one carrying underline
 *  marks, or the bare paragraph immediately after evidence — a card is a cite
 *  followed by its body, and the exporter flattens the card node away.
 *
 *  the inference deliberately reaches exactly one paragraph past the last
 *  mark it saw. guessing further would indent ordinary paragraphs, and a
 *  paragraph left bare merely renders as the profile's Normal, which is the
 *  cheaper mistake. */
function classifyBare(paragraph: Element, openCard: boolean): BlockType | null {
  const marks = new Set(
    elements(paragraph, 'w:rStyle')
      .map((el) => el.getAttribute('w:val'))
      .filter((val): val is string => val !== null),
  );

  if (marks.has(CITE_EXPORT)) return 'cite_paragraph';
  if (marks.has(UNDERLINE_EXPORT)) return 'card_body';
  return openCard ? 'card_body' : null;
}

/** cardmirror's export style ids -> the profile's, plus the two styles it
 *  never writes. */
function applyStyles(documentXml: string, profile: Profile): string {
  const doc = parseXml(documentXml, DOCUMENT);
  // the section we are about to write references its header and footer by
  // r:id, so the prefix has to be bound even if the export never used it
  if (!doc.documentElement.hasAttribute('xmlns:r')) {
    doc.documentElement.setAttribute('xmlns:r', R);
  }
  let openCard = false;

  for (const paragraph of elements(doc.documentElement, 'w:p')) {
    const pPr = directChild(paragraph, 'w:pPr');
    const pStyle = pPr ? directChild(pPr, 'w:pStyle') : null;
    const exported = pStyle?.getAttribute('w:val') ?? null;

    if (exported !== null) {
      const type = TYPE_BY_EXPORT_STYLE[exported];
      if (type) pStyle!.setAttribute('w:val', styleIdFor(profile, type));
      // a heading, a tag or an analytic closes whatever card was open
      openCard = false;
    } else {
      const type = classifyBare(paragraph, openCard);
      if (type) setPStyle(doc, paragraph, styleIdFor(profile, type));
      openCard =
        type === 'cite_paragraph' ||
        elements(paragraph, 'w:rStyle').some(
          (el) => el.getAttribute('w:val') === UNDERLINE_EXPORT,
        );
    }

    for (const rStyle of elements(paragraph, 'w:rStyle')) {
      const type = TYPE_BY_EXPORT_STYLE[rStyle.getAttribute('w:val') ?? ''];
      if (type) rStyle.setAttribute('w:val', styleIdFor(profile, type));
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
    settings.replace(/<w:settings\b([^>]*)>/, `<w:settings$1><w:attachedTemplate r:id="${TEMPLATE_REL_ID}"/>`),
  );
}

/** throws rather than returning something half-written: this runs against a
 *  file the user is actively saving, and a partial read must never become a
 *  partial write. */
export function rewriteDocx(bytes: Uint8Array, profile: Profile, meta: DocMeta): Uint8Array {
  const parts = unzip(bytes);
  if (!isDocx(parts)) throw new Error('not a complete docx — read again in a moment');

  const documentXml = readText(parts, DOCUMENT);
  if (!documentXml) throw new Error('document.xml is unreadable');

  writeText(parts, 'word/styles.xml', buildStylesXml(profile));

  const refs = writeHeaderFooter(parts, profile, meta);
  const styled = applyStyles(documentXml, profile);
  writeText(parts, DOCUMENT, replaceSectPr(styled, buildSectPr(profile.page, refs)));

  pointAttachedTemplate(parts, profile.attachedTemplate);
  writeMarker(parts, profile.id);

  return zip(parts);
}
