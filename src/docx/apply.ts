// the save pipeline.
//
// cardmirror's exporter rebuilds the package from scratch on every save: its
// own styles.xml, one hardcoded letter section with 1" margins, and no header,
// footer or theme at all. so this is not a formatter. it is the thing that puts
// back what the exporter has just thrown away.
//
// the template is authoritative, every time. an earlier version asked whether
// word had written the file and adopted its header if so, which made a header
// something you edited in word and laymirror preserved. that is backwards for a
// squad: the school's header is fixed, and the two or three words inside it
// that change are typed into laymirror's panel, so a file that has been through
// word comes out looking exactly like a file that has not.

import { fillFields, type Values } from './fields.js';
import { writeMarker } from './marker.js';
import { restoreSnapshot } from './snapshot.js';
import { headerParts, type Blueprint } from '../template/template.js';
import { EXPORT_STYLE_BY_TYPE } from '../template/styles.js';
import { parseXml, serializeXml } from './xml.js';
import { isDocx, readText, strToBytes, unzip, writeText, zip, type Parts } from './zip.js';

const DOCUMENT = 'word/document.xml';
const SETTINGS = 'word/settings.xml';
const SETTINGS_RELS = 'word/_rels/settings.xml.rels';

const TEMPLATE_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate';
const TEMPLATE_REL_ID = 'rIdLayMirrorTemplate';

const CITE_EXPORT = EXPORT_STYLE_BY_TYPE.cite_mark!;
const UNDERLINE_EXPORT = EXPORT_STYLE_BY_TYPE.underline_mark!;

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

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
 *  only evidence left in the file. the inference reaches exactly one paragraph
 *  past the last mark it saw: guessing further would indent ordinary prose,
 *  while leaving a paragraph bare merely renders it as the template's Normal,
 *  which is the cheaper mistake. */
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
export function applyStyles(documentXml: string, blueprint: Blueprint): string {
  const doc = parseXml(documentXml, DOCUMENT);
  let openCard = false;

  for (const paragraph of elements(doc.documentElement, 'w:p')) {
    const pPr = directChild(paragraph, 'w:pPr');
    const pStyle = pPr ? directChild(pPr, 'w:pStyle') : null;
    const exported = pStyle?.getAttribute('w:val') ?? null;

    if (exported !== null) {
      const mapped = blueprint.styleMap[exported];
      if (mapped) pStyle!.setAttribute('w:val', mapped);
      // a heading, a tag or an analytic closes whatever card was open
      openCard = false;
    } else {
      const type = classifyBare(paragraph, openCard);
      const target = type ? blueprint.bareStyles[type] : null;
      if (target) setPStyle(doc, paragraph, target);
      openCard =
        type === 'cite_paragraph' ||
        elements(paragraph, 'w:rStyle').some(
          (el) => el.getAttribute('w:val') === UNDERLINE_EXPORT,
        );
    }

    for (const rStyle of elements(paragraph, 'w:rStyle')) {
      const mapped = blueprint.styleMap[rStyle.getAttribute('w:val') ?? ''];
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

/** put the school's document onto a package cardmirror has just written.
 *
 *  throws rather than returning something half-written: this runs against a
 *  file the user is actively saving, and a partial read must never become a
 *  partial write. */
export function applyTemplate(
  bytes: Uint8Array,
  blueprint: Blueprint,
  values: Values,
  templateId: string,
): Uint8Array {
  const parts = unzip(bytes);
  if (!isDocx(parts)) throw new Error('not a complete docx — read again in a moment');

  const documentXml = readText(parts, DOCUMENT);
  if (!documentXml) throw new Error('document.xml is unreadable');
  writeText(parts, DOCUMENT, applyStyles(documentXml, blueprint));

  const filled = fillFields(headerParts(blueprint.snapshot), values);
  const override: Record<string, Uint8Array> = {};
  for (const [name, xml] of Object.entries(filled)) override[name] = strToBytes(xml);

  restoreSnapshot(parts, blueprint.snapshot, override);
  pointAttachedTemplate(parts, blueprint.snapshot.attachedTemplate);
  writeMarker(parts, templateId);

  return zip(parts);
}
