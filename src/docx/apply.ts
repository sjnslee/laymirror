// the save pipeline: put back what cardmirror's exporter threw away.
//
// the template is authoritative every time, so a file that has been through word
// comes out like one that has not. the few words in the header that do change
// are typed into laymirror's panel.

import { fillFields, type Values } from './fields.js';
import { writeMarker } from './marker.js';
import { EMPTY_RELS, restoreSnapshot } from './snapshot.js';
import { headerParts, type Blueprint } from '../template/template.js';
import { EXPORT_STYLE_BY_TYPE } from '../template/styles.js';
import { elements, parseXml, serializeXml } from './xml.js';
import { isDocx, readText, unzip, writeText, zip, type Parts } from './zip.js';

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

/** which type a bare paragraph is, judged by the marks its runs carry: a cite
 *  paragraph and a card body export with no style at all.
 *
 *  the inference reaches exactly one paragraph past the last mark — further
 *  would indent ordinary prose. */
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
function applyStyles(documentXml: string, blueprint: Blueprint): string {
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

/** word matches an attached template by basename out of the user's templates
 *  folder, so a basename is both the safe target and the working one.
 *
 *  only where cardmirror's export has none: the one it writes is what lights up
 *  verbatim's ribbon, and replacing it loses that for teammates on verbatim. */
function pointAttachedTemplate(parts: Parts, template: string | null): void {
  if (!template) return;

  const rels = readText(parts, SETTINGS_RELS);
  if (rels?.includes(TEMPLATE_REL_TYPE)) return;

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

/** put the template onto a package cardmirror has just written. throws on a
 *  partial read rather than turning it into a partial write. */
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

  restoreSnapshot(parts, blueprint.snapshot, fillFields(headerParts(blueprint.snapshot), values));
  pointAttachedTemplate(parts, blueprint.snapshot.attachedTemplate);
  writeMarker(parts, templateId);

  return zip(parts);
}
