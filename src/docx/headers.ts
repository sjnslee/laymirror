// word/header1.xml and word/footer1.xml, plus the wiring that makes word
// believe in them.
//
// the donor's header is copied through as-is unless it carries tokens. that
// is deliberate: a real school header is a text box, a rule, a team code and
// a `PAGE of NUMPAGES` field pair, and nothing here can tell which literal
// run means "team code" and which means "1AC". guessing would print the
// authors where the template says something else. a school that wants a slot
// filled marks it `{{title}}`, `{{authors}}` or `{{team}}` and gets it.
//
// the page number is always left as a field. it is what keeps the judge's
// copy right even when our own paginator has drifted a page — flattening it
// to a literal number would put the drift on paper.

import { CONTENT_TYPES, readText, writeText, type Parts } from './zip.js';
import { parseXml, serializeXml } from './xml.js';
import type { SectRefs } from './sect.js';
import type { PageSetup, Profile, TypeSpec } from '../profile/profile.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export const HEADER_PART = 'word/header1.xml';
export const FOOTER_PART = 'word/footer1.xml';
const DOC_RELS = 'word/_rels/document.xml.rels';

/** named rather than numbered, like the marker's relationship: rewriting the
 *  same file twice must reuse the entry, not add a second one. */
export const HEADER_REL_ID = 'rIdLayMirrorHeader';
export const FOOTER_REL_ID = 'rIdLayMirrorFooter';

const HEADER_TYPE = `${R}/header`;
const FOOTER_TYPE = `${R}/footer`;
const HEADER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml';
const FOOTER_CT = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml';

const PACKAGE_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

const EMPTY_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
  `<Relationships xmlns="${PACKAGE_RELS_NS}"></Relationships>`;

export interface DocMeta {
  title: string;
  authors: string;
  teamCode: string;
}

export interface HeaderFooter {
  headerXml: string;
  footerXml: string;
}

// `team code` with or without the space, because both get typed
const TOKEN = /\{\{\s*(title|authors|team ?code|team)\s*\}\}/gi;

function valueFor(token: string, meta: DocMeta): string {
  const key = token.toLowerCase().replace(/\s+/g, '');
  if (key === 'title') return meta.title;
  if (key === 'authors') return meta.authors;
  return meta.teamCode;
}

/** fill the tokens in one paragraph, if it has any.
 *
 *  word splits a run wherever a revision id changes, so a donor's `{{title}}`
 *  can arrive as four `<w:t>`s and no per-element replace would ever see it.
 *  the paragraph's text is joined, matched, and written back run by run: the
 *  value lands in the run the token opened in, so that run's formatting is
 *  the one that survives, and tabs, ptabs and field characters are never part
 *  of the string so they cannot be disturbed. */
function fillTokens(paragraph: Element, meta: DocMeta): boolean {
  const nodes = paragraph.getElementsByTagName('w:t');
  const texts: Element[] = [];
  for (let i = 0; i < nodes.length; i++) texts.push(nodes.item(i)!);
  if (texts.length === 0) return false;

  // which run owns each character of the joined text
  const owner: number[] = [];
  let joined = '';
  texts.forEach((t, i) => {
    const s = t.textContent ?? '';
    joined += s;
    for (let c = 0; c < s.length; c++) owner.push(i);
  });

  const matches = [...joined.matchAll(TOKEN)];
  if (matches.length === 0) return false;

  const pieces = texts.map(() => '');
  const add = (run: number, s: string) => {
    pieces[run] = pieces[run]! + s;
  };

  let cursor = 0;
  for (const match of matches) {
    const start = match.index;
    for (let c = cursor; c < start; c++) add(owner[c]!, joined[c]!);
    add(owner[start]!, valueFor(match[1]!, meta));
    cursor = start + match[0].length;
  }
  for (let c = cursor; c < joined.length; c++) add(owner[c]!, joined[c]!);

  texts.forEach((t, i) => {
    const value = pieces[i]!;
    t.textContent = value;
    // word trims a run's leading and trailing space without this
    if (value !== value.trim()) t.setAttribute('xml:space', 'preserve');
  });
  return true;
}

/** returns the donor xml byte-identical when it holds no tokens — an
 *  untouched part is the strongest guarantee we can give a school. */
function substitute(xml: string, meta: DocMeta, what: string): string {
  // cheap gate. a token split across runs never matches the raw xml, but its
  // opening braces are still in there, so this only skips parts that could
  // not possibly hold one.
  if (!xml.includes('{{')) return xml;

  const doc = parseXml(xml, what);
  const paragraphs = doc.getElementsByTagName('w:p');
  let touched = false;
  for (let i = 0; i < paragraphs.length; i++) {
    if (fillTokens(paragraphs.item(i)!, meta)) touched = true;
  }
  return touched ? serializeXml(doc) : xml;
}

const escapeXml = (s: string): string =>
  s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const text = (s: string): string =>
  `<w:t xml:space="preserve">${escapeXml(s)}</w:t>`;

/** a live field, with a cached result word replaces on open. */
const field = (instruction: string): string =>
  '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
  `<w:r><w:instrText xml:space="preserve"> ${instruction} </w:instrText></w:r>` +
  '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
  '<w:r><w:t>1</w:t></w:r>' +
  '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

function runProps(spec: TypeSpec, bold: boolean): string {
  const bits: string[] = [];
  if (spec.font) bits.push(`<w:rFonts w:ascii="${spec.font}" w:hAnsi="${spec.font}"/>`);
  if (bold) bits.push('<w:b/>');
  if (spec.sizePt !== undefined) {
    const half = Math.round(spec.sizePt * 2);
    bits.push(`<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`);
  }
  return bits.length ? `<w:rPr>${bits.join('')}</w:rPr>` : '';
}

/** tab stops at the middle and the right edge of the text column, so a
 *  synthesised header lines up on whatever page the profile describes. */
function tabs(page: PageSetup): string {
  const usable = page.widthTwips - page.margin.left - page.margin.right;
  return (
    '<w:tabs>' +
    `<w:tab w:val="center" w:pos="${Math.round(usable / 2)}"/>` +
    `<w:tab w:val="right" w:pos="${usable}"/>` +
    '</w:tabs>'
  );
}

const open = (tag: 'w:hdr' | 'w:ftr'): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<${tag} xmlns:w="${W}" xmlns:r="${R}">`;

/** used when the profile has no donor header — the built-in profile, or a
 *  template whose header word never wrote. team code left, title right,
 *  authors underneath, over a rule. */
function synthesiseHeader(profile: Profile, meta: DocMeta): string {
  const spec = profile.types.paragraph;
  const rPr = runProps(spec, true);
  const rule = '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="auto"/></w:pBdr>';

  const top =
    `<w:p><w:pPr>${tabs(profile.page)}${meta.authors ? '' : rule}</w:pPr>` +
    `<w:r>${rPr}${text(meta.teamCode)}</w:r>` +
    `<w:r>${rPr}<w:tab/></w:r>` +
    `<w:r>${rPr}${text(meta.title)}</w:r></w:p>`;

  const second = meta.authors
    ? `<w:p><w:pPr>${tabs(profile.page)}${rule}</w:pPr><w:r>${rPr}${text(meta.authors)}</w:r></w:p>`
    : '';

  return `${open('w:hdr')}${top}${second}</w:hdr>`;
}

/** page numbers stay fields, always. */
function synthesiseFooter(profile: Profile): string {
  const rPr = runProps(profile.types.paragraph, false);
  return (
    `${open('w:ftr')}<w:p><w:pPr><w:jc w:val="center"/></w:pPr>` +
    `<w:r>${rPr}${text('Page ')}</w:r>${field('PAGE')}` +
    `<w:r>${rPr}${text(' of ')}</w:r>${field('NUMPAGES')}` +
    '</w:p></w:ftr>'
  );
}

export function buildHeaderFooter(profile: Profile, meta: DocMeta): HeaderFooter {
  return {
    headerXml: profile.headerXml
      ? substitute(profile.headerXml, meta, 'donor header1.xml')
      : synthesiseHeader(profile, meta),
    footerXml: profile.footerXml
      ? substitute(profile.footerXml, meta, 'donor footer1.xml')
      : synthesiseFooter(profile),
  };
}

function ensureOverride(parts: Parts, partName: string, contentType: string): void {
  const ct = readText(parts, CONTENT_TYPES);
  if (!ct || ct.includes(`PartName="/${partName}"`)) return;
  writeText(
    parts,
    CONTENT_TYPES,
    ct.replace('</Types>', `<Override PartName="/${partName}" ContentType="${contentType}"/></Types>`),
  );
}

function ensureRel(parts: Parts, id: string, type: string, target: string): void {
  const rels = readText(parts, DOC_RELS) ?? EMPTY_RELS;
  if (rels.includes(`Id="${id}"`)) {
    if (!readText(parts, DOC_RELS)) writeText(parts, DOC_RELS, rels);
    return;
  }
  writeText(
    parts,
    DOC_RELS,
    rels.replace(
      '</Relationships>',
      `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`,
    ),
  );
}

/** write both parts and everything word needs to find them, and hand back the
 *  ids `buildSectPr` references. */
export function writeHeaderFooter(parts: Parts, profile: Profile, meta: DocMeta): SectRefs {
  const { headerXml, footerXml } = buildHeaderFooter(profile, meta);

  writeText(parts, HEADER_PART, headerXml);
  writeText(parts, FOOTER_PART, footerXml);

  ensureOverride(parts, HEADER_PART, HEADER_CT);
  ensureOverride(parts, FOOTER_PART, FOOTER_CT);
  ensureRel(parts, HEADER_REL_ID, HEADER_TYPE, 'header1.xml');
  ensureRel(parts, FOOTER_REL_ID, FOOTER_TYPE, 'footer1.xml');

  return { headerRelId: HEADER_REL_ID, footerRelId: FOOTER_REL_ID };
}
