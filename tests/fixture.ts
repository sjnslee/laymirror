// a minimal but structurally valid docx, built in memory. hermetic, and it
// lets a test choose whether docProps/custom.xml exists — the real donors
// differ on that and the two paths behave differently.

import { writeText, type Parts } from '../src/docx/zip.js';

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  '</Types>';

const RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

const DOCUMENT =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body></w:document>';

const CUSTOM_WITH_DOCID =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"' +
  ' xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="ContentTypeId">' +
  '<vt:lpwstr>0x0101000248</vt:lpwstr></property>' +
  '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="cmirDocId">' +
  '<vt:lpwstr>doc-abc-123</vt:lpwstr></property>' +
  '</Properties>';

export function makeDocx(opts: { custom?: boolean } = {}): Parts {
  const parts: Parts = {};
  writeText(parts, '[Content_Types].xml', CONTENT_TYPES);
  writeText(parts, '_rels/.rels', RELS);
  writeText(parts, 'word/document.xml', DOCUMENT);
  if (opts.custom ?? true) writeText(parts, 'docProps/custom.xml', CUSTOM_WITH_DOCID);
  return parts;
}

// ── a synthetic donor template ────────────────────────────────────────
// exercises everything readTemplate has to cope with in a real school
// template: basedOn chains, theme-referenced fonts, a section with a header
// and footer, and an attachedTemplate carrying an absolute path.

import { zip } from '../src/docx/zip.js';

const THEME = (major: string, minor: string) =>
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="t">' +
  '<a:themeElements><a:fontScheme name="f">' +
  `<a:majorFont><a:latin typeface="${major}"/></a:majorFont>` +
  `<a:minorFont><a:latin typeface="${minor}"/></a:minorFont>` +
  '</a:fontScheme></a:themeElements></a:theme>';

const STYLES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
  '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>' +
  // Normal names a real face
  '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>' +
  '<w:rPr><w:rFonts w:ascii="Palatino Linotype" w:hAnsi="Palatino Linotype"/><w:sz w:val="20"/></w:rPr></w:style>' +
  // heading 1: page break, centred, small caps, outline 0
  '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:pageBreakBefore/><w:spacing w:after="60"/><w:jc w:val="center"/><w:outlineLvl w:val="0"/></w:pPr>' +
  '<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:b/><w:smallCaps/><w:sz w:val="40"/></w:rPr></w:style>' +
  // headings 2 and 3 use the theme major font and carry no page break
  // heading 2 carries word's stock accent1 blue, as every real donor does —
  // a themed colour, so it must not reach the profile
  '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200"/><w:outlineLvl w:val="1"/></w:pPr>' +
  '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/><w:b/><w:color w:val="4F81BD" w:themeColor="accent1"/><w:sz w:val="26"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200"/><w:outlineLvl w:val="2"/></w:pPr>' +
  // heading 3's colour is named outright, so it is a real choice and stays
  '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/><w:b/><w:color w:val="7A0019"/></w:rPr></w:style>' +
  // Tag inherits font and size from Normal
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Tag"><w:name w:val="Tag"/><w:basedOn w:val="Normal"/>' +
  '<w:rPr><w:b/></w:rPr></w:style>' +
  // Cite inherits bold from Tag and the face from Normal, two links up
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Cite"><w:name w:val="Cite"/><w:basedOn w:val="Tag"/>' +
  '<w:rPr><w:u w:val="thick"/></w:rPr></w:style>' +
  // the card body takes the theme minor font, not Normal's face
  '<w:style w:type="paragraph" w:customStyle="1" w:styleId="card"><w:name w:val="card"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/><w:ind w:left="288" w:right="288"/></w:pPr>' +
  '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/><w:u w:val="single"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:customStyle="1" w:styleId="Underline"><w:name w:val="Underline"/>' +
  '<w:rPr><w:b w:val="0"/><w:sz w:val="20"/><w:u w:val="single"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:customStyle="1" w:styleId="Style13ptBold"><w:name w:val="Style 13 pt Bold"/>' +
  '<w:rPr><w:b/><w:sz w:val="26"/><w:u w:val="none"/></w:rPr></w:style>' +
  // something laymirror does not model, which must survive the rewrite
  '<w:style w:type="table" w:styleId="TableNormal"><w:name w:val="Normal Table"/></w:style>' +
  '</w:styles>';

const TEMPLATE_DOC =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
  '<w:body><w:p><w:r><w:t>donor</w:t></w:r></w:p>' +
  '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>' +
  '<w:footerReference w:type="default" r:id="rId11"/>' +
  '<w:pgSz w:w="12240" w:h="15840"/>' +
  '<w:pgMar w:top="720" w:right="720" w:bottom="1008" w:left="720" w:header="720" w:footer="720"/>' +
  '</w:sectPr></w:body></w:document>';

const SETTINGS_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"' +
  ' Target="file:////Users/someone/Library/Templates/Lay%20Cut%20Cards.dotx" TargetMode="External"/>' +
  '</Relationships>';

/** a donor template as raw docx bytes. */
export function makeTemplate(): Uint8Array {
  const parts = makeDocx();
  writeText(parts, 'word/document.xml', TEMPLATE_DOC);
  writeText(parts, 'word/styles.xml', STYLES);
  writeText(parts, 'word/theme/theme1.xml', THEME('Calibri', 'Cambria'));
  writeText(parts, 'word/_rels/settings.xml.rels', SETTINGS_RELS);
  writeText(parts, 'word/header1.xml', '<w:hdr><w:p><w:r><w:t>Team Code</w:t></w:r></w:p></w:hdr>');
  writeText(parts, 'word/footer1.xml', '<w:ftr><w:p><w:r><w:t>Page</w:t></w:r></w:p></w:ftr>');
  return zip(parts);
}
