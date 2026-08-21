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
