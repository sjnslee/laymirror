// parse and serialise an ooxml part.
//
// `XMLSerializer` drops the xml declaration and `DOMParser` never puts it in
// the tree, so it goes back on by hand — word writes one on every part and a
// part is easier to diff against word's own output with it there.

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export function parseXml(xml: string, what: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`${what} did not parse`);
  }
  return doc;
}

export function serializeXml(doc: Document): string {
  return DECL + new XMLSerializer().serializeToString(doc);
}
