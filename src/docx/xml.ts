// parse and serialise an ooxml part.
//
// `DOMParser` never puts the xml declaration in the tree, so it has to go back
// on by hand — word writes one on every part, and a part is easier to diff
// against word's own output with it there.
//
// but `XMLSerializer` is not consistent about it. chromium re-emits a
// declaration for a document that was parsed with one; jsdom emits none. so a
// naive prepend produces one declaration under test and two in the app, and a
// second declaration makes word and every other reader call the file corrupt.
// whatever the serialiser produced is stripped, and ours is the only one.

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

export function parseXml(xml: string, what: string): Document {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.getElementsByTagName('parsererror').length > 0) {
    throw new Error(`${what} did not parse`);
  }
  return doc;
}

export function serializeXml(doc: Document): string {
  const xml = new XMLSerializer().serializeToString(doc);
  return DECL + xml.replace(/^\s*<\?xml\b[^?]*\?>\s*/, '');
}
