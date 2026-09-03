// parse and serialise an ooxml part.
//
// `DOMParser` drops the xml declaration and `XMLSerializer` is inconsistent about
// putting one back — chromium does, jsdom does not — so a naive prepend writes
// one under test and two in the app. two makes word call the file corrupt, so
// whatever the serialiser produced is stripped and ours is the only one.

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
