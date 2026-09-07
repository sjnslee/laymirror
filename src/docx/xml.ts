// parse and serialise an ooxml part.
//
// `DOMParser` drops the xml declaration and `XMLSerializer` is inconsistent
// about putting one back — chromium does, jsdom does not — and two declarations
// make word call the file corrupt. so whatever it produced is stripped first.

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

/** a live HTMLCollection is awkward to iterate and changes under an edit. */
export function elements(parent: Element, tag: string): Element[] {
  const found = parent.getElementsByTagName(tag);
  const out: Element[] = [];
  for (let i = 0; i < found.length; i++) out.push(found.item(i)!);
  return out;
}
