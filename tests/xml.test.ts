// @vitest-environment jsdom
import { describe, expect, it, vi } from 'vitest';
import { parseXml, serializeXml } from '../src/docx/xml.js';

const DOC = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:p xmlns:w="urn:w"/>';

const declarations = (xml: string) => [...xml.matchAll(/<\?xml\b/g)].length;

describe('serializeXml', () => {
  it('writes exactly one declaration, at the start', () => {
    const out = serializeXml(parseXml(DOC, 'test'));
    expect(declarations(out)).toBe(1);
    expect(out.startsWith('<?xml version="1.0"')).toBe(true);
  });

  // jsdom emits none and chromium emits one, so the app and the tests disagree
  // unless whatever the serialiser produced is stripped first. two declarations
  // make word call the file corrupt.
  it('writes one whatever the serialiser did', () => {
    vi.spyOn(XMLSerializer.prototype, 'serializeToString').mockReturnValue(DOC);
    const out = serializeXml(parseXml(DOC, 'test'));
    expect(declarations(out)).toBe(1);
    expect(out).toContain('<w:p');
    vi.restoreAllMocks();
  });
});

describe('parseXml', () => {
  it('names the part it could not read', () => {
    expect(() => parseXml('<w:p', 'word/document.xml')).toThrow('word/document.xml');
  });
});
