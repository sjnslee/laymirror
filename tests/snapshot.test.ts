import { describe, expect, it } from 'vitest';
import { makeExport, makeTemplate } from './fixture.js';
import { readText, unzip, writeText, type Parts } from '../src/docx/zip.js';
import {
  captureSnapshot,
  readSectPr,
  restoreSnapshot,
  retargetSectPr,
} from '../src/docx/snapshot.js';

const template = () => unzip(makeTemplate());
const exported = () => unzip(makeExport());
const snapshot = () => captureSnapshot(template())!;

const hasHeader = (parts: Parts): boolean =>
  (readText(parts, 'word/_rels/document.xml.rels') ?? '').includes('relationships/header');

describe('captureSnapshot', () => {
  it('carries the header and footer byte for byte', () => {
    const parts = template();
    const snap = captureSnapshot(parts)!;
    expect(snap.parts['word/header1.xml']).toEqual(parts['word/header1.xml']);
    expect(snap.parts['word/footer1.xml']).toEqual(parts['word/footer1.xml']);
  });

  it('carries styles, theme, font table and numbering', () => {
    const snap = snapshot();
    expect(readText(snap.parts, 'word/styles.xml')).toContain('Palatino Linotype');
    expect(readText(snap.parts, 'word/theme/theme1.xml')).toContain('Calibri');
    expect(readText(snap.parts, 'word/numbering.xml')).toContain('w:numId="7"');
  });

  // a header's crest is a part the header points at, and a header copied
  // without it puts a red x on every page
  it('carries what the header itself relates to', () => {
    const snap = snapshot();
    expect(snap.parts['word/media/crest.png']).toBeDefined();
    expect(readText(snap.parts, 'word/_rels/header1.xml.rels')).toContain('crest.png');
  });

  it('leaves an external target alone', () => {
    expect(Object.keys(snapshot().parts)).not.toContain('https://example.org');
  });

  it('keeps the section with its real margins', () => {
    const snap = snapshot();
    expect(snap.sectPr).toContain('w:bottom="1008"');
    expect(snap.sectPr).toContain('w:left="720"');
  });

  it('reduces the attached template to a basename', () => {
    expect(snapshot().attachedTemplate).toBe('Lay%20Cut%20Cards.dotx');
  });

  it('returns null when there is no identity to keep', () => {
    const bare: Parts = {};
    writeText(bare, '[Content_Types].xml', '<Types></Types>');
    writeText(bare, 'word/document.xml', '<w:document><w:body/></w:document>');
    expect(captureSnapshot(bare)).toBeNull();
  });
});

describe('restoreSnapshot', () => {
  it('puts the school header back onto a cardmirror export', () => {
    const parts = exported();
    expect(hasHeader(parts)).toBe(false);
    restoreSnapshot(parts, snapshot());
    expect(hasHeader(parts)).toBe(true);
    expect(readText(parts, 'word/header1.xml')).toContain('PAGE');
  });

  it('replaces cardmirror 1in margins with the school section', () => {
    const parts = exported();
    expect(readText(parts, 'word/document.xml')).toContain('w:bottom="1440"');
    restoreSnapshot(parts, snapshot());
    const doc = readText(parts, 'word/document.xml')!;
    expect(doc).toContain('w:bottom="1008"');
    expect(doc).not.toContain('w:bottom="1440"');
  });

  it('rewrites section references onto ids it minted', () => {
    const parts = exported();
    restoreSnapshot(parts, snapshot());
    const doc = readText(parts, 'word/document.xml')!;
    const rels = readText(parts, 'word/_rels/document.xml.rels')!;
    // the template's own rId10 would collide with the fresh package's ids
    expect(doc).not.toContain('r:id="rId10"');
    const id = /<w:headerReference[^>]*r:id="([^"]+)"/.exec(doc)?.[1];
    expect(id).toMatch(/^rIdLayMirror/);
    expect(rels).toContain(`Id="${id}"`);
  });

  it('binds the r prefix the restored section needs', () => {
    const parts = exported();
    expect(readText(parts, 'word/document.xml')).not.toContain('xmlns:r=');
    restoreSnapshot(parts, snapshot());
    expect(readText(parts, 'word/document.xml')).toContain('xmlns:r=');
  });

  it('declares the restored parts in content types', () => {
    const parts = exported();
    restoreSnapshot(parts, snapshot());
    const ct = readText(parts, '[Content_Types].xml')!;
    expect(ct).toContain('/word/header1.xml');
    expect(ct).toContain('/word/numbering.xml');
  });

  // a part with no declared content type makes word call the whole file
  // corrupt, and an image is declared by extension rather than by name
  it('declares the extension a carried image needs', () => {
    const parts = exported();
    expect(readText(parts, '[Content_Types].xml')).not.toContain('image/png');
    restoreSnapshot(parts, snapshot());
    expect(readText(parts, '[Content_Types].xml')).toContain('image/png');
  });

  it('is idempotent — restoring twice adds one relationship, not two', () => {
    const parts = exported();
    const snap = snapshot();
    restoreSnapshot(parts, snap);
    restoreSnapshot(parts, snap);
    const rels = readText(parts, 'word/_rels/document.xml.rels')!;
    expect([...rels.matchAll(/relationships\/header/g)]).toHaveLength(1);
  });

  it('leaves parts it does not own alone', () => {
    const parts = exported();
    restoreSnapshot(parts, snapshot());
    expect(readText(parts, 'word/settings.xml')).toContain('attachedTemplate');
  });

  // the header the user has filled in, without the snapshot ever being written
  it('takes an overriding part in place of the snapshot', () => {
    const parts = exported();
    const snap = snapshot();
    restoreSnapshot(parts, snap, {
      'word/header1.xml': new TextEncoder().encode('<w:hdr>filled</w:hdr>'),
    });
    expect(readText(parts, 'word/header1.xml')).toBe('<w:hdr>filled</w:hdr>');
    expect(readText(snap.parts, 'word/header1.xml')).toContain('Team ');
  });
});

describe('retargetSectPr', () => {
  // a dangling r:id makes word declare the file corrupt, so a reference whose
  // part did not come back has to go rather than point at nothing
  it('drops a reference whose part is missing', () => {
    const sect =
      '<w:sectPr><w:headerReference w:type="default" r:id="rId10"/>' +
      '<w:footerReference w:type="default" r:id="rId11"/><w:pgSz w:w="12240"/></w:sectPr>';
    const out = retargetSectPr(sect, { rId10: 'rIdNew' });
    expect(out).toContain('r:id="rIdNew"');
    expect(out).not.toContain('rId11');
    expect(out).not.toContain('footerReference');
  });
});

describe('readSectPr', () => {
  it('takes the body section, which is the last one', () => {
    const xml =
      '<w:body><w:p><w:pPr><w:sectPr><w:pgSz w:w="1"/></w:sectPr></w:pPr></w:p>' +
      '<w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:body>';
    expect(readSectPr(xml)).toContain('12240');
  });
});

// word resolves the theme and the font table through a relationship, and
// cardmirror's exporter never writes one — so a carried theme would be a part
// word simply never reads
describe('restoreSnapshot — relationships word needs', () => {
  it('relates the theme, font table and numbering it carried', () => {
    const parts = exported();
    restoreSnapshot(parts, snapshot());
    const rels = readText(parts, 'word/_rels/document.xml.rels')!;
    expect(rels).toContain('relationships/theme');
    expect(rels).toContain('relationships/fontTable');
    expect(rels).toContain('relationships/numbering');
  });

  it('leaves the one cardmirror already wrote alone', () => {
    const parts = exported();
    restoreSnapshot(parts, snapshot());
    const rels = readText(parts, 'word/_rels/document.xml.rels')!;
    expect([...rels.matchAll(/relationships\/styles/g)]).toHaveLength(1);
  });
});
