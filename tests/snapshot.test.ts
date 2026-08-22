import { describe, expect, it } from 'vitest';
import { makeExport, makeTemplate } from './fixture.js';
import { readText, unzip, writeText } from '../src/docx/zip.js';
import {
  captureSnapshot,
  hasOwnHeader,
  readSectPr,
  restoreSnapshot,
  retargetSectPr,
} from '../src/docx/snapshot.js';

const template = () => unzip(makeTemplate());
const exported = () => unzip(makeExport());

describe('hasOwnHeader', () => {
  it('is true for a document word wrote', () => {
    const parts = template();
    writeText(
      parts,
      'word/_rels/document.xml.rels',
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '</Relationships>',
    );
    expect(hasOwnHeader(parts)).toBe(true);
  });

  // the whole save pipeline turns on this: cardmirror's exporter never emits
  // a header reference, so its absence means cardmirror just wrote the file
  it('is false for a cardmirror export', () => {
    expect(hasOwnHeader(exported())).toBe(false);
  });
});

describe('captureSnapshot', () => {
  const withRels = () => {
    const parts = template();
    writeText(
      parts,
      'word/_rels/document.xml.rels',
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
        '</Relationships>',
    );
    return parts;
  };

  it('carries the header and footer byte for byte', () => {
    const parts = withRels();
    const snap = captureSnapshot(parts)!;
    expect(snap.parts['word/header1.xml']).toBe(readText(parts, 'word/header1.xml'));
    expect(snap.parts['word/footer1.xml']).toBe(readText(parts, 'word/footer1.xml'));
  });

  it('carries styles, theme and font table', () => {
    const snap = captureSnapshot(withRels())!;
    expect(snap.parts['word/styles.xml']).toContain('Palatino Linotype');
    expect(snap.parts['word/theme/theme1.xml']).toContain('Calibri');
  });

  it('keeps the section with its real margins', () => {
    const snap = captureSnapshot(withRels())!;
    expect(snap.sectPr).toContain('w:bottom="1008"');
    expect(snap.sectPr).toContain('w:left="720"');
  });

  it('reduces the attached template to a basename', () => {
    const snap = captureSnapshot(withRels())!;
    expect(snap.attachedTemplate).toBe('Lay%20Cut%20Cards.dotx');
  });

  it('returns null when there is no identity to keep', () => {
    const bare: Record<string, Uint8Array> = {};
    writeText(bare, '[Content_Types].xml', '<Types></Types>');
    writeText(bare, 'word/document.xml', '<w:document><w:body/></w:document>');
    expect(captureSnapshot(bare)).toBeNull();
  });
});

describe('restoreSnapshot', () => {
  const snapshotOfTemplate = () => {
    const parts = template();
    writeText(
      parts,
      'word/_rels/document.xml.rels',
      '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
        '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
        '</Relationships>',
    );
    return captureSnapshot(parts)!;
  };

  it('puts the school header back onto a cardmirror export', () => {
    const parts = exported();
    expect(hasOwnHeader(parts)).toBe(false);
    restoreSnapshot(parts, snapshotOfTemplate());
    expect(hasOwnHeader(parts)).toBe(true);
    expect(readText(parts, 'word/header1.xml')).toContain('PAGE');
  });

  it('replaces cardmirror 1in margins with the school section', () => {
    const parts = exported();
    expect(readText(parts, 'word/document.xml')).toContain('w:bottom="1440"');
    restoreSnapshot(parts, snapshotOfTemplate());
    const doc = readText(parts, 'word/document.xml')!;
    expect(doc).toContain('w:bottom="1008"');
    expect(doc).not.toContain('w:bottom="1440"');
  });

  it('rewrites section references onto ids it minted', () => {
    const parts = exported();
    restoreSnapshot(parts, snapshotOfTemplate());
    const doc = readText(parts, 'word/document.xml')!;
    const rels = readText(parts, 'word/_rels/document.xml.rels')!;
    // the donor's own rId10 would collide with the fresh package's ids
    expect(doc).not.toContain('r:id="rId10"');
    const id = /<w:headerReference[^>]*r:id="([^"]+)"/.exec(doc)?.[1];
    expect(id).toMatch(/^rIdLayMirror/);
    expect(rels).toContain(`Id="${id}"`);
  });

  it('binds the r prefix the restored section needs', () => {
    const parts = exported();
    expect(readText(parts, 'word/document.xml')).not.toContain('xmlns:r=');
    restoreSnapshot(parts, snapshotOfTemplate());
    expect(readText(parts, 'word/document.xml')).toContain('xmlns:r=');
  });

  it('declares the restored parts in content types', () => {
    const parts = exported();
    restoreSnapshot(parts, snapshotOfTemplate());
    expect(readText(parts, '[Content_Types].xml')).toContain('/word/header1.xml');
  });

  it('is idempotent — restoring twice adds one relationship, not two', () => {
    const parts = exported();
    const snap = snapshotOfTemplate();
    restoreSnapshot(parts, snap);
    restoreSnapshot(parts, snap);
    const rels = readText(parts, 'word/_rels/document.xml.rels')!;
    expect([...rels.matchAll(/relationships\/header/g)]).toHaveLength(1);
  });

  it('leaves parts it does not own alone', () => {
    const parts = exported();
    restoreSnapshot(parts, snapshotOfTemplate());
    expect(readText(parts, 'word/numbering.xml')).toBe('<w:numbering/>');
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
