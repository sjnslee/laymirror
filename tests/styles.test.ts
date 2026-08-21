// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { buildStylesXml } from '../src/docx/styles.js';
import { buildSectPr, replaceSectPr } from '../src/docx/sect.js';
import { readTemplate } from '../src/profile/read-template.js';
import { DEFAULT_LAY } from '../src/profile/defaults.js';
import { unzip, readText } from '../src/docx/zip.js';

const bytes = new Uint8Array(readFileSync('tests/fixtures/donor.docx'));
const { profile } = readTemplate(bytes, DEFAULT_LAY);
const xml = buildStylesXml(profile);

/** the <w:style> element for a styleId. */
function styleFor(id: string): string {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const all = doc.getElementsByTagName('w:style');
  for (let i = 0; i < all.length; i++) {
    const s = all.item(i)!;
    if (s.getAttribute('w:styleId') === id) return new XMLSerializer().serializeToString(s);
  }
  throw new Error(`no style ${id}`);
}

describe('buildStylesXml', () => {
  it('writes the lay names cardmirror reimports by', () => {
    expect(styleFor('Tag')).toContain('w:val="Tag"');
    expect(styleFor('Cite')).toContain('w:val="Cite"');
    expect(styleFor('card')).toContain('w:val="card"');
  });

  it('keeps the pocket page-breaking, centred and small-caps', () => {
    const s = styleFor('Heading1');
    expect(s).toContain('<w:pageBreakBefore');
    expect(s).toContain('w:val="center"');
    expect(s).toContain('<w:smallCaps');
    expect(s).toContain('w:val="40"'); // 20pt in half-points
  });

  it('resolves the card body font to a real face, not a theme reference', () => {
    const s = styleFor('card');
    expect(s).toContain('w:ascii="Cambria"');
    expect(s).not.toContain('asciiTheme');
    expect(s).toContain('w:left="288"');
    expect(s).toContain('w:right="288"');
  });

  it('emits the sentinels that keep cardmirror on its native import path', () => {
    // without these, a cite mark is silently dropped on reimport
    expect(() => styleFor('StyleUnderline')).not.toThrow();
    expect(() => styleFor('Emphasis')).not.toThrow();
    expect(() => styleFor('Style13ptBold')).not.toThrow();
  });

  it('turns bold off explicitly where the profile says false', () => {
    expect(styleFor('Underline')).toContain('<w:b w:val="0"/>');
  });

  it('leaves donor styles it does not model alone', () => {
    const donorStyleCount = (readText(unzip(bytes), 'word/styles.xml') ?? '').match(
      /<w:style\b/g,
    )?.length;
    const outCount = xml.match(/<w:style\b/g)?.length ?? 0;
    expect(outCount).toBeGreaterThanOrEqual(donorStyleCount ?? 0);
    expect(xml).toContain('TableNormal');
  });

  it('round-trips through a parser without error', () => {
    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    expect(doc.getElementsByTagName('parsererror')).toHaveLength(0);
  });
});

describe('buildSectPr', () => {
  it('writes the donor page size and margins', () => {
    const sect = buildSectPr(profile.page);
    expect(sect).toContain('w:w="12240"');
    expect(sect).toContain('w:h="15840"');
    expect(sect).toContain('w:top="720"');
    expect(sect).toContain('w:bottom="1008"');
  });

  it('references the header and footer when there are any', () => {
    const sect = buildSectPr(profile.page, { headerRelId: 'rId10', footerRelId: 'rId11' });
    expect(sect).toContain('<w:headerReference w:type="default" r:id="rId10"/>');
    expect(sect).toContain('<w:footerReference w:type="default" r:id="rId11"/>');
  });

  it('omits the references when there are none', () => {
    expect(buildSectPr(profile.page)).not.toContain('headerReference');
  });

  it('replaces cardmirror hardcoded section wholesale', () => {
    const before =
      '<w:body><w:p/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/>' +
      '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body>';
    const after = replaceSectPr(before, buildSectPr(profile.page));
    expect(after.match(/<w:sectPr/g)).toHaveLength(1);
    expect(after).toContain('w:top="720"');
    expect(after).not.toContain('w:top="1440"');
  });

  it('adds a section to a body that has none', () => {
    const after = replaceSectPr('<w:body><w:p/></w:body>', buildSectPr(profile.page));
    expect(after).toContain('<w:sectPr>');
    expect(after.indexOf('<w:sectPr>')).toBeLessThan(after.indexOf('</w:body>'));
  });
});
