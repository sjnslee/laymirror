// donor, read off the donor's styles.xml and inlined so the plugin works
// before anyone picks a template. values are resolved (basedOn chains already
// followed) because css needs concrete numbers.
//
// note the donor's card body is the theme's minor font — cambria — not
// palatino, and its headings 2-4 are the major font, calibri. only heading 1
// names times new roman outright.

import type { Profile } from './profile.js';

/** cardmirror ships metric-compatible substitutes for cambria (caladea),
 *  calibri (carlito) and times new roman (tinos). palatino linotype and
 *  garamond have none, which is where pagination drift comes from. */
export const FONT_FALLBACKS: Record<string, string> = {
  'Palatino Linotype': '"Palatino Linotype", Palatino, "Book Antiqua", serif',
  'Times New Roman': 'Tinos, "Times New Roman", serif',
  Cambria: 'Caladea, Cambria, serif',
  Calibri: 'Carlito, Calibri, sans-serif',
  Garamond: 'Garamond, "EB Garamond", serif',
};

export const DEFAULT_LAY: Profile = {
  id: 'sample-lay',
  name: 'generic lay',

  // letter, 0.5" margins except a 0.7" bottom
  page: {
    widthTwips: 12240,
    heightTwips: 15840,
    margin: { top: 720, right: 720, bottom: 1008, left: 720, header: 720, footer: 720 },
  },

  types: {
    paragraph: {
      styleId: 'Normal',
      styleName: 'Normal',
      font: 'Palatino Linotype',
      sizePt: 10,
    },
    pocket: {
      styleId: 'Heading1',
      styleName: 'heading 1',
      font: 'Times New Roman',
      sizePt: 20,
      bold: true,
      smallCaps: true,
      align: 'center',
      spaceAfterPt: 3,
      pageBreakBefore: true,
      outlineLevel: 0,
    },
    hat: {
      styleId: 'Heading2',
      styleName: 'heading 2',
      font: 'Calibri',
      sizePt: 13,
      bold: true,
      color: '4F81BD',
      spaceBeforePt: 10,
      keepNext: true,
      keepLines: true,
      outlineLevel: 1,
    },
    block: {
      styleId: 'Heading3',
      styleName: 'heading 3',
      font: 'Calibri',
      sizePt: 10,
      bold: true,
      color: '4F81BD',
      spaceBeforePt: 10,
      keepNext: true,
      keepLines: true,
      outlineLevel: 2,
    },
    tag: {
      styleId: 'Tag',
      styleName: 'Tag',
      font: 'Palatino Linotype',
      sizePt: 10,
      bold: true,
      keepNext: true,
    },
    cite_paragraph: {
      styleId: 'Cite',
      styleName: 'Cite',
      font: 'Palatino Linotype',
      sizePt: 10,
      bold: true,
      underline: 'thick',
    },
    card_body: {
      styleId: 'card',
      styleName: 'card',
      font: 'Cambria',
      sizePt: 10,
      underline: 'single',
      indentLeftDxa: 288,
      indentRightDxa: 288,
      spaceAfterPt: 8,
      lineSpacing: { rule: 'auto', value: 259 },
    },
    // absent from the donor — ours to define
    analytic: {
      styleId: 'Analytic',
      styleName: 'Analytic',
      font: 'Palatino Linotype',
      sizePt: 10,
      bold: true,
    },
    undertag: {
      styleId: 'Undertag',
      styleName: 'Undertag',
      font: 'Palatino Linotype',
      sizePt: 10,
      italic: true,
    },

    underline_mark: {
      styleId: 'Underline',
      styleName: 'Underline',
      sizePt: 10,
      bold: false,
      underline: 'single',
    },
    cite_mark: {
      styleId: 'Style13ptBold',
      styleName: 'Style 13 pt Bold',
      sizePt: 13,
      bold: true,
      underline: 'none',
    },
    emphasis_mark: {
      styleId: 'Emphasis',
      styleName: 'Emphasis',
      bold: true,
    },
    analytic_mark: {
      styleId: 'AnalyticChar',
      styleName: 'Analytic Char',
      bold: true,
    },
    undertag_mark: {
      styleId: 'UndertagChar',
      styleName: 'Undertag Char',
      italic: true,
    },
  },

  headerXml: null,
  footerXml: null,
  attachedTemplate: 'Lay Cut Cards.dotx',
  donorStylesXml: '',
  fontFallbacks: FONT_FALLBACKS,
};
