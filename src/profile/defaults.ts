// the built-in fallback profile: plain lay conventions, deliberately generic.
// it is not any school's format. a school's template docx is the real profile
// — load one through the panel and every value below is replaced by theirs.

import type { Profile } from './profile.js';

/** cardmirror ships metric-compatible substitutes for cambria (caladea),
 *  calibri (carlito) and times new roman (tinos). palatino linotype and
 *  garamond have none, so a template calling for them will paginate a little
 *  differently here than in word. */
export const FONT_FALLBACKS: Record<string, string> = {
  'Palatino Linotype': '"Palatino Linotype", Palatino, "Book Antiqua", serif',
  'Times New Roman': 'Tinos, "Times New Roman", serif',
  Cambria: 'Caladea, Cambria, serif',
  Calibri: 'Carlito, Calibri, sans-serif',
  Garamond: 'Garamond, "EB Garamond", serif',
};

const SERIF = 'Times New Roman';

export const DEFAULT_LAY: Profile = {
  id: 'default',
  name: 'generic lay',

  // us letter, one inch all round
  page: {
    widthTwips: 12240,
    heightTwips: 15840,
    margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 },
  },

  types: {
    paragraph: { styleId: 'Normal', styleName: 'Normal', font: SERIF, sizePt: 12 },
    pocket: {
      styleId: 'Heading1',
      styleName: 'heading 1',
      font: SERIF,
      sizePt: 16,
      bold: true,
      smallCaps: true,
      align: 'center',
      spaceAfterPt: 6,
      pageBreakBefore: true,
      outlineLevel: 0,
    },
    hat: {
      styleId: 'Heading2',
      styleName: 'heading 2',
      font: SERIF,
      sizePt: 14,
      bold: true,
      spaceBeforePt: 10,
      keepNext: true,
      keepLines: true,
      outlineLevel: 1,
    },
    block: {
      styleId: 'Heading3',
      styleName: 'heading 3',
      font: SERIF,
      sizePt: 13,
      bold: true,
      spaceBeforePt: 8,
      keepNext: true,
      keepLines: true,
      outlineLevel: 2,
    },
    tag: { styleId: 'Tag', styleName: 'Tag', font: SERIF, sizePt: 12, bold: true, keepNext: true },
    cite_paragraph: {
      styleId: 'Cite',
      styleName: 'Cite',
      font: SERIF,
      sizePt: 12,
      bold: true,
      underline: 'single',
    },
    card_body: {
      styleId: 'card',
      styleName: 'card',
      font: SERIF,
      sizePt: 12,
      indentLeftDxa: 360,
      indentRightDxa: 360,
      spaceAfterPt: 8,
      lineSpacing: { rule: 'auto', value: 240 },
    },
    analytic: { styleId: 'Analytic', styleName: 'Analytic', font: SERIF, sizePt: 12, bold: true },
    undertag: { styleId: 'Undertag', styleName: 'Undertag', font: SERIF, sizePt: 12, italic: true },

    underline_mark: {
      styleId: 'Underline',
      styleName: 'Underline',
      sizePt: 12,
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
    emphasis_mark: { styleId: 'Emphasis', styleName: 'Emphasis', bold: true },
    analytic_mark: { styleId: 'AnalyticChar', styleName: 'Analytic Char', bold: true },
    undertag_mark: { styleId: 'UndertagChar', styleName: 'Undertag Char', italic: true },
  },

  headerXml: null,
  footerXml: null,
  attachedTemplate: null,
  donorStylesXml: '',
  fontFallbacks: FONT_FALLBACKS,
};
