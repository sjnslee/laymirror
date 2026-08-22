// a profile is the school's document identity plus the vocabulary bridge to
// it. nothing here describes typography: the template's own styles.xml and
// theme travel verbatim inside the snapshot, so word sees the school's styles
// rather than our reconstruction of them.

import type { Snapshot } from '../docx/snapshot.js';

export interface StyleInfo {
  id: string;
  name: string;
  kind: 'paragraph' | 'character' | 'table' | 'numbering';
}

export interface Profile {
  id: string;
  name: string;
  /** the parts that make a document the school's, carried byte for byte. */
  snapshot: Snapshot | null;
  /** cardmirror's exported style id -> the id this template defines.
   *  an id absent from the map is left as cardmirror wrote it. */
  styleMap: Record<string, string>;
  /** cite paragraphs and card bodies leave cardmirror with no `w:pStyle` at
   *  all, so they cannot be remapped by id — they are recognised from the
   *  marks their runs carry and given these ids instead. null leaves them as
   *  the template's Normal. */
  bareStyles: {
    cite_paragraph: string | null;
    card_body: string | null;
  };
  /** every style the template defines, so the panel can offer real choices
   *  instead of asking the user to type an id. */
  styles: StyleInfo[];
}

export const isProfile = (value: unknown): value is Profile =>
  !!value &&
  typeof value === 'object' &&
  typeof (value as Profile).id === 'string' &&
  typeof (value as Profile).name === 'string';
