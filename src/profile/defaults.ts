// the profile in force before a school template has been loaded.
//
// it carries nothing and maps nothing, deliberately. laymirror's whole value
// is the school's own template, and inventing a house style to stand in for
// one would put a document on a judge's desk in a format no school chose.
// until a template is loaded, marking a document lay changes what is on
// screen and leaves the file alone.

import type { Profile } from './profile.js';

export const DEFAULT_PROFILE: Profile = {
  id: 'none',
  name: 'no template loaded',
  snapshot: null,
  styleMap: {},
  bareStyles: { cite_paragraph: null, card_body: null },
  styles: [],
};

export const hasTemplate = (profile: Profile): boolean => profile.snapshot !== null;
