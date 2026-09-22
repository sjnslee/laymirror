// cardmirror's card numbering against the template's lists.

/** matches `<tag .../>` and `<tag ...>...</tag>`. the `\b` is what keeps
 *  `w:num` off `w:numbering`, `w:numFmt` and `w:numPicBullet`. */
const element = (tag: string): RegExp =>
  new RegExp(`<${tag}\\b[^>]*/>|<${tag}\\b[^>]*>[\\s\\S]*?</${tag}>`, "g");

const ABSTRACT_NUM = element("w:abstractNum");
const NUM = element("w:num");

const CLOSE = "</w:numbering>";

interface Definitions {
  abstractNums: string[];
  nums: string[];
}

const matches = (xml: string, pattern: RegExp): string[] =>
  [...xml.matchAll(pattern)].map((match) => match[0]);

const definitionsIn = (xml: string): Definitions => ({
  abstractNums: matches(xml, ABSTRACT_NUM),
  nums: matches(xml, NUM),
});

/** read off the open tag only: a `w:num` carries an `abstractNumId` child, and
 *  an `abstractNum` a whole tree of levels. */
function idOf(element: string, attribute: string): number | null {
  const open = /^<[^>]*>/.exec(element)![0];
  const raw = new RegExp(`\\b${attribute}="(\\d+)"`).exec(open)?.[1];
  return raw === undefined ? null : Number(raw);
}

const idsOf = (elements: readonly string[], attribute: string): number[] =>
  elements
    .map((el) => idOf(el, attribute))
    .filter((id): id is number => id !== null);

/** the shift that puts every one of our ids past the last of theirs. */
const shiftFor = (
  theirs: readonly number[],
  ours: readonly number[],
): number =>
  theirs.length && ours.length
    ? Math.max(...theirs) + 1 - Math.min(...ours)
    : 0;

function renumber(element: string, nums: number, abstracts: number): string {
  const bump =
    (by: number) => (whole: string, head: string, id: string, tail: string) =>
      by === 0 ? whole : `${head}${Number(id) + by}${tail}`;
  return element
    .replace(/(\bw:abstractNumId=")(\d+)(")/g, bump(abstracts))
    .replace(/(<w:abstractNumId\b[^>]*\bw:val=")(\d+)(")/g, bump(abstracts))
    .replace(/(\bw:numId=")(\d+)(")/g, bump(nums));
}

const insertAt = (xml: string, at: number, text: string): string =>
  xml.slice(0, at) + text + xml.slice(at);

/** the schema orders `numPicBullet*`, `abstractNum*`, `num*`, and word calls a
 *  file that breaks it corrupt. */
function addAbstractNums(xml: string, abstractNums: string): string {
  if (!abstractNums) return xml;
  const last = [...xml.matchAll(ABSTRACT_NUM)].pop();
  if (last) return insertAt(xml, last.index! + last[0].length, abstractNums);
  const first = [...xml.matchAll(NUM)][0];
  return insertAt(
    xml,
    first ? first.index! : xml.lastIndexOf(CLOSE),
    abstractNums,
  );
}

export interface MergedNumbering {
  /** the part to write, or null to leave whatever is already in place. */
  xml: string | null;
  /** cardmirror's numId -> the merged one, empty when nothing moved. */
  remap: Record<number, number>;
}

/** merge what cardmirror just exported into the template's own definitions. */
export function mergeNumbering(
  exportXml: string | null,
  templateXml: string | null,
): MergedNumbering {
  const nothing: MergedNumbering = { xml: null, remap: {} };
  if (!exportXml || !templateXml) return nothing;

  const ours = definitionsIn(exportXml);
  if (!ours.nums.length) return nothing;

  const theirs = definitionsIn(templateXml);
  // nothing of the template's to lose, so cardmirror's part stands as written
  if (!theirs.nums.length && !theirs.abstractNums.length)
    return { xml: exportXml, remap: {} };
  // an apply can land on a file laymirror has already applied to, whose
  // numbering.xml is the merged part: merging it again would shift twice
  if (theirs.nums.every((num) => exportXml.includes(num)))
    return { xml: exportXml, remap: {} };

  const nums = shiftFor(
    idsOf(theirs.nums, "w:numId"),
    idsOf(ours.nums, "w:numId"),
  );
  const abstracts = shiftFor(
    idsOf(theirs.abstractNums, "w:abstractNumId"),
    idsOf(ours.abstractNums, "w:abstractNumId"),
  );

  const remap: Record<number, number> = {};
  for (const id of idsOf(ours.nums, "w:numId")) remap[id] = id + nums;

  const merged = addAbstractNums(
    templateXml,
    ours.abstractNums.map((el) => renumber(el, nums, abstracts)).join(""),
  );
  const moved = ours.nums.map((el) => renumber(el, nums, abstracts)).join("");

  return { xml: insertAt(merged, merged.lastIndexOf(CLOSE), moved), remap };
}

/** point the document's list references at the ids the merge minted. */
export function remapNumIds(
  documentXml: string,
  remap: Record<number, number>,
): string {
  if (!Object.keys(remap).length) return documentXml;
  return documentXml.replace(/<w:numPr\b[^>]*>[\s\S]*?<\/w:numPr>/g, (numPr) =>
    numPr.replace(
      /(<w:numId\b[^>]*\bw:val=")(\d+)(")/g,
      (whole, head, id, tail) => {
        const next = remap[Number(id)];
        return next === undefined ? whole : `${head}${next}${tail}`;
      },
    ),
  );
}
