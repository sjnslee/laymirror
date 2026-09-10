# Template guidelines

Templates must be a `.docx`, `.docm`, `.dotx` or `.dotm`, under 2 MB. Upload in the panel under
**template**.

## Styles

Template must have styles set for each style in Cardmirror.

| Cardmirror block   | Template style name                                                            |
| ------------------ | ------------------------------------------------------------------------------ |
| Pocket, Hat, Block | Heading 1, Heading 2, Heading 3                                                |
| Tag                | Tag, Tags, Debate Tag or Heading 4                                             |
| Cite paragraph     | Cite, Cites, Debate Cite Main, Debate Secondary Cite or NormalCite             |
| Card body          | Card, Cards, Card Text, Card (Indented), Nothing, Normal Text or Evidence Text |
| Analytic           | Analytic                                                                       |
| Undertag           | Undertag                                                                       |

| Cardmirror mark | Template character style name                                                             |
| --------------- | ----------------------------------------------------------------------------------------- |
| Underline       | Underline, Debate Underline, Debate Highlighted, Dotted Underline or Style Bold Underline |
| Cite            | Author-Date or Style Style Bold + 12 pt                                                   |
| Emphasis        | Emphasis                                                                                  |

If page breaks are desired for printing, set one of the headings to **page break before** and use
the heading at the start of each page.

## Header fields

The plugin will directly copy over all headers and footers from the template file to each page
on the output file.

However, editable fields can be inserted in the header on the template. This allows these fields in
the header to be modified within the plugin interface. These fields should be are demarcated with the
following:

Mark each editable field with a zero width character on both sides, `‹zw›` below:

```
‹zw›Team code‹zw›  →tab→  ‹zw›1AC‹zw›  →tab→  Page {PAGE} of {NUMPAGES}
```

Rows: `Team code`, `1AC`.

Marks: U+200B, U+200C, U+200D, U+2060, U+FEFF. Type `200b` alt+x, or
Insert → Symbol → Special Characters → No-Width Optional Break.

- One pair per stretch.
- Page numbers as a Word field.
- Row label and default value: the template's text.
- Unmarked paragraph: every stretch between tabs, line breaks and Word fields.
