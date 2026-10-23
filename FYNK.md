# fynk fork of mammoth.js

fynk imports .docx files into its editor. It uses mammoth only to read the document: `transformDocument` captures
the document element, and the HTML output is ignored. This fork exposes the formatting fynk needs on top of upstream
mammoth. HTML and Markdown output are the same as upstream's.

## Additions

| Where | Field |
| --- | --- |
| `convertToHtml` result | `customStyles`: every style keyed by style ID: `{name, type, properties}`, formatting resolved through `w:basedOn` |
| `convertToHtml` result | `numberingStyles`: abstract numbering definitions keyed by `abstractNumId` |
| numbering level | `format` (`w:numFmt`), `delimiter` (`w:lvlText`); a paragraph's `numbering` also has `numId` |
| paragraph | `spacing {before, after, line, lineRule}`, `indent.left`, `indent.right` |
| run | `color` |
| `ins` / `del` elements | tracked changes with `authorName`, `date`, `changeId`, each wrapped in a run |
| `commentRangeStart` / `commentRangeEnd` elements, comment `date` | comment ranges |
| document | `headers`, `footers`: one element per part, in relationship order |
| table, table cell | `isBordered`; `width` (`w:tcW` / 12), `bgColor` |
| image | `width`, `height` in pixels, `imageProperties {floating, wrappingStyle}`, `isCropped` |

`lib/index.d.ts` declares these additions together with every document element mammoth produces, in the `mammoth`
namespace (`import type { Document, Paragraph } from "mammoth"`).

## Syncing with upstream

Rebase the `fynk/*` branch onto the latest upstream release tag, one commit per addition, and set the version to
`<upstream version>-fynk.<n>`. fynk pins the fork by commit hash in its `pnpm-workspace.yaml` catalog, so a fork
change reaches fynk only when that hash is bumped.
