# @markdownkit/md-docx

Convert between Markdown and Microsoft Word (`.docx`) with a TypeScript API and CLI.

`@markdownkit/md-docx` supports:

- **Markdown → DOCX** conversion with support for:

- Headings, paragraphs, emphasis, links, and inline code
- Ordered and unordered lists (including nested lists)
- Code blocks
- Tables (including alignment markers and inline formatting)
- Images (remote URLs and data URLs)
- Comments and page breaks
- Multi-section documents with per-section style/page/header/footer/page-number settings
- Table of contents placeholder support
- **DOCX → Markdown** conversion (robust extraction via DOCX parsing + HTML-to-Markdown pipeline)

## Installation

```bash
pnpm add @markdownkit/md-docx
```

Or run directly with npx:

```bash
npx @markdownkit/md-docx input.md output.docx
```

## CLI Usage

```bash
md-docx <input.md> [output.docx] [--options <options.json>]
md-docx <input-dir> [--recursive] [--options <options.json>]
md-docx --from-docx <input.docx> [output.md] [--options <options.json>]
md-docx --from-docx <input-dir> [--recursive] [--options <options.json>]
```

### Examples

```bash
md-docx README.md README.docx
md-docx README.md
md-docx .
md-docx docs --recursive
md-docx docs -r
md-docx docs -r --options docx-options.json
md-docx README.md README.docx --options docx-options.json
md-docx --from-docx contract.docx contract.md
md-docx --from-docx docs -r
md-to-docx README.md README.docx
mtd README.md README.docx
mtd README.md
mtd .
mtd --from-docx contract.docx
dtm contract.docx
dtm .
docx-to-md docs -r
```

> `md-to-docx` is kept as a compatibility alias. `mtd` is a short alias for `md-docx`. `dtm` and `docx-to-md` default to DOCX → Markdown mode (so `dtm .` works without `--from-docx`).

### Options file

`--options` accepts JSON.

- For Markdown → DOCX, use the exported `Options` type.
- For DOCX → Markdown (`--from-docx`), use the exported `DocxToMarkdownOptions` type.

Markdown → DOCX example:

```json
{
  "documentType": "report",
  "style": {
    "fontFamily": "Inter",
    "paragraphAlignment": "JUSTIFIED",
    "heading1Alignment": "CENTER"
  },
  "template": {
    "pageNumbering": {
      "display": "current",
      "alignment": "CENTER"
    }
  },
  "sections": [
    {
      "markdown": "# Cover\n\nAcme Proposal",
      "pageNumbering": { "display": "none" }
    },
    {
      "markdown": "# Body\n\nMain content",
      "pageNumbering": { "start": 1 }
    }
  ]
}
```

DOCX → Markdown example:

```json
{
  "mammoth": {
    "preserveEmptyParagraphs": true
  },
  "turndown": {
    "headingStyle": "atx",
    "codeBlockStyle": "fenced",
    "bulletListMarker": "-"
  }
}
```

## API Usage

```ts
import { convertDocxToMarkdown, convertMarkdownToDocx } from "@markdownkit/md-docx"
import { readFile } from "node:fs/promises"

const markdown = `# Hello\n\nThis is a **DOCX** file.`
const blob = await convertMarkdownToDocx(markdown, {
  style: {
    fontFamily: "Arial",
    heading1Alignment: "CENTER"
  }
})

const docxBuffer = await readFile("./input.docx")
const fromDocx = await convertDocxToMarkdown(docxBuffer)
```

### Exported API

- `convertMarkdownToDocx(markdown, options?) => Promise<Blob>`
- `convertDocxToMarkdown(docxInput, options?) => Promise<string>`
- `parseToDocxOptions(markdown, options?) => Promise<IPropertiesOptions>`
- `downloadDocx(blob, filename?)`
- `MarkdownConversionError`
- `DocxToMarkdownError`
- Types exported from `src/types` via package entry declarations

## Notes

- Runtime target: **Node 22+**.
- Browser download helper is available via `downloadDocx`, while core conversion works in Node.
- `convertDocxToMarkdown` is Node-focused and expects DOCX binary input (`Buffer`, `Uint8Array`, or `ArrayBuffer`).
- Legacy style option `fontFamilly` is still accepted for compatibility; prefer `fontFamily`.

## Development

```bash
pnpm install
pnpm run typecheck
pnpm test
pnpm run build
```

Generate an npm tarball for inspection:

```bash
pnpm pack
```

## Scope

Current release scope is **Markdown ↔ DOCX** conversion.

## License

MIT
