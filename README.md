# @markdownkit/md-docx

Convert Markdown into production-quality Microsoft Word (`.docx`) documents with a TypeScript API and CLI.

`@markdownkit/md-docx` focuses on **Markdown → DOCX** conversion with support for:

- Headings, paragraphs, emphasis, links, and inline code
- Ordered and unordered lists (including nested lists)
- Code blocks
- Tables (including alignment markers and inline formatting)
- Images (remote URLs and data URLs)
- Comments and page breaks
- Multi-section documents with per-section style/page/header/footer/page-number settings
- Table of contents placeholder support

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
md-docx <input.md> <output.docx> [--options <options.json>]
md-docx <input-dir> [--recursive] [--options <options.json>]
```

### Examples

```bash
md-docx README.md README.docx
md-docx .
md-docx docs --recursive
md-docx docs -r
md-docx docs -r --options docx-options.json
md-docx README.md README.docx --options docx-options.json
md-to-docx README.md README.docx
mtd README.md README.docx
mtd .
```

> `md-to-docx` is kept as a compatibility alias. `mtd` is a short alias for `md-docx`.

### Options file

`--options` accepts JSON matching the exported `Options` type. Example:

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

## API Usage

```ts
import { convertMarkdownToDocx } from "@markdownkit/md-docx"

const markdown = `# Hello\n\nThis is a **DOCX** file.`
const blob = await convertMarkdownToDocx(markdown, {
  style: {
    fontFamily: "Arial",
    heading1Alignment: "CENTER"
  }
})
```

### Exported API

- `convertMarkdownToDocx(markdown, options?) => Promise<Blob>`
- `parseToDocxOptions(markdown, options?) => Promise<IPropertiesOptions>`
- `downloadDocx(blob, filename?)`
- `MarkdownConversionError`
- Types exported from `src/types` via package entry declarations

## Notes

- Runtime target: **Node 22+**.
- Browser download helper is available via `downloadDocx`, while core conversion works in Node.
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

Current release scope is intentionally focused on **Markdown → DOCX** conversion.

## License

MIT
