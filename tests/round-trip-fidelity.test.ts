import { Buffer } from 'node:buffer'

import { describe, expect, it } from 'vitest'

import { convertDocxToMarkdown } from '../src/docxToMarkdown.js'
import { convertMarkdownToDocx } from '../src/index.js'
import { readDocumentXml } from './testUtils.js'

async function md2docxXml(markdown: string): Promise<string> {
  const blob = await convertMarkdownToDocx(markdown)
  return readDocumentXml(blob)
}

async function roundTrip(markdown: string): Promise<string> {
  const blob = await convertMarkdownToDocx(markdown)
  const buffer = Buffer.from(await blob.arrayBuffer())
  return convertDocxToMarkdown(buffer)
}

describe('md-docx forward fidelity', () => {
  it('F3: a loose ordered list item does not create extra numbered markers', async () => {
    const xml = await md2docxXml('1. First\n\n   More text for first\n\n2. Second\n')
    // Two list items => exactly two numbered markers (numPr), even though item 1
    // has a continuation paragraph.
    const numPrCount = (xml.match(/<w:numPr>/g) ?? []).length
    expect(numPrCount).toBe(2)
  })

  it('F4: a table nested in a list item is preserved (not dropped)', async () => {
    const markdown = [
      '- Item with a table:',
      '',
      '  | A | B |',
      '  | - | - |',
      '  | 1 | 2 |',
      '',
    ].join('\n')
    const xml = await md2docxXml(markdown)
    expect(xml).toContain('<w:tbl>')
  })
})

describe('md-docx round-trip (md -> docx -> md)', () => {
  it('F5: strikethrough round-trips', async () => {
    const md = await roundTrip('This is ~~struck~~ text.\n')
    expect(md).toContain('~~struck~~')
  })

  it('F5: underline round-trips to ++ syntax', async () => {
    const md = await roundTrip('This is ++underlined++ text.\n')
    expect(md).toContain('++underlined++')
  })

  it('F6: inline code round-trips', async () => {
    const md = await roundTrip('Call `myFunction()` here.\n')
    expect(md).toContain('`myFunction()`')
  })
})
