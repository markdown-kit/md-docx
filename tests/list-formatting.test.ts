import * as fs from 'node:fs/promises'
import * as path from 'node:path'

import { describe, expect, it } from 'vitest'

import { convertMarkdownToDocx } from '../src/index'
import { readDocumentXml, readDocxEntry } from './testUtils'

describe('List Item Formatting', () => {
  it('should handle bold text within list items', async () => {
    const markdown = `- Regular list item
- **List** item with **bold text** inside
- Another list item with **multiple** bold **words**
- List item
  **Bold text on next line**
1. Numbered list with **bold text**
2. Another numbered item with **multiple** bold **sections**
- **List** item with **bold text** inside
- **List** item with **bold text** inside`

    const docxBlob = await convertMarkdownToDocx(markdown)

    // Save the blob to a file
    const buffer = await docxBlob.arrayBuffer()
    const outputPath = path.join(process.cwd(), 'test-output')
    await fs.mkdir(outputPath, { recursive: true })
    await fs.writeFile(path.join(outputPath, 'list-formatting-test.docx'), Buffer.from(buffer))

    // Verify the blob was created successfully
    expect(docxBlob).toBeInstanceOf(Blob)
    expect(docxBlob.size).toBeGreaterThan(0)
    expect(docxBlob.type).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
  })

  it('should handle mixed formatting in list items', async () => {
    const markdown = `- List item with **bold** and *italic* text
- List item with \`code\` and **bold** formatting
- List item with **bold _italic_ text**`

    const docxBlob = await convertMarkdownToDocx(markdown)

    // Save the blob to a file
    const buffer = await docxBlob.arrayBuffer()
    const outputPath = path.join(process.cwd(), 'test-output')
    await fs.mkdir(outputPath, { recursive: true })
    await fs.writeFile(
      path.join(outputPath, 'list-formatting-mixed-test.docx'),
      Buffer.from(buffer),
    )

    // Verify the blob was created successfully
    expect(docxBlob).toBeInstanceOf(Blob)
    expect(docxBlob.size).toBeGreaterThan(0)
    expect(docxBlob.type).toBe(
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )
  })

  it('numbers nested ordered list items at their own level', async () => {
    const blob = await convertMarkdownToDocx('1. a\n   1. a1\n   2. a2\n2. b\n')
    const documentXml = await readDocumentXml(blob)
    const numberingXml = await readDocxEntry(blob, 'word/numbering.xml')

    const itemNumbering = new Map(
      [
        ...documentXml.matchAll(
          /<w:ilvl w:val="(\d+)"\/><w:numId w:val="(\d+)"\/>[\s\S]*?<w:t[^>]*>([^<]*)<\/w:t>/gu,
        ),
      ].map((match) => [match[3], { ilvl: match[1], numId: match[2] }]),
    )
    const nested = itemNumbering.get('a1')
    expect(nested?.ilvl).toBe('1')

    const abstractNumId = new RegExp(
      `<w:num w:numId="${nested?.numId}"[^>]*><w:abstractNumId w:val="(\\d+)"/>`,
      'u',
    ).exec(numberingXml)?.[1]
    const abstractNum = new RegExp(
      `<w:abstractNum [^>]*w:abstractNumId="${abstractNumId}"[^>]*>([\\s\\S]*?)</w:abstractNum>`,
      'u',
    ).exec(numberingXml)?.[1]
    const levels = [
      ...(abstractNum ?? '').matchAll(/<w:lvl w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvl>/gu),
    ]
    expect(levels.map((level) => level[1])).toEqual(['0', '1', '2', '3', '4', '5', '6', '7', '8'])

    const level1 = levels[1][2]
    expect(level1).toContain('<w:numFmt w:val="decimal"/>')
    expect(level1).toContain('<w:lvlText w:val="%2."/>')
    const indentOf = (level: string): number =>
      Number(/<w:ind w:left="(\d+)"/u.exec(level)?.[1] ?? Number.NaN)
    expect(indentOf(level1)).toBeGreaterThan(indentOf(levels[0][2]))
  })
})
