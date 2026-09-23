import { Document, Packer } from 'docx'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { convertMarkdownToDocx, parseToDocxOptions } from '../src/index'
import { modelToDocx } from '../src/modelToDocx'
import type { RunInfo } from './testUtils'
import { readDocumentXml, readDocxEntry, readRuns, writeBlobDocx } from './testUtils'

describe('regression coverage', () => {
  it('clamps generated heading style sizes to valid DOCX ranges', async () => {
    const options = await parseToDocxOptions('# Heading 1\n\n## Heading 2\n\n### Heading 3', {
      style: {
        titleSize: 8,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
      },
    })

    const paragraphStyles = options.styles?.paragraphStyles ?? []
    const headingStyles = paragraphStyles.filter((style) =>
      ['Title', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5'].includes(style.id),
    )
    const headingSizes = headingStyles
      .map((style) => style.run?.size)
      .filter((size): size is number => typeof size === 'number')

    expect(headingSizes.length).toBeGreaterThan(0)
    expect(headingSizes.every((size) => size >= 8)).toBe(true)
  })

  it('renders COMMENT paragraph markers and TOC placeholders', async () => {
    const markdown = `[TOC]

# Primary Heading

COMMENT: reviewer note`

    const blob = await convertMarkdownToDocx(markdown)
    await writeBlobDocx(blob, 'regression-comment-toc.docx')

    const xml = await readDocumentXml(blob)
    expect(xml).toContain('Table of Contents')
    expect(xml).toContain('Comment: reviewer note')
    expect(xml).toContain('Primary Heading')
  })

  it('renders inline <sup>/<sub> tags as superscript and subscript runs', async () => {
    const blob = await convertMarkdownToDocx(
      'Water is H<sub>2</sub>O and the 2<sup>nd</sup> **bold<sup>x</sup>** clause.\n\n> Quote E = mc<sup>2</sup>\n',
    )
    const xml = await readDocumentXml(blob)
    expect(xml).toMatch(/<w:vertAlign w:val="subscript"\/>[\s\S]{0,120}<w:t[^>]*>2<\/w:t>/u)
    expect(xml).toMatch(/<w:vertAlign w:val="superscript"\/>[\s\S]{0,120}<w:t[^>]*>nd<\/w:t>/u)
    expect(xml).toMatch(
      /<w:b\/>[\s\S]{0,220}<w:vertAlign w:val="superscript"\/>[\s\S]{0,120}<w:t[^>]*>x<\/w:t>/u,
    )
    expect(xml).toMatch(/<w:vertAlign w:val="superscript"\/>[\s\S]{0,160}<w:t[^>]*>2<\/w:t>/u)
    expect(xml).not.toMatch(/&lt;sup&gt;|<sup>/u)
  })

  it('sizes list item text with listItemSize rather than paragraphSize', async () => {
    const blob = await convertMarkdownToDocx('Body text\n\n- Bullet item\n\n1. Numbered item\n', {
      style: { paragraphSize: 24, listItemSize: 18 },
    })
    const xml = await readDocumentXml(blob)
    const runSize = (text: string): string | undefined =>
      xml
        .split('<w:r>')
        .find((run) => run.includes(`>${text}<`))
        ?.match(/<w:sz w:val="(\d+)"\/>/u)?.[1]
    expect(runSize('Body text')).toBe('24')
    expect(runSize('Bullet item')).toBe('18')
    expect(runSize('Numbered item')).toBe('18')
  })

  describe('literal markup characters in text stay literal', () => {
    const plain = (text: string, flags: Partial<RunInfo> = {}): RunInfo => ({
      text,
      bold: false,
      italic: false,
      underline: false,
      strike: false,
      code: false,
      hyperlink: false,
      ...flags,
    })

    /** Merge adjacent runs with identical formatting so run splits don't matter. */
    function mergedRuns(runs: RunInfo[]): RunInfo[] {
      const merged: RunInfo[] = []
      for (const run of runs) {
        const previous = merged.at(-1)
        const { text: _text, ...flags } = run
        if (previous) {
          const { text: _previousText, ...previousFlags } = previous
          if (JSON.stringify(flags) === JSON.stringify(previousFlags)) {
            previous.text += run.text
            continue
          }
        }
        merged.push({ ...run })
      }
      return merged
    }

    const cases: Array<{ name: string; markdown: string; expected: RunInfo[] }> = [
      {
        name: 'spaced asterisks',
        markdown: 'Compute 2 * 3 * 4 now.',
        expected: [plain('Compute 2 * 3 * 4 now.')],
      },
      {
        name: 'escaped asterisks',
        markdown: 'Use \\*literal asterisks\\* here.',
        expected: [plain('Use *literal asterisks* here.')],
      },
      {
        name: 'C++ twice',
        markdown: 'I like C++ and C++ too.',
        expected: [plain('I like C++ and C++ too.')],
      },
      {
        name: 'escaped tildes',
        markdown: 'Keep \\~\\~not struck\\~\\~ here.',
        expected: [plain('Keep ~~not struck~~ here.')],
      },
      {
        name: 'escaped backticks',
        markdown: 'a \\` b \\` c',
        expected: [plain('a ` b ` c')],
      },
      {
        name: 'backslash inside inline code',
        markdown: 'Run `a\\*b` now',
        expected: [plain('Run '), plain('a\\*b', { code: true }), plain(' now')],
      },
      {
        name: 'link URL containing a parenthesis',
        markdown: 'Wiki [Rust](https://en.wikipedia.org/wiki/Rust_(programming_language)) end.',
        expected: [
          plain('Wiki '),
          plain('Rust', { hyperlink: true, underline: true }),
          plain(' end.'),
        ],
      },
      {
        name: 'link text containing markup characters',
        markdown: 'See [2*3 [x]](https://x.y) ok',
        expected: [
          plain('See '),
          plain('2*3 [x]', { hyperlink: true, underline: true }),
          plain(' ok'),
        ],
      },
      {
        name: 'underline syntax still applies',
        markdown: '++underline++ done',
        expected: [plain('underline', { underline: true }), plain(' done')],
      },
      {
        name: 'regular formatting still applies',
        markdown: '**b** *i* ~~s~~',
        expected: [
          plain('b', { bold: true }),
          plain(' '),
          plain('i', { italic: true }),
          plain(' '),
          plain('s', { strike: true }),
        ],
      },
      {
        name: 'heading with literal markers',
        markdown: '## Price \\*not italic\\* and C++ or C++',
        expected: [plain('Price *not italic* and C++ or C++')],
      },
      {
        name: 'table cell with spaced asterisks',
        markdown: '| h |\n| --- |\n| 2 * 3 * 4 |',
        expected: [plain('h', { bold: true }), plain('2 * 3 * 4')],
      },
    ]

    it.each(cases)('$name', async ({ markdown, expected }) => {
      const xml = await readDocumentXml(await convertMarkdownToDocx(markdown))
      expect(mergedRuns(readRuns(xml))).toEqual(expected)
    })

    it('keeps the full link URL when it contains a parenthesis', async () => {
      const blob = await convertMarkdownToDocx(
        'Wiki [Rust](https://en.wikipedia.org/wiki/Rust_(programming_language)) end.',
      )
      const rels = await readDocxEntry(blob, 'word/_rels/document.xml.rels')
      expect(rels).toContain('Target="https://en.wikipedia.org/wiki/Rust_(programming_language%29"')
    })
  })

  describe('hard line breaks', () => {
    /** Body paragraphs of document.xml, in order. */
    function bodyParagraphs(xml: string): string[] {
      return [...(xml.split('<w:body>')[1] ?? '').matchAll(/<w:p>[\s\S]*?<\/w:p>/gu)].map(
        (match) => match[0],
      )
    }

    function expectBreakBetween(paragraph: string, before: string, after: string): void {
      expect(paragraph).toMatch(
        new RegExp(`>${before}</w:t>[\\s\\S]*<w:br/>[\\s\\S]*>${after}</w:t>`, 'u'),
      )
      expect(paragraph).not.toMatch(/<w:t[^>]*>[^<]*\n[^<]*<\/w:t>/u)
    }

    it('renders a hard break in a paragraph as a Word line break', async () => {
      const xml = await readDocumentXml(await convertMarkdownToDocx('line one  \nline two'))
      const paragraphs = bodyParagraphs(xml)
      expect(paragraphs).toHaveLength(1)
      expectBreakBetween(paragraphs[0], 'line one', 'line two')
    })

    it('keeps formatting across a hard break', async () => {
      const xml = await readDocumentXml(await convertMarkdownToDocx('**bold  \nstill bold**'))
      expect(readRuns(xml).map((run) => [run.text, run.bold])).toEqual([
        ['bold', true],
        ['still bold', true],
      ])
      expect(xml).toContain('<w:br/>')
    })

    it('renders a hard break in a list item as a Word line break', async () => {
      const xml = await readDocumentXml(await convertMarkdownToDocx('- item one\\\n  continued\n'))
      const paragraphs = bodyParagraphs(xml)
      expect(paragraphs).toHaveLength(1)
      expect(paragraphs[0]).toContain('<w:numPr>')
      expectBreakBetween(paragraphs[0], 'item one', 'continued')
    })

    it('renders a hard break in a table cell as a Word line break', async () => {
      // GFM rows are single lines, so a break node in a cell only arrives via the model.
      const { children } = await modelToDocx(
        {
          children: [
            {
              type: 'table',
              headers: [[{ type: 'text', value: 'h' }]],
              rows: [
                [
                  [
                    { type: 'text', value: 'cell one' },
                    { type: 'text', value: '\n', lineBreak: true },
                    { type: 'text', value: 'cell two' },
                  ],
                ],
              ],
            },
          ],
        },
        { titleSize: 32, headingSpacing: 240, paragraphSpacing: 240, lineSpacing: 1.15 },
        {},
      )
      const buffer = await Packer.toBuffer(new Document({ sections: [{ children }] }))
      const xml = await (await JSZip.loadAsync(buffer)).file('word/document.xml')!.async('string')
      const cellParagraph = bodyParagraphs(xml).find((paragraph) => paragraph.includes('cell one'))
      expectBreakBetween(cellParagraph ?? '', 'cell one', 'cell two')
    })

    it('renders a hard break in a blockquote without a stray backslash', async () => {
      const xml = await readDocumentXml(await convertMarkdownToDocx('> quote one  \n> quote two'))
      expectBreakBetween(bodyParagraphs(xml)[0], 'quote one', 'quote two')
      expect(xml).not.toContain('\\')
    })

    it('keeps a soft line break inside the same line', async () => {
      const xml = await readDocumentXml(await convertMarkdownToDocx('line one\nline two'))
      expect(xml).not.toContain('<w:br/>')
    })
  })

  it('maps a level-6 heading to the Heading6 style', async () => {
    const xml = await readDocumentXml(await convertMarkdownToDocx('###### Six'))
    const styles = [...xml.matchAll(/<w:pStyle w:val="([^"]+)"\/>/gu)].map((match) => match[1])
    expect(new Set(styles)).toEqual(new Set(['Heading6']))
  })
})
