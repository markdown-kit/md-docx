import { describe, expect, it } from 'vitest'

import { convertMarkdownToDocx, parseToDocxOptions } from '../src/index'
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
})
