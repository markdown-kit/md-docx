import { describe, expect, it } from 'vitest'

import { convertMarkdownToDocx, parseToDocxOptions } from '../src/index'
import { readDocumentXml, writeBlobDocx } from './testUtils'

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
})
