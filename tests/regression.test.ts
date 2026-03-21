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
})
