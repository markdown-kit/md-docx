import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { convertDocxToMarkdown, convertMarkdownToDocx } from '../src/index'

describe('convertDocxToMarkdown', () => {
  let tempDir = ''

  beforeEach(async () => {
    tempDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docx-to-markdown-'))
  })

  afterEach(async () => {
    if (tempDir) {
      await fs.rm(tempDir, { recursive: true, force: true })
    }
  })

  it('converts DOCX binary input back to markdown text', async () => {
    const sourceMarkdown = '# Roundtrip Title\n\nA paragraph with **bold** text.\n\n- one\n- two\n'
    const docxBlob = await convertMarkdownToDocx(sourceMarkdown)
    const docxBuffer = Buffer.from(await docxBlob.arrayBuffer())

    const markdown = await convertDocxToMarkdown(docxBuffer)

    expect(markdown).toContain('Roundtrip Title')
    expect(markdown).toContain('**bold**')
    expect(markdown).toContain('one')
    expect(markdown).toContain('two')
  })

  it('removes bold-only wrappers from standalone heading-like lines', async () => {
    const sourceMarkdown = '# **CAPYDB**\n\n## **Technical Specification**\n\nParagraph.'
    const docxBlob = await convertMarkdownToDocx(sourceMarkdown)
    const docxBuffer = Buffer.from(await docxBlob.arrayBuffer())

    const markdown = await convertDocxToMarkdown(docxBuffer)

    expect(markdown).toContain('CAPYDB')
    expect(markdown).toContain('Technical Specification')
    expect(markdown).not.toContain('**CAPYDB**')
    expect(markdown).not.toContain('**Technical Specification**')
  })

  it('converts table content without leaking raw html table tags', async () => {
    const sourceMarkdown =
      '# Table Spec\n\n| Pillar | Responsibility |\n| --- | --- |\n| Pillar A | Storage |\n| Pillar B | Proxy |\n'
    const docxBlob = await convertMarkdownToDocx(sourceMarkdown)
    const docxBuffer = Buffer.from(await docxBlob.arrayBuffer())

    const markdown = await convertDocxToMarkdown(docxBuffer)

    expect(markdown).toContain('|')
    expect(markdown).toContain('Pillar')
    expect(markdown).toContain('Responsibility')
    expect(markdown).not.toMatch(/<\/?table/i)
    expect(markdown).not.toMatch(/<\/?tbody/i)
    expect(markdown).not.toMatch(/<\/?tr/i)
  })

  it('throws on empty DOCX content', async () => {
    await expect(convertDocxToMarkdown(Buffer.alloc(0))).rejects.toThrow(
      'Invalid DOCX input: file content is empty',
    )
  })
})
