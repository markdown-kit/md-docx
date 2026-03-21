import * as fs from 'node:fs'
import * as path from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import { convertMarkdownToDocx } from '../src/index'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const outputDir = path.join(__dirname, '..', 'test-output')

if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir)
}

const INLINE_IMAGE_DATA_URL =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAGgwJ/vk9yBgAAAABJRU5ErkJggg=='

describe('image sizing', () => {
  it('preserves aspect ratio when only width is provided', async () => {
    const markdown = `![Wide Banner](${INLINE_IMAGE_DATA_URL}#w=400)`

    const blob = await convertMarkdownToDocx(markdown)
    expect(blob).toBeInstanceOf(Blob)
    expect(blob.size).toBeGreaterThan(0)

    const outputPath = path.join(outputDir, 'image-size-width-only.docx')
    const arrayBuffer = await blob.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))
  })

  it('uses explicit width and height when both provided', async () => {
    const markdown = `![Exact Size](${INLINE_IMAGE_DATA_URL}#w=120&h=60)`

    const blob = await convertMarkdownToDocx(markdown)
    expect(blob).toBeInstanceOf(Blob)
    expect(blob.size).toBeGreaterThan(0)

    const outputPath = path.join(outputDir, 'image-size-both.docx')
    const arrayBuffer = await blob.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))
  })

  it('supports data URLs with explicit size', async () => {
    // 1x1 PNG (transparent)
    const onePx =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAGgwJ/vk9yBgAAAABJRU5ErkJggg==#w=50&h=50'
    const markdown = `![one px](${onePx})`

    const blob = await convertMarkdownToDocx(markdown)
    expect(blob).toBeInstanceOf(Blob)
    expect(blob.size).toBeGreaterThan(0)

    const outputPath = path.join(outputDir, 'image-size-dataurl.docx')
    const arrayBuffer = await blob.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))
  })
})
