import * as fs from 'node:fs'
import * as path from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import { convertMarkdownToDocx } from '../src/index'
import type { Options } from '../src/types'

const __dirname = path.dirname(fileURLToPath(import.meta.url))
const outputDir = path.join(__dirname, '..', 'test-output')

// Create output directory if it doesn't exist
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir)
}

describe('Text Alignment Tests', () => {
  it('should apply different alignments to various elements', async () => {
    const markdown = `# Centered Title

## Right-Aligned Subtitle

This is a justified paragraph that demonstrates how text can be spread evenly across the width of the page. This creates a clean, professional look with straight edges on both the left and right margins.

> This is a centered blockquote that stands out from the regular text.

This is a left-aligned paragraph (default alignment) that shows the standard text positioning. It's easy to read and follows traditional document formatting.

### Center-Aligned Heading

This is another justified paragraph to show consistency in formatting. When you have longer paragraphs of text, justified alignment can make them look more organized and professional in formal documents.`

    const options: Options = {
      documentType: 'document',
      style: {
        // Font sizes
        titleSize: 32,
        paragraphSize: 24,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,

        // Alignment settings
        paragraphAlignment: 'JUSTIFIED',
        blockquoteAlignment: 'CENTER',
      },
    }

    const buffer = await convertMarkdownToDocx(markdown, options)

    // Save the file for manual inspection
    const outputPath = path.join(outputDir, 'text-alignment-test.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    // Verify the buffer is not empty
    expect(buffer).toBeInstanceOf(Blob)
    expect(buffer.size).toBeGreaterThan(0)
  })

  it('should handle mixed paragraph alignments', async () => {
    const markdown = `This is a left-aligned paragraph.

This is a justified paragraph with enough text to demonstrate the justification. It should spread across the width of the page evenly, creating straight edges on both sides.

This is a center-aligned paragraph.

This is a right-aligned paragraph.`

    const mixedOptions: Options = {
      documentType: 'document',
      style: {
        titleSize: 32,
        paragraphSize: 24,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
        // Test each paragraph with different alignment
        paragraphAlignment: 'LEFT', // This should be the default
      },
    }

    const buffer = await convertMarkdownToDocx(markdown, mixedOptions)
    const outputPath = path.join(outputDir, 'mixed-alignment-test.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    expect(buffer).toBeInstanceOf(Blob)
    expect(buffer.size).toBeGreaterThan(0)
  })

  it('should respect heading alignment configuration', async () => {
    const markdown = `# Level 1 Heading (Centered by default)

## Level 2 Heading (Right-aligned by config)

### Level 3 Heading (Left-aligned by default)`

    const headingOptions: Options = {
      documentType: 'document',
      style: {
        titleSize: 32,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
        // No explicit headingAlignment here as we're testing the default configs
      },
    }

    const buffer = await convertMarkdownToDocx(markdown, headingOptions)
    const outputPath = path.join(outputDir, 'heading-alignment-test.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    expect(buffer).toBeInstanceOf(Blob)
    expect(buffer.size).toBeGreaterThan(0)
  })

  it('should handle blockquote alignments', async () => {
    const markdown = `> This is a centered blockquote.

Regular paragraph here.

> This is another blockquote with center alignment.`

    const blockquoteOptions: Options = {
      documentType: 'document',
      style: {
        titleSize: 32,
        paragraphSize: 24,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
        blockquoteAlignment: 'CENTER',
      },
    }

    const buffer = await convertMarkdownToDocx(markdown, blockquoteOptions)
    const outputPath = path.join(outputDir, 'blockquote-alignment-test.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    expect(buffer).toBeInstanceOf(Blob)
    expect(buffer.size).toBeGreaterThan(0)
  })
})
