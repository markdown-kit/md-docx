import * as fs from 'node:fs'
import * as path from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import { convertMarkdownToDocx } from '../src/index'
import type { Options } from '../src/index'

// Get the directory path using import.meta.url
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// Create output directory if it doesn't exist
const outputDir = path.join(__dirname, '..', 'test-output')
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true })
}

describe('Heading Alignment Tests', () => {
  it('should apply default left alignment to all headings', async () => {
    const markdown = `# Heading 1
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5`

    const options: Options = {
      documentType: 'document',
      style: {
        titleSize: 32,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
        // No explicit heading alignments - should use defaults (all LEFT)
      },
    }

    const buffer = await convertMarkdownToDocx(markdown, options)
    const outputPath = path.join(outputDir, 'default-heading-alignment.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    expect(buffer).toBeInstanceOf(Blob)
    expect(buffer.size).toBeGreaterThan(0)
  })

  it('should apply custom alignments to individual heading levels', async () => {
    const markdown = `# Heading 1 (should be centered)
## Heading 2 (should be right-aligned)
### Heading 3 (should be justified)
#### Heading 4 (should be left-aligned)
##### Heading 5 (should be centered)`

    const options: Options = {
      documentType: 'document',
      style: {
        titleSize: 32,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
        // Custom alignments for each heading level
        heading1Alignment: 'CENTER',
        heading2Alignment: 'RIGHT',
        heading3Alignment: 'JUSTIFIED',
        heading4Alignment: 'LEFT',
        heading5Alignment: 'CENTER',
      },
    }

    const buffer = await convertMarkdownToDocx(markdown, options)
    const outputPath = path.join(outputDir, 'custom-heading-alignment.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    expect(buffer).toBeInstanceOf(Blob)
    expect(buffer.size).toBeGreaterThan(0)
  })

  it('should fall back to general headingAlignment when level-specific not provided', async () => {
    const markdown = `# Heading 1
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5`

    const options: Options = {
      documentType: 'document',
      style: {
        titleSize: 32,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
        // Only general heading alignment, no level-specific ones
        headingAlignment: 'CENTER',
      },
    }

    const buffer = await convertMarkdownToDocx(markdown, options)
    const outputPath = path.join(outputDir, 'fallback-heading-alignment.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    expect(buffer).toBeInstanceOf(Blob)
    expect(buffer.size).toBeGreaterThan(0)
  })

  it('should prioritize level-specific alignment over general headingAlignment', async () => {
    const markdown = `# Heading 1 (should be centered)
## Heading 2 (should be right-aligned)
### Heading 3 (should be justified)
#### Heading 4 (should be left-aligned)
##### Heading 5 (should be centered)`

    const options: Options = {
      documentType: 'document',
      style: {
        titleSize: 32,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
        // Both level-specific and general heading alignment
        headingAlignment: 'LEFT', // This should be ignored for levels with specific alignments
        heading1Alignment: 'CENTER',
        heading2Alignment: 'RIGHT',
        heading3Alignment: 'JUSTIFIED',
        heading4Alignment: 'LEFT',
        heading5Alignment: 'CENTER',
      },
    }

    const buffer = await convertMarkdownToDocx(markdown, options)
    const outputPath = path.join(outputDir, 'priority-heading-alignment.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    expect(buffer).toBeInstanceOf(Blob)
    expect(buffer.size).toBeGreaterThan(0)
  })

  it('should handle mixed heading alignments with other document elements', async () => {
    const markdown = `# Centered Heading 1

## Right-aligned Heading 2

This is a justified paragraph that demonstrates how text can be spread evenly across the width of the page.

### Justified Heading 3

> This is a centered blockquote that stands out from the regular text.

#### Left-aligned Heading 4

This is a left-aligned paragraph (default alignment) that shows the standard text positioning.

##### Centered Heading 5`

    const options: Options = {
      documentType: 'document',
      style: {
        titleSize: 32,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
        // Mixed alignments
        paragraphAlignment: 'JUSTIFIED',
        blockquoteAlignment: 'CENTER',
        heading1Alignment: 'CENTER',
        heading2Alignment: 'RIGHT',
        heading3Alignment: 'JUSTIFIED',
        heading4Alignment: 'LEFT',
        heading5Alignment: 'CENTER',
      },
    }

    const buffer = await convertMarkdownToDocx(markdown, options)
    const outputPath = path.join(outputDir, 'mixed-alignment.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    expect(buffer).toBeInstanceOf(Blob)
    expect(buffer.size).toBeGreaterThan(0)
  })
})
