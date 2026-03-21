import * as fs from 'node:fs'
import * as path from 'node:path'
import { fileURLToPath } from 'node:url'

import { describe, expect, it } from 'vitest'

import { convertMarkdownToDocx } from '../src/index'
import type { Options } from '../src/types'

// Get the directory name in ES module
const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

// Create output directory if it doesn't exist
const outputDir = path.join(__dirname, 'output')
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir)
}

describe('Markdown Newline Handling', () => {
  it('should handle various newline scenarios correctly', async () => {
    const markdown = `# Newline Test Document

## Regular Paragraphs
This is a paragraph.
This line should be part of the same paragraph.

This is a new paragraph after an empty line.

## Lists with Bold Items
- First item
**Bold continuation**
- Second item
**Another bold part**

## Code Blocks
Here's a code block:

\`\`\`javascript
function test() {
  console.log("Line 1");

  console.log("Line 3 after empty line");
  // Comment
  return true;
}
\`\`\`

## Multiple Empty Lines


This text comes after multiple empty lines.

## Blockquotes
> First line of quote
> Second line of quote

> New quote after empty line

## Mixed Content
Regular paragraph.
\`\`\`python
def function():
    # Indented line
    print("test")

    print("after empty line")
\`\`\`
Final paragraph.`

    const options: Options = {
      documentType: 'document',
      style: {
        titleSize: 32,
        headingSpacing: 240,
        paragraphSpacing: 240,
        lineSpacing: 1.15,
        paragraphAlignment: 'LEFT',
      },
    }

    // Convert markdown to docx
    const buffer = await convertMarkdownToDocx(markdown, options)

    // Save the file for manual inspection
    const outputPath = path.join(outputDir, 'newline-test.docx')
    const arrayBuffer = await buffer.arrayBuffer()
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    // Verify the buffer is not empty
    const { size } = buffer
    expect(size).toBeGreaterThan(0)
  })
})
