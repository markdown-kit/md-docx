import fs from 'node:fs'
import fsp from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'

import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import type { CliOutput } from '../src/cli'
import { runCli } from '../src/cli'

function captureOutput(): CliOutput & { logs: string[]; errors: string[] } {
  const logs: string[] = []
  const errors: string[] = []
  return {
    logs,
    errors,
    log: (message: string) => logs.push(message),
    error: (message: string) => errors.push(message),
  }
}

describe('standalone CLI', () => {
  let tempDir = ''

  beforeEach(() => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'md-to-docx-cli-'))
  })

  afterEach(async () => {
    if (tempDir) {
      await fsp.rm(tempDir, { recursive: true, force: true })
    }
  })

  it('converts a markdown file to docx', async () => {
    const inputPath = path.join(tempDir, 'input.md')
    const outputPath = path.join(tempDir, 'output.docx')
    const output = captureOutput()

    await fsp.writeFile(inputPath, '# CLI Test\n\nThis file should convert.')

    const exitCode = await runCli([inputPath, outputPath], output)

    expect(exitCode).toBe(0)
    expect(output.errors).toHaveLength(0)
    expect(output.logs.join('\n')).toContain('DOCX created at:')

    const stat = await fsp.stat(outputPath)
    expect(stat.size).toBeGreaterThan(0)
  })

  it('supports options loaded from JSON file', async () => {
    const inputPath = path.join(tempDir, 'input.md')
    const outputPath = path.join(tempDir, 'output-with-options.docx')
    const optionsPath = path.join(tempDir, 'options.json')
    const output = captureOutput()

    await fsp.writeFile(inputPath, '# Styled CLI Test\n\nContent.')
    await fsp.writeFile(
      optionsPath,
      JSON.stringify({
        documentType: 'report',
        style: {
          heading1Alignment: 'CENTER',
          paragraphAlignment: 'JUSTIFIED',
        },
      }),
    )

    const exitCode = await runCli([inputPath, outputPath, '--options', optionsPath], output)

    expect(exitCode).toBe(0)
    const stat = await fsp.stat(outputPath)
    expect(stat.size).toBeGreaterThan(0)
  })

  it('supports template and multi-section options from JSON file', async () => {
    const inputPath = path.join(tempDir, 'input.md')
    const outputPath = path.join(tempDir, 'output-multi-section.docx')
    const optionsPath = path.join(tempDir, 'multi-section-options.json')
    const output = captureOutput()

    await fsp.writeFile(inputPath, '# Placeholder\n\nCLI should ignore this with sections.')
    await fsp.writeFile(
      optionsPath,
      JSON.stringify({
        template: {
          page: {
            margin: {
              top: 1440,
              right: 1080,
              bottom: 1440,
              left: 1080,
            },
          },
          pageNumbering: {
            display: 'current',
            alignment: 'CENTER',
          },
        },
        sections: [
          {
            markdown: '# Cover\n\nPrepared by CLI',
            footers: {
              default: null,
            },
            pageNumbering: {
              display: 'none',
            },
            style: {
              paragraphAlignment: 'CENTER',
            },
          },
          {
            markdown: '# Body\n\nMain content starts here.\n\n1. One\n2. Two',
            headers: {
              default: {
                text: 'Main Section',
                alignment: 'RIGHT',
              },
            },
            pageNumbering: {
              start: 1,
              formatType: 'decimal',
            },
          },
        ],
      }),
    )

    const exitCode = await runCli([inputPath, outputPath, '--options', optionsPath], output)

    expect(exitCode).toBe(0)
    expect(output.errors).toHaveLength(0)
    const stat = await fsp.stat(outputPath)
    expect(stat.size).toBeGreaterThan(0)
  })

  it('supports -o short flag for options', async () => {
    const inputPath = path.join(tempDir, 'input.md')
    const outputPath = path.join(tempDir, 'output.docx')
    const optionsPath = path.join(tempDir, 'options.json')
    const output = captureOutput()

    await fsp.writeFile(inputPath, '# Short Flag\n\nContent.')
    await fsp.writeFile(optionsPath, JSON.stringify({ documentType: 'report' }))

    const exitCode = await runCli([inputPath, outputPath, '-o', optionsPath], output)

    expect(exitCode).toBe(0)
    expect(output.errors).toHaveLength(0)
  })

  it('creates nested output directories automatically', async () => {
    const inputPath = path.join(tempDir, 'input.md')
    const outputPath = path.join(tempDir, 'nested', 'deep', 'output.docx')
    const output = captureOutput()

    await fsp.writeFile(inputPath, '# Nested\n\nContent.')

    const exitCode = await runCli([inputPath, outputPath], output)

    expect(exitCode).toBe(0)
    const stat = await fsp.stat(outputPath)
    expect(stat.size).toBeGreaterThan(0)
  })

  it('returns non-zero on invalid arguments', async () => {
    const output = captureOutput()

    const exitCode = await runCli([], output)

    expect(exitCode).toBe(1)
    expect(output.errors.join('\n')).toContain('Usage:')
  })

  it('prints help text with --help and exits 0', async () => {
    const output = captureOutput()

    const exitCode = await runCli(['--help'], output)

    expect(exitCode).toBe(0)
    expect(output.logs.join('\n')).toContain('Usage:')
    expect(output.logs.join('\n')).toContain('--options')
  })

  it('prints help text with -h and exits 0', async () => {
    const output = captureOutput()

    const exitCode = await runCli(['-h'], output)

    expect(exitCode).toBe(0)
    expect(output.logs.join('\n')).toContain('Usage:')
  })

  it('fails on nonexistent input file', async () => {
    const outputPath = path.join(tempDir, 'output.docx')
    const output = captureOutput()

    const exitCode = await runCli([path.join(tempDir, 'missing.md'), outputPath], output)

    expect(exitCode).toBe(1)
    expect(output.errors.join('\n')).toMatch(/no such file|ENOENT/)
  })

  it('fails on unknown flags', async () => {
    const output = captureOutput()

    const exitCode = await runCli(['a.md', 'b.docx', '--verbose'], output)

    expect(exitCode).toBe(1)
    expect(output.errors.join('\n')).toContain('Unknown argument: --verbose')
  })

  it('fails when --options is given without a value', async () => {
    const output = captureOutput()

    const exitCode = await runCli(['a.md', 'b.docx', '--options'], output)

    expect(exitCode).toBe(1)
    expect(output.errors.join('\n')).toContain('Missing value for --options')
  })

  it('fails when options file contains invalid JSON', async () => {
    const inputPath = path.join(tempDir, 'input.md')
    const outputPath = path.join(tempDir, 'output.docx')
    const optionsPath = path.join(tempDir, 'bad.json')
    const output = captureOutput()

    await fsp.writeFile(inputPath, '# Test\n\nContent.')
    await fsp.writeFile(optionsPath, 'not valid json')

    const exitCode = await runCli([inputPath, outputPath, '--options', optionsPath], output)

    expect(exitCode).toBe(1)
    expect(output.errors.join('\n')).toContain('Invalid JSON')
  })

  it('fails when options file contains a non-object JSON value', async () => {
    const inputPath = path.join(tempDir, 'input.md')
    const outputPath = path.join(tempDir, 'output.docx')
    const optionsPath = path.join(tempDir, 'array.json')
    const output = captureOutput()

    await fsp.writeFile(inputPath, '# Test\n\nContent.')
    await fsp.writeFile(optionsPath, '["not", "an", "object"]')

    const exitCode = await runCli([inputPath, outputPath, '--options', optionsPath], output)

    expect(exitCode).toBe(1)
    expect(output.errors.join('\n')).toContain('Options JSON must be an object')
  })

  it('fails when too many positional arguments are given', async () => {
    const output = captureOutput()

    const exitCode = await runCli(['a.md', 'b.docx', 'extra.md'], output)

    expect(exitCode).toBe(1)
    expect(output.errors.join('\n')).toContain(
      'Expected either <input.md> <output.docx> or <input-dir>',
    )
  })

  it('fails when only one file positional argument is given', async () => {
    const inputPath = path.join(tempDir, 'single.md')
    const output = captureOutput()
    await fsp.writeFile(inputPath, '# Single\n\nMissing output path should fail.')

    const exitCode = await runCli([inputPath], output)

    expect(exitCode).toBe(1)
    expect(output.errors.join('\n')).toContain('Output path is required when input is a file')
  })

  it('converts markdown files from current directory input', async () => {
    const docsDir = path.join(tempDir, 'docs')
    const output = captureOutput()
    await fsp.mkdir(docsDir, { recursive: true })
    await fsp.writeFile(path.join(docsDir, 'a.md'), '# A\n\nAlpha')
    await fsp.writeFile(path.join(docsDir, 'b.markdown'), '# B\n\nBeta')

    const exitCode = await runCli([docsDir], output)

    expect(exitCode).toBe(0)
    await expect(fsp.stat(path.join(docsDir, 'a.docx'))).resolves.toBeDefined()
    await expect(fsp.stat(path.join(docsDir, 'b.docx'))).resolves.toBeDefined()
    expect(output.logs.join('\n')).toContain('Converted 2 file(s) from directory:')
  })

  it('supports recursive directory conversion with -r', async () => {
    const docsDir = path.join(tempDir, 'docs')
    const nestedDir = path.join(docsDir, 'nested')
    const output = captureOutput()
    await fsp.mkdir(nestedDir, { recursive: true })
    await fsp.writeFile(path.join(docsDir, 'root.md'), '# Root\n\nTop')
    await fsp.writeFile(path.join(nestedDir, 'child.md'), '# Child\n\nNested')

    const exitCode = await runCli([docsDir, '-r'], output)

    expect(exitCode).toBe(0)
    await expect(fsp.stat(path.join(docsDir, 'root.docx'))).resolves.toBeDefined()
    await expect(fsp.stat(path.join(nestedDir, 'child.docx'))).resolves.toBeDefined()
    expect(output.logs.join('\n')).toContain('Converted 2 file(s) from directory:')
  })
})
