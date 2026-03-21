#!/usr/bin/env node
import fsSync from 'node:fs'
import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

import { convertDocxToMarkdown, convertMarkdownToDocx } from './index.js'
import type { DocxToMarkdownOptions, Options } from './types.js'

export interface CliOutput {
  log: (message: string) => void
  error: (message: string) => void
}

interface ParsedCliArgs {
  inputPath: string
  outputPath?: string
  optionsPath?: string
  recursive: boolean
  fromDocx: boolean
}

interface HelpCliArgs {
  showHelp: true
}

type CliArgs = ParsedCliArgs | HelpCliArgs

const HELP_TEXT = `Usage:
  md-docx <input.md> <output.docx> [--options <options.json>]
  md-docx <input-dir> [--recursive] [--options <options.json>]
  md-docx --from-docx <input.docx> [output.md] [--options <options.json>]
  md-docx --from-docx <input-dir> [--recursive] [--options <options.json>]

Examples:
  md-docx a.md b.docx
  mtd a.md b.docx
  mtd .
  mtd docs --recursive
  mtd --from-docx proposal.docx
  mtd --from-docx docs -r
  dtm proposal.docx
  docx-to-md --from-docx docs -r
  md-to-docx a.md b.docx
  npx @markdownkit/md-docx a.md b.docx
  md-docx a.md b.docx --options options.json`

const DEFAULT_CLI_OUTPUT: CliOutput = {
  log: console.log,
  error: console.error,
}

function parseCliArgs(args: string[]): CliArgs {
  if (args.includes('-h') || args.includes('--help')) {
    return { showHelp: true }
  }

  const positional: string[] = []
  let optionsPath: string | undefined
  let recursive = false
  let fromDocx = false

  for (let i = 0; i < args.length; i++) {
    const arg = args[i]

    if (arg === '--options' || arg === '-o') {
      const nextArg = args[i + 1]
      if (!nextArg || nextArg.startsWith('-')) {
        throw new Error('Missing value for --options')
      }
      optionsPath = nextArg
      i++
      continue
    }

    if (arg === '--recursive' || arg === '-r') {
      recursive = true
      continue
    }

    if (arg === '--from-docx' || arg === '--docx-to-md' || arg === '-d') {
      fromDocx = true
      continue
    }

    if (arg.startsWith('-')) {
      throw new Error(`Unknown argument: ${arg}`)
    }

    positional.push(arg)
  }

  if (positional.length < 1 || positional.length > 2) {
    throw new Error(
      'Expected one input path and optional output path (examples: input.md output.docx, --from-docx input.docx output.md, or directory input)',
    )
  }

  return {
    inputPath: positional[0],
    outputPath: positional[1],
    optionsPath,
    recursive,
    fromDocx,
  }
}

function isMarkdownFile(filePath: string): boolean {
  const lower = filePath.toLowerCase()
  return lower.endsWith('.md') || lower.endsWith('.markdown')
}

function isDocxFile(filePath: string): boolean {
  const lower = filePath.toLowerCase()
  return lower.endsWith('.docx') && !path.basename(filePath).startsWith('~$')
}

async function collectFiles(
  dirPath: string,
  recursive: boolean,
  matcher: (filePath: string) => boolean,
): Promise<string[]> {
  const entries = await fs.readdir(dirPath, { withFileTypes: true })
  const files: string[] = []

  for (const entry of entries) {
    const entryPath = path.join(dirPath, entry.name)

    if (entry.isDirectory()) {
      if (recursive) {
        files.push(...(await collectFiles(entryPath, true, matcher)))
      }
      continue
    }

    if (entry.isFile() && matcher(entry.name)) {
      files.push(entryPath)
    }
  }

  return files
}

async function convertSingleMarkdownFile(
  inputPath: string,
  outputPath: string,
  options: Options | undefined,
): Promise<void> {
  const markdown = await fs.readFile(inputPath, 'utf8')
  const blob = await convertMarkdownToDocx(markdown, options)
  const arrayBuffer = await blob.arrayBuffer()

  await fs.mkdir(path.dirname(outputPath), { recursive: true })
  await fs.writeFile(outputPath, Buffer.from(arrayBuffer))
}

async function convertSingleDocxFile(
  inputPath: string,
  outputPath: string,
  options: DocxToMarkdownOptions | undefined,
): Promise<void> {
  const docxBuffer = await fs.readFile(inputPath)
  const markdown = await convertDocxToMarkdown(docxBuffer, options)

  await fs.mkdir(path.dirname(outputPath), { recursive: true })
  await fs.writeFile(outputPath, markdown, 'utf8')
}

async function readOptionsFile(optionsPath: string): Promise<Record<string, unknown>> {
  const content = await fs.readFile(optionsPath, 'utf8')

  try {
    const parsed: unknown = JSON.parse(content)

    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error('Options JSON must be an object')
    }

    return parsed as Record<string, unknown>
  } catch (err) {
    if (err instanceof SyntaxError) {
      throw new TypeError(`Invalid JSON in options file "${optionsPath}": ${err.message}`)
    }
    throw err
  }
}

export async function runCli(
  args: string[],
  output: CliOutput = DEFAULT_CLI_OUTPUT,
): Promise<number> {
  try {
    const parsedArgs = parseCliArgs(args)
    if ('showHelp' in parsedArgs) {
      output.log(HELP_TEXT)
      return 0
    }

    const inputPath = path.resolve(parsedArgs.inputPath)
    const optionsObject = parsedArgs.optionsPath
      ? await readOptionsFile(path.resolve(parsedArgs.optionsPath))
      : undefined
    const inputStat = await fs.stat(inputPath)

    if (inputStat.isFile()) {
      const inputIsDocx = isDocxFile(inputPath)
      const shouldConvertFromDocx = parsedArgs.fromDocx || inputIsDocx

      if (shouldConvertFromDocx) {
        if (parsedArgs.fromDocx && !inputIsDocx) {
          throw new Error('When --from-docx is set, input file must have .docx extension')
        }

        const outputPath = parsedArgs.outputPath
          ? path.resolve(parsedArgs.outputPath)
          : inputPath.replace(/\.docx$/i, '.md')
        await convertSingleDocxFile(
          inputPath,
          outputPath,
          optionsObject as DocxToMarkdownOptions | undefined,
        )
        output.log(`Markdown created at: ${outputPath}`)
        return 0
      }

      if (!parsedArgs.outputPath) {
        throw new Error('Output path is required when input is a markdown file')
      }

      const outputPath = path.resolve(parsedArgs.outputPath)
      await convertSingleMarkdownFile(inputPath, outputPath, optionsObject as Options | undefined)
      output.log(`DOCX created at: ${outputPath}`)
      return 0
    }

    if (!inputStat.isDirectory()) {
      throw new Error(`Input path is neither a file nor directory: ${parsedArgs.inputPath}`)
    }

    if (parsedArgs.outputPath) {
      throw new Error('Output path must not be provided when input is a directory')
    }

    if (parsedArgs.fromDocx) {
      const docxFiles = await collectFiles(inputPath, parsedArgs.recursive, isDocxFile)
      if (docxFiles.length === 0) {
        throw new Error(
          `No DOCX files found in directory: ${parsedArgs.inputPath}. ` +
            `Supported input extension: .docx${parsedArgs.recursive ? '' : '. ' +
              'If your files are in subfolders, run with -r/--recursive'}`,
        )
      }

      let convertedCount = 0
      for (const docxFilePath of docxFiles) {
        const outputPath = docxFilePath.replace(/\.docx$/i, '.md')
        await convertSingleDocxFile(
          docxFilePath,
          outputPath,
          optionsObject as DocxToMarkdownOptions | undefined,
        )
        convertedCount++
        output.log(`Markdown created at: ${outputPath}`)
      }

      output.log(`Converted ${convertedCount} file(s) from DOCX to Markdown in directory: ${inputPath}`)
      return 0
    }

    const markdownFiles = await collectFiles(inputPath, parsedArgs.recursive, isMarkdownFile)
    if (markdownFiles.length === 0) {
      throw new Error(
        `No markdown files found in directory: ${parsedArgs.inputPath}. ` +
          `Supported input extensions: .md, .markdown. ` +
          `${parsedArgs.recursive ? '' : 'If your files are in subfolders, run with -r/--recursive. '}` +
          `For DOCX-to-Markdown conversion, use --from-docx.`,
      )
    }

    let convertedCount = 0
    for (const markdownFilePath of markdownFiles) {
      const outputPath = markdownFilePath.replace(/\.(md|markdown)$/i, '.docx')
      await convertSingleMarkdownFile(markdownFilePath, outputPath, optionsObject as Options | undefined)
      convertedCount++
      output.log(`DOCX created at: ${outputPath}`)
    }

    output.log(`Converted ${convertedCount} file(s) from directory: ${inputPath}`)
    return 0
  } catch (err) {
    const message =
      err instanceof Error
        ? err.message
        : typeof err === 'object' && err !== null && 'message' in err
          ? String((err as { message: unknown }).message)
          : String(err)
    output.error(`Error: ${message}`)
    output.error('')
    output.error(HELP_TEXT)
    return 1
  }
}

const currentFilePath = fileURLToPath(import.meta.url)
function resolveRealPath(filePath: string): string {
  try {
    return fsSync.realpathSync(filePath)
  } catch {
    return path.resolve(filePath)
  }
}

const invokedFilePath = process.argv[1] ? resolveRealPath(process.argv[1]) : ''
const currentRealFilePath = resolveRealPath(currentFilePath)

if (invokedFilePath === currentRealFilePath) {
  void runCli(process.argv.slice(2)).then((exitCode) => {
    process.exitCode = exitCode
  })
}
