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

const DEFAULT_CLI_OUTPUT: CliOutput = {
  log: console.log,
  error: console.error,
}

interface CliTheme {
  colorEnabled: boolean
}

const ANSI = {
  reset: '\u001B[0m',
  bold: '\u001B[1m',
  dim: '\u001B[2m',
  cyan: '\u001B[36m',
  green: '\u001B[32m',
  yellow: '\u001B[33m',
  red: '\u001B[31m',
}

function style(text: string, code: string, theme: CliTheme): string {
  if (!theme.colorEnabled) {
    return text
  }
  return `${code}${text}${ANSI.reset}`
}

function inferColorEnabled(): boolean {
  if (process.env.NO_COLOR !== undefined) {
    return false
  }

  return process.stdout.isTTY === true && process.stderr.isTTY === true
}

function normalizeInvocationName(invocationName: string): string {
  const basename = path.basename(invocationName).toLowerCase()
  return basename.endsWith('.js') ? basename.slice(0, -3) : basename
}

function inferDefaultFromDocx(invocationName: string): boolean {
  const normalized = normalizeInvocationName(invocationName)
  return normalized === 'dtm' || normalized === 'docx-to-md'
}

function formatHelpText(invocationName: string, theme: CliTheme): string {
  const cmd = normalizeInvocationName(invocationName) || 'md-docx'

  return [
    `${style('md-docx CLI', `${ANSI.bold}${ANSI.cyan}`, theme)} ${style('Markdown ↔ DOCX converter', ANSI.dim, theme)}`,
    '',
    `${style('Usage:', ANSI.bold, theme)}`,
    `  ${cmd} <input.md> [output.docx] [--options <options.json>]`,
    `  ${cmd} <input-dir> [--recursive] [--options <options.json>]`,
    `  ${cmd} --from-docx <input.docx> [output.md] [--options <options.json>]`,
    `  ${cmd} --from-docx <input-dir> [--recursive] [--options <options.json>]`,
    '',
    `${style('Examples:', ANSI.bold, theme)}`,
    `  mtd .`,
    `  mtd ./docs -r`,
    `  dtm .`,
    `  dtm ./contracts -r`,
    `  md-docx report.md`,
    `  md-docx --from-docx contract.docx`,
    '',
    `${style('Aliases:', ANSI.bold, theme)}`,
    `  Markdown → DOCX: ${style('md-docx, mtd, md-to-docx', ANSI.green, theme)}`,
    `  DOCX → Markdown: ${style('dtm, docx-to-md', ANSI.green, theme)}`,
  ].join('\n')
}

function formatErrorMessage(message: string, theme: CliTheme): string {
  return `${style('✖ Error:', `${ANSI.bold}${ANSI.red}`, theme)} ${message}`
}

function formatInfoMessage(message: string, theme: CliTheme): string {
  return `${style('ℹ', ANSI.cyan, theme)} ${message}`
}

function formatSuccessMessage(message: string, theme: CliTheme): string {
  return `${style('✔', ANSI.green, theme)} ${message}`
}

function parseCliArgs(args: string[], defaultFromDocx: boolean): CliArgs {
  if (args.includes('-h') || args.includes('--help')) {
    return { showHelp: true }
  }

  const positional: string[] = []
  let optionsPath: string | undefined
  let recursive = false
  let fromDocx = defaultFromDocx

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
  invocationName = 'md-docx',
): Promise<number> {
  const theme: CliTheme = { colorEnabled: inferColorEnabled() }

  try {
    const parsedArgs = parseCliArgs(args, inferDefaultFromDocx(invocationName))
    if ('showHelp' in parsedArgs) {
      output.log(formatHelpText(invocationName, theme))
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
        output.log(formatSuccessMessage(`Markdown created at: ${outputPath}`, theme))
        return 0
      }

      const outputPath = parsedArgs.outputPath
        ? path.resolve(parsedArgs.outputPath)
        : isMarkdownFile(inputPath)
          ? inputPath.replace(/\.(md|markdown)$/i, '.docx')
          : undefined

      if (!outputPath) {
        throw new Error(
          'Output path is required when input file extension is not .md/.markdown/.docx',
        )
      }

      await convertSingleMarkdownFile(inputPath, outputPath, optionsObject as Options | undefined)
      output.log(formatSuccessMessage(`DOCX created at: ${outputPath}`, theme))
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
            `Supported input extension: .docx${
              parsedArgs.recursive
                ? ''
                : '. ' + 'If your files are in subfolders, run with -r/--recursive'
            }`,
        )
      }

      output.log(
        formatInfoMessage(
          `DOCX mode detected. Converting ${docxFiles.length} file(s) from ${inputPath}`,
          theme,
        ),
      )

      let convertedCount = 0
      for (const docxFilePath of docxFiles) {
        const outputPath = docxFilePath.replace(/\.docx$/i, '.md')
        await convertSingleDocxFile(
          docxFilePath,
          outputPath,
          optionsObject as DocxToMarkdownOptions | undefined,
        )
        convertedCount++
        output.log(formatSuccessMessage(`Markdown created at: ${outputPath}`, theme))
      }

      output.log(
        formatSuccessMessage(
          `Converted ${convertedCount} file(s) from DOCX to Markdown in directory: ${inputPath}`,
          theme,
        ),
      )
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

    output.log(
      formatInfoMessage(
        `Markdown mode detected. Converting ${markdownFiles.length} file(s) from ${inputPath}`,
        theme,
      ),
    )

    let convertedCount = 0
    for (const markdownFilePath of markdownFiles) {
      const outputPath = markdownFilePath.replace(/\.(md|markdown)$/i, '.docx')
      await convertSingleMarkdownFile(
        markdownFilePath,
        outputPath,
        optionsObject as Options | undefined,
      )
      convertedCount++
      output.log(formatSuccessMessage(`DOCX created at: ${outputPath}`, theme))
    }

    output.log(
      formatSuccessMessage(
        `Converted ${convertedCount} file(s) from directory: ${inputPath}`,
        theme,
      ),
    )
    return 0
  } catch (err) {
    const message =
      err instanceof Error
        ? err.message
        : typeof err === 'object' && err !== null && 'message' in err
          ? String((err as { message: unknown }).message)
          : String(err)
    output.error(formatErrorMessage(message, theme))
    output.error('')
    output.error(formatHelpText(invocationName, theme))
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
  void runCli(process.argv.slice(2), DEFAULT_CLI_OUTPUT, process.argv[1] ?? 'md-docx').then(
    (exitCode) => {
      process.exitCode = exitCode
    },
  )
}
