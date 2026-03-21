#!/usr/bin/env node
import fs from 'node:fs/promises'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

import { convertMarkdownToDocx } from './index.js'
import type { Options } from './types.js'

export interface CliOutput {
  log: (message: string) => void
  error: (message: string) => void
}

interface ParsedCliArgs {
  inputPath: string
  outputPath?: string
  optionsPath?: string
  recursive: boolean
}

interface HelpCliArgs {
  showHelp: true
}

type CliArgs = ParsedCliArgs | HelpCliArgs

const HELP_TEXT = `Usage:
  md-docx <input.md> <output.docx> [--options <options.json>]
  md-docx <input-dir> [--recursive] [--options <options.json>]

Examples:
  md-docx a.md b.docx
  mtd a.md b.docx
  mtd .
  mtd docs --recursive
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

    if (arg.startsWith('-')) {
      throw new Error(`Unknown argument: ${arg}`)
    }

    positional.push(arg)
  }

  if (positional.length < 1 || positional.length > 2) {
    throw new Error('Expected either <input.md> <output.docx> or <input-dir>')
  }

  return {
    inputPath: positional[0],
    outputPath: positional[1],
    optionsPath,
    recursive,
  }
}

function isMarkdownFile(filePath: string): boolean {
  const lower = filePath.toLowerCase()
  return lower.endsWith('.md') || lower.endsWith('.markdown')
}

async function collectMarkdownFiles(dirPath: string, recursive: boolean): Promise<string[]> {
  const entries = await fs.readdir(dirPath, { withFileTypes: true })
  const files: string[] = []

  for (const entry of entries) {
    const entryPath = path.join(dirPath, entry.name)

    if (entry.isDirectory()) {
      if (recursive) {
        files.push(...(await collectMarkdownFiles(entryPath, true)))
      }
      continue
    }

    if (entry.isFile() && isMarkdownFile(entry.name)) {
      files.push(entryPath)
    }
  }

  return files
}

async function convertSingleFile(
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

async function readOptionsFile(optionsPath: string): Promise<Options> {
  const content = await fs.readFile(optionsPath, 'utf8')

  try {
    const parsed: unknown = JSON.parse(content)

    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
      throw new Error('Options JSON must be an object')
    }

    return parsed as Options
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
    const options = parsedArgs.optionsPath
      ? await readOptionsFile(path.resolve(parsedArgs.optionsPath))
      : undefined
    const inputStat = await fs.stat(inputPath)

    if (inputStat.isFile()) {
      if (!parsedArgs.outputPath) {
        throw new Error('Output path is required when input is a file')
      }

      const outputPath = path.resolve(parsedArgs.outputPath)
      await convertSingleFile(inputPath, outputPath, options)
      output.log(`DOCX created at: ${outputPath}`)
      return 0
    }

    if (!inputStat.isDirectory()) {
      throw new Error(`Input path is neither a file nor directory: ${parsedArgs.inputPath}`)
    }

    if (parsedArgs.outputPath) {
      throw new Error('Output path must not be provided when input is a directory')
    }

    const markdownFiles = await collectMarkdownFiles(inputPath, parsedArgs.recursive)
    if (markdownFiles.length === 0) {
      throw new Error(`No markdown files found in directory: ${parsedArgs.inputPath}`)
    }

    let convertedCount = 0
    for (const markdownFilePath of markdownFiles) {
      const outputPath = markdownFilePath.replace(/\.(md|markdown)$/i, '.docx')
      await convertSingleFile(markdownFilePath, outputPath, options)
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
const invokedFilePath = process.argv[1] ? path.resolve(process.argv[1]) : ''

if (invokedFilePath === path.resolve(currentFilePath)) {
  void runCli(process.argv.slice(2)).then((exitCode) => {
    process.exitCode = exitCode
  })
}
