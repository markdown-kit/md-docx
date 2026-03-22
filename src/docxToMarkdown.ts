import { Buffer } from 'node:buffer'
import { createRequire } from 'node:module'

import * as mammoth from 'mammoth'
import TurndownService from 'turndown'

import type { DocxToMarkdownOptions, DocxToMarkdownTurndownOptions } from './types.js'

type TurndownGfmPlugin = (service: TurndownService) => void

const require = createRequire(import.meta.url)
const { gfm } = require('turndown-plugin-gfm') as {
  gfm: TurndownGfmPlugin
}

export type DocxInput = ArrayBuffer | Uint8Array | Buffer

const defaultTurndownOptions: DocxToMarkdownTurndownOptions = {
  headingStyle: 'atx',
  bulletListMarker: '-',
  codeBlockStyle: 'fenced',
  fence: '```',
  emDelimiter: '*',
  strongDelimiter: '**',
  linkStyle: 'inlined',
  linkReferenceStyle: 'full',
}

export class DocxToMarkdownError extends Error {
  constructor(
    message: string,
    public context?: unknown,
  ) {
    super(message)
    this.name = 'DocxToMarkdownError'
  }
}

function toUint8Array(input: DocxInput): Uint8Array {
  if (input instanceof Uint8Array) {
    return input
  }

  if (input instanceof ArrayBuffer) {
    return new Uint8Array(input)
  }

  throw new TypeError('Unsupported DOCX input. Expected Buffer, Uint8Array, or ArrayBuffer.')
}

function normalizeMarkdownOutput(markdown: string): string {
  const normalized = markdown
    .replace(/\r\n/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()

  return normalized.length > 0 ? `${normalized}\n` : ''
}

function createTurndownService(options: DocxToMarkdownOptions): TurndownService {
  const turndown = new TurndownService({
    ...defaultTurndownOptions,
    ...options.turndown,
  })
  turndown.use(gfm)
  return turndown
}

/**
 * Convert DOCX binary content to Markdown.
 */
export async function convertDocxToMarkdown(
  docxInput: DocxInput,
  options: DocxToMarkdownOptions = {},
): Promise<string> {
  try {
    const inputBytes = toUint8Array(docxInput)
    if (inputBytes.byteLength === 0) {
      throw new DocxToMarkdownError('Invalid DOCX input: file content is empty')
    }

    const mammothResult = await mammoth.convertToHtml(
      {
        buffer: Buffer.from(inputBytes),
      },
      {
        styleMap: options.mammoth?.styleMap,
        includeDefaultStyleMap: options.mammoth?.includeDefaultStyleMap,
        includeEmbeddedStyleMap: options.mammoth?.includeEmbeddedStyleMap,
        ignoreEmptyParagraphs: options.mammoth?.preserveEmptyParagraphs === true ? false : true,
        convertImage: mammoth.images.dataUri,
      },
    )

    const turndown = createTurndownService(options)
    const markdown = turndown.turndown(mammothResult.value)

    if (options.normalizeWhitespace === false) {
      return markdown
    }

    return normalizeMarkdownOutput(markdown)
  } catch (err) {
    if (err instanceof DocxToMarkdownError) {
      throw err
    }

    throw new DocxToMarkdownError(
      `Failed to convert DOCX to Markdown: ${err instanceof Error ? err.message : String(err)}`,
      { originalError: err },
    )
  }
}
