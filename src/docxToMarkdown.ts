import { Buffer } from 'node:buffer'
import { createRequire } from 'node:module'

import * as mammoth from 'mammoth'
import TurndownService from 'turndown'

import { parseDocxToMarkdownOptionsFile } from './options-file.js'
import type { DocxToMarkdownOptions, DocxToMarkdownTurndownOptions } from './types.js'

type TurndownGfmPlugin = (service: TurndownService) => void

const require = createRequire(import.meta.url)

function isObjectRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function loadGfmPlugin(): TurndownGfmPlugin {
  const pluginModule: unknown = require('turndown-plugin-gfm')

  if (!isObjectRecord(pluginModule) || typeof pluginModule.gfm !== 'function') {
    throw new TypeError('turndown-plugin-gfm did not expose a callable gfm plugin')
  }

  const plugin = pluginModule.gfm
  return (service: TurndownService) => {
    plugin(service)
  }
}

const gfm = loadGfmPlugin()

export type DocxInput = ArrayBuffer | Uint8Array | Buffer

const defaultMammothStyleMap = [
  "p[style-name='Title'] => h1:fresh",
  "p[style-name='Heading 1'] => h1:fresh",
  "p[style-name='Heading 2'] => h2:fresh",
  "p[style-name='Heading 3'] => h3:fresh",
  "p[style-name='Heading 4'] => h4:fresh",
  "p[style-name='Heading 5'] => h5:fresh",
  "p[style-name='Heading1'] => h1:fresh",
  "p[style-name='Heading2'] => h2:fresh",
  "p[style-name='Heading3'] => h3:fresh",
  "p[style-name='Heading4'] => h4:fresh",
  "p[style-name='Heading5'] => h5:fresh",
  "p[style-id='1'] => h1:fresh",
  "p[style-id='2'] => h2:fresh",
  "p[style-id='3'] => h3:fresh",
  "p[style-id='4'] => h4:fresh",
  "p[style-id='5'] => h5:fresh",
  "p[style-id='6'] => h6:fresh",
  // Synthetic style names assigned by `buildReverseTransform` so direct run/
  // paragraph formatting (which mammoth otherwise drops) round-trips back to
  // markdown: strikethrough, underline, inline code, and code blocks.
  "r[style-name='MddStrike'] => del",
  "r[style-name='MddUnderline'] => u",
  "r[style-name='MddCode'] => code",
  "p[style-name='MddCodeBlock'] => pre:separator('\\n')",
] as const

/** A monospace font implies code in the forward (md→docx) path. */
function isMonospaceFont(font: unknown): boolean {
  return typeof font === 'string' && /courier|consolas|monaco|mono/iu.test(font)
}

/**
 * Build a mammoth document transform that assigns synthetic style names to
 * runs/paragraphs carrying direct formatting (strikethrough, underline,
 * monospace) so the style map can convert them to del/u/code/pre — the inverse
 * of the md→docx renderer, which applies these as direct formatting rather than
 * named styles (which mammoth would otherwise silently discard).
 */
interface MammothTransforms {
  run(fn: (run: Record<string, unknown>) => unknown): (document: unknown) => unknown
  paragraph(fn: (paragraph: Record<string, unknown>) => unknown): (document: unknown) => unknown
}

// mammoth.transforms is a documented runtime API but is absent from the bundled
// type definitions; this narrowed interop access is justified and guarded.
const mammothTransforms = (mammoth as unknown as { transforms?: MammothTransforms }).transforms

function buildReverseTransform(): (document: unknown) => unknown {
  if (!mammothTransforms) {
    return (document: unknown) => document
  }

  const runTransform = mammothTransforms.run((run: Record<string, unknown>) => {
    if (run.styleName) {
      return run
    }
    if (run.isStrikethrough) {
      return { ...run, styleName: 'MddStrike' }
    }
    if (run.isUnderline) {
      return { ...run, styleName: 'MddUnderline' }
    }
    if (isMonospaceFont(run.font)) {
      return { ...run, styleName: 'MddCode' }
    }
    return run
  })

  const paragraphTransform = mammothTransforms.paragraph((paragraph: Record<string, unknown>) => {
    if (paragraph.styleName) {
      return paragraph
    }
    const children = Array.isArray(paragraph.children) ? paragraph.children : []
    const runs = children.filter(
      (child): child is Record<string, unknown> => isObjectRecord(child) && child.type === 'run',
    )
    if (
      runs.length > 0 &&
      runs.every((run) => isMonospaceFont(run.font) || run.styleName === 'MddCode')
    ) {
      return { ...paragraph, styleName: 'MddCodeBlock' }
    }
    return paragraph
  })

  return (document: unknown) => paragraphTransform(runTransform(document))
}

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
    .replace(/<table[\s\S]*?<\/table>/gi, (htmlTable) => {
      const turndown = new TurndownService({
        ...defaultTurndownOptions,
      })
      turndown.use(gfm)
      return `\n${turndown.turndown(htmlTable).trim()}\n`
    })
    .replace(/^(#{1,6})\s+\*\*(.+?)\*\*\s*$/gm, '$1 $2')
    .replace(/^(#{1,6})\s+_(.+?)_\s*$/gm, '$1 $2')
    .replace(/^\*\*(.+?)\*\*\s*$/gm, '$1')
    .replace(/^_(.+?)_\s*$/gm, '$1')
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
  // Underline (`<u>`) has no standard markdown; round-trip it to the `++text++`
  // syntax the md→docx renderer uses for underline.
  turndown.addRule('mddUnderline', {
    filter: ['u'],
    replacement: (content: string) => (content ? `++${content}++` : ''),
  })
  return turndown
}

function preprocessMammothHtml(html: string): string {
  return html
    .replace(/<(td|th)>\s*<p>([\s\S]*?)<\/p>\s*<\/(td|th)>/gi, (_match, openTag, content) => {
      return `<${String(openTag).toLowerCase()}>${content}</${String(openTag).toLowerCase()}>`
    })
    .replace(/<(\/?)([A-Za-z][A-Za-z0-9-]*)/g, (_match, slash: string, tagName: string) => {
      return `<${slash}${tagName.toLowerCase()}`
    })
}

/**
 * Convert DOCX binary content to Markdown.
 */
export async function convertDocxToMarkdown(
  docxInput: DocxInput,
  options: DocxToMarkdownOptions = {},
): Promise<string> {
  try {
    const validatedOptions = parseDocxToMarkdownOptionsFile(options, 'options')
    const inputBytes = toUint8Array(docxInput)
    if (inputBytes.byteLength === 0) {
      throw new DocxToMarkdownError('Invalid DOCX input: file content is empty')
    }

    const mammothResult = await mammoth.convertToHtml(
      {
        buffer: Buffer.from(inputBytes),
      },
      {
        styleMap: [...defaultMammothStyleMap, ...(validatedOptions.mammoth?.styleMap ?? [])],
        includeDefaultStyleMap: validatedOptions.mammoth?.includeDefaultStyleMap,
        includeEmbeddedStyleMap: validatedOptions.mammoth?.includeEmbeddedStyleMap,
        ignoreEmptyParagraphs:
          validatedOptions.mammoth?.preserveEmptyParagraphs === true ? false : true,
        convertImage: mammoth.images.dataUri,
        // Surface direct strike/underline/monospace formatting as styled HTML so
        // it round-trips back to markdown (mammoth drops it otherwise).
        transformDocument: buildReverseTransform(),
      },
    )

    const turndown = createTurndownService(validatedOptions)
    const markdown = turndown.turndown(preprocessMammothHtml(mammothResult.value))

    if (validatedOptions.normalizeWhitespace === false) {
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
