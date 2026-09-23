import * as fs from 'node:fs'
import * as path from 'node:path'

import JSZip from 'jszip'

export const TEST_OUTPUT_DIR = path.join(process.cwd(), 'test-output')

export function ensureTestOutputDir(): void {
  if (!fs.existsSync(TEST_OUTPUT_DIR)) {
    fs.mkdirSync(TEST_OUTPUT_DIR, { recursive: true })
  }
}

export async function writeBlobDocx(blob: Blob, filename: string): Promise<string> {
  ensureTestOutputDir()
  const outputPath = path.join(TEST_OUTPUT_DIR, filename)
  const arrayBuffer = await blob.arrayBuffer()
  fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))
  return outputPath
}

export async function readDocxEntry(blob: Blob, entryPath: string): Promise<string> {
  const buffer = await blob.arrayBuffer()
  const zip = await JSZip.loadAsync(Buffer.from(buffer))
  const entry = zip.file(entryPath)

  if (!entry) {
    throw new Error(`${entryPath} not found in DOCX`)
  }

  return entry.async('string')
}

export async function readDocumentXml(blob: Blob): Promise<string> {
  return readDocxEntry(blob, 'word/document.xml')
}

export interface RunInfo {
  text: string
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  code: boolean
  hyperlink: boolean
}

/**
 * Flatten the body of a document.xml into its text runs (in order) with the
 * formatting flags the tests care about. Runs without text are skipped.
 */
export function readRuns(documentXml: string): RunInfo[] {
  const body = documentXml.split('<w:body>')[1] ?? documentXml
  const runs: RunInfo[] = []
  const tokenPattern = /<w:hyperlink\b[^>]*>|<\/w:hyperlink>|<w:r>([\s\S]*?)<\/w:r>/gu
  let inHyperlink = false
  for (const match of body.matchAll(tokenPattern)) {
    if (match[0].startsWith('<w:hyperlink')) {
      inHyperlink = true
      continue
    }
    if (match[0] === '</w:hyperlink>') {
      inHyperlink = false
      continue
    }
    const run = match[1] ?? ''
    const text = [...run.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/gu)].map((t) => t[1]).join('')
    if (!text) {
      continue
    }
    runs.push({
      text: text.replace(/&amp;/gu, '&').replace(/&lt;/gu, '<').replace(/&gt;/gu, '>'),
      bold: /<w:b(?: w:val="true")?\/>/u.test(run),
      italic: /<w:i(?: w:val="true")?\/>/u.test(run),
      underline: /<w:u w:val="single"\/>/u.test(run),
      strike: /<w:strike(?: w:val="true")?\/>/u.test(run),
      code: /Courier New/u.test(run),
      hyperlink: inHyperlink,
    })
  }
  return runs
}
