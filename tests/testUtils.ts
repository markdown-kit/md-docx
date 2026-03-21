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

export async function readDocumentXml(blob: Blob): Promise<string> {
  const buffer = await blob.arrayBuffer()
  const zip = await JSZip.loadAsync(Buffer.from(buffer))
  const documentXml = zip.file('word/document.xml')

  if (!documentXml) {
    throw new Error('word/document.xml not found in DOCX')
  }

  return documentXml.async('string')
}
