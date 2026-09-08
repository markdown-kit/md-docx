import type { Table } from 'docx'
import { Paragraph, TextRun, PageBreak, AlignmentType } from 'docx'

import type {
  DocxDocumentModel,
  DocxBlockNode,
  DocxListNode,
  DocxListItemNode,
} from './docxModel.js'
import {
  processHeading,
  processTable,
  processCodeBlock,
  processBlockquote,
  processComment,
  processImage,
  processParagraph,
  processListItem,
} from './helpers.js'
import type { Style, Options } from './types.js'

/**
 * Converts internal docx model to docx Paragraph/Table objects
 * Handles nested lists with proper level tracking
 */
export async function modelToDocx(
  model: DocxDocumentModel,
  style: Style,
  options: Options,
  renderOptions: { sequenceIdOffset?: number } = {},
): Promise<{
  children: Array<Paragraph | Table>
  headings: Array<{ text: string; level: number; bookmarkId: string }>
  maxSequenceId: number
  sequenceStarts: Map<number, number>
  tocPlaceholders: WeakSet<Paragraph>
}> {
  const children: Array<Paragraph | Table> = []
  const headings: Array<{ text: string; level: number; bookmarkId: string }> = []
  const documentType = options.documentType ?? 'document'
  const sequenceIdOffset = renderOptions.sequenceIdOffset ?? 0
  const tocPlaceholders = new WeakSet<Paragraph>()

  // Track numbering sequences for nested lists
  let maxSequenceId = 0
  // Per-sequence starting number for ordered lists that use `start: N`.
  const sequenceStarts = new Map<number, number>()

  function encodeInlineNode(node: {
    value: string
    bold?: boolean
    italic?: boolean
    underline?: boolean
    strikethrough?: boolean
    superScript?: boolean
    subScript?: boolean
    code?: boolean
    link?: string
  }): string {
    if (node.code) {
      return `\`${node.value}\``
    }

    let text = node.link ? `[${node.value}](${node.link})` : node.value

    if (node.superScript) {
      text = `<sup>${text}</sup>`
    } else if (node.subScript) {
      text = `<sub>${text}</sub>`
    }
    if (node.strikethrough) {
      text = `~~${text}~~`
    }
    if (node.underline) {
      text = `++${text}++`
    }

    if (node.bold && node.italic) return `***${text}***`
    if (node.bold) return `**${text}**`
    if (node.italic) return `*${text}*`
    return text
  }

  /**
   * Encode a run of inline nodes to markdown-like text. Adjacent nodes that
   * share an outer marker (e.g. bold text followed by bold superscript) would
   * otherwise produce `****`, which the inline parser reads as bold+italic
   * toggles; matching closing/opening markers are collapsed instead.
   */
  function encodeInlineNodes(nodes: Array<Parameters<typeof encodeInlineNode>[0]>): string {
    let out = ''
    for (const part of nodes.map((c) => encodeInlineNode(c))) {
      const trailingStars = /\*+$/.exec(out)?.[0].length ?? 0
      const leadingStars = /^\*+/.exec(part)?.[0].length ?? 0
      if (trailingStars > 0 && trailingStars === leadingStars && out.length > trailingStars) {
        out = out.slice(0, -trailingStars) + part.slice(leadingStars)
        continue
      }
      let merged = false
      for (const marker of ['~~', '++']) {
        if (out.endsWith(marker) && part.startsWith(marker) && out.length > marker.length) {
          out = out.slice(0, -marker.length) + part.slice(marker.length)
          merged = true
          break
        }
      }
      if (!merged) {
        out += part
      }
    }
    return out
  }

  function renderBlockNode(node: DocxBlockNode, listLevel = 0): Array<Paragraph | Table> {
    switch (node.type) {
      case 'heading': {
        // Re-encode inline formatting into markdown-like syntax for helpers.
        const headingText = encodeInlineNodes(node.children)

        const headingLine = `${'#'.repeat(node.level)} ${headingText}`
        const config = {
          level: node.level,
          size: 0,
          style: node.level === 1 ? 'Title' : undefined,
        }
        const { paragraph, bookmarkId } = processHeading(headingLine, config, style)
        headings.push({
          text: headingText,
          level: node.level,
          bookmarkId,
        })
        return [paragraph]
      }

      case 'paragraph': {
        const paragraphText = encodeInlineNodes(node.children)
        return [processParagraph(paragraphText, style)]
      }

      case 'list': {
        return renderList(node, listLevel || 0)
      }

      case 'codeBlock': {
        return [processCodeBlock(node.value, node.language, style)]
      }

      case 'blockquote': {
        // Preserve inline formatting (bold/italic/link/code) by re-encoding to
        // markdown, and include non-paragraph children (headings, lists, code)
        // that were previously dropped.
        const quoteText = node.children
          .map((child) => {
            switch (child.type) {
              case 'paragraph':
              case 'heading':
                return encodeInlineNodes(child.children)
              case 'list':
                return child.children
                  .map((listItem, index) => {
                    const marker = child.ordered ? `${(child.start ?? 1) + index}.` : '-'
                    const itemText = listItem.children
                      .map((block) =>
                        block.type === 'paragraph' ? encodeInlineNodes(block.children) : '',
                      )
                      .filter(Boolean)
                      .join(' ')
                    return `${marker} ${itemText}`
                  })
                  .join('\n')
              case 'codeBlock':
                return child.value
              default:
                return ''
            }
          })
          .filter(Boolean)
          .join('\n')
        return [processBlockquote(quoteText, style)]
      }

      case 'thematicBreak': {
        // Render a horizontal rule as an empty paragraph with a bottom border.
        return [
          new Paragraph({
            border: {
              bottom: { color: '999999', space: 1, style: 'single', size: 6 },
            },
            spacing: { before: 120, after: 120 },
          }),
        ]
      }

      case 'image': {
        // processImage returns Promise<Paragraph[]>, so we need to handle it specially
        // For now, return empty array and handle images separately
        return []
      }

      case 'table': {
        const tableData = {
          headers: node.headers.map((cells) => encodeInlineNodes(cells)),
          rows: node.rows.map((row) => row.map((cells) => encodeInlineNodes(cells))),
          align: node.align,
        }
        return [processTable(tableData, documentType, style)]
      }

      case 'comment': {
        return [processComment(node.value, style)]
      }

      case 'pageBreak': {
        return [new Paragraph({ children: [new PageBreak()] })]
      }

      case 'tocPlaceholder': {
        const placeholder = new Paragraph({})
        tocPlaceholders.add(placeholder)
        return [placeholder]
      }

      default:
        return []
    }
  }

  function renderList(list: DocxListNode, currentLevel: number): Array<Paragraph | Table> {
    const nodes: Array<Paragraph | Table> = []
    const startNumber = list.ordered ? (list.start ?? 1) : 1
    let itemNumber = startNumber
    const adjustedSequenceId = list.sequenceId ? list.sequenceId + sequenceIdOffset : undefined

    // Track max sequence ID
    if (adjustedSequenceId && adjustedSequenceId > maxSequenceId) {
      maxSequenceId = adjustedSequenceId
    }

    // Record a non-default start so the numbering definition begins at it.
    if (adjustedSequenceId && startNumber !== 1) {
      sequenceStarts.set(adjustedSequenceId, startNumber)
    }

    for (const item of list.children) {
      nodes.push(
        ...renderListItem(item, list.ordered, currentLevel, adjustedSequenceId, itemNumber),
      )
      itemNumber++
    }

    return nodes
  }

  /**
   * Render one list item. Only the FIRST paragraph receives the list marker;
   * subsequent paragraphs (loose list items) render as continuation paragraphs
   * without a marker so they do not create spurious extra numbers/bullets.
   * Non-paragraph block children (nested lists, code blocks, tables,
   * blockquotes) are preserved rather than dropped.
   */
  function renderListItem(
    item: DocxListItemNode,
    isOrdered: boolean,
    level: number,
    sequenceId: number | undefined,
    itemNumber: number,
  ): Array<Paragraph | Table> {
    const nodes: Array<Paragraph | Table> = []
    let markerEmitted = false

    for (const child of item.children) {
      if (child.type === 'list') {
        nodes.push(...renderList(child, level + 1))
      } else if (child.type === 'paragraph') {
        const paragraphText = encodeInlineNodes(child.children)

        if (!markerEmitted) {
          // First paragraph: the actual list item with its number/bullet.
          nodes.push(
            processListItem(
              {
                text: paragraphText,
                isNumbered: isOrdered,
                listNumber: itemNumber,
                sequenceId: sequenceId ?? 1,
                level,
              },
              style,
            ),
          )
          markerEmitted = true
        } else {
          // Continuation paragraph: no list marker and no numbering, so loose
          // list items do not produce spurious extra numbers/bullets.
          nodes.push(processParagraph(paragraphText, style))
        }
      } else {
        // Other block content (code blocks, tables, blockquotes) — preserve it
        // (previously tables were silently dropped).
        nodes.push(...renderBlockNode(child, level))
      }
    }

    // Ensure every item is visible even if it had no paragraph child.
    if (!markerEmitted && nodes.length === 0) {
      nodes.push(
        processListItem(
          {
            text: '',
            isNumbered: isOrdered,
            listNumber: itemNumber,
            sequenceId: sequenceId ?? 1,
            level,
          },
          style,
        ),
      )
    }

    return nodes
  }

  // Process all top-level nodes
  for (const node of model.children) {
    if (node.type === 'image') {
      // Handle images asynchronously
      try {
        const imageParagraphs = await processImage(node.alt, node.url, style)
        children.push(...imageParagraphs)
      } catch {
        children.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `[Image could not be loaded: ${node.alt}]`,
                italics: true,
                color: 'FF0000',
              }),
            ],
            alignment: AlignmentType.CENTER,
            bidirectional: style.direction === 'RTL',
          }),
        )
      }
    } else {
      const rendered = renderBlockNode(node)
      children.push(...rendered)
    }
  }

  return { children, headings, maxSequenceId, sequenceStarts, tocPlaceholders }
}
