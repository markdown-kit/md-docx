import * as fs from 'node:fs'
import * as path from 'node:path'

import { describe, expect, it } from 'vitest'

import { convertMarkdownToDocx, MarkdownConversionError, parseToDocxOptions } from '../src/index'
import type { Options } from '../src/types'
import { readDocxEntry } from './testUtils'

const outputDir = path.join(process.cwd(), 'test-output')

if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true })
}

describe('sections API', () => {
  it('builds distinct section properties and footer behavior', async () => {
    const options: Options = {
      template: {
        pageNumbering: {
          display: 'current',
          alignment: 'RIGHT',
        },
      },
      sections: [
        {
          markdown: '# Cover\n\nCustom cover content.',
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
          markdown: '# Body\n\nMain content starts here.',
          footers: {
            default: {
              text: 'Page',
              pageNumberDisplay: 'currentAndSectionTotal',
              alignment: 'RIGHT',
            },
          },
          pageNumbering: {
            start: 1,
            formatType: 'decimal',
          },
        },
      ],
    }

    const docxOptions = await parseToDocxOptions('', options)

    expect(docxOptions.sections).toHaveLength(2)
    expect(docxOptions.sections[0].footers).toBeUndefined()
    expect(docxOptions.sections[1].footers?.default).toBeDefined()
    expect(docxOptions.sections[1].properties?.page?.pageNumbers?.start).toBe(1)
    expect(docxOptions.sections[1].properties?.page?.pageNumbers?.formatType).toBe('decimal')
  })

  it('creates independent numbered-list references across sections', async () => {
    const docxOptions = await parseToDocxOptions('', {
      sections: [
        {
          markdown: '1. First section item\n2. Second item',
        },
        {
          markdown: '1. New section item\n2. Another section item',
        },
      ],
    })

    expect(docxOptions.numbering?.config).toHaveLength(2)
    expect(docxOptions.numbering?.config[0].reference).toBe('numbered-list-1')
    expect(docxOptions.numbering?.config[1].reference).toBe('numbered-list-2')
  })

  it('supports style overrides per section during full conversion', async () => {
    const blob = await convertMarkdownToDocx('', {
      style: {
        paragraphSize: 24,
      },
      sections: [
        {
          markdown: '# Section One\n\nSmaller text section.',
          style: {
            paragraphSize: 20,
          },
          pageNumbering: {
            display: 'none',
          },
        },
        {
          markdown: '# Section Two\n\nLarger text section.',
          style: {
            paragraphSize: 30,
          },
          pageNumbering: {
            start: 1,
          },
        },
      ],
    })

    const arrayBuffer = await blob.arrayBuffer()
    const outputPath = path.join(outputDir, 'sections-style-overrides.docx')
    fs.writeFileSync(outputPath, Buffer.from(arrayBuffer))

    expect(blob).toBeInstanceOf(Blob)
    expect(blob.size).toBeGreaterThan(0)
  })

  it('applies advanced section properties from template and section overrides', async () => {
    const docxOptions = await parseToDocxOptions('', {
      template: {
        page: {
          margin: {
            top: 1600,
            left: 1200,
          },
        },
        headers: {
          default: {
            text: 'Template Header',
            alignment: 'LEFT',
          },
        },
        pageNumbering: {
          display: 'current',
          alignment: 'CENTER',
        },
      },
      sections: [
        {
          markdown: '# Intro\n\nTemplate should apply here.',
          titlePage: true,
          type: 'ODD_PAGE',
          page: {
            size: {
              orientation: 'LANDSCAPE',
            },
          },
          headers: {
            first: {
              text: 'First Page Header',
              alignment: 'RIGHT',
            },
          },
          footers: {
            default: {
              text: 'Custom Footer',
              pageNumberDisplay: 'currentAndTotal',
              alignment: 'RIGHT',
            },
          },
          pageNumbering: {
            start: 5,
            formatType: 'upperRoman',
            separator: 'colon',
          },
        },
      ],
    })

    expect(docxOptions.sections).toHaveLength(1)
    expect(docxOptions.sections[0].headers?.default).toBeDefined()
    expect(docxOptions.sections[0].headers?.first).toBeDefined()
    expect(docxOptions.sections[0].footers?.default).toBeDefined()
    expect(docxOptions.sections[0].properties?.titlePage).toBe(true)
    expect(docxOptions.sections[0].properties?.type).toBeDefined()
    expect(docxOptions.sections[0].properties?.page?.size?.orientation).toBeDefined()
    expect(docxOptions.sections[0].properties?.page?.margin?.top).toBe(1600)
    expect(docxOptions.sections[0].properties?.page?.margin?.left).toBe(1200)
    expect(docxOptions.sections[0].properties?.page?.pageNumbers?.start).toBe(5)
    expect(docxOptions.sections[0].properties?.page?.pageNumbers?.formatType).toBeDefined()
    expect(docxOptions.sections[0].properties?.page?.pageNumbers?.separator).toBeDefined()
  })

  it('supports section type and page numbering variants across sections', async () => {
    const docxOptions = await parseToDocxOptions('', {
      template: {
        pageNumbering: {
          display: 'current',
          alignment: 'RIGHT',
        },
      },
      sections: [
        {
          markdown: '# A\n\nOne',
          type: 'NEXT_PAGE',
          pageNumbering: {
            start: 1,
            formatType: 'decimal',
          },
        },
        {
          markdown: '# B\n\nTwo',
          type: 'CONTINUOUS',
          pageNumbering: {
            start: 1,
            formatType: 'lowerRoman',
            separator: 'period',
          },
        },
        {
          markdown: '# C\n\nThree',
          type: 'EVEN_PAGE',
          pageNumbering: {
            start: 1,
            formatType: 'upperLetter',
            separator: 'hyphen',
          },
        },
      ],
    })

    expect(docxOptions.sections).toHaveLength(3)
    expect(docxOptions.sections[0].properties?.type).toBeDefined()
    expect(docxOptions.sections[1].properties?.type).toBeDefined()
    expect(docxOptions.sections[2].properties?.type).toBeDefined()
    expect(docxOptions.sections[1].properties?.page?.pageNumbers?.separator).toBeDefined()
    expect(docxOptions.sections[2].properties?.page?.pageNumbers?.formatType).toBeDefined()
    expect(docxOptions.sections[2].footers?.default).toBeDefined()
  })

  it('throws for invalid section type', async () => {
    await expect(
      parseToDocxOptions('', {
        sections: [
          {
            markdown: '# Bad',
            type: 'BAD_TYPE' as any,
          },
        ],
      }),
    ).rejects.toThrow(MarkdownConversionError)
  })

  it('throws for invalid page orientation', async () => {
    await expect(
      parseToDocxOptions('', {
        sections: [
          {
            markdown: '# Bad',
            page: {
              size: {
                orientation: 'SIDEWAYS' as any,
              },
            },
          },
        ],
      }),
    ).rejects.toThrow('Invalid page orientation')
  })

  it('throws for invalid header alignment', async () => {
    await expect(
      parseToDocxOptions('', {
        sections: [
          {
            markdown: '# Bad',
            headers: {
              default: {
                text: 'Header',
                alignment: 'MIDDLE' as any,
              },
            },
          },
        ],
      }),
    ).rejects.toThrow('Invalid header/footer alignment')
  })

  it('renders {{page}} inline in header/footer text and suppresses the trailing field', async () => {
    const blob = await convertMarkdownToDocx('# Body\n\nContent.', {
      template: {
        headers: {
          default: { text: 'Acme | Page {{page}} | Confidential', alignment: 'CENTER' },
        },
        footers: {
          default: { text: 'Questions? billing@acme.com', pageNumberDisplay: 'none' },
        },
      },
    })

    const headerXml = await readDocxEntry(blob, 'word/header1.xml')
    const footerXml = await readDocxEntry(blob, 'word/footer1.xml')

    // The literal token must not survive; a PAGE field must sit between the text parts.
    expect(headerXml).not.toContain('{{page}}')
    expect(headerXml).toMatch(/Acme \| Page <\/w:t>[\s\S]*PAGE[\s\S]*<w:t[^>]*> \| Confidential/u)
    expect((headerXml.match(/w:instrText/gu) ?? []).length).toBe(2)

    // Footer text without a token and with display 'none' carries no PAGE field.
    expect(footerXml).toContain('Questions? billing@acme.com')
    expect(footerXml).not.toMatch(/PAGE/u)
  })

  it('enables even/odd pages when an even header or footer is configured', async () => {
    const evenOptions = await parseToDocxOptions('# Doc', {
      template: {
        headers: { default: { text: 'Odd pages' }, even: { text: 'Even pages' } },
      },
    })
    expect(evenOptions.evenAndOddHeaderAndFooters).toBe(true)
    expect(evenOptions.sections[0].headers?.even).toBeDefined()

    const defaultOnly = await parseToDocxOptions('# Doc', {
      template: { headers: { default: { text: 'Every page' } } },
    })
    expect(defaultOnly.evenAndOddHeaderAndFooters).toBeUndefined()

    const blob = await convertMarkdownToDocx('# Doc', {
      template: { footers: { default: { text: 'Odd' }, even: { text: 'Even' } } },
    })
    const settingsXml = await readDocxEntry(blob, 'word/settings.xml')
    expect(settingsXml).toContain('<w:evenAndOddHeaders')
  })

  it('renders multi-line header text as separate lines', async () => {
    const blob = await convertMarkdownToDocx('# Body\n\nContent.', {
      template: { headers: { default: { text: 'Acme Corp\nConfidential' } } },
    })
    const headerXml = await readDocxEntry(blob, 'word/header1.xml')
    expect(headerXml).toMatch(/Acme Corp<\/w:t>[\s\S]*<w:br\/>[\s\S]*Confidential/u)
    expect(headerXml).not.toMatch(/Acme Corp\nConfidential/u)
  })

  it('throws for invalid titlePage type', async () => {
    await expect(
      parseToDocxOptions('', {
        sections: [
          {
            markdown: '# Bad',
            titlePage: 'yes' as any,
          },
        ],
      }),
    ).rejects.toThrow('Invalid titlePage')
  })

  it('blanks a later section header/footer set to none instead of inheriting it', async () => {
    const blob = await convertMarkdownToDocx('', {
      template: { headers: { default: { text: 'Company' } } },
      sections: [
        { markdown: '# Body' },
        {
          markdown: '# Appendix',
          pageNumbering: { display: 'none' },
          headers: { default: null },
        },
      ],
    })
    const documentXml = await readDocxEntry(blob, 'word/document.xml')
    const relsXml = await readDocxEntry(blob, 'word/_rels/document.xml.rels')
    const sectionProperties = [...documentXml.matchAll(/<w:sectPr>([\s\S]*?)<\/w:sectPr>/gu)].map(
      (match) => match[1],
    )
    expect(sectionProperties).toHaveLength(2)

    async function referencedPart(sectPr: string, kind: 'header' | 'footer'): Promise<string> {
      const relId = new RegExp(`<w:${kind}Reference w:type="default" r:id="([^"]+)"`, 'u').exec(
        sectPr,
      )?.[1]
      expect(relId, `section has a default ${kind} reference`).toBeDefined()
      const target = new RegExp(`Id="${relId}"[^>]*Target="([^"]+)"`, 'u').exec(relsXml)?.[1]
      return readDocxEntry(blob, `word/${target}`)
    }

    const [firstSection, secondSection] = sectionProperties
    expect(await referencedPart(firstSection, 'header')).toContain('Company')
    expect(await referencedPart(firstSection, 'footer')).toContain('PAGE')

    const blankHeader = await referencedPart(secondSection, 'header')
    const blankFooter = await referencedPart(secondSection, 'footer')
    expect(blankHeader).not.toMatch(/<w:t[^>]*>[^<]+<\/w:t>/u)
    expect(blankFooter).not.toMatch(/<w:t[^>]*>[^<]+<\/w:t>/u)
    expect(blankFooter).not.toContain('PAGE')

    const settingsXml = await readDocxEntry(blob, 'word/settings.xml')
    expect(settingsXml).not.toMatch(/<w:evenAndOddHeaders(?: w:val="true")?\/>/u)
  })

  it('only blanks a later even slot when the document uses even pages', async () => {
    const withoutEvenPages = await parseToDocxOptions('', {
      sections: [{ markdown: '# One' }, { markdown: '# Two', headers: { even: null } }],
    })
    expect(withoutEvenPages.evenAndOddHeaderAndFooters).toBeUndefined()
    expect(withoutEvenPages.sections[1].headers?.even).toBeUndefined()

    const withEvenPages = await parseToDocxOptions('', {
      sections: [
        { markdown: '# One', headers: { even: { text: 'Even' } } },
        { markdown: '# Two', headers: { even: null } },
      ],
    })
    expect(withEvenPages.evenAndOddHeaderAndFooters).toBe(true)
    expect(withEvenPages.sections[1].headers?.even).toBeDefined()
    expect(withEvenPages.sections[0].headers?.default).toBeUndefined()
  })
})
