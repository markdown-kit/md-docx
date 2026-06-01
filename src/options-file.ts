import type {
  AlignmentOption,
  DocxToMarkdownOptions,
  HeaderFooterGroup,
  HeaderFooterSlot,
  JsonTextReplacement,
  MarkdownCliOptionsFile,
  SectionConfig,
  Style,
} from './types.js'

type JsonObject = Record<string, unknown>

const ALIGNMENT_OPTIONS = [
  'LEFT',
  'CENTER',
  'RIGHT',
  'JUSTIFIED',
] as const satisfies readonly AlignmentOption[]
const DOCUMENT_TYPES = ['document', 'report'] as const
const PAGE_NUMBER_DISPLAYS = [
  'none',
  'current',
  'currentAndTotal',
  'currentAndSectionTotal',
] as const
const PAGE_NUMBER_FORMATS = [
  'decimal',
  'upperRoman',
  'lowerRoman',
  'upperLetter',
  'lowerLetter',
] as const
const PAGE_NUMBER_SEPARATORS = ['hyphen', 'period', 'colon', 'emDash', 'endash'] as const
const SECTION_TYPES = ['NEXT_PAGE', 'NEXT_COLUMN', 'CONTINUOUS', 'EVEN_PAGE', 'ODD_PAGE'] as const
const PAGE_ORIENTATIONS = ['PORTRAIT', 'LANDSCAPE'] as const
const TABLE_LAYOUTS = ['autofit', 'fixed'] as const
const MAMMOTH_KEYS = [
  'styleMap',
  'includeDefaultStyleMap',
  'includeEmbeddedStyleMap',
  'preserveEmptyParagraphs',
] as const
const TURNDOWN_KEYS = [
  'headingStyle',
  'hr',
  'bulletListMarker',
  'codeBlockStyle',
  'fence',
  'emDelimiter',
  'strongDelimiter',
  'linkStyle',
  'linkReferenceStyle',
] as const
const STYLE_KEYS = [
  'titleSize',
  'headingSpacing',
  'paragraphSpacing',
  'lineSpacing',
  'fontFamily',
  'fontFamilly',
  'direction',
  'heading1Size',
  'heading2Size',
  'heading3Size',
  'heading4Size',
  'heading5Size',
  'paragraphSize',
  'listItemSize',
  'codeBlockSize',
  'blockquoteSize',
  'tocFontSize',
  'tocHeading1FontSize',
  'tocHeading2FontSize',
  'tocHeading3FontSize',
  'tocHeading4FontSize',
  'tocHeading5FontSize',
  'tocHeading1Bold',
  'tocHeading2Bold',
  'tocHeading3Bold',
  'tocHeading4Bold',
  'tocHeading5Bold',
  'tocHeading1Italic',
  'tocHeading2Italic',
  'tocHeading3Italic',
  'tocHeading4Italic',
  'tocHeading5Italic',
  'paragraphAlignment',
  'headingAlignment',
  'heading1Alignment',
  'heading2Alignment',
  'heading3Alignment',
  'heading4Alignment',
  'heading5Alignment',
  'blockquoteAlignment',
  'codeBlockAlignment',
  'tableLayout',
] as const
const SECTION_KEYS = [
  'style',
  'page',
  'headers',
  'footers',
  'pageNumbering',
  'titlePage',
  'type',
] as const
const DOCUMENT_SECTION_KEYS = ['markdown', ...SECTION_KEYS] as const
const PAGE_KEYS = ['margin', 'size'] as const
const PAGE_MARGIN_KEYS = ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'] as const
const PAGE_SIZE_KEYS = ['width', 'height', 'orientation'] as const
const PAGE_NUMBERING_KEYS = ['start', 'formatType', 'separator', 'display', 'alignment'] as const
const HEADER_FOOTER_GROUP_KEYS = ['default', 'first', 'even'] as const
const HEADER_FOOTER_SLOT_KEYS = ['text', 'alignment', 'pageNumberDisplay'] as const
const MARKDOWN_CLI_OPTION_KEYS = [
  'documentType',
  'style',
  'template',
  'sections',
  'textReplacements',
] as const
const DOCX_TO_MARKDOWN_OPTION_KEYS = ['mammoth', 'turndown', 'normalizeWhitespace'] as const
const TEXT_REPLACEMENT_KEYS = ['find', 'replace'] as const

function isPlainObject(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function expectPlainObject(value: unknown, context: string): JsonObject {
  if (!isPlainObject(value)) {
    throw new TypeError(`${context} must be a JSON object`)
  }

  return value
}

function assertAllowedKeys(
  value: JsonObject,
  allowedKeys: readonly string[],
  context: string,
): void {
  const allowed = new Set(allowedKeys)

  for (const key of Object.keys(value)) {
    if (!allowed.has(key)) {
      throw new TypeError(`${context}.${key} is not a supported option`)
    }
  }
}

function readOptionalString(value: unknown, context: string): string | undefined {
  if (value === undefined) {
    return undefined
  }

  if (typeof value !== 'string') {
    throw new TypeError(`${context} must be a string`)
  }

  return value
}

function readOptionalBoolean(value: unknown, context: string): boolean | undefined {
  if (value === undefined) {
    return undefined
  }

  if (typeof value !== 'boolean') {
    throw new TypeError(`${context} must be a boolean`)
  }

  return value
}

function readOptionalNumber(value: unknown, context: string): number | undefined {
  if (value === undefined) {
    return undefined
  }

  if (typeof value !== 'number' || Number.isNaN(value)) {
    throw new TypeError(`${context} must be a number`)
  }

  return value
}

function readOptionalStringArray(value: unknown, context: string): string[] | undefined {
  if (value === undefined) {
    return undefined
  }

  if (!Array.isArray(value) || value.some((entry) => typeof entry !== 'string')) {
    throw new TypeError(`${context} must be an array of strings`)
  }

  return [...value]
}

function readOptionalEnum<T extends string>(
  value: unknown,
  allowedValues: readonly T[],
  context: string,
): T | undefined {
  if (value === undefined) {
    return undefined
  }

  if (typeof value !== 'string') {
    throw new TypeError(`${context} must be one of: ${allowedValues.join(', ')}`)
  }

  const match = allowedValues.find((entry) => entry === value)
  if (!match) {
    throw new TypeError(`${context} must be one of: ${allowedValues.join(', ')}`)
  }

  return match
}

function parseStyle(value: unknown, context: string): Partial<Style> | undefined {
  if (value === undefined) {
    return undefined
  }

  const record = expectPlainObject(value, context)
  assertAllowedKeys(record, STYLE_KEYS, context)

  const style: Partial<Style> = {}

  for (const key of [
    'titleSize',
    'headingSpacing',
    'paragraphSpacing',
    'lineSpacing',
    'heading1Size',
    'heading2Size',
    'heading3Size',
    'heading4Size',
    'heading5Size',
    'paragraphSize',
    'listItemSize',
    'codeBlockSize',
    'blockquoteSize',
    'tocFontSize',
    'tocHeading1FontSize',
    'tocHeading2FontSize',
    'tocHeading3FontSize',
    'tocHeading4FontSize',
    'tocHeading5FontSize',
  ] as const) {
    const numericValue = readOptionalNumber(record[key], `${context}.${key}`)
    if (numericValue !== undefined) {
      style[key] = numericValue
    }
  }

  for (const key of [
    'tocHeading1Bold',
    'tocHeading2Bold',
    'tocHeading3Bold',
    'tocHeading4Bold',
    'tocHeading5Bold',
    'tocHeading1Italic',
    'tocHeading2Italic',
    'tocHeading3Italic',
    'tocHeading4Italic',
    'tocHeading5Italic',
  ] as const) {
    const booleanValue = readOptionalBoolean(record[key], `${context}.${key}`)
    if (booleanValue !== undefined) {
      style[key] = booleanValue
    }
  }

  for (const key of ['fontFamily', 'fontFamilly'] as const) {
    const stringValue = readOptionalString(record[key], `${context}.${key}`)
    if (stringValue !== undefined) {
      style[key] = stringValue
    }
  }

  const direction = readOptionalEnum(record.direction, ['LTR', 'RTL'], `${context}.direction`)
  if (direction !== undefined) {
    style.direction = direction
  }

  for (const key of [
    'paragraphAlignment',
    'headingAlignment',
    'heading1Alignment',
    'heading2Alignment',
    'heading3Alignment',
    'heading4Alignment',
    'heading5Alignment',
    'blockquoteAlignment',
    'codeBlockAlignment',
  ] as const) {
    const alignment = readOptionalEnum(record[key], ALIGNMENT_OPTIONS, `${context}.${key}`)
    if (alignment !== undefined) {
      style[key] = alignment
    }
  }

  const tableLayout = readOptionalEnum(record.tableLayout, TABLE_LAYOUTS, `${context}.tableLayout`)
  if (tableLayout !== undefined) {
    style.tableLayout = tableLayout
  }

  return style
}

function parseHeaderFooterSlot(value: unknown, context: string): HeaderFooterSlot | undefined {
  if (value === undefined) {
    return undefined
  }

  if (value === null) {
    return null
  }

  const record = expectPlainObject(value, context)
  assertAllowedKeys(record, HEADER_FOOTER_SLOT_KEYS, context)

  const slot: NonNullable<HeaderFooterSlot> = {}
  const text = readOptionalString(record.text, `${context}.text`)
  if (text !== undefined) {
    slot.text = text
  }

  const alignment = readOptionalEnum(record.alignment, ALIGNMENT_OPTIONS, `${context}.alignment`)
  if (alignment !== undefined) {
    slot.alignment = alignment
  }

  const pageNumberDisplay = readOptionalEnum(
    record.pageNumberDisplay,
    PAGE_NUMBER_DISPLAYS,
    `${context}.pageNumberDisplay`,
  )
  if (pageNumberDisplay !== undefined) {
    slot.pageNumberDisplay = pageNumberDisplay
  }

  return slot
}

function parseHeaderFooterGroup(value: unknown, context: string): HeaderFooterGroup | undefined {
  if (value === undefined) {
    return undefined
  }

  const record = expectPlainObject(value, context)
  assertAllowedKeys(record, HEADER_FOOTER_GROUP_KEYS, context)

  const group: HeaderFooterGroup = {}

  for (const key of HEADER_FOOTER_GROUP_KEYS) {
    const slot = parseHeaderFooterSlot(record[key], `${context}.${key}`)
    if (slot !== undefined) {
      group[key] = slot
    }
  }

  return group
}

function parsePageConfig(value: unknown, context: string): SectionConfig['page'] | undefined {
  if (value === undefined) {
    return undefined
  }

  const record = expectPlainObject(value, context)
  assertAllowedKeys(record, PAGE_KEYS, context)

  const page: NonNullable<SectionConfig['page']> = {}

  if (record.margin !== undefined) {
    const marginRecord = expectPlainObject(record.margin, `${context}.margin`)
    assertAllowedKeys(marginRecord, PAGE_MARGIN_KEYS, `${context}.margin`)
    const margin: NonNullable<NonNullable<SectionConfig['page']>['margin']> = {}

    for (const key of PAGE_MARGIN_KEYS) {
      const numericValue = readOptionalNumber(marginRecord[key], `${context}.margin.${key}`)
      if (numericValue !== undefined) {
        margin[key] = numericValue
      }
    }

    page.margin = margin
  }

  if (record.size !== undefined) {
    const sizeRecord = expectPlainObject(record.size, `${context}.size`)
    assertAllowedKeys(sizeRecord, PAGE_SIZE_KEYS, `${context}.size`)
    const size: NonNullable<NonNullable<SectionConfig['page']>['size']> = {}

    const width = readOptionalNumber(sizeRecord.width, `${context}.size.width`)
    if (width !== undefined) {
      size.width = width
    }

    const height = readOptionalNumber(sizeRecord.height, `${context}.size.height`)
    if (height !== undefined) {
      size.height = height
    }

    const orientation = readOptionalEnum(
      sizeRecord.orientation,
      PAGE_ORIENTATIONS,
      `${context}.size.orientation`,
    )
    if (orientation !== undefined) {
      size.orientation = orientation
    }

    page.size = size
  }

  return page
}

function parsePageNumbering(
  value: unknown,
  context: string,
): SectionConfig['pageNumbering'] | undefined {
  if (value === undefined) {
    return undefined
  }

  const record = expectPlainObject(value, context)
  assertAllowedKeys(record, PAGE_NUMBERING_KEYS, context)

  const pageNumbering: NonNullable<SectionConfig['pageNumbering']> = {}
  const start = readOptionalNumber(record.start, `${context}.start`)
  if (start !== undefined) {
    pageNumbering.start = start
  }

  const formatType = readOptionalEnum(
    record.formatType,
    PAGE_NUMBER_FORMATS,
    `${context}.formatType`,
  )
  if (formatType !== undefined) {
    pageNumbering.formatType = formatType
  }

  const separator = readOptionalEnum(
    record.separator,
    PAGE_NUMBER_SEPARATORS,
    `${context}.separator`,
  )
  if (separator !== undefined) {
    pageNumbering.separator = separator
  }

  const display = readOptionalEnum(record.display, PAGE_NUMBER_DISPLAYS, `${context}.display`)
  if (display !== undefined) {
    pageNumbering.display = display
  }

  const alignment = readOptionalEnum(record.alignment, ALIGNMENT_OPTIONS, `${context}.alignment`)
  if (alignment !== undefined) {
    pageNumbering.alignment = alignment
  }

  return pageNumbering
}

function parseSectionConfig(value: unknown, context: string): SectionConfig | undefined {
  if (value === undefined) {
    return undefined
  }

  const record = expectPlainObject(value, context)
  assertAllowedKeys(record, SECTION_KEYS, context)

  const config: SectionConfig = {}

  const style = parseStyle(record.style, `${context}.style`)
  if (style !== undefined) {
    config.style = style
  }

  const page = parsePageConfig(record.page, `${context}.page`)
  if (page !== undefined) {
    config.page = page
  }

  const headers = parseHeaderFooterGroup(record.headers, `${context}.headers`)
  if (headers !== undefined) {
    config.headers = headers
  }

  const footers = parseHeaderFooterGroup(record.footers, `${context}.footers`)
  if (footers !== undefined) {
    config.footers = footers
  }

  const pageNumbering = parsePageNumbering(record.pageNumbering, `${context}.pageNumbering`)
  if (pageNumbering !== undefined) {
    config.pageNumbering = pageNumbering
  }

  const titlePage = readOptionalBoolean(record.titlePage, `${context}.titlePage`)
  if (titlePage !== undefined) {
    config.titlePage = titlePage
  }

  const type = readOptionalEnum(record.type, SECTION_TYPES, `${context}.type`)
  if (type !== undefined) {
    config.type = type
  }

  return config
}

function parseJsonTextReplacement(value: unknown, context: string): JsonTextReplacement {
  const record = expectPlainObject(value, context)
  assertAllowedKeys(record, TEXT_REPLACEMENT_KEYS, context)

  if (typeof record.find !== 'string') {
    throw new TypeError(`${context}.find must be a string`)
  }

  if (typeof record.replace !== 'string') {
    throw new TypeError(`${context}.replace must be a string`)
  }

  return {
    find: record.find,
    replace: record.replace,
  }
}

export function parseMarkdownCliOptionsFile(
  value: unknown,
  context = 'options',
): MarkdownCliOptionsFile {
  const record = expectPlainObject(value, context)
  assertAllowedKeys(record, MARKDOWN_CLI_OPTION_KEYS, context)

  const options: MarkdownCliOptionsFile = {}

  const documentType = readOptionalEnum(
    record.documentType,
    DOCUMENT_TYPES,
    `${context}.documentType`,
  )
  if (documentType !== undefined) {
    options.documentType = documentType
  }

  const style = parseStyle(record.style, `${context}.style`)
  if (style !== undefined) {
    options.style = style
  }

  const template = parseSectionConfig(record.template, `${context}.template`)
  if (template !== undefined) {
    options.template = template
  }

  if (record.sections !== undefined) {
    if (!Array.isArray(record.sections)) {
      throw new TypeError(`${context}.sections must be an array`)
    }

    options.sections = record.sections.map((section, index) => {
      const sectionRecord = expectPlainObject(section, `${context}.sections[${index}]`)
      assertAllowedKeys(sectionRecord, DOCUMENT_SECTION_KEYS, `${context}.sections[${index}]`)

      if (typeof sectionRecord.markdown !== 'string') {
        throw new TypeError(`${context}.sections[${index}].markdown must be a string`)
      }

      const { markdown: _markdown, ...sectionConfigRecord } = sectionRecord
      const sectionConfig =
        parseSectionConfig(sectionConfigRecord, `${context}.sections[${index}]`) ?? {}
      return {
        ...sectionConfig,
        markdown: sectionRecord.markdown,
      }
    })
  }

  if (record.textReplacements !== undefined) {
    if (!Array.isArray(record.textReplacements)) {
      throw new TypeError(`${context}.textReplacements must be an array`)
    }

    options.textReplacements = record.textReplacements.map((entry, index) =>
      parseJsonTextReplacement(entry, `${context}.textReplacements[${index}]`),
    )
  }

  return options
}

export function parseDocxToMarkdownOptionsFile(
  value: unknown,
  context = 'options',
): DocxToMarkdownOptions {
  const record = expectPlainObject(value, context)
  assertAllowedKeys(record, DOCX_TO_MARKDOWN_OPTION_KEYS, context)

  const options: DocxToMarkdownOptions = {}

  if (record.mammoth !== undefined) {
    const mammothRecord = expectPlainObject(record.mammoth, `${context}.mammoth`)
    assertAllowedKeys(mammothRecord, MAMMOTH_KEYS, `${context}.mammoth`)

    options.mammoth = {
      styleMap: readOptionalStringArray(mammothRecord.styleMap, `${context}.mammoth.styleMap`),
      includeDefaultStyleMap: readOptionalBoolean(
        mammothRecord.includeDefaultStyleMap,
        `${context}.mammoth.includeDefaultStyleMap`,
      ),
      includeEmbeddedStyleMap: readOptionalBoolean(
        mammothRecord.includeEmbeddedStyleMap,
        `${context}.mammoth.includeEmbeddedStyleMap`,
      ),
      preserveEmptyParagraphs: readOptionalBoolean(
        mammothRecord.preserveEmptyParagraphs,
        `${context}.mammoth.preserveEmptyParagraphs`,
      ),
    }
  }

  if (record.turndown !== undefined) {
    const turndownRecord = expectPlainObject(record.turndown, `${context}.turndown`)
    assertAllowedKeys(turndownRecord, TURNDOWN_KEYS, `${context}.turndown`)

    options.turndown = {
      headingStyle: readOptionalEnum(
        turndownRecord.headingStyle,
        ['setext', 'atx'],
        `${context}.turndown.headingStyle`,
      ),
      hr: readOptionalString(turndownRecord.hr, `${context}.turndown.hr`),
      bulletListMarker: readOptionalEnum(
        turndownRecord.bulletListMarker,
        ['-', '*', '+'],
        `${context}.turndown.bulletListMarker`,
      ),
      codeBlockStyle: readOptionalEnum(
        turndownRecord.codeBlockStyle,
        ['indented', 'fenced'],
        `${context}.turndown.codeBlockStyle`,
      ),
      fence: readOptionalEnum(turndownRecord.fence, ['```', '~~~'], `${context}.turndown.fence`),
      emDelimiter: readOptionalEnum(
        turndownRecord.emDelimiter,
        ['_', '*'],
        `${context}.turndown.emDelimiter`,
      ),
      strongDelimiter: readOptionalEnum(
        turndownRecord.strongDelimiter,
        ['**', '__'],
        `${context}.turndown.strongDelimiter`,
      ),
      linkStyle: readOptionalEnum(
        turndownRecord.linkStyle,
        ['inlined', 'referenced'],
        `${context}.turndown.linkStyle`,
      ),
      linkReferenceStyle: readOptionalEnum(
        turndownRecord.linkReferenceStyle,
        ['full', 'collapsed', 'shortcut'],
        `${context}.turndown.linkReferenceStyle`,
      ),
    }
  }

  const normalizeWhitespace = readOptionalBoolean(
    record.normalizeWhitespace,
    `${context}.normalizeWhitespace`,
  )
  if (normalizeWhitespace !== undefined) {
    options.normalizeWhitespace = normalizeWhitespace
  }

  return options
}
