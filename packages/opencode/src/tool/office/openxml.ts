import { mkdir, readFile, writeFile } from "node:fs/promises"
import path from "node:path"
import {
  TextReader,
  Uint8ArrayReader,
  Uint8ArrayWriter,
  ZipReader,
  ZipWriter,
} from "@zip.js/zip.js"

export class OpenXmlError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "OpenXmlError"
  }
}

function escapeXml(input: unknown) {
  return String(input ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;")
}

function unescapeXml(input: string) {
  return input
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&")
}

function paragraphs(text: string) {
  return text
    .replaceAll("\r\n", "\n")
    .replaceAll("\r", "\n")
    .split("\n")
    .map((line) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`)
    .join("")
}

async function zip(entries: Record<string, string | Uint8Array>) {
  const writer = new ZipWriter(new Uint8ArrayWriter())
  for (const [filename, content] of Object.entries(entries)) {
    await writer.add(filename, typeof content === "string" ? new TextReader(content) : new Uint8ArrayReader(content))
  }
  return writer.close()
}

async function readZipEntries(filepath: string) {
  const reader = new ZipReader(new Uint8ArrayReader(await readFile(filepath)))
  try {
    const result = new Map<string, Uint8Array>()
    for (const entry of await reader.getEntries()) {
      if (entry.directory) continue
      const data = await entry.getData?.(new Uint8ArrayWriter())
      if (data) result.set(entry.filename, data)
    }
    return result
  } finally {
    await reader.close()
  }
}

function decode(bytes: Uint8Array) {
  return new TextDecoder().decode(bytes)
}

function encode(text: string) {
  return new TextEncoder().encode(text)
}

export async function createWordDocument(input: { outputPath: string; content: string; title?: string; overwrite?: boolean }) {
  if (!input.overwrite && (await Bun.file(input.outputPath).exists())) throw new OpenXmlError(`Output file already exists: ${input.outputPath}`)
  const now = new Date().toISOString()
  const body = `${input.title ? `<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>${escapeXml(input.title)}</w:t></w:r></w:p>` : ""}${paragraphs(input.content)}`
  await mkdir(path.dirname(input.outputPath), { recursive: true })
  await writeFile(
    input.outputPath,
    await zip({
      "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
      "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`,
      "docProps/core.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXml(input.title ?? "")}</dc:title><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`,
      "docProps/app.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>opencode</Application></Properties>`,
      "word/_rels/document.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
      "word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
      "word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style></w:styles>`,
    }),
  )
  return { outputPath: input.outputPath, format: path.extname(input.outputPath), title: input.title }
}

export async function inspectWordDocument(input: { filePath: string; includeText?: boolean; includeTables?: boolean; maxTextChars?: number }) {
  const entries = await readZipEntries(input.filePath)
  const document = entries.get("word/document.xml")
  if (!document) throw new OpenXmlError("Only .docx files with word/document.xml can be inspected without Microsoft Word.")
  const xml = decode(document)
  const text = Array.from(xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g), (match) => unescapeXml(match[1])).join("")
  const result: Record<string, unknown> = {
    filePath: input.filePath,
    name: path.basename(input.filePath),
    fullName: input.filePath,
    characters: text.length,
    words: text.trim() ? text.trim().split(/\s+/).length : 0,
    paragraphs: (xml.match(/<w:p[\s>]/g) ?? []).length,
    tables: (xml.match(/<w:tbl[\s>]/g) ?? []).length,
    headings: [],
  }
  if (input.includeText) result.text = text.slice(0, input.maxTextChars ?? 12000)
  if (input.includeTables) result.tablePreview = Array.from(xml.matchAll(/<w:tbl[\s\S]*?<\/w:tbl>/g), (match, index) => ({
    index: index + 1,
    rows: (match[0].match(/<w:tr[\s>]/g) ?? []).length,
    columns: Math.max(0, ...Array.from(match[0].matchAll(/<w:tr[\s\S]*?<\/w:tr>/g), (row) => (row[0].match(/<w:tc[\s>]/g) ?? []).length)),
  })).slice(0, 20)
  return result
}

export async function replaceWordDocument(input: {
  filePath: string
  replacements: Array<{ find: string; replace: string }>
  saveAs?: string
}) {
  if (path.extname(input.filePath).toLowerCase() !== ".docx") {
    throw new OpenXmlError("LibreOffice fallback can only replace text directly in .docx files.")
  }
  const entries = await readZipEntries(input.filePath)
  const document = entries.get("word/document.xml")
  if (!document) throw new OpenXmlError("Only .docx files with word/document.xml can be modified without Microsoft Word.")

  let xml = decode(document)
  const results = input.replacements.map((item) => {
    if (!item.find) throw new OpenXmlError("Replacement find text cannot be empty.")
    const find = escapeXml(item.find)
    const replace = escapeXml(item.replace)
    const count = xml.split(find).length - 1
    xml = xml.replaceAll(find, replace)
    return { find: item.find, replace: item.replace, count }
  })
  entries.set("word/document.xml", encode(xml))
  const outputPath = input.saveAs ?? input.filePath
  await mkdir(path.dirname(outputPath), { recursive: true })
  await writeFile(outputPath, await zip(Object.fromEntries(entries)))
  return { filePath: input.filePath, outputPath, replacements: results }
}

function workbookSheets(entries: Map<string, Uint8Array>) {
  const workbook = entries.get("xl/workbook.xml")
  if (!workbook) throw new OpenXmlError("Only .xlsx files with xl/workbook.xml can be inspected without Microsoft Excel.")
  const rels = entries.get("xl/_rels/workbook.xml.rels")
  const relations = new Map(
    rels
      ? Array.from(decode(rels).matchAll(/<Relationship\s+([^>]+)>/g), (match) => [
          match[1].match(/Id="([^"]+)"/)?.[1] ?? "",
          match[1].match(/Target="([^"]+)"/)?.[1] ?? "",
        ] as const)
      : [],
  )
  return Array.from(decode(workbook).matchAll(/<sheet\s+([^>]+)>/g), (match, index) => {
    const rel = match[1].match(/r:id="([^"]+)"/)?.[1] ?? ""
    const target = relations.get(rel)
    return {
      name: unescapeXml(match[1].match(/name="([^"]*)"/)?.[1] ?? `Sheet${index + 1}`),
      index: Number(match[1].match(/sheetId="(\d+)"/)?.[1] ?? index + 1),
      visible: match[1].includes('state="hidden"') ? 0 : 1,
      path: target ? path.posix.normalize(path.posix.join("xl", target)) : `xl/worksheets/sheet${index + 1}.xml`,
    }
  })
}

function columnIndex(column: string) {
  return column.toUpperCase().split("").reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0)
}

function parseCellRef(ref: string) {
  const match = ref.match(/^([A-Za-z]+)(\d+)$/)
  if (!match) throw new OpenXmlError(`Invalid cell reference: ${ref}`)
  return { column: columnIndex(match[1]), row: Number(match[2]) }
}

function parseRangeRef(range: string) {
  const [start, end = start] = range.replaceAll("$", "").split(":")
  return { start: parseCellRef(start), end: parseCellRef(end) }
}

function sharedStrings(entries: Map<string, Uint8Array>) {
  const xml = entries.get("xl/sharedStrings.xml")
  if (!xml) return []
  return Array.from(decode(xml).matchAll(/<si[\s\S]*?<\/si>/g), (match) =>
    Array.from(match[0].matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g), (text) => unescapeXml(text[1])).join(""),
  )
}

function cellText(cell: string, strings: string[]) {
  const type = cell.match(/\st="([^"]+)"/)?.[1]
  const value = cell.match(/<v>([\s\S]*?)<\/v>/)?.[1]
  if (type === "s") return strings[Number(value ?? 0)] ?? ""
  if (type === "inlineStr") {
    return Array.from(cell.matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g), (match) => unescapeXml(match[1])).join("")
  }
  return value ? unescapeXml(value) : ""
}

export async function inspectExcelWorkbook(input: { filePath: string; includeSheets?: boolean }) {
  const entries = await readZipEntries(input.filePath)
  const sheets = workbookSheets(entries)
  return {
    filePath: input.filePath,
    name: path.basename(input.filePath),
    fullName: input.filePath,
    sheets: sheets.length,
    hasMacros: [".xlsm", ".xls"].includes(path.extname(input.filePath).toLowerCase()),
    ...(input.includeSheets !== false
      ? { sheetSummary: sheets.map((sheet) => ({ name: sheet.name, index: sheet.index, visible: sheet.visible })) }
      : {}),
  }
}

export async function readExcelRange(input: {
  filePath: string
  sheet: string
  range: string
  includeFormulas?: boolean
  includeNumberFormats?: boolean
  maxCells?: number
}) {
  const entries = await readZipEntries(input.filePath)
  const sheet = workbookSheets(entries).find((item) => item.name === input.sheet)
  if (!sheet) throw new OpenXmlError(`Worksheet not found: ${input.sheet}`)
  const xml = entries.get(sheet.path)
  if (!xml) throw new OpenXmlError(`Worksheet data not found: ${input.sheet}`)

  const bounds = parseRangeRef(input.range)
  const cells = (bounds.end.row - bounds.start.row + 1) * (bounds.end.column - bounds.start.column + 1)
  const maxCells = input.maxCells ?? 1000
  if (cells > maxCells) throw new OpenXmlError(`Requested range has ${cells} cells, which exceeds maxCells=${maxCells}.`)

  const strings = sharedStrings(entries)
  const map = new Map(
    Array.from(decode(xml).matchAll(/<c\s+([^>]*\sr="([^"]+)"[^>]*)>([\s\S]*?)<\/c>/g), (match) => {
      const formula = match[3].match(/<f[^>]*>([\s\S]*?)<\/f>/)?.[1]
      return [
        match[2].replaceAll("$", ""),
        {
          address: match[2].replaceAll("$", ""),
          text: cellText(match[0], strings),
          value: cellText(match[0], strings),
          ...(input.includeFormulas && formula ? { formula: unescapeXml(formula) } : {}),
          ...(input.includeNumberFormats ? { numberFormat: undefined } : {}),
        },
      ] as const
    }),
  )

  const rows = []
  for (let row = bounds.start.row; row <= bounds.end.row; row++) {
    const items = []
    for (let column = bounds.start.column; column <= bounds.end.column; column++) {
      const address = `${columnName(column)}${row}`
      items.push(map.get(address) ?? { address, text: "", value: null })
    }
    rows.push(items)
  }
  return { filePath: input.filePath, sheet: input.sheet, range: input.range.replaceAll("$", ""), rows }
}

function columnName(index: number) {
  let name = ""
  for (let current = index; current > 0; current = Math.floor((current - 1) / 26)) {
    name = String.fromCharCode(((current - 1) % 26) + 65) + name
  }
  return name
}
