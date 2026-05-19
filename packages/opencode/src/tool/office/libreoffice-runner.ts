import { copyFile, mkdir } from "node:fs/promises"
import path from "node:path"
import { convertOfficeToPdf, convertOfficeToText } from "./libreoffice"
import { createWordDocumentWithPandoc } from "./pandoc"
import {
  createWordDocument,
  inspectExcelWorkbook,
  inspectWordDocument,
  readExcelRange,
  replaceWordDocument,
  OpenXmlError,
} from "./openxml"
import type { OfficeRequest } from "./powershell"

export class LibreOfficeAutomationError extends Error {
  constructor(message: string, readonly details?: string) {
    super(details ? `${message}\n${details}` : message)
    this.name = "LibreOfficeAutomationError"
  }
}

async function ensureOutput(outputPath: string, overwrite: unknown) {
  if (!overwrite && (await Bun.file(outputPath).exists())) throw new LibreOfficeAutomationError(`Output file already exists: ${outputPath}`)
  await mkdir(path.dirname(outputPath), { recursive: true })
}

async function exportPdf(filePath: string, outputPath: string, overwrite: unknown, options: { abort?: AbortSignal; timeout?: number }) {
  await ensureOutput(outputPath, overwrite)
  const converted = await convertOfficeToPdf(filePath, options)
  try {
    await copyFile(converted.filepath, outputPath)
    return { filePath, outputPath, format: "pdf" }
  } finally {
    await converted.cleanup()
  }
}

async function createBestWordDocument(request: OfficeRequest, options: { abort?: AbortSignal }) {
  const input = {
    outputPath: String(request.outputPath),
    content: String(request.content ?? ""),
    title: typeof request.title === "string" ? request.title : undefined,
    overwrite: Boolean(request.overwrite),
    timeout: request.timeout,
    abort: options.abort,
  }
  try {
    return await createWordDocumentWithPandoc(input)
  } catch (cause) {
    if (cause instanceof Error && cause.message.includes("already exists")) throw cause
    return await createWordDocument(input)
  }
}

async function inspectTextDocument(request: OfficeRequest, options: { abort?: AbortSignal; timeout?: number }) {
  const converted = await convertOfficeToText(String(request.filePath), options)
  try {
    const text = await Bun.file(converted.filepath).text()
    const truncated = text.slice(0, Number(request.maxTextChars ?? 12000))
    return {
      filePath: request.filePath,
      name: path.basename(String(request.filePath)),
      fullName: request.filePath,
      characters: text.length,
      words: text.trim() ? text.trim().split(/\s+/).length : 0,
      paragraphs: text.split(/\n\s*\n/).filter(Boolean).length,
      tables: 0,
      headings: [],
      ...(request.includeText ? { text: truncated } : {}),
      note: "Inspected via LibreOffice text conversion; comments, revisions, and detailed layout metadata are unavailable on this platform.",
    }
  } finally {
    await converted.cleanup()
  }
}

function unsupported(operation: string, details: string) {
  throw new LibreOfficeAutomationError(`${operation} is not supported by the LibreOffice backend.`, details)
}

export async function runLibreOfficeAutomation(request: OfficeRequest, options: { abort?: AbortSignal } = {}) {
  const timeout = request.timeout
  try {
    switch (request.operation) {
      case "word_create":
        return await createBestWordDocument(request, options)
      case "word_inspect":
        if (path.extname(String(request.filePath)).toLowerCase() === ".docx") {
          return await inspectWordDocument({
            filePath: String(request.filePath),
            includeText: Boolean(request.includeText),
            includeTables: Boolean(request.includeTables),
            maxTextChars: Number(request.maxTextChars ?? 12000),
          })
        }
        return await inspectTextDocument(request, { abort: options.abort, timeout })
      case "word_export_pdf":
        return await exportPdf(String(request.filePath), String(request.outputPath), request.overwrite, {
          abort: options.abort,
          timeout,
        })
      case "word_replace":
        return await replaceWordDocument({
          filePath: String(request.filePath),
          replacements: (request.replacements as Array<{ find: string; replace: string }> | undefined) ?? [],
          saveAs: typeof request.saveAs === "string" ? request.saveAs : undefined,
        })
      case "excel_inspect":
        if (![".xlsx", ".xlsm"].includes(path.extname(String(request.filePath)).toLowerCase())) {
          unsupported("excel_inspect", "Only .xlsx/.xlsm workbook inspection is available without Microsoft Excel.")
        }
        return await inspectExcelWorkbook({ filePath: String(request.filePath), includeSheets: request.includeSheets !== false })
      case "excel_export_pdf":
        if (request.sheet) unsupported("excel_export_pdf", "Sheet-specific export requires Microsoft Excel on Windows.")
        return await exportPdf(String(request.filePath), String(request.outputPath), request.overwrite, {
          abort: options.abort,
          timeout,
        })
      case "excel_range":
        if (![".xlsx", ".xlsm"].includes(path.extname(String(request.filePath)).toLowerCase())) {
          unsupported("excel_range", "Only .xlsx/.xlsm range reads are available without Microsoft Excel.")
        }
        return await readExcelRange({
          filePath: String(request.filePath),
          sheet: String(request.sheet),
          range: String(request.range),
          includeFormulas: Boolean(request.includeFormulas),
          includeNumberFormats: Boolean(request.includeNumberFormats),
          maxCells: Number(request.maxCells ?? 1000),
        })
      case "excel_write":
        unsupported("excel_write", "Writing Excel cell matrices while preserving workbook fidelity requires Microsoft Excel on Windows.")
      case "excel_recalculate":
        unsupported("excel_recalculate", "Recalculating formulas requires Microsoft Excel on Windows.")
    }
  } catch (cause) {
    if (cause instanceof LibreOfficeAutomationError) throw cause
    if (cause instanceof OpenXmlError) throw new LibreOfficeAutomationError(cause.message)
    throw new LibreOfficeAutomationError(
      `LibreOffice backend failed while running ${request.operation}.`,
      cause instanceof Error ? cause.message : String(cause),
    )
  }
}
