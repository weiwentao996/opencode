import { mkdir, mkdtemp, readdir, rm } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { pathToFileURL } from "node:url"
import { Process } from "@/util/process"

const OFFICE_EXTENSIONS = new Set([
  ".doc",
  ".docx",
  ".docm",
  ".xls",
  ".xlsx",
  ".xlsm",
  ".ppt",
  ".pptx",
  ".pptm",
  ".odt",
  ".ods",
  ".odp",
])

const DEFAULT_TIMEOUT = 120_000

export class LibreOfficeError extends Error {
  constructor(message: string, readonly details?: string) {
    super(details ? `${message}\n${details}` : message)
    this.name = "LibreOfficeError"
  }
}

export function isOfficeDocument(filepath: string) {
  return OFFICE_EXTENSIONS.has(path.extname(filepath).toLowerCase())
}

function candidates() {
  const configured = process.env.OPENCODE_LIBREOFFICE_PATH
  if (configured) return [configured]
  if (process.platform === "darwin") {
    return [
      "soffice",
      "libreoffice",
      "/Applications/LibreOffice.app/Contents/MacOS/soffice",
      path.join(os.homedir(), "Applications/LibreOffice.app/Contents/MacOS/soffice"),
    ]
  }
  if (process.platform === "win32") {
    return [
      "soffice.exe",
      "libreoffice.exe",
      "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
      "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe",
    ]
  }
  return ["soffice", "libreoffice"]
}

function outputName(filepath: string, extension: string) {
  return path.basename(filepath, path.extname(filepath)) + extension
}

function conversionFailure(errors: string[]) {
  return new LibreOfficeError(
    "LibreOffice is required to read Office documents. Install LibreOffice or set OPENCODE_LIBREOFFICE_PATH to the soffice executable.",
    errors.filter(Boolean).join("\n"),
  )
}

export async function convertOfficeDocument(
  filepath: string,
  target: { format: string; extension: string },
  options: { abort?: AbortSignal; timeout?: number } = {},
) {
  const errors: string[] = []
  for (const executable of candidates()) {
    const dir = await mkdtemp(path.join(os.tmpdir(), "opencode-office-convert-"))
    const profile = path.join(dir, "profile")
    const outdir = path.join(dir, "out")
    try {
      await mkdir(outdir, { recursive: true })
      const timeout = options.timeout ?? DEFAULT_TIMEOUT
      const abort = options.abort ? AbortSignal.any([options.abort, AbortSignal.timeout(timeout)]) : AbortSignal.timeout(timeout)
      const result = await Process.text(
        [
          executable,
          "--headless",
          "--nologo",
          "--nofirststartwizard",
          "--nodefault",
          "--norestore",
          `-env:UserInstallation=${pathToFileURL(profile).href}`,
          "--convert-to",
          target.format,
          "--outdir",
          outdir,
          filepath,
        ],
        { abort, timeout: 5_000, nothrow: true },
      )
      if (result.code !== 0) {
        errors.push(`${executable}: ${result.stderr.toString().trim() || result.text.trim() || `exit ${result.code}`}`)
        await rm(dir, { recursive: true, force: true })
        continue
      }

      const expected = path.join(outdir, outputName(filepath, target.extension))
      const outputs = await readdir(outdir).catch(() => [] as string[])
      const fallback = outputs.find((item) => item.toLowerCase().endsWith(target.extension))
      const output = outputs.includes(path.basename(expected)) ? expected : fallback ? path.join(outdir, fallback) : undefined
      if (!output) {
        errors.push(`${executable}: conversion completed but no ${target.extension} file was produced`)
        await rm(dir, { recursive: true, force: true })
        continue
      }

      return { filepath: output, cleanup: () => rm(dir, { recursive: true, force: true }) }
    } catch (cause) {
      errors.push(`${executable}: ${cause instanceof Error ? cause.message : String(cause)}`)
      await rm(dir, { recursive: true, force: true })
    }
  }
  throw conversionFailure(errors)
}

export function convertOfficeToPdf(filepath: string, options: { abort?: AbortSignal; timeout?: number } = {}) {
  return convertOfficeDocument(filepath, { format: "pdf", extension: ".pdf" }, options)
}

export function convertOfficeToText(filepath: string, options: { abort?: AbortSignal; timeout?: number } = {}) {
  return convertOfficeDocument(filepath, { format: "txt", extension: ".txt" }, options)
}
