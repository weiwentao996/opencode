import { mkdir, mkdtemp, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { Process } from "@/util/process"

const DEFAULT_TIMEOUT = 120_000

export class PandocError extends Error {
  constructor(message: string, readonly details?: string) {
    super(details ? `${message}\n${details}` : message)
    this.name = "PandocError"
  }
}

function candidates() {
  const configured = process.env.OPENCODE_PANDOC_PATH
  if (configured) return [configured]
  return ["pandoc"]
}

function markdown(input: { title?: string; content: string }) {
  return [input.title ? `# ${input.title}` : undefined, input.content].filter(Boolean).join("\n\n")
}

export async function createWordDocumentWithPandoc(input: {
  outputPath: string
  content: string
  title?: string
  overwrite?: boolean
  abort?: AbortSignal
  timeout?: number
}) {
  if (!input.overwrite && (await Bun.file(input.outputPath).exists())) throw new PandocError(`Output file already exists: ${input.outputPath}`)

  const errors: string[] = []
  for (const executable of candidates()) {
    const dir = await mkdtemp(path.join(os.tmpdir(), "opencode-pandoc-docx-"))
    const source = path.join(dir, "input.md")
    try {
      await mkdir(path.dirname(input.outputPath), { recursive: true })
      await writeFile(source, markdown(input), "utf8")
      const timeout = input.timeout ?? DEFAULT_TIMEOUT
      const abort = input.abort ? AbortSignal.any([input.abort, AbortSignal.timeout(timeout)]) : AbortSignal.timeout(timeout)
      const result = await Process.text(
        [executable, source, "--from", "markdown", "--to", "docx", "--output", input.outputPath],
        { abort, timeout: 5_000, nothrow: true },
      )
      if (result.code === 0 && (await Bun.file(input.outputPath).exists())) {
        return { outputPath: input.outputPath, format: path.extname(input.outputPath), title: input.title, backend: "pandoc" }
      }
      errors.push(`${executable}: ${result.stderr.toString().trim() || result.text.trim() || `exit ${result.code}`}`)
    } catch (cause) {
      errors.push(`${executable}: ${cause instanceof Error ? cause.message : String(cause)}`)
    } finally {
      await rm(dir, { recursive: true, force: true })
    }
  }
  throw new PandocError("Pandoc document generation failed.", errors.filter(Boolean).join("\n"))
}
