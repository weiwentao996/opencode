import { afterEach, describe, expect, test } from "bun:test"
import { chmod, mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { Result, Schema } from "effect"
import { runLibreOfficeAutomation } from "../../src/tool/office/libreoffice-runner"
import { OfficeAutomationError, runOfficeAutomation } from "../../src/tool/office/powershell"
import {
  ExcelRangeParameters,
  ExcelWriteParameters,
  WordReplaceParameters,
} from "../../src/tool/office/schema"

const accepts = (schema: Schema.Decoder<unknown>, input: unknown): boolean =>
  Result.isSuccess(Schema.decodeUnknownResult(schema)(input))

const parse = <S extends Schema.Decoder<unknown>>(schema: S, input: unknown): S["Type"] =>
  Schema.decodeUnknownSync(schema)(input)

const dirs: string[] = []
afterEach(async () => {
  await Promise.all(dirs.splice(0).map((dir) => rm(dir, { recursive: true, force: true })))
})

async function tmpdir() {
  const dir = await mkdtemp(path.join(os.tmpdir(), "opencode-office-test-"))
  dirs.push(dir)
  return dir
}

async function withEnv<T>(values: Record<string, string>, fn: () => Promise<T>) {
  const previous = Object.fromEntries(Object.keys(values).map((key) => [key, process.env[key]]))
  Object.assign(process.env, values)
  try {
    return await fn()
  } finally {
    for (const [key, value] of Object.entries(previous)) {
      if (value === undefined) delete process.env[key]
      else process.env[key] = value
    }
  }
}

async function fakePandoc(dir: string) {
  if (process.platform === "win32") {
    const script = path.join(dir, "pandoc.cmd")
    await writeFile(
      script,
      [
        "@echo off",
        "set output=",
        ":loop",
        "if \"%~1\"==\"\" goto done",
        "if \"%~1\"==\"--output\" (",
        "  shift",
        "  set output=%~1",
        ")",
        "shift",
        "goto loop",
        ":done",
        "copy /Y \"%OPENCODE_FAKE_PANDOC_DOCX%\" \"%output%\" >NUL",
      ].join("\r\n"),
      "utf8",
    )
    return script
  }

  const script = path.join(dir, "pandoc")
  await writeFile(
    script,
    [
      "#!/bin/sh",
      "output=",
      "prev=",
      "for arg in \"$@\"; do",
      "  if [ \"$prev\" = \"--output\" ]; then output=\"$arg\"; fi",
      "  prev=\"$arg\"",
      "done",
      "cp \"$OPENCODE_FAKE_PANDOC_DOCX\" \"$output\"",
    ].join("\n"),
    "utf8",
  )
  await chmod(script, 0o755)
  return script
}

async function fakeSoffice(dir: string) {
  if (process.platform === "win32") {
    const script = path.join(dir, "soffice.cmd")
    await writeFile(
      script,
      [
        "@echo off",
        "set outdir=",
        "set input=",
        ":loop",
        "if \"%~1\"==\"\" goto done",
        "if \"%~1\"==\"--outdir\" (",
        "  shift",
        "  set outdir=%~1",
        ")",
        "set input=%~1",
        "shift",
        "goto loop",
        ":done",
        "if not exist \"%outdir%\" mkdir \"%outdir%\"",
        "for %%F in (\"%input%\") do set name=%%~nF",
        "copy /Y \"%OPENCODE_FAKE_LIBREOFFICE_OUTPUT%\" \"%outdir%\\%name%.pdf\" >NUL",
      ].join("\r\n"),
      "utf8",
    )
    return script
  }

  const script = path.join(dir, "soffice")
  await writeFile(
    script,
    [
      "#!/bin/sh",
      "outdir=",
      "input=",
      "prev=",
      "for arg in \"$@\"; do",
      "  if [ \"$prev\" = \"--outdir\" ]; then outdir=\"$arg\"; fi",
      "  input=\"$arg\"",
      "  prev=\"$arg\"",
      "done",
      "mkdir -p \"$outdir\"",
      "base=$(basename \"$input\")",
      "name=${base%.*}.pdf",
      "cp \"$OPENCODE_FAKE_LIBREOFFICE_OUTPUT\" \"$outdir/$name\"",
    ].join("\n"),
    "utf8",
  )
  await chmod(script, 0o755)
  return script
}

describe("office tool parameters", () => {
  test("word replacements require the top-level fields", () => {
    expect(accepts(WordReplaceParameters, { filePath: "a.docx", replacements: [] })).toBe(true)
    expect(accepts(WordReplaceParameters, { filePath: "a.docx" })).toBe(false)
    expect(accepts(WordReplaceParameters, { replacements: [] })).toBe(false)
  })

  test("excel range accepts A1 ranges and rejects malformed ranges", () => {
    expect(parse(ExcelRangeParameters, { filePath: "a.xlsx", sheet: "Sheet1", range: "A1:D20" }).range).toBe("A1:D20")
    expect(accepts(ExcelRangeParameters, { filePath: "a.xlsx", sheet: "Sheet1", range: "$A$1:$D$20" })).toBe(true)
    expect(accepts(ExcelRangeParameters, { filePath: "a.xlsx", sheet: "Sheet1", range: "A:" })).toBe(false)
    expect(accepts(ExcelRangeParameters, { filePath: "a.xlsx", sheet: "Sheet1", range: "Sheet1!A1" })).toBe(false)
  })

  test("excel write requires a matrix and start cell", () => {
    expect(
      accepts(ExcelWriteParameters, {
        filePath: "a.xlsx",
        sheet: "Sheet1",
        startCell: "B2",
        values: [[1, "=A1+1"]],
      }),
    ).toBe(true)
    expect(accepts(ExcelWriteParameters, { filePath: "a.xlsx", sheet: "Sheet1", values: [[1]] })).toBe(false)
    expect(accepts(ExcelWriteParameters, { filePath: "a.xlsx", sheet: "Sheet1", startCell: "B:" })).toBe(false)
  })
})

describe("office libreoffice backend", () => {
  test("creates and inspects Word documents with automatic fallback", async () => {
    const dir = await tmpdir()
    const outputPath = path.join(dir, "generated.docx")
    await withEnv({ OPENCODE_PANDOC_PATH: path.join(dir, "missing-pandoc") }, () =>
      runLibreOfficeAutomation({
        operation: "word_create",
        outputPath,
        title: "Plan",
        content: "First line\nSecond line",
        overwrite: true,
      }),
    )

    const result = await runLibreOfficeAutomation({
      operation: "word_inspect",
      filePath: outputPath,
      includeText: true,
      includeTables: true,
    })
    expect(result).toMatchObject({ name: "generated.docx", paragraphs: 3, tables: 0 })
    expect((result as { text: string }).text).toContain("First line")
  })

  test("creates Word documents with Pandoc when available", async () => {
    const dir = await tmpdir()
    const template = path.join(dir, "template.docx")
    const outputPath = path.join(dir, "pandoc.docx")
    await withEnv({ OPENCODE_PANDOC_PATH: path.join(dir, "missing-pandoc") }, () =>
      runLibreOfficeAutomation({ operation: "word_create", outputPath: template, content: "from pandoc", overwrite: true }),
    )
    const pandoc = await fakePandoc(dir)

    const result = await withEnv({ OPENCODE_PANDOC_PATH: pandoc, OPENCODE_FAKE_PANDOC_DOCX: template }, () =>
      runLibreOfficeAutomation({ operation: "word_create", outputPath, content: "ignored", overwrite: true }),
    )
    expect(result).toMatchObject({ outputPath, backend: "pandoc" })
    const inspected = await runLibreOfficeAutomation({ operation: "word_inspect", filePath: outputPath, includeText: true })
    expect((inspected as { text: string }).text).toContain("from pandoc")
  })

  test("replaces text in .docx OpenXML", async () => {
    const dir = await tmpdir()
    const filePath = path.join(dir, "replace.docx")
    await runLibreOfficeAutomation({ operation: "word_create", outputPath: filePath, content: "alpha beta", overwrite: true })

    const result = await runLibreOfficeAutomation({
      operation: "word_replace",
      filePath,
      replacements: [{ find: "beta", replace: "gamma" }],
    })
    expect(result).toMatchObject({ outputPath: filePath })

    const inspected = await runLibreOfficeAutomation({ operation: "word_inspect", filePath, includeText: true })
    expect((inspected as { text: string }).text).toContain("alpha gamma")
  })

  test("exports PDF through configured soffice", async () => {
    const dir = await tmpdir()
    const input = path.join(dir, "input.docx")
    const sourcePdf = path.join(dir, "source.pdf")
    const outputPath = path.join(dir, "output.pdf")
    await writeFile(input, "office")
    await writeFile(sourcePdf, "%PDF-1.4\n%%EOF\n")
    const soffice = await fakeSoffice(dir)

    await withEnv(
      { OPENCODE_LIBREOFFICE_PATH: soffice, OPENCODE_FAKE_LIBREOFFICE_OUTPUT: sourcePdf },
      async () => {
        const result = await runLibreOfficeAutomation({
          operation: "word_export_pdf",
          filePath: input,
          outputPath,
          overwrite: true,
        })
        expect(result).toMatchObject({ filePath: input, outputPath, format: "pdf" })
      },
    )
    expect(await readFile(outputPath, "utf8")).toContain("%PDF")
  })
})

describe("office powershell bridge", () => {
  test("parses successful JSON responses from the runner", async () => {
    if (process.platform !== "win32") return
    const result = await runOfficeAutomation(
      { operation: "excel_inspect", filePath: "C:/tmp/book.xlsx" },
      {
        runner: async (_cmd, options) => {
          const response = options?.env?.OPENCODE_OFFICE_RESPONSE
          const request = options?.env?.OPENCODE_OFFICE_REQUEST
          expect(response).toEqual(expect.any(String))
          expect(request).toEqual(expect.any(String))
          const body = JSON.parse(await readFile(String(request), "utf8"))
          expect(body.operation).toBe("excel_inspect")
          await writeFile(String(response), "﻿" + JSON.stringify({ ok: true, result: { sheets: 2 } }), "utf8")
          return { text: "", code: 0 }
        },
      },
    )
    expect(result).toEqual({ sheets: 2 })
  })

  test("maps failed JSON responses to OfficeAutomationError", async () => {
    if (process.platform !== "win32") return
    await expect(
      runOfficeAutomation(
        { operation: "word_inspect", filePath: "C:/tmp/doc.docx" },
        {
          runner: async (_cmd, options) => {
            await writeFile(
              String(options?.env?.OPENCODE_OFFICE_RESPONSE),
              JSON.stringify({ ok: false, error: "Word failed", details: "COM error" }),
              "utf8",
            )
            return { text: "", code: 1 }
          },
        },
      ),
    ).rejects.toBeInstanceOf(OfficeAutomationError)
  })
})
