import { describe, expect, test } from "bun:test"
import { readFile, writeFile } from "node:fs/promises"
import { Result, Schema } from "effect"
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
