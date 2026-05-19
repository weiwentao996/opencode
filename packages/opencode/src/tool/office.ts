import path from "node:path"
import { Effect, Schema } from "effect"
import { AppFileSystem } from "@opencode-ai/core/filesystem"
import { Bus } from "@/bus"
import { File } from "@/file"
import { FileWatcher } from "@/file/watcher"
import { InstanceState } from "@/effect/instance-state"
import type { InstanceContext } from "@/project/instance"
import { assertExternalDirectoryEffect } from "./external-directory"
import * as Tool from "./tool"
import { runLibreOfficeAutomation } from "./office/libreoffice-runner"
import { runOfficeAutomation, type OfficeOperation } from "./office/powershell"
import {
  ExcelExportPdfParameters,
  ExcelInspectParameters,
  ExcelRangeParameters,
  ExcelRecalculateParameters,
  ExcelWriteParameters,
  WordCreateParameters,
  WordExportPdfParameters,
  WordInspectParameters,
  WordReplaceParameters,
  type ExcelExportPdfInput,
  type ExcelInspectInput,
  type ExcelRangeInput,
  type ExcelRecalculateInput,
  type ExcelWriteInput,
  type WordCreateInput,
  type WordExportPdfInput,
  type WordInspectInput,
  type WordReplaceInput,
} from "./office/schema"

function resultOutput(result: unknown) {
  return JSON.stringify(result, null, 2)
}

type RuntimePath = {
  absolute: string
  relative: string
}

function resolveRuntimePath(input: string, instance: InstanceContext): RuntimePath {
  const absolute = AppFileSystem.normalizePath(
    AppFileSystem.windowsPath(path.isAbsolute(input) ? input : path.join(instance.directory, input)),
  )
  return {
    absolute,
    relative: path.relative(instance.worktree, absolute),
  }
}

function askRead(ctx: Tool.Context, filepath: RuntimePath, operation: string) {
  return ctx.ask({
    permission: "read",
    patterns: [filepath.relative],
    always: ["*"],
    metadata: {
      operation,
      filepath: filepath.absolute,
    },
  })
}

function askEdit(ctx: Tool.Context, patterns: RuntimePath[], operation: string, metadata: Record<string, unknown>) {
  return ctx.ask({
    permission: "edit",
    patterns: patterns.map((item) => item.relative),
    always: ["*"],
    metadata: {
      operation,
      ...metadata,
    },
  })
}

function askMacros(ctx: Tool.Context, filepath: RuntimePath, operation: string, allowMacros: boolean | undefined) {
  if (!allowMacros) return Effect.void
  return ctx.ask({
    permission: "office_macro",
    patterns: [filepath.relative],
    always: [],
    metadata: {
      operation,
      filepath: filepath.absolute,
      warning: "Opening Office documents with macros enabled can execute document code.",
    },
  })
}

function run(operation: OfficeOperation, input: Record<string, unknown>, ctx: Tool.Context) {
  return Effect.tryPromise({
    try: () =>
      process.platform === "win32"
        ? runOfficeAutomation({ operation, ...input }, { abort: ctx.abort })
        : runLibreOfficeAutomation({ operation, ...input }, { abort: ctx.abort }),
    catch: (cause) => cause,
  })
}

function publishUpdated(bus: Bus.Interface, filepath: string, existed: boolean) {
  return Effect.gen(function* () {
    yield* bus.publish(File.Event.Edited, { file: filepath })
    yield* bus.publish(FileWatcher.Event.Updated, {
      file: filepath,
      event: existed ? "change" : "add",
    })
  })
}

export const OfficeWordCreateTool = Tool.define(
  "office_word_create",
  Effect.gen(function* () {
    const fs = yield* AppFileSystem.Service
    const bus = yield* Bus.Service
    return {
      description:
        "Create a new editable Word document using the best available backend for the current platform.",
      parameters: WordCreateParameters,
      execute: (params: WordCreateInput, ctx: Tool.Context) =>
        Effect.gen(function* () {
          const instance = yield* InstanceState.context
          const output = resolveRuntimePath(params.outputPath, instance)
          yield* assertExternalDirectoryEffect(ctx, output.absolute)
          yield* askEdit(ctx, [output], "office_word_create", { outputPath: output.absolute, title: params.title })
          const existed = yield* fs.existsSafe(output.absolute)
          const result = yield* run(
            "word_create",
            {
              ...params,
              outputPath: output.absolute,
            },
            ctx,
          )
          yield* publishUpdated(bus, output.absolute, existed)
          return {
            title: output.relative,
            metadata: { outputPath: output.absolute },
            output: resultOutput(result),
          }
        }).pipe(Effect.orDie),
    }
  }),
)

export const OfficeWordInspectTool = Tool.define(
  "office_word_inspect",
  Effect.succeed({
    description:
      "Inspect a Word .doc or .docx document. Uses Microsoft Word automation on Windows; macOS/Linux use OpenXML or LibreOffice text conversion with reduced metadata fidelity.",
    parameters: WordInspectParameters,
    execute: (params: WordInspectInput, ctx: Tool.Context) =>
      Effect.gen(function* () {
        const instance = yield* InstanceState.context
        const file = resolveRuntimePath(params.filePath, instance)
        yield* assertExternalDirectoryEffect(ctx, file.absolute)
        yield* askRead(ctx, file, "office_word_inspect")
        yield* askMacros(ctx, file, "office_word_inspect", params.allowMacros)
        const result = yield* run(
          "word_inspect",
          {
            ...params,
            filePath: file.absolute,
            allowMacros: params.allowMacros ?? false,
          },
          ctx,
        )
        return {
          title: file.relative,
          metadata: { filepath: file.absolute },
          output: resultOutput(result),
        }
      }).pipe(Effect.orDie),
  }),
)

export const OfficeWordExportPdfTool = Tool.define(
  "office_word_export_pdf",
  Effect.gen(function* () {
    const fs = yield* AppFileSystem.Service
    const bus = yield* Bus.Service
    return {
      description:
        "Export a Word .doc or .docx document to PDF. Uses Microsoft Word on Windows and LibreOffice headless on macOS/Linux.",
      parameters: WordExportPdfParameters,
      execute: (params: WordExportPdfInput, ctx: Tool.Context) =>
        Effect.gen(function* () {
          const instance = yield* InstanceState.context
          const file = resolveRuntimePath(params.filePath, instance)
          const output = resolveRuntimePath(params.outputPath, instance)
          yield* assertExternalDirectoryEffect(ctx, file.absolute)
          yield* assertExternalDirectoryEffect(ctx, output.absolute)
          yield* askRead(ctx, file, "office_word_export_pdf")
          yield* askEdit(ctx, [output], "office_word_export_pdf", { filepath: file.absolute, outputPath: output.absolute })
          yield* askMacros(ctx, file, "office_word_export_pdf", params.allowMacros)
          const existed = yield* fs.existsSafe(output.absolute)
          const result = yield* run(
            "word_export_pdf",
            {
              ...params,
              filePath: file.absolute,
              outputPath: output.absolute,
              allowMacros: params.allowMacros ?? false,
            },
            ctx,
          )
          yield* publishUpdated(bus, output.absolute, existed)
          return {
            title: output.relative,
            metadata: { filepath: file.absolute, outputPath: output.absolute },
            output: resultOutput(result),
          }
        }).pipe(Effect.orDie),
    }
  }),
)

export const OfficeWordReplaceTool = Tool.define(
  "office_word_replace",
  Effect.gen(function* () {
    const fs = yield* AppFileSystem.Service
    const bus = yield* Bus.Service
    return {
      description:
        "Replace text in a Word document while preserving formatting where possible. Uses Microsoft Word on Windows; macOS/Linux support direct .docx OpenXML replacement only.",
      parameters: WordReplaceParameters,
      execute: (params: WordReplaceInput, ctx: Tool.Context) =>
        Effect.gen(function* () {
          if (params.replacements.length === 0) throw new Error("At least one replacement is required.")
          if (params.replacements.some((item) => item.find.length === 0)) throw new Error("Replacement find text cannot be empty.")
          const instance = yield* InstanceState.context
          const file = resolveRuntimePath(params.filePath, instance)
          const output = params.saveAs ? resolveRuntimePath(params.saveAs, instance) : file
          yield* assertExternalDirectoryEffect(ctx, file.absolute)
          yield* assertExternalDirectoryEffect(ctx, output.absolute)
          yield* askEdit(ctx, [output], "office_word_replace", {
            filepath: file.absolute,
            outputPath: output.absolute,
            replacements: params.replacements.map((item) => item.find),
          })
          yield* askMacros(ctx, file, "office_word_replace", params.allowMacros)
          const existed = yield* fs.existsSafe(output.absolute)
          const result = yield* run(
            "word_replace",
            {
              ...params,
              filePath: file.absolute,
              saveAs: params.saveAs ? output.absolute : undefined,
              allowMacros: params.allowMacros ?? false,
            },
            ctx,
          )
          yield* publishUpdated(bus, output.absolute, existed)
          return {
            title: output.relative,
            metadata: { filepath: file.absolute, outputPath: output.absolute },
            output: resultOutput(result),
          }
        }).pipe(Effect.orDie),
    }
  }),
)

export const OfficeExcelInspectTool = Tool.define(
  "office_excel_inspect",
  Effect.succeed({
    description:
      "Inspect an Excel .xls, .xlsx, or .xlsm workbook. Uses Microsoft Excel on Windows; macOS/Linux support .xlsx/.xlsm OpenXML workbook summaries.",
    parameters: ExcelInspectParameters,
    execute: (params: ExcelInspectInput, ctx: Tool.Context) =>
      Effect.gen(function* () {
        const instance = yield* InstanceState.context
        const file = resolveRuntimePath(params.filePath, instance)
        yield* assertExternalDirectoryEffect(ctx, file.absolute)
        yield* askRead(ctx, file, "office_excel_inspect")
        yield* askMacros(ctx, file, "office_excel_inspect", params.allowMacros)
        const result = yield* run(
          "excel_inspect",
          {
            ...params,
            filePath: file.absolute,
            allowMacros: params.allowMacros ?? false,
          },
          ctx,
        )
        return {
          title: file.relative,
          metadata: { filepath: file.absolute },
          output: resultOutput(result),
        }
      }).pipe(Effect.orDie),
  }),
)

export const OfficeExcelRangeTool = Tool.define(
  "office_excel_range",
  Effect.succeed({
    description:
      "Read a bounded Excel worksheet range. Uses Microsoft Excel on Windows; macOS/Linux support basic .xlsx/.xlsm cell values and formulas via OpenXML.",
    parameters: ExcelRangeParameters,
    execute: (params: ExcelRangeInput, ctx: Tool.Context) =>
      Effect.gen(function* () {
        const instance = yield* InstanceState.context
        const file = resolveRuntimePath(params.filePath, instance)
        yield* assertExternalDirectoryEffect(ctx, file.absolute)
        yield* askRead(ctx, file, "office_excel_range")
        yield* askMacros(ctx, file, "office_excel_range", params.allowMacros)
        const result = yield* run(
          "excel_range",
          {
            ...params,
            filePath: file.absolute,
            maxCells: params.maxCells ?? 1000,
            allowMacros: params.allowMacros ?? false,
          },
          ctx,
        )
        return {
          title: `${file.relative} ${params.sheet}!${params.range}`,
          metadata: { filepath: file.absolute, sheet: params.sheet, range: params.range },
          output: resultOutput(result),
        }
      }).pipe(Effect.orDie),
  }),
)

export const OfficeExcelWriteTool = Tool.define(
  "office_excel_write",
  Effect.gen(function* () {
    const fs = yield* AppFileSystem.Service
    const bus = yield* Bus.Service
    return {
      description:
        "Write values or formulas into an Excel workbook using Microsoft Excel automation while preserving workbook fidelity. Strings beginning with = are written as formulas. This operation requires Windows Excel automation.",
      parameters: ExcelWriteParameters,
      execute: (params: ExcelWriteInput, ctx: Tool.Context) =>
        Effect.gen(function* () {
          if (params.values.length === 0) throw new Error("At least one row of values is required.")
          if (params.values.every((row) => row.length === 0)) throw new Error("At least one value is required.")
          const instance = yield* InstanceState.context
          const file = resolveRuntimePath(params.filePath, instance)
          const output = params.saveAs ? resolveRuntimePath(params.saveAs, instance) : file
          yield* assertExternalDirectoryEffect(ctx, file.absolute)
          yield* assertExternalDirectoryEffect(ctx, output.absolute)
          yield* askEdit(ctx, [output], "office_excel_write", {
            filepath: file.absolute,
            outputPath: output.absolute,
            sheet: params.sheet,
            startCell: params.startCell,
            rows: params.values.length,
          })
          yield* askMacros(ctx, file, "office_excel_write", params.allowMacros)
          const existed = yield* fs.existsSafe(output.absolute)
          const result = yield* run(
            "excel_write",
            {
              ...params,
              filePath: file.absolute,
              saveAs: params.saveAs ? output.absolute : undefined,
              allowMacros: params.allowMacros ?? false,
            },
            ctx,
          )
          yield* publishUpdated(bus, output.absolute, existed)
          return {
            title: output.relative,
            metadata: { filepath: file.absolute, outputPath: output.absolute },
            output: resultOutput(result),
          }
        }).pipe(Effect.orDie),
    }
  }),
)

export const OfficeExcelRecalculateTool = Tool.define(
  "office_excel_recalculate",
  Effect.gen(function* () {
    const bus = yield* Bus.Service
    return {
      description:
        "Recalculate an Excel workbook with Microsoft Excel's native calculation engine on Windows. This operation requires Windows Excel automation.",
      parameters: ExcelRecalculateParameters,
      execute: (params: ExcelRecalculateInput, ctx: Tool.Context) =>
        Effect.gen(function* () {
          const instance = yield* InstanceState.context
          const file = resolveRuntimePath(params.filePath, instance)
          yield* assertExternalDirectoryEffect(ctx, file.absolute)
          if (params.save) {
            yield* askEdit(ctx, [file], "office_excel_recalculate", { filepath: file.absolute, save: true })
          } else {
            yield* askRead(ctx, file, "office_excel_recalculate")
          }
          yield* askMacros(ctx, file, "office_excel_recalculate", params.allowMacros)
          const result = yield* run(
            "excel_recalculate",
            {
              ...params,
              filePath: file.absolute,
              allowMacros: params.allowMacros ?? false,
            },
            ctx,
          )
          if (params.save) yield* publishUpdated(bus, file.absolute, true)
          return {
            title: file.relative,
            metadata: { filepath: file.absolute, saved: Boolean(params.save) },
            output: resultOutput(result),
          }
        }).pipe(Effect.orDie),
    }
  }),
)

export const OfficeExcelExportPdfTool = Tool.define(
  "office_excel_export_pdf",
  Effect.gen(function* () {
    const fs = yield* AppFileSystem.Service
    const bus = yield* Bus.Service
    return {
      description:
        "Export an Excel workbook to PDF. Uses Microsoft Excel on Windows and LibreOffice headless on macOS/Linux; sheet-specific export requires Windows Excel automation.",
      parameters: ExcelExportPdfParameters,
      execute: (params: ExcelExportPdfInput, ctx: Tool.Context) =>
        Effect.gen(function* () {
          const instance = yield* InstanceState.context
          const file = resolveRuntimePath(params.filePath, instance)
          const output = resolveRuntimePath(params.outputPath, instance)
          yield* assertExternalDirectoryEffect(ctx, file.absolute)
          yield* assertExternalDirectoryEffect(ctx, output.absolute)
          yield* askRead(ctx, file, "office_excel_export_pdf")
          yield* askEdit(ctx, [output], "office_excel_export_pdf", {
            filepath: file.absolute,
            outputPath: output.absolute,
            sheet: params.sheet,
          })
          yield* askMacros(ctx, file, "office_excel_export_pdf", params.allowMacros)
          const existed = yield* fs.existsSafe(output.absolute)
          const result = yield* run(
            "excel_export_pdf",
            {
              ...params,
              filePath: file.absolute,
              outputPath: output.absolute,
              allowMacros: params.allowMacros ?? false,
            },
            ctx,
          )
          yield* publishUpdated(bus, output.absolute, existed)
          return {
            title: output.relative,
            metadata: { filepath: file.absolute, outputPath: output.absolute },
            output: resultOutput(result),
          }
        }).pipe(Effect.orDie),
    }
  }),
)

export const OfficeTools: ReadonlyArray<Effect.Effect<Tool.Info<any, any>, never, any>> = [
  OfficeWordCreateTool,
  OfficeWordInspectTool,
  OfficeWordExportPdfTool,
  OfficeWordReplaceTool,
  OfficeExcelInspectTool,
  OfficeExcelRangeTool,
  OfficeExcelWriteTool,
  OfficeExcelRecalculateTool,
  OfficeExcelExportPdfTool,
]
