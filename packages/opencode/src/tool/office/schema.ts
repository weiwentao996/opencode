import { Schema } from "effect"

export const OfficeFilePath = Schema.String.annotate({ description: "Path to the Word or Excel document" })
export const OfficeOutputPath = Schema.String.annotate({ description: "Path where the generated document should be written" })
export const OfficeTimeout = Schema.optional(
  Schema.Int.check(Schema.isGreaterThanOrEqualTo(1_000)).annotate({
    description: "Optional timeout in milliseconds for Microsoft Office automation",
  }),
)
export const AllowMacros = Schema.optional(
  Schema.Boolean.annotate({
    description: "Allow Office macros while opening the document. Defaults to false for safety.",
  }),
)

const ExcelRangeAddress = Schema.String.check(
  Schema.isPattern(/^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?$/),
).annotate({ description: "Excel range address, for example A1:D20" })

const ExcelCellAddress = Schema.String.check(Schema.isPattern(/^\$?[A-Za-z]{1,3}\$?\d+$/)).annotate({
  description: "Excel cell address, for example A1",
})

export const WordCreateParameters = Schema.Struct({
  outputPath: OfficeOutputPath,
  content: Schema.String.annotate({ description: "Plain text content to insert into the new Word document" }),
  title: Schema.optional(Schema.String).annotate({ description: "Optional document title" }),
  overwrite: Schema.optional(Schema.Boolean).annotate({ description: "Overwrite outputPath if it already exists" }),
  timeout: OfficeTimeout,
})
export type WordCreateInput = Schema.Schema.Type<typeof WordCreateParameters>

export const WordInspectParameters = Schema.Struct({
  filePath: OfficeFilePath,
  includeText: Schema.optional(Schema.Boolean).annotate({ description: "Include bounded document text" }),
  includeTables: Schema.optional(Schema.Boolean).annotate({ description: "Include bounded table previews" }),
  includeComments: Schema.optional(Schema.Boolean).annotate({ description: "Include comment details" }),
  includeRevisions: Schema.optional(Schema.Boolean).annotate({ description: "Include tracked revision details" }),
  maxTextChars: Schema.optional(Schema.Int.check(Schema.isGreaterThanOrEqualTo(1))).annotate({
    description: "Maximum number of characters to return for extracted text",
  }),
  allowMacros: AllowMacros,
  timeout: OfficeTimeout,
})
export type WordInspectInput = Schema.Schema.Type<typeof WordInspectParameters>

export const WordExportPdfParameters = Schema.Struct({
  filePath: OfficeFilePath,
  outputPath: OfficeOutputPath,
  overwrite: Schema.optional(Schema.Boolean).annotate({ description: "Overwrite outputPath if it already exists" }),
  allowMacros: AllowMacros,
  timeout: OfficeTimeout,
})
export type WordExportPdfInput = Schema.Schema.Type<typeof WordExportPdfParameters>

export const WordReplacement = Schema.Struct({
  find: Schema.String.annotate({ description: "Text to find" }),
  replace: Schema.String.annotate({ description: "Replacement text" }),
  matchCase: Schema.optional(Schema.Boolean).annotate({ description: "Match case during replacement" }),
  wholeWord: Schema.optional(Schema.Boolean).annotate({ description: "Match whole words only" }),
})

export const WordReplaceParameters = Schema.Struct({
  filePath: OfficeFilePath,
  replacements: Schema.Array(WordReplacement).annotate({ description: "Find/replace operations to apply" }),
  saveAs: Schema.optional(Schema.String).annotate({
    description: "Optional output path. If omitted, the original document is modified.",
  }),
  trackChanges: Schema.optional(Schema.Boolean).annotate({ description: "Enable Word tracked changes while replacing" }),
  allowMacros: AllowMacros,
  timeout: OfficeTimeout,
})
export type WordReplaceInput = Schema.Schema.Type<typeof WordReplaceParameters>

export const ExcelInspectParameters = Schema.Struct({
  filePath: OfficeFilePath,
  includeSheets: Schema.optional(Schema.Boolean).annotate({ description: "Include worksheet summaries" }),
  includeNamedRanges: Schema.optional(Schema.Boolean).annotate({ description: "Include workbook named ranges" }),
  includePivotTables: Schema.optional(Schema.Boolean).annotate({ description: "Include pivot table counts" }),
  includeCharts: Schema.optional(Schema.Boolean).annotate({ description: "Include chart counts" }),
  allowMacros: AllowMacros,
  timeout: OfficeTimeout,
})
export type ExcelInspectInput = Schema.Schema.Type<typeof ExcelInspectParameters>

export const ExcelRangeParameters = Schema.Struct({
  filePath: OfficeFilePath,
  sheet: Schema.String.annotate({ description: "Worksheet name" }),
  range: ExcelRangeAddress,
  includeFormulas: Schema.optional(Schema.Boolean).annotate({ description: "Include cell formulas" }),
  includeNumberFormats: Schema.optional(Schema.Boolean).annotate({ description: "Include cell number formats" }),
  maxCells: Schema.optional(Schema.Int.check(Schema.isGreaterThanOrEqualTo(1))).annotate({
    description: "Maximum cells to return. Defaults to 1000.",
  }),
  allowMacros: AllowMacros,
  timeout: OfficeTimeout,
})
export type ExcelRangeInput = Schema.Schema.Type<typeof ExcelRangeParameters>

const ExcelCellValue = Schema.NullOr(Schema.Union([Schema.String, Schema.Number, Schema.Boolean]))

export const ExcelWriteParameters = Schema.Struct({
  filePath: OfficeFilePath,
  sheet: Schema.String.annotate({ description: "Worksheet name" }),
  startCell: ExcelCellAddress,
  values: Schema.Array(Schema.Array(ExcelCellValue)).annotate({
    description: "Two-dimensional matrix of values to write. Strings beginning with = are written as formulas.",
  }),
  saveAs: Schema.optional(Schema.String).annotate({
    description: "Optional output path. If omitted, the original workbook is modified.",
  }),
  preserveFormats: Schema.optional(Schema.Boolean).annotate({ description: "Preserve existing cell formats where possible" }),
  allowMacros: AllowMacros,
  timeout: OfficeTimeout,
})
export type ExcelWriteInput = Schema.Schema.Type<typeof ExcelWriteParameters>

export const ExcelRecalculateParameters = Schema.Struct({
  filePath: OfficeFilePath,
  full: Schema.optional(Schema.Boolean).annotate({ description: "Run CalculateFullRebuild instead of normal calculation" }),
  save: Schema.optional(Schema.Boolean).annotate({ description: "Save the workbook after recalculation" }),
  allowMacros: AllowMacros,
  timeout: OfficeTimeout,
})
export type ExcelRecalculateInput = Schema.Schema.Type<typeof ExcelRecalculateParameters>

export const ExcelExportPdfParameters = Schema.Struct({
  filePath: OfficeFilePath,
  outputPath: OfficeOutputPath,
  sheet: Schema.optional(Schema.String).annotate({ description: "Optional worksheet name to export instead of workbook" }),
  overwrite: Schema.optional(Schema.Boolean).annotate({ description: "Overwrite outputPath if it already exists" }),
  allowMacros: AllowMacros,
  timeout: OfficeTimeout,
})
export type ExcelExportPdfInput = Schema.Schema.Type<typeof ExcelExportPdfParameters>
