import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
import os from "node:os"
import path from "node:path"
import { Process } from "@/util/process"

export type OfficeOperation =
  | "word_create"
  | "word_inspect"
  | "word_export_pdf"
  | "word_replace"
  | "excel_inspect"
  | "excel_range"
  | "excel_write"
  | "excel_recalculate"
  | "excel_export_pdf"

export type OfficeRequest = {
  operation: OfficeOperation
  timeout?: number
  [key: string]: unknown
}

type OfficeResponse =
  | { ok: true; result: unknown; warnings?: string[] }
  | { ok: false; error: string; details?: string; hresult?: string }

export class OfficeAutomationError extends Error {
  constructor(
    message: string,
    readonly details?: string,
  ) {
    super(details ? `${message}\n${details}` : message)
    this.name = "OfficeAutomationError"
  }
}

export type OfficeRunner = (cmd: string[], options: Parameters<typeof Process.text>[1]) => Promise<{ text: string; code: number }>

async function ensureOfficeScript() {
  const dir = path.join(os.tmpdir(), "opencode-office")
  const scriptPath = path.join(dir, "office.ps1")
  await mkdir(dir, { recursive: true })
  const current = await readFile(scriptPath, "utf8").catch(() => undefined)
  if (current !== POWERSHELL_SCRIPT) await writeFile(scriptPath, POWERSHELL_SCRIPT, "utf8")
  return scriptPath
}

export async function runOfficeAutomation(request: OfficeRequest, options: { abort?: AbortSignal; runner?: OfficeRunner } = {}) {
  if (process.platform !== "win32") {
    throw new OfficeAutomationError("Microsoft Office automation is only available on Windows.")
  }

  const dir = await mkdtemp(path.join(os.tmpdir(), "opencode-office-"))
  const requestPath = path.join(dir, "request.json")
  const responsePath = path.join(dir, "response.json")
  const scriptPath = await ensureOfficeScript()
  const runner = options.runner ?? ((cmd, opts) => Process.text(cmd, { ...opts, nothrow: true }))

  try {
    await writeFile(requestPath, JSON.stringify(request, null, 2), "utf8")
    const out = await runner(
      ["powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", scriptPath],
      {
        env: {
          OPENCODE_OFFICE_REQUEST: requestPath,
          OPENCODE_OFFICE_RESPONSE: responsePath,
        },
        abort: options.abort,
        timeout: request.timeout ?? 120_000,
      },
    )

    let raw = ""
    try {
      raw = await readFile(responsePath, "utf8")
    } catch {
      const text = out.text.trim()
      throw new OfficeAutomationError("Microsoft Office automation failed before producing a response.", text)
    }

    const response = JSON.parse(raw.replace(/^﻿/, "")) as OfficeResponse
    if (!response.ok) throw new OfficeAutomationError(response.error, response.details ?? response.hresult)
    return response.result
  } finally {
    await rm(dir, { recursive: true, force: true })
  }
}

export const POWERSHELL_SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Write-OfficeResponse($Payload) {
  $json = $Payload | ConvertTo-Json -Depth 64 -Compress
  $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
  [System.IO.File]::WriteAllText($env:OPENCODE_OFFICE_RESPONSE, $json, $utf8NoBom)
}

function Release-Com($Object) {
  if ($null -eq $Object) { return }
  try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch {}
}

function Finish-ComCleanup() {
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

function Disable-Macros($App, $AllowMacros) {
  if ($AllowMacros) { return }
  try { $App.AutomationSecurity = 3 } catch {}
}

function Open-WordDocument($Word, $Path, $ReadOnly) {
  return $Word.Documents.Open($Path, $false, $ReadOnly, $false)
}

function Get-WordFileFormat($Path) {
  $extension = [System.IO.Path]::GetExtension([string]$Path).ToLowerInvariant()
  switch ($extension) {
    '.doc' { return 0 }
    '.docm' { return 13 }
    default { return 16 }
  }
}

function Save-WordDocumentAs($Document, $Path, $Overwrite) {
  if ((Test-Path -LiteralPath $Path) -and -not $Overwrite) { throw "Output file already exists: $Path" }
  $directory = [System.IO.Path]::GetDirectoryName([string]$Path)
  if (-not [string]::IsNullOrEmpty($directory)) { [System.IO.Directory]::CreateDirectory($directory) | Out-Null }
  $format = Get-WordFileFormat $Path
  $hadExisting = Test-Path -LiteralPath $Path
  if ($hadExisting -and $Overwrite) { Remove-Item -LiteralPath $Path -Force }
  try {
    $Document.SaveAs2([ref]$Path, [ref]$format)
  } catch {
    if (-not $hadExisting -and (Test-Path -LiteralPath $Path)) {
      Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
    throw
  }
}

function Open-ExcelWorkbook($Excel, $Path, $ReadOnly) {
  return $Excel.Workbooks.Open($Path, $null, $ReadOnly)
}

function Convert-ToWordText($Text) {
  $value = [string]$Text
  $value = [System.Text.RegularExpressions.Regex]::Replace($value, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
  return [System.Text.RegularExpressions.Regex]::Replace($value, '\r\n|\n|\r', [Environment]::NewLine)
}

function Invoke-WordCreate($Request) {
  $word = $null; $doc = $null
  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $selection = $word.Selection
    if ($Request.title) {
      $selection.Font.Bold = $true
      $selection.Font.Size = 16
      $selection.TypeText((Convert-ToWordText $Request.title))
      $selection.TypeParagraph()
      $selection.TypeParagraph()
      $selection.Font.Bold = $false
      $selection.Font.Size = 11
    }
    $selection.TypeText((Convert-ToWordText $Request.content))
    Save-WordDocumentAs $doc $Request.outputPath ([bool]$Request.overwrite)
    return @{ outputPath = $Request.outputPath; format = [System.IO.Path]::GetExtension([string]$Request.outputPath); title = $Request.title }
  } finally {
    if ($doc) { try { $doc.Close($false) } catch {}; Release-Com $doc }
    if ($word) { try { $word.Quit() } catch {}; Release-Com $word }
    Finish-ComCleanup
  }
}

function Get-WordHeadings($Document) {
  $items = New-Object System.Collections.ArrayList
  $limit = [Math]::Min($Document.Paragraphs.Count, 300)
  for ($i = 1; $i -le $limit; $i++) {
    $paragraph = $Document.Paragraphs.Item($i)
    $style = ''
    try { $style = [string]$paragraph.Range.Style.NameLocal } catch {}
    if ($style -match 'Heading|标题|標題') {
      $text = ([string]$paragraph.Range.Text).Trim() -replace "\r|\a", ''
      if ($text.Length -gt 0) { [void]$items.Add(@{ index = $i; style = $style; text = $text }) }
    }
    Release-Com $paragraph
  }
  return $items
}

function Get-WordTables($Document) {
  $items = New-Object System.Collections.ArrayList
  $limit = [Math]::Min($Document.Tables.Count, 20)
  for ($i = 1; $i -le $limit; $i++) {
    $table = $Document.Tables.Item($i)
    [void]$items.Add(@{ index = $i; rows = $table.Rows.Count; columns = $table.Columns.Count })
    Release-Com $table
  }
  return $items
}

function Invoke-WordInspect($Request) {
  $word = $null; $doc = $null
  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    Disable-Macros $word $Request.allowMacros
    $doc = Open-WordDocument $word $Request.filePath $true
    $result = [ordered]@{
      filePath = $Request.filePath
      name = $doc.Name
      fullName = $doc.FullName
      pages = $doc.ComputeStatistics(2)
      words = $doc.ComputeStatistics(0)
      characters = $doc.ComputeStatistics(3)
      paragraphs = $doc.Paragraphs.Count
      tables = $doc.Tables.Count
      sections = $doc.Sections.Count
      comments = $doc.Comments.Count
      revisions = $doc.Revisions.Count
      headings = Get-WordHeadings $doc
    }
    if ($Request.includeText) {
      $max = 12000
      if ($Request.maxTextChars) { $max = [int]$Request.maxTextChars }
      $text = [string]$doc.Content.Text
      if ($text.Length -gt $max) { $text = $text.Substring(0, $max) }
      $result.text = $text
    }
    if ($Request.includeTables) { $result.tablePreview = Get-WordTables $doc }
    if ($Request.includeComments) {
      $comments = New-Object System.Collections.ArrayList
      for ($i = 1; $i -le [Math]::Min($doc.Comments.Count, 50); $i++) {
        $comment = $doc.Comments.Item($i)
        [void]$comments.Add(@{ index = $i; author = [string]$comment.Author; text = ([string]$comment.Range.Text).Trim() })
        Release-Com $comment
      }
      $result.commentPreview = $comments
    }
    if ($Request.includeRevisions) {
      $revisions = New-Object System.Collections.ArrayList
      for ($i = 1; $i -le [Math]::Min($doc.Revisions.Count, 50); $i++) {
        $revision = $doc.Revisions.Item($i)
        [void]$revisions.Add(@{ index = $i; author = [string]$revision.Author; type = [int]$revision.Type; text = ([string]$revision.Range.Text).Trim() })
        Release-Com $revision
      }
      $result.revisionPreview = $revisions
    }
    return $result
  } finally {
    if ($doc) { try { $doc.Close($false) } catch {}; Release-Com $doc }
    if ($word) { try { $word.Quit() } catch {}; Release-Com $word }
    Finish-ComCleanup
  }
}

function Invoke-WordExportPdf($Request) {
  if ((Test-Path -LiteralPath $Request.outputPath) -and -not $Request.overwrite) { throw "Output file already exists: $($Request.outputPath)" }
  $word = $null; $doc = $null
  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    Disable-Macros $word $Request.allowMacros
    $doc = Open-WordDocument $word $Request.filePath $true
    $doc.ExportAsFixedFormat($Request.outputPath, 17)
    return @{ filePath = $Request.filePath; outputPath = $Request.outputPath; format = 'pdf' }
  } finally {
    if ($doc) { try { $doc.Close($false) } catch {}; Release-Com $doc }
    if ($word) { try { $word.Quit() } catch {}; Release-Com $word }
    Finish-ComCleanup
  }
}

function Invoke-WordReplace($Request) {
  if (-not $Request.replacements -or $Request.replacements.Count -eq 0) { throw 'At least one replacement is required.' }
  $word = $null; $doc = $null
  try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    Disable-Macros $word $Request.allowMacros
    $doc = Open-WordDocument $word $Request.filePath $false
    if ($null -ne $Request.trackChanges) { $doc.TrackRevisions = [bool]$Request.trackChanges }
    $results = New-Object System.Collections.ArrayList
    foreach ($item in $Request.replacements) {
      if ([string]::IsNullOrEmpty([string]$item.find)) { throw 'Replacement find text cannot be empty.' }
      $range = $doc.Content
      $find = $range.Find
      $find.ClearFormatting()
      $find.Replacement.ClearFormatting()
      $count = 0
      while ($find.Execute([string]$item.find, [bool]$item.matchCase, [bool]$item.wholeWord, $false, $false, $false, $true, 1, $false, [string]$item.replace, 1)) {
        $count++
      }
      [void]$results.Add(@{ find = [string]$item.find; replace = [string]$item.replace; count = $count })
      Release-Com $find
      Release-Com $range
    }
    $finalPath = $Request.filePath
    if ($Request.saveAs) {
      $finalPath = [string]$Request.saveAs
      Save-WordDocumentAs $doc $finalPath $true
    } else {
      $doc.Save()
    }
    return @{ filePath = $Request.filePath; outputPath = $finalPath; replacements = $results }
  } finally {
    if ($doc) { try { $doc.Close($false) } catch {}; Release-Com $doc }
    if ($word) { try { $word.Quit() } catch {}; Release-Com $word }
    Finish-ComCleanup
  }
}

function Get-ExcelSheetSummary($Sheet, $IncludePivotTables, $IncludeCharts) {
  $used = $Sheet.UsedRange
  $formulaCount = 0
  try { $formulaCount = $used.SpecialCells(-4123).Count } catch {}
  $errorCount = 0
  $maxScan = [Math]::Min($used.Rows.Count * $used.Columns.Count, 5000)
  $scanned = 0
  foreach ($cell in $used.Cells) {
    if ($scanned -ge $maxScan) { break }
    $scanned++
    try { if (([string]$cell.Text).StartsWith('#')) { $errorCount++ } } catch {}
    Release-Com $cell
  }
  $chartCount = 0
  if ($IncludeCharts) { try { $chartCount = $Sheet.ChartObjects().Count } catch {} }
  $pivotCount = 0
  if ($IncludePivotTables) { try { $pivotCount = $Sheet.PivotTables().Count } catch {} }
  $summary = [ordered]@{
    name = $Sheet.Name
    index = $Sheet.Index
    visible = $Sheet.Visible
    usedRange = $used.Address($false, $false)
    rows = $used.Rows.Count
    columns = $used.Columns.Count
    formulaCount = $formulaCount
    scannedErrorCount = $errorCount
  }
  if ($IncludeCharts) { $summary.chartCount = $chartCount }
  if ($IncludePivotTables) { $summary.pivotTableCount = $pivotCount }
  Release-Com $used
  return $summary
}

function Invoke-ExcelInspect($Request) {
  $excel = $null; $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Disable-Macros $excel $Request.allowMacros
    $wb = Open-ExcelWorkbook $excel $Request.filePath $true
    $result = [ordered]@{
      filePath = $Request.filePath
      name = $wb.Name
      fullName = $wb.FullName
      sheets = $wb.Worksheets.Count
      hasMacros = ([System.IO.Path]::GetExtension($Request.filePath).ToLowerInvariant() -in @('.xlsm', '.xls'))
      calculationMode = $excel.Calculation
    }
    if ($Request.includeSheets -ne $false) {
      $sheets = New-Object System.Collections.ArrayList
      foreach ($sheet in $wb.Worksheets) {
        [void]$sheets.Add((Get-ExcelSheetSummary $sheet ([bool]$Request.includePivotTables) ([bool]$Request.includeCharts)))
        Release-Com $sheet
      }
      $result.sheetSummary = $sheets
    }
    if ($Request.includeNamedRanges) {
      $names = New-Object System.Collections.ArrayList
      foreach ($name in $wb.Names) {
        [void]$names.Add(@{ name = [string]$name.Name; refersTo = [string]$name.RefersTo })
        Release-Com $name
      }
      $result.namedRanges = $names
    }
    return $result
  } finally {
    if ($wb) { try { $wb.Close($false) } catch {}; Release-Com $wb }
    if ($excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
    Finish-ComCleanup
  }
}

function Get-Worksheet($Workbook, $Name) {
  foreach ($sheet in $Workbook.Worksheets) {
    if ($sheet.Name -eq $Name) { return $sheet }
    Release-Com $sheet
  }
  throw "Worksheet not found: $Name"
}

function Invoke-ExcelRange($Request) {
  $excel = $null; $wb = $null; $sheet = $null; $range = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Disable-Macros $excel $Request.allowMacros
    $wb = Open-ExcelWorkbook $excel $Request.filePath $true
    $sheet = Get-Worksheet $wb $Request.sheet
    $range = $sheet.Range($Request.range)
    $cells = $range.Rows.Count * $range.Columns.Count
    $max = 1000
    if ($Request.maxCells) { $max = [int]$Request.maxCells }
    if ($cells -gt $max) { throw "Requested range has $cells cells, which exceeds maxCells=$max." }
    $rows = New-Object System.Collections.ArrayList
    for ($r = 1; $r -le $range.Rows.Count; $r++) {
      $row = New-Object System.Collections.ArrayList
      for ($c = 1; $c -le $range.Columns.Count; $c++) {
        $cell = $range.Cells.Item($r, $c)
        $item = [ordered]@{ address = $cell.Address($false, $false); text = [string]$cell.Text; value = $cell.Value2 }
        if ($Request.includeFormulas) { $item.formula = [string]$cell.Formula }
        if ($Request.includeNumberFormats) { $item.numberFormat = [string]$cell.NumberFormat }
        if ($cell.MergeCells) { $item.mergeArea = [string]$cell.MergeArea.Address($false, $false) }
        [void]$row.Add($item)
        Release-Com $cell
      }
      [void]$rows.Add($row)
    }
    return @{ filePath = $Request.filePath; sheet = $Request.sheet; range = $range.Address($false, $false); rows = $rows }
  } finally {
    if ($range) { Release-Com $range }
    if ($sheet) { Release-Com $sheet }
    if ($wb) { try { $wb.Close($false) } catch {}; Release-Com $wb }
    if ($excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
    Finish-ComCleanup
  }
}

function Invoke-ExcelWrite($Request) {
  if (-not $Request.values -or $Request.values.Count -eq 0) { throw 'At least one row of values is required.' }
  $excel = $null; $wb = $null; $sheet = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Disable-Macros $excel $Request.allowMacros
    $wb = Open-ExcelWorkbook $excel $Request.filePath $false
    $sheet = Get-Worksheet $wb $Request.sheet
    $rowCount = $Request.values.Count
    $columnCount = 0
    foreach ($row in $Request.values) { if ($row.Count -gt $columnCount) { $columnCount = $row.Count } }
    if ($columnCount -eq 0) { throw 'At least one value is required.' }
    for ($r = 0; $r -lt $rowCount; $r++) {
      $row = $Request.values[$r]
      for ($c = 0; $c -lt $row.Count; $c++) {
        $cell = $sheet.Range($Request.startCell).Offset($r, $c)
        $value = $row[$c]
        if ($value -is [string] -and $value.StartsWith('=')) { $cell.Formula = $value } else { $cell.Value2 = $value }
        Release-Com $cell
      }
    }
    $finalPath = $Request.filePath
    if ($Request.saveAs) { $finalPath = [string]$Request.saveAs; $wb.SaveAs($finalPath) } else { $wb.Save() }
    return @{ filePath = $Request.filePath; outputPath = $finalPath; sheet = $Request.sheet; startCell = $Request.startCell; rows = $rowCount; columns = $columnCount }
  } finally {
    if ($sheet) { Release-Com $sheet }
    if ($wb) { try { $wb.Close($false) } catch {}; Release-Com $wb }
    if ($excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
    Finish-ComCleanup
  }
}

function Invoke-ExcelRecalculate($Request) {
  $excel = $null; $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Disable-Macros $excel $Request.allowMacros
    $wb = Open-ExcelWorkbook $excel $Request.filePath $false
    if ($Request.full) { $excel.CalculateFullRebuild() } else { $excel.Calculate() }
    try { $excel.CalculateUntilAsyncQueriesDone() } catch {}
    if ($Request.save) { $wb.Save() }
    return @{ filePath = $Request.filePath; full = [bool]$Request.full; saved = [bool]$Request.save; calculationState = $excel.CalculationState }
  } finally {
    if ($wb) { try { $wb.Close($false) } catch {}; Release-Com $wb }
    if ($excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
    Finish-ComCleanup
  }
}

function Invoke-ExcelExportPdf($Request) {
  if ((Test-Path -LiteralPath $Request.outputPath) -and -not $Request.overwrite) { throw "Output file already exists: $($Request.outputPath)" }
  $excel = $null; $wb = $null; $sheet = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Disable-Macros $excel $Request.allowMacros
    $wb = Open-ExcelWorkbook $excel $Request.filePath $true
    if ($Request.sheet) {
      $sheet = Get-Worksheet $wb $Request.sheet
      $sheet.ExportAsFixedFormat(0, $Request.outputPath)
    } else {
      $wb.ExportAsFixedFormat(0, $Request.outputPath)
    }
    return @{ filePath = $Request.filePath; outputPath = $Request.outputPath; sheet = $Request.sheet; format = 'pdf' }
  } finally {
    if ($sheet) { Release-Com $sheet }
    if ($wb) { try { $wb.Close($false) } catch {}; Release-Com $wb }
    if ($excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
    Finish-ComCleanup
  }
}

try {
  $requestText = [System.IO.File]::ReadAllText($env:OPENCODE_OFFICE_REQUEST, [System.Text.Encoding]::UTF8)
  $request = $requestText | ConvertFrom-Json
  $result = switch ($request.operation) {
    'word_create' { Invoke-WordCreate $request }
    'word_inspect' { Invoke-WordInspect $request }
    'word_export_pdf' { Invoke-WordExportPdf $request }
    'word_replace' { Invoke-WordReplace $request }
    'excel_inspect' { Invoke-ExcelInspect $request }
    'excel_range' { Invoke-ExcelRange $request }
    'excel_write' { Invoke-ExcelWrite $request }
    'excel_recalculate' { Invoke-ExcelRecalculate $request }
    'excel_export_pdf' { Invoke-ExcelExportPdf $request }
    default { throw "Unknown Office operation: $($request.operation)" }
  }
  Write-OfficeResponse @{ ok = $true; result = $result }
} catch {
  $exception = $_.Exception
  Write-OfficeResponse @{ ok = $false; error = $exception.Message; details = [string]$_; hresult = ('0x{0:X8}' -f $exception.HResult) }
  exit 1
}
`
