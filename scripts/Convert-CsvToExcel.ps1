<#
.SYNOPSIS
    CSV ファイルを Excel ファイル (.xlsx) に変換します。

.DESCRIPTION
    CSV を読み取り、Excel COM オブジェクトで .xlsx に変換します。
    ヘッダー行の書式設定、列幅自動調整、罫線付きで見栄えの良い Excel を生成します。

.PARAMETER InputFile
    変換元の CSV ファイルパス（必須）

.PARAMETER OutputFile
    出力先の Excel ファイルパス（省略時: 入力ファイルと同名 .xlsx）

.EXAMPLE
    .\Convert-CsvToExcel.ps1 -InputFile .\data.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile
)

function Test-CsvColumnConsistency {
    param([string]$Path)

    Add-Type -AssemblyName Microsoft.VisualBasic
    $parser = $null
    try {
        $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [System.Text.Encoding]::UTF8, $true)
        $parser.SetDelimiters(",")
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false

        $expectedColumns = $null
        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields()
            if ($null -eq $fields) { continue }

            if ($fields.Count -eq 1 -and [string]::IsNullOrWhiteSpace($fields[0])) {
                continue
            }

            if ($null -eq $expectedColumns) {
                $expectedColumns = $fields.Count
                continue
            }

            if ($fields.Count -ne $expectedColumns) {
                return [PSCustomObject]@{
                    IsValid         = $false
                    ExpectedColumns = $expectedColumns
                    ActualColumns   = $fields.Count
                    LineNumber      = $parser.LineNumber
                }
            }
        }

        return [PSCustomObject]@{
            IsValid         = $true
            ExpectedColumns = $expectedColumns
            ActualColumns   = $expectedColumns
            LineNumber      = 0
        }
    }
    finally {
        if ($parser) {
            $parser.Close()
        }
    }
}

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
    Write-Error "❌ ファイルが見つかりません: $InputFile"
    exit 1
}
$InputFile = (Resolve-Path -LiteralPath $InputFile).Path

if ([string]::IsNullOrEmpty($OutputFile)) {
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, ".xlsx")
}
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)

# 既存ファイルがあれば削除
if (Test-Path $OutputFile) {
    Remove-Item $OutputFile -Force
}

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  📄 CSV → Excel 変換" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  入力: $InputFile" -ForegroundColor White
Write-Host "  出力: $OutputFile" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

Write-Host "📂 CSV を読み込み中..." -ForegroundColor Yellow
$consistency = Test-CsvColumnConsistency -Path $InputFile
if (-not $consistency.IsValid) {
    Write-Error "❌ CSV の列数が不一致です (行 $($consistency.LineNumber)): 想定 $($consistency.ExpectedColumns) 列 / 実際 $($consistency.ActualColumns) 列"
    exit 1
}

$csvData = Import-Csv -Path $InputFile -Encoding UTF8
$csvRows = @($csvData)
$rowCount = $csvRows.Count

if ($rowCount -eq 0) {
    Write-Warning "⚠ CSV にデータがありません"
    exit 0
}

$headers = $csvRows[0].PSObject.Properties.Name
Write-Host "📋 $($headers.Count) 列 × $rowCount 行を変換中..." -ForegroundColor Yellow

$excel = $null
$workbook = $null

try {
    Write-Host "⏳ Excel を起動中..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $sheet = $workbook.Sheets.Item(1)

    # ヘッダー行の書き込み
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $sheet.Cells.Item(1, $c + 1).Value2 = $headers[$c]
    }

    # データ行の書き込み
    $rowIndex = 2
    foreach ($row in $csvRows) {
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $sheet.Cells.Item($rowIndex, $c + 1).Value2 = [string]$row.($headers[$c])
        }
        $rowIndex++
    }

    # ヘッダーの書式設定
    $headerRange = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(1, $headers.Count))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 0xD9E1F2  # 薄い青
    $headerRange.Font.Color = 0x000000

    # 列幅の自動調整
    $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null

    # 罫線
    $tableRange = $sheet.UsedRange
    $borders = @(7, 8, 9, 10, 11, 12)
    foreach ($border in $borders) {
        $tableRange.Borders.Item($border).LineStyle = 1
        $tableRange.Borders.Item($border).Weight = 2
        $tableRange.Borders.Item($border).Color = 0xC0C0C0
    }

    # 保存 (xlOpenXMLWorkbook = 51)
    $workbook.SaveAs($OutputFile, 51)

    Write-Host ""
    Write-Host "✅ 変換完了！" -ForegroundColor Green
    Write-Host "  出力先: $OutputFile" -ForegroundColor White
    Write-Host ""
}
catch {
    Write-Error "❌ エラーが発生しました: $_"
    exit 1
}
finally {
    if ($workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
