<#
.SYNOPSIS
    Excel ファイル (.xlsx/.xls/.xlsm) を CSV に変換します。

.DESCRIPTION
    Excel COM オブジェクトでファイルを読み取り、UTF-8 BOM 付き CSV に変換します。

.PARAMETER InputFile
    変換元の Excel ファイルパス（必須）

.PARAMETER OutputFile
    出力先の CSV ファイルパス（省略時: 入力ファイルと同名 .csv）

.PARAMETER SheetName
    変換対象のシート名（省略時: 最初のシート）

.EXAMPLE
    .\Convert-ExcelToCsv.ps1 -InputFile .\data.xlsx
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [Parameter(Mandatory = $false)]
    [string]$SheetName
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
    Write-Error "❌ ファイルが見つかりません: $InputFile"
    exit 1
}
$InputFile = (Resolve-Path -LiteralPath $InputFile).Path

if ([string]::IsNullOrEmpty($OutputFile)) {
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, ".csv")
}
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  📊 Excel → CSV 変換" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  入力: $InputFile" -ForegroundColor White
Write-Host "  出力: $OutputFile" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

$excel = $null
$workbook = $null

try {
    Write-Host "⏳ Excel を起動中..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "📂 ファイルを読み込み中..." -ForegroundColor Yellow
    $workbook = $excel.Workbooks.Open($InputFile)

    if ($SheetName) {
        $sheet = $workbook.Sheets.Item($SheetName)
        if ($null -eq $sheet) {
            Write-Error "❌ シート '$SheetName' が見つかりません"
            exit 1
        }
    }
    else {
        $sheet = $workbook.Sheets.Item(1)
    }

    Write-Host "📋 シート '$($sheet.Name)' を変換中..." -ForegroundColor Yellow

    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    $startRow = $usedRange.Row
    $startCol = $usedRange.Column

    $csvLines = @()
    for ($r = 0; $r -lt $rowCount; $r++) {
        $cells = @()
        for ($c = 0; $c -lt $colCount; $c++) {
            $val = [string]$sheet.Cells.Item($startRow + $r, $startCol + $c).Text
            # カンマ・改行・ダブルクォートを含む場合はクォートで囲む
            if ($val -match '[,"\r\n]') {
                $val = '"' + ($val -replace '"', '""') + '"'
            }
            $cells += $val
        }
        $csvLines += $cells -join ","
    }

    # UTF-8 BOM 付きで書き出し
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllLines($OutputFile, $csvLines, $utf8Bom)

    Write-Host ""
    Write-Host "✅ 変換完了！" -ForegroundColor Green
    Write-Host "  出力先: $OutputFile" -ForegroundColor White
    Write-Host "  行数: $rowCount  列数: $colCount" -ForegroundColor White
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
