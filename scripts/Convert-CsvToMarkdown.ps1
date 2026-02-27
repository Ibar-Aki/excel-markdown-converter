<#
.SYNOPSIS
    CSV ファイルをマークダウンテーブルに変換します。

.DESCRIPTION
    CSV を Import-Csv で読み取り、マークダウン形式のテーブルに変換します。
    Excel COM 不要で軽量に動作します。

.PARAMETER InputFile
    変換元の CSV ファイルパス（必須）

.PARAMETER OutputFile
    出力先のマークダウンファイルパス（省略時: 入力ファイルと同名 .md）

.EXAMPLE
    .\Convert-CsvToMarkdown.ps1 -InputFile .\data.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile
)

# ──────────────────────────────────────────────
# ヘルパー関数
# ──────────────────────────────────────────────
function Format-CellValue {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return "" }
    $Value = $Value -replace '\|', '\|'
    $Value = $Value -replace "`r`n", " "
    $Value = $Value -replace "`n", " "
    $Value = $Value -replace "`r", " "
    return $Value.Trim()
}

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

            # 空行は無視（末尾空行など）
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

# ──────────────────────────────────────────────
# メイン処理
# ──────────────────────────────────────────────
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
    Write-Error "❌ ファイルが見つかりません: $InputFile"
    exit 1
}
$InputFile = (Resolve-Path -LiteralPath $InputFile).Path

if ([string]::IsNullOrEmpty($OutputFile)) {
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, ".md")
}
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  📄 CSV → Markdown 変換" -ForegroundColor Cyan
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

# 列幅の計算
$colWidths = @()
foreach ($h in $headers) {
    $maxWidth = (Format-CellValue -Value $h).Length
    if ($maxWidth -lt 3) { $maxWidth = 3 }
    foreach ($row in $csvRows) {
        $val = (Format-CellValue -Value ([string]$row.$h)).Length
        if ($val -gt $maxWidth) { $maxWidth = $val }
    }
    $colWidths += $maxWidth
}

$lines = @()

# ヘッダー行
$headerCells = @()
for ($i = 0; $i -lt $headers.Count; $i++) {
    $headerCells += (Format-CellValue -Value $headers[$i]).PadRight($colWidths[$i])
}
$lines += "| " + ($headerCells -join " | ") + " |"

# セパレータ行
$sepCells = @()
for ($i = 0; $i -lt $headers.Count; $i++) {
    $sepCells += "-" * $colWidths[$i]
}
$lines += "| " + ($sepCells -join " | ") + " |"

# データ行
foreach ($row in $csvRows) {
    $dataCells = @()
    for ($i = 0; $i -lt $headers.Count; $i++) {
        $val = Format-CellValue -Value ([string]$row.($headers[$i]))
        $dataCells += $val.PadRight($colWidths[$i])
    }
    $lines += "| " + ($dataCells -join " | ") + " |"
}

$markdown = $lines -join "`r`n"
$markdown | Out-File -FilePath $OutputFile -Encoding utf8

Write-Host ""
Write-Host "✅ 変換完了！" -ForegroundColor Green
Write-Host "  出力先: $OutputFile" -ForegroundColor White
Write-Host ""
