<#
.SYNOPSIS
    マークダウンテーブルを CSV ファイルに変換します。

.DESCRIPTION
    マークダウンファイルからテーブルを検出・パースし、CSV に変換します。
    Excel COM 不要で軽量に動作します。

.PARAMETER InputFile
    変換元のマークダウンファイルパス（必須）

.PARAMETER OutputFile
    出力先の CSV ファイルパス（省略時: 入力ファイルと同名 .csv）

.EXAMPLE
    .\Convert-MarkdownToCsv.ps1 -InputFile .\data.md
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
function Test-MarkdownTableRow {
    param([string]$Line)
    $trimmed = $Line.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) { return $false }
    if ($trimmed -notmatch '\|') { return $false }

    $cells = Split-MarkdownRow -Line $trimmed
    return $cells.Count -ge 1
}

function Split-MarkdownRow {
    param([string]$Line)

    $content = $Line.Trim()
    if ($content.StartsWith("|")) { $content = $content.Substring(1) }
    if ($content.EndsWith("|")) { $content = $content.Substring(0, $content.Length - 1) }

    $cells = [System.Collections.Generic.List[string]]::new()
    $buffer = New-Object System.Text.StringBuilder

    for ($i = 0; $i -lt $content.Length; $i++) {
        $ch = $content[$i]

        if ($ch -eq '\' -and $i + 1 -lt $content.Length) {
            $next = $content[$i + 1]
            if ($next -eq '|' -or $next -eq '\') {
                [void]$buffer.Append($next)
                $i++
                continue
            }
        }

        if ($ch -eq '|') {
            $cells.Add($buffer.ToString().Trim())
            $null = $buffer.Clear()
            continue
        }

        [void]$buffer.Append($ch)
    }

    $cells.Add($buffer.ToString().Trim())
    return $cells.ToArray()
}

function Test-MarkdownSeparatorRow {
    param([string]$Line)

    if (-not (Test-MarkdownTableRow -Line $Line)) { return $false }

    $cells = Split-MarkdownRow -Line $Line
    if ($cells.Count -eq 0) { return $false }

    foreach ($cell in $cells) {
        $t = $cell.Trim()
        if ($t.Length -eq 0 -or $t -notmatch '^:?-{1,}:?$') {
            return $false
        }
    }
    return $true
}

function Parse-MarkdownTables {
    param([string[]]$Lines)

    $tables = [System.Collections.Generic.List[object]]::new()
    $lastHeading = $null
    $index = 0

    while ($index -lt $Lines.Count) {
        $line = $Lines[$index].Trim()

        if ($line -match '^#{1,6}\s+(.+)$') {
            $lastHeading = $Matches[1].Trim()
            $index++
            continue
        }

        if ((Test-MarkdownTableRow -Line $line) -and $index + 1 -lt $Lines.Count) {
            $sepLine = $Lines[$index + 1].Trim()
            if (Test-MarkdownSeparatorRow -Line $sepLine) {
                $rows = [System.Collections.Generic.List[object]]::new()
                $rows.Add((Split-MarkdownRow -Line $line))
                $index += 2

                while ($index -lt $Lines.Count) {
                    $dataLine = $Lines[$index].Trim()
                    if (-not (Test-MarkdownTableRow -Line $dataLine)) { break }
                    $rows.Add((Split-MarkdownRow -Line $dataLine))
                    $index++
                }

                $title = if ([string]::IsNullOrWhiteSpace($lastHeading)) {
                    "Table$($tables.Count + 1)"
                }
                else {
                    $lastHeading
                }

                $tables.Add([PSCustomObject]@{
                    Title = $title
                    Rows  = $rows.ToArray()
                })

                $lastHeading = $null
                continue
            }
        }

        $index++
    }

    return $tables.ToArray()
}

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
Write-Host "  📝 Markdown → CSV 変換" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  入力: $InputFile" -ForegroundColor White
Write-Host "  出力: $OutputFile" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

Write-Host "📂 マークダウンファイルを読み込み中..." -ForegroundColor Yellow
$lines = Get-Content -Path $InputFile -Encoding UTF8

$tables = @(Parse-MarkdownTables -Lines $lines)
if ($tables.Count -eq 0) {
    Write-Warning "⚠ マークダウンファイルにテーブルが見つかりませんでした"
    exit 0
}

if ($tables.Count -gt 1) {
    Write-Host "📌 複数テーブルを検出: $($tables.Count) 個（空行区切りでCSVへ連結）" -ForegroundColor DarkYellow
}

$totalRows = 0
foreach ($table in $tables) { $totalRows += $table.Rows.Count }
$firstColCount = $tables[0].Rows[0].Count
Write-Host "📋 $firstColCount 列 × 合計 $totalRows 行を変換中..." -ForegroundColor Yellow

# CSV 出力
$csvLines = @()
for ($tableIndex = 0; $tableIndex -lt $tables.Count; $tableIndex++) {
    if ($tableIndex -gt 0) {
        $csvLines += ""
    }

    foreach ($row in $tables[$tableIndex].Rows) {
        $csvCells = @()
        foreach ($cell in $row) {
            if ($cell -match '[,"\r\n]') {
                $csvCells += '"' + ($cell -replace '"', '""') + '"'
            }
            else {
                $csvCells += $cell
            }
        }
        $csvLines += $csvCells -join ","
    }
}

# UTF-8 BOM 付きで書き出し
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllLines($OutputFile, $csvLines, $utf8Bom)

Write-Host ""
Write-Host "✅ 変換完了！" -ForegroundColor Green
Write-Host "  出力先: $OutputFile" -ForegroundColor White
Write-Host "  テーブル数: $($tables.Count)  行数: $totalRows" -ForegroundColor White
Write-Host ""
