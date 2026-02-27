<#
.SYNOPSIS
    マークダウンテーブルを Excel ファイル (.xlsx) に変換します。

.DESCRIPTION
    マークダウンファイルからテーブル部分を検出・パースし、
    Excel COM オブジェクトで .xlsx に書き出します。
    複数テーブルが含まれる場合、各テーブルを別シートに出力します。

.PARAMETER InputFile
    変換元のマークダウンファイルパス（必須）

.PARAMETER OutputFile
    出力先の Excel ファイルパス（省略時: 入力ファイルと同名 .xlsx）

.EXAMPLE
    .\Convert-MarkdownToExcel.ps1 -InputFile .\data.md
    .\Convert-MarkdownToExcel.ps1 -InputFile .\data.md -OutputFile .\output.xlsx
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
                    Title   = $title
                    RowData = $rows.ToArray()
                })

                $lastHeading = $null
                continue
            }
        }

        $index++
    }

    return $tables.ToArray()
}

function Get-SafeSheetName {
    param(
        [string]$RawName,
        [hashtable]$UsedSheetNames,
        [int]$Index
    )

    $baseName = if ([string]::IsNullOrWhiteSpace($RawName)) { "Sheet$Index" } else { $RawName.Trim() }
    $baseName = $baseName -replace '[\\/:?*\[\]]', '_'
    if ($baseName.Length -gt 31) {
        $baseName = $baseName.Substring(0, 31)
    }
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $baseName = "Sheet$Index"
    }

    $candidate = $baseName
    $suffix = 2
    while ($UsedSheetNames.ContainsKey($candidate)) {
        $suffixText = "_$suffix"
        $maxBaseLength = 31 - $suffixText.Length
        if ($maxBaseLength -lt 1) { $maxBaseLength = 1 }
        $trimmed = if ($baseName.Length -gt $maxBaseLength) { $baseName.Substring(0, $maxBaseLength) } else { $baseName }
        $candidate = "$trimmed$suffixText"
        $suffix++
    }

    $UsedSheetNames[$candidate] = $true
    return $candidate
}

# ──────────────────────────────────────────────
# メイン処理
# ──────────────────────────────────────────────
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
    Write-Error "ファイルが見つかりません: $InputFile"
    exit 1
}
$InputFile = (Resolve-Path -LiteralPath $InputFile).Path

if ([string]::IsNullOrEmpty($OutputFile)) {
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, ".xlsx")
}
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)

if (Test-Path $OutputFile) {
    Remove-Item $OutputFile -Force
}

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  📝 Markdown → Excel 変換" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  入力: $InputFile" -ForegroundColor White
Write-Host "  出力: $OutputFile" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

# ──────────────────────────────────────────────
# マークダウンファイルの読み込み・パース
# ──────────────────────────────────────────────
Write-Host "📂 マークダウンファイルを読み込み中..." -ForegroundColor Yellow
$lines = Get-Content -Path $InputFile -Encoding UTF8

$allTables = @(Parse-MarkdownTables -Lines $lines)

if ($allTables.Count -eq 0) {
    Write-Warning "マークダウンファイルにテーブルが見つかりませんでした"
    exit 0
}

Write-Host "📋 $($allTables.Count) 個のテーブルを検出しました" -ForegroundColor Yellow

# ──────────────────────────────────────────────
# Excel COM でファイル出力
# ──────────────────────────────────────────────
$excel = $null
$workbook = $null

try {
    Write-Host "⏳ Excel を起動中..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()

    while ($workbook.Sheets.Count -gt 1) {
        $workbook.Sheets.Item($workbook.Sheets.Count).Delete()
    }

    $sheetIndex = 0
    $usedSheetNames = @{}
    for ($t = 0; $t -lt $allTables.Count; $t++) {
        $tbl = $allTables[$t]
        $rows = $tbl.RowData
        $sheetTitle = $tbl.Title

        if ($null -eq $rows -or $rows.Length -eq 0) {
            Write-Host "  ├─ テーブル $($t+1) はデータなしのためスキップ" -ForegroundColor DarkYellow
            continue
        }

        $sheetTitle = Get-SafeSheetName -RawName $sheetTitle -UsedSheetNames $usedSheetNames -Index ($sheetIndex + 1)

        if ($sheetIndex -eq 0) {
            $sheet = $workbook.Sheets.Item(1)
            $sheet.Name = $sheetTitle
        }
        else {
            $sheet = $workbook.Sheets.Add([System.Reflection.Missing]::Value, $workbook.Sheets.Item($workbook.Sheets.Count))
            $sheet.Name = $sheetTitle
        }

        Write-Host "  ├─ シート '$sheetTitle' に書き込み中... ($($rows.Length) 行)" -ForegroundColor Gray

        for ($r = 0; $r -lt $rows.Length; $r++) {
            $row = $rows[$r]
            for ($c = 0; $c -lt $row.Length; $c++) {
                $sheet.Cells.Item($r + 1, $c + 1).Value2 = [string]$row[$c]
            }
        }

        # ヘッダー行装飾
        $colCount = $rows[0].Length
        if ($colCount -gt 0) {
            $headerRange = $sheet.Range($sheet.Cells.Item(1, 1), $sheet.Cells.Item(1, $colCount))
            $headerRange.Font.Bold = $true
            $headerRange.Interior.Color = 0xD9E1F2
            $headerRange.Font.Color = 0x000000
        }

        $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null

        if ($rows.Length -gt 1 -and $colCount -gt 0) {
            $tableRange = $sheet.UsedRange
            foreach ($border in @(7, 8, 9, 10, 11, 12)) {
                $tableRange.Borders.Item($border).LineStyle = 1
                $tableRange.Borders.Item($border).Weight = 2
                $tableRange.Borders.Item($border).Color = 0xC0C0C0
            }
        }

        $sheetIndex++
    }

    if ($sheetIndex -eq 0) {
        Write-Warning "有効なテーブルが見つかりませんでした"
    }
    else {
        $workbook.Sheets.Item(1).Activate()
        $workbook.SaveAs($OutputFile, 51)

        Write-Host ""
        Write-Host "✅ 変換完了！" -ForegroundColor Green
        Write-Host "  出力先: $OutputFile" -ForegroundColor White
        Write-Host "  テーブル数: $sheetIndex" -ForegroundColor White
        Write-Host ""
    }
}
catch {
    Write-Error "エラーが発生しました: $_"
    exit 1
}
finally {
    if ($workbook) {
        try { $workbook.Close($false) } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null } catch {}
    }
    if ($excel) {
        try { $excel.Quit() } catch {}
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
