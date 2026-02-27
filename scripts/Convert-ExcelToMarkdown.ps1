<#
.SYNOPSIS
    Excel ファイル (.xlsx/.xls/.xlsm) をマークダウンテーブルに変換します。

.DESCRIPTION
    Excel COM オブジェクトを利用して Excel ファイルを読み取り、
    マークダウン形式のテーブルに変換して出力します。
    M365 / Office がインストールされた Windows 環境で動作します。

.PARAMETER InputFile
    変換元の Excel ファイルパス（必須）

.PARAMETER OutputFile
    出力先のマークダウンファイルパス（省略時: 入力ファイルと同名 .md）

.PARAMETER SheetName
    変換対象のシート名（省略時: 最初のシート）

.PARAMETER AllSheets
    すべてのシートを変換する（各シートは見出し付きで出力）

.EXAMPLE
    .\Convert-ExcelToMarkdown.ps1 -InputFile .\data.xlsx
    .\Convert-ExcelToMarkdown.ps1 -InputFile .\data.xlsx -OutputFile .\output.md -SheetName "Sheet2"
    .\Convert-ExcelToMarkdown.ps1 -InputFile .\data.xlsx -AllSheets
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$InputFile,

    [Parameter(Mandatory = $false)]
    [string]$OutputFile,

    [Parameter(Mandatory = $false)]
    [string]$SheetName,

    [Parameter(Mandatory = $false)]
    [switch]$AllSheets
)

# ──────────────────────────────────────────────
# ヘルパー関数
# ──────────────────────────────────────────────
function Format-CellValue {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return "" }
    # パイプ文字をエスケープ
    $Value = $Value -replace '\|', '\|'
    # 改行をスペースに変換
    $Value = $Value -replace "`r`n", " "
    $Value = $Value -replace "`n", " "
    $Value = $Value -replace "`r", " "
    return $Value.Trim()
}

function Convert-SheetToMarkdown {
    param(
        $Sheet,
        [string]$Title = ""
    )

    $usedRange = $Sheet.UsedRange
    if ($null -eq $usedRange -or $usedRange.Rows.Count -eq 0) {
        return ""
    }

    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    $startRow = $usedRange.Row
    $startCol = $usedRange.Column

    $lines = @()

    # シートタイトル
    if ($Title -ne "") {
        $lines += "## $Title"
        $lines += ""
    }

    # 各列の最大幅を計算（見やすさのため）
    $colWidths = @()
    for ($c = 0; $c -lt $colCount; $c++) {
        $maxWidth = 3  # 最小幅
        for ($r = 0; $r -lt $rowCount; $r++) {
            $cell = $Sheet.Cells.Item($startRow + $r, $startCol + $c)
            $val = Format-CellValue -Value ([string]$cell.Text)
            if ($val.Length -gt $maxWidth) { $maxWidth = $val.Length }
        }
        $colWidths += $maxWidth
    }

    # ヘッダー行
    $headerCells = @()
    for ($c = 0; $c -lt $colCount; $c++) {
        $cell = $Sheet.Cells.Item($startRow, $startCol + $c)
        $val = Format-CellValue -Value ([string]$cell.Text)
        $headerCells += $val.PadRight($colWidths[$c])
    }
    $lines += "| " + ($headerCells -join " | ") + " |"

    # セパレータ行
    $sepCells = @()
    for ($c = 0; $c -lt $colCount; $c++) {
        $sepCells += "-" * $colWidths[$c]
    }
    $lines += "| " + ($sepCells -join " | ") + " |"

    # データ行
    for ($r = 1; $r -lt $rowCount; $r++) {
        $dataCells = @()
        for ($c = 0; $c -lt $colCount; $c++) {
            $cell = $Sheet.Cells.Item($startRow + $r, $startCol + $c)
            $val = Format-CellValue -Value ([string]$cell.Text)
            $dataCells += $val.PadRight($colWidths[$c])
        }
        $lines += "| " + ($dataCells -join " | ") + " |"
    }

    $lines += ""
    return ($lines -join "`r`n")
}

# ──────────────────────────────────────────────
# メイン処理
# ──────────────────────────────────────────────
$ErrorActionPreference = "Stop"

# 入力ファイルの検証
if (-not (Test-Path -LiteralPath $InputFile -PathType Leaf)) {
    Write-Error "❌ ファイルが見つかりません: $InputFile"
    exit 1
}
$InputFile = (Resolve-Path -LiteralPath $InputFile).Path

# 出力ファイルの決定
if ([string]::IsNullOrEmpty($OutputFile)) {
    $OutputFile = [System.IO.Path]::ChangeExtension($InputFile, ".md")
}
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)

Write-Host ""
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  📊 Excel → Markdown 変換" -ForegroundColor Cyan
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host "  入力: $InputFile" -ForegroundColor White
Write-Host "  出力: $OutputFile" -ForegroundColor White
Write-Host "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" -ForegroundColor Cyan
Write-Host ""

# Excel COM の起動
$excel = $null
$workbook = $null

try {
    Write-Host "⏳ Excel を起動中..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    Write-Host "📂 ファイルを読み込み中..." -ForegroundColor Yellow
    $workbook = $excel.Workbooks.Open($InputFile)

    $markdown = ""

    if ($AllSheets) {
        # 全シート変換
        Write-Host "📋 全 $($workbook.Sheets.Count) シートを変換中..." -ForegroundColor Yellow
        foreach ($sheet in $workbook.Sheets) {
            Write-Host "  ├─ シート: $($sheet.Name)" -ForegroundColor Gray
            $markdown += Convert-SheetToMarkdown -Sheet $sheet -Title $sheet.Name
        }
    }
    elseif ($SheetName) {
        # 指定シート変換
        $sheet = $workbook.Sheets.Item($SheetName)
        if ($null -eq $sheet) {
            Write-Error "❌ シート '$SheetName' が見つかりません"
            exit 1
        }
        Write-Host "📋 シート '$SheetName' を変換中..." -ForegroundColor Yellow
        $markdown = Convert-SheetToMarkdown -Sheet $sheet
    }
    else {
        # 最初のシート
        $sheet = $workbook.Sheets.Item(1)
        Write-Host "📋 シート '$($sheet.Name)' を変換中..." -ForegroundColor Yellow
        $markdown = Convert-SheetToMarkdown -Sheet $sheet
    }

    # ファイル出力
    $markdown | Out-File -FilePath $OutputFile -Encoding utf8
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
    # COM オブジェクトの解放
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
