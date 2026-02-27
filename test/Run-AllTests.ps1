<#
.SYNOPSIS
    全変換パターンの一括テストを実行します。

.DESCRIPTION
    1. テスト用 Excel の生成
    2. Excel → Markdown（直接）
    3. Markdown → Excel（直接）
    4. Excel → CSV → Markdown（CSV 経由）
    5. Markdown → CSV → Excel（CSV 経由）
    の全パイプラインを自動テストします。
#>

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$scriptDir = Join-Path $projectRoot "scripts"
$testDir = Join-Path $projectRoot "test"
$outputDir = Join-Path $testDir "output"

# 出力フォルダの作成
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

Write-Host ""
Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor Magenta
Write-Host "║     🧪 Excel ↔ Markdown 変換ツール テスト       ║" -ForegroundColor Magenta
Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor Magenta
Write-Host ""

$passed = 0
$failed = 0
$total = 0

function Test-Step {
    param(
        [string]$Name,
        [scriptblock]$Action
    )
    $script:total++
    Write-Host "┌─ テスト $($script:total): $Name" -ForegroundColor Yellow
    try {
        & $Action
        Write-Host "└─ ✅ PASS" -ForegroundColor Green
        Write-Host ""
        $script:passed++
    }
    catch {
        Write-Host "└─ ❌ FAIL: $_" -ForegroundColor Red
        Write-Host ""
        $script:failed++
    }
}

# ─── テスト1: サンプル Excel の生成 ───
Test-Step "テスト用 Excel 生成" {
    & "$testDir\Create-SampleExcel.ps1"
    $sampleXlsx = Join-Path $testDir "sample.xlsx"
    if (-not (Test-Path $sampleXlsx)) { throw "sample.xlsx が生成されませんでした" }
}

# ─── テスト2: Excel → Markdown（直接） ───
Test-Step "Excel → Markdown（直接変換）" {
    $input_ = Join-Path $testDir "sample.xlsx"
    $output_ = Join-Path $outputDir "direct_excel2md.md"
    & "$scriptDir\Convert-ExcelToMarkdown.ps1" -InputFile $input_ -OutputFile $output_
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }
    $content = Get-Content $output_ -Raw
    if ($content -notmatch '\|') { throw "マークダウンテーブルが含まれていません" }
    Write-Host "  📄 出力内容（先頭5行）:" -ForegroundColor Gray
    Get-Content $output_ -TotalCount 5 | ForEach-Object { Write-Host "     $_" -ForegroundColor Gray }
}

# ─── テスト3: Excel → Markdown（全シート） ───
Test-Step "Excel → Markdown（全シート変換）" {
    $input_ = Join-Path $testDir "sample.xlsx"
    $output_ = Join-Path $outputDir "direct_excel2md_allsheets.md"
    & "$scriptDir\Convert-ExcelToMarkdown.ps1" -InputFile $input_ -OutputFile $output_ -AllSheets
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }
    $content = Get-Content $output_ -Raw
    if ($content -notmatch 'プロジェクト管理表') { throw "シート1のタイトルが見つかりません" }
    if ($content -notmatch '部材リスト') { throw "シート2のタイトルが見つかりません" }
}

# ─── テスト4: Markdown → Excel（直接） ───
Test-Step "Markdown → Excel（直接変換）" {
    $input_ = Join-Path $testDir "sample.md"
    $output_ = Join-Path $outputDir "direct_md2excel.xlsx"
    & "$scriptDir\Convert-MarkdownToExcel.ps1" -InputFile $input_ -OutputFile $output_
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }
    $fileSize = (Get-Item $output_).Length
    if ($fileSize -lt 1000) { throw "Excel ファイルが小さすぎます ($fileSize bytes)" }

    $verifyMd = Join-Path $outputDir "direct_md2excel_verify.md"
    & "$scriptDir\Convert-ExcelToMarkdown.ps1" -InputFile $output_ -OutputFile $verifyMd -AllSheets
    $verifyContent = Get-Content $verifyMd -Raw
    if ($verifyContent -notmatch '##\s+プロジェクト管理表') { throw "シート名 'プロジェクト管理表' が反映されていません" }
    if ($verifyContent -notmatch '##\s+部材リスト') { throw "シート名 '部材リスト' が反映されていません" }
}

# ─── テスト5: Excel → CSV ───
Test-Step "Excel → CSV" {
    $input_ = Join-Path $testDir "sample.xlsx"
    $output_ = Join-Path $outputDir "excel2csv.csv"
    & "$scriptDir\Convert-ExcelToCsv.ps1" -InputFile $input_ -OutputFile $output_
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }
    $content = Get-Content $output_ -Raw
    if ($content -notmatch ',') { throw "CSV にカンマが含まれていません" }
    Write-Host "  📄 出力内容（先頭3行）:" -ForegroundColor Gray
    Get-Content $output_ -TotalCount 3 | ForEach-Object { Write-Host "     $_" -ForegroundColor Gray }
}

# ─── テスト6: CSV → Markdown ───
Test-Step "CSV → Markdown" {
    $input_ = Join-Path $outputDir "excel2csv.csv"
    $output_ = Join-Path $outputDir "csv2md.md"
    & "$scriptDir\Convert-CsvToMarkdown.ps1" -InputFile $input_ -OutputFile $output_
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }
    $content = Get-Content $output_ -Raw
    if ($content -notmatch '\|') { throw "マークダウンテーブルが含まれていません" }
}

# ─── テスト7: Markdown → CSV ───
Test-Step "Markdown → CSV" {
    $input_ = Join-Path $testDir "sample.md"
    $output_ = Join-Path $outputDir "md2csv.csv"
    & "$scriptDir\Convert-MarkdownToCsv.ps1" -InputFile $input_ -OutputFile $output_
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }
    $content = Get-Content $output_ -Raw
    if ($content -notmatch ',') { throw "CSV にカンマが含まれていません" }
}

# ─── テスト8: CSV → Excel ───
Test-Step "CSV → Excel" {
    $input_ = Join-Path $outputDir "excel2csv.csv"
    $output_ = Join-Path $outputDir "csv2excel.xlsx"
    & "$scriptDir\Convert-CsvToExcel.ps1" -InputFile $input_ -OutputFile $output_
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }
}

# ─── テスト9: ラウンドトリップ（Excel → MD → Excel） ───
Test-Step "ラウンドトリップ: Excel → MD → Excel" {
    $md_ = Join-Path $outputDir "direct_excel2md.md"
    $roundtrip_ = Join-Path $outputDir "roundtrip.xlsx"
    & "$scriptDir\Convert-MarkdownToExcel.ps1" -InputFile $md_ -OutputFile $roundtrip_
    if (-not (Test-Path $roundtrip_)) { throw "ラウンドトリップ Excel が生成されませんでした" }
}

# ─── テスト10: CSV パイプライン（Excel → CSV → MD → CSV → Excel） ───
Test-Step "CSV パイプライン: Excel → CSV → MD → CSV → Excel" {
    $csv1 = Join-Path $outputDir "pipeline_step1.csv"
    $md1 = Join-Path $outputDir "pipeline_step2.md"
    $csv2 = Join-Path $outputDir "pipeline_step3.csv"
    $xlsx1 = Join-Path $outputDir "pipeline_step4.xlsx"

    & "$scriptDir\Convert-ExcelToCsv.ps1" -InputFile (Join-Path $testDir "sample.xlsx") -OutputFile $csv1
    & "$scriptDir\Convert-CsvToMarkdown.ps1" -InputFile $csv1 -OutputFile $md1
    & "$scriptDir\Convert-MarkdownToCsv.ps1" -InputFile $md1 -OutputFile $csv2
    & "$scriptDir\Convert-CsvToExcel.ps1" -InputFile $csv2 -OutputFile $xlsx1

    if (-not (Test-Path $xlsx1)) { throw "パイプライン最終出力が生成されませんでした" }
}

# ─── テスト11: Markdown のエスケープパイプ（CSV） ───
Test-Step "Markdown → CSV（エスケープパイプ保持）" {
    $input_ = Join-Path $outputDir "escaped_pipe.md"
    $output_ = Join-Path $outputDir "escaped_pipe.csv"

    @(
        "| Col1 | Col2 |"
        "|------|------|"
        "| a\|b | c |"
    ) | Set-Content -Path $input_ -Encoding UTF8

    & "$scriptDir\Convert-MarkdownToCsv.ps1" -InputFile $input_ -OutputFile $output_
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }

    $rows = @(Import-Csv -Path $output_ -Encoding UTF8)
    if ($rows.Count -ne 1) { throw "行数不正: $($rows.Count)" }
    if ($rows[0].Col1 -ne "a|b") { throw "Col1 が不正: '$($rows[0].Col1)'" }
    if ($rows[0].Col2 -ne "c") { throw "Col2 が不正: '$($rows[0].Col2)'" }
}

# ─── テスト12: Markdown のエスケープパイプ（Excel） ───
Test-Step "Markdown → Excel（エスケープパイプ保持）" {
    $input_ = Join-Path $outputDir "escaped_pipe_excel.md"
    $xlsx_ = Join-Path $outputDir "escaped_pipe_excel.xlsx"
    $csv_ = Join-Path $outputDir "escaped_pipe_excel.csv"

    @(
        "| Col1 | Col2 |"
        "|------|------|"
        "| a\|b | c |"
    ) | Set-Content -Path $input_ -Encoding UTF8

    & "$scriptDir\Convert-MarkdownToExcel.ps1" -InputFile $input_ -OutputFile $xlsx_
    & "$scriptDir\Convert-ExcelToCsv.ps1" -InputFile $xlsx_ -OutputFile $csv_

    $rows = @(Import-Csv -Path $csv_ -Encoding UTF8)
    if ($rows.Count -ne 1) { throw "行数不正: $($rows.Count)" }
    if ($rows[0].Col1 -ne "a|b") { throw "Col1 が不正: '$($rows[0].Col1)'" }
    if ($rows[0].Col2 -ne "c") { throw "Col2 が不正: '$($rows[0].Col2)'" }
}

# ─── テスト13: Markdown 複数テーブルのCSV連結 ───
Test-Step "Markdown → CSV（複数テーブル保持）" {
    $input_ = Join-Path $outputDir "multi_table.md"
    $output_ = Join-Path $outputDir "multi_table.csv"
    if (Test-Path $output_) { Remove-Item -Force $output_ }

    @(
        "## First"
        ""
        "| A |"
        "|---|"
        "| 1 |"
        ""
        "## Second"
        ""
        "| B |"
        "|---|"
        "| 2 |"
    ) | Set-Content -Path $input_ -Encoding UTF8

    & "$scriptDir\Convert-MarkdownToCsv.ps1" -InputFile $input_ -OutputFile $output_
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }
    $lines = Get-Content -Path $output_
    if ($lines -notcontains "A") { throw "1つ目テーブルのヘッダーが欠落しています" }
    if ($lines -notcontains "B") { throw "2つ目テーブルのヘッダーが欠落しています" }
    if ($lines -notcontains "2") { throw "2つ目テーブルのデータが欠落しています" }
}

# ─── テスト14: スクリプト失敗時の終了コード ───
Test-Step "スクリプト失敗時に非0終了コードを返す" {
    $missingInput = Join-Path $outputDir "not_found.md"
    $output_ = Join-Path $outputDir "not_found.csv"
    if (Test-Path $output_) { Remove-Item -Force $output_ }

    $proc = Start-Process -FilePath "powershell" `
        -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "$scriptDir\Convert-MarkdownToCsv.ps1", "-InputFile", $missingInput, "-OutputFile", $output_) `
        -WindowStyle Hidden -Wait -PassThru
    if ($proc.ExitCode -eq 0) { throw "失敗時に 0 が返されました" }
}

# ─── テスト15: BAT失敗時の終了コード伝播 ───
Test-Step "BAT 失敗時に非0終了コードを返す" {
    $missingInput = Join-Path $outputDir "not_found_bat.md"
    $batFile = Join-Path $projectRoot "bat\md2csv.bat"
    $batCommand = "`"$batFile`" `"$missingInput`" --no-pause"
    $proc = Start-Process -FilePath "cmd.exe" -ArgumentList @("/c", $batCommand) -WindowStyle Hidden -Wait -PassThru
    if ($proc.ExitCode -eq 0) { throw "BAT が失敗終了コードを返していません" }
}

# ─── テスト16: 外側パイプ無しMarkdown → CSV ───
Test-Step "Markdown → CSV（外側パイプ無し表記）" {
    $input_ = Join-Path $outputDir "no_outer_pipe.md"
    $output_ = Join-Path $outputDir "no_outer_pipe.csv"

    @(
        "A | B"
        "---|---"
        "1 | 2"
    ) | Set-Content -Path $input_ -Encoding UTF8

    & "$scriptDir\Convert-MarkdownToCsv.ps1" -InputFile $input_ -OutputFile $output_
    if (-not (Test-Path $output_)) { throw "出力ファイルが生成されませんでした" }

    $rows = @(Import-Csv -Path $output_ -Encoding UTF8)
    if ($rows.Count -ne 1) { throw "行数不正: $($rows.Count)" }
    if ($rows[0].A -ne "1") { throw "A列が不正: '$($rows[0].A)'" }
    if ($rows[0].B -ne "2") { throw "B列が不正: '$($rows[0].B)'" }
}

# ─── テスト17: 外側パイプ無しMarkdown → Excel ───
Test-Step "Markdown → Excel（外側パイプ無し表記）" {
    $input_ = Join-Path $outputDir "no_outer_pipe_excel.md"
    $xlsx_ = Join-Path $outputDir "no_outer_pipe_excel.xlsx"
    $csv_ = Join-Path $outputDir "no_outer_pipe_excel.csv"

    @(
        "A | B"
        "---|---"
        "1 | 2"
    ) | Set-Content -Path $input_ -Encoding UTF8

    & "$scriptDir\Convert-MarkdownToExcel.ps1" -InputFile $input_ -OutputFile $xlsx_
    & "$scriptDir\Convert-ExcelToCsv.ps1" -InputFile $xlsx_ -OutputFile $csv_

    $rows = @(Import-Csv -Path $csv_ -Encoding UTF8)
    if ($rows.Count -ne 1) { throw "行数不正: $($rows.Count)" }
    if ($rows[0].A -ne "1") { throw "A列が不正: '$($rows[0].A)'" }
    if ($rows[0].B -ne "2") { throw "B列が不正: '$($rows[0].B)'" }
}

# ─── テスト18: 列不整合CSVはMarkdown変換で失敗 ───
Test-Step "CSV 列不整合時に Markdown 変換が失敗する" {
    $input_ = Join-Path $outputDir "invalid_columns_md.csv"
    @(
        "A,B"
        "1,2"
        "3,4,5"
    ) | Set-Content -Path $input_ -Encoding UTF8

    $proc = Start-Process -FilePath "powershell" `
        -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "$scriptDir\Convert-CsvToMarkdown.ps1", "-InputFile", $input_) `
        -WindowStyle Hidden -Wait -PassThru
    if ($proc.ExitCode -eq 0) { throw "列不整合CSVで成功してしまいました" }
}

# ─── テスト19: 列不整合CSVはExcel変換で失敗 ───
Test-Step "CSV 列不整合時に Excel 変換が失敗する" {
    $input_ = Join-Path $outputDir "invalid_columns_excel.csv"
    @(
        "A,B"
        "1,2"
        "3,4,5"
    ) | Set-Content -Path $input_ -Encoding UTF8

    $proc = Start-Process -FilePath "powershell" `
        -ArgumentList @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "$scriptDir\Convert-CsvToExcel.ps1", "-InputFile", $input_) `
        -WindowStyle Hidden -Wait -PassThru
    if ($proc.ExitCode -eq 0) { throw "列不整合CSVで成功してしまいました" }
}

# ─── 結果サマリー ───
Write-Host ""
Write-Host "╔══════════════════════════════════════════════════╗" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })
Write-Host "║  テスト結果: $passed/$total PASS, $failed/$total FAIL" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })
Write-Host "╚══════════════════════════════════════════════════╝" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })
Write-Host ""

if ($failed -gt 0) { exit 1 }
