<#
.SYNOPSIS
    テスト用の Excel ファイルを生成します。

.DESCRIPTION
    動作テスト用のサンプル Excel ファイル (.xlsx) を作成します。
    日本語データを含む2つのシートを生成します。
#>

$ErrorActionPreference = "Stop"

$outputDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$outputFile = Join-Path $outputDir "sample.xlsx"

# 既存ファイルがあれば削除
if (Test-Path $outputFile) {
    Remove-Item $outputFile -Force
}

Write-Host "⏳ テスト用 Excel を作成中..." -ForegroundColor Yellow

$excel = $null
$workbook = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Add()

    # ─── シート1: プロジェクト管理表 ───
    $sheet1 = $workbook.Sheets.Item(1)
    $sheet1.Name = "プロジェクト管理表"

    $headers1 = @("ID", "タスク名", "担当者", "ステータス", "優先度", "期限")
    $data1 = @(
        @("1", "要件定義", "田中", "完了", "高", "2026-01-15"),
        @("2", "基本設計", "鈴木", "進行中", "高", "2026-02-01"),
        @("3", "詳細設計", "佐藤", "未着手", "中", "2026-02-15"),
        @("4", "実装", "田中", "未着手", "高", "2026-03-01"),
        @("5", "テスト", "山田", "未着手", "中", "2026-03-15")
    )

    for ($c = 0; $c -lt $headers1.Count; $c++) {
        $sheet1.Cells.Item(1, $c + 1).Value2 = $headers1[$c]
    }
    for ($r = 0; $r -lt $data1.Count; $r++) {
        for ($c = 0; $c -lt $data1[$r].Count; $c++) {
            $sheet1.Cells.Item($r + 2, $c + 1).Value2 = $data1[$r][$c]
        }
    }

    # ヘッダー装飾
    $hdr1 = $sheet1.Range($sheet1.Cells.Item(1, 1), $sheet1.Cells.Item(1, $headers1.Count))
    $hdr1.Font.Bold = $true
    $hdr1.Interior.Color = 0xD9E1F2
    $sheet1.UsedRange.EntireColumn.AutoFit() | Out-Null

    # ─── シート2: 部材リスト ───
    $sheet2 = $workbook.Sheets.Add([System.Reflection.Missing]::Value, $sheet1)
    $sheet2.Name = "部材リスト"

    $headers2 = @("品番", "品名", "数量", "単価", "金額", "備考")
    $data2 = @(
        @("A-001", "ボルト M10×30", "100", "50", "5,000", "ステンレス"),
        @("A-002", "ナット M10", "100", "30", "3,000", "ステンレス"),
        @("B-001", "フランジ 50A", "10", "2,500", "25,000", "SUS304"),
        @("C-001", "パイプ 50A×2m", "5", "8,000", "40,000", "SGP")
    )

    for ($c = 0; $c -lt $headers2.Count; $c++) {
        $sheet2.Cells.Item(1, $c + 1).Value2 = $headers2[$c]
    }
    for ($r = 0; $r -lt $data2.Count; $r++) {
        for ($c = 0; $c -lt $data2[$r].Count; $c++) {
            $sheet2.Cells.Item($r + 2, $c + 1).Value2 = $data2[$r][$c]
        }
    }

    $hdr2 = $sheet2.Range($sheet2.Cells.Item(1, 1), $sheet2.Cells.Item(1, $headers2.Count))
    $hdr2.Font.Bold = $true
    $hdr2.Interior.Color = 0xD9E1F2
    $sheet2.UsedRange.EntireColumn.AutoFit() | Out-Null

    $sheet1.Activate()

    # 保存 (xlOpenXMLWorkbook = 51)
    $workbook.SaveAs($outputFile, 51)

    Write-Host "✅ テスト用 Excel を作成しました: $outputFile" -ForegroundColor Green
}
catch {
    Write-Error "❌ エラー: $_"
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
