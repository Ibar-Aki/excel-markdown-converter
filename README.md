# 📊 Excel ↔ Markdown 変換ツール
更新日: 2026-02-27

Excel ファイルとマークダウンテーブルを双方向に変換する PowerShell ツール集です。  
**ソフトのインストール不要**。Windows + M365（Excel）の標準環境だけで動作します。

---

## 📁 ファイル構成

```
void-curiosity/
├── scripts/                          # PowerShell スクリプト
│   ├── Convert-ExcelToMarkdown.ps1   # Excel → Markdown
│   ├── Convert-MarkdownToExcel.ps1   # Markdown → Excel
│   ├── Convert-ExcelToCsv.ps1        # Excel → CSV
│   ├── Convert-CsvToMarkdown.ps1     # CSV → Markdown
│   ├── Convert-MarkdownToCsv.ps1     # Markdown → CSV
│   └── Convert-CsvToExcel.ps1        # CSV → Excel
├── bat/                              # BAT ファイル（ドラッグ＆ドロップ対応）
│   ├── excel2md.bat                  # Excel → Markdown
│   ├── md2excel.bat                  # Markdown → Excel
│   ├── excel2csv.bat                 # Excel → CSV
│   ├── csv2md.bat                    # CSV → Markdown
│   ├── md2csv.bat                    # Markdown → CSV
│   └── csv2excel.bat                 # CSV → Excel
├── test/                             # テスト
│   ├── sample.md                     # テスト用マークダウン
│   ├── Create-SampleExcel.ps1        # テスト用 Excel 生成スクリプト
│   └── Run-AllTests.ps1              # 全変換テストスクリプト
└── README.md                         # この説明書
```

---

## 🚀 使い方

### 方法1: ドラッグ＆ドロップ（一番簡単）

1. `bat/` フォルダ内の BAT ファイルを開く
2. 変換したいファイルを BAT ファイルにドラッグ＆ドロップ
3. 変換完了！入力ファイルと同じフォルダに出力ファイルが生成されます

```
📄 data.xlsx  →  🖱️ excel2md.bat にドロップ  →  📄 data.md
📄 table.md   →  🖱️ md2excel.bat にドロップ  →  📄 table.xlsx
```

### 方法2: コマンドライン

```powershell
# Excel → マークダウン
.\scripts\Convert-ExcelToMarkdown.ps1 -InputFile .\data.xlsx

# マークダウン → Excel
.\scripts\Convert-MarkdownToExcel.ps1 -InputFile .\table.md

# 出力先を指定
.\scripts\Convert-ExcelToMarkdown.ps1 -InputFile .\data.xlsx -OutputFile .\output.md

# 特定のシートを指定
.\scripts\Convert-ExcelToMarkdown.ps1 -InputFile .\data.xlsx -SheetName "Sheet2"

# 全シートを変換
.\scripts\Convert-ExcelToMarkdown.ps1 -InputFile .\data.xlsx -AllSheets
```

### 方法3: BAT ファイルをコマンドラインから実行

```cmd
bat\excel2md.bat C:\path\to\data.xlsx
bat\excel2md.bat C:\path\to\data.xlsx C:\path\to\output.md
bat\excel2md.bat C:\path\to\data.xlsx --no-pause
```

---

## 🔄 変換パターン

6 つのスクリプトで、以下の全変換パターンをカバーしています。

```
  ┌──────────┐        直接変換         ┌──────────┐
  │  Excel   │ ◄══════════════════════► │ Markdown │
  │ .xlsx    │  excel2md / md2excel     │   .md    │
  └────┬─────┘                          └────┬─────┘
       │                                     │
       │  excel2csv     csv2md               │  md2csv
       │  csv2excel                          │
       ▼                                     ▼
  ┌──────────┐        csv2md            ┌──────────┐
  │   CSV    │ ════════════════════════► │   CSV    │
  │  .csv    │ ◄════════════════════════ │  .csv    │
  └──────────┘        md2csv            └──────────┘
```

### 直接変換（推奨）

| BAT ファイル | スクリプト | 変換方向 | Excel COM |
|:---:|:---|:---|:---:|
| `excel2md.bat` | `Convert-ExcelToMarkdown.ps1` | Excel → Markdown | 必要 |
| `md2excel.bat` | `Convert-MarkdownToExcel.ps1` | Markdown → Excel | 必要 |

### CSV 経由変換

| BAT ファイル | スクリプト | 変換方向 | Excel COM |
|:---:|:---|:---|:---:|
| `excel2csv.bat` | `Convert-ExcelToCsv.ps1` | Excel → CSV | 必要 |
| `csv2md.bat` | `Convert-CsvToMarkdown.ps1` | CSV → Markdown | **不要** |
| `md2csv.bat` | `Convert-MarkdownToCsv.ps1` | Markdown → CSV | **不要** |
| `csv2excel.bat` | `Convert-CsvToExcel.ps1` | CSV → Excel | 必要 |

> **💡 Tips:** CSV 経由の変換は2ステップになりますが、CSV → Markdown / Markdown → CSV は
> **Excel COM が不要**なので、Office がない環境でも使えます。

---

## 📋 各スクリプトの詳細

### Convert-ExcelToMarkdown.ps1

Excel ファイルをマークダウンテーブルに変換します。

| パラメータ | 必須 | 説明 |
|:---|:---:|:---|
| `-InputFile` | ✅ | 入力 Excel ファイルパス（.xlsx / .xls / .xlsm） |
| `-OutputFile` | - | 出力先（省略時: 同名 .md） |
| `-SheetName` | - | 変換対象のシート名（省略時: 最初のシート） |
| `-AllSheets` | - | 全シートを変換（各シートは見出し付き） |

**特徴:**

- 列幅を自動計算して整列されたテーブルを生成
- セル内のパイプ文字(`|`)を自動エスケープ
- セル内改行をスペースに変換

---

### Convert-MarkdownToExcel.ps1

マークダウンテーブルを Excel ファイルに変換します。

| パラメータ | 必須 | 説明 |
|:---|:---:|:---|
| `-InputFile` | ✅ | 入力マークダウンファイルパス |
| `-OutputFile` | - | 出力先（省略時: 同名 .xlsx） |

**特徴:**

- 複数テーブルを検出 → 各テーブルを別シートとして出力
- 外側パイプなしテーブル記法（`A | B`）にも対応
- ヘッダー行を太字＋薄い青背景で装飾
- 罫線（薄グレー）を自動追加
- 列幅を AutoFit で自動調整

---

### Convert-ExcelToCsv.ps1

Excel ファイルを CSV（UTF-8 BOM 付き）に変換します。

| パラメータ | 必須 | 説明 |
|:---|:---:|:---|
| `-InputFile` | ✅ | 入力 Excel ファイルパス |
| `-OutputFile` | - | 出力先（省略時: 同名 .csv） |
| `-SheetName` | - | 変換対象のシート名 |

**特徴:**

- UTF-8 BOM 付き出力（Excel で開いても文字化けしない）
- カンマ・改行・ダブルクォートを含むセルを適切にクォート

---

### Convert-CsvToMarkdown.ps1

CSV ファイルをマークダウンテーブルに変換します。**Excel COM 不要**。

| パラメータ | 必須 | 説明 |
|:---|:---:|:---|
| `-InputFile` | ✅ | 入力 CSV ファイルパス |
| `-OutputFile` | - | 出力先（省略時: 同名 .md） |

**特徴:**

- CSV の列数不整合を検知した場合はエラー終了（データ欠損を防止）

---

### Convert-MarkdownToCsv.ps1

マークダウンテーブルを CSV に変換します。**Excel COM 不要**。

| パラメータ | 必須 | 説明 |
|:---|:---:|:---|
| `-InputFile` | ✅ | 入力マークダウンファイルパス |
| `-OutputFile` | - | 出力先（省略時: 同名 .csv） |

**特徴:**

- エスケープ済みパイプ文字（`\|`）をセル内文字として復元
- 外側パイプなしテーブル記法（`A | B`）にも対応
- 複数テーブルを検出し、空行区切りで 1 つの CSV に連結出力

---

### Convert-CsvToExcel.ps1

CSV ファイルを Excel ファイル（.xlsx）に変換します。

| パラメータ | 必須 | 説明 |
|:---|:---:|:---|
| `-InputFile` | ✅ | 入力 CSV ファイルパス |
| `-OutputFile` | - | 出力先（省略時: 同名 .xlsx） |

**特徴:**

- ヘッダー行を装飾（太字・背景色）
- 罫線・列幅自動調整つき
- CSV の列数不整合を検知した場合はエラー終了（データ欠損を防止）

---

## 🧪 テスト

```powershell
# 全テストを実行（19パターン）
powershell -ExecutionPolicy Bypass -File .\test\Run-AllTests.ps1
```

テスト内容:

1. テスト用 Excel の生成
2. Excel → Markdown（直接）
3. Excel → Markdown（全シート）
4. Markdown → Excel（直接）
5. Excel → CSV
6. CSV → Markdown
7. Markdown → CSV
8. CSV → Excel
9. ラウンドトリップ: Excel → MD → Excel
10. CSV パイプライン: Excel → CSV → MD → CSV → Excel
11. Markdown → CSV（`\|` を保持）
12. Markdown → Excel（`\|` を保持）
13. Markdown → CSV（複数テーブル保持）
14. スクリプト失敗時の非0終了コード検証
15. BAT 失敗時の非0終了コード検証
16. Markdown → CSV（外側パイプ無し表記）
17. Markdown → Excel（外側パイプ無し表記）
18. CSV 列不整合時に Markdown 変換が失敗することを検証
19. CSV 列不整合時に Excel 変換が失敗することを検証

---

## 🔧 動作環境

| 要件 | 詳細 |
|:---|:---|
| OS | Windows 10 / 11 |
| PowerShell | 5.1 以降（Windows 標準搭載） |
| Office | M365 / Office 2016 以降（Excel COM に必要） |
| 追加インストール | **不要** |

> **⚠ 注意:** Excel COM を使用するスクリプトの実行中は、バックグラウンドで
> Excel プロセスが起動します。スクリプト終了時に自動で解放されます。

---

## 📖 技術解説: Excel COM と ImportExcel

### Excel COM オブジェクトとは？

**COM (Component Object Model)** は Microsoft が開発した技術で、
アプリケーション間でオブジェクトを共有する仕組みです。

Excel COM は、**Excel をプログラムから操作するためのインターフェース**です。
PowerShell から以下のように利用できます:

```powershell
# Excel COM の基本的な使い方
$excel = New-Object -ComObject Excel.Application   # Excel を起動
$excel.Visible = $false                             # 画面に表示しない
$excel.DisplayAlerts = $false                       # ダイアログを表示しない

$workbook = $excel.Workbooks.Open("C:\data.xlsx")   # ファイルを開く
$sheet = $workbook.Sheets.Item(1)                   # 最初のシートを取得
$value = $sheet.Cells.Item(1, 1).Text               # A1 セルの値を取得

# 必ず後片付け！
$workbook.Close($false)
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
```

#### Excel COM のメリット・デメリット

| ✅ メリット | ❌ デメリット |
|:---|:---|
| 追加インストール不要 | Excel のインストールが必要 |
| Excel の全機能にアクセス可能 | 処理速度がやや遅い（COM間通信のオーバーヘッド） |
| 書式・数式・グラフも扱える | COM オブジェクトの解放を正しく行わないとメモリリークする |
| .xls / .xlsx / .xlsm 全対応 | サーバー環境（非GUI）では動作しない場合がある |

#### COM オブジェクト解放のベストプラクティス

COM を使う際の**最も重要なポイント**はオブジェクトの解放です。
解放を忘れると Excel プロセスがゾンビ化（バックグラウンドに残り続ける）します。

```powershell
try {
    $excel = New-Object -ComObject Excel.Application
    # ... 処理 ...
}
finally {
    # 必ず finally ブロックで解放する
    if ($workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    # GC を強制実行
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
```

#### よく使う Excel COM のプロパティ・メソッド

```powershell
# ──── ワークブック操作 ────
$wb = $excel.Workbooks.Open("C:\file.xlsx")    # 開く
$wb = $excel.Workbooks.Add()                    # 新規作成
$wb.SaveAs("C:\output.xlsx", 51)               # 保存 (51 = xlOpenXMLWorkbook)
$wb.Close($false)                               # 閉じる（保存しない）

# ──── シート操作 ────
$sheet = $wb.Sheets.Item(1)                     # インデックスで取得
$sheet = $wb.Sheets.Item("Sheet1")              # 名前で取得
$sheet.Name = "新しい名前"                      # シート名変更
$newSheet = $wb.Sheets.Add()                    # シート追加

# ──── セル操作 ────
$cell = $sheet.Cells.Item(1, 1)                 # A1 セル
$cell.Value2 = "Hello"                          # 値の設定
$text = $cell.Text                              # 表示テキストの取得
$range = $sheet.Range("A1:C5")                  # セル範囲

# ──── 書式設定 ────
$cell.Font.Bold = $true                         # 太字
$cell.Font.Size = 12                            # フォントサイズ
$cell.Interior.Color = 0xFFFF00                 # 背景色（黄色）
$range.EntireColumn.AutoFit()                   # 列幅自動調整

# ──── UsedRange（データ範囲の取得） ────
$used = $sheet.UsedRange
$rowCount = $used.Rows.Count                    # 行数
$colCount = $used.Columns.Count                 # 列数
```

---

### ImportExcel モジュールとは？

**ImportExcel** は、Doug Finke 氏が作成した PowerShell モジュールで、
**Excel をインストールせずに** .xlsx ファイルを読み書きできます。

内部的には **EPPlus** (.NET の Excel 操作ライブラリ) を使用しており、
COM ではなく直接 .xlsx ファイル（実体は ZIP + XML）を操作します。

#### インストール方法

```powershell
# PowerShell Gallery からインストール（管理者権限不要）
Install-Module -Name ImportExcel -Scope CurrentUser
```

#### 基本的な使い方

```powershell
# ──── Excel → PowerShell オブジェクト ────
$data = Import-Excel -Path "C:\data.xlsx"
$data = Import-Excel -Path "C:\data.xlsx" -WorksheetName "Sheet2"

# ──── PowerShell オブジェクト → Excel ────
$data | Export-Excel -Path "C:\output.xlsx" -AutoSize -BoldTopRow
Get-Process | Export-Excel -Path "C:\processes.xlsx" -Show  # 作成後に Excel で開く

# ──── CSV → Excel（一発変換） ────
Import-Csv "C:\data.csv" | Export-Excel -Path "C:\output.xlsx" -AutoSize

# ──── Excel → マークダウン（パイプライン） ────
$data = Import-Excel -Path "C:\data.xlsx"
$data | ConvertTo-Markdown  # ※ PowerShell 7 以降で利用可能
```

#### Excel COM vs ImportExcel 比較

| 項目 | Excel COM | ImportExcel |
|:---|:---|:---|
| **Excel のインストール** | 必要 | **不要** |
| **追加モジュール** | 不要 | Install-Module が必要 |
| **処理速度** | やや遅い（COM通信） | **高速**（直接ファイル操作） |
| **対応形式** | .xls / .xlsx / .xlsm 全対応 | .xlsx のみ |
| **書式設定** | 全機能対応 | 基本的な書式のみ |
| **グラフ・ピボット** | 対応 | 一部対応 |
| **サーバー環境** | 非推奨 | **対応** |
| **メモリ管理** | 手動解放が必要 | 自動（.NET GC） |

#### ImportExcel の便利な機能

```powershell
# ──── 条件付き書式 ────
$data | Export-Excel -Path "C:\report.xlsx" -AutoSize -ConditionalText $(
    New-ConditionalText -Text "Error" -BackgroundColor Red -ConditionalTextColor White
)

# ──── チャートの追加 ────
$chartDef = New-ExcelChartDefinition -Title "売上推移" -ChartType Line -XRange "Month" -YRange "Sales"
$data | Export-Excel -Path "C:\report.xlsx" -ExcelChartDefinition $chartDef

# ──── ピボットテーブル ────
$data | Export-Excel -Path "C:\report.xlsx" -IncludePivotTable -PivotRows "Category" -PivotData @{Amount="Sum"}

# ──── 複数シートへの書き出し ────
$data1 | Export-Excel -Path "C:\report.xlsx" -WorksheetName "売上"
$data2 | Export-Excel -Path "C:\report.xlsx" -WorksheetName "経費" -Append
```

---

### 本ツールが Excel COM を採用した理由

| 理由 | 説明 |
|:---|:---|
| ゼロインストール | M365 環境なら追加モジュール不要 |
| .xls 対応 | 古い Excel ファイルにも対応可能 |
| 完全な互換性 | Excel 本体と同じエンジンで読み書き |

> **💡 Tips:** ImportExcel を使いたい場合は、`Install-Module -Name ImportExcel -Scope CurrentUser`
> でインストールすれば、本ツールを使わなくても同等の変換が可能です。

---

## 📝 変換例

### 入力: Excel テーブル

| ID | タスク名 | 担当者 | ステータス |
|----|----------|--------|------------|
| 1  | 要件定義 | 田中   | 完了       |
| 2  | 基本設計 | 鈴木   | 進行中     |
| 3  | 詳細設計 | 佐藤   | 未着手     |

### 出力: マークダウン

```markdown
| ID | タスク名 | 担当者 | ステータス |
|----|----------|--------|------------|
| 1  | 要件定義 | 田中   | 完了       |
| 2  | 基本設計 | 鈴木   | 進行中     |
| 3  | 詳細設計 | 佐藤   | 未着手     |
```

---

## ❓ トラブルシューティング

### 「Excel.Application が登録されていません」エラー

→ Excel / M365 がインストールされていない環境です。  
→ CSV 経由の変換（`csv2md.bat` / `md2csv.bat`）は Excel 不要で使えます。

### Excel プロセスが残る

→ スクリプトが異常終了した場合、タスクマネージャーで `EXCEL.EXE` を手動終了してください。

```powershell
# Excel プロセスを全て終了
Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
```

### 文字化けする

→ 出力は UTF-8 (BOM 付き) です。メモ帳やVSCodeで正しく表示されます。  
→ CSV を Excel で直接開く場合は BOM 付きなので文字化けしません。

### 実行ポリシーエラー

→ BAT ファイルから実行する場合は自動で `-ExecutionPolicy Bypass` されます。  
→ 直接実行する場合:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

---

## 📜 ライセンス

MIT License - 自由にご利用ください。
