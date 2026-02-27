# UTF-8 BOM を追加するスクリプト
$projectRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

$dirs = @(
    (Join-Path $projectRoot "scripts"),
    (Join-Path $projectRoot "test")
)

foreach ($dir in $dirs) {
    $files = Get-ChildItem -Path $dir -Filter "*.ps1"
    foreach ($f in $files) {
        $content = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
        $utf8Bom = New-Object System.Text.UTF8Encoding($true)
        [System.IO.File]::WriteAllText($f.FullName, $content, $utf8Bom)
        Write-Host "BOM added: $($f.Name)"
    }
}

# MDファイルにもBOM追加
$mdFile = Join-Path $projectRoot "test/sample.md"
if (Test-Path $mdFile) {
    $content = [System.IO.File]::ReadAllText($mdFile, [System.Text.Encoding]::UTF8)
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($mdFile, $content, $utf8Bom)
    Write-Host "BOM added: sample.md"
}
