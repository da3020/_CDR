# =====================================================
# FLAG ARCHIVE INDEX BUILDER
# Version: 1.0
# =====================================================

$ARCHIVE_ROOT = "\\Keenetic-5026\ugreen\STORE\! СУБЛИМАЦИЯ\! ! ! ФЛАГИ"
$SCRIPT_DIR = Split-Path -Parent $MyInvocation.MyCommand.Path
$INDEX_FILE = Join-Path $SCRIPT_DIR "_FLAG_INDEX.txt"

Write-Host "Indexing archive..."
Write-Host "Source: $ARCHIVE_ROOT"
Write-Host "Target: $INDEX_FILE"

$result = @()

Get-ChildItem $ARCHIVE_ROOT -Recurse -Filter *.cdr | ForEach-Object {

    # имя вида 000458_Флаг ....cdr
    if ($_.Name -match '^0*(\d+)_') {

        $article = [int]$matches[1]
        $path = $_.FullName

        $result += "$article|$path"
    }
}

$result |
Sort-Object |
Set-Content -Encoding UTF8 $INDEX_FILE

Write-Host "Done. Found $($result.Count) files."
