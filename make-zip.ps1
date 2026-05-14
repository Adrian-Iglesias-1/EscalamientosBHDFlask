$dest = "C:\Users\the_f\OneDrive\Escritorio\EscalamientosApp-1.1\EscalamientosApp-v2.1.zip"
$src  = "C:\Users\the_f\OneDrive\Escritorio\EscalamientosApp-1.1"
$tmp  = "$env:TEMP\EscalamientosApp_build"

Remove-Item $dest -ErrorAction SilentlyContinue
Remove-Item $tmp -Recurse -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $tmp | Out-Null

$excludeDirs  = @('venv', '.git', '__pycache__', '.claude', 'fotos')
$excludeFiles = @('make-zip.ps1', 'PlanillaEscalamientos.xlsx')
$excludeExt   = @('.zip', '.log', '.db', '.sqlite', '.pyc')

Get-ChildItem $src -Recurse | Where-Object {
    $item = $_
    $skip = $false
    foreach ($d in $excludeDirs)  { if ($item.FullName -match "\\$d(\\|$)") { $skip = $true; break } }
    if ($excludeFiles -contains $item.Name) { $skip = $true }
    if ($excludeExt   -contains $item.Extension) { $skip = $true }
    -not $skip -and -not $item.PSIsContainer
} | ForEach-Object {
    $relative = $_.FullName.Substring($src.Length + 1)
    $target   = Join-Path $tmp $relative
    $dir      = Split-Path $target -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    Copy-Item $_.FullName -Destination $target -Force
}

Compress-Archive -Path "$tmp\*" -DestinationPath $dest
Remove-Item $tmp -Recurse -Force
Write-Host "Listo: $dest"
