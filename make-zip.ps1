$dest = "C:\Users\the_f\OneDrive\Escritorio\EscalamientosApp-1.1\EscalamientosApp-v2.1.zip"
$src  = "C:\Users\the_f\OneDrive\Escritorio\EscalamientosApp-1.1"

Remove-Item $dest -ErrorAction SilentlyContinue

$items = Get-ChildItem $src | Where-Object { $_.Name -ne ".git" -and $_.Name -notmatch "\.zip$" -and $_.Name -ne "make-zip.ps1" }
Compress-Archive -Path $items.FullName -DestinationPath $dest
Write-Host "Creado: $dest"
