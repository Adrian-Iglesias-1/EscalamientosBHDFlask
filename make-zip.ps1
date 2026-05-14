$source = "C:\Users\the_f\OneDrive\Escritorio\EscalamientosApp-1.1\EscalamientosApp-1.2"
$dest = "C:\Users\the_f\OneDrive\Escritorio\EscalamientosApp-1.1\EscalamientosApp-v1.2.zip"
Remove-Item $dest -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($dest, 1)
$files = Get-ChildItem $source -Recurse -File | Where-Object { $_.FullName -notmatch "\\venv\\" -and $_.FullName -notmatch "__pycache__" }
$prefixLen = ($source.TrimEnd("\") + "\").Length
foreach ($f in $files) {
    $rel = $f.FullName.Substring($prefixLen)
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $f.FullName, $rel, 0) | Out-Null
}
$zip.Dispose()
$size = [math]::Round((Get-Item $dest).Length / 1MB, 1)
Write-Host "OK! $size MB - $($files.Count) archivos"
