Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-Docx($path) {
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    $entry = $zip.GetEntry('word/document.xml')
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $content = $reader.ReadToEnd()
    $reader.Close()
    $zip.Dispose()
    $text = $content -replace '<[^>]+>', ' '
    $text = $text -replace '\s+', ' '
    return $text.Substring(0, [Math]::Min(8000, $text.Length))
}

$folder = 'c:/Users/Redouane/Portfolio-V3/assets/pdf/'

Write-Host "=== PROJET 1 ==="
Write-Host (Read-Docx ($folder + 'Documentation_Technique_BTS_SIO_Redouane_Assani.docx'))

Write-Host ""
Write-Host "=== PROJET 2 ==="
Write-Host (Read-Docx ($folder + 'Documentation_Projet2_v2.docx'))
