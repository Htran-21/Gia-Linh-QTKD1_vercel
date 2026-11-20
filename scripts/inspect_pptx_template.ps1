Add-Type -AssemblyName "System.IO.Compression.FileSystem"

$pptxPath = "C:\Users\ADMIN\.gemini\antigravity\brain\e7d8edc4-dee3-4d3a-8049-d9e8951fcc8f\.user_uploaded\media_1794963781967_0aef2d03.pptx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($pptxPath)

Write-Output ("Total entries in PPTX template: " + $zip.Entries.Count)
$slides = @($zip.Entries | Where-Object { $_.FullName -match "^ppt/slides/slide\d+\.xml$" })
Write-Output ("Slide count: " + $slides.Count)

foreach ($s in $slides) {
    Write-Output ("Found: " + $s.FullName + " (Size: " + $s.Length + " bytes)")
}

$zip.Dispose()
