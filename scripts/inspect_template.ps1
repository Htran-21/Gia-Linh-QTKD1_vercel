$pptxTemplate = "C:\Users\ADMIN\.gemini\antigravity\brain\e7d8edc4-dee3-4d3a-8049-d9e8951fcc8f\.user_uploaded\media_1794963781967_0aef2d03.pptx"
Write-Output ("Template exists: " + (Test-Path $pptxTemplate))
if (Test-Path $pptxTemplate) {
    Write-Output ("File size: " + (Get-Item $pptxTemplate).Length)
}

# Check Word / PowerPoint COM
$wordAvailable = $false
try {
    $w = New-Object -ComObject Word.Application
    $wordAvailable = $true
    $w.Quit()
} catch {}
Write-Output ("Word COM available: " + $wordAvailable)
