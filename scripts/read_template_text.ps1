Add-Type -AssemblyName "System.IO.Compression.FileSystem"
Add-Type -AssemblyName "System.Xml.Linq"

$pptxPath = "C:\Users\ADMIN\.gemini\antigravity\brain\e7d8edc4-dee3-4d3a-8049-d9e8951fcc8f\.user_uploaded\media_1794963781967_0aef2d03.pptx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($pptxPath)

for ($i = 1; $i -le 10; $i++) {
    $entry = $zip.GetEntry("ppt/slides/slide$i.xml")
    if ($entry) {
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $xmlContent = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()

        $xdoc = [System.Xml.Linq.XDocument]::Parse($xmlContent)
        $nsA = [System.Xml.Linq.XNamespace]::Get("http://schemas.openxmlformats.org/drawingml/2006/main")
        $texts = @($xdoc.Descendants($nsA + "t") | ForEach-Object { $_.Value.Trim() } | Where-Object { $_ -ne "" })

        Write-Output ("=== SLIDE " + $i + " (Count: " + $texts.Count + ") ===")
        Write-Output ($texts -join " | ")
    }
}

$zip.Dispose()
