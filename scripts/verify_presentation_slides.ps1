Add-Type -AssemblyName "System.IO.Compression.FileSystem"
Add-Type -AssemblyName "System.Xml.Linq"

$pptxPath = "c:\Users\ADMIN\Downloads\AI Agent CK\outputs\Thuyet_Trinh_He_Thong_Phan_Tich_Du_Lieu_Doanh_Nghiep.pptx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($pptxPath)

for ($i = 1; $i -le 10; $i++) {
    $entry = $zip.GetEntry("ppt/slides/slide$i.xml")
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    $xml = $reader.ReadToEnd()
    $reader.Close()
    $stream.Close()

    $xdoc = [System.Xml.Linq.XDocument]::Parse($xml)
    $nsA = [System.Xml.Linq.XNamespace]::Get("http://schemas.openxmlformats.org/drawingml/2006/main")
    Write-Output ("=== SLIDE $i ===")
    foreach ($p in $xdoc.Descendants($nsA + "p")) {
        $pTexts = @($p.Descendants($nsA + "t") | ForEach-Object { $_.Value })
        if ($pTexts.Count -gt 0) {
            Write-Output ("- " + ($pTexts -join ""))
        }
    }
}
$zip.Dispose()
