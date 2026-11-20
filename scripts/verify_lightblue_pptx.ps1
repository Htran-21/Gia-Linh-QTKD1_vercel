Add-Type -AssemblyName "System.IO.Compression.FileSystem"
$pptxPath = "c:\Users\ADMIN\Downloads\AI Agent CK\outputs\Thuyet_Trinh_He_Thong_Phan_Tich_Du_Lieu_Doanh_Nghiep.pptx"

$zip = [System.IO.Compression.ZipFile]::OpenRead($pptxPath)
Write-Output ("Total Entries in PPTX: " + $zip.Entries.Count)

$slides = $zip.Entries | Where-Object { $_.FullName -match "^ppt/slides/slide\d+\.xml$" }
Write-Output ("Slide count: " + $slides.Count)

foreach ($s in $slides | Sort-Object { [int]($_.FullName -replace '\D', '') }) {
    $reader = New-Object System.IO.StreamReader($s.Open(), [System.Text.Encoding]::UTF8)
    $xmlText = $reader.ReadToEnd()
    $reader.Close()

    $xmlDoc = [xml]$xmlText
    $ns = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
    $ns.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main")
    $ns.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main")

    $titleNode = $xmlDoc.SelectSingleNode("//p:sp[p:nvSpPr/p:cNvPr[@name='SlideTitleBox']]//a:t", $ns)
    $t = if ($titleNode) { $titleNode.InnerText } else { "(No Title Node)" }
    Write-Output ($s.FullName + " -> XML OK | " + $t)
}
$zip.Dispose()
