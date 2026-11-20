Add-Type -AssemblyName "System.IO.Compression.FileSystem"
$zip = [System.IO.Compression.ZipFile]::OpenRead("c:\Users\ADMIN\Downloads\AI Agent CK\outputs\Thuyet_Trinh_He_Thong_Phan_Tich_Du_Lieu_Doanh_Nghiep.pptx")
Write-Output ("First 10 entries:")
for ($i = 0; $i -lt [Math]::Min(10, $zip.Entries.Count); $i++) {
    Write-Output ($zip.Entries[$i].FullName)
}
$zip.Dispose()
