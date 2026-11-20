# ==============================================================================
# SCRIPT TO BUILD A SAMPLE TEST CUSTOMER EXCEL & JSON FILE FOR USER UPLOAD TESTING
# Files: data/Khach_Hang_Doanh_Nghiep_Moi_Test.xlsx
#        data/Khach_Hang_Doanh_Nghiep_Moi_Test.json
# ==============================================================================

Add-Type -AssemblyName "System.IO.Compression"
Add-Type -AssemblyName "System.IO.Compression.FileSystem"

$ProjectRoot = "c:\Users\ADMIN\Downloads\AI Agent CK"
$DataDir = Join-Path $ProjectRoot "data"
$targetXlsx = Join-Path $DataDir "Khach_Hang_Doanh_Nghiep_Moi_Test.xlsx"
$targetJson = Join-Path $DataDir "Khach_Hang_Doanh_Nghiep_Moi_Test.json"

$testData = @(
    [ordered]@{
        "Mã Khách Hàng" = "DN-CUSTOM-001"
        "Tên Doanh Nghiệp" = "Công ty CP Dịch Vụ Hàng Hải Biển Đông"
        "Người Đại Diện" = "Vũ Đình Cường"
        "Lĩnh Vực" = "Vận Tải Biển & Cảng Biển"
        "Phân Khúc" = "Enterprise VIP"
        "Gói Dịch Vụ Hiện Tại" = "Hạ Tầng Điện Toán Mở Rộng"
        "Hạn Hợp Đồng" = "2026-12-25"
        "Tỷ Lệ Sử Dụng (%)" = 95.0
        "Sự Cố Kỹ Thuật" = "Không có sự cố"
        "Giá Trị Hợp Đồng" = 120000000
        "Chiết Khấu (%)" = 10
    },
    [ordered]@{
        "Mã Khách Hàng" = "DN-CUSTOM-002"
        "Tên Doanh Nghiệp" = "Tập đoàn Dệt May Thăng Long"
        "Người Đại Diện" = "Hoàng Kim Oanh"
        "Lĩnh Vực" = "Sản Xuất & Xuất Khẩu"
        "Phân Khúc" = "SME"
        "Gói Dịch Vụ Hiện Tại" = "Điện Toán Đám Mây Cơ Bản"
        "Hạn Hợp Đồng" = "2026-12-18"
        "Tỷ Lệ Sử Dụng (%)" = 62.5
        "Sự Cố Kỹ Thuật" = "Không có sự cố"
        "Giá Trị Hợp Đồng" = 65000000
        "Chiết Khấu (%)" = 5
    },
    [ordered]@{
        "Mã Khách Hàng" = "DN-CUSTOM-003"
        "Tên Doanh Nghiệp" = "Chuỗi Nhà Thuốc An Tâm"
        "Người Đại Diện" = "Bùi Minh Tâm"
        "Lĩnh Vực" = "Bán Lẻ Y Tế & Dược Phẩm"
        "Phân Khúc" = "SME"
        "Gói Dịch Vụ Hiện Tại" = "Điện Toán Đám Mây Cơ Bản"
        "Hạn Hợp Đồng" = "2026-11-30"
        "Tỷ Lệ Sử Dụng (%)" = $null # NULL đo lường
        "Sự Cố Kỹ Thuật" = "Không có sự cố"
        "Giá Trị Hợp Đồng" = 45000000
        "Chiết Khấu (%)" = 0
    },
    [ordered]@{
        "Mã Khách Hàng" = "DN-CUSTOM-004"
        "Tên Doanh Nghiệp" = "Công ty TNHH Vận Tải Sao Vàng"
        "Người Đại Diện" = "Đỗ Cao Thắng"
        "Lĩnh Vực" = "Giao Nhận Nhanh & Kho Bãi"
        "Phân Khúc" = "Enterprise VIP"
        "Gói Dịch Vụ Hiện Tại" = "Hạ Tầng Điện Toán Mở Rộng"
        "Hạn Hợp Đồng" = "2026-12-10"
        "Tỷ Lệ Sử Dụng (%)" = 88.0
        "Sự Cố Kỹ Thuật" = "Sự cố Mở (Ticket TK-9901 CSDL)" # Kỹ thuật mở -> Hard-Block
        "Giá Trị Hợp Đồng" = 150000000
        "Chiết Khấu (%)" = 10
    },
    [ordered]@{
        "Mã Khách Hàng" = "DN-CUSTOM-005"
        "Tên Doanh Nghiệp" = "Công ty Năng Lượng Mặt Trời SolarTech"
        "Người Đại Diện" = "Nguyễn Văn Đạt"
        "Lĩnh Vực" = "Năng Lượng Tái Tạo"
        "Phân Khúc" = "Doanh Nghiệp Tiêu Chuẩn"
        "Gói Dịch Vụ Hiện Tại" = "Điện Toán Đám Mây Cơ Bản"
        "Hạn Hợp Đồng" = "2026-12-05"
        "Tỷ Lệ Sử Dụng (%)" = 0 # 0% thực tế
        "Sự Cố Kỹ Thuật" = "Không có sự cố"
        "Giá Trị Hợp Đồng" = 35000000
        "Chiết Khấu (%)" = 0
    }
)

# 1. Write JSON
$jsonText = $testData | ConvertTo-Json -Depth 5
[System.IO.File]::WriteAllText($targetJson, $jsonText, [System.Text.Encoding]::UTF8)

# 2. Write XLSX (OpenXML)
if (Test-Path $targetXlsx) { Remove-Item $targetXlsx -Force }
$tempDir = Join-Path $env:TEMP ("xlsx_custom_" + (Get-Random))
New-Item -ItemType Directory -Path $tempDir | Out-Null
$xlDir = Join-Path $tempDir "xl"
$worksheetsDir = Join-Path $xlDir "worksheets"
$relsDir = Join-Path $tempDir "_rels"
$xlRelsDir = Join-Path $xlDir "_rels"

New-Item -ItemType Directory -Path $worksheetsDir -Force | Out-Null
New-Item -ItemType Directory -Path $relsDir -Force | Out-Null
New-Item -ItemType Directory -Path $xlRelsDir -Force | Out-Null

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# [Content_Types].xml
$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "[Content_Types].xml"), $contentTypes.Trim(), $utf8NoBom)

# _rels/.rels
$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $relsDir ".rels"), $rootRels.Trim(), $utf8NoBom)

# xl/workbook.xml
$workbook = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Khach_Hang_Doanh_Nghiep" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
'@
[System.IO.File]::WriteAllText((Join-Path $xlDir "workbook.xml"), $workbook.Trim(), $utf8NoBom)

# xl/_rels/workbook.xml.rels
$xlRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $xlRelsDir "workbook.xml.rels"), $xlRels.Trim(), $utf8NoBom)

# sheet1.xml
$headers = @("Mã Khách Hàng", "Tên Doanh Nghiệp", "Người Đại Diện", "Lĩnh Vực", "Phân Khúc", "Gói Dịch Vụ Hiện Tại", "Hạn Hợp Đồng", "Tỷ Lệ Sử Dụng (%)", "Sự Cố Kỹ Thuật", "Giá Trị Hợp Đồng", "Chiết Khấu (%)")

function Get-ColLetter([int]$colIdx) {
    $div = $colIdx
    $colLetter = ""
    while ($div -gt 0) {
        $mod = ($div - 1) % 26
        $colLetter = [char](65 + $mod) + $colLetter
        $div = [int](($div - $mod) / 26)
    }
    return $colLetter
}

$sb = New-Object System.Text.StringBuilder
$sb.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>') | Out-Null
$sb.AppendLine('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">') | Out-Null
$sb.AppendLine('  <sheetData>') | Out-Null

# Header
$sb.AppendLine('    <row r="1">') | Out-Null
for ($c = 0; $c -lt $headers.Count; $c++) {
    $col = Get-ColLetter ($c + 1)
    $val = [System.Security.SecurityElement]::Escape($headers[$c])
    $sb.AppendLine("      <c r=`"${col}1`" t=`"inlineStr`"><is><t>$val</t></is></c>") | Out-Null
}
$sb.AppendLine('    </row>') | Out-Null

$rowNum = 2
foreach ($row in $testData) {
    $sb.AppendLine("    <row r=`"$rowNum`">") | Out-Null
    $c = 1
    foreach ($h in $headers) {
        $col = Get-ColLetter $c
        $val = $row[$h]
        if ($null -ne $val) {
            $escVal = [System.Security.SecurityElement]::Escape([string]$val)
            $sb.AppendLine("      <c r=`"${col}${rowNum}`" t=`"inlineStr`"><is><t>$escVal</t></is></c>") | Out-Null
        }
        $c++
    }
    $sb.AppendLine('    </row>') | Out-Null
    $rowNum++
}

$sb.AppendLine('  </sheetData>') | Out-Null
$sb.AppendLine('</worksheet>') | Out-Null

[System.IO.File]::WriteAllText((Join-Path $worksheetsDir "sheet1.xml"), $sb.ToString(), $utf8NoBom)
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $targetXlsx)
Remove-Item -Path $tempDir -Recurse -Force

Write-Output ("Created Test Custom Files:")
Write-Output ("1. " + $targetXlsx + " (" + (Get-Item $targetXlsx).Length + " bytes)")
Write-Output ("2. " + $targetJson + " (" + (Get-Item $targetJson).Length + " bytes)")
