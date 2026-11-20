# ==============================================================================
# SCRIPT TO EXPORT CLEANED ENTERPRISE EXCEL FILE
# Target: outputs/Doanh_Nghiep_Phan_Tich_Output.xlsx
# ==============================================================================

Add-Type -AssemblyName "System.IO.Compression"
Add-Type -AssemblyName "System.IO.Compression.FileSystem"

$ProjectRoot = "c:\Users\ADMIN\Downloads\AI Agent CK"
$jsonPath = Join-Path $ProjectRoot "outputs\enterprise_analysis_results.json"
$targetXlsx = Join-Path $ProjectRoot "outputs\Doanh_Nghiep_Phan_Tich_Output.xlsx"

$dataObj = Get-Content -Path $jsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
$profiles = $dataObj.profiles

if (Test-Path $targetXlsx) { Remove-Item $targetXlsx -Force }

$tempDir = Join-Path $env:TEMP ("xlsx_ent_" + (Get-Random))
New-Item -ItemType Directory -Path $tempDir | Out-Null
$xlDir = Join-Path $tempDir "xl"
$worksheetsDir = Join-Path $xlDir "worksheets"
$relsDir = Join-Path $tempDir "_rels"
$xlRelsDir = Join-Path $xlDir "_rels"

New-Item -ItemType Directory -Path $worksheetsDir -Force | Out-Null
New-Item -ItemType Directory -Path $relsDir -Force | Out-Null
New-Item -ItemType Directory -Path $xlRelsDir -Force | Out-Null

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# 1. [Content_Types].xml
$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "[Content_Types].xml"), $contentTypes.Trim(), $utf8NoBom)

# 2. _rels/.rels
$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $relsDir ".rels"), $rootRels.Trim(), $utf8NoBom)

# 3. xl/workbook.xml
$workbook = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Du_Lieu_Khach_Hang_Da_Lam_Sach" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
'@
[System.IO.File]::WriteAllText((Join-Path $xlDir "workbook.xml"), $workbook.Trim(), $utf8NoBom)

# 4. xl/_rels/workbook.xml.rels
$xlRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $xlRelsDir "workbook.xml.rels"), $xlRels.Trim(), $utf8NoBom)

# 5. xl/styles.xml
$styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><color theme="1"/><name val="Segoe UI"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0D9488"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
  </cellXfs>
</styleSheet>
'@
[System.IO.File]::WriteAllText((Join-Path $xlDir "styles.xml"), $styles.Trim(), $utf8NoBom)

# 6. Build xl/worksheets/sheet1.xml
$headers = @(
    "Ma Khach Hang", "Ten Doanh Nghiep", "Nguoi Dai Dien", "Linh Vuc", "Phan Khuc",
    "Goi Dich Vu Hien Tai", "Han Hop Dong", "So Ngay Con Lai", "Ty Le Su Dung (%)",
    "Su Co Ky Thuat", "De Xuat Hanh Dong", "Goi Dich Vu Moi", "Gia Goc (VND)",
    "Ma Uu Dai", "Chiet Khau (%)", "Gia Sau Cung (VND)", "Can Cu Doanh Thu", "Trang Thai",
    "Nhan Xet & Danh Gia AI", "Ke Hoach Thuc Thi"
)

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

# Header Row
$sb.AppendLine('    <row r="1">') | Out-Null
for ($c = 0; $c -lt $headers.Count; $c++) {
    $col = Get-ColLetter ($c + 1)
    $val = [System.Security.SecurityElement]::Escape($headers[$c])
    $sb.AppendLine("      <c r=`"${col}1`" t=`"inlineStr`" s=`"1`"><is><t>$val</t></is></c>") | Out-Null
}
$sb.AppendLine('    </row>') | Out-Null

# Data Rows
$rowNum = 2
foreach ($p in $profiles) {
    $usageStr = if ($null -eq $p.service_usage_rate_pct) { "NULL" } else { "$($p.service_usage_rate_pct)%" }
    $groundedStr = if ($p.is_revenue_grounded) { "DU CAN CU" } else { "TAM GIU" }
    $vals = @(
        $p.account_id,
        $p.customer_name,
        $p.representative,
        $p.industry,
        $p.segment,
        $p.current_package_name,
        $p.expiry_date,
        $p.days_to_expiry.ToString(),
        $usageStr,
        $p.open_tickets,
        $p.recommendation_type,
        $p.recommended_package_name,
        $p.base_price_vnd.ToString(),
        $p.discount_code,
        $p.discount_percent.ToString(),
        $p.final_price_vnd.ToString(),
        $groundedStr,
        $p.overall_status,
        $p.ai_remarks,
        $p.action_plan
    )

    $sb.AppendLine("    <row r=`"$rowNum`">") | Out-Null
    for ($c = 0; $c -lt $vals.Count; $c++) {
        $col = Get-ColLetter ($c + 1)
        $escVal = [System.Security.SecurityElement]::Escape([string]$vals[$c])
        $sb.AppendLine("      <c r=`"${col}${rowNum}`" t=`"inlineStr`"><is><t>$escVal</t></is></c>") | Out-Null
    }
    $sb.AppendLine('    </row>') | Out-Null
    $rowNum++
}

$sb.AppendLine('  </sheetData>') | Out-Null
$sb.AppendLine('</worksheet>') | Out-Null

[System.IO.File]::WriteAllText((Join-Path $worksheetsDir "sheet1.xml"), $sb.ToString(), $utf8NoBom)

# Zip
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $targetXlsx)
Remove-Item -Path $tempDir -Recurse -Force

Write-Output ("Successfully exported Cleaned Enterprise Excel: " + $targetXlsx + " (" + (Get-Item $targetXlsx).Length + " bytes)")
