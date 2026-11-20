# ==============================================================================
# SCRIPT TO BUILD A BRAND NEW 10-SLIDE POWERPOINT PRESENTATION (.PPTX) FROM SCRATCH
# Color Theme: Light Pastel Blue ("xanh dương nhạt")
# Modern, clean presentation design for Gia Linh Data Studio
# ==============================================================================

param(
    [string]$OutputPath = "c:\Users\ADMIN\Downloads\AI Agent CK\outputs\Thuyet_Trinh_He_Thong_Phan_Tich_Du_Lieu_Doanh_Nghiep.pptx"
)

Add-Type -AssemblyName "System.IO.Compression"
Add-Type -AssemblyName "System.IO.Compression.FileSystem"

$tempDir = Join-Path $env:TEMP "pptx_build_gialinh_$(Get-Random)"
if (Test-Path $tempDir) { Remove-Item -Recurse -Force $tempDir }
New-Item -ItemType Directory -Path $tempDir | Out-Null

$relsDir = New-Item -ItemType Directory -Path (Join-Path $tempDir "_rels")
$pptDir = New-Item -ItemType Directory -Path (Join-Path $tempDir "ppt")
$pptRelsDir = New-Item -ItemType Directory -Path (Join-Path $pptDir "_rels")
$slidesDir = New-Item -ItemType Directory -Path (Join-Path $pptDir "slides")
$slidesRelsDir = New-Item -ItemType Directory -Path (Join-Path $slidesDir "_rels")
$layoutsDir = New-Item -ItemType Directory -Path (Join-Path $pptDir "slideLayouts")
$layoutsRelsDir = New-Item -ItemType Directory -Path (Join-Path $layoutsDir "_rels")
$mastersDir = New-Item -ItemType Directory -Path (Join-Path $pptDir "slideMasters")
$mastersRelsDir = New-Item -ItemType Directory -Path (Join-Path $mastersDir "_rels")
$themeDir = New-Item -ItemType Directory -Path (Join-Path $pptDir "theme")

$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# 1. [Content_Types].xml
$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
'@

for ($i = 1; $i -le 10; $i++) {
    $contentTypes += "`r`n  <Override PartName=`"/ppt/slides/slide$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.presentationml.slide+xml`"/>"
}
$contentTypes += "`r`n</Types>"
[System.IO.File]::WriteAllText((Join-Path $tempDir "[Content_Types].xml"), $contentTypes, $utf8NoBom)

# 2. _rels/.rels
$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $relsDir ".rels"), $rootRels, $utf8NoBom)

# 3. ppt/presentation.xml
$presXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
'@
for ($i = 1; $i -le 10; $i++) {
    $sldId = 255 + $i
    $rId = "rId" + ($i + 2)
    $presXml += "`r`n    <p:sldId id=`"$sldId`" r:id=`"$rId`"/>"
}
$presXml += @'

  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
'@
[System.IO.File]::WriteAllText((Join-Path $pptDir "presentation.xml"), $presXml, $utf8NoBom)

# 4. ppt/_rels/presentation.xml.rels
$presRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
'@
for ($i = 1; $i -le 10; $i++) {
    $rId = "rId" + ($i + 2)
    $presRelsXml += "`r`n  <Relationship Id=`"$rId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide`" Target=`"slides/slide$i.xml`"/>"
}
$presRelsXml += "`r`n</Relationships>"
[System.IO.File]::WriteAllText((Join-Path $pptRelsDir "presentation.xml.rels"), $presRelsXml, $utf8NoBom)

# 5. ppt/theme/theme1.xml (LIGHT BLUE PALETTE: Sky Blue #0284C7, Light Blue #E0F2FE, Dark Navy #0F172A)
$themeXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="GiaLinhLightBlueTheme">
  <a:themeElements>
    <a:clrScheme name="LightBluePalette">
      <a:dk1><a:srgbClr val="0F172A"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="334155"/></a:dk2>
      <a:lt2><a:srgbClr val="F0F9FF"/></a:lt2>
      <a:accent1><a:srgbClr val="0284C7"/></a:accent1>
      <a:accent2><a:srgbClr val="38BDF8"/></a:accent2>
      <a:accent3><a:srgbClr val="0EA5E9"/></a:accent3>
      <a:accent4><a:srgbClr val="BAE6FD"/></a:accent4>
      <a:accent5><a:srgbClr val="6366F1"/></a:accent5>
      <a:accent6><a:srgbClr val="10B981"/></a:accent6>
      <a:hlink><a:srgbClr val="0284C7"/></a:hlink>
      <a:folHlink><a:srgbClr val="0369A1"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="ModernSans">
      <a:majorFont>
        <a:latin typeface="Segoe UI"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Segoe UI"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satOff val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satOff val="200000"/></a:schemeClr></a:gs></a:gsLst></a:gradFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:gs></a:gsLst></a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="s" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="25400" cap="flat" cmpd="s" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="38100" cap="flat" cmpd="s" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="lt2"/></a:solidFill>
        <a:solidFill><a:schemeClr val="lt1"/></a:solidFill>
        <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="40000"/></a:schemeClr></a:gs></a:gsLst></a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>
'@
[System.IO.File]::WriteAllText((Join-Path $themeDir "theme1.xml"), $themeXml, $utf8NoBom)

# 6. ppt/slideMasters/slideMaster1.xml
$slideMasterXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="F0F9FF"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>
'@
[System.IO.File]::WriteAllText((Join-Path $mastersDir "slideMaster1.xml"), $slideMasterXml, $utf8NoBom)

# 7. ppt/slideMasters/_rels/slideMaster1.xml.rels
$masterRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $mastersRelsDir "slideMaster1.xml.rels"), $masterRelsXml, $utf8NoBom)

# 8. ppt/slideLayouts/slideLayout1.xml
$slideLayoutXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="blank" preserve="1">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>
'@
[System.IO.File]::WriteAllText((Join-Path $layoutsDir "slideLayout1.xml"), $slideLayoutXml, $utf8NoBom)

# 9. ppt/slideLayouts/_rels/slideLayout1.xml.rels
$layoutRelsXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $layoutsRelsDir "slideLayout1.xml.rels"), $layoutRelsXml, $utf8NoBom)

# 10. Generate 10 Slide XMLs with Light Blue Theme
# Helper function to generate shape XML
function Build-ShapeXml {
    param(
        [int]$id,
        [string]$name,
        [long]$x,
        [long]$y,
        [long]$cx,
        [long]$cy,
        [string]$geom = "roundRect", # roundRect or rect
        [string]$fillClr = "FFFFFF",
        [string]$lineClr = "BAE6FD",
        [int]$lineW = 12700,
        [string]$paragraphsXml = "",
        [int]$radius = 3000
    )

    $lineTag = if ($lineClr -ne "") {
        "<a:ln w=`"$lineW`"><a:solidFill><a:srgbClr val=`"$lineClr`"/></a:solidFill></a:ln>"
    } else { "<a:ln><a:noFill/></a:ln>" }

    $fillTag = if ($fillClr -ne "") {
        "<a:solidFill><a:srgbClr val=`"$fillClr`"/></a:solidFill>"
    } else { "<a:noFill/>" }

    $geomTag = if ($geom -eq "roundRect") {
        "<a:prstGeom prst=`"roundRect`"><a:avLst><a:gd name=`"adj`" fmla=`"val $radius`"/></a:avLst></a:prstGeom>"
    } else {
        "<a:prstGeom prst=`"$geom`"><a:avLst/></a:prstGeom>"
    }

    return @"
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="$id" name="$name"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="$x" y="$y"/><a:ext cx="$cx" cy="$cy"/></a:xfrm>
          $geomTag
          $fillTag
          $lineTag
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="182880" tIns="182880" rIns="182880" bIns="182880" anchor="ctr"/>
          <a:lstStyle/>
          $paragraphsXml
        </p:txBody>
      </p:sp>
"@
}

function Build-ParaXml {
    param(
        [string]$text,
        [int]$size = 1400, # 14pt = 1400
        [bool]$bold = $false,
        [string]$color = "0F172A",
        [string]$align = "l", # l, r, ctr
        [int]$spaceAfter = 100000,
        [int]$spaceBefore = 0
    )
    $bAttr = if ($bold) { ' b="1"' } else { '' }
    $spcAfterXml = if ($spaceAfter -gt 0) { "<a:spcAft><a:spcPts val=`"$spaceAfter`"/></a:spcAft>" } else { "" }
    $spcBeforeXml = if ($spaceBefore -gt 0) { "<a:spcBef><a:spcPts val=`"$spaceBefore`"/></a:spcBef>" } else { "" }
    $safeText = [System.Security.SecurityElement]::Escape($text)

    return @"
          <a:p>
            <a:pPr algn="$align">$spcBeforeXml$spcAfterXml</a:pPr>
            <a:r>
              <a:rPr lang="vi-VN" sz="$size"$bAttr>
                <a:solidFill><a:srgbClr val="$color"/></a:solidFill>
                <a:latin typeface="Segoe UI"/>
                <a:ea typeface="Segoe UI"/>
              </a:rPr>
              <a:t>$safeText</a:t>
            </a:r>
          </a:p>
"@
}

# SLIDES CONTENT DEFINITION (10 HIGH-IMPACT SLIDES)
$slidesData = @(
    # SLIDE 1: COVER
    @{
        HeaderBadge = "GIA LINH DATA STUDIO • PHIÊN BẢN 4.0"
        Title = "GIA LINH DATA STUDIO"
        Subtitle = "Nền Tảng Phân Tích & Đối Soát Dữ Liệu Khách Hàng Doanh Nghiệp"
        Cards = @(
            @{
                x = 800000; y = 2800000; cx = 10592000; cy = 3400000; fill = "FFFFFF"; line = "BAE6FD"; radius = 4000;
                Paras = @(
                    @{ text = "TỰ ĐỘNG HÓA QUY TRÌNH • KIỂM SOÁT AN TOÀN DOANH THU • BẢO MẬT ĐA TẦNG"; sz = 1300; b = $true; clr = "0284C7"; align = "ctr"; spAfter = 140000 },
                    @{ text = "Hệ thống chuyên sâu giải quyết triệt để 3 bài toán lớn: dữ liệu phân mảnh đa nguồn, sai lệch dữ liệu đo lường công suất (NULL ≠ 0), và xung đột quy trình chào hợp đồng khi khách hàng đang gặp sự cố kỹ thuật."; sz = 1500; b = $false; clr = "334155"; align = "ctr"; spAfter = 200000 },
                    @{ text = "✦ Phân Biệt NULL ≠ 0    ✦ Khóa Chặn Kỹ Thuật Hard-Block    ✦ Cổng Phê Duyệt Cấp Quản Lý HITL"; sz = 1400; b = $true; clr = "0369A1"; align = "ctr"; spAfter = 240000 },
                    @{ text = "Đơn vị xây dựng: Gia Linh Data Studio  |  Môi trường: http://localhost:8080  |  Năm 2026 - 2027"; sz = 1200; b = $false; clr = "64748B"; align = "ctr" }
                )
            }
        )
    },
    # SLIDE 2: BỐI CẢNH & THÁCH THỨC
    @{
        HeaderBadge = "BỐI CẢNH & THÁCH THỨC QUẢN TRỊ DỮ LIỆU"
        Title = "3 Thách Thức Lớn Trong Quản Trị Khách Hàng B2B"
        Subtitle = "Thực trạng dữ liệu phân mảnh và rủi ro thất thoát doanh thu trước khi có Gia Linh Data Studio"
        Cards = @(
            @{
                x = 800000; y = 2200000; cx = 3300000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "THÁCH THỨC 01"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 80000 },
                    @{ text = "Dữ Liệu Phân Mảnh"; sz = 1800; b = $true; clr = "0F172A"; align = "l"; spAfter = 140000 },
                    @{ text = "Hợp đồng nằm tại CRM, đo lường công suất nằm tại máy chủ kỹ thuật, lịch sử thanh toán lưu trữ riêng biệt."; sz = 1300; b = $false; clr = "475569"; align = "l"; spAfter = 140000 },
                    @{ text = "Hệ quả: Mất hơn 48 giờ đối soát thủ công liên phòng ban cho mỗi chu kỳ tái ký."; sz = 1200; b = $true; clr = "DC2626"; align = "l" }
                )
            },
            @{
                x = 4446000; y = 2200000; cx = 3300000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "THÁCH THỨC 02"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 80000 },
                    @{ text = "Nhầm Lẫn Dữ Liệu NULL"; sz = 1800; b = $true; clr = "0F172A"; align = "l"; spAfter = 140000 },
                    @{ text = "Khi cổng API lỗi, trường dữ liệu đo lường trả về NULL. Hệ thống cũ thường tự động gán về 0%."; sz = 1300; b = $false; clr = "475569"; align = "l"; spAfter = 140000 },
                    @{ text = "Hệ quả: Phán đoán sai nhu cầu thực, bỏ sót khách hàng tiềm năng hoặc chào giá sai lệch."; sz = 1200; b = $true; clr = "DC2626"; align = "l" }
                )
            },
            @{
                x = 8092000; y = 2200000; cx = 3300000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "THÁCH THỨC 03"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 80000 },
                    @{ text = "Chào Tái Ký Khi Lỗi Hạ Tầng"; sz = 1800; b = $true; clr = "0F172A"; align = "l"; spAfter = 140000 },
                    @{ text = "Kinh doanh không nắm bắt sự cố kỹ thuật mở, tiếp tục gửi báo giá nâng cấp cho khách hàng đang bức xúc."; sz = 1300; b = $false; clr = "475569"; align = "l"; spAfter = 140000 },
                    @{ text = "Hệ quả: Gây khủng hoảng truyền thông, mất uy tín thương hiệu và nguy cơ hủy dịch vụ."; sz = 1200; b = $true; clr = "DC2626"; align = "l" }
                )
            }
        )
    },
    # SLIDE 3: 4 TRỤ CỘT CHIẾN LƯỢC
    @{
        HeaderBadge = "KIẾN TRÚC & NGUYÊN TẮC VẬN HÀNH"
        Title = "4 Trụ Cột Chiến Lược Của Gia Linh Data Studio"
        Subtitle = "Khung tiêu chuẩn bảo vệ an toàn dữ liệu và tối ưu hiệu suất kinh doanh"
        Cards = @(
            @{
                x = 800000; y = 2200000; cx = 2500000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "TRỤ CỘT 1"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 60000 },
                    @{ text = "Hợp Nhất Đa Nguồn"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Đồng bộ tự động 6 bảng dữ liệu: hợp đồng, khách hàng, biểu phí, chiết khấu, đo lường và sự cố."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "Khóa mốc đối chiếu chung (as_of_date) bảo đảm tính nhất quán số liệu."; sz = 1100; b = $true; clr = "0284C7"; align = "l" }
                )
            },
            @{
                x = 3500000; y = 2200000; cx = 2500000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "TRỤ CỘT 2"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 60000 },
                    @{ text = "Tách Biệt NULL ≠ 0"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Phân định rạch ròi giữa khuyết dữ liệu do lỗi mạng với khách hàng chưa phát sinh lưu lượng thực tế."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "Chỉ ghi nhận doanh thu khi có đủ căn cứ đo lường chuẩn xác."; sz = 1100; b = $true; clr = "0284C7"; align = "l" }
                )
            },
            @{
                x = 6200000; y = 2200000; cx = 2500000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "TRỤ CỘT 3"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 60000 },
                    @{ text = "Khóa Chặn Kỹ Thuật"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Cơ chế Hard-Block tự động khóa tài khoản khi có Ticket sự cố máy chủ mức độ cao chưa giải quyết."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "Cấm bộ phận Kinh doanh tiếp cận chào hợp đồng trước khi sự cố được khắc phục."; sz = 1100; b = $true; clr = "0284C7"; align = "l" }
                )
            },
            @{
                x = 8900000; y = 2200000; cx = 2500000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "TRỤ CỘT 4"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 60000 },
                    @{ text = "Cổng Phê Duyệt HITL"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Tự động hóa 95% trường hợp hợp lệ; chuyển các ca chiết khấu vượt khung lên Giám đốc phê duyệt."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "Kết hợp sức mạnh AI và quyền kiểm soát tối thượng của con người."; sz = 1100; b = $true; clr = "0284C7"; align = "l" }
                )
            }
        )
    },
    # SLIDE 4: QUY TRÌNH 4 BƯỚC PHÂN TÍCH
    @{
        HeaderBadge = "QUY TRÌNH VẬN HÀNH DỮ LIỆU"
        Title = "Quy Trình 4 Bước Phân Tích & Đối Soát Tự Động"
        Subtitle = "Từ kết nối nguồn thô đến bảng điều khiển tính toán và xuất báo cáo đa định dạng"
        Cards = @(
            @{
                x = 800000; y = 2200000; cx = 2500000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "BƯỚC 01"; sz = 1200; b = $true; clr = "0284C7"; align = "l"; spAfter = 60000 },
                    @{ text = "Nạp & Kết Nối Nguồn"; sz = 1500; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Tiếp nhận tệp Excel/JSON, kết nối API dữ liệu đo lường. Khóa mốc đối chiếu chung ngày 18/11/2026."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "➔ Kiểm tra tính toàn vẹn tệp đầu vào."; sz = 1100; b = $true; clr = "0284C7"; align = "l" }
                )
            },
            @{
                x = 3500000; y = 2200000; cx = 2500000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "BƯỚC 02"; sz = 1200; b = $true; clr = "0284C7"; align = "l"; spAfter = 60000 },
                    @{ text = "Quét & Chẩn Đoán Lỗi"; sz = 1500; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Phát hiện lỗi định dạng ngày (32/13/2026), khuyết thiếu NULL và sự cố máy chủ mở."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "➔ Bật bảng Where - Why - Fix Guide."; sz = 1100; b = $true; clr = "0284C7"; align = "l" }
                )
            },
            @{
                x = 6200000; y = 2200000; cx = 2500000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "BƯỚC 03"; sz = 1200; b = $true; clr = "0284C7"; align = "l"; spAfter = 60000 },
                    @{ text = "Làm Sạch & Phân Khúc"; sz = 1500; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Áp dụng ma trận logic đề xuất gói, tính toán chiết khấu hợp lệ, gắn cờ khóa chặn kỹ thuật an toàn."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "➔ Phân loại VIP, SME, Tạm giữ."; sz = 1100; b = $true; clr = "0284C7"; align = "l" }
                )
            },
            @{
                x = 8900000; y = 2200000; cx = 2500000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "BƯỚC 04"; sz = 1200; b = $true; clr = "0284C7"; align = "l"; spAfter = 60000 },
                    @{ text = "Xác Nhận Doanh Thu"; sz = 1500; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Hiển thị Dashboard trực quan, kích hoạt trợ lý AI đối soát và xuất file Excel, Word, PowerPoint."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "➔ Bảo vệ doanh thu có căn cứ vững chắc."; sz = 1100; b = $true; clr = "0284C7"; align = "l" }
                )
            }
        )
    },
    # SLIDE 5: CHUYÊN ĐỀ TÁCH BIỆT NULL VS 0%
    @{
        HeaderBadge = "CHUYÊN ĐỀ KỸ THUẬT DỮ LIỆU ĐẶC BIỆT"
        Title = "Phân Định Rạch Ròi: Khuyết Dữ Liệu (NULL) vs Mức Dùng 0%"
        Subtitle = "So sánh thực tiễn 2 hồ sơ doanh nghiệp tiêu biểu DN-2026-004 và DN-2026-005"
        Cards = @(
            @{
                x = 800000; y = 2200000; cx = 5150000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "HỒ SƠ DN-2026-004 (TIẾP VẬN LOGITRANS)"; sz = 1200; b = $true; clr = "DC2626"; align = "l"; spAfter = 80000 },
                    @{ text = "Trạng Thái: Khuyết Dữ Liệu Đo Lường (NULL)"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "• Nguyên nhân: Cổng API đo lường công suất máy chủ bị lỗi đồng bộ (Ticket TK-8004). Dữ liệu hoàn toàn trống rỗng."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 100000 },
                    @{ text = "• Sai lầm thường gặp: Gán bừa NULL = 0% dẫn đến kết luận sai lệch rằng khách hàng không sử dụng hạ tầng."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 100000 },
                    @{ text = "• Quy tắc Gia Linh Data Studio: Tạm giữ 50.000.000 VNĐ, không tính vào doanh thu đủ căn cứ. Cấm chào tái ký trước khi Kỹ thuật hoàn tất cấu hình API."; sz = 1200; b = $true; clr = "0369A1"; align = "l" }
                )
            },
            @{
                x = 6242000; y = 2200000; cx = 5150000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "HỒ SƠ DN-2026-005 (KHÁCH SẠN LOTUS)"; sz = 1200; b = $true; clr = "16A34A"; align = "l"; spAfter = 80000 },
                    @{ text = "Trạng Thái: Mức Dùng 0% Đạt Chuẩn Hợp Lệ"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "• Nguyên nhân: Hệ thống đo lường hoạt động bình thường, khách hàng tạm hoãn đưa vào sử dụng do doanh nghiệp đang tái cấu trúc."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 100000 },
                    @{ text = "• Phân tích dữ liệu: Con số 0% là dữ liệu thực tế chuẩn xác, không phải do lỗi hệ thống đo lường."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 100000 },
                    @{ text = "• Quy tắc Gia Linh Data Studio: Xác nhận đủ căn cứ đề xuất gói duy trì 31.500.000 VNĐ, phân công chuyên viên CSKH đồng hành hỗ trợ kích hoạt."; sz = 1200; b = $true; clr = "0369A1"; align = "l" }
                )
            }
        )
    },
    # SLIDE 6: KHÓA CHẶN AN TOÀN HARD-BLOCK
    @{
        HeaderBadge = "CƠ CHẾ BẢO VỆ THƯƠNG HIỆU & KHÁCH HÀNG"
        Title = "Cơ Chế Khóa Chặn An Toàn Kiểm Soát Lỗi Kỹ Thuật (Hard-Block)"
        Subtitle = "Bảo đảm uy tín doanh nghiệp và ngăn chặn xung đột giữa Kỹ thuật và Kinh doanh"
        Cards = @(
            @{
                x = 800000; y = 2200000; cx = 10592000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "TRƯỜNG HỢP ĐIỂN HÌNH: DOANH NGHIỆP DN-2026-007 (ĐỊA ỐC PHƯƠNG NAM)"; sz = 1400; b = $true; clr = "0284C7"; align = "l"; spAfter = 100000 },
                    @{ text = "Tình Huống Đối Soát:"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 100000 },
                    @{ text = "• Khách hàng có mức sử dụng hạ tầng rất cao (94%), đủ tiêu chuẩn nâng cấp lên gói Enterprise Suite trị giá 220.000.000 VNĐ."; sz = 1300; b = $false; clr = "334155"; align = "l"; spAfter = 80000 },
                    @{ text = "• Tuy nhiên, bộ phận Kỹ thuật đang ghi nhận sự cố máy chủ CSDL cấp độ nghiêm trọng (Ticket TK-8002 đang ở trạng thái Mở)."; sz = 1300; b = $false; clr = "334155"; align = "l"; spAfter = 120000 },
                    @{ text = "Hành Động Của Gia Linh Data Studio:"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 100000 },
                    @{ text = "1. Kích hoạt chốt chặn an toàn Hard-Block, tạm đình chỉ báo giá nâng cấp."; sz = 1300; b = $true; clr = "DC2626"; align = "l"; spAfter = 60000 },
                    @{ text = "2. Tạm giữ toàn bộ giá trị 220.000.000 VNĐ, không đưa vào danh mục doanh thu đủ căn cứ."; sz = 1300; b = $true; clr = "0369A1"; align = "l"; spAfter = 60000 },
                    @{ text = "3. Phát thông báo khẩn yêu cầu Kỹ thuật xử lý xong Ticket trong 12h trước khi Kinh doanh liên hệ khách hàng."; sz = 1300; b = $true; clr = "16A34A"; align = "l" }
                )
            }
        )
    },
    # SLIDE 7: QUẢN TRỊ CHIẾT KHẤU & CỔNG HITL
    @{
        HeaderBadge = "CỔNG PHÊ DUYỆT CẤP QUẢN LÝ (HITL GATE)"
        Title = "Quản Trị Chiết Khấu & Cổng Phê Duyệt Cấp Quản Lý"
        Subtitle = "Cân bằng giữa tự động hóa linh hoạt và quyền kiểm soát tài chính tối cao"
        Cards = @(
            @{
                x = 800000; y = 2200000; cx = 3300000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "NGOẠI LỆ CHIẾT KHẤU"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 80000 },
                    @{ text = "Hồ Sơ DN-2026-003"; sz = 1700; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Dược phẩm Medilife yêu cầu chiết khấu 20%, vượt khung thẩm quyền tự động thông thường (tối đa 15%)."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "➔ Hệ thống chuyển sang trạng thái 'CHỜ QUẢN LÝ PHÊ DUYỆT', yêu cầu Giám đốc Kinh doanh ký duyệt."; sz = 1200; b = $true; clr = "7C3AED"; align = "l" }
                )
            },
            @{
                x = 4446000; y = 2200000; cx = 3300000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "RANH GIỚI NÂNG CẤP"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 80000 },
                    @{ text = "Hồ Sơ DN-2026-008"; sz = 1700; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Cơ khí Sakura đạt mức sử dụng 79.2% (sát mốc 80%), yêu cầu mã VIP nhưng mới ký hợp đồng 1 năm."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "➔ Hệ thống từ chối mã VIP, tự động áp dụng mã Tái ký sớm 10% hợp lệ, giữ nguyên giá trị tối ưu."; sz = 1200; b = $true; clr = "0284C7"; align = "l" }
                )
            },
            @{
                x = 8092000; y = 2200000; cx = 3300000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "TÁI KÝ KHẨN CẤP"; sz = 1100; b = $true; clr = "0284C7"; align = "l"; spAfter = 80000 },
                    @{ text = "Hồ Sơ DN-2026-010"; sz = 1700; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "Sáng Tạo Việt chỉ còn 3 ngày hợp đồng, mức sử dụng thấp 35%, yêu cầu chiết khấu lớn."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "➔ Hệ thống từ chối chiết khấu sai quy tắc, hướng dẫn tái ký nguyên giá khẩn cấp để không gián đoạn dịch vụ."; sz = 1200; b = $true; clr = "DC2626"; align = "l" }
                )
            }
        )
    },
    # SLIDE 8: KIẾN TRÚC DASHBOARD
    @{
        HeaderBadge = "GIAO DIỆN & TRẢI NGHIỆM NGƯỜI DÙNG"
        Title = "Kiến Trúc Bảng Điều Khiển Trực Quan (Dashboard)"
        Subtitle = "Bố cục 2 cột hiện đại, không gian làm việc rộng rãi, khởi tạo 0-Data tinh khôi"
        Cards = @(
            @{
                x = 800000; y = 2200000; cx = 5150000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "THANH CÔNG CỤ DỌC (CỘT TRÁI)"; sz = 1300; b = $true; clr = "0284C7"; align = "l"; spAfter = 80000 },
                    @{ text = "Trung Tâm Điều Hành & Trợ Lý AI"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "• Điều hướng linh hoạt: Tổng quan, Bảng dữ liệu sạch, Chẩn đoán lỗi (0 lỗi ban đầu)."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 80000 },
                    @{ text = "• Bộ lọc phân khúc 4 nhóm: Tất cả, Khách VIP, Khối SME, Tạm giữ an toàn."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 80000 },
                    @{ text = "• Trợ lý AI chat trực tiếp: Tra cứu căn cứ từng mã DN, phân tích NULL vs 0% bằng 100% tiếng Việt."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 80000 },
                    @{ text = "• Thẻ gợi ý thông minh (Smart Insight) tự động cập nhật cảnh báo."; sz = 1200; b = $true; clr = "0369A1"; align = "l" }
                )
            },
            @{
                x = 6242000; y = 2200000; cx = 5150000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "KHÔNG GIAN LÀM VIỆC CHÍNH (FULL WIDTH)"; sz = 1300; b = $true; clr = "0284C7"; align = "l"; spAfter = 80000 },
                    @{ text = "Trực Quan Hóa Số Liệu Đa Chiều"; sz = 1600; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "• Lời chào tương tác: 'Xin chào, khách hàng 👋' cùng thông báo động khi nhấp chuột."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 80000 },
                    @{ text = "• 4 Thẻ chỉ số Gradient: Tổng khách hàng, Doanh thu đủ căn cứ, Mức dùng TB, Doanh thu tạm giữ."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 80000 },
                    @{ text = "• Biểu đồ đường kép (Dual Spline): Thể hiện biến động doanh thu duy trì và nâng cấp."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 80000 },
                    @{ text = "• Biểu đồ tròn (Donut): Phân bổ cơ cấu danh mục gói dịch vụ doanh nghiệp."; sz = 1200; b = $true; clr = "0369A1"; align = "l" }
                )
            }
        )
    },
    # SLIDE 9: CHẨN ĐOÁN LỖI TỰ ĐỘNG
    @{
        HeaderBadge = "CHẨN ĐOÁN LỖI DỮ LIỆU ĐẦU VÀO"
        Title = "Cơ Chế Chẩn Đoán Lỗi 0-Lỗi Tự Động (Where - Why - Fix)"
        Subtitle = "Chỉ khi dữ liệu chuẩn xác mới xác nhận doanh thu, bảo vệ tuyệt đối tính kỷ luật số liệu"
        Cards = @(
            @{
                x = 800000; y = 2200000; cx = 10592000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "MA TRẬN CHẨN ĐOÁN VÀ KHẮC PHỤC LỖI TỰ ĐỘNG"; sz = 1400; b = $true; clr = "0284C7"; align = "l"; spAfter = 100000 },
                    @{ text = "1. Lỗi Định Dạng Ngày Tháng (Dòng 5 | Cột: Ngày hết hạn):"; sz = 1400; b = $true; clr = "0F172A"; align = "l"; spAfter = 40000 },
                    @{ text = "   - Giá trị sai: '32/13/2026' từ hệ thống CRM cũ ➔ Giải pháp: Chuẩn hóa tự động về chuẩn ISO YYYY-MM-DD (2026-11-28)."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 80000 },
                    @{ text = "2. Lỗi Khuyết Dữ Liệu Đo Lường (Dòng 7 | Cột: Mức sử dụng dịch vụ):"; sz = 1400; b = $true; clr = "0F172A"; align = "l"; spAfter = 40000 },
                    @{ text = "   - Giá trị sai: 'NULL / Để trống' do lỗi cổng API ➔ Giải pháp: Tạm giữ doanh thu, tuyệt đối không gán bằng 0; chờ đồng bộ API."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 80000 },
                    @{ text = "3. Lỗi Sự Cố Kỹ Thuật Máy Chủ (Dòng 10 | Bảng: Sự cố kỹ thuật):"; sz = 1400; b = $true; clr = "0F172A"; align = "l"; spAfter = 40000 },
                    @{ text = "   - Giá trị sai: Sự cố Mở (Ticket TK-8002) ➔ Giải pháp: Kích hoạt chốt chặn an toàn cấm chào tái ký; ưu tiên Kỹ thuật xử lý."; sz = 1200; b = $false; clr = "475569"; align = "l"; spAfter = 120000 },
                    @{ text = "✦ Nút [🔧 Tự Động Chuẩn Hóa & Bắt Đầu Phân Tích] cho phép xử lý và đưa hệ thống về 0 Lỗi chỉ trong 1 thao tác."; sz = 1300; b = $true; clr = "16A34A"; align = "l" }
                )
            }
        )
    },
    # SLIDE 10: TỔNG KẾT & LIÊN HỆ
    @{
        HeaderBadge = "TỔNG KẾT & ĐỊNH HƯỚNG PHÁT TRIỂN"
        Title = "Gia Linh Data Studio — Nâng Tầm Quản Trị Dữ Liệu Doanh Nghiệp"
        Subtitle = "Sẵn sàng vận hành thực tế và mở rộng quy mô toàn diện"
        Cards = @(
            @{
                x = 800000; y = 2200000; cx = 5150000; cy = 4000000; fill = "FFFFFF"; line = "BAE6FD"; radius = 3000;
                Paras = @(
                    @{ text = "GIÁ TRỊ CỐT LÕI MANG LẠI"; sz = 1400; b = $true; clr = "0284C7"; align = "l"; spAfter = 100000 },
                    @{ text = "• Tiết kiệm hơn 90% thời gian đối soát dữ liệu thủ công."; sz = 1300; b = $false; clr = "334155"; align = "l"; spAfter = 80000 },
                    @{ text = "• Triệt tiêu hoàn toàn rủi ro thất thoát doanh thu từ dữ liệu khuyết thiếu."; sz = 1300; b = $false; clr = "334155"; align = "l"; spAfter = 80000 },
                    @{ text = "• Bảo vệ uy tín thương hiệu nhờ cơ chế khóa chặn sự cố tự động."; sz = 1300; b = $false; clr = "334155"; align = "l"; spAfter = 80000 },
                    @{ text = "• Hỗ trợ lãnh đạo doanh nghiệp ra quyết định chiến lược tức thì dựa trên dữ liệu đủ căn cứ."; sz = 1300; b = $false; clr = "334155"; align = "l"; spAfter = 120000 },
                    @{ text = "Tỷ lệ đối soát chuẩn hóa: 100%  |  Thời gian chẩn đoán: < 2 giây"; sz = 1200; b = $true; clr = "16A34A"; align = "l" }
                )
            },
            @{
                x = 6242000; y = 2200000; cx = 5150000; cy = 4000000; fill = "E0F2FE"; line = "38BDF8"; radius = 3000;
                Paras = @(
                    @{ text = "THÔNG TIN HỆ THỐNG & LIÊN HỆ"; sz = 1400; b = $true; clr = "0369A1"; align = "l"; spAfter = 100000 },
                    @{ text = "GIA LINH DATA STUDIO"; sz = 2000; b = $true; clr = "0F172A"; align = "l"; spAfter = 120000 },
                    @{ text = "• Địa chỉ máy chủ: http://localhost:8080"; sz = 1300; b = $true; clr = "0284C7"; align = "l"; spAfter = 80000 },
                    @{ text = "• Đơn vị phát triển: Ban Công Nghệ & Vận Hành Dữ Liệu"; sz = 1200; b = $false; clr = "334155"; align = "l"; spAfter = 80000 },
                    @{ text = "• Email hỗ trợ: contact@gialinh-datastudio.vn"; sz = 1200; b = $false; clr = "334155"; align = "l"; spAfter = 80000 },
                    @{ text = "• Tài liệu đính kèm: Báo cáo Word (.docx) & Dữ liệu Excel (.xlsx)"; sz = 1200; b = $false; clr = "334155"; align = "l"; spAfter = 140000 },
                    @{ text = "XIN TRÂN TRỌNG CẢM ƠN!"; sz = 1600; b = $true; clr = "0F172A"; align = "ctr" }
                )
            }
        )
    }
)

# 11. BUILD EACH SLIDE XML
for ($i = 0; $i -lt $slidesData.Count; $i++) {
    $slideNum = $i + 1
    $data = $slidesData[$i]

    $spTreeXml = ""
    $shapeId = 2

    # A. TOP HEADER BADGE (XANH DƯƠNG NHẠT PILL)
    $badgeText = $data.HeaderBadge
    $badgeParas = Build-ParaXml -text $badgeText -size 1000 -bold $true -color "0284C7" -align "l" -spaceAfter 0
    $badgeShape = Build-ShapeXml -id $shapeId -name "HeaderBadge" -x 800000 -y 500000 -cx 4500000 -cy 350000 -geom "roundRect" -fillClr "E0F2FE" -lineClr "BAE6FD" -lineW 12700 -paragraphsXml $badgeParas -radius 5000
    $spTreeXml += $badgeShape + "`r`n"
    $shapeId++

    # B. SLIDE NUMBER BADGE (RIGHT TOP PILL)
    $numText = "$slideNum / 10"
    $numParas = Build-ParaXml -text $numText -size 1100 -bold $true -color "0369A1" -align "ctr" -spaceAfter 0
    $numShape = Build-ShapeXml -id $shapeId -name "SlideNumBadge" -x 10592000 -y 500000 -cx 800000 -cy 350000 -geom "roundRect" -fillClr "E0F2FE" -lineClr "BAE6FD" -lineW 12700 -paragraphsXml $numParas -radius 5000
    $spTreeXml += $numShape + "`r`n"
    $shapeId++

    # C. TITLE & SUBTITLE
    $titleParas = (Build-ParaXml -text $data.Title -size 2400 -bold $true -color "0F172A" -align "l" -spaceAfter 40000) + `
                  (Build-ParaXml -text $data.Subtitle -size 1200 -bold $false -color "64748B" -align "l" -spaceAfter 0)
    $titleShape = Build-ShapeXml -id $shapeId -name "SlideTitleBox" -x 800000 -y 950000 -cx 10592000 -cy 1050000 -geom "rect" -fillClr "" -lineClr "" -lineWidth 0 -paragraphsXml $titleParas
    $spTreeXml += $titleShape + "`r`n"
    $shapeId++

    # D. CONTENT CARDS (THIẾT KẾ CARD HIỆN ĐẠI MÀU XANH DƯƠNG NHẠT)
    foreach ($card in $data.Cards) {
        $cParasXml = ""
        foreach ($p in $card.Paras) {
            $spAft = if ($p.ContainsKey('spAfter')) { $p.spAfter } else { 60000 }
            $cParasXml += Build-ParaXml -text $p.text -size $p.sz -bold $p.b -color $p.clr -align $p.align -spaceAfter $spAft
        }
        $cShape = Build-ShapeXml -id $shapeId -name "ContentCard" -x $card.x -y $card.y -cx $card.cx -cy $card.cy -geom "roundRect" -fillClr $card.fill -lineClr $card.line -lineW 12700 -paragraphsXml $cParasXml -radius $card.radius
        $spTreeXml += $cShape + "`r`n"
        $shapeId++
    }

    # ASSEMBLE FULL SLIDE XML
    $slideXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="F0F9FF"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
$spTreeXml
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"@
    [System.IO.File]::WriteAllText((Join-Path $slidesDir "slide$slideNum.xml"), $slideXml, $utf8NoBom)

    # Relationship for each slide -> layout1
    $slideRelXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>
"@
    [System.IO.File]::WriteAllText((Join-Path $slidesRelsDir "slide$slideNum.xml.rels"), $slideRelXml, $utf8NoBom)
}

# 12. ZIP ENTIRE DIRECTORY INTO OUTPUT PPTX PRESERVING FORWARD SLASHES
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

$zip = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Create)
$allFiles = Get-ChildItem -Path $tempDir -Recurse -File
foreach ($file in $allFiles) {
    $relPath = $file.FullName.Substring($tempDir.Length + 1).Replace("\", "/")
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $relPath, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
}
$zip.Dispose()

# Also copy to root workspace
$rootCopy = "c:\Users\ADMIN\Downloads\AI Agent CK\Thuyet_Trinh_He_Thong_Phan_Tich_Du_Lieu_Doanh_Nghiep.pptx"
Copy-Item -Path $OutputPath -Destination $rootCopy -Force

Remove-Item -Recurse -Force $tempDir

$fileSize = (Get-Item $OutputPath).Length
Write-Output "Successfully created brand new Light Blue PowerPoint deck: $OutputPath ($fileSize bytes)"
