# ==============================================================================
# SCRIPT TO BUILD EXECUTIVE WORD REPORT (.DOCX) USING OPENXML
# Target: outputs/Bao_Cao_Tong_Quan_He_Thong_Phan_Tich_Du_Lieu.docx
# ==============================================================================

Add-Type -AssemblyName "System.IO.Compression"
Add-Type -AssemblyName "System.IO.Compression.FileSystem"

$targetDocx = "c:\Users\ADMIN\Downloads\AI Agent CK\outputs\Bao_Cao_Tong_Quan_He_Thong_Phan_Tich_Du_Lieu.docx"
$rootDocx = "c:\Users\ADMIN\Downloads\AI Agent CK\Bao_Cao_Tong_Quan_He_Thong_Phan_Tich_Du_Lieu.docx"

if (Test-Path $targetDocx) { Remove-Item $targetDocx -Force }
if (Test-Path $rootDocx) { Remove-Item $rootDocx -Force }

$tempDir = Join-Path $env:TEMP ("docx_ent_" + (Get-Random))
New-Item -ItemType Directory -Path $tempDir | Out-Null
$wordDir = Join-Path $tempDir "word"
$relsDir = Join-Path $tempDir "_rels"
$wordRelsDir = Join-Path $wordDir "_rels"

New-Item -ItemType Directory -Path $wordDir -Force | Out-Null
New-Item -ItemType Directory -Path $relsDir -Force | Out-Null
New-Item -ItemType Directory -Path $wordRelsDir -Force | Out-Null

# 1. [Content_Types].xml
$contentTypes = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
'@
[System.IO.File]::WriteAllText((Join-Path $tempDir "[Content_Types].xml"), $contentTypes.Trim(), [System.Text.Encoding]::UTF8)

# 2. _rels/.rels
$rootRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $relsDir ".rels"), $rootRels.Trim(), [System.Text.Encoding]::UTF8)

# 3. word/_rels/document.xml.rels
$docRels = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
'@
[System.IO.File]::WriteAllText((Join-Path $wordRelsDir "document.xml.rels"), $docRels.Trim(), [System.Text.Encoding]::UTF8)

# 4. word/styles.xml
$styles = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="23"/>
        <w:color w:val="1E1B4B"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="276" w:lineRule="auto" w:after="140"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>
'@
[System.IO.File]::WriteAllText((Join-Path $wordDir "styles.xml"), $styles.Trim(), [System.Text.Encoding]::UTF8)

# Helper function to generate WordprocessingML XML elements
function W-Para([string]$text, [string]$color = "1E1B4B", [int]$sz = 23, [bool]$bold = $false, [string]$align = "left", [int]$spaceBefore = 0, [int]$spaceAfter = 140) {
    $bTag = if ($bold) { "<w:b/>" } else { "" }
    $jcTag = if ($align -ne "left") { "<w:jc w:val=`"$align`"/>" } else { "" }
    $spTag = "<w:spacing w:before=`"$spaceBefore`" w:after=`"$spaceAfter`" w:line=`"280`" w:lineRule=`"auto`"/>"
    $escaped = [System.Security.SecurityElement]::Escape($text)
    return @"
    <w:p>
      <w:pPr>
        $jcTag
        $spTag
      </w:pPr>
      <w:r>
        <w:rPr>
          $bTag
          <w:sz w:val="$sz"/>
          <w:color w:val="$color"/>
        </w:rPr>
        <w:t xml:space="preserve">$escaped</w:t>
      </w:r>
    </w:p>
"@
}

function W-Heading1([string]$text) {
    return W-Para -text $text -color "4F46E5" -sz 34 -bold $true -spaceBefore 280 -spaceAfter 120
}

function W-Heading2([string]$text) {
    return W-Para -text $text -color "3730A3" -sz 28 -bold $true -spaceBefore 200 -spaceAfter 100
}

function W-Bullet([string]$title, [string]$content) {
    $escTitle = [System.Security.SecurityElement]::Escape($title)
    $escCont = [System.Security.SecurityElement]::Escape($content)
    return @"
    <w:p>
      <w:pPr>
        <w:ind w:left="360" w:hanging="240"/>
        <w:spacing w:before="60" w:after="80" w:line="260" w:lineRule="auto"/>
      </w:pPr>
      <w:r>
        <w:rPr><w:b/><w:color w:val="4F46E5"/><w:sz w:val="23"/></w:rPr>
        <w:t xml:space="preserve">• $($escTitle): </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:color w:val="334155"/><w:sz w:val="23"/></w:rPr>
        <w:t xml:space="preserve">$escCont</w:t>
      </w:r>
    </w:p>
"@
}

function W-Table2Col([array]$rows) {
    $sbTbl = New-Object System.Text.StringBuilder
    $sbTbl.AppendLine('    <w:tbl>') | Out-Null
    $sbTbl.AppendLine('      <w:tblPr>') | Out-Null
    $sbTbl.AppendLine('        <w:tblW w:w="9400" w:type="dxa"/>') | Out-Null
    $sbTbl.AppendLine('        <w:tblBorders>') | Out-Null
    $sbTbl.AppendLine('          <w:top w:val="single" w:sz="6" w:space="0" w:color="C7D2FE"/>') | Out-Null
    $sbTbl.AppendLine('          <w:bottom w:val="single" w:sz="6" w:space="0" w:color="C7D2FE"/>') | Out-Null
    $sbTbl.AppendLine('          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E0E7FF"/>') | Out-Null
    $sbTbl.AppendLine('          <w:insideV w:val="none"/>') | Out-Null
    $sbTbl.AppendLine('          <w:left w:val="none"/><w:right w:val="none"/>') | Out-Null
    $sbTbl.AppendLine('        </w:tblBorders>') | Out-Null
    $sbTbl.AppendLine('      </w:tblPr>') | Out-Null

    # Header Row
    $sbTbl.AppendLine('      <w:tr>') | Out-Null
    $sbTbl.AppendLine('        <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EEF2FF"/></w:tcPr><w:p><w:pPr><w:spacing w:before="80" w:after="80"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="312E81"/></w:rPr><w:t>Tiêu Chí / Thành Phần</w:t></w:r></w:p></w:tc>') | Out-Null
    $sbTbl.AppendLine('        <w:tc><w:tcPr><w:tcW w:w="6400" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="EEF2FF"/></w:tcPr><w:p><w:pPr><w:spacing w:before="80" w:after="80"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="312E81"/></w:rPr><w:t>Nội Dung Chi Tiết &amp; Ý Nghĩa Vận Hành</w:t></w:r></w:p></w:tc>') | Out-Null
    $sbTbl.AppendLine('      </w:tr>') | Out-Null

    foreach ($r in $rows) {
        $c1 = [System.Security.SecurityElement]::Escape($r[0])
        $c2 = [System.Security.SecurityElement]::Escape($r[1])
        $sbTbl.AppendLine('      <w:tr>') | Out-Null
        $sbTbl.AppendLine("        <w:tc><w:tcPr><w:tcW w:w=`"3000`" w:type=`"dxa`"/></w:tcPr><w:p><w:pPr><w:spacing w:before=`"80`" w:after=`"80`"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val=`"4338CA`"/></w:rPr><w:t>$c1</w:t></w:r></w:p></w:tc>") | Out-Null
        $sbTbl.AppendLine("        <w:tc><w:tcPr><w:tcW w:w=`"6400`" w:type=`"dxa`"/></w:tcPr><w:p><w:pPr><w:spacing w:before=`"80`" w:after=`"80`"/></w:pPr><w:r><w:rPr><w:color w:val=`"334155`"/></w:rPr><w:t>$c2</w:t></w:r></w:p></w:tc>") | Out-Null
        $sbTbl.AppendLine('      </w:tr>') | Out-Null
    }

    $sbTbl.AppendLine('    </w:tbl>') | Out-Null
    return $sbTbl.ToString()
}

$sbDoc = New-Object System.Text.StringBuilder
$sbDoc.AppendLine('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>') | Out-Null
$sbDoc.AppendLine('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">') | Out-Null
$sbDoc.AppendLine('  <w:body>') | Out-Null

# DOCUMENT COVER & TITLE
$sbDoc.AppendLine((W-Para -text "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" -sz 20 -bold $true -align "center" -spaceAfter 20)) | Out-Null
$sbDoc.AppendLine((W-Para -text "Độc lập - Tự do - Hạnh phúc" -sz 20 -bold $true -align "center" -spaceAfter 240)) | Out-Null
$sbDoc.AppendLine((W-Para -text "BÁO CÁO TỔNG QUAN HỆ THỐNG" -color "4338CA" -sz 42 -bold $true -align "center" -spaceBefore 200 -spaceAfter 60)) | Out-Null
$sbDoc.AppendLine((W-Para -text "HỆ THỐNG PHÂN TÍCH VÀ ĐỐI SOÁT DỮ LIỆU KHÁCH HÀNG DOANH NGHIỆP" -color "1E1B4B" -sz 30 -bold $true -align "center" -spaceAfter 120)) | Out-Null
$sbDoc.AppendLine((W-Para -text "Nền tảng Tự Động Hóa Chuẩn Hóa Dữ Liệu, Kiểm Soát An Toàn Doanh Thu & Hỗ Trợ Ra Quyết Định" -color "64748B" -sz 22 -align "center" -spaceAfter 300)) | Out-Null

# SECTION 1
$sbDoc.AppendLine((W-Heading1 -text "1. TỔNG QUAN VÀ BỐI CẢNH XÂY DỰNG HỆ THỐNG")) | Out-Null
$sbDoc.AppendLine((W-Para -text "Trong kỷ nguyên số hóa và quản trị dữ liệu quy mô lớn, các doanh nghiệp cung cấp dịch vụ hạ tầng, phần mềm và giải pháp B2B phải đối mặt với khối lượng thông tin khách hàng khổng lồ từ nhiều nguồn khác nhau. Những bất cập phổ biến gồm: dữ liệu bị phân mảnh, tình trạng khuyết thiếu dữ liệu đo lường, sai lệch định dạng ngày tháng và đặc biệt là sự thiếu đồng bộ giữa các bộ phận Kỹ thuật, Kinh doanh và Chăm sóc khách hàng.")) | Out-Null
$sbDoc.AppendLine((W-Para -text "Hệ thống Phân Tích Dữ Liệu Khách Hàng Doanh Nghiệp (Gia Linh Data Studio) được nghiên cứu và phát triển nhằm mục tiêu giải quyết triệt để các bài toán trên: tự động hóa 100% quy trình đối soát dữ liệu, làm sạch và phát hiện lỗi tự động, tính toán doanh thu có căn cứ vững chắc, và cung cấp bảng điều khiển trực quan giúp lãnh đạo doanh nghiệp ra quyết định chính xác, tức thì.")) | Out-Null

# SECTION 2
$sbDoc.AppendLine((W-Heading1 -text "2. MỤC TIÊU CHIẾN LƯỢC VÀ 4 TRỤ CỘT CỐT LÕI")) | Out-Null
$sbDoc.AppendLine((W-Para -text "Hệ thống được thiết kế dựa trên 4 trụ cột chiến lược cốt lõi:")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Tập Trung & Kết Nối Đa Nguồn" -content "Hợp nhất toàn bộ dữ liệu hợp đồng khách hàng, thông tin đo lường công suất, lịch sử hỗ trợ kỹ thuật và chính sách biểu phí vào một kho dữ liệu thống nhất, có khóa mốc thời gian đối chiếu chung.")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Chuẩn Hóa & Tách Biệt Tuyệt Đối NULL ≠ 0" -content "Loại bỏ hoàn toàn rủi ro hiểu nhầm giữa khuyết dữ liệu hệ thống (NULL) với giá trị thực tế bằng 0, bảo đảm tính kỷ luật và toàn vẹn của số liệu.")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Khóa Chặn An Toàn Kiểm Soát Lỗi Kỹ Thuật (Hard-Block)" -content "Bảo vệ uy tín thương hiệu và quyền lợi khách hàng bằng cách tự động khóa, cấm chào mời tái ký hoặc nâng cấp khi tài khoản đang gặp sự cố kỹ thuật chưa được xử lý.")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Cổng Kiểm Soát Phê Duyệt Cấp Quản Lý (HITL Gate)" -content "Kết hợp sức mạnh tính toán tự động với năng lực ra quyết định của con người; chỉ chuyển các trường hợp ngoại lệ (chiết khấu lớn, hồ sơ ranh giới) lên cấp quản lý xét duyệt.")) | Out-Null

# SECTION 3
$sbDoc.AppendLine((W-Heading1 -text "3. QUY TRÌNH 4 GIAI ĐOẠN XỬ LÝ DỮ LIỆU TỰ ĐỘNG")) | Out-Null
$rowsQuyTrinh = @(
    @("Giai Đoạn 1: Kết Nối Nguồn Dữ Liệu", "Tiếp nhận tệp dữ liệu khách hàng (.xlsx, .csv, .json), kiểm tra tính toàn vẹn và đồng bộ theo mốc ngày đối chiếu chung."),
    @("Giai Đoạn 2: Làm Sạch & Phân Định Dữ Liệu", "Khử trùng lặp, chuẩn hóa định dạng ngày chuẩn quốc tế, cô lập các trường khuyết thiếu (NULL) và phân loại riêng biệt với giá trị 0%."),
    @("Giai Đoạn 3: Phân Tích & Tính Toán Chính Sách", "Xác định mức độ sử dụng hạ tầng, đối chiếu điều kiện nâng cấp gói dịch vụ và áp dụng chính sách chiết khấu theo quy tắc xác định rõ ràng."),
    @("Giai Đoạn 4: Trực Quan Hóa & Xuất Kết Quả", "Cung cấp bảng điều khiển BI (Dashboard), xuất tệp Excel đã làm sạch kèm nhận xét AI và kích hoạt trợ lý tra cứu tự động.")
)
$sbDoc.AppendLine((W-Table2Col -rows $rowsQuyTrinh)) | Out-Null

# SECTION 4
$sbDoc.AppendLine((W-Heading1 -text "4. CƠ CHẾ KIỂM SOÁT AN TOÀN SỐ LIỆU VÀ PHÂN ĐỊNH NULL ≠ 0")) | Out-Null
$sbDoc.AppendLine((W-Para -text "Một trong những đột phá lớn nhất của hệ thống là khả năng phân biệt chính xác bản chất của dữ liệu khuyết thiếu:")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Dữ Liệu Khuyết Thiếu (NULL / Missing)" -content "Xảy ra khi cổng API đo lường bị gián đoạn kết nối. Hệ thống tuyệt đối không tự ý gán bằng 0%; thay vào đó, gắn cờ cảnh báo, tạm giữ doanh thu không tính vào doanh thu đủ căn cứ và yêu cầu Kỹ thuật hoàn tất kết nối.")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Giá Trị Bằng 0 (0% Usage)" -content "Xác nhận khách hàng chưa phát sinh lưu lượng thực tế (dữ liệu hợp lệ). Hệ thống đề xuất gói duy trì dịch vụ chi phí thấp kèm kỹ thuật viên đồng hành phục hồi.")) | Out-Null

# SECTION 5
$sbDoc.AppendLine((W-Heading1 -text "5. BỘ CHẨN ĐOÁN LỖI DỮ LIỆU MINH BẠCH (WHERE – WHY – FIX GUIDE)")) | Out-Null
$sbDoc.AppendLine((W-Para -text "Để hỗ trợ người dùng vận hành, hệ thống tích hợp công cụ chẩn đoán lỗi tệp đầu vào theo cấu trúc 3 thành phần rõ ràng:")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Vị Trí Lỗi (Where)" -content "Xác định chính xác số dòng và tên cột dữ liệu phát sinh lỗi (VD: Dòng 5 | Cột: Ngày hết hạn).")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Nguyên Nhân Sai (Why)" -content "Chỉ rõ lý do không hợp lệ (VD: Ngày 32 tháng 13 không tồn tại, số âm bất hợp lý, hoặc khuyết log).")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Hướng Dẫn Sửa (Fix Guide) & Nút Tự Động Sửa" -content "Cung cấp hướng dẫn quy chuẩn khắc phục và cung cấp nút bấm [Tự Động Sửa Lỗi] để hệ thống tự động chuẩn hóa chỉ bằng 1 cú click.")) | Out-Null

# SECTION 6
$sbDoc.AppendLine((W-Heading1 -text "6. GIAO DIỆN VẬN HÀNH 4 BƯỚC TUẦN TỰ & TÔNG MÀU GRADIENT HIỆN ĐẠI")) | Out-Null
$sbDoc.AppendLine((W-Para -text "Giao diện hệ thống được thiết kế theo phong cách hiện đại, trực quan và tối ưu triệt để trải nghiệm người dùng:")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Thanh Công Cụ Dọc 4 Bước Tuần Tự (Collapsible Sidebar)" -content "Toàn bộ chu trình thao tác được tối giản thành 4 bước duy nhất: Bước 1: Nạp Dữ Liệu; Bước 2: Quét & Chẩn Đoán Lỗi; Bước 3: Chạy Tính Toán; Bước 4: Xuất Kết Quả Excel. Thanh công cụ có thể thu gọn linh hoạt để mở rộng không gian quan sát.")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Tông Màu Gradient Ngọc Lục Bảo (Emerald/Mint) - Xanh Cyan - Xanh Dương Hoàng Gia" -content "Bảng màu chuyển sắc tinh tế (Emerald #10b981 sang Cyan #06b6d4, Sky Blue #0284c7 và Royal Blue/Indigo #6366f1) kết hợp thẻ nội dung trắng tinh tế trên nền sáng, mang lại cảm giác công nghệ cao và chuyên nghiệp.")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "Chính Sách Khởi Đầu Zero-Data & Khóa An Toàn" -content "Khi mở ứng dụng, mọi chỉ số và biểu đồ đều ở mức 0, tuyệt đối không hiển thị bất kỳ lỗi nào khi chưa bấm quét thử; nút tải file Excel được khóa an toàn cho đến khi dữ liệu được nạp và tính toán thành công.")) | Out-Null

# SECTION 7
$sbDoc.AppendLine((W-Heading1 -text "7. PHÂN TÍCH MA TRẬN SWOT CỦA HỆ THỐNG")) | Out-Null
$rowsSwot = @(
    @("Điểm Mạnh (Strengths - S)", "Quy tắc xác định 100% minh bạch, loại bỏ hoàn toàn sai lệch số liệu; chốt chặn an toàn bảo vệ khách hàng; xử lý dữ liệu tức thì < 2 giây; giao diện thuần Việt trực quan."),
    @("Điểm Yếu (Weaknesses - W)", "Phụ thuộc vào mức độ hoàn thiện của log đo lường hạ tầng; đòi hỏi ý thức phối hợp cập nhật trạng thái đóng ticket của đội ngũ kỹ thuật."),
    @("Cơ Hội (Opportunities - O)", "Khả năng mở rộng quy mô xử lý hàng chục ngàn doanh nghiệp; tích hợp thêm các kênh thông báo tự động (Email, SMS, Zalo ZNS); tối ưu mô hình học máy dự báo rủi ro rời bỏ."),
    @("Thách Thức (Threats - T)", "Chính sách biểu phí và khuyến mại doanh nghiệp thay đổi liên tục đòi hỏi bảng quy tắc phải linh hoạt; yêu cầu bảo mật thông tin tài chính nghiêm ngặt.")
)
$sbDoc.AppendLine((W-Table2Col -rows $rowsSwot)) | Out-Null

# SECTION 8
$sbDoc.AppendLine((W-Heading1 -text "8. KẾ HOẠCH TRIỂN KHAI VÀ MỞ RỘNG NĂM 2037")) | Out-Null
$sbDoc.AppendLine((W-Para -text "Lộ trình phát triển hệ thống trong năm 2037 được chia làm hai giai đoạn trọng tâm:")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "6 Tháng Đầu Năm 2037 (Chuẩn Hóa & Vận Hành Chính Thức)" -content "Đưa hệ thống vào vận hành thường trực tại doanh nghiệp; kết nối trực tiếp cơ sở dữ liệu qua API; đào tạo toàn diện nhân sự vận hành; duy trì tỷ lệ đối soát tự động đạt trên 98%.")) | Out-Null
$sbDoc.AppendLine((W-Bullet -title "6 Tháng Cuối Năm 2037 (Mở Rộng Quy Mô & Trí Tuệ Nhân Tạo)" -content "Mở rộng hệ thống xử lý đa chi nhánh; tích hợp mô hình AI phân tích hành vi khách hàng; tự động hóa luồng phê duyệt cấp quản lý và giảm tỷ lệ rời bỏ khách hàng xuống dưới 3%.")) | Out-Null

# SECTION 9
$sbDoc.AppendLine((W-Heading1 -text "9. KẾT LUẬN")) | Out-Null
$sbDoc.AppendLine((W-Para -text "Hệ thống Phân Tích Dữ Liệu Khách Hàng Doanh Nghiệp (Gia Linh Data Studio) không chỉ là một công cụ báo cáo thuần túy, mà là một nền tảng vận hành dữ liệu toàn diện, bảo vệ an toàn doanh thu và chuẩn hóa quy trình làm việc liên phòng ban. Việc đưa hệ thống vào áp dụng thực tế sẽ tạo bước ngoặt lớn trong năng lực quản trị, tối ưu hóa chi phí và thúc đẩy tăng trưởng bền vững cho doanh nghiệp.")) | Out-Null

# FOOTER INFO
$sbDoc.AppendLine((W-Para -text "Hà Nội, Ngày 18 Tháng 11 Năm 2026" -sz 20 -align "right" -spaceBefore 200 -spaceAfter 40)) | Out-Null
$sbDoc.AppendLine((W-Para -text "BAN PHÁT TRIỂN & VẬN HÀNH HỆ THỐNG DỮ LIỆU" -sz 22 -bold $true -align "right" -spaceAfter 20)) | Out-Null

$sbDoc.AppendLine('  </w:body>') | Out-Null
$sbDoc.AppendLine('</w:document>') | Out-Null

[System.IO.File]::WriteAllText((Join-Path $wordDir "document.xml"), $sbDoc.ToString(), [System.Text.Encoding]::UTF8)

# Zip into target DOCX
[System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $targetDocx)
Copy-Item -Path $targetDocx -Destination $rootDocx -Force
Remove-Item -Path $tempDir -Recurse -Force

Write-Output ("Successfully created Word Report: " + $targetDocx + " (" + (Get-Item $targetDocx).Length + " bytes)")
