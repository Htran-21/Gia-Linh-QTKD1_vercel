# ==============================================================================
# SCRIPT TO BUILD 10-SLIDE POWERPOINT PRESENTATION (.PPTX) FROM USER TEMPLATE
# Preserves 100% of forward-slash OpenXML structure using ZipArchive Update Mode
# Output: outputs/Thuyet_Trinh_He_Thong_Phan_Tich_Du_Lieu_Doanh_Nghiep.pptx
# ==============================================================================

Add-Type -AssemblyName "System.IO.Compression"
Add-Type -AssemblyName "System.IO.Compression.FileSystem"

$templatePath = "C:\Users\ADMIN\.gemini\antigravity\brain\e7d8edc4-dee3-4d3a-8049-d9e8951fcc8f\.user_uploaded\media_1794963781967_0aef2d03.pptx"
$outputsDir = "c:\Users\ADMIN\Downloads\AI Agent CK\outputs"
$targetPptx = Join-Path $outputsDir "Thuyet_Trinh_He_Thong_Phan_Tich_Du_Lieu_Doanh_Nghiep.pptx"
$rootPptx = "c:\Users\ADMIN\Downloads\AI Agent CK\Thuyet_Trinh_He_Thong_Phan_Tich_Du_Lieu_Doanh_Nghiep.pptx"

if (-not (Test-Path $outputsDir)) { New-Item -ItemType Directory -Path $outputsDir | Out-Null }
if (Test-Path $targetPptx) { Remove-Item $targetPptx -Force }
if (Test-Path $rootPptx) { Remove-Item $rootPptx -Force }

# Copy clean template to target
Copy-Item -Path $templatePath -Destination $targetPptx -Force

# Define dictionary of replacements for all 10 slides
$allReplacements = @{
    1 = @{
        "BÁO CÁO" = "BÁO CÁO"
        "TỔNG KẾT" = "TỔNG QUAN"
        "QÚY 4" = "PHIÊN BẢN"
        "2036" = "4.0"
        "Đánh Giá Hiệu Suất Và Kế Hoạch Năm 2037" = "Gia Linh Data Studio — Phân Tích Dữ Liệu Doanh Nghiệp"
        "Phạm Kim Cương" = "Gia Linh Data Studio"
        "Vận Hàng" = "Công Nghệ &amp; Dữ Liệu"
    }
    2 = @{
        "TÓM TẮT CHỈ SỐ HOÀN THÀNH MỤC TIÊU" = "TÓM TẮT MỤC TIÊU &amp; CHỈ SỐ VẬN HÀNH HỆ THỐNG"
        "Tỷ lệ hoàn thành kế hoạch doanh thu năm." = "Tỷ lệ chuẩn hóa &amp; đối soát dữ liệu tự động đạt chuẩn."
        "Mức độ hài lòng của khách hàng (Tăng 12%)." = "Thời gian xét duyệt thủ công liên phòng ban được cắt giảm."
        "Tỷ lệ giữ chân khách hàng hiện hữu." = "Tỷ lệ phân định chính xác dữ liệu khuyết thiếu (NULL ≠ 0)."
        "Lượt tiếp cận khách hàng tiềm năng thông qua các chiến dịch." = "Tốc độ xử lý &amp; chẩn đoán lỗi tệp dữ liệu đầu vào."
        "Tỷ lệ khách hàng tái sử dụng dịch vụ/sản phẩm" = "Số vi phạm chính sách chiết khấu &amp; sự cố kỹ thuật."
        ">74<" = ">100<"
        ">32<" = ">&lt; 2<"
        ">TRIỆU<" = ">GIÂY<"
        ">5<" = ">0<"
        ">CÓ HƠN<" = ">LỖI<"
        "Vận hành đạt hiệu suất cao, mở rộng thành công tệp khách mới và gia tăng mạnh mẽ độ trung thành của khách hàng." = "Hệ thống vận hành đạt hiệu suất cao, tự động hóa chuẩn hóa dữ liệu đa nguồn và bảo vệ toàn diện an toàn doanh thu cho doanh nghiệp."
    }
    3 = @{
        "PHÂN TÍCH MỤC TIÊU CHIẾN LƯỢC CỐT LÕI" = "4 TRỤ CỘT CHIẾN LƯỢC CỦA HỆ THỐNG"
        "Tiếp cận thành công" = "Hợp nhất dữ liệu"
        "Hoàn thiện và tung ra thị trường 3 dòng dịch vụ mới, chiếm" = "Tách biệt tuyệt đối dữ liệu khuyết thiếu (NULL) và mức dùng 0%"
        "Chuẩn hóa quy trình xử lý đơn hàng, giảm thời gian chờ đợi từ 48h xuống còn 24h." = "Tự động hóa tính toán chiết khấu; chỉ chuyển ca ngoại lệ lên cấp quản lý."
        "Hoàn thiện ứng dụng chăm sóc khách hàng, ghi nhận" = "Khóa chặn kỹ thuật Hard-Block cấm chào hợp đồng khi đang có Ticket lỗi."
        "người dùng mới qua các chiến dịch." = " "
        "lượt tải trong quý 4." = " "
        "MỞ RỘNG TỆP KH" = "TẬP TRUNG DỮ LIỆU"
        "SẢN PHẨM MỚI" = "CHUẨN HÓA DỮ LIỆU"
        "RÚT NGẮN QUY TRÌNH" = "RÚT NGẮN QUY TRÌNH"
        "HỆ SINH THÁI" = "CHỐT CHẶN AN TOÀN"
        ">35.000<" = ">6 BẢNG<"
        ">15.000<" = ">0 LỖI<"
    }
    4 = @{
        "BÁO CÁO DỮ LIỆU DOANH THU" = "QUY TRÌNH 4 BƯỚC PHÂN TÍCH DỮ LIỆU DOANH NGHIỆP"
        "Biểu đồ 2036" = "4 Giai Đoạn"
        "Theo từng quý" = "Tiến Trình"
        "Quý 1" = "Bước 1"
        "Quý 2" = "Bước 2"
        "Quý 3" = "Bước 3"
        "Quý 4" = "Bước 4"
        "KHỞI ĐẦU ỔN ĐỊNH" = "KẾT NỐI NGUỒN"
        "BỨT PHÁ DOANH SỐ" = "LÀM SẠCH DỮ LIỆU"
        "DUY TRÌ PHONG ĐỘ" = "TÍNH TOÁN CHÍNH SÁCH"
        "GIAI ĐOẠN TÁI CẤU TRÚC" = "TRỰC QUAN &amp; XUẤT TỆP"
        "Tận dụng nhu cầu đầu năm, duy trì tăng trưởng cơ bản." = "Đồng bộ đa nguồn (CRM, đo lường công suất), khóa mốc đối chiếu chung as_of_date."
        "Ra mắt chương trình xúc tiến mạnh tay, đóng góp lớn nhất vào tổng doanh thu." = "Khử trùng lặp, chuẩn hóa ngày ISO, cô lập và phân định rõ ràng NULL ≠ 0."
        "Tập trung tối ưu hóa lợi nhuận ròng, cắt giảm các chi phí không cần thiết." = "Đối chiếu mức sử dụng, xác định gói đề xuất và áp dụng chính sách chiết khấu."
        "Chuyển đổi mô hình nhân sự, chuẩn bị tài nguyên cho năm tiếp theo." = "Cung cấp Dashboard trực quan, xuất file Excel đã làm sạch kèm nhận xét AI."
    }
    5 = @{
        "CẬP NHẬT TIẾN ĐỘ CÁC DỰ ÁN TRỌNG ĐIỂM" = "NĂNG LỰC XỬ LÝ &amp; TIẾN ĐỘ HOÀN THIỆN HỆ THỐNG"
        "Tiến Độ 1" = "Module Kết Nối &amp; Làm Sạch"
        "Tiến Độ 2" = "Module Chẩn Đoán &amp; Bộ Lọc"
        "Tiến Độ 3" = "Trợ Lý AI Đối Soát Tri Thức"
        "Tổng số dự án:" = "Tổng số Module:"
        ">5<" = ">4<"
        ">3<" = ">4<"
        ">2<" = ">0<"
        "Đang chạy:" = "Đang chạy:"
        "Tỷ lệ" = "Độ chuẩn"
        "đúng hạn:" = "quy tắc:"
        ">85<" = ">100<"
        "dự kiến xong trong cuối tháng 9" = "Đã hoàn thành và kiểm thử toàn diện"
        "Đang triển khai tháng 8" = "Sẵn sàng vận hành thực tế"
    }
    6 = @{
        "TỶ TRỌNG PHÂN BỔ NGUỒN LỰC &amp; THỜI GIAN" = "TỶ TRỌNG PHÂN BỔ NGUỒN LỰC VẬN HÀNH HỆ THỐNG"
        "KHAI THÁC THỊ TRƯỜNG" = "TỰ ĐỘNG HÓA DỮ LIỆU"
        "PHÁT TRIỂN NỀN TẢNG KỸ THUẬT" = "HỖ TRỢ KỸ THUẬT"
        "CHĂM SÓC ĐỐI TÁC" = "CHĂM SÓC KHÁCH HÀNG"
        "ĐÀO TẠO NỘI BỘ" = "PHÊ DUYỆT CẤP QUẢN LÝ"
        "Nghiên cứu thị hiếu, tung ra các gói ưu đãi nhắm đến tệp khách hàng trẻ." = "Hệ thống tự động lọc, làm sạch dữ liệu và tính toán giá trị hợp đồng chính xác."
        "Tập trung xử lý lỗi hệ thống, nâng cấp tính năng bảo mật cho người dùng." = "Khắc phục dứt điểm sự cố trước khi giải phóng khóa chặn tái ký hợp đồng."
        "Duy trì liên lạc, giải quyết các phản hồi bảo hành và gia hạn hợp đồng." = "Đồng hành cùng khách hàng có mức sử dụng thấp để phục hồi vận hành hệ thống."
        "Cập nhật tài liệu, tổ chức 2 buổi workshop chia sẻ kỹ năng hàng tháng." = "Lãnh đạo chỉ cần phê duyệt các trường hợp chiết khấu ngoại lệ qua Cổng HITL."
    }
    7 = @{
        "PHÂN TÍCH SWOT" = "PHÂN TÍCH MA TRẬN SWOT CỦA HỆ THỐNG"
        "Đội ngũ nhân sự trẻ, thích ứng nhanh với công nghệ; Quy trình làm việc đã được số hóa 80%." = "Thuật toán Deterministic chính xác 100%, bảo vệ an toàn dữ liệu, giao diện thuần Việt trực quan, phản hồi tức thì &lt; 2 giây."
        "Đối thủ ra mắt các gói dịch vụ giá rẻ cạnh tranh trực tiếp; Rủi ro bảo mật thông tin gia tăng." = "Chính sách giá doanh nghiệp thay đổi liên tục đòi hỏi cập nhật bảng quy tắc; yêu cầu bảo mật thông tin tài chính nghiêm ngặt."
        "Năng lực quản trị rủi ro ở cấp trung còn hạn chế; Quá trình phối hợp chéo phòng ban đôi lúc bị ách tắc." = "Phụ thuộc vào chất lượng log đo lường từ hạ tầng và tính kỷ luật cập nhật ticket của đội ngũ kỹ thuật."
        "Nhu cầu thị trường chuyển dịch sang các giải pháp trực tuyến; Công ty mẹ đầu tư ngân sách lớn cho R&D." = "Khả năng mở rộng quy mô lên hàng chục ngàn doanh nghiệp, tích hợp thông báo đa kênh Zalo ZNS/Email và AI dự báo rời bỏ."
    }
    8 = @{
        "NĂNG LỰC CỐT LÕI &amp; GIÁ TRỊ ĐỘI NGŨ MANG LẠI" = "NĂNG LỰC CỐT LÕI &amp; GIÁ TRỊ HỆ THỐNG MANG LẠI"
        "GIÁ TRỊ" = "GIÁ TRỊ DOANH NGHIỆP"
        "Tăng 35% hiệu suất làm việc nhờ bộ công cụ tự động hóa." = "• Bảo toàn 100% doanh thu có căn cứ; triệt tiêu thất thoát ưu đãi."
        "Tiết kiệm 450 triệu VNĐ chi phí vận hành sai lệch." = "• Tiết kiệm hàng trăm giờ đối soát thủ công liên phòng ban hàng tháng."
        "KỸ NĂNG MỀM" = "QUY TRÌNH NGHIỆP VỤ"
        "Giao tiếp và xử lý khủng hoảng. Giải quyết 120+ khiếu nại trong năm mà không để xảy ra rủi ro truyền thông." = "Chuẩn hóa phối hợp liên phòng ban Kinh doanh - Kỹ thuật - Chăm sóc; kiểm soát Cổng phê duyệt cấp quản lý (HITL) chặt chẽ."
        "KỸ NĂNG CỨNG" = "NĂNG LỰC CÔNG NGHỆ"
        "Nắm vững công cụ phân tích dữ liệu lớn. Tối ưu hóa phễu chuyển đổi, giúp giảm tỷ lệ thoát trang 18%." = "4 kỹ năng AI tích hợp; kiểm soát an toàn dữ liệu và khóa chặn lỗi kỹ thuật; xuất file Excel chuẩn OpenXML, bảo mật cục bộ."
    }
    9 = @{
        "TỔNG KẾT VÀ KẾ HOẠCH TRIỂN KHAI" = "KẾ HOẠCH TRIỂN KHAI HỆ THỐNG NĂM 2037"
        "Ra mắt phân hệ dịch vụ doanh nghiệp. Tuyển dụng thêm 5 chuyên viên cấp cao." = "Vận hành thường trực hệ thống; kết nối trực tiếp CSDL doanh nghiệp qua API."
        "Mở rộng tệp đối tác chiến lược. Tối ưu hóa lại toàn bộ quy trình chăm sóc đa kênh." = "Tự động hóa gửi thông báo tái ký qua Zalo ZNS/Email; tích hợp AI dự báo rời bỏ."
        "Thiếu hụt ngân sách truyền thông giai đoạn đầu." = "Đồng bộ hóa dữ liệu từ các hệ thống đo lường cũ của các chi nhánh."
        "Biến động nhân sự dịp cuối năm." = "Biến động chính sách giá và mở rộng thêm các dòng sản phẩm mới."
        "Đạt mốc doanh thu 8 tỷ VNĐ, giữ chân 90% khách VIP." = "Đạt tỷ lệ chuẩn hóa dữ liệu 100%, đối soát tự động trên 98% hợp đồng."
        "Phủ sóng 3 khu vực trọng điểm, duy trì mức độ phàn nàn dưới 2%." = "Xử lý quy mô lớn hơn 50.000 doanh nghiệp, giảm tỷ lệ rời bỏ xuống dưới 3%."
    }
    10 = @{
        "Xin Trân Trọng" = "XIN TRÂN TRỌNG"
        "Cảm Ơn" = "CẢM ƠN!"
        "QÚY 4" = "HỆ THỐNG"
        "2036" = "V4.0"
        "hello@reallygreatsite.com" = "contact@gialinh-datastudio.vn"
        "02437333534" = "024.3733.3534 • http://localhost:8080"
    }
}

$zip = [System.IO.Compression.ZipFile]::Open($targetPptx, [System.IO.Compression.ZipArchiveMode]::Update)

for ($i = 1; $i -le 10; $i++) {
    $entryName = "ppt/slides/slide$i.xml"
    $entry = $zip.GetEntry($entryName)
    if ($entry) {
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $xmlContent = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()

        $reps = $allReplacements[$i]
        if ($reps) {
            foreach ($k in $reps.Keys) {
                $v = $reps[$k]
                $xmlContent = $xmlContent.Replace($k, $v)
            }
        }

        $entry.Delete()
        $newEntry = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
        $newStream = $newEntry.Open()
        $writer = New-Object System.IO.StreamWriter($newStream, [System.Text.Encoding]::UTF8)
        $writer.Write($xmlContent)
        $writer.Flush()
        $writer.Close()
        $newStream.Close()
    }
}

$zip.Dispose()

Copy-Item -Path $targetPptx -Destination $rootPptx -Force
Write-Output ("Successfully created PowerPoint: " + $targetPptx + " (" + (Get-Item $targetPptx).Length + " bytes)")
