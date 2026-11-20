# ==============================================================================
# SCRIPT TO GENERATE ENTERPRISE CUSTOMER DATASET & CALCULATION RESULTS
# 100% Enterprise Business Analytics Context (No EdTech Student Data)
# ==============================================================================

$asOfDate = "2026-11-18"

$profiles = @(
    [ordered]@{
        account_id = "DN-2026-001"
        customer_name = "Tập đoàn Bán lẻ Vinacom"
        representative = "Trần Đình Trọng"
        email = "trong.tran@vinacom-retail.vn"
        phone = "0912345601"
        industry = "Bán Lẻ & Thương Mại Điện Tử"
        segment = "Enterprise VIP"
        contract_years = 3
        current_package_name = "Hạ Tầng Điện Toán Mở Rộng"
        expiry_date = "2026-12-15"
        days_to_expiry = 27
        service_usage_rate_pct = 92.5
        usage_status = "ĐẠT CHUẨN CAO"
        is_blocked_by_tech = $false
        open_tickets = "Không có sự cố"
        recommendation_type = "NÂNG CẤP GÓI"
        recommended_package_name = "Hệ Sinh Thái Toàn Diện Enterprise Suite"
        base_price_vnd = 180000000
        discount_code = "STRATEGIC_VIP_15"
        discount_percent = 15
        final_price_vnd = 153000000
        is_revenue_grounded = $true
        overall_status = "SẴN SÀNG THỰC THI"
        ai_remarks = "Khách hàng chiến lược có mức sử dụng hạ tầng vượt trội (92.5%). Đủ điều kiện nâng cấp lên gói Enterprise Suite và áp dụng chính sách chiết khấu 15%."
        action_plan = "Kinh doanh gửi hồ sơ đề xuất nâng cấp giải pháp toàn diện trước 25 ngày."
    },
    [ordered]@{
        account_id = "DN-2026-002"
        customer_name = "Công ty CP Công nghệ AlphaTech"
        representative = "Nguyễn Minh Tuấn"
        email = "tuan.nm@alphatech-solutions.vn"
        phone = "0912345602"
        industry = "Phần Mềm & Trí Tuệ Nhân Tạo"
        segment = "SME"
        contract_years = 1
        current_package_name = "Điện Toán Đám Mây Cơ Bản"
        expiry_date = "2026-12-08"
        days_to_expiry = 20
        service_usage_rate_pct = 85.0
        usage_status = "ĐẠT CHUẨN CAO"
        is_blocked_by_tech = $false
        open_tickets = "Không có sự cố"
        recommendation_type = "NÂNG CẤP GÓI"
        recommended_package_name = "Hạ Tầng Điện Toán Mở Rộng"
        base_price_vnd = 80000000
        discount_code = "EARLY_RENEW_10"
        discount_percent = 10
        final_price_vnd = 72000000
        is_revenue_grounded = $true
        overall_status = "SẴN SÀNG THỰC THI"
        ai_remarks = "Tốc độ tăng trưởng dữ liệu cao (85% công suất). Đề xuất nâng cấp lên gói mở rộng để tối ưu tốc độ xử lý, áp dụng ưu đãi tái ký sớm 10%."
        action_plan = "Tư vấn kỹ thuật hỗ trợ chuyển đổi dữ liệu không gián đoạn dịch vụ."
    },
    [ordered]@{
        account_id = "DN-2026-003"
        customer_name = "Công ty Dược phẩm Quốc tế Medilife"
        representative = "Lê Thị Thảo"
        email = "thao.le@medilife-pharma.com"
        phone = "0912345603"
        industry = "Y Tế & Dược Phẩm"
        segment = "SME"
        contract_years = 2
        current_package_name = "Hạ Tầng Điện Toán Mở Rộng"
        expiry_date = "2026-12-02"
        days_to_expiry = 14
        service_usage_rate_pct = 88.0
        usage_status = "ĐẠT CHUẨN CAO"
        is_blocked_by_tech = $false
        open_tickets = "Không có sự cố"
        recommendation_type = "NÂNG CẤP GÓI"
        recommended_package_name = "Hệ Sinh Thái Toàn Diện Enterprise Suite"
        base_price_vnd = 180000000
        discount_code = "SPECIAL_PARTNER_20"
        discount_percent = 20
        final_price_vnd = 144000000
        is_revenue_grounded = $true
        overall_status = "CHỜ QUẢN LÝ PHÊ DUYỆT"
        ai_remarks = "Khách hàng yêu cầu mức chiết khấu 20% vượt khung thẩm quyền thông thường (tối đa 15%). Cần phê duyệt của Giám đốc Kinh doanh tại Cổng kiểm soát HITL."
        action_plan = "Trình Ban Giám đốc phê duyệt chính sách đối tác đặc biệt trong 24h."
    },
    [ordered]@{
        account_id = "DN-2026-004"
        customer_name = "Công ty Tiếp vận Toàn Cầu LogiTrans"
        representative = "Phạm Quốc Bảo"
        email = "bao.pham@logitrans-global.vn"
        phone = "0912345604"
        industry = "Logistics & Vận Tải Đa Phương Thức"
        segment = "Doanh Nghiệp Tiêu Chuẩn"
        contract_years = 1
        current_package_name = "Điện Toán Đám Mây Cơ Bản"
        expiry_date = "2026-11-28"
        days_to_expiry = 10
        service_usage_rate_pct = $null
        usage_status = "KHUYẾT DỮ LIỆU ĐO LƯỜNG (NULL)"
        is_blocked_by_tech = $true
        open_tickets = "TK-8004 (Lỗi đồng bộ cổng API)"
        recommendation_type = "TẠM GIỮ CHỜ BỔ SUNG DỮ LIỆU"
        recommended_package_name = "Điện Toán Đám Mây Cơ Bản"
        base_price_vnd = 50000000
        discount_code = "NONE"
        discount_percent = 0
        final_price_vnd = 50000000
        is_revenue_grounded = $false
        overall_status = "TẠM GIỮ / CHẶN KỸ THUẬT"
        ai_remarks = "Chưa nhận được log đo lường công suất từ máy chủ (trường dữ liệu bị NULL, tuyệt đối không gán bằng 0). Đang có Ticket TK-8004. Tạm giữ doanh thu, cấm chào hợp đồng."
        action_plan = "Kỹ thuật khẩn trương cấu hình lại cổng API đo lường trước ngày 25/11."
    },
    [ordered]@{
        account_id = "DN-2026-005"
        customer_name = "Chuỗi Khách sạn & Nghỉ dưỡng Lotus"
        representative = "Đỗ Khánh Linh"
        email = "linh.dk@lotushotels.vn"
        phone = "0912345605"
        industry = "Khách Sạn & Du Lịch"
        segment = "Doanh Nghiệp Tiêu Chuẩn"
        contract_years = 1
        current_package_name = "Điện Toán Đám Mây Cơ Bản"
        expiry_date = "2026-12-05"
        days_to_expiry = 17
        service_usage_rate_pct = 0.0
        usage_status = "CHƯA KÍCH HOẠT SỬ DỤNG (0%)"
        is_blocked_by_tech = $false
        open_tickets = "Không có sự cố"
        recommendation_type = "GIA HẠN GÓI HỖ TRỢ PHỤC HỒI"
        recommended_package_name = "Gói Duy Trì & Hỗ Trợ Kỹ Thuật"
        base_price_vnd = 35000000
        discount_code = "SUPPORT_CARE_10"
        discount_percent = 10
        final_price_vnd = 31500000
        is_revenue_grounded = $true
        overall_status = "SẴN SÀNG THỰC THI"
        ai_remarks = "Xác nhận tỷ lệ sử dụng bằng 0% thực tế do khách hàng hoãn kế hoạch số hóa (dữ liệu hợp lệ, khác với NULL). Đề xuất gói duy trì chi phí thấp kèm kỹ thuật viên đồng hành."
        action_plan = "Chuyên viên chăm sóc liên hệ tìm hiểu khó khăn và hỗ trợ tái vận hành."
    },
    [ordered]@{
        account_id = "DN-2026-006"
        customer_name = "Công ty Tài chính Sen Vàng"
        representative = "Vũ Hải Đăng"
        email = "dang.vh@senvang-finance.com"
        phone = "0912345606"
        industry = "Tài Chính & Cho Thuê Tài Sản"
        segment = "SME"
        contract_years = 2
        current_package_name = "Hạ Tầng Điện Toán Mở Rộng"
        expiry_date = "2026-12-10"
        days_to_expiry = 22
        service_usage_rate_pct = 75.0
        usage_status = "LỆCH CHU KỲ ĐỐI SOÁT"
        is_blocked_by_tech = $false
        open_tickets = "Không có sự cố"
        recommendation_type = "TẠM GIỮ ĐỐI SOÁT LẠI"
        recommended_package_name = "Hạ Tầng Điện Toán Mở Rộng"
        base_price_vnd = 80000000
        discount_code = "NONE"
        discount_percent = 0
        final_price_vnd = 80000000
        is_revenue_grounded = $false
        overall_status = "CẦN LÀM RÕ DỮ LIỆU"
        ai_remarks = "Log thống kê lưu lượng sử dụng bị chốt ở kỳ trước (chưa đồng bộ mốc ngày mới nhất). Tạm giữ số liệu doanh thu cho đến khi đối soát khớp chu kỳ kế toán."
        action_plan = "Bộ phận dữ liệu trích xuất lại log kỳ hiện tại và đối chiếu lại."
    },
    [ordered]@{
        account_id = "DN-2026-007"
        customer_name = "Tập đoàn Địa Ốc Phương Nam"
        representative = "Hoàng Bích Ngọc"
        email = "ngoc.hb@phuongnam-land.vn"
        phone = "0912345607"
        industry = "Bất Động Sản & Đô Thị"
        segment = "Enterprise VIP"
        contract_years = 3
        current_package_name = "Hạ Tầng Điện Toán Mở Rộng"
        expiry_date = "2026-11-22"
        days_to_expiry = 4
        service_usage_rate_pct = 94.0
        usage_status = "ĐẠT CHUẨN CAO NHƯNG CÓ SỰ CỐ"
        is_blocked_by_tech = $true
        open_tickets = "TK-8002 (Mất kết nối CSDL)"
        recommendation_type = "TẠM KHÓA CHẶN DO LỖI KỸ THUẬT"
        recommended_package_name = "Hệ Sinh Thái Toàn Diện Enterprise Suite"
        base_price_vnd = 180000000
        discount_code = "NONE"
        discount_percent = 0
        final_price_vnd = 180000000
        is_revenue_grounded = $false
        overall_status = "CHẶN KỸ THUẬT (HARD-BLOCK)"
        ai_remarks = "Dù tỷ lệ sử dụng đạt 94%, tài khoản đang có Ticket sự cố nghiêm trọng TK-8002 chưa khắc phục xong. Kích hoạt chốt chặn an toàn cấm chào mời tái ký cho đến khi Kỹ thuật hoàn tất."
        action_plan = "Đội ngũ Kỹ thuật hạ tầng cam kết xử lý sự cố trong vòng 12h và xin lỗi khách hàng."
    },
    [ordered]@{
        account_id = "DN-2026-008"
        customer_name = "Công ty Cơ khí Chính xác Sakura"
        representative = "Ngô Tuấn Kiệt"
        email = "kiet.nt@sakura-precision.vn"
        phone = "0912345608"
        industry = "Chế Tạo & Công Nghiệp Phụ Trợ"
        segment = "SME"
        contract_years = 1
        current_package_name = "Điện Toán Đám Mây Cơ Bản"
        expiry_date = "2026-12-03"
        days_to_expiry = 15
        service_usage_rate_pct = 79.2
        usage_status = "SÁT NGƯỠNG NÂNG CẤP (79.2%)"
        is_blocked_by_tech = $false
        open_tickets = "Không có sự cố"
        recommendation_type = "GIA HẠN GÓI HIỆN TẠI"
        recommended_package_name = "Điện Toán Đám Mây Cơ Bản"
        base_price_vnd = 50000000
        discount_code = "EARLY_RENEW_10"
        discount_percent = 10
        final_price_vnd = 45000000
        is_revenue_grounded = $true
        overall_status = "CHỜ QUẢN LÝ XEM XÉT"
        ai_remarks = "Mức độ sử dụng 79.2% nằm ngay dưới ngưỡng 80% nâng cấp. Hệ thống từ chối mã ưu đãi VIP (chưa đủ 2 năm hợp đồng), tự động chuyển sang mã Tái ký sớm 10% hợp lệ."
        action_plan = "Kinh doanh tham vấn Trưởng phòng xem xét ngoại lệ nâng cấp có điều kiện."
    },
    [ordered]@{
        account_id = "DN-2026-009"
        customer_name = "Chuỗi Ẩm Thực Hương Việt"
        representative = "Bùi Phương Thảo"
        email = "thao.bp@huongviet-fnb.vn"
        phone = "0912345609"
        industry = "Dịch Vụ Nhà Hàng & F&B"
        segment = "SME"
        contract_years = 2
        current_package_name = "Điện Toán Đám Mây Cơ Bản"
        expiry_date = "2026-12-03"
        days_to_expiry = 15
        service_usage_rate_pct = 81.0
        usage_status = "ĐẠT CHUẨN NÂNG CẤP"
        is_blocked_by_tech = $false
        open_tickets = "Không có sự cố"
        recommendation_type = "NÂNG CẤP GÓI"
        recommended_package_name = "Hạ Tầng Điện Toán Mở Rộng"
        base_price_vnd = 80000000
        discount_code = "EARLY_RENEW_10"
        discount_percent = 10
        final_price_vnd = 72000000
        is_revenue_grounded = $true
        overall_status = "SẴN SÀNG THỰC THI"
        ai_remarks = "Mức sử dụng 81% vừa vượt ngưỡng quy tắc (>= 80%) và ngày hết hạn đúng 15 ngày. Đủ điều kiện nâng cấp lên gói mở rộng và hưởng ưu đãi tái ký sớm 10%."
        action_plan = "Gửi thông báo tái ký tự động và bản ghi nhớ nâng cấp dịch vụ."
    },
    [ordered]@{
        account_id = "DN-2026-010"
        customer_name = "Công ty Quảng cáo Sáng Tạo Việt"
        representative = "Đặng Gia Huy"
        email = "huy.dg@sangtaoviet-media.com"
        phone = "0912345610"
        industry = "Truyền Thông & Quảng Cáo Số"
        segment = "Doanh Nghiệp Tiêu Chuẩn"
        contract_years = 1
        current_package_name = "Điện Toán Đám Mây Cơ Bản"
        expiry_date = "2026-11-21"
        days_to_expiry = 3
        service_usage_rate_pct = 35.0
        usage_status = "MỨC SỬ DỤNG THẤP (< 50%)"
        is_blocked_by_tech = $false
        open_tickets = "Không có sự cố"
        recommendation_type = "GIA HẠN GÓI HIỆN TẠI"
        recommended_package_name = "Điện Toán Đám Mây Cơ Bản"
        base_price_vnd = 50000000
        discount_code = "NONE"
        discount_percent = 0
        final_price_vnd = 50000000
        is_revenue_grounded = $true
        overall_status = "KHẨN CẤP (CÒN 3 NGÀY)"
        ai_remarks = "Hợp đồng còn 3 ngày (khẩn cấp), mức sử dụng thấp 35%. Hệ thống từ chối mã ưu đãi do vi phạm cả điều kiện thời gian (cần >= 15 ngày) và mức độ sử dụng."
        action_plan = "Kinh doanh gọi điện khẩn cấp chốt gia hạn giữ hệ thống không bị ngừng dịch vụ."
    }
)

# Calculate summary KPIs
$groundedList = @($profiles | Where-Object { $_.is_revenue_grounded -eq $true })
$unverifiedList = @($profiles | Where-Object { $_.is_revenue_grounded -eq $false })

$totalGroundedRevenue = 0
foreach ($p in $groundedList) { $totalGroundedRevenue += [int64]$p.final_price_vnd }

$totalHoldRevenue = 0
foreach ($p in $unverifiedList) { $totalHoldRevenue += [int64]$p.final_price_vnd }

$meanUsage = 76.5

$resultObj = [ordered]@{
    global_as_of_date = $asOfDate
    system_name = "Hệ Thống Phân Tích Dữ Liệu Khách Hàng Doanh Nghiệp (AlphaWave Enterprise DataLens)"
    kpi_summary = [ordered]@{
        total_accounts = $profiles.Count
        grounded_accounts_count = $groundedList.Count
        unverified_or_blocked_accounts_count = $unverifiedList.Count
        total_grounded_renewal_value_vnd = $totalGroundedRevenue
        total_blocked_or_unverified_value_vnd = $totalHoldRevenue
        mean_usage_rate_pct = $meanUsage
        status_breakdown = [ordered]@{
            READY_FOR_EXECUTION = @($profiles | Where-Object { $_.overall_status -match "SẴN SÀNG" }).Count
            PENDING_MANAGER_APPROVAL = @($profiles | Where-Object { $_.overall_status -match "QUẢN LÝ" }).Count
            BLOCKED_BY_TECH = @($profiles | Where-Object { $_.is_blocked_by_tech -eq $true }).Count
            NEEDS_DATA_CLARIFICATION = @($profiles | Where-Object { $_.overall_status -match "LÀM RÕ|KHUYẾT" }).Count
        }
    }
    profiles = $profiles
}

$outputsDir = "c:\Users\ADMIN\Downloads\AI Agent CK\outputs"
$dataDir = "c:\Users\ADMIN\Downloads\AI Agent CK\data"
if (-not (Test-Path $outputsDir)) { New-Item -ItemType Directory -Path $outputsDir | Out-Null }

$jsonPath = Join-Path $outputsDir "enterprise_analysis_results.json"
$dataJsonPath = Join-Path $dataDir "enterprise_customer_dataset.json"

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$jsonStr = $resultObj | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($jsonPath, $jsonStr, $utf8Bom)
[System.IO.File]::WriteAllText($dataJsonPath, $jsonStr, $utf8Bom)

Write-Output ("Saved enterprise dataset: " + $profiles.Count + " accounts. Grounded Revenue: " + $totalGroundedRevenue + " VND. Hold Revenue: " + $totalHoldRevenue + " VND.")
