# BÁO CÁO TỔNG QUAN HỆ THỐNG PHÂN TÍCH DỮ LIỆU KHÁCH HÀNG & ĐIỀU PHỐI GIA HẠN
**Đơn vị thực hiện:** Nhóm Phát triển Giải pháp AI Agent  
**Doanh nghiệp giả định:** Công ty Cổ phần Giáo dục Trực tuyến EdTech Future  
**Nền tảng vận hành:** DataLens Analytics Engine (Localhost:8080)  
**Mốc ngày đối chiếu chung (Global As-Of Date):** 2026-11-16  

---

## PHẦN 1: TÓM TẮT ĐỀ BÀI & BỐI CẢNH DOANH NGHIỆP

### 1.1. Bối cảnh hoạt động
Công ty Giáo dục EdTech Future là doanh nghiệp cung cấp các gói tài khoản học ngoại ngữ trực tuyến theo năm. Mỗi tháng có khoảng **150 tài khoản học viên** đến kỳ hạn cần gia hạn hoặc nâng cấp lộ trình học mới. 

Để đưa ra đề xuất khóa học phù hợp, 3 bộ phận nghiệp vụ bắt buộc phải phối hợp chéo:
1. **Đội Chăm sóc Học viên (CS):** Rà soát tiến độ học tập, tần suất tham gia lớp học, kiểm tra mức độ đạt cam kết đầu ra và nhu cầu hỗ trợ.
2. **Nhân viên Tư vấn Tuyển sinh (Sales):** Xác định thời điểm gọi điện tư vấn (trước ngày hết hạn 30 ngày), đề xuất giữ gói cũ hoặc nâng cấp gói mới và áp dụng chính sách giảm giá.
3. **Bộ phận Kỹ thuật - Học vụ (Tech/Support):** Rà soát các sự cố kỹ thuật (lỗi micro, giật lag, văng khỏi lớp học, tài khoản bị treo) và tình trạng xử lý các yêu cầu hỗ trợ (Support Tickets).

### 1.2. Ba Vấn đề Bất cập Cốt lõi & Hậu quả Thực tế
1. **Chào gói nâng cao khi tài khoản đang gặp sự cố kỹ thuật chưa xử lý:**
   - *Thực trạng:* Sales không nắm được học viên đang mở Ticket khiếu nại văng khỏi lớp, vẫn gọi điện chào mua gói nâng cao.
   - *Hậu quả:* Học viên bức xúc, đánh giá dịch vụ thiếu chuyên nghiệp, từ chối gia hạn và có nguy cơ lan truyền phản hồi tiêu cực trên mạng xã hội.
2. **Áp dụng mã giảm giá sai điều kiện chính sách:**
   - *Thực trạng:* Nhân viên tự ý áp các mã ưu đãi lớn (Early Bird 10%, High Achiever 15%, Loyalty VIP 20%) cho học viên không đủ điều kiện (ví dụ: ngày liên hệ quá sát ngày hết hạn, hoặc tỷ lệ chuyên cần quá thấp).
   - *Hậu quả:* Doanh nghiệp thất thoát doanh thu, tạo tiền lệ xấu, vi phạm kỷ luật tài chính.
3. **Nhầm lẫn dữ liệu ngày tháng & Số buổi học:**
   - *Thực trạng:* Nhầm ngày hết hạn hợp đồng (`expiry_date`) với ngày cần bắt đầu gọi tư vấn (`contact_start_date`); nhầm lẫn số buổi đăng ký với số buổi thực học; đặc biệt nhầm lẫn giữa **dữ liệu chưa ghi nhận (`NULL`)** với **học viên vắng học (`0 buổi`)**.
   - *Hậu quả:* Gọi điện quá trễ sát ngày hết hạn khiến tỷ lệ mất khách (Churn) tăng cao; phân loại sai nhu cầu học viên; Trưởng phòng Kinh doanh bị quá tải vì phải kiểm tra thủ công 100% hồ sơ trước khi duyệt.

---

## PHẦN 2: THIẾT KẾ GIẢI PHÁP & KIẾN TRÚC HỆ THỐNG AI AGENT

### 2.1. Kiến trúc Tổng thể: Mục tiêu – Đầu vào – Quy trình – Đầu ra (GIPO)

```
┌────────────────────────────────────────────────────────────────────────┐
│ MỤC TIÊU (GOAL)                                                        │
│ • Tự động hóa 100% khâu chuẩn bị danh sách gia hạn có căn cứ.          │
│ • Triệt tiêu xung đột dữ liệu giữa Sales - CS - Tech.                  │
│ • Giải phóng 85% thời gian phê duyệt thủ công của Trưởng phòng.        │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
┌───────────────────────────────────▼────────────────────────────────────┐
│ ĐẦU VÀO (INPUT) - ĐA NGUỒN                                             │
│ 1. CRM Accounts & Expiry (.xlsx / Database)                            │
│ 2. LMS Attendance Logs (Google Sheets / API)                           │
│ 3. Tech Support Tickets (Lark Base Bitable / API)                      │
│ 4. Course Catalog & Learning Paths (JSON)                              │
│ 5. Pricing Table (Biểu phí niêm yết)                                   │
│ 6. Discount Policies (Điều kiện ưu đãi & Thẩm quyền duyệt)             │
│ * Tất cả đồng bộ về mốc ngày đối chiếu chung: as_of_date = 2026-11-16  │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
┌───────────────────────────────────▼────────────────────────────────────┐
│ QUY TRÌNH (PROCESS) - 4 AI SKILLS & 2 HITL GATES                       │
│ • Skill 1: Ingestion & Validation (Phát hiện lỗi, chẩn đoán, sửa lỗi) │
│ • Skill 2: Data Cleaning (Khử trùng, chuẩn ISO, tách NULL ≠ 0)         │
│ • Skill 3: Deterministic Pricing Engine & Hard-Block Tech              │
│ • Skill 4: Guardrail Message Synthesis & Structured Audit Log          │
│ • HITL Gate 1: Quản lý duyệt ưu đãi ngoại lệ > 15%                     │
│ • HITL Gate 2: Kỹ thuật ký xác nhận đóng Ticket trước khi mở gia hạn   │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
┌───────────────────────────────────▼────────────────────────────────────┐
│ ĐẦU RA (OUTPUT) - 4 KẾT QUẢ ĐỒNG BỘ                                    │
│ 1. Dashboard Phân tích Trực quan (BI KPIs, Area, Donut, Bar, Reps)     │
│ 2. Báo cáo Tổng quan Chiến lược (Executive Summary, SWOT, Roadmap)     │
│ 3. Bộ Slide PowerPoint 10 Trang (Trình chiếu tương tác & tải file)     │
│ 4. File Excel Đã Làm Sạch Kèm Nhận Xét AI Từng Dòng (.xlsx)            │
└────────────────────────────────────────────────────────────────────────┘
```

---

## PHẦN 3: CƠ CHẾ CHẨN ĐOÁN LỖI DỮ LIỆU ĐẦU VÀO (ERROR DIAGNOSTICS)

Hệ thống được trang bị bộ kiểm tra tính toàn vẹn (Data Integrity Guard) ngay tại cửa ngõ nhập liệu:

| Loại Lỗi | Vị Trí Điển Hình | Nguyên Nhân Sai | Cách Khắc Phục (Fix Guide) |
| :--- | :--- | :--- | :--- |
| **Sai định dạng ngày** | Cột `expiry_date` dòng 7 | Giá trị `32/13/2026` không tồn tại ngày 32, tháng 13 | Quy chuẩn về định dạng ISO-8601 `YYYY-MM-DD` (VD: `2026-12-04`) hoặc dùng tính năng Auto-Fix |
| **Giá trị số âm bất hợp lý** | Cột `attended_sessions` dòng 12 | Giá trị âm (`-5 buổi`) vi phạm logic học tập | Nhập số buổi thực tế từ $0$ đến số buổi đăng ký, hoặc để trống (`NULL`) |
| **Khuyết bảng dữ liệu** | File tải lên thiếu Sheet | Tệp Excel thiếu sheet `accounts` hoặc `Sheet1` | Tải file mẫu chuẩn của hệ thống hoặc đặt lại tên sheet đúng quy ước |
| **Xung đột Kỹ thuật** | Bảng `support_tickets` | Đang tồn tại Ticket mở (VD: `TK-9002 - Open` văng lớp) | Kỹ thuật phải sửa dứt điểm lỗi và đóng Ticket, hoặc kích hoạt cờ Tạm hoãn gia hạn |

---

## PHẦN 4: CHI TIẾT 4 KẾT QUẢ ĐẦU RA TRÊN WEB APP LOCALHOST

### Kết quả 1: Dashboard Phân Tích (Interactive BI Dashboard)
- **Hàng 4 chỉ số KPI chính:**
  - *Tổng số lượng tài khoản:* 10 (mẫu chuẩn) hoặc 150 (toàn kỳ tháng).
  - *Doanh thu đủ căn cứ ghi nhận:* **105,6 Tr ₫** (kỳ mẫu) / **1,48 Tỷ ₫** (kỳ 150 học viên) — chiếm 68% tổng kỳ vọng.
  - *Doanh thu tạm giữ (Lỗi Tech/Khuyết log):* **50,0 Tr ₫** — được cô lập an toàn, chưa ghi nhận.
  - *Tỷ lệ chuyên cần trung bình:* **71.8%** — căn cứ vững chắc để nâng cấp hoặc giữ gói.
- **Biểu đồ miền (Area Chart):** Tách lớp rõ rệt giữa *Doanh thu Giữ gói cũ* và *Doanh thu Nâng cấp (Up-Sell)* qua các tháng 8, 9, 10, 11.
- **Biểu đồ tròn (Donut Chart):** Tỷ trọng Doanh thu đủ căn cứ (68%) vs Tạm giữ do lỗi Tech (22%) và Khuyết log (10%).
- **Thẻ Nhân viên Tư vấn (Sales Reps):** Mô phỏng chính xác mẫu báo cáo điều hành, đặc biệt xử lý chuẩn trường hợp `(Blank)` của Tobi Schmidt khi học viên khuyết dữ liệu LMS (`HV-004`).
- **Cổng Phê Duyệt Con Người (HITL):** Nút duyệt 1-click cho Quản lý (duyệt ưu đãi 20% cho `HV-003`) và Kỹ thuật (ký đóng lỗi micro cho `HV-007`), cập nhật số liệu Dashboard tức thì.

### Kết quả 2: Báo Cáo Tổng Quan (Executive Overview Report)
- Tóm tắt đề bài và bối cảnh vận hành EdTech Future.
- Đánh giá chỉ số tài chính, phân loại học viên theo 3 nhóm rủi ro (Thấp, Trung bình, Cao).
- Ma trận phân tích SWOT toàn diện và Lộ trình triển khai 2 giai đoạn cho năm 2037.
- Tích hợp nút in báo cáo và xuất PDF trực tiếp.

### Kết quả 3: Bộ Slide PowerPoint 10 Trang (Interactive Presentation Deck)
Thiết kế chuẩn hóa theo đúng bố cục 10 slide từ template mẫu của người dùng:
1. **Slide 1:** Bìa Báo cáo Đánh giá Hiệu suất & Kế hoạch Năm 2037.
2. **Slide 2:** Tóm tắt 5 chỉ số KPI Hoàn thành Mục tiêu (98%, 85%, 74%, 1,48 Tỷ ₫, 0% lỗi).
3. **Slide 3:** Phân tích 4 Trụ cột Chiến lược Cốt lõi (Tập trung dữ liệu, Tách NULL ≠ 0, Rút ngắn duyệt, Guardrail khóa chặn).
4. **Slide 4:** Báo cáo Cơ cấu Doanh thu (Donut Chart) & 4 Giai đoạn Quy trình (Kết nối, Làm sạch, Phân tích, Báo cáo).
5. **Slide 5:** Tiến độ Kiểm thử 3 Kịch bản Trọng điểm (Chuẩn 100%, Thiếu dữ liệu 100%, Ranh giới quy tắc 100%).
6. **Slide 6:** Tỷ trọng Phân bổ Nguồn lực & Thời gian (40% Sales, 30% Tech, 20% CS, 10% Quản lý).
7. **Slide 7:** Phân tích Ma trận SWOT Hệ thống (S - W - O - T).
8. **Slide 8:** Năng lực Cốt lõi & Giá trị Đội ngũ mang lại (Giá trị Doanh nghiệp, Nghiệp vụ, Công nghệ).
9. **Slide 9:** Kế hoạch Triển khai 6 Tháng Đầu Năm vs 6 Tháng Cuối Năm.
10. **Slide 10:** Lời Cảm Ơn, Thông tin Liên hệ và Đường dẫn Trải nghiệm.

*Tính năng slide trên Web:* Trình chiếu Fullscreen, chuyển slide Trước/Sau, thanh Thumbnail chọn nhanh 10 slide.

### Kết quả 4: File Excel Đã Làm Sạch Kèm Nhận Xét AI (.xlsx)
Bảng dữ liệu khách hàng được làm sạch và bổ sung cột **"Nhận Xét AI & Khuyến Nghị Hành Động"** cho từng dòng:
- `HV-2026-001` (Nguyễn Minh Anh): *"Chuyên cần 91.7% (>=80%), không có sự cố kỹ thuật. Đủ điều kiện nâng cấp gói IELTS Intensive 1 Năm với ưu đãi Early Bird 10%."*
- `HV-2026-004` (Phạm Quốc Bảo): *"[KHUYẾT DỮ LIỆU] Không có log LMS (NULL ≠ 0 buổi). Đang mở Ticket lỗi điểm danh. Tạm giữ doanh thu, không được tự ý gán = 0."*
- `HV-2026-005` (Đỗ Khánh Linh): *"[VẮNG 100%] Tham gia 0/48 buổi. Đề xuất giữ nguyên lộ trình English Foundation kèm khóa phụ đạo 1-1 bổ trợ."*
- `HV-2026-007` (Hoàng Bích Ngọc): *"[KHÓA CHẶN KỸ THUẬT] Đang mở Ticket TK-9002 (lỗi micro văng lớp). Kích hoạt Hard-Block cấm chào nâng cấp cho đến khi Kỹ thuật đóng Ticket."*
- Tích hợp nút tải file `.xlsx` trực tiếp về máy.

---

## PHẦN 5: KẾT QUẢ KIỂM THỬ TRÊN 3 KỊCH BẢN THỰC TẾ

| Kịch Bản | Học Viên Tiêu Biểu | Tình Huống Đầu Vào | Xử Lý Của AI Agent | Kết Quả Đạt Được |
| :--- | :--- | :--- | :--- | :--- |
| **Kịch bản 1: Thông thường** | HV-001, HV-002, HV-008 | Chuyên cần cao $\ge 80\%$, không có sự cố kỹ thuật | Tự động đề xuất nâng cấp gói, tính đúng học phí kèm mã ưu đãi hợp lệ | Doanh thu được ghi nhận đầy đủ, xuất bản nháp tư vấn chuẩn |
| **Kịch bản 2: Thiếu dữ liệu** | HV-004, HV-005, HV-006 | HV-004 khuyết log LMS; HV-005 có log 0 buổi; HV-006 lệch mốc ngày đối chiếu | Phân định tuyệt đối NULL ≠ 0; cô lập doanh thu vào nhóm tạm giữ; gắn cờ yêu cầu đồng bộ | Không bị ngộ nhận học viên bỏ học; bảo toàn tính chính xác |
| **Kịch bản 3: Ranh giới quy tắc** | HV-007, HV-009, HV-010 | HV-007 mở ticket văng lớp; HV-010 chuyên cần 37.5% đòi ưu đãi 10% | Kích hoạt Hard-Block cấm chào nâng cấp; từ chối mã ưu đãi vi phạm chính sách | Bảo vệ uy tín thương hiệu; triệt tiêu thất thoát doanh thu |

---

## PHẦN 6: HƯỚNG DẪN TRẢI NGHIỆM TRỰC TIẾP
1. Truy cập trình duyệt tại: **`http://localhost:8080`**.
2. Khi vừa vào, hệ thống ở **Chế độ chờ sạch (Zero initial data)**.
3. Bấm nút **"⚠️ Thử nghiệm phát hiện lỗi dữ liệu"** để kiểm tra cơ chế chẩn đoán lỗi, chỉ rõ sai ở đâu và hướng dẫn sửa.
4. Bấm thanh **"⚡ DEMO NHANH"** (hoặc bấm nút "🔧 Tự Động Sửa Lỗi") để hệ thống tính toán tức thì.
5. Chuyển đổi linh hoạt giữa 4 Tab kết quả: **Dashboard**, **Báo Cáo**, **Slide PowerPoint (10 Slides)**, và **Excel Đã Làm Sạch**.
