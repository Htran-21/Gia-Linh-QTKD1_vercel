# TASK 2: XÂY DỰNG GIẢI PHÁP KỸ THUẬT, THỰC THI & BÁO CÁO KIỂM THỬ 3 KỊCH BẢN

Toàn bộ kết quả trong báo cáo này được sinh tự động 100% từ việc thực thi hệ thống AI Agent thông qua 4 kỹ năng nghiệp vụ (`Skills/Tools`) trên dữ liệu thực tế tại mốc đối chiếu `as_of_date = 2026-11-16`. Mọi bước đọc dữ liệu, gọi công cụ, kiểm tra quy tắc (`RULE_ID`) và hậu kiểm chống ảo giác (`Guardrails`) đều được lưu vết tại `logs/agent_execution_audit.jsonl` và `logs/agent_execution_trace.log`.

---

## 1. Kiến trúc Giải pháp Kỹ thuật & Tích hợp Kỹ năng Nghiệp vụ trên Localhost

### 1.1. Danh mục Tệp tin Hệ thống đã Xây dựng
- **Giao diện Web Dashboard & Chatbox Đa nguồn (`http://localhost:8080`)**: `web/index.html`
- **Localhost Server & REST API Orchestrator**: `server.ps1` và `start_localhost.bat`
- **Pipeline Thực thi 4 Kỹ năng AI Agent & OpenXML `.xlsx` Engine**: `scripts/run_agent_pipeline.ps1` và `agent_engine.py`
- **Dữ liệu Đầu vào Mẫu Chuẩn (6 Bảng)**:
  - File Excel 6 Sheets: `data/EdTech_Future_Input_Data.xlsx`
  - File JSON cấu trúc: `data/edtech_sample_dataset.json`
- **Kết quả Đầu ra & Nhật ký Truy vết (Audit Logs)**:
  - Báo cáo đề xuất xuất ra Excel: `outputs/EdTech_Renewal_Proposals_Output.xlsx`
  - Dữ liệu phân tích chi tiết JSON: `outputs/renewal_analysis_results.json`
  - Nhật ký truy vết JSONL: `logs/agent_execution_audit.jsonl`

---

## 2. Tổng hợp Kết quả Thực thi Toàn hệ thống (Mốc đối chiếu: `2026-11-16`)

| Nhóm Trạng thái Hồ sơ | Mã Trạng thái (`overall_status`) | Số lượng TK | Giá trị Gia hạn (VNĐ) | Được tính vào "Doanh thu Đủ căn cứ"? | Cổng Kiểm soát (HITL Gate) |
| :--- | :--- | :---: | ---: | :---: | :--- |
| **1. Đủ điều kiện thực thi ngay** | `READY_FOR_EXECUTION` | **5** | **73,400,000 VNĐ** | **CÓ (Đủ căn cứ)** | Gate 3: Sales/CS rà soát & gửi tin |
| **2. Chờ Trưởng phòng KD phê duyệt** | `PENDING_MANAGER_APPROVAL` | **2** | **32,200,000 VNĐ** | **CÓ (Đủ căn cứ)** | Gate 2: Trưởng phòng KD duyệt ưu đãi/ranh giới |
| **CỘNG NHÓM ĐỦ CĂN CỨ TÍNH TOÁN** | **`is_revenue_grounded = True`** | **7 TK** | **105,600,000 VNĐ** | **✅ TỔNG GIÁ TRỊ ĐỦ CĂN CỨ** | — |
| **3. Tạm dừng do Sự cố Kỹ thuật đang mở** | `BLOCKED_BY_TECH` | **2** | *28,000,000 VNĐ (Tạm giữ)* | **KHÔNG (Chưa đủ căn cứ)** | Gate 1: Bộ phận Tech xử lý dứt điểm Ticket |
| **4. Tạm dừng do Lệch/Thiếu Dữ liệu Chuyên cần** | `NEEDS_DATA_CLARIFICATION` | **1** | *22,000,000 VNĐ (Tạm giữ)* | **KHÔNG (Chưa đủ căn cứ)** | Gate 1: CS & Học vụ đồng bộ bổ sung log |
| **CỘNG NHÓM TẠM GIỮ / CHỜ XỬ LÝ** | **`is_revenue_grounded = False`** | **3 TK** | *50,000,000 VNĐ (Tạm giữ)* | **⏸ TÁCH RIÊNG (KHÔNG GỘP)** | — |

---

## 3. Báo cáo Kiểm thử Chi tiết trong 3 Kịch bản Nghiệp vụ

### 3.1. Kịch bản 1: Trường hợp Thông thường (`SCENARIO_1_NORMAL`)
- **`HV-2026-001` (Nguyễn Minh Anh)**: Hết hạn `2026-12-06` (còn 20 ngày, gọi từ `2026-11-06`), chuyên cần `44/48` (91.67%), không có lỗi kỹ thuật -> Đề xuất `UPGRADE` lên `ENG-INTER-02` (16,000,000đ), áp mã `HIGH_ACHIEVER_15` (-15%) -> Học phí: **13,600,000 VNĐ** (`READY_FOR_EXECUTION`).
- **`HV-2026-002` (Trần Hoàng Nam)**: Hết hạn `2026-12-04` (còn 18 ngày, gọi từ `2026-11-04`), chuyên cần `40/60` (66.67%), ticket cũ `TK-9001` đã `Resolved` -> Đề xuất `RENEW_SAME` giữ gói `ENG-INTER-02`, áp mã `EARLY_BIRD_10` (-10%) -> Học phí: **14,400,000 VNĐ** (`READY_FOR_EXECUTION`).
- **`HV-2026-003` (Lê Thu Hà)**: Hết hạn `2026-11-28` (còn 12 ngày, gọi từ `2026-10-29`), chuyên cần `54/60` (90%), học viên năm 2 -> Đề xuất `UPGRADE` lên `ENG-ADV-03` (22,000,000đ), mã `LOYALTY_VIP_20` (-20%, dự kiến **17,600,000 VNĐ**) -> Gắn cờ `PENDING_MANAGER_APPROVAL` (Chờ Trưởng phòng KD duyệt tại Gate 2, bản nháp không cam kết đã duyệt chính thức).

### 3.2. Kịch bản 2: Trường hợp Thiếu hụt / Bất thường Dữ liệu (`SCENARIO_2_MISSING_DATA`)
- **`HV-2026-004` (Phạm Quốc Bảo)**: `attended_sessions = NULL` (khuyết log LMS) và có Ticket `TK-9003` (`In Progress`) -> Hệ thống phân biệt rõ `NULL` với `0`, gắn nhãn `MISSING_ATTENDANCE_DATA` & `BLOCKED_BY_TECH`, không cộng 12,000,000đ vào doanh thu đủ căn cứ.
- **`HV-2026-005` (Đỗ Khánh Linh)**: `attended_sessions = 0` trên 48 buổi (0%) -> Xác nhận `ZERO_ATTENDANCE`, đề xuất `RENEW_WITH_SUPPORT` (12,000,000 VNĐ), yêu cầu CS liên hệ tìm hiểu nguyên nhân vắng học.
- **`HV-2026-006` (Vũ Đức Trọng)**: Bảng chuyên cần có `as_of_date = 2026-10-15` lệch so với mốc chuẩn `2026-11-16` -> Gắn nhãn `STALE_AS_OF_DATE` (`NEEDS_DATA_CLARIFICATION`), tạm giữ 22,000,000đ chờ đồng bộ log tháng 11.

### 3.3. Kịch bản 3: Trường hợp Nằm ở Ranh giới Quy tắc Nghiệp vụ (`SCENARIO_3_BOUNDARY`)
- **`HV-2026-007` (Hoàng Bích Ngọc)**: Chuyên cần 91.67% (`55/60` buổi) nhưng đang có Ticket `TK-9002` (`Open` - lỗi âm thanh phòng học) -> Bị chặn cứng (`BLOCKED_BY_TECH`), cấm chào nâng cấp lúc này, bản nháp tuyệt đối không cam kết đã khắc phục xong sự cố.
- **`HV-2026-008` (Ngô Tuấn Kiệt)**: Chuyên cần `38/48` (79.17% - sát dưới ngưỡng 80%), yêu cầu mã `LOYALTY_VIP_20` dù mới học 1 năm -> Đề xuất `RENEW_SAME` kèm cờ `PENDING_MANAGER_APPROVAL` cho ranh giới chuyên cần; từ chối `LOYALTY_VIP_20` và tự động chuyển sang `EARLY_BIRD_10` (-10% do còn đúng 15 ngày) -> **10,800,000 VNĐ**.
- **`HV-2026-009` (Bùi Phương Thảo)**: Chuyên cần `58/72` (80.56% - vừa qua ngưỡng 80%) và còn đúng **15 ngày** (chạm biên `>= 15 ngày` của `EARLY_BIRD_10`) -> Đề xuất `UPGRADE` lên `ENG-BIZ-04` (28,000,000đ), giảm 10% còn **25,200,000 VNĐ** (`READY_FOR_EXECUTION`).
- **`HV-2026-010` (Đặng Gia Huy)**: Còn 3 ngày hết hạn (`URGENT_7DAYS`), chuyên cần 37.5%, đòi mã `EARLY_BIRD_10` -> Từ chối mã giảm giá do vi phạm cả số ngày (< 15) lẫn chuyên cần (< 50%), giữ gói `ENG-FOUND-01` (**12,000,000 VNĐ**).
