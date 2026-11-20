# TASK 1: PHÂN TÍCH QUY TRÌNH VÀ THIẾT KẾ GIẢI PHÁP AI AGENT (EDTECH FUTURE)

---

## 1. Phân tích Quy trình Hiện tại (As-Is) và Các Vấn đề Bất cập

### 1.1. Diễn tiến quy trình hiện tại (Từ tiếp nhận dữ liệu đến bàn giao Quản lý)
Hàng tháng, để xử lý khoảng **150 tài khoản học viên đến kỳ gia hạn**, quy trình thủ công hiện tại diễn ra qua 5 khâu rời rạc:
1. **Thu thập dữ liệu thủ công**: Nhân viên Chăm sóc học viên (CS) tải danh sách tài khoản sắp hết hạn từ hệ thống CRM, xin file theo dõi chuyên cần từ Học vụ, kiểm tra bảng ticket lỗi kỹ thuật từ bộ phận Tech/Support và tra cứu bảng giá/ưu đãi từ phòng Kinh doanh (Sales).
2. **Đối chiếu bằng mắt và Excel rời rạc**: CS/Sales tự ghép nối (`VLOOKUP` thủ công hoặc dò tay) theo mã học viên giữa các tệp có định dạng ngày tháng và mốc cập nhật khác nhau.
3. **Đánh giá và lên phương án tư vấn cảm tính**: Nhân viên tự ước lượng tiến độ học tập, tự chọn gói gia hạn/nâng cấp và tự áp mã giảm giá theo trí nhớ hoặc ghi chú cá nhân.
4. **Liên hệ học viên thiếu đồng bộ**: Sales/CS gọi điện hoặc nhắn tin mời gia hạn/nâng cấp mà không biết trạng thái xử lý sự cố kỹ thuật thực tế của học viên.
5. **Kiểm duyệt thủ công 100% ở khâu cuối**: Trưởng phòng Kinh doanh phải mở lại từng hồ sơ, đối chiếu lại từng con số chuyên cần, lịch sử lỗi và điều kiện mã giảm giá trước khi ký duyệt.

### 1.2. Bốn (04) vấn đề bất cập cốt lõi và Ảnh hưởng cụ thể

| STT | Vấn đề bất cập trong thực tế | Nguyên nhân gốc rễ | Ảnh hưởng cụ thể đến EdTech Future |
| :--- | :--- | :--- | :--- |
| **1** | **Tư vấn nâng cấp gói học khi tài khoản đang gặp sự cố kỹ thuật chưa giải quyết** | Dữ liệu Ticket của bộ phận Tech/Support không được đồng bộ và không có "cơ chế chặn" (Hard-block) trước khi hồ sơ chuyển sang Sales. | Gây bức xúc lớn cho học viên, làm giảm uy tín thương hiệu, tăng tỷ lệ hủy gia hạn (Churn rate) và phát sinh khiếu nại gay gắt trên mạng xã hội. |
| **2** | **Nhầm lẫn giữa "Ngày hết hạn" với "Ngày cần liên hệ", và giữa "Số buổi đăng ký" với "Số buổi thực học" (hoặc dữ liệu trống)** | Thiếu mốc ngày đối chiếu chuẩn (`as_of_date`), định dạng ngày tháng không đồng nhất; không phân biệt được giá trị rỗng (`NULL` - chưa đồng bộ log) với giá trị `0` (học viên bỏ học). | Gọi điện quá muộn (khi tài khoản đã khóa) hoặc quá sớm; đánh giá sai năng lực/cam kết đầu ra của học viên dẫn đến tư vấn sai lộ trình học tập. |
| **3** | **Áp dụng mã giảm giá và chính sách ưu đãi sai điều kiện quy định** | Chính sách ưu đãi có nhiều điều kiện ràng buộc (hạng thành viên, tỷ lệ chuyên cần tối thiểu, hạn áp dụng, điều kiện gộp mã) nhưng nhân viên tính nhẩm hoặc áp dụng tùy tiện để chốt đơn. | Thất thoát doanh thu trực tiếp cho công ty, hoặc phải hủy/điều chỉnh lại báo giá sau khi đã hứa với học viên, gây mất niềm tin. |
| **4** | **Nút cổ chai (Bottleneck) tại khâu phê duyệt của Trưởng phòng Kinh doanh** | Không có hệ thống kiểm tra quy tắc tự động và nhật ký truy vết (Audit Log), buộc cấp quản lý phải rà soát lại toàn bộ 150 hồ sơ/tháng từ đầu. | Lãng phí hàng chục giờ làm việc của cấp quản lý mỗi tháng, làm chậm tốc độ phản hồi học viên và giảm hiệu suất toàn đội Sales – CS – Tech. |

---

## 2. Quy trình Đề xuất Mới (To-Be) Tích hợp 4 Bước Phân tích Dữ liệu & AI Agent

Quy trình mới chuẩn hóa toàn bộ luồng vận hành theo **4 giai đoạn cốt lõi** (bám sát sơ đồ *Quy trình làm việc phân tích dữ liệu*) kết hợp chặt chẽ giữa **AI Agent** và các bộ phận **CS – Sales – Tech/Support – Trưởng phòng Kinh doanh**:

### Bảng chi tiết Quy trình Đề xuất Mới (Đầu vào – Các bước – Đầu ra – Trách nhiệm)

| Bước quy trình | Chi tiết công việc thực hiện | Đầu vào (Inputs) | Đầu ra (Outputs) | Phụ trách (PIC) |
| :--- | :--- | :--- | :--- | :--- |
| **Bước 1: Nhập Dữ liệu từ các Nguồn (Data Ingestion)** | Tiếp nhận dữ liệu qua **Chatbox thông minh** hoặc trình kết nối trực tiếp từ file `.xlsx`, **Google Sheets** (URL/API), **Lark Base** (App Token/Table). Thiết lập mốc ngày đối chiếu chung (`as_of_date`). | 6 nguồn: (1) Tài khoản & Ngày hết hạn, (2) Nhật ký chuyên cần, (3) Ticket kỹ thuật, (4) Danh mục khóa học, (5) Biểu phí, (6) Chính sách ưu đãi. | Tập dữ liệu thô tập trung (Raw Dataset) gắn nhãn thời gian thu thập và mốc `as_of_date` thống nhất. | **AI Agent (Data Ingestion Skill)** + Nhân viên CS (cung cấp link/file qua Chatbox). |
| **Bước 2: Làm sạch & Chuẩn hóa Dữ liệu (Data Cleaning)** | Chuẩn hóa định dạng ngày (`YYYY-MM-DD`), chuẩn hóa mã tài khoản/khóa học, khử trùng lặp. **Phân tách rõ:** Bản ghi thiếu log chuyên cần (`MISSING_DATA`) và bản ghi có chuyên cần bằng `0` (`ZERO_ATTENDANCE`). | Tập dữ liệu thô từ Bước 1. | Bộ dữ liệu sạch (Validated Dataset) + Danh sách cảnh báo dữ liệu khuyết thiếu/bất thường cần làm rõ. | **AI Agent (Data Cleaning & Validation Skill)**. |
| **Bước 3A: Đối chiếu Nghiệp vụ & Phân loại Trạng thái (Core Analysis)** | Tính số ngày còn lại đến hạn (`days_to_expiry`), xác định **Ngày bắt đầu liên hệ** (`contact_start_date = expiry_date - 30 ngày`). Tính tỷ lệ thực học (`attended / registered`). Kiểm tra Ticket kỹ thuật đang mở (`Open`/`In Progress`). | Bộ dữ liệu sạch + Quy tắc nghiệp vụ chuẩn. | Hồ sơ trạng thái từng học viên: `READY_TO_RENEW`, `READY_TO_UPGRADE`, `BLOCKED_BY_TECH`, `NEEDS_DATA_CLARIFICATION`, `OUT_OF_WINDOW`. | **AI Agent (Cross-Check & Triage Skill)**. |
| **Bước 3B: Đề xuất Gói, Tính Học phí & Soạn thảo Nội dung (Proposal & Drafting)** | Đề xuất **Giữ gói cũ** hoặc **Nâng cấp gói mới** kèm căn cứ; tính học phí gốc, kiểm tra điều kiện ưu đãi hợp lệ, xuất giá sau ưu đãi. Gọi **Gemini/OpenAI API** sinh lý do tư vấn, kế hoạch phối hợp CS-Sales-Tech và bản nháp Email/Tin nhắn. | Hồ sơ trạng thái + Danh mục khóa học + Biểu phí + Chính sách ưu đãi. | Hồ sơ đề xuất chi tiết từng học viên + Bản nháp tin nhắn/email + Log truy vết quy tắc (Audit Log). | **AI Agent (Pricing Engine & LLM GenAI Skill)**. |
| **Bước 4A: Phối hợp Xử lý Sự cố & Phê duyệt (HITL Gate 1 & Gate 2)** | - **Tech/Support**: Xử lý dứt điểm các tài khoản gắn cờ `BLOCKED_BY_TECH`, xác nhận dữ liệu chuyên cần bị khuyết.<br>- **Trưởng phòng KD**: Duyệt nhanh trên Dashboard các hồ sơ đủ điều kiện và xem xét các hồ sơ `PENDING_MANAGER_APPROVAL`. | Danh sách tài khoản bị chặn kỹ thuật, danh sách thiếu dữ liệu, danh sách đề xuất ưu đãi/nâng cấp. | Xác nhận `Tech Resolved` từ bộ phận Kỹ thuật; Quyết định `Approved` / `Rejected` từ Trưởng phòng KD. | **Bộ phận Tech/Support**, **Chăm sóc học viên (CS)** và **Trưởng phòng Kinh doanh**. |
| **Bước 4B: Trực quan hóa Dashboard & Triển khai Tư vấn (Dashboard & Execution)** | Hiển thị Dashboard KPI trực quan (Bar, Line, Pie chart), tổng hợp số lượng tài khoản theo trạng thái và **Tổng giá trị gia hạn có đủ căn cứ**. Sales/CS kiểm tra bản nháp lần cuối và gửi cho học viên. | Hồ sơ đã phê duyệt + Dữ liệu tổng hợp toàn kỳ. | Dashboard tương tác, Báo cáo tổng hợp (`.xlsx`/`.json`/`.md`), Lịch trình chăm sóc & Tin nhắn gửi học viên. | **AI Agent (Dashboard & Export)** + **Nhân viên Sales/CS** (thực thi liên hệ). |

---

## 3. Lựa chọn AI Agent: Các bước Tự động hóa và Lý do Lựa chọn

Hệ thống lựa chọn **4 bước cốt lõi** (tương ứng với 4 kỹ năng nghiệp vụ phối hợp) để AI Agent tự động đảm nhận:

1. **Bước 1 & 2 – Thu thập đa nguồn (`.xlsx`, Google Sheets, Lark Base) và Làm sạch, Chuẩn hóa Dữ liệu**:
   - **Lý do lựa chọn**: Con người rất dễ mắc lỗi khi gộp dữ liệu từ nhiều nền tảng (Excel, Google Sheets, Lark Base) có cấu trúc cột và định dạng thời gian khác nhau. Agent kết hợp công cụ xử lý dữ liệu tự động giúp chuẩn hóa 150+ hồ sơ trong vài giây, tự động phát hiện bản ghi trùng lặp và đặc biệt là **không bao giờ đánh đồng dữ liệu trống (`NULL`) với chuyên cần bằng `0`**.
2. **Bước 3A – Đối chiếu Đa chiều, Xác định Thời điểm Liên hệ & Chặn Sự cố Kỹ thuật (Cross-Check & Risk Gate)**:
   - **Lý do lựa chọn**: Nhân viên thường nhầm lẫn giữa ngày hết hạn và ngày bắt đầu phải gọi tư vấn, hoặc quên kiểm tra bảng Ticket kỹ thuật trước khi gọi mời nâng cấp. Agent thực thi quy tắc đối chiếu 100% không bỏ sót bất kỳ ticket `Open`/`In Progress` nào, tự động gắn cờ `BLOCKED_BY_TECH` để ngăn chặn tuyệt đối việc chào bán sai thời điểm.
3. **Bước 3B (Phần 1) – Đề xuất Gói học & Tính toán Học phí/Ưu đãi Chính xác theo Chính sách (Deterministic Pricing & Policy Engine)**:
   - **Lý do lựa chọn**: Việc kết hợp giữa tỷ lệ thực học (`attended_sessions / registered_sessions`), cam kết đầu ra và điều kiện áp dụng mã ưu đãi đòi hỏi tính toán chính xác tuyệt đối. Sử dụng Tool tính toán tất định (Deterministic Function Tool) được Agent gọi giúp loại bỏ hoàn toàn lỗi tính nhẩm hoặc áp sai mã giảm giá của nhân viên Sales.
4. **Bước 3B (Phần 2) – Khởi tạo Kế hoạch Phối hợp Liên phòng ban & Soạn thảo Bản nháp Cá nhân hóa (LLM Reasoning & Drafting qua Gemini/OpenAI)**:
   - **Lý do lựa chọn**: Mỗi học viên có lịch sử học tập, lý do gián đoạn hoặc vấn đề kỹ thuật khác nhau. Mô hình ngôn ngữ lớn (Gemini/OpenAI) có khả năng tổng hợp ngữ cảnh từ kết quả công cụ (Tool Outputs) để viết bản nháp email/tin nhắn tự nhiên, đúng tông giọng, đồng thời sinh chỉ dẫn phối hợp cụ thể cho từng bộ phận CS – Sales – Tech.

---

## 4. Kiểm soát và Ranh giới (Human-in-the-Loop & Guardrails)

### 4.1. Ba (03) Điểm Con người Kiểm tra và Ra Quyết định (Human-in-the-Loop - HITL)

| Điểm kiểm soát (HITL Gate) | Người chịu trách nhiệm | Điều kiện kích hoạt chuyển giao (Escalation / Handoff Conditions) | Hành động của Con người |
| :--- | :--- | :--- | :--- |
| **HITL Gate 1: Xác minh Kỹ thuật & Dữ liệu khuyết** | **Bộ phận Tech/Support & CS** | 1. Tài khoản có ít nhất 1 ticket kỹ thuật ở trạng thái `Open`, `Pending` hoặc `In Progress`.<br>2. Dữ liệu chuyên cần bị `NULL`/trống hoặc mâu thuẫn (`attended_sessions > registered_sessions`, ngày cập nhật log cũ hơn quy định). | Tech xử lý sự cố và cập nhật trạng thái `Resolved` trên hệ thống; CS xác minh lại nhật ký học tập thực tế trước khi cho phép Agent mở khóa tư vấn gia hạn. |
| **HITL Gate 2: Phê duyệt Chính sách, Ưu đãi & Ngoại lệ** | **Trưởng phòng Kinh doanh (Sales Manager)** | 1. Hồ sơ đề xuất nâng cấp gói hoặc áp dụng mã ưu đãi đặc biệt/ưu đãi có điều kiện cần phê duyệt (`requires_approval = True`).<br>2. Hồ sơ nằm ở ranh giới quy tắc (ví dụ: chuyên cần sát ngưỡng 79%–80%, hoặc hết hạn đúng ngày giới hạn ưu đãi Early Bird). | Xem bảng tóm tắt căn cứ + Audit Log trên Dashboard để nhấn **Phê duyệt (Approve)** hàng loạt hoặc điều chỉnh phương án cho từng hồ sơ ngoại lệ. |
| **HITL Gate 3: Kiểm duyệt Thông điệp Gửi Học viên** | **Nhân viên Sales / CS** | Tất cả các bản nháp Email/Tin nhắn Zalo/SMS do AI Agent soạn thảo. (Agent chỉ tạo trạng thái `DRAFT`, không tự động gửi trực tiếp). | Kiểm tra ngữ điệu, thông tin học phí và xác nhận gửi đi sau khi đã có đủ phê duyệt từ Gate 1 và Gate 2. |

### 4.2. Ranh giới Cấm (Hard Boundaries) của AI Agent
- **Cấm tự ý cam kết khắc phục sự cố kỹ thuật**: Khi ticket chưa có trạng thái `Resolved`/`Closed` trong dữ liệu đầu vào, Agent tuyệt đối không được viết trong tin nhắn cho học viên rằng *"Sự cố của bạn đã được khắc phục xong"*.
- **Cấm tự ý cam kết đã duyệt ưu đãi**: Khi mã giảm giá thuộc diện cần Trưởng phòng KD phê duyệt hoặc chưa đủ điều kiện, Agent chỉ được ghi nhận là *"Dự kiến đề xuất trình phê duyệt"* trong hồ sơ nội bộ và không được báo giá đã giảm như một cam kết chính thức cho học viên khi chưa có cờ `Manager_Approved = True`.
- **Cấm gộp doanh thu chưa đủ căn cứ vào Tổng giá trị gia hạn khả thi**: Các tài khoản đang bị `BLOCKED_BY_TECH` hoặc `MISSING_ATTENDANCE_DATA` phải được tách riêng thành nhóm *"Giá trị chờ xử lý/chưa đủ căn cứ"*, tuyệt đối không cộng dồn vào *"Tổng giá trị gia hạn có đủ căn cứ tính toán"*.

### 4.3. Cách thức Phát hiện và Ngăn chặn Kết quả Sai sót của Agent
1. **Kiến trúc Tách biệt Tính toán và Ngôn ngữ (Neuro-Symbolic Separation)**: Mọi phép tính ngày tháng, tỷ lệ chuyên cần, kiểm tra điều kiện mã giảm giá và tính tiền đều do **Python Rule Engine & Pricing Tool** thực hiện chính xác 100%. LLM (Gemini/OpenAI) bắt buộc gọi Tool (`Function Calling`) và sử dụng đúng con số Tool trả về.
2. **Bộ Kiểm tra Hậu kiểm Tự động (Output Guardrail Validator)**: Tự động quét văn bản của LLM để kiểm tra số tiền có khớp với `final_price` không và chặn mọi từ khóa cam kết sai sự thật khi `open_tickets > 0` hoặc `approval_status == "PENDING"`.
3. **Nhật ký Truy vết Toàn diện (Immutable Execution & Rule Trace Log)**: Mọi bước từ đọc dữ liệu, làm sạch, gọi Tool, áp dụng mã quy tắc (`RULE_ID`), prompt gửi LLM, phản hồi của LLM đều được ghi vào file log `JSONL`.

---

## 5. Bộ Quy tắc Nghiệp vụ Chuẩn hóa (Đề xuất áp dụng cho Task 2)

1. **Mốc thời gian & Cửa sổ liên hệ (`Contact Window Rules`)**:
   - Mốc đối chiếu chung: `as_of_date`.
   - Số ngày còn lại: `days_to_expiry = (expiry_date - as_of_date).days`.
   - Ngày bắt đầu liên hệ chuẩn: `contact_start_date = expiry_date - 30 ngày`.
   - Phân loại độ khẩn cấp liên hệ:
     - `URGENT (Khẩn cấp)`: `0 <= days_to_expiry <= 7`
     - `STANDARD_WINDOW (Đúng kỳ hạn)`: `8 <= days_to_expiry <= 30`
     - `EARLY_PREP (Chuẩn bị sớm)`: `31 <= days_to_expiry <= 45`
     - `EXPIRED_GRACE (Đã quá hạn/Gia hạn muộn)`: `-7 <= days_to_expiry < 0`
2. **Quy tắc Chuyên cần & Phân biệt Dữ liệu thiếu (`Attendance & Data Integrity Rules`)**:
   - Nếu `attended_sessions` bị trống (`NULL`, `NaN`, `""`) hoặc không tìm thấy bản ghi: Gắn cờ `ATTENDANCE_MISSING` (không gán bằng `0`), trạng thái `NEEDS_DATA_CLARIFICATION`.
   - Nếu `attended_sessions == 0`: Gắn cờ `ZERO_ATTENDANCE` (`attendance_rate = 0%`).
   - Tỷ lệ thực học: `attendance_rate = attended_sessions / registered_sessions`.
3. **Quy tắc Chặn Sự cố Kỹ thuật (`Technical Ticket Blocker Rules`)**:
   - Nếu có bất kỳ Ticket nào ở trạng thái `Open`, `In Progress`, `Pending` hoặc `Reopened`: Gắn cờ `BLOCKED_BY_TECH = True`, cấm đề xuất nâng cấp gói cao hơn và không cộng vào tổng giá trị gia hạn đủ căn cứ.
4. **Quy tắc Đề xuất Gói học (`Package Recommendation Rules`)**:
   - **Nâng cấp gói mới (`UPGRADE`)**: `BLOCKED_BY_TECH == False`, đủ dữ liệu, `attendance_rate >= 80%`, đạt cam kết đầu ra.
   - **Giữ gói cũ (`RENEW_SAME`)**: `BLOCKED_BY_TECH == False`, đủ dữ liệu, `50% <= attendance_rate < 80%`.
   - **Giữ gói cũ kèm Lộ trình Phụ đạo (`RENEW_WITH_SUPPORT`)**: `attendance_rate < 50%` (bao gồm cả `0%`).
5. **Quy tắc Tính Học phí & Chính sách Ưu đãi (`Pricing & Discount Policy Rules`)**:
   - `EARLY_BIRD_10` (Giảm 10%): `days_to_expiry >= 15`.
   - `HIGH_ACHIEVER_15` (Giảm 15%): `attendance_rate >= 85%` và không có ticket khiếu nại.
   - `LOYALTY_VIP_20` (Giảm 20%): Dành cho học viên năm thứ 2+ hoặc ngoại lệ, bắt buộc `requires_manager_approval = True`.
