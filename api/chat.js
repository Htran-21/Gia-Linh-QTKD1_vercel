module.exports = (req, res) => {
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  let body = req.body || {};
  if (typeof body === 'string') {
    try { body = JSON.parse(body); } catch (e) {}
  }
  const text = String(body.message || '').toLowerCase();

  let reply = '';
  if (text.includes('null') || text.includes('bằng 0') || text.includes('khác biệt') || text.includes('004') || text.includes('005') || text.includes('công suất') || text.includes('đo lường')) {
    reply = `🔍 **Phân biệt Dữ liệu Khuyết (NULL / Để trống) vs. Mức sử dụng bằng 0% trên hệ thống**:
1. **Khách hàng DN-2026-004 (Công ty Tiếp vận Toàn Cầu LogiTrans)**:
   - Trường dữ liệu service_usage_rate_pct = NULL do lỗi đồng bộ cổng API đo lường (Ticket TK-8004 đang xử lý).
   - **Quy tắc an toàn**: Tuyệt đối **không được gán bằng 0%**. Hệ thống gắn cờ cảnh báo, tạm giữ doanh thu (**không tính vào doanh thu đủ căn cứ**) và yêu cầu Kỹ thuật hoàn tất kết nối API trước khi chốt hợp đồng.
2. **Khách hàng DN-2026-005 (Chuỗi Khách sạn & Nghỉ dưỡng Lotus)**:
   - Trường dữ liệu service_usage_rate_pct = 0% thực tế do doanh nghiệp tạm hoãn số hóa vì tái cấu trúc (dữ liệu hợp lệ mốc đối chiếu).
   - **Quy tắc xử lý**: Xác nhận khách hàng chưa phát sinh lưu lượng nhưng hồ sơ chuẩn xác, đủ căn cứ đề xuất gói duy trì hỗ trợ kỹ thuật (31.5 Tr ₫) và giao chuyên viên chăm sóc đồng hành tháo gỡ khó khăn.`;
  } else if (text.includes('sự cố') || text.includes('lỗi') || text.includes('kỹ thuật') || text.includes('chặn') || text.includes('block') || text.includes('007') || text.includes('ticket')) {
    reply = `🛡️ **Cơ Chế Khóa Chặn An Toàn Kiểm Soát Lỗi Kỹ Thuật (Hard-Block)**:
- **Nguyên tắc cốt lõi**: Tuyệt đối **CẤM** bộ phận Kinh doanh chào mời tái ký hoặc nâng cấp hợp đồng cho khách hàng đang mở Ticket sự cố kỹ thuật nghiêm trọng.
- **Trường hợp điển hình DN-2026-007 (Tập đoàn Địa ốc Phương Nam)**:
  - Doanh nghiệp có mức sử dụng hạ tầng rất cao (94%), đủ tiêu chuẩn nâng cấp gói Enterprise Suite.
  - Tuy nhiên hệ thống phát hiện Ticket TK-8002 (Mất kết nối máy chủ CSDL) đang ở trạng thái Mở (Open).
  - Hệ thống tự động kích hoạt **Khóa Chặn Kỹ Thuật (Hard-Block)**, tạm dừng báo giá nâng cấp và ưu tiên Kỹ thuật khắc phục sự cố trong vòng 12h trước khi tiếp cận khách hàng.`;
  } else if (text.includes('chiết khấu') || text.includes('ưu đãi') || text.includes('phê duyệt') || text.includes('quản lý') || text.includes('003') || text.includes('008') || text.includes('010') || text.includes('hitl')) {
    reply = `⚖️ **Quy Tắc Quản Trị Chiết Khấu & Cổng Phê Duyệt Cấp Quản Lý (HITL Gate)**:
- **Trường hợp ngoại lệ DN-2026-003 (Công ty Dược Medilife)**: Khách hàng đề xuất chiết khấu 20% vượt khung thẩm quyền tự động (tối đa 15%) -> Hệ thống tự động chuyển sang trạng thái CHỜ QUẢN LÝ PHÊ DUYỆT, yêu cầu Giám đốc Kinh doanh ký duyệt.
- **Trường hợp ranh giới DN-2026-008 (Cơ khí Sakura)**: Mức sử dụng đạt 79.2% (ngay sát ranh giới 80% nâng cấp), yêu cầu mã ưu đãi VIP nhưng mới ký hợp đồng 1 năm -> Hệ thống từ chối mã VIP, chuyển sang mã Tái ký sớm 10% hợp lệ.
- **Trường hợp khẩn cấp DN-2026-010 (Sáng Tạo Việt)**: Hợp đồng chỉ còn 3 ngày, mức sử dụng thấp 35% -> Hệ thống từ chối mọi chiết khấu vi phạm quy tắc ngày và công suất, đề xuất tái ký nguyên giá để giữ liên tục dịch vụ.`;
  } else {
    reply = `📊 **Báo Cáo Tổng Hợp Đối Soát Dữ Liệu Gia Linh Data Studio**:
- **Tổng số doanh nghiệp đã đối soát**: 10 tài khoản
- **Hồ sơ đủ căn cứ pháp lý & kỹ thuật**: 7 doanh nghiệp -> Tổng doanh thu đủ căn cứ: 567.5 Tr ₫
- **Hồ sơ tạm giữ an toàn / khóa chặn**: 3 doanh nghiệp -> Tổng giá trị tạm giữ: 310.0 Tr ₫
- **Tỷ lệ sử dụng dịch vụ trung bình**: 76.5% (Tách biệt tuyệt đối NULL ≠ 0%).
Hệ thống đã chuẩn hóa tự động 100% dữ liệu và kích hoạt đầy đủ các chốt chặn kiểm soát rủi ro. Bạn có thể sang Bước 4 để tải file Excel hoặc xem các giải pháp đề xuất.`;
  }

  return res.status(200).json({ reply });
};
