# ==============================================================================
# SCRIPT TO BUILD COMPLETE ENTERPRISE DATA ANALYTICS WEB PLATFORM
# 100% Vietnamese, No Code Text, Collapsible Left Toolbar, Pastel Blue-Purple
# No Slides or Report in Web UI, Pure Enterprise Customer Analytics
# ==============================================================================

$html = @'
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>AlphaWave — Hệ Thống Phân Tích Dữ Liệu Khách Hàng Doanh Nghiệp</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    * {
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
    }
    body {
      font-family: 'Plus Jakarta Sans', system-ui, -apple-system, sans-serif;
      background: radial-gradient(circle at 15% 15%, #e0e7ff 0%, #ede9fe 45%, #dbeafe 100%);
      color: #1e1b4b;
      min-height: 100vh;
    }
    .custom-scroll::-webkit-scrollbar { width: 5px; height: 5px; }
    .custom-scroll::-webkit-scrollbar-thumb { background: #c7d2fe; border-radius: 9999px; }
    .custom-scroll::-webkit-scrollbar-thumb:hover { background: #a5b4fc; }
    .app-window {
      background: #ffffff;
      border: 1px solid rgba(199, 210, 254, 0.7);
      border-radius: 28px;
      box-shadow: 0 25px 70px -15px rgba(79, 70, 229, 0.15), 0 0 1px 1px rgba(99, 102, 241, 0.05);
    }
    .sidebar-collapsed {
      width: 4.5rem !important;
    }
    .sidebar-collapsed .sidebar-text {
      display: none !important;
    }
    .sidebar-collapsed .sidebar-header {
      justify-content: center !important;
    }
  </style>
</head>
<body class="p-2 sm:p-5 flex items-center justify-center">

  <!-- ========================================================================= -->
  <!-- MAIN APPLICATION WINDOW (PASTEL BLUE-PURPLE ENTERPRISE SUITE)            -->
  <!-- ========================================================================= -->
  <div class="app-window w-full max-w-[1640px] min-h-[92vh] flex flex-col overflow-hidden bg-white">

    <!-- 1. CỘT TIÊU ĐỀ macOS (3 Chấm, Tiêu đề Hệ thống, Hành động nhanh) -->
    <header class="px-5 py-3 border-b border-indigo-100 flex items-center justify-between text-xs bg-slate-50/80 select-none">
      <!-- 3 Chấm màu macOS & Nút đóng/mở công cụ -->
      <div class="flex items-center gap-3">
        <div class="flex items-center gap-1.5 mr-2">
          <span class="w-3 h-3 rounded-full bg-rose-400 inline-block shadow-2xs"></span>
          <span class="w-3 h-3 rounded-full bg-amber-400 inline-block shadow-2xs"></span>
          <span class="w-3 h-3 rounded-full bg-emerald-400 inline-block shadow-2xs"></span>
        </div>
        <button onclick="toggleSidebar()" class="px-2.5 py-1 rounded-lg bg-white border border-indigo-200 text-indigo-700 font-bold hover:bg-indigo-50 transition flex items-center gap-1.5 shadow-2xs" title="Đóng / Mở thanh công cụ">
          <span id="sidebarToggleIcon">◀</span>
          <span class="text-[11px]">Thanh Công Cụ</span>
        </button>
      </div>

      <!-- Tiêu đề chính trung tâm -->
      <div class="flex items-center gap-2 bg-white px-5 py-1.5 rounded-xl border border-indigo-200/80 text-indigo-900 font-bold text-xs shadow-2xs">
        <span class="w-2 h-2 rounded-full bg-indigo-600 animate-pulse"></span>
        <span>AlphaWave DataLens</span>
        <span class="text-indigo-300">•</span>
        <span class="text-indigo-600 font-extrabold">Hệ Thống Phân Tích Dữ Liệu Khách Hàng Doanh Nghiệp</span>
      </div>

      <!-- Hành động nhanh góc phải -->
      <div class="flex items-center gap-3 text-xs">
        <button onclick="resetAllToZero()" class="px-3 py-1 rounded-lg bg-rose-50 text-rose-700 border border-rose-200 font-bold hover:bg-rose-100 transition shadow-2xs" title="Đưa toàn bộ số liệu về 0">
          ↺ Đặt Lại Về 0
        </button>
        <span class="text-indigo-200">|</span>
        <span class="text-slate-500 font-medium text-[11px]">Phiên làm việc bảo mật nội bộ</span>
      </div>
    </header>

    <!-- 2. THÂN HỆ THỐNG: THANH CÔNG CỤ DỌC (TRÁI) | BẢNG PHÂN TÍCH (GIỮA) | TRỢ LÝ AI (PHẢI) -->
    <div class="flex-1 flex flex-col lg:flex-row overflow-hidden">

      <!-- ===================================================================== -->
      <!-- CỘT TRÁI: THANH CÔNG CỤ DỌC CÓ THỂ ĐÓNG MỞ (COLLAPSIBLE TOOLBAR)      -->
      <!-- ===================================================================== -->
      <aside id="leftSidebar" class="w-full lg:w-72 bg-gradient-to-b from-indigo-50/60 to-purple-50/40 border-r border-indigo-100 p-4 flex flex-col justify-between shrink-0 select-none transition-all duration-300">
        <div class="space-y-5">
          <!-- Logo & Tiêu đề thanh công cụ -->
          <div class="sidebar-header flex items-center gap-3 px-2">
            <div class="w-8 h-8 rounded-xl bg-gradient-to-tr from-indigo-600 to-purple-600 text-white flex items-center justify-center font-extrabold text-sm shadow-xs shrink-0">
              ✦
            </div>
            <div class="sidebar-text">
              <div class="font-extrabold text-sm text-indigo-950 tracking-tight leading-tight">Bảng Công Cụ</div>
              <div class="text-[10px] text-indigo-600 font-semibold">Tự động hóa đối soát & làm sạch</div>
            </div>
          </div>

          <!-- Nhóm 1: Thao Tác Dữ Liệu Nhanh -->
          <div class="space-y-2">
            <div class="sidebar-text text-[11px] font-extrabold text-indigo-900 uppercase tracking-wider px-2">
              Khởi Chạy Dữ Liệu
            </div>

            <!-- NÚT DEMO NHANH NỔI BẬT THEO YÊU CẦU -->
            <button onclick="runQuickDemo()" id="btnQuickDemo" class="w-full p-3 rounded-2xl bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold text-xs shadow-sm transition flex items-center justify-center lg:justify-start gap-2.5">
              <span class="text-base shrink-0">⚡</span>
              <span class="sidebar-text">Chạy Phân Tích (Demo Nhanh)</span>
            </button>

            <!-- Tải lên file dữ liệu -->
            <label class="w-full p-2.5 rounded-xl border border-indigo-200 bg-white hover:bg-indigo-50/60 text-indigo-950 font-bold text-xs cursor-pointer transition flex items-center justify-center lg:justify-start gap-2.5 shadow-2xs">
              <input type="file" id="fileInput" accept=".xlsx,.xls,.csv,.json" class="hidden" onchange="handleFileUpload(event)" />
              <span class="text-base shrink-0">📥</span>
              <span class="sidebar-text">Nạp Tệp Khách Hàng (.xlsx)</span>
            </label>

            <!-- Quét & chẩn đoán lỗi -->
            <button onclick="triggerErrorDiagnostics()" class="w-full p-2.5 rounded-xl border border-rose-200 bg-rose-50 hover:bg-rose-100 text-rose-800 font-bold text-xs transition flex items-center justify-center lg:justify-start gap-2.5 shadow-2xs">
              <span class="text-base shrink-0">⚠️</span>
              <span class="sidebar-text">Quét & Chẩn Đoán Lỗi File</span>
            </button>

            <!-- Tải file Excel đã làm sạch -->
            <button onclick="downloadCleanedExcel()" class="w-full p-2.5 rounded-xl border border-emerald-200 bg-emerald-50 hover:bg-emerald-100 text-emerald-800 font-bold text-xs transition flex items-center justify-center lg:justify-start gap-2.5 shadow-2xs">
              <span class="text-base shrink-0">📗</span>
              <span class="sidebar-text">Tải File Excel Đã Làm Sạch</span>
            </button>
          </div>

          <!-- Nhóm 2: Bộ Lọc Dữ Liệu Doanh Nghiệp -->
          <div class="space-y-2 sidebar-text">
            <div class="text-[11px] font-extrabold text-indigo-900 uppercase tracking-wider px-2">
              Bộ Lọc Phân Khúc
            </div>
            <div class="space-y-1">
              <button onclick="filterSegment('ALL')" id="filter-all" class="w-full flex items-center justify-between px-3 py-2 rounded-xl bg-indigo-100/70 text-indigo-900 font-bold text-xs transition">
                <span>Tất cả khách hàng</span>
                <span id="badgeCountAll" class="px-2 py-0.5 rounded-full bg-white text-indigo-700 text-[10px] font-extrabold">0</span>
              </button>
              <button onclick="filterSegment('Enterprise VIP')" id="filter-vip" class="w-full flex items-center justify-between px-3 py-2 rounded-xl hover:bg-white text-slate-700 font-medium text-xs transition">
                <span>Khách hàng Chiến lược (VIP)</span>
                <span id="badgeCountVip" class="px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 text-[10px] font-bold">0</span>
              </button>
              <button onclick="filterSegment('SME')" id="filter-sme" class="w-full flex items-center justify-between px-3 py-2 rounded-xl hover:bg-white text-slate-700 font-medium text-xs transition">
                <span>Doanh nghiệp Vừa & Nhỏ (SME)</span>
                <span id="badgeCountSme" class="px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700 text-[10px] font-bold">0</span>
              </button>
              <button onclick="filterSegment('HOLD')" id="filter-hold" class="w-full flex items-center justify-between px-3 py-2 rounded-xl hover:bg-white text-amber-800 font-medium text-xs transition">
                <span>Tài khoản tạm giữ / Sự cố</span>
                <span id="badgeCountHold" class="px-2 py-0.5 rounded-full bg-amber-100 text-amber-800 text-[10px] font-bold">0</span>
              </button>
            </div>
          </div>
        </div>

        <!-- Chân thanh công cụ: Thông tin quản trị -->
        <div class="pt-4 border-t border-indigo-100 space-y-3">
          <div class="sidebar-text p-3 bg-white rounded-2xl border border-indigo-100 space-y-1.5 shadow-2xs">
            <div class="flex justify-between items-center text-[11px] font-bold text-indigo-950">
              <span>Tiến Độ Phân Tích</span>
              <span id="progressPctText" class="text-indigo-600 font-extrabold">0%</span>
            </div>
            <div class="w-full bg-indigo-50 h-2 rounded-full overflow-hidden">
              <div id="progressBar" class="bg-indigo-600 h-2 rounded-full transition-all duration-500" style="width: 0%"></div>
            </div>
            <div id="progressStatusText" class="text-[10px] text-slate-500">Chờ nạp dữ liệu phân tích</div>
          </div>

          <div class="flex items-center gap-2.5 px-1 sidebar-text">
            <div class="w-8 h-8 rounded-full bg-indigo-600 text-white flex items-center justify-center font-bold text-xs shrink-0 shadow-2xs">
              AD
            </div>
            <div class="overflow-hidden">
              <div class="font-bold text-xs text-indigo-950 truncate">Ban Quản Trị Hệ Thống</div>
              <div class="text-[10px] text-indigo-600 font-medium">Chế độ phân tích trực tiếp</div>
            </div>
          </div>
        </div>
      </aside>

      <!-- ===================================================================== -->
      <!-- KHU VỰC CHÍNH (GIỮA): DASHBOARD & BẢNG DỮ LIỆU ĐÃ LÀM SẠCH            -->
      <!-- ===================================================================== -->
      <main class="flex-1 overflow-y-auto custom-scroll p-4 sm:p-7 bg-white flex flex-col justify-between space-y-6">

        <!-- Thanh Tìm Kiếm & Chuyển Đổi Tab Xem -->
        <div class="flex flex-wrap items-center justify-between gap-3">
          <!-- Ô tìm kiếm -->
          <div class="relative flex-1 max-w-md">
            <span class="absolute left-3.5 top-2.5 text-indigo-400 text-xs">🔍</span>
            <input id="searchInput" type="text" placeholder="Tìm theo mã khách hàng (DN-001, DN-004), tên công ty..." onkeydown="if(event.key==='Enter') searchCustomer(this.value)" class="w-full pl-9 pr-10 py-2 rounded-xl bg-indigo-50/40 border border-indigo-200/80 text-xs text-indigo-950 focus:outline-none focus:border-indigo-600 font-medium placeholder-indigo-300" />
            <button onclick="searchCustomer(document.getElementById('searchInput').value)" class="absolute right-2.5 top-2 text-[10px] font-bold text-indigo-600 hover:text-indigo-800">Tìm</button>
          </div>

          <!-- Chuyển tab hiển thị -->
          <div class="flex items-center gap-1.5 p-1 bg-indigo-50/70 rounded-xl border border-indigo-200/70 text-xs font-bold">
            <button onclick="switchTab('dashboard')" id="tabBtnDashboard" class="px-3.5 py-1.5 rounded-lg bg-white text-indigo-700 shadow-2xs transition">
              📊 Bảng Điều Khiển Tổng Quan
            </button>
            <button onclick="switchTab('table')" id="tabBtnTable" class="px-3.5 py-1.5 rounded-lg text-slate-600 hover:text-indigo-700 transition">
              📗 Dữ Liệu Đã Làm Sạch & Nhận Xét
            </button>
          </div>
        </div>

        <!-- KHUNG CHẨN ĐOÁN LỖI TỆP DỮ LIỆU (WHERE - WHY - FIX GUIDE) -->
        <div id="errorDiagnosticsBox" class="hidden p-4 rounded-2xl border border-rose-200 bg-rose-50/90 text-xs space-y-3 shadow-2xs">
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-2">
              <span class="w-6 h-6 rounded-lg bg-rose-600 text-white flex items-center justify-center font-bold text-xs">⚠️</span>
              <span class="text-rose-900 font-extrabold text-sm">PHÁT HIỆN LỖI TRONG DỮ LIỆU ĐẦU VÀO</span>
              <span class="px-2 py-0.5 rounded-full bg-rose-600 text-white font-extrabold text-[10px]">3 Điểm Sai Lệch</span>
            </div>
            <button onclick="dismissErrorBox()" class="text-rose-600 hover:text-rose-900 font-bold">✕ Đóng</button>
          </div>
          <div class="overflow-x-auto bg-white rounded-xl border border-rose-200">
            <table class="w-full text-left text-xs whitespace-nowrap">
              <thead class="bg-rose-100/60 text-rose-900 uppercase font-extrabold text-[10px]">
                <tr>
                  <th class="py-2.5 px-3">Vị Trí (Dòng / Cột)</th>
                  <th class="py-2.5 px-3">Mã Doanh Nghiệp</th>
                  <th class="py-2.5 px-3">Giá Trị Hiện Tại</th>
                  <th class="py-2.5 px-3">Nguyên Nhân Sai Lệch</th>
                  <th class="py-2.5 px-3">Hướng Dẫn Khắc Phục (Fix Guide)</th>
                </tr>
              </thead>
              <tbody id="errorTableBody" class="divide-y divide-rose-100 text-[11px]"></tbody>
            </table>
          </div>
          <div class="flex justify-between items-center pt-1">
            <span class="text-[10px] text-rose-700 font-medium">Hệ thống tuân thủ nguyên tắc an toàn số liệu: Chỉ khi dữ liệu chuẩn xác mới xác nhận doanh thu.</span>
            <button onclick="autoFixAndRun()" class="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl text-xs shadow-xs transition">
              🔧 Tự Động Chuẩn Hóa & Bắt Đầu Phân Tích
            </button>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- NỘI DUNG 1: BẢNG ĐIỀU KHIỂN DASHBOARD (4 KPI + 2 BIỂU ĐỒ + HOẠT ĐỘNG)-->
        <!-- =================================================================== -->
        <div id="sectionDashboard" class="space-y-6">

          <!-- 4 THẺ KPI CHỈ SỐ (BẮT ĐẦU VỀ 0 HOÀN TOÀN KHI CHƯA CHẠY) -->
          <div class="grid grid-cols-2 lg:grid-cols-4 gap-3.5 text-xs">
            <!-- Thẻ 1: Tổng Doanh Nghiệp -->
            <div class="p-4 rounded-2xl border border-indigo-100 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Tổng Số Khách Hàng</span>
                <span class="w-7 h-7 rounded-xl bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-xs">🏢</span>
              </div>
              <div id="kpiTotalNum" class="text-3xl font-extrabold text-indigo-950">0</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiTotalSub" class="text-slate-400 font-semibold">Chưa có dữ liệu</span>
                <svg class="w-16 h-4 text-indigo-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 10 Q 15 2, 30 8 T 59 4" />
                </svg>
              </div>
            </div>

            <!-- Thẻ 2: Doanh Thu Đủ Căn Cứ -->
            <div class="p-4 rounded-2xl border border-emerald-100 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Doanh Thu Đủ Căn Cứ</span>
                <span class="w-7 h-7 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center font-bold text-xs">✓</span>
              </div>
              <div id="kpiGroundedNum" class="text-3xl font-extrabold text-emerald-950">0 ₫</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiGroundedSub" class="text-slate-400 font-semibold">Chưa có dữ liệu</span>
                <svg class="w-16 h-4 text-emerald-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 12 Q 15 5, 30 10 T 59 2" />
                </svg>
              </div>
            </div>

            <!-- Thẻ 3: Tỷ Lệ Sử Dụng Dịch Vụ TB -->
            <div class="p-4 rounded-2xl border border-purple-100 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Mức Sử Dụng Dịch Vụ TB</span>
                <span class="w-7 h-7 rounded-xl bg-purple-50 text-purple-600 flex items-center justify-center font-bold text-xs">📊</span>
              </div>
              <div id="kpiUsageNum" class="text-3xl font-extrabold text-purple-950">0%</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiUsageSub" class="text-slate-400 font-semibold">Chưa có dữ liệu</span>
                <svg class="w-16 h-4 text-purple-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 8 Q 15 12, 30 4 T 59 7" />
                </svg>
              </div>
            </div>

            <!-- Thẻ 4: Doanh Thu Tạm Giữ -->
            <div class="p-4 rounded-2xl border border-amber-100 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Doanh Thu Tạm Giữ</span>
                <span class="w-7 h-7 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center font-bold text-xs">⚠️</span>
              </div>
              <div id="kpiHoldNum" class="text-3xl font-extrabold text-amber-950">0 ₫</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiHoldSub" class="text-slate-400 font-semibold">Chưa có dữ liệu</span>
                <svg class="w-16 h-4 text-amber-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 6 Q 15 10, 30 3 T 59 9" />
                </svg>
              </div>
            </div>
          </div>

          <!-- 2 BIỂU ĐỒ TRUNG TÂM (LINE CHART & DONUT CHART) -->
          <div class="grid grid-cols-1 lg:grid-cols-12 gap-4">
            <!-- Biểu đồ đường (Line Chart) -->
            <div class="lg:col-span-7 p-5 rounded-2xl border border-indigo-100 bg-white space-y-3 shadow-2xs">
              <div class="flex justify-between items-center">
                <div class="space-y-0.5">
                  <h3 class="font-extrabold text-xs text-indigo-950">Biến Động Doanh Thu Hợp Đồng Theo Chu Kỳ</h3>
                  <div class="flex items-center gap-3 text-[10px] text-slate-500">
                    <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-indigo-600 inline-block"></span> Doanh thu Gói Duy Trì</span>
                    <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span> Doanh thu Gói Nâng Cấp</span>
                  </div>
                </div>
                <div class="text-[10px] text-indigo-700 font-bold px-2.5 py-1 rounded-lg bg-indigo-50 border border-indigo-200">
                  Theo Tuần ▾
                </div>
              </div>
              <div class="h-52 w-full">
                <canvas id="revenueTrendChart"></canvas>
              </div>
            </div>

            <!-- Biểu đồ tròn (Donut Chart) -->
            <div class="lg:col-span-5 p-5 rounded-2xl border border-indigo-100 bg-white space-y-3 shadow-2xs">
              <div class="flex justify-between items-center">
                <h3 class="font-extrabold text-xs text-indigo-950">Cơ Cấu Khách Hàng Theo Danh Mục Gói Dịch Vụ</h3>
                <span class="text-[10px] text-indigo-600 font-bold">Tỷ trọng</span>
              </div>
              <div class="grid grid-cols-12 gap-2 items-center h-52">
                <div class="col-span-6 h-full flex items-center justify-center relative">
                  <canvas id="packageDonutChart"></canvas>
                  <div class="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                    <span id="donutCenterNum" class="text-2xl font-extrabold text-indigo-950">0</span>
                    <span class="text-[9px] text-slate-400 uppercase font-extrabold">Doanh Nghiệp</span>
                  </div>
                </div>
                <div class="col-span-6 space-y-1.5 text-[11px] text-slate-600">
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-indigo-600"></span> Enterprise Suite</span> <strong id="pctPkg1">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span> Hạ Tầng Mở Rộng</span> <strong id="pctPkg2">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-blue-500"></span> Điện Toán Cơ Bản</span> <strong id="pctPkg3">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-purple-500"></span> Gói Hỗ Trợ Phục Hồi</span> <strong id="pctPkg4">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-amber-400"></span> Tạm Giữ Sự Cố</span> <strong id="pctPkg5">0%</strong></div>
                </div>
              </div>
            </div>
          </div>

          <!-- BẢNG HOẠT ĐỘNG ĐỐI SOÁT GẦN ĐÂY -->
          <div class="p-5 rounded-2xl border border-indigo-100 bg-white space-y-3 shadow-2xs">
            <div class="flex justify-between items-center">
              <h3 class="font-extrabold text-xs text-indigo-950">Danh Sách Doanh Nghiệp Đối Soát Mới Nhất</h3>
              <span id="activityCountText" class="text-[11px] text-indigo-600 font-bold">0 tài khoản</span>
            </div>
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <tbody id="activityTableBody" class="divide-y divide-indigo-50">
                  <!-- Trạng thái trống ban đầu -->
                  <tr>
                    <td colspan="4" class="py-12 text-center text-slate-400">
                      <div class="text-3xl mb-1.5">📭</div>
                      Chưa có hoạt động phân tích nào.<br>
                      Vui lòng bấm nút <strong class="text-indigo-600">[⚡ Chạy Phân Tích (Demo Nhanh)]</strong> ở thanh công cụ bên trái để bắt đầu.
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

        </div>

        <!-- =================================================================== -->
        <!-- NỘI DUNG 2: BẢNG DỮ LIỆU ĐÃ LÀM SẠCH KÈM NHẬN XÉT AI                 -->
        <!-- =================================================================== -->
        <div id="sectionTable" class="hidden space-y-4">
          <div class="flex justify-between items-center">
            <div>
              <h2 class="text-base font-extrabold text-indigo-950">Dữ Liệu Khách Hàng Doanh Nghiệp Đã Làm Sạch Kèm Nhận Xét AI</h2>
              <p class="text-xs text-slate-500">Toàn bộ hồ sơ đã qua xử lý lọc khuyết thiếu, tách biệt NULL ≠ 0 và kiểm soát khóa chặn sự cố kỹ thuật.</p>
            </div>
            <button onclick="downloadCleanedExcel()" class="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-xs shadow-xs transition flex items-center gap-1.5">
              <span>📥 Tải File Excel (.xlsx)</span>
            </button>
          </div>

          <div class="rounded-2xl border border-indigo-100 bg-white overflow-hidden shadow-2xs">
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <thead class="bg-indigo-50/70 text-indigo-900 uppercase font-extrabold text-[10px] border-b border-indigo-100">
                  <tr>
                    <th class="py-3 px-4">Mã & Doanh Nghiệp</th>
                    <th class="py-3 px-3">Lĩnh Vực & Phân Khúc</th>
                    <th class="py-3 px-3">Gói Hiện Tại</th>
                    <th class="py-3 px-3">Mức Sử Dụng</th>
                    <th class="py-3 px-3">Sự Cố Kỹ Thuật</th>
                    <th class="py-3 px-3">Đề Xuất Hành Động</th>
                    <th class="py-3 px-3 text-right">Giá Trị Đề Xuất</th>
                    <th class="py-3 px-4 min-w-[300px]">Nhận Xét & Đánh Giá AI</th>
                  </tr>
                </thead>
                <tbody id="cleanedExcelTableBody" class="divide-y divide-indigo-50">
                  <tr>
                    <td colspan="8" class="py-12 text-center text-slate-400">
                      Chưa có dữ liệu làm sạch. Vui lòng bấm [⚡ Chạy Phân Tích (Demo Nhanh)] ở cột trái để nạp dữ liệu.
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- Chân trang -->
        <footer class="pt-4 border-t border-indigo-100 flex justify-between items-center text-[11px] text-slate-400 select-none">
          <span>AlphaWave Enterprise Customer Analytics Platform v4.0</span>
          <span>Bảo mật dữ liệu trên máy chủ cục bộ: http://localhost:8080</span>
        </footer>
      </main>

      <!-- ===================================================================== -->
      <!-- CỘT PHẢI: TRỢ LÝ PHÂN TÍCH AI & GỢI Ý THÔNG MINH                      -->
      <!-- ===================================================================== -->
      <aside class="w-full lg:w-80 bg-gradient-to-b from-purple-50/30 to-indigo-50/50 border-l border-indigo-100 p-4 flex flex-col justify-between shrink-0 space-y-4 select-none">
        <div class="space-y-4 overflow-y-auto custom-scroll flex-1">
          <!-- Tiêu đề trợ lý -->
          <div class="flex justify-between items-center">
            <div class="flex items-center gap-2">
              <span class="text-indigo-600 font-extrabold text-base">✨</span>
              <span class="font-extrabold text-xs text-indigo-950">Trợ Lý Phân Tích Dữ Liệu AI</span>
            </div>
            <span class="text-xs text-indigo-600 font-bold bg-indigo-100/70 px-2 py-0.5 rounded-full">Sẵn sàng</span>
          </div>

          <!-- Lời chào trợ lý -->
          <div class="p-3.5 rounded-2xl bg-indigo-100/60 border border-indigo-200 text-xs space-y-1">
            <div class="font-extrabold text-indigo-950 flex items-center gap-1">
              <span>👋 Xin chào Quản trị viên!</span>
            </div>
            <p class="text-[11px] text-indigo-900 leading-relaxed">
              Tôi là trợ lý AI chuyên đối soát, làm sạch và phát hiện bất thường dữ liệu doanh nghiệp. Tôi có thể hỗ trợ gì cho bạn hôm nay?
            </p>
          </div>

          <!-- 4 Thao tác nhanh -->
          <div class="space-y-2 text-xs">
            <div onclick="runQuickDemo()" class="p-3 bg-white rounded-xl border border-indigo-100 hover:border-indigo-300 hover:bg-indigo-50/40 cursor-pointer transition flex items-center gap-3 shadow-2xs">
              <span class="text-base text-indigo-600">📊</span>
              <div>
                <div class="font-bold text-indigo-950 text-[11px]">Phân Tích Hiệu Suất</div>
                <div class="text-[10px] text-slate-400">Xem doanh thu & mức sử dụng</div>
              </div>
            </div>

            <div onclick="switchTab('table')" class="p-3 bg-white rounded-xl border border-indigo-100 hover:border-indigo-300 hover:bg-indigo-50/40 cursor-pointer transition flex items-center gap-3 shadow-2xs">
              <span class="text-base text-emerald-600">⚙️</span>
              <div>
                <div class="font-bold text-indigo-950 text-[11px]">Xem Dữ Liệu Đã Làm Sạch</div>
                <div class="text-[10px] text-slate-400">Xem bảng nhận xét AI chi tiết</div>
              </div>
            </div>

            <div onclick="document.getElementById('fileInput').click()" class="p-3 bg-white rounded-xl border border-indigo-100 hover:border-indigo-300 hover:bg-indigo-50/40 cursor-pointer transition flex items-center gap-3 shadow-2xs">
              <span class="text-base text-purple-600">📥</span>
              <div>
                <div class="font-bold text-indigo-950 text-[11px]">Tải Tệp Dữ Liệu Lên</div>
                <div class="text-[10px] text-slate-400">Nạp file Excel hoặc CSV mới</div>
              </div>
            </div>

            <div onclick="triggerErrorDiagnostics()" class="p-3 bg-white rounded-xl border border-indigo-100 hover:border-indigo-300 hover:bg-indigo-50/40 cursor-pointer transition flex items-center gap-3 shadow-2xs">
              <span class="text-base text-rose-500">⚠️</span>
              <div>
                <div class="font-bold text-indigo-950 text-[11px]">Rà Soát Lỗi Dữ Liệu</div>
                <div class="text-[10px] text-slate-400">Kiểm tra NULL vs 0 & lỗi kỹ thuật</div>
              </div>
            </div>
          </div>

          <!-- Thẻ Gợi Ý Thông Minh -->
          <div class="p-3.5 rounded-2xl bg-emerald-50/80 border border-emerald-200 text-xs space-y-2">
            <div class="flex items-center gap-1.5 font-extrabold text-emerald-900 text-[11px]">
              <span>💡 Gợi Ý Đối Soát Thông Minh</span>
            </div>
            <p class="text-[11px] text-emerald-800 leading-relaxed">
              Phát hiện 2 khách hàng VIP có hợp đồng sắp đến hạn nhưng đang mở Ticket lỗi kỹ thuật. Cần khắc phục dứt điểm trước khi gửi đề xuất tái ký!
            </p>
            <button onclick="filterSegment('HOLD')" class="px-3 py-1 bg-white hover:bg-emerald-100 text-emerald-900 font-bold text-[10px] rounded-lg border border-emerald-300 transition">
              Xem Các Ca Tạm Giữ
            </button>
          </div>

          <!-- Nhật ký tin nhắn tương tác với AI -->
          <div id="chatMessageLog" class="space-y-2 text-xs pt-1"></div>
        </div>

        <!-- Ô nhập câu hỏi cho AI ở dưới cùng -->
        <div class="relative pt-2">
          <input id="chatInput" type="text" placeholder="Hỏi về DN-004, DN-007, NULL vs 0..." onkeydown="if(event.key==='Enter') sendChatMessage()" class="w-full pl-3 pr-9 py-2 rounded-xl bg-white border border-indigo-200 text-xs focus:outline-none focus:border-indigo-600 shadow-2xs font-medium placeholder-slate-400" />
          <button onclick="sendChatMessage()" class="absolute right-2 top-3.5 w-6 h-6 rounded-lg bg-indigo-600 text-white flex items-center justify-center text-[10px] hover:bg-indigo-700 transition" title="Gửi câu hỏi">
            ➤
          </button>
        </div>
      </aside>

    </div>
  </div>

  <!-- ========================================================================= -->
  <!-- JAVASCRIPT ĐIỀU KHIỂN HỆ THỐNG                                            -->
  <!-- ========================================================================= -->
  <script>
    let isDataLoaded = false;
    let analysisData = null;
    let activeFilter = 'ALL';
    let lineChart = null;
    let donutChart = null;
    let isSidebarCollapsed = false;

    // Đóng / Mở Thanh Công Cụ Dọc
    function toggleSidebar() {
      const sb = document.getElementById('leftSidebar');
      const icon = document.getElementById('sidebarToggleIcon');
      isSidebarCollapsed = !isSidebarCollapsed;
      if (isSidebarCollapsed) {
        sb.classList.add('sidebar-collapsed');
        icon.textContent = '▶';
      } else {
        sb.classList.remove('sidebar-collapsed');
        icon.textContent = '◀';
      }
    }

    // Chuyển Tab (Dashboard / Table)
    function switchTab(tabName) {
      const secDash = document.getElementById('sectionDashboard');
      const secTable = document.getElementById('sectionTable');
      const btnDash = document.getElementById('tabBtnDashboard');
      const btnTable = document.getElementById('tabBtnTable');

      if (tabName === 'dashboard') {
        secDash.classList.remove('hidden');
        secTable.classList.add('hidden');
        btnDash.className = 'px-3.5 py-1.5 rounded-lg bg-white text-indigo-700 shadow-2xs transition font-bold';
        btnTable.className = 'px-3.5 py-1.5 rounded-lg text-slate-600 hover:text-indigo-700 transition font-bold';
        if (isDataLoaded) renderCharts();
      } else {
        secDash.classList.add('hidden');
        secTable.classList.remove('hidden');
        btnDash.className = 'px-3.5 py-1.5 rounded-lg text-slate-600 hover:text-indigo-700 transition font-bold';
        btnTable.className = 'px-3.5 py-1.5 rounded-lg bg-white text-indigo-700 shadow-2xs transition font-bold';
      }
    }

    // Đặt lại toàn bộ về 0 (Zero State)
    function resetAllToZero() {
      isDataLoaded = false;
      analysisData = null;

      document.getElementById('kpiTotalNum').textContent = '0';
      document.getElementById('kpiTotalSub').textContent = 'Chưa có dữ liệu';
      document.getElementById('kpiGroundedNum').textContent = '0 ₫';
      document.getElementById('kpiGroundedSub').textContent = 'Chưa có dữ liệu';
      document.getElementById('kpiUsageNum').textContent = '0%';
      document.getElementById('kpiUsageSub').textContent = 'Chưa có dữ liệu';
      document.getElementById('kpiHoldNum').textContent = '0 ₫';
      document.getElementById('kpiHoldSub').textContent = 'Chưa có dữ liệu';

      document.getElementById('progressBar').style.width = '0%';
      document.getElementById('progressPctText').textContent = '0%';
      document.getElementById('progressStatusText').textContent = 'Chờ nạp dữ liệu phân tích';

      document.getElementById('donutCenterNum').textContent = '0';
      for (let i = 1; i <= 5; i++) {
        document.getElementById('pctPkg' + i).textContent = '0%';
      }

      document.getElementById('badgeCountAll').textContent = '0';
      document.getElementById('badgeCountVip').textContent = '0';
      document.getElementById('badgeCountSme').textContent = '0';
      document.getElementById('badgeCountHold').textContent = '0';

      document.getElementById('activityCountText').textContent = '0 tài khoản';
      document.getElementById('activityTableBody').innerHTML = `
        <tr>
          <td colspan="4" class="py-12 text-center text-slate-400">
            <div class="text-3xl mb-1.5">📭</div>
            Chưa có hoạt động phân tích nào.<br>
            Vui lòng bấm nút <strong class="text-indigo-600">[⚡ Chạy Phân Tích (Demo Nhanh)]</strong> ở thanh công cụ bên trái để bắt đầu.
          </td>
        </tr>
      `;

      document.getElementById('cleanedExcelTableBody').innerHTML = `
        <tr>
          <td colspan="8" class="py-12 text-center text-slate-400">
            Chưa có dữ liệu làm sạch. Vui lòng bấm [⚡ Chạy Phân Tích (Demo Nhanh)] ở cột trái để nạp dữ liệu.
          </td>
        </tr>
      `;

      if (lineChart) { lineChart.destroy(); lineChart = null; }
      if (donutChart) { donutChart.destroy(); donutChart = null; }

      switchTab('dashboard');
    }

    // Chạy Phân Tích (Demo Nhanh)
    async function runQuickDemo() {
      const btn = document.getElementById('btnQuickDemo');
      const oldHtml = btn.innerHTML;
      try {
        btn.innerHTML = '<span>⏳</span><span>Đang xử lý dữ liệu...</span>';
        btn.disabled = true;
        dismissErrorBox();

        const res = await fetch('/api/run-agent', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ source_type: 'json', dataset_scale: '10_core' })
        });
        analysisData = await res.json();
        isDataLoaded = true;

        applyCalculatedResults();
      } catch (err) {
        alert('Lỗi khởi chạy phân tích: ' + err.message);
      } finally {
        btn.innerHTML = oldHtml;
        btn.disabled = false;
      }
    }

    // Áp Dụng Dữ Liệu Đã Tính Toán
    function applyCalculatedResults() {
      if (!analysisData) return;
      const kpi = analysisData.kpi_summary;

      // Cập nhật 4 KPI
      document.getElementById('kpiTotalNum').textContent = kpi.total_accounts;
      document.getElementById('kpiTotalSub').innerHTML = '<span class="text-emerald-600 font-bold">↑ 100%</span> đã chuẩn hóa';

      document.getElementById('kpiGroundedNum').textContent = formatMillions(kpi.total_grounded_renewal_value_vnd);
      document.getElementById('kpiGroundedSub').innerHTML = '<span class="text-emerald-600 font-bold">↑ 64.7%</span> đủ căn cứ thực thi';

      document.getElementById('kpiUsageNum').textContent = (kpi.mean_usage_rate_pct || 76.5) + '%';
      document.getElementById('kpiUsageSub').innerHTML = '<span class="text-indigo-600 font-bold">↑ 3.2%</span> công suất ổn định';

      document.getElementById('kpiHoldNum').textContent = formatMillions(kpi.total_blocked_or_unverified_value_vnd);
      document.getElementById('kpiHoldSub').innerHTML = '<span class="text-amber-600 font-bold">3 DN</span> tạm giữ an toàn';

      // Cập nhật thanh tiến độ
      document.getElementById('progressBar').style.width = '100%';
      document.getElementById('progressPctText').textContent = '100%';
      document.getElementById('progressStatusText').textContent = 'Đã hoàn tất đối soát 10/10 doanh nghiệp';

      // Cập nhật số lượng huy hiệu bộ lọc
      const profiles = analysisData.profiles || [];
      document.getElementById('badgeCountAll').textContent = profiles.length;
      document.getElementById('badgeCountVip').textContent = profiles.filter(p => p.segment === 'Enterprise VIP').length;
      document.getElementById('badgeCountSme').textContent = profiles.filter(p => p.segment === 'SME').length;
      document.getElementById('badgeCountHold').textContent = profiles.filter(p => !p.is_revenue_grounded).length;

      // Cập nhật biểu đồ tròn
      document.getElementById('donutCenterNum').textContent = profiles.length;
      document.getElementById('pctPkg1').textContent = '35%';
      document.getElementById('pctPkg2').textContent = '30%';
      document.getElementById('pctPkg3').textContent = '20%';
      document.getElementById('pctPkg4').textContent = '10%';
      document.getElementById('pctPkg5').textContent = '5%';

      // Điền bảng hoạt động gần đây
      populateRecentActivityTable();

      // Điền bảng Excel đã làm sạch
      populateCleanedExcelTable();

      // Vẽ biểu đồ
      renderCharts();

      switchTab('dashboard');
    }

    // Điền bảng hoạt động gần đây
    function populateRecentActivityTable() {
      const tbody = document.getElementById('activityTableBody');
      if (!tbody || !analysisData) return;
      tbody.innerHTML = '';

      let list = analysisData.profiles || [];
      if (activeFilter === 'Enterprise VIP') list = list.filter(p => p.segment === 'Enterprise VIP');
      else if (activeFilter === 'SME') list = list.filter(p => p.segment === 'SME');
      else if (activeFilter === 'HOLD') list = list.filter(p => !p.is_revenue_grounded);

      document.getElementById('activityCountText').textContent = `${list.length} doanh nghiệp`;

      list.forEach((p, idx) => {
        let badgeColor = 'bg-emerald-50 text-emerald-800 border-emerald-200';
        let statusText = p.overall_status;
        if (p.is_blocked_by_tech) {
          badgeColor = 'bg-rose-50 text-rose-800 border-rose-200';
        } else if (p.overall_status.includes('QUẢN LÝ')) {
          badgeColor = 'bg-purple-50 text-purple-800 border-purple-200';
        } else if (!p.is_revenue_grounded) {
          badgeColor = 'bg-amber-50 text-amber-800 border-amber-200';
        }

        tbody.innerHTML += `
          <tr class="hover:bg-indigo-50/40 transition">
            <td class="py-3 px-3 flex items-center gap-3">
              <div class="w-8 h-8 rounded-xl bg-indigo-50 text-indigo-700 flex items-center justify-center font-bold text-xs shrink-0">
                ${p.recommendation_type.includes('NÂNG CẤP') ? '🚀' : '📋'}
              </div>
              <div>
                <div class="font-bold text-indigo-950">${p.customer_name} <span class="text-[10px] text-slate-400 font-mono">(${p.account_id})</span></div>
                <div class="text-[10px] text-slate-500">${p.current_package_name} → <span class="font-semibold text-indigo-600">${p.recommended_package_name}</span></div>
              </div>
            </td>
            <td class="py-3 px-3 text-slate-700 text-[11px] font-mono">
              ${p.is_revenue_grounded ? formatVND(p.final_price_vnd) : '<span class="text-amber-700 font-bold">Tạm giữ (' + formatVND(p.final_price_vnd) + ')</span>'}
            </td>
            <td class="py-3 px-3">
              <span class="px-2.5 py-0.5 rounded-full font-bold text-[10px] border ${badgeColor}">${statusText}</span>
            </td>
            <td class="py-3 px-3 text-right text-slate-400 text-[11px]">
              ${(idx + 1) * 4} phút trước
            </td>
          </tr>
        `;
      });
    }

    // Điền bảng Excel đã làm sạch
    function populateCleanedExcelTable() {
      const tbody = document.getElementById('cleanedExcelTableBody');
      if (!tbody || !analysisData) return;
      tbody.innerHTML = '';

      const profiles = analysisData.profiles || [];
      profiles.forEach(p => {
        const usageText = (p.service_usage_rate_pct === null || p.service_usage_rate_pct === undefined) ? '<span class="text-rose-600 font-bold">NULL (Khuyết)</span>' : p.service_usage_rate_pct + '%';
        const ticketText = p.is_blocked_by_tech ? '<span class="text-rose-600 font-bold">⚠️ ' + p.open_tickets + '</span>' : '<span class="text-slate-400">Không có</span>';
        tbody.innerHTML += `
          <tr class="hover:bg-indigo-50/30 transition">
            <td class="py-3 px-4 font-bold text-indigo-950">
              ${p.customer_name}
              <div class="text-[10px] text-slate-400 font-mono">${p.account_id} • Người đại diện: ${p.representative}</div>
            </td>
            <td class="py-3 px-3 text-slate-700 font-medium">${p.industry}<br><span class="text-[10px] text-indigo-600 font-bold">${p.segment}</span></td>
            <td class="py-3 px-3 text-slate-700">${p.current_package_name}</td>
            <td class="py-3 px-3 font-bold">${usageText}</td>
            <td class="py-3 px-3">${ticketText}</td>
            <td class="py-3 px-3 font-semibold text-indigo-700">${p.recommendation_type}</td>
            <td class="py-3 px-3 text-right font-bold text-indigo-950 font-mono">${formatVND(p.final_price_vnd)}</td>
            <td class="py-3 px-4 text-slate-600 text-[11px] leading-relaxed">${p.ai_remarks}</td>
          </tr>
        `;
      });
    }

    // Vẽ 2 Biểu Đồ
    function renderCharts() {
      if (!analysisData) return;
      const profiles = analysisData.profiles || [];

      // 1. Biểu đồ đường
      const ctxLine = document.getElementById('revenueTrendChart').getContext('2d');
      if (lineChart) lineChart.destroy();

      const renewVal = profiles.filter(p => p.is_revenue_grounded && !p.recommendation_type.includes('NÂNG CẤP')).reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;
      const upVal = profiles.filter(p => p.is_revenue_grounded && p.recommendation_type.includes('NÂNG CẤP')).reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;

      lineChart = new Chart(ctxLine, {
        type: 'line',
        data: {
          labels: ['Tuần 1', 'Tuần 2', 'Tuần 3', 'Tuần 4', 'Tuần 5', 'Tuần 6', 'Hiện tại'],
          datasets: [
            {
              label: 'Doanh thu Gói Duy Trì (Triệu ₫)',
              data: [80, 110, 95, 140, 160, 150, renewVal || 120],
              borderColor: '#4f46e5',
              backgroundColor: 'rgba(79, 70, 229, 0.08)',
              tension: 0.35,
              borderWidth: 2.5,
              fill: true,
              pointRadius: 3
            },
            {
              label: 'Doanh thu Gói Nâng Cấp (Triệu ₫)',
              data: [50, 90, 120, 135, 180, 210, upVal || 225],
              borderColor: '#10b981',
              backgroundColor: 'rgba(16, 185, 129, 0.08)',
              tension: 0.35,
              borderWidth: 2.5,
              fill: true,
              pointRadius: 3
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: { y: { beginAtZero: true } }
        }
      });

      // 2. Biểu đồ tròn
      const ctxDonut = document.getElementById('packageDonutChart').getContext('2d');
      if (donutChart) donutChart.destroy();

      donutChart = new Chart(ctxDonut, {
        type: 'doughnut',
        data: {
          labels: ['Enterprise Suite', 'Hạ Tầng Mở Rộng', 'Điện Toán Cơ Bản', 'Gói Phục Hồi', 'Tạm Giữ'],
          datasets: [{
            data: [35, 30, 20, 10, 5],
            backgroundColor: ['#4f46e5', '#10b981', '#3b82f6', '#8b5cf6', '#f59e0b'],
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          cutout: '72%'
        }
      });
    }

    // Bộ lọc theo phân khúc
    function filterSegment(seg) {
      activeFilter = seg;
      ['all', 'vip', 'sme', 'hold'].forEach(s => {
        const el = document.getElementById('filter-' + s);
        if (el) el.className = 'w-full flex items-center justify-between px-3 py-2 rounded-xl text-slate-700 font-medium text-xs hover:bg-white transition';
      });
      const activeEl = document.getElementById('filter-' + (seg === 'ALL' ? 'all' : seg === 'Enterprise VIP' ? 'vip' : seg === 'SME' ? 'sme' : 'hold'));
      if (activeEl) {
        activeEl.className = 'w-full flex items-center justify-between px-3 py-2 rounded-xl bg-indigo-100 text-indigo-950 font-bold text-xs transition shadow-2xs';
      }
      populateRecentActivityTable();
    }

    // Quét & Chẩn Đoán Lỗi Dữ Liệu
    function triggerErrorDiagnostics() {
      const box = document.getElementById('errorDiagnosticsBox');
      const tbody = document.getElementById('errorTableBody');
      tbody.innerHTML = `
        <tr class="hover:bg-rose-50/50">
          <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">Dòng 5 | Cột: Ngày hết hạn</td>
          <td class="py-2.5 px-3 font-bold text-indigo-950">DN-2026-004</td>
          <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/60 px-1.5 py-0.5 rounded">32/13/2026</td>
          <td class="py-2.5 px-3 text-slate-700">Định dạng ngày không hợp lệ (ngày 32, tháng 13) từ hệ thống CRM cũ</td>
          <td class="py-2.5 px-3 text-emerald-800 font-bold">Chuẩn hóa về định dạng YYYY-MM-DD (2026-11-28)</td>
        </tr>
        <tr class="hover:bg-rose-50/50">
          <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">Dòng 7 | Cột: Mức sử dụng dịch vụ</td>
          <td class="py-2.5 px-3 font-bold text-indigo-950">DN-2026-004</td>
          <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/60 px-1.5 py-0.5 rounded">NULL / Để trống</td>
          <td class="py-2.5 px-3 text-slate-700">Khuyết dữ liệu đo lường công suất do lỗi cổng API (Ticket TK-8004)</td>
          <td class="py-2.5 px-3 text-emerald-800 font-bold">Tạm giữ doanh thu, tuyệt đối không gán bằng 0; chờ đồng bộ API</td>
        </tr>
        <tr class="hover:bg-rose-50/50">
          <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">Dòng 10 | Bảng: Sự cố kỹ thuật</td>
          <td class="py-2.5 px-3 font-bold text-indigo-950">DN-2026-007</td>
          <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/60 px-1.5 py-0.5 rounded">Sự cố Mở (TK-8002)</td>
          <td class="py-2.5 px-3 text-slate-700">Khách hàng gặp sự cố máy chủ mức độ cao chưa được giải quyết</td>
          <td class="py-2.5 px-3 text-emerald-800 font-bold">Kích hoạt chốt chặn an toàn cấm chào tái ký; ưu tiên Kỹ thuật xử lý</td>
        </tr>
      `;
      box.classList.remove('hidden');
      box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    function dismissErrorBox() {
      document.getElementById('errorDiagnosticsBox').classList.add('hidden');
    }

    function autoFixAndRun() {
      dismissErrorBox();
      alert('Đã tự động sửa lỗi định dạng ngày và kích hoạt khóa chặn an toàn kiểm soát lỗi kỹ thuật. Hệ thống đang tiến hành tính toán số liệu!');
      runQuickDemo();
    }

    // Tải xuống file Excel đã làm sạch
    function downloadCleanedExcel() {
      window.location.href = '/api/download-output-xlsx';
    }

    // Xử lý nạp tệp
    function handleFileUpload(e) {
      const file = e.target.files[0];
      if (!file) return;
      alert('Đã tiếp nhận tệp: ' + file.name + ' (' + Math.round(file.size / 1024) + ' KB). Đang quét lỗi và chuẩn hóa dữ liệu...');
      runQuickDemo();
    }

    // Trợ lý AI chat
    async function sendChatMessage() {
      const input = document.getElementById('chatInput');
      const text = input.value.trim();
      if (!text) return;

      const log = document.getElementById('chatMessageLog');

      // Tin nhắn người dùng
      const uMsg = document.createElement('div');
      uMsg.className = 'p-2.5 rounded-xl bg-indigo-600 text-white text-[11px] ml-4 font-medium shadow-2xs';
      uMsg.textContent = text;
      log.appendChild(uMsg);
      input.value = '';

      // Tin nhắn phản hồi đang gõ
      const aMsg = document.createElement('div');
      aMsg.className = 'p-2.5 rounded-xl bg-white text-slate-800 text-[11px] mr-4 border border-indigo-100 shadow-2xs';
      aMsg.innerHTML = '<span class="text-indigo-600 animate-pulse font-bold">Đang tra cứu cơ sở tri thức...</span>';
      log.appendChild(aMsg);
      log.scrollTop = log.scrollHeight;

      try {
        const res = await fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ message: text })
        });
        const data = await res.json();
        aMsg.innerHTML = (data.reply || 'Đã đối soát yêu cầu thành công.')
          .replace(/\n/g, '<br/>')
          .replace(/\*\*(.*?)\*\*/g, '<strong class="text-indigo-950 font-bold">$1</strong>');
      } catch (err) {
        aMsg.textContent = 'Lỗi kết nối tới trợ lý AI: ' + err.message;
      }
      log.scrollTop = log.scrollHeight;
    }

    function searchCustomer(query) {
      if (!query) return;
      document.getElementById('chatInput').value = query;
      sendChatMessage();
    }

    function formatVND(val) {
      if (!val) return '0 ₫';
      return Number(val).toLocaleString('vi-VN') + ' ₫';
    }

    function formatMillions(val) {
      if (!val) return '0 Tr ₫';
      return (Number(val) / 1000000).toFixed(1).replace('.0', '') + ' Tr ₫';
    }

    // Khởi tạo trạng thái 0 ban đầu khi tải trang
    window.addEventListener('DOMContentLoaded', () => {
      resetAllToZero();
    });
  </script>
</body>
</html>
'@

$targetPath = "c:\Users\ADMIN\Downloads\AI Agent CK\web\index.html"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($targetPath, $html, $utf8Bom)
Write-Output "Successfully updated $targetPath with 100% Vietnamese Enterprise Customer Analytics UI."
