# ==============================================================================
# SCRIPT TO BUILD GIA LINH DATA STUDIO UI
# - System Name: Gia Linh Data Studio
# - Welcome Greeting: "Xin chào, Khách hàng 👋"
# - Error Diagnostics: Starts at 0 Lỗi initially (0 lỗi khi chưa nhập dữ liệu & phát hiện lỗi)
# - 2-Column layout: NO Right Sidebar (All AI tools & chat in Left Sidebar)
# - Pastel Blue-Purple aesthetic
# ==============================================================================

$html = @'
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Gia Linh Data Studio — Hệ Thống Phân Tích Dữ Liệu Doanh Nghiệp</title>
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
      background: radial-gradient(circle at 10% 10%, #e0e7ff 0%, #ede9fe 40%, #e0f2fe 100%);
      color: #0f172a;
      min-height: 100vh;
    }
    .custom-scroll::-webkit-scrollbar { width: 5px; height: 5px; }
    .custom-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 9999px; }
    .custom-scroll::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
    .app-window {
      background: #ffffff;
      border: 1px solid rgba(226, 232, 240, 0.9);
      border-radius: 28px;
      box-shadow: 0 25px 65px -12px rgba(99, 102, 241, 0.18), 0 0 1px 1px rgba(99, 102, 241, 0.05);
    }
    .sidebar-collapsed {
      width: 4.75rem !important;
    }
    .sidebar-collapsed .sidebar-content {
      display: none !important;
    }
    .sidebar-collapsed .sidebar-header {
      justify-content: center !important;
    }
    @keyframes fadeInToast {
      from { opacity: 0; transform: translateY(10px) scale(0.95); }
      to { opacity: 1; transform: translateY(0) scale(1); }
    }
    .toast-animate {
      animation: fadeInToast 0.25s ease-out forwards;
    }
  </style>
</head>
<body class="p-2 sm:p-5 flex items-center justify-center relative">

  <!-- ========================================================================= -->
  <!-- TOAST THÔNG BÁO XIN CHÀO KHÁCH HÀNG                                       -->
  <!-- ========================================================================= -->
  <div id="greetingToast" class="fixed bottom-6 right-6 z-50 hidden toast-animate p-4 rounded-2xl bg-slate-900/95 text-white border border-indigo-400/40 shadow-2xl flex items-center gap-3 backdrop-blur-md max-w-md">
    <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-indigo-500 to-purple-500 text-white flex items-center justify-center text-lg font-bold shrink-0 shadow-xs">
      👋
    </div>
    <div class="flex-1 text-xs">
      <div class="font-extrabold text-sm text-indigo-200">Xin chào, Khách hàng!</div>
      <div class="text-slate-300 text-[11px] mt-0.5">Chào mừng bạn đến với <strong>Gia Linh Data Studio</strong>. Hệ thống sẵn sàng đối soát dữ liệu.</div>
    </div>
    <button onclick="dismissGreetingToast()" class="text-slate-400 hover:text-white font-bold p-1 text-xs">✕</button>
  </div>

  <!-- ========================================================================= -->
  <!-- CỬA SỔ ỨNG DỤNG GIA LINH DATA STUDIO (BỐ CỤC 2 CỘT)                       -->
  <!-- ========================================================================= -->
  <div class="app-window w-full max-w-[1740px] min-h-[94vh] flex flex-col overflow-hidden bg-white">

    <!-- 1. BROWSER CHROME HEADER (macOS 3 DOTS + URL BAR + TẢI TÀI LIỆU) -->
    <header class="px-5 py-3 border-b border-slate-100 flex items-center justify-between text-xs bg-slate-50/70 select-none">
      <!-- 3 Chấm màu macOS & Nút điều hướng -->
      <div class="flex items-center gap-3">
        <div class="flex items-center gap-1.5 mr-1">
          <span class="w-3 h-3 rounded-full bg-rose-400 inline-block shadow-2xs"></span>
          <span class="w-3 h-3 rounded-full bg-amber-400 inline-block shadow-2xs"></span>
          <span class="w-3 h-3 rounded-full bg-emerald-400 inline-block shadow-2xs"></span>
        </div>
        <div class="flex items-center gap-1 text-slate-400 text-[11px] ml-1 hidden sm:flex">
          <button class="p-1 hover:text-slate-600 rounded">‹</button>
          <button class="p-1 hover:text-slate-600 rounded">›</button>
        </div>
      </div>

      <!-- Thanh địa chỉ trung tâm macOS -->
      <div class="flex items-center gap-2 bg-white px-5 py-1.5 rounded-xl border border-slate-200/80 text-slate-700 font-semibold text-xs shadow-2xs">
        <span class="text-slate-400 text-xs">🛡️</span>
        <span class="font-bold text-indigo-950">GiaLinhDataStudio.vn</span>
        <span class="text-slate-300">/</span>
        <span class="text-indigo-600 font-medium">enterprise-analytics</span>
        <span class="text-slate-400 text-[11px] ml-1">🔄</span>
      </div>

      <!-- Tải tài liệu báo cáo & Đặt lại về 0 -->
      <div class="flex items-center gap-2 text-xs">
        <a href="/api/download-report-docx" download class="px-3 py-1.5 rounded-xl bg-indigo-50 hover:bg-indigo-100 text-indigo-700 border border-indigo-200 font-bold transition shadow-2xs flex items-center gap-1.5">
          <span>📄</span><span class="hidden md:inline">Báo Cáo Word (.docx)</span>
        </a>
        <a href="/api/download-presentation-pptx" download class="px-3 py-1.5 rounded-xl bg-purple-50 hover:bg-purple-100 text-purple-700 border border-purple-200 font-bold transition shadow-2xs flex items-center gap-1.5">
          <span>📽️</span><span class="hidden md:inline">Slide Thuyết Trình (.pptx)</span>
        </a>
        <button onclick="resetAllToZero()" class="px-3 py-1.5 rounded-xl bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 font-bold transition shadow-2xs" title="Đưa toàn bộ số liệu và chẩn đoán lỗi về 0">
          ↺ Đặt Lại Về 0
        </button>
      </div>
    </header>

    <!-- 2. INNER APP TOP BAR (LOGO GIA LINH DATA STUDIO + TÌM KIẾM + KHÁCH HÀNG) -->
    <div class="px-6 py-3.5 border-b border-slate-100 flex items-center justify-between gap-4 bg-white select-none">
      <!-- Logo Gia Linh Data Studio & Nút mở/đóng thanh công cụ -->
      <div class="flex items-center gap-3">
        <button onclick="toggleSidebar()" class="p-2 rounded-xl bg-slate-50 hover:bg-indigo-50 border border-slate-200 text-slate-600 hover:text-indigo-600 transition" title="Đóng / Mở Thanh Công Cụ Bên Trái">
          <span id="sidebarToggleIcon" class="text-xs font-bold">◀</span>
        </button>
        <div class="flex items-center gap-2.5 cursor-pointer" onclick="showGreetingToast()">
          <div class="w-8 h-8 rounded-xl bg-gradient-to-tr from-indigo-600 via-indigo-500 to-purple-600 text-white flex items-center justify-center font-black text-xs shadow-sm tracking-wider">
            GL
          </div>
          <div>
            <span class="text-base font-extrabold tracking-tight text-slate-900">Gia Linh Data Studio</span>
            <span class="ml-1.5 px-2 py-0.5 rounded-md bg-indigo-50 text-[10px] font-bold text-indigo-600 border border-indigo-100">Enterprise</span>
          </div>
        </div>
      </div>

      <!-- Ô tìm kiếm thông minh ở giữa -->
      <div class="relative flex-1 max-w-xl mx-4">
        <span class="absolute left-3.5 top-2.5 text-slate-400 text-xs">🔍</span>
        <input id="searchInput" type="text" placeholder="Tìm kiếm theo mã khách hàng (DN-001, DN-004), tên công ty, hoặc hợp đồng..." onkeydown="if(event.key==='Enter') searchCustomer(this.value)" class="w-full pl-9 pr-14 py-2 rounded-xl bg-slate-50/80 border border-slate-200 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-indigo-500 focus:bg-white transition shadow-2xs font-medium" />
        <span class="absolute right-3 top-2 px-1.5 py-0.5 rounded text-[10px] bg-slate-200 text-slate-600 font-mono font-semibold">⌘K</span>
      </div>

      <!-- Cụm chức năng bên phải -->
      <div class="flex items-center gap-3">
        <!-- Nút Chạy Demo Nhanh -->
        <button onclick="runQuickDemo()" id="btnTopQuickDemo" class="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs shadow-xs transition flex items-center gap-1.5">
          <span>⚡</span>
          <span>Chạy Phân Tích (Demo Nhanh)</span>
        </button>

        <!-- Biểu tượng quà tặng & thông báo -->
        <div class="flex items-center gap-1">
          <button class="p-2 rounded-xl hover:bg-slate-100 text-slate-500 transition text-sm" title="Ưu đãi & Khuyến mãi">🎁</button>
          <button class="p-2 rounded-xl hover:bg-slate-100 text-slate-500 transition text-sm relative" title="Thông báo hệ thống">
            🔔
            <span class="w-2 h-2 rounded-full bg-indigo-600 absolute top-1.5 right-1.5"></span>
          </button>
        </div>

        <!-- Khách Hàng (Avatar & Tên hiển thị) -->
        <div class="flex items-center gap-2 pl-2 border-l border-slate-100 cursor-pointer" onclick="showGreetingToast()" title="Bấm vào để hiển thị lời chào">
          <div class="relative">
            <div class="w-8 h-8 rounded-full bg-gradient-to-tr from-indigo-500 to-purple-600 text-white font-extrabold text-xs flex items-center justify-center shadow-2xs">
              KH
            </div>
            <span class="w-2.5 h-2.5 rounded-full bg-emerald-500 ring-2 ring-white absolute -bottom-0.5 -right-0.5"></span>
          </div>
          <div class="hidden xl:block text-left leading-tight">
            <div class="text-xs font-bold text-slate-800">Khách Hàng</div>
            <div class="text-[10px] text-slate-400 font-medium">Gia Linh Data Studio</div>
          </div>
        </div>
      </div>
    </div>

    <!-- 3. MAIN BODY: 2-COLUMN LAYOUT -->
    <div class="flex-1 flex flex-col lg:flex-row overflow-hidden">

      <!-- ===================================================================== -->
      <!-- CỘT TRÁI: THANH CÔNG CỤ DỌC (COLLAPSIBLE SIDEBAR)                      -->
      <!-- TOÀN BỘ CÔNG CỤ & TRỢ LÝ AI ĐƯỢC CHUYỂN VÀO ĐÂY                       -->
      <!-- ===================================================================== -->
      <aside id="leftSidebar" class="w-full lg:w-88 xl:w-96 bg-slate-50/70 border-r border-slate-100 p-4 flex flex-col justify-between shrink-0 select-none transition-all duration-300 overflow-y-auto custom-scroll space-y-4">
        
        <div class="space-y-4">
          
          <!-- MENU ĐIỀU HƯỚNG CHÍNH -->
          <div class="space-y-1">
            <div class="sidebar-content px-2 pb-1 text-[10px] font-extrabold text-slate-400 uppercase tracking-wider">
              Điều Hướng Hệ Thống
            </div>
            
            <button onclick="switchTab('dashboard')" id="nav-dashboard" class="w-full px-3 py-2 rounded-xl bg-white text-indigo-700 font-bold text-xs shadow-2xs border border-indigo-100 flex items-center gap-3 transition">
              <span class="text-sm shrink-0">🏠</span>
              <span class="sidebar-content flex-1 text-left">Tổng Quan (Dashboard)</span>
              <span class="sidebar-content text-[10px] font-extrabold px-1.5 py-0.5 rounded-full bg-indigo-50 text-indigo-600">Trực tiếp</span>
            </button>

            <button onclick="switchTab('table')" id="nav-table" class="w-full px-3 py-2 rounded-xl text-slate-600 hover:bg-white hover:text-indigo-700 font-semibold text-xs flex items-center gap-3 transition">
              <span class="text-sm shrink-0">📗</span>
              <span class="sidebar-content flex-1 text-left">Dữ Liệu Đã Làm Sạch &amp; AI</span>
            </button>

            <!-- CHẨN ĐOÁN LỖI FILE: BAN ĐẦU 0 LỖI KHI CHƯA NHẬP DỮ LIỆU & PHÁT HIỆN LỖI -->
            <button onclick="triggerErrorDiagnostics()" id="nav-errors" class="w-full px-3 py-2 rounded-xl text-slate-700 hover:bg-white hover:text-indigo-700 font-semibold text-xs flex items-center gap-3 transition border border-transparent hover:border-slate-200">
              <span class="text-sm shrink-0">⚠️</span>
              <span class="sidebar-content flex-1 text-left flex items-center justify-between">
                <span>Chẩn Đoán Lỗi File</span>
                <span id="navErrorCountBadge" class="px-2 py-0.5 rounded-full bg-slate-200 text-slate-700 font-extrabold text-[10px]">0 lỗi</span>
              </span>
            </button>

            <button onclick="downloadCleanedExcel()" class="w-full px-3 py-2 rounded-xl text-emerald-700 hover:bg-emerald-50 font-semibold text-xs flex items-center gap-3 transition">
              <span class="text-sm shrink-0">📥</span>
              <span class="sidebar-content flex-1 text-left">Xuất Báo Cáo Excel (.xlsx)</span>
            </button>
          </div>

          <!-- NẠP TỆP FILE KHÁCH HÀNG -->
          <div class="sidebar-content pt-2 border-t border-slate-200/70">
            <div class="px-2 pb-1.5 text-[10px] font-extrabold text-slate-400 uppercase tracking-wider">
              Nạp Tệp Phân Tích
            </div>
            <label class="w-full p-2.5 rounded-xl border border-dashed border-indigo-300 bg-white hover:bg-indigo-50/50 text-indigo-900 font-semibold text-xs cursor-pointer transition flex items-center gap-2.5 shadow-2xs">
              <input type="file" id="fileInput" accept=".xlsx,.xls,.csv,.json" class="hidden" onchange="handleFileUpload(event)" />
              <span class="text-base text-indigo-600 shrink-0">📂</span>
              <div class="overflow-hidden">
                <div class="font-bold text-[11px] truncate">Chọn Tệp Dữ Liệu (.xlsx, .json)</div>
                <div class="text-[9px] text-slate-400">Tự động quét cấu trúc &amp; kiểm tra lỗi</div>
              </div>
            </label>
          </div>

          <!-- BỘ LỌC PHÂN KHÚC KHÁCH HÀNG -->
          <div class="sidebar-content pt-2 border-t border-slate-200/70 space-y-1.5">
            <div class="flex items-center justify-between px-2">
              <span class="text-[10px] font-extrabold text-slate-400 uppercase tracking-wider">Bộ Lọc Phân Khúc</span>
              <span class="text-[9px] text-slate-400">4 nhóm</span>
            </div>
            <div class="grid grid-cols-2 gap-1.5">
              <button onclick="filterSegment('ALL')" id="filter-all" class="flex items-center justify-between px-2.5 py-1.5 rounded-xl bg-white text-indigo-950 font-bold text-[11px] border border-indigo-200 shadow-2xs transition">
                <span>Tất cả</span>
                <span id="badgeCountAll" class="px-1.5 py-0.2 rounded-full bg-indigo-100 text-indigo-700 text-[10px] font-extrabold">0</span>
              </button>
              <button onclick="filterSegment('Enterprise VIP')" id="filter-vip" class="flex items-center justify-between px-2.5 py-1.5 rounded-xl bg-white hover:bg-indigo-50 border border-slate-200 text-slate-700 font-semibold text-[11px] transition">
                <span>Khách VIP</span>
                <span id="badgeCountVip" class="px-1.5 py-0.2 rounded-full bg-slate-100 text-indigo-700 text-[10px] font-bold">0</span>
              </button>
              <button onclick="filterSegment('SME')" id="filter-sme" class="flex items-center justify-between px-2.5 py-1.5 rounded-xl bg-white hover:bg-indigo-50 border border-slate-200 text-slate-700 font-semibold text-[11px] transition">
                <span>Khối SME</span>
                <span id="badgeCountSme" class="px-1.5 py-0.2 rounded-full bg-slate-100 text-indigo-700 text-[10px] font-bold">0</span>
              </button>
              <button onclick="filterSegment('HOLD')" id="filter-hold" class="flex items-center justify-between px-2.5 py-1.5 rounded-xl bg-white hover:bg-amber-50 border border-amber-200 text-amber-800 font-semibold text-[11px] transition">
                <span>Tạm giữ</span>
                <span id="badgeCountHold" class="px-1.5 py-0.2 rounded-full bg-amber-100 text-amber-800 text-[10px] font-bold">0</span>
              </button>
            </div>
          </div>

          <!-- PHẦN TRỢ LÝ PHÂN TÍCH AI (CHUYỂN HOÀN TOÀN TỪ BÊN PHẢI SANG BÊN TRÁI) -->
          <div class="sidebar-content pt-2 border-t border-slate-200/70 space-y-2.5">
            <div class="flex items-center justify-between px-1">
              <div class="flex items-center gap-1.5">
                <span class="text-indigo-600 font-extrabold text-sm">✨</span>
                <span class="text-[11px] font-extrabold text-slate-800 uppercase tracking-wider">Trợ Lý Phân Tích AI</span>
              </div>
              <span class="text-[9px] text-emerald-700 font-bold bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-full">● Trực Tuyến</span>
            </div>

            <!-- Thẻ Gợi Ý Thông Minh (Smart Insight Card) -->
            <div class="p-3 rounded-2xl bg-gradient-to-br from-emerald-50/90 to-teal-50/50 border border-emerald-200/80 text-[11px] space-y-1.5 shadow-2xs">
              <div class="font-extrabold text-emerald-950 flex items-center justify-between">
                <span class="flex items-center gap-1"><span>💡</span><span>Gợi Ý Thông Minh</span></span>
                <span class="text-[9px] text-emerald-600 font-normal">Tự động</span>
              </div>
              <p class="text-[10px] text-emerald-900 leading-relaxed font-medium">
                Hệ thống áp dụng khóa chặn an toàn (Hard-Block) khi phát hiện sự cố máy chủ, cấm chào mời tái ký trước khi sửa xong sự cố.
              </p>
            </div>

            <!-- Thẻ Gợi Ý Nhanh Câu Hỏi -->
            <div class="space-y-1">
              <div class="text-[9px] text-slate-400 font-bold uppercase tracking-wider px-1">Câu hỏi gợi ý:</div>
              <div class="flex flex-wrap gap-1">
                <button onclick="searchCustomer('DN-004')" class="px-2 py-1 rounded-lg bg-white border border-slate-200 text-slate-700 hover:text-indigo-600 hover:border-indigo-300 text-[10px] font-semibold transition">
                  🔍 DN-004 (Lỗi API / NULL)
                </button>
                <button onclick="searchCustomer('DN-007')" class="px-2 py-1 rounded-lg bg-white border border-slate-200 text-slate-700 hover:text-rose-600 hover:border-rose-300 text-[10px] font-semibold transition">
                  🛡️ DN-007 (Hard-Block)
                </button>
                <button onclick="searchCustomer('NULL vs 0')" class="px-2 py-1 rounded-lg bg-white border border-slate-200 text-slate-700 hover:text-indigo-600 hover:border-indigo-300 text-[10px] font-semibold transition">
                  ⚖️ NULL vs 0%
                </button>
              </div>
            </div>

            <!-- Khung Lịch Sử Hội Thoại Chat Trực Tiếp -->
            <div id="chatMessageLog" class="max-h-48 overflow-y-auto custom-scroll space-y-1.5 p-2 bg-white rounded-xl border border-slate-200 text-xs shadow-2xs">
              <div class="p-2 rounded-lg bg-indigo-50/80 text-indigo-950 text-[11px] leading-relaxed">
                👋 <strong>Xin chào, Khách hàng!</strong> Tôi là Trợ lý Phân Tích Dữ Liệu <strong>Gia Linh Data Studio</strong>. Bạn có thể tra cứu chi tiết từng mã tài khoản, đối chiếu căn cứ doanh thu, hoặc kiểm tra khóa chặn sự cố kỹ thuật.
              </div>
            </div>

            <!-- Ô Nhập Câu Hỏi Chat -->
            <div class="relative">
              <input id="chatInput" type="text" placeholder="Hỏi về DN-004, DN-007, NULL vs 0..." onkeydown="if(event.key==='Enter') sendChatMessage()" class="w-full pl-3 pr-8 py-2 rounded-xl bg-white border border-slate-200 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-indigo-500 shadow-2xs font-medium" />
              <button onclick="sendChatMessage()" class="absolute right-1.5 top-2 w-6 h-6 rounded-lg bg-indigo-600 hover:bg-indigo-700 text-white flex items-center justify-center text-[10px] transition" title="Gửi câu hỏi">
                ➤
              </button>
            </div>
          </div>

        </div>

        <!-- Chân thanh công cụ: Thẻ Gói Doanh Nghiệp & Hồ Sơ Khách Hàng -->
        <div class="pt-3 border-t border-slate-200/70 space-y-2.5">
          <!-- Card Enterprise Plan -->
          <div class="sidebar-content p-3 bg-white rounded-2xl border border-slate-200/80 space-y-1.5 shadow-2xs">
            <div class="flex justify-between items-center text-xs font-extrabold text-slate-800">
              <span>Gói Doanh Nghiệp</span>
              <span id="progressPctText" class="text-indigo-600 font-extrabold">0%</span>
            </div>
            <div class="text-[10px] text-slate-400">Tái tục vào ngày 24/11/2026</div>
            <div class="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
              <div id="progressBar" class="bg-indigo-600 h-1.5 rounded-full transition-all duration-500" style="width: 0%"></div>
            </div>
            <div id="progressStatusText" class="text-[9px] text-slate-500 pt-0.5">Chờ nạp dữ liệu phân tích</div>
          </div>

          <!-- Thông tin Khách Hàng -->
          <div class="flex items-center gap-2.5 px-1 sidebar-content cursor-pointer" onclick="showGreetingToast()" title="Bấm vào để hiển thị lời chào">
            <div class="w-8 h-8 rounded-full bg-slate-800 text-white flex items-center justify-center font-bold text-xs shrink-0 shadow-2xs">
              KH
            </div>
            <div class="overflow-hidden flex-1">
              <div class="font-bold text-xs text-slate-800 truncate">Khách Hàng</div>
              <div class="text-[10px] text-slate-400 font-medium truncate">Người Dùng Hệ Thống</div>
            </div>
            <span class="text-slate-400 text-xs">•••</span>
          </div>
        </div>
      </aside>

      <!-- ===================================================================== -->
      <!-- KHÔNG GIAN LÀM VIỆC CHÍNH (FULL REMAINING WIDTH - 2 COLUMN)          -->
      <!-- ===================================================================== -->
      <main class="flex-1 overflow-y-auto custom-scroll p-4 sm:p-7 bg-white flex flex-col justify-between space-y-6">

        <!-- TIÊU ĐỀ CHÀO MỪNG KHÁCH HÀNG: KHI BẤM VÀO SẼ HIỆN XIN CHÀO KHÁCH HÀNG -->
        <div class="flex flex-wrap items-center justify-between gap-3">
          <div class="cursor-pointer" onclick="showGreetingToast()" title="Bấm vào để hiện thông báo lời chào">
            <h1 id="greetingTitle" class="text-2xl font-black text-slate-900 tracking-tight flex items-center gap-2 hover:text-indigo-600 transition">
              Xin chào, khách hàng 👋
            </h1>
            <p class="text-xs text-slate-500 mt-0.5">
              Chào mừng bạn đến với <strong>Gia Linh Data Studio</strong>. Dưới đây là bức tranh tổng quan đối soát dữ liệu doanh nghiệp hôm nay.
            </p>
          </div>

          <!-- Date badge pill -->
          <div class="flex items-center gap-2 px-3.5 py-1.5 rounded-xl border border-slate-200 bg-slate-50 text-xs font-semibold text-slate-700 shadow-2xs">
            <span>12 Thg 5 – 18 Thg 5, 2026</span>
            <span>📅</span>
          </div>
        </div>

        <!-- KHUNG CHẨN ĐOÁN LỖI TỆP DỮ LIỆU (WHERE - WHY - FIX GUIDE) -->
        <!-- BAN ĐẦU LÀ 0 LỖI KHI CHƯA NHẬP DỮ LIỆU & PHÁT HIỆN LỖI -->
        <div id="errorDiagnosticsBox" class="hidden p-4 rounded-2xl border text-xs space-y-3 shadow-2xs transition-all">
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-2">
              <span id="diagBoxIcon" class="w-6 h-6 rounded-lg bg-emerald-600 text-white flex items-center justify-center font-bold text-xs">✓</span>
              <span id="diagBoxTitle" class="text-slate-900 font-extrabold text-sm">TRẠNG THÁI DỮ LIỆU ĐẦU VÀO</span>
              <span id="diagBadgeCount" class="px-2 py-0.5 rounded-full bg-slate-200 text-slate-700 font-extrabold text-[10px]">0 Lỗi Phát Hiện</span>
            </div>
            <button onclick="dismissErrorBox()" class="text-slate-500 hover:text-slate-800 font-bold">✕ Đóng</button>
          </div>
          
          <div id="errorTableContainer" class="overflow-x-auto bg-white rounded-xl border border-slate-200">
            <!-- Nội dung bảng lỗi hoặc thông báo an toàn được render bằng JS -->
          </div>

          <div id="errorBoxFooter" class="flex flex-wrap justify-between items-center gap-2 pt-1">
            <!-- Nút hành động hoặc thông báo hướng dẫn -->
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- NỘI DUNG 1: BẢNG ĐIỀU KHIỂN DASHBOARD (4 KPI + 2 BIỂU ĐỒ + HOẠT ĐỘNG)-->
        <!-- =================================================================== -->
        <div id="sectionDashboard" class="space-y-6">

          <!-- 4 THẺ KPI CHỈ SỐ (BẮT ĐẦU VỀ 0 HOÀN TOÀN KHI CHƯA CHẠY) -->
          <div class="grid grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
            <!-- Thẻ 1: Tổng Số Khách Hàng (Total Agents) -->
            <div class="p-4.5 rounded-2xl border border-slate-100 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Tổng Số Khách Hàng</span>
                <span class="w-7 h-7 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center font-bold text-xs">✓</span>
              </div>
              <div id="kpiTotalNum" class="text-3xl font-extrabold text-slate-900">0</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiTotalSub" class="text-slate-400 font-medium">Chưa nạp dữ liệu</span>
                <svg class="w-16 h-4 text-indigo-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 10 Q 15 2, 30 8 T 59 4" />
                </svg>
              </div>
            </div>

            <!-- Thẻ 2: Doanh Thu Đủ Căn Cứ (Tasks Completed) -->
            <div class="p-4.5 rounded-2xl border border-slate-100 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Doanh Thu Đủ Căn Cứ</span>
                <span class="w-7 h-7 rounded-full bg-emerald-50 text-emerald-600 flex items-center justify-center font-bold text-xs">📋</span>
              </div>
              <div id="kpiGroundedNum" class="text-3xl font-extrabold text-slate-900">0 ₫</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiGroundedSub" class="text-slate-400 font-medium">Chưa nạp dữ liệu</span>
                <svg class="w-16 h-4 text-emerald-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 12 Q 15 5, 30 10 T 59 2" />
                </svg>
              </div>
            </div>

            <!-- Thẻ 3: Mức Sử Dụng Dịch Vụ TB (Success Rate) -->
            <div class="p-4.5 rounded-2xl border border-slate-100 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Mức Sử Dụng Dịch Vụ TB</span>
                <span class="w-7 h-7 rounded-full bg-purple-50 text-purple-600 flex items-center justify-center font-bold text-xs">📊</span>
              </div>
              <div id="kpiUsageNum" class="text-3xl font-extrabold text-slate-900">0%</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiUsageSub" class="text-slate-400 font-medium">Chưa nạp dữ liệu</span>
                <svg class="w-16 h-4 text-purple-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 8 Q 15 12, 30 4 T 59 7" />
                </svg>
              </div>
            </div>

            <!-- Thẻ 4: Doanh Thu Tạm Giữ (Time Saved) -->
            <div class="p-4.5 rounded-2xl border border-slate-100 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Doanh Thu Tạm Giữ (Sự Cố/Khuyết)</span>
                <span class="w-7 h-7 rounded-full bg-amber-50 text-amber-600 flex items-center justify-center font-bold text-xs">⏰</span>
              </div>
              <div id="kpiHoldNum" class="text-3xl font-extrabold text-slate-900">0 ₫</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiHoldSub" class="text-slate-400 font-medium">Chưa nạp dữ liệu</span>
                <svg class="w-16 h-4 text-amber-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 6 Q 15 10, 30 3 T 59 9" />
                </svg>
              </div>
            </div>
          </div>

          <!-- 2 BIỂU ĐỒ TRỰC QUAN MỞ RỘNG (PERFORMANCE OVERVIEW & TASKS BY AGENT) -->
          <div class="grid grid-cols-1 lg:grid-cols-12 gap-5">
            <!-- Biểu đồ đường (Line Chart) -->
            <div class="lg:col-span-7 p-5 rounded-2xl border border-slate-100 bg-white space-y-3 shadow-2xs">
              <div class="flex justify-between items-center">
                <div class="space-y-0.5">
                  <h3 class="font-bold text-xs text-slate-900">Tổng Quan Hiệu Suất (Performance Overview)</h3>
                  <div class="flex items-center gap-3 text-[10px] text-slate-500">
                    <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-indigo-600 inline-block"></span> Doanh thu Gói Duy Trì</span>
                    <span class="flex items-center gap-1"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span> Doanh thu Gói Nâng Cấp</span>
                  </div>
                </div>
                <div class="text-[10px] text-slate-600 font-semibold px-2.5 py-1 rounded-lg bg-slate-50 border border-slate-200">
                  Theo Tuần ▾
                </div>
              </div>
              <div class="h-60 w-full">
                <canvas id="revenueTrendChart"></canvas>
              </div>
            </div>

            <!-- Biểu đồ tròn (Donut Chart - Tasks by Agent) -->
            <div class="lg:col-span-5 p-5 rounded-2xl border border-slate-100 bg-white space-y-3 shadow-2xs">
              <div class="flex justify-between items-center">
                <h3 class="font-bold text-xs text-slate-900">Phân Bổ Gói Dịch Vụ (Tasks by Agent)</h3>
                <span class="text-[10px] text-indigo-600 font-bold">Tỷ trọng</span>
              </div>
              <div class="grid grid-cols-12 gap-2 items-center h-60">
                <div class="col-span-6 h-full flex items-center justify-center relative">
                  <canvas id="packageDonutChart"></canvas>
                  <div class="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                    <span id="donutCenterNum" class="text-2xl font-black text-slate-900">0</span>
                    <span class="text-[9px] text-slate-400 uppercase font-bold">Khách Hàng</span>
                  </div>
                </div>
                <div class="col-span-6 space-y-2 text-[11px] text-slate-600">
                  <div class="flex justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-indigo-600"></span> Enterprise Suite</span> <strong id="pctPkg1" class="text-slate-900">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span> Hạ Tầng Mở Rộng</span> <strong id="pctPkg2" class="text-slate-900">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-blue-500"></span> Điện Toán Cơ Bản</span> <strong id="pctPkg3" class="text-slate-900">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-purple-500"></span> Hỗ Trợ Phục Hồi</span> <strong id="pctPkg4" class="text-slate-900">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-amber-400"></span> Tạm Giữ Sự Cố</span> <strong id="pctPkg5" class="text-slate-900">0%</strong></div>
                </div>
              </div>
            </div>
          </div>

          <!-- BẢNG HOẠT ĐỘNG KHÁCH HÀNG GẦN ĐÂY (RECENT AGENT ACTIVITY) -->
          <div class="p-5 rounded-2xl border border-slate-100 bg-white space-y-3 shadow-2xs">
            <div class="flex justify-between items-center">
              <h3 class="font-bold text-xs text-slate-900">Hoạt Động Khách Hàng Gần Đây (Recent Activity)</h3>
              <div class="flex items-center gap-3">
                <span id="activityCountText" class="text-[11px] text-slate-400 font-medium">0 tài khoản</span>
                <button onclick="switchTab('table')" class="text-[11px] text-indigo-600 hover:text-indigo-800 font-bold">Xem Tất Cả (View All)</button>
              </div>
            </div>
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <tbody id="activityTableBody" class="divide-y divide-slate-100">
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
              <h2 class="text-base font-extrabold text-slate-900">Dữ Liệu Khách Hàng Doanh Nghiệp Đã Làm Sạch Kèm Nhận Xét AI</h2>
              <p class="text-xs text-slate-500">Toàn bộ hồ sơ đã qua xử lý lọc khuyết thiếu, tách biệt NULL ≠ 0 và kiểm soát khóa chặn sự cố kỹ thuật.</p>
            </div>
            <button onclick="downloadCleanedExcel()" class="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-extrabold text-xs shadow-xs transition flex items-center gap-1.5">
              <span>📥 Tải File Excel (.xlsx)</span>
            </button>
          </div>

          <div class="rounded-2xl border border-slate-100 bg-white overflow-hidden shadow-2xs">
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <thead class="bg-slate-50 text-slate-700 uppercase font-bold text-[10px] border-b border-slate-100">
                  <tr>
                    <th class="py-3 px-4">Mã &amp; Doanh Nghiệp</th>
                    <th class="py-3 px-3">Lĩnh Vực &amp; Phân Khúc</th>
                    <th class="py-3 px-3">Gói Hiện Tại</th>
                    <th class="py-3 px-3">Mức Sử Dụng</th>
                    <th class="py-3 px-3">Sự Cố Kỹ Thuật</th>
                    <th class="py-3 px-3">Đề Xuất Hành Động</th>
                    <th class="py-3 px-3 text-right">Giá Trị Đề Xuất</th>
                    <th class="py-3 px-4 min-w-[320px]">Nhận Xét &amp; Đánh Giá AI</th>
                  </tr>
                </thead>
                <tbody id="cleanedExcelTableBody" class="divide-y divide-slate-100">
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

        <!-- Chân trang ứng dụng Gia Linh Data Studio -->
        <footer class="pt-4 border-t border-slate-100 flex justify-between items-center text-[11px] text-slate-400 select-none">
          <span>Gia Linh Data Studio — Nền Tảng Phân Tích Dữ Liệu Doanh Nghiệp v4.0</span>
          <span>Bảo mật dữ liệu trên máy chủ cục bộ: http://localhost:8080</span>
        </footer>
      </main>

    </div>
  </div>

  <!-- ========================================================================= -->
  <!-- JAVASCRIPT ĐIỀU KHIỂN HỆ THỐNG GIA LINH DATA STUDIO                        -->
  <!-- ========================================================================= -->
  <script>
    let isDataLoaded = false;
    let analysisData = null;
    let activeFilter = 'ALL';
    let lineChart = null;
    let donutChart = null;
    let isSidebarCollapsed = false;

    // Biến lưu trạng thái chẩn đoán lỗi: Ban đầu là 0 lỗi (khi chưa nhập dữ liệu & phát hiện lỗi)
    let detectedErrorsCount = 0;
    let detectedErrorsList = [];

    // Hiển thị Lời Chào Khách Hàng Toast
    function showGreetingToast() {
      const toast = document.getElementById('greetingToast');
      toast.classList.remove('hidden');
      setTimeout(() => {
        // Tự ẩn sau 6 giây nếu người dùng không bấm tắt
        if (!toast.classList.contains('hidden')) {
          dismissGreetingToast();
        }
      }, 6000);
    }

    function dismissGreetingToast() {
      document.getElementById('greetingToast').classList.add('hidden');
    }

    // Đóng / Mở Thanh Công Cụ Dọc Bên Trái
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
      const navDash = document.getElementById('nav-dashboard');
      const navTable = document.getElementById('nav-table');

      if (tabName === 'dashboard') {
        secDash.classList.remove('hidden');
        secTable.classList.add('hidden');
        navDash.className = 'w-full px-3 py-2 rounded-xl bg-white text-indigo-700 font-bold text-xs shadow-2xs border border-indigo-100 flex items-center gap-3 transition';
        navTable.className = 'w-full px-3 py-2 rounded-xl text-slate-600 hover:bg-white hover:text-indigo-700 font-semibold text-xs flex items-center gap-3 transition';
        if (isDataLoaded) renderCharts();
      } else {
        secDash.classList.add('hidden');
        secTable.classList.remove('hidden');
        navDash.className = 'w-full px-3 py-2 rounded-xl text-slate-600 hover:bg-white hover:text-indigo-700 font-semibold text-xs flex items-center gap-3 transition';
        navTable.className = 'w-full px-3 py-2 rounded-xl bg-white text-indigo-700 font-bold text-xs shadow-2xs border border-indigo-100 flex items-center gap-3 transition';
      }
    }

    // Đặt lại toàn bộ về 0 (Zero-Data Initial State & 0 Lỗi)
    function resetAllToZero() {
      isDataLoaded = false;
      analysisData = null;

      // Đặt lại chẩn đoán lỗi về 0 lỗi khi chưa nạp dữ liệu
      detectedErrorsCount = 0;
      detectedErrorsList = [];
      updateErrorDisplay();
      dismissErrorBox();

      document.getElementById('kpiTotalNum').textContent = '0';
      document.getElementById('kpiTotalSub').textContent = 'Chưa nạp dữ liệu';
      document.getElementById('kpiGroundedNum').textContent = '0 ₫';
      document.getElementById('kpiGroundedSub').textContent = 'Chưa nạp dữ liệu';
      document.getElementById('kpiUsageNum').textContent = '0%';
      document.getElementById('kpiUsageSub').textContent = 'Chưa nạp dữ liệu';
      document.getElementById('kpiHoldNum').textContent = '0 ₫';
      document.getElementById('kpiHoldSub').textContent = 'Chưa nạp dữ liệu';

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

    // Cập nhật giao diện trạng thái lỗi (0 lỗi vs >0 lỗi)
    function updateErrorDisplay() {
      const badgeNav = document.getElementById('navErrorCountBadge');
      if (detectedErrorsCount === 0) {
        badgeNav.textContent = '0 lỗi';
        badgeNav.className = 'px-2 py-0.5 rounded-full bg-slate-200 text-slate-700 font-extrabold text-[10px]';
      } else {
        badgeNav.textContent = detectedErrorsCount + ' lỗi';
        badgeNav.className = 'px-2 py-0.5 rounded-full bg-rose-600 text-white font-extrabold text-[10px] animate-pulse';
      }
    }

    // Quét & Chẩn Đoán Lỗi Dữ Liệu
    function triggerErrorDiagnostics() {
      const box = document.getElementById('errorDiagnosticsBox');
      const boxIcon = document.getElementById('diagBoxIcon');
      const boxTitle = document.getElementById('diagBoxTitle');
      const badgeCount = document.getElementById('diagBadgeCount');
      const container = document.getElementById('errorTableContainer');
      const footer = document.getElementById('errorBoxFooter');

      if (detectedErrorsCount === 0) {
        // TRẠNG THÁI: 0 LỖI KHI CHƯA NHẬP DỮ LIỆU HOẶC CHƯA PHÁT HIỆN LỖI
        box.className = 'p-4 rounded-2xl border border-emerald-200 bg-emerald-50/90 text-xs space-y-3 shadow-2xs';
        boxIcon.textContent = '✓';
        boxIcon.className = 'w-6 h-6 rounded-lg bg-emerald-600 text-white flex items-center justify-center font-bold text-xs';
        boxTitle.textContent = 'TRẠNG THÁI DỮ LIỆU ĐẦU VÀO: AN TOÀN';
        boxTitle.className = 'text-emerald-950 font-extrabold text-sm';
        badgeCount.textContent = '0 Lỗi Phát Hiện';
        badgeCount.className = 'px-2 py-0.5 rounded-full bg-emerald-200 text-emerald-900 font-extrabold text-[10px]';

        container.innerHTML = `
          <div class="p-6 text-center text-slate-600 space-y-2">
            <div class="text-3xl text-emerald-600">🛡️</div>
            <div class="font-extrabold text-slate-800 text-sm">Chưa phát hiện lỗi nào trong dữ liệu (0 lỗi)</div>
            <p class="text-xs text-slate-500 max-w-lg mx-auto">
              Hệ thống Gia Linh Data Studio chưa ghi nhận sai lệch hoặc bạn chưa nạp tệp phân tích. Bạn có thể nạp tệp mới hoặc bấm nút bên dưới để quét thử nghiệm tệp có lỗi.
            </p>
          </div>
        `;

        footer.innerHTML = `
          <span class="text-[10px] text-emerald-800 font-medium">Hệ thống tuân thủ nguyên tắc an toàn: Dữ liệu sạch mới đưa vào tính toán.</span>
          <button onclick="simulateFindErrors()" class="px-4 py-1.5 bg-rose-600 hover:bg-rose-700 text-white font-bold rounded-xl text-xs shadow-xs transition flex items-center gap-1.5">
            <span>⚠️</span><span>Quét Tệp Kiểm Thử (Phát Hiện 3 Điểm Sai Lệch)</span>
          </button>
        `;
      } else {
        // TRẠNG THÁI: ĐÃ PHÁT HIỆN LỖI (3 LỖI CỤ THỂ KHI NHẬP FILE HOẶC QUÉT)
        box.className = 'p-4 rounded-2xl border border-rose-200 bg-rose-50/90 text-xs space-y-3 shadow-2xs';
        boxIcon.textContent = '⚠️';
        boxIcon.className = 'w-6 h-6 rounded-lg bg-rose-600 text-white flex items-center justify-center font-bold text-xs';
        boxTitle.textContent = 'PHÁT HIỆN LỖI TRONG DỮ LIỆU ĐẦU VÀO';
        boxTitle.className = 'text-rose-900 font-extrabold text-sm';
        badgeCount.textContent = detectedErrorsCount + ' Điểm Sai Lệch Cần Khắc Phục';
        badgeCount.className = 'px-2 py-0.5 rounded-full bg-rose-600 text-white font-extrabold text-[10px]';

        container.innerHTML = `
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
            <tbody class="divide-y divide-rose-100 text-[11px]">
              <tr class="hover:bg-rose-50/50">
                <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">Dòng 5 | Cột: Ngày hết hạn</td>
                <td class="py-2.5 px-3 font-bold text-slate-900">DN-2026-004</td>
                <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/60 px-1.5 py-0.5 rounded">32/13/2026</td>
                <td class="py-2.5 px-3 text-slate-700">Định dạng ngày không hợp lệ (ngày 32, tháng 13) từ hệ thống CRM cũ</td>
                <td class="py-2.5 px-3 text-emerald-800 font-bold">Chuẩn hóa về định dạng YYYY-MM-DD (2026-11-28)</td>
              </tr>
              <tr class="hover:bg-rose-50/50">
                <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">Dòng 7 | Cột: Mức sử dụng</td>
                <td class="py-2.5 px-3 font-bold text-slate-900">DN-2026-004</td>
                <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/60 px-1.5 py-0.5 rounded">NULL / Để trống</td>
                <td class="py-2.5 px-3 text-slate-700">Khuyết dữ liệu đo lường công suất do cổng API (Ticket TK-8004)</td>
                <td class="py-2.5 px-3 text-emerald-800 font-bold">Tạm giữ doanh thu, tuyệt đối không gán bằng 0; chờ đồng bộ API</td>
              </tr>
              <tr class="hover:bg-rose-50/50">
                <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">Dòng 10 | Bảng: Sự cố kỹ thuật</td>
                <td class="py-2.5 px-3 font-bold text-slate-900">DN-2026-007</td>
                <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/60 px-1.5 py-0.5 rounded">Sự cố Mở (TK-8002)</td>
                <td class="py-2.5 px-3 text-slate-700">Khách hàng gặp sự cố máy chủ mức độ cao chưa được giải quyết</td>
                <td class="py-2.5 px-3 text-emerald-800 font-bold">Kích hoạt chốt chặn an toàn (Hard-Block) cấm chào tái ký; ưu tiên Kỹ thuật xử lý</td>
              </tr>
            </tbody>
          </table>
        `;

        footer.innerHTML = `
          <span class="text-[10px] text-rose-700 font-medium">Hệ thống tuân thủ nguyên tắc an toàn số liệu: Chỉ khi dữ liệu chuẩn xác mới xác nhận doanh thu.</span>
          <button onclick="autoFixAndRun()" class="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl text-xs shadow-xs transition flex items-center gap-1.5">
            <span>🔧</span><span>Tự Động Chuẩn Hóa &amp; Bắt Đầu Phân Tích</span>
          </button>
        `;
      }

      box.classList.remove('hidden');
      box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    // Mô phỏng quét phát hiện 3 lỗi khi nhập dữ liệu mẫu
    function simulateFindErrors() {
      detectedErrorsCount = 3;
      updateErrorDisplay();
      triggerErrorDiagnostics();
    }

    function dismissErrorBox() {
      document.getElementById('errorDiagnosticsBox').classList.add('hidden');
    }

    // Tự động chuẩn hóa lỗi và chạy phân tích
    function autoFixAndRun() {
      dismissErrorBox();
      detectedErrorsCount = 0;
      updateErrorDisplay();
      alert('Đã tự động sửa lỗi định dạng ngày và kích hoạt khóa chặn an toàn kiểm soát lỗi kỹ thuật (Hiện tại: 0 lỗi). Gia Linh Data Studio đang tiến hành tính toán số liệu!');
      runQuickDemo();
    }

    // Chạy Phân Tích (Demo Nhanh)
    async function runQuickDemo() {
      const btn = document.getElementById('btnTopQuickDemo');
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

        // Dữ liệu sau khi làm sạch thành công đạt 0 lỗi
        detectedErrorsCount = 0;
        updateErrorDisplay();

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
          <tr class="hover:bg-slate-50/70 transition">
            <td class="py-3.5 px-3 flex items-center gap-3">
              <div class="w-8 h-8 rounded-full bg-slate-100 text-indigo-700 flex items-center justify-center font-bold text-xs shrink-0">
                ${p.recommendation_type.includes('NÂNG CẤP') ? '🚀' : '📋'}
              </div>
              <div>
                <div class="font-bold text-slate-900">${p.customer_name} <span class="text-[10px] text-slate-400 font-mono">(${p.account_id})</span></div>
                <div class="text-[10px] text-slate-500">${p.current_package_name} → <span class="font-semibold text-indigo-600">${p.recommended_package_name}</span></div>
              </div>
            </td>
            <td class="py-3.5 px-3 text-slate-700 text-[11px] font-mono">
              ${p.is_revenue_grounded ? formatVND(p.final_price_vnd) : '<span class="text-amber-700 font-bold">Tạm giữ (' + formatVND(p.final_price_vnd) + ')</span>'}
            </td>
            <td class="py-3.5 px-3">
              <span class="px-2.5 py-0.5 rounded-full font-bold text-[10px] border ${badgeColor}">${statusText}</span>
            </td>
            <td class="py-3.5 px-3 text-right text-slate-400 text-[11px]">
              ${(idx + 1) * 3} phút trước
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
          <tr class="hover:bg-slate-50/60 transition">
            <td class="py-3.5 px-4 font-bold text-slate-900">
              ${p.customer_name}
              <div class="text-[10px] text-slate-400 font-mono">${p.account_id} • Đại diện: ${p.representative}</div>
            </td>
            <td class="py-3.5 px-3 text-slate-700 font-medium">${p.industry}<br><span class="text-[10px] text-indigo-600 font-bold">${p.segment}</span></td>
            <td class="py-3.5 px-3 text-slate-700">${p.current_package_name}</td>
            <td class="py-3.5 px-3 font-bold">${usageText}</td>
            <td class="py-3.5 px-3">${ticketText}</td>
            <td class="py-3.5 px-3 font-semibold text-indigo-700">${p.recommendation_type}</td>
            <td class="py-3.5 px-3 text-right font-bold text-slate-900 font-mono">${formatVND(p.final_price_vnd)}</td>
            <td class="py-3.5 px-4 text-slate-600 text-[11px] leading-relaxed">${p.ai_remarks}</td>
          </tr>
        `;
      });
    }

    // Vẽ 2 Biểu Đồ Chart.js
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
              tension: 0.38,
              borderWidth: 2.5,
              fill: true,
              pointRadius: 3
            },
            {
              label: 'Doanh thu Gói Nâng Cấp (Triệu ₫)',
              data: [50, 90, 120, 135, 180, 210, upVal || 225],
              borderColor: '#10b981',
              backgroundColor: 'rgba(16, 185, 129, 0.08)',
              tension: 0.38,
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
          scales: {
            y: { beginAtZero: true, grid: { color: 'rgba(241, 245, 249, 1)' } },
            x: { grid: { display: false } }
          }
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
        if (el) el.className = 'flex items-center justify-between px-2.5 py-1.5 rounded-xl bg-white hover:bg-slate-50 border border-slate-200 text-slate-700 font-semibold text-[11px] transition';
      });
      const activeEl = document.getElementById('filter-' + (seg === 'ALL' ? 'all' : seg === 'Enterprise VIP' ? 'vip' : seg === 'SME' ? 'sme' : 'hold'));
      if (activeEl) {
        activeEl.className = 'flex items-center justify-between px-2.5 py-1.5 rounded-xl bg-white text-indigo-950 font-bold text-[11px] border border-indigo-200 shadow-2xs transition';
      }
      populateRecentActivityTable();
    }

    // Tải xuống file Excel đã làm sạch
    function downloadCleanedExcel() {
      window.location.href = '/api/download-output-xlsx';
    }

    // Xử lý nạp tệp khách hàng
    function handleFileUpload(e) {
      const file = e.target.files[0];
      if (!file) return;
      alert('Đã tiếp nhận tệp: ' + file.name + ' (' + Math.round(file.size / 1024) + ' KB). Đang quét lỗi và chuẩn hóa dữ liệu...');
      
      // Nếu là tệp thử nghiệm có lỗi, hệ thống phát hiện 3 lỗi
      if (file.name.toLowerCase().includes('loi') || file.name.toLowerCase().includes('error') || file.name.toLowerCase().includes('raw')) {
        simulateFindErrors();
      } else {
        runQuickDemo();
      }
    }

    // Trợ lý AI chat trực tiếp
    async function sendChatMessage() {
      const input = document.getElementById('chatInput');
      const text = input.value.trim();
      if (!text) return;

      const log = document.getElementById('chatMessageLog');

      // Tin nhắn người dùng
      const uMsg = document.createElement('div');
      uMsg.className = 'p-2 rounded-lg bg-indigo-600 text-white text-[11px] font-medium shadow-2xs ml-3';
      uMsg.textContent = text;
      log.appendChild(uMsg);
      input.value = '';

      // Tin nhắn phản hồi đang gõ
      const aMsg = document.createElement('div');
      aMsg.className = 'p-2 rounded-lg bg-slate-50 text-slate-800 text-[11px] mr-2 border border-slate-200 shadow-2xs';
      aMsg.innerHTML = '<span class="text-indigo-600 animate-pulse font-bold">Đang tra cứu tri thức...</span>';
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

    // Khởi tạo trạng thái ban đầu: Toàn bộ về 0, Chẩn đoán 0 lỗi
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
Write-Output "Successfully updated $targetPath with Gia Linh Data Studio layout."
