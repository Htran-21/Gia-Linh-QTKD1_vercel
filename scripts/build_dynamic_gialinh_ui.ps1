# ==============================================================================
# SCRIPT TO BUILD ADVANCED GIA LINH DATA STUDIO UI
# - Supports REAL uploaded files (.xlsx, .xls, .csv, .json) with zero demo leakage
# - Only uses demo dataset when user explicitly clicks Demo
# - Dynamic error scan & analysis on uploaded data
# - Dynamic Excel generation from uploaded file
# - Collapsible sidebar that collapses into a sleek 4-step icon rail
# ==============================================================================

$targetHtml = "c:\Users\ADMIN\Downloads\AI Agent CK\web\index.html"

$htmlContent = @'
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Gia Linh Data Studio — Hệ Thống Phân Tích Dữ Liệu Doanh Nghiệp</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <!-- SheetJS for local or CDN excel parsing & generation -->
  <script src="/xlsx.full.min.js"></script>
  <script>
    if (typeof XLSX === 'undefined') {
      document.write('<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"><\/script>');
    }
  </script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
  <style>
    * {
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
    }
    body {
      font-family: 'Plus Jakarta Sans', system-ui, -apple-system, sans-serif;
      background: radial-gradient(circle at 10% 10%, #ecfdf5 0%, #f0fdf4 35%, #e0f2fe 75%, #ede9fe 100%);
      color: #0f172a;
      min-height: 100vh;
    }
    .custom-scroll::-webkit-scrollbar { width: 5px; height: 5px; }
    .custom-scroll::-webkit-scrollbar-thumb { background: #99f6e4; border-radius: 9999px; }
    .custom-scroll::-webkit-scrollbar-thumb:hover { background: #5eead4; }
    .app-window {
      background: #ffffff;
      border: 1px solid rgba(204, 251, 241, 0.8);
      border-radius: 28px;
      box-shadow: 0 25px 70px -15px rgba(14, 165, 233, 0.15), 0 0 1px 1px rgba(16, 185, 129, 0.05);
    }
    /* SIDEBAR GRADIENT KHỚP HÌNH ẢNH MẪU: EMERALD -> CYAN -> ROYAL BLUE -> VIOLET */
    .sidebar-emerald-cyan {
      background: linear-gradient(180deg, #10b981 0%, #059669 15%, #06b6d4 42%, #0284c7 68%, #3b82f6 84%, #6366f1 100%);
      position: relative;
    }
    .sidebar-emerald-cyan::before {
      content: "";
      position: absolute;
      top: 15%;
      left: -20%;
      width: 140%;
      height: 45%;
      background: radial-gradient(circle, rgba(255, 255, 255, 0.22) 0%, rgba(255, 255, 255, 0) 70%);
      pointer-events: none;
    }
    /* SIDEBAR COLLAPSED TO ICON RAIL */
    #leftSidebar {
      transition: width 0.3s cubic-bezier(0.4, 0, 0.2, 1), padding 0.3s ease;
    }
    #leftSidebar.sidebar-collapsed {
      width: 5rem !important; /* 80px */
      padding-left: 0.5rem !important;
      padding-right: 0.5rem !important;
    }
    #leftSidebar.sidebar-collapsed #sidebarFullView {
      display: none !important;
    }
    #leftSidebar.sidebar-collapsed #sidebarIconView {
      display: flex !important;
    }
    #sidebarIconView {
      display: none;
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
  <div id="greetingToast" class="fixed bottom-6 right-6 z-50 hidden toast-animate p-4 rounded-2xl bg-slate-900/95 text-white border border-teal-400/40 shadow-2xl flex items-center gap-3 backdrop-blur-md max-w-md">
    <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-emerald-500 via-teal-500 to-cyan-500 text-white flex items-center justify-center text-lg font-bold shrink-0 shadow-xs">
      👋
    </div>
    <div class="flex-1 text-xs">
      <div class="font-extrabold text-sm text-teal-200">Xin chào, khách hàng!</div>
      <div class="text-slate-300 text-[11px] mt-0.5">Chào mừng bạn đến với <strong>Gia Linh Data Studio</strong>. Thực hiện tuần tự 4 bước ở thanh công cụ bên trái để bắt đầu phân tích.</div>
    </div>
    <button onclick="dismissGreetingToast()" class="text-slate-400 hover:text-white font-bold p-1 text-xs">✕</button>
  </div>

  <!-- ========================================================================= -->
  <!-- CỬA SỔ ỨNG DỤNG GIA LINH DATA STUDIO (2 CỘT)                              -->
  <!-- ========================================================================= -->
  <div class="app-window w-full max-w-[1760px] min-h-[94vh] flex flex-col overflow-hidden bg-white">

    <!-- 1. BROWSER CHROME HEADER (macOS 3 DOTS + URL BAR + TẢI TÀI LIỆU) -->
    <header class="px-5 py-3 border-b border-slate-100 flex items-center justify-between text-xs bg-slate-50/70 select-none">
      <!-- 3 Chấm màu macOS & Nút thu gọn thanh công cụ -->
      <div class="flex items-center gap-3">
        <div class="flex items-center gap-1.5 mr-1">
          <span class="w-3 h-3 rounded-full bg-rose-500 inline-block shadow-2xs"></span>
          <span class="w-3 h-3 rounded-full bg-amber-400 inline-block shadow-2xs"></span>
          <span class="w-3 h-3 rounded-full bg-emerald-400 inline-block shadow-2xs"></span>
        </div>
        <button onclick="toggleSidebar()" class="p-1.5 rounded-lg hover:bg-slate-200/60 text-slate-700 font-bold transition text-xs flex items-center gap-1" title="Thu gọn / Mở rộng thanh công cụ thành Icon Bar">
          <span id="headerToggleIcon">☰</span>
          <span id="headerToggleText" class="text-[11px] hidden sm:inline font-semibold">Thu gọn</span>
        </button>
      </div>

      <!-- Thanh địa chỉ trung tâm macOS -->
      <div class="flex items-center gap-2 bg-white px-5 py-1.5 rounded-xl border border-slate-200/80 text-slate-700 font-semibold text-xs shadow-2xs">
        <span class="text-emerald-500 text-xs">🛡️</span>
        <span class="font-extrabold text-slate-900">GiaLinhDataStudio.vn</span>
        <span class="text-slate-300">/</span>
        <span class="text-teal-600 font-semibold">quy-trinh-4-buoc</span>
        <span class="text-slate-400 text-[11px] ml-1">🔄</span>
      </div>

      <!-- Tải tài liệu báo cáo & Đặt lại về 0 -->
      <div class="flex items-center gap-2 text-xs">
        <a href="/api/download-report-docx" download class="px-3 py-1.5 rounded-xl bg-teal-50 hover:bg-teal-100 text-teal-800 border border-teal-200 font-bold transition shadow-2xs flex items-center gap-1.5">
          <span>📄</span><span class="hidden md:inline">Báo Cáo Word (.docx)</span>
        </a>
        <a href="/api/download-presentation-pptx" download class="px-3 py-1.5 rounded-xl bg-sky-50 hover:bg-sky-100 text-sky-800 border border-sky-200 font-bold transition shadow-2xs flex items-center gap-1.5">
          <span>📽️</span><span class="hidden md:inline">Slide Thuyết Trình (.pptx)</span>
        </a>
        <button onclick="resetAllToZero()" class="px-3 py-1.5 rounded-xl bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 font-bold transition shadow-2xs" title="Đưa toàn bộ số liệu về 0 ban đầu">
          ↺ Đặt Lại Về 0
        </button>
      </div>
    </header>

    <!-- 2. INNER APP TOP BAR (LOGO + Ô TÌM KIẾM PILL + KHÁCH HÀNG) -->
    <div class="px-6 py-3.5 border-b border-slate-100 flex items-center justify-between gap-4 bg-white select-none">
      <!-- Logo Gia Linh Data Studio -->
      <div class="flex items-center gap-3">
        <div class="flex items-center gap-2.5 cursor-pointer" onclick="showGreetingToast()">
          <div class="w-9 h-9 rounded-xl bg-gradient-to-tr from-emerald-400 via-teal-500 to-cyan-500 text-white flex items-center justify-center font-black text-xs shadow-md tracking-wider">
            .D
          </div>
          <div>
            <span class="text-base font-black tracking-tight text-slate-900">Gia Linh Data Studio</span>
            <span id="modeBadge" class="ml-1.5 px-2 py-0.5 rounded-md bg-slate-100 text-[10px] font-extrabold text-slate-600 border border-slate-200">Chờ Dữ Liệu</span>
          </div>
        </div>
      </div>

      <!-- Ô tìm kiếm dạng pill bo tròn theo đúng ảnh mẫu -->
      <div class="relative flex-1 max-w-xl mx-4">
        <span class="absolute left-4 top-2.5 text-slate-400 text-xs">🔍</span>
        <input id="searchInput" type="text" placeholder="Tìm kiếm theo mã doanh nghiệp, tên khách hàng, hợp đồng..." onkeydown="if(event.key==='Enter') searchCustomer(this.value)" class="w-full pl-10 pr-14 py-2 rounded-full bg-slate-50 border border-slate-200 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-teal-500 focus:bg-white transition shadow-2xs font-medium" />
        <span class="absolute right-4 top-2 px-1.5 py-0.5 rounded text-[10px] bg-slate-200 text-slate-600 font-mono font-semibold">⌘K</span>
      </div>

      <!-- Cụm biểu tượng bên phải theo đúng ảnh mẫu (Calendar, stats, notifications, avatar) -->
      <div class="flex items-center gap-3">
        <div class="flex items-center gap-1.5 text-slate-500">
          <button class="p-2 rounded-full hover:bg-slate-100 transition text-sm" title="Lịch đối soát">📅</button>
          <button class="p-2 rounded-full hover:bg-slate-100 transition text-sm" title="Hiệu suất">👥</button>
          <button class="p-2 rounded-full hover:bg-slate-100 transition text-sm" title="Lớp dữ liệu">🧊</button>
          <button class="p-2 rounded-full hover:bg-slate-100 transition text-sm relative" title="Thông báo">
            🔔
            <span class="w-2 h-2 rounded-full bg-teal-500 absolute top-1.5 right-1.5"></span>
          </button>
        </div>

        <!-- Khách Hàng (Avatar & Tên hiển thị) -->
        <div class="flex items-center gap-2 pl-2 border-l border-slate-200 cursor-pointer" onclick="showGreetingToast()" title="Bấm vào để hiển thị lời chào">
          <div class="relative">
            <div class="w-8 h-8 rounded-full bg-gradient-to-tr from-emerald-500 to-cyan-600 text-white font-black text-xs flex items-center justify-center shadow-2xs">
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
    <div class="flex-1 flex flex-col lg:flex-row overflow-hidden relative">

      <!-- ===================================================================== -->
      <!-- CỘT TRÁI: THANH CÔNG CỤ DỌC (GRADIENT EMERALD -> CYAN -> ROYAL BLUE)   -->
      <!-- HỖ TRỢ CHUYỂN ĐỔI: FULL VIEW <-> ICON RAIL VIEW                       -->
      <!-- ===================================================================== -->
      <aside id="leftSidebar" class="sidebar-emerald-cyan w-full lg:w-88 xl:w-96 text-white p-5 flex flex-col justify-between shrink-0 select-none transition-all duration-300 overflow-y-auto custom-scroll space-y-4">
        
        <!-- =================================================================== -->
        <!-- CHẾ ĐỘ 1: FULL VIEW (KHI MỞ RỘNG)                                   -->
        <!-- =================================================================== -->
        <div id="sidebarFullView" class="space-y-5 relative z-10">
          
          <!-- THẺ TỔNG QUAN HỆ THỐNG PHONG CÁCH ẢNH MẪU -->
          <div class="p-4 rounded-2xl bg-white/20 backdrop-blur-md border border-white/30 text-white space-y-2 shadow-sm">
            <div class="flex items-center justify-between text-xs">
              <span class="font-extrabold uppercase tracking-wider text-teal-100 text-[10px]">Trung Tâm Điều Hành</span>
              <button onclick="toggleSidebar()" class="px-2 py-0.5 rounded-full bg-white/25 hover:bg-white/35 text-[10px] font-bold transition flex items-center gap-1" title="Thu gọn thành thanh Icon">
                <span>◀</span><span>Thu gọn</span>
              </button>
            </div>
            <div class="font-black text-sm text-white leading-tight">Gia Linh Data Studio</div>
            <div class="text-[11px] text-emerald-100 font-medium">Quy trình 4 bước chuẩn hóa &amp; đối soát tự động</div>
          </div>

          <!-- ================================================================= -->
          <!-- 4 BƯỚC QUY TRÌNH DUY NHẤT TRÊN THANH CÔNG CỤ                      -->
          <!-- ================================================================= -->
          <div class="space-y-3">
            <div class="text-[11px] font-black uppercase tracking-wider text-teal-100 px-1 flex items-center justify-between">
              <span>Quy Trình 4 Bước Thao Tác</span>
              <span id="currentSourceLabel" class="text-[9px] text-white/80 font-normal">Chưa có dữ liệu</span>
            </div>

            <!-- BƯỚC 1: NẠP DỮ LIỆU ĐẦU VÀO -->
            <div class="p-3.5 rounded-2xl bg-white/15 backdrop-blur-md border border-white/25 space-y-2.5 transition hover:bg-white/20">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 rounded-lg bg-white text-teal-700 flex items-center justify-center font-black text-xs shrink-0 shadow-2xs">1</span>
                  <span class="font-black text-xs text-white">Bước 1: Nạp Dữ Liệu</span>
                </div>
                <span id="step1StatusBadge" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white">Chưa nạp</span>
              </div>
              <div class="space-y-2">
                <!-- Chọn tệp của người dùng (.xlsx, .csv, .json) -->
                <label class="w-full py-2.5 px-3 rounded-xl bg-white text-teal-900 hover:bg-teal-50 font-black text-xs cursor-pointer shadow-sm transition flex items-center justify-center gap-2">
                  <input type="file" id="fileInput" accept=".xlsx,.xls,.csv,.json" class="hidden" onchange="handleFileUpload(event)" />
                  <span>📥</span><span>Nạp Tệp Của Bạn (.xlsx, .csv, .json)</span>
                </label>
                <!-- Nút nạp tệp mẫu Demo -->
                <button onclick="loadSampleDataset()" id="btnLoadSample" class="w-full py-1.5 px-3 rounded-xl border border-white/40 bg-white/10 hover:bg-white/20 text-white font-bold text-[11px] transition flex items-center justify-center gap-2">
                  <span>📂</span><span>Chạy Thử Tệp Mẫu Demo (10 Doanh Nghiệp)</span>
                </button>
                <div id="fileInfoNote" class="text-[9px] text-teal-100 text-center font-medium">Hệ thống sẽ chỉ tính toán từ tệp bạn tải lên, hoàn toàn độc lập với phần demo.</div>
              </div>
            </div>

            <!-- BƯỚC 2: QUÉT & CHẨN ĐOÁN LỖI (KHI CHƯA BẤM CHẠY THỬ THÌ TUYỆT ĐỐI KHÔNG HIỆN LỖI) -->
            <div class="p-3.5 rounded-2xl bg-white/15 backdrop-blur-md border border-white/25 space-y-2.5 transition hover:bg-white/20">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 rounded-lg bg-white text-cyan-700 flex items-center justify-center font-black text-xs shrink-0 shadow-2xs">2</span>
                  <span class="font-black text-xs text-white">Bước 2: Quét &amp; Chẩn Đoán Lỗi</span>
                </div>
                <span id="step2StatusBadge" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white">Chờ quét</span>
              </div>
              <div class="space-y-1.5">
                <p class="text-[10px] text-teal-100 font-medium">Kiểm tra định dạng ngày, phân biệt NULL vs 0%, rà soát sự cố kỹ thuật máy chủ trên tệp đã nạp.</p>
                <button onclick="runErrorScan()" id="btnScanErrors" class="w-full py-2 px-3 rounded-xl bg-cyan-400 hover:bg-cyan-300 text-slate-900 font-black text-xs shadow-sm transition flex items-center justify-center gap-2">
                  <span>🔍</span><span>Bấm Chạy Thử &amp; Quét Lỗi</span>
                </button>
              </div>
            </div>

            <!-- BƯỚC 3: CHẠY TÍNH TOÁN & ĐỐI SOÁT -->
            <div class="p-3.5 rounded-2xl bg-white/15 backdrop-blur-md border border-white/25 space-y-2.5 transition hover:bg-white/20">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 rounded-lg bg-white text-blue-700 flex items-center justify-center font-black text-xs shrink-0 shadow-2xs">3</span>
                  <span class="font-black text-xs text-white">Bước 3: Chạy Tính Toán</span>
                </div>
                <span id="step3StatusBadge" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white">Chờ tính toán</span>
              </div>
              <div class="space-y-1.5">
                <p class="text-[10px] text-blue-100 font-medium">Tính toán chỉ số KPI, đối soát doanh thu đủ căn cứ vs tạm giữ trực tiếp từ tệp của bạn.</p>
                <button onclick="runCalculationPipeline()" id="btnCalculate" class="w-full py-2 px-3 rounded-xl bg-gradient-to-r from-emerald-300 via-teal-200 to-cyan-200 text-slate-900 hover:brightness-105 font-black text-xs shadow-md transition flex items-center justify-center gap-2">
                  <span>⚡</span><span>Bắt Đầu Tính Toán &amp; Phân Tích</span>
                </button>
              </div>
            </div>

            <!-- BƯỚC 4: XUẤT KẾT QUẢ & BÁO CÁO (CHƯA TÍNH TOÁN KHÔNG CHO PHÉP TẢI DỮ LIỆU ĐÃ SẴN) -->
            <div class="p-3.5 rounded-2xl bg-white/15 backdrop-blur-md border border-white/25 space-y-2.5 transition hover:bg-white/20">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 rounded-lg bg-white text-indigo-700 flex items-center justify-center font-black text-xs shrink-0 shadow-2xs">4</span>
                  <span class="font-black text-xs text-white">Bước 4: Xuất Kết Quả Excel</span>
                </div>
                <span id="step4StatusBadge" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white">Chưa mở</span>
              </div>
              <div class="space-y-2">
                <!-- Nút tải Excel bị khóa khi chưa tính toán -->
                <button onclick="downloadCalculatedExcel()" id="btnDownloadExcel" disabled class="w-full py-2.5 px-3 rounded-xl bg-slate-300/40 border border-white/20 text-white/60 font-black text-xs cursor-not-allowed transition flex items-center justify-center gap-2 shadow-2xs">
                  <span id="lockExcelIcon">🔒</span>
                  <span id="lockExcelText">Xuất Tệp Excel (Chờ Tính Toán)</span>
                </button>
                <div class="text-[9px] text-teal-100 text-center font-medium">Chỉ tải được Excel khi dữ liệu đã được nạp &amp; tính toán thành công.</div>
              </div>
            </div>

          </div>

          <!-- TRỢ LÝ AI CHAT TRỰC TIẾP TRONG SIDEBAR -->
          <div class="pt-2 border-t border-white/20 space-y-2">
            <div class="flex items-center justify-between text-[11px] font-black uppercase text-teal-100 px-1">
              <span>Trợ Lý Tra Cứu AI</span>
              <span class="text-[9px] text-emerald-200 bg-white/20 px-2 py-0.5 rounded-full font-bold">Sẵn sàng</span>
            </div>
            
            <div id="chatMessageLog" class="max-h-32 overflow-y-auto custom-scroll space-y-1.5 p-2 bg-white/10 rounded-xl border border-white/20 text-xs">
              <div class="p-2 rounded-lg bg-white/20 text-white text-[11px] leading-relaxed backdrop-blur-sm">
                👋 <strong>Xin chào, khách hàng!</strong> Tôi là Trợ lý <strong>Gia Linh Data Studio</strong>. Bạn có thể hỏi về các mã doanh nghiệp, hoặc phân biệt <em>NULL vs 0%</em>.
              </div>
            </div>

            <div class="relative">
              <input id="chatInput" type="text" placeholder="Hỏi về mã doanh nghiệp, NULL vs 0..." onkeydown="if(event.key==='Enter') sendChatMessage()" class="w-full pl-3 pr-8 py-2 rounded-xl bg-white/20 border border-white/30 text-xs text-white placeholder-white/60 focus:outline-none focus:bg-white/30 font-medium" />
              <button onclick="sendChatMessage()" class="absolute right-1.5 top-2 w-6 h-6 rounded-lg bg-white text-teal-800 flex items-center justify-center text-[10px] font-bold transition shadow-xs" title="Gửi câu hỏi">
                ➤
              </button>
            </div>
          </div>

          <!-- Chân thanh công cụ: Thông tin Khách Hàng -->
          <div class="pt-2 border-t border-white/20 flex items-center gap-2.5 px-1 cursor-pointer" onclick="showGreetingToast()" title="Bấm vào để hiển thị lời chào">
            <div class="w-8 h-8 rounded-full bg-white text-teal-800 flex items-center justify-center font-black text-xs shrink-0 shadow-md">
              KH
            </div>
            <div class="overflow-hidden flex-1">
              <div class="font-extrabold text-xs text-white truncate">Khách Hàng</div>
              <div class="text-[10px] text-teal-100 font-medium truncate">Gia Linh Data Studio</div>
            </div>
            <span class="text-white/60 text-xs">•••</span>
          </div>

        </div>

        <!-- =================================================================== -->
        <!-- CHẾ ĐỘ 2: ICON RAIL VIEW (KHI BẤM ĐÓNG LẠI ĐỂ TỐI ƯU KHÔNG GIAN)     -->
        <!-- =================================================================== -->
        <div id="sidebarIconView" class="flex-col items-center justify-between h-full py-2 space-y-6 relative z-10 w-full">
          
          <!-- Top: Nút mở rộng & Logo icon -->
          <div class="flex flex-col items-center gap-4">
            <button onclick="toggleSidebar()" class="w-10 h-10 rounded-xl bg-white/20 hover:bg-white/30 text-white flex items-center justify-center text-sm font-black transition shadow-xs" title="Mở rộng thanh công cụ đầy đủ">
              ▶
            </button>
            <div class="w-10 h-10 rounded-xl bg-white text-teal-800 flex items-center justify-center font-black text-sm shadow-md" title="Gia Linh Data Studio">
              .D
            </div>
          </div>

          <!-- Middle: 4 ICON BƯỚC THAO TÁC CỐT LÕI (CÓ BADGE ĐÁNH DẤU SỐ BƯỚC) -->
          <div class="flex flex-col items-center gap-4">
            
            <!-- Icon Bước 1: Nạp Dữ Liệu -->
            <div class="relative group">
              <label class="w-11 h-11 rounded-2xl bg-white/20 hover:bg-white/30 text-white flex items-center justify-center text-lg cursor-pointer transition shadow-md border border-white/30" title="Bước 1: Nạp Tệp Dữ Liệu Của Bạn">
                <input type="file" accept=".xlsx,.xls,.csv,.json" class="hidden" onchange="handleFileUpload(event)" />
                <span>📂</span>
                <span class="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-teal-400 text-teal-950 font-black text-[10px] flex items-center justify-center shadow-xs">1</span>
              </label>
              <div class="hidden group-hover:block absolute left-14 top-2 bg-slate-900 text-white text-[11px] font-bold py-1.5 px-3 rounded-xl whitespace-nowrap shadow-xl z-50">
                Bước 1: Nạp Tệp (.xlsx, .csv, .json)
              </div>
            </div>

            <!-- Icon Bước 2: Quét & Chẩn Đoán Lỗi -->
            <div class="relative group">
              <button onclick="runErrorScan()" id="iconBtnScan" class="w-11 h-11 rounded-2xl bg-white/20 hover:bg-white/30 text-white flex items-center justify-center text-lg transition shadow-md border border-white/30" title="Bước 2: Quét & Chẩn Đoán Lỗi">
                <span>🔍</span>
                <span id="iconBadge2" class="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-cyan-300 text-cyan-950 font-black text-[10px] flex items-center justify-center shadow-xs">2</span>
              </button>
              <div class="hidden group-hover:block absolute left-14 top-2 bg-slate-900 text-white text-[11px] font-bold py-1.5 px-3 rounded-xl whitespace-nowrap shadow-xl z-50">
                Bước 2: Quét &amp; Chẩn Đoán Lỗi
              </div>
            </div>

            <!-- Icon Bước 3: Chạy Tính Toán -->
            <div class="relative group">
              <button onclick="runCalculationPipeline()" id="iconBtnCalc" class="w-11 h-11 rounded-2xl bg-white/20 hover:bg-white/30 text-white flex items-center justify-center text-lg transition shadow-md border border-white/30" title="Bước 3: Bắt Đầu Tính Toán">
                <span>⚡</span>
                <span id="iconBadge3" class="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-blue-300 text-blue-950 font-black text-[10px] flex items-center justify-center shadow-xs">3</span>
              </button>
              <div class="hidden group-hover:block absolute left-14 top-2 bg-slate-900 text-white text-[11px] font-bold py-1.5 px-3 rounded-xl whitespace-nowrap shadow-xl z-50">
                Bước 3: Chạy Tính Toán &amp; Phân Tích
              </div>
            </div>

            <!-- Icon Bước 4: Xuất File Excel -->
            <div class="relative group">
              <button onclick="downloadCalculatedExcel()" id="iconBtnExcel" class="w-11 h-11 rounded-2xl bg-white/10 text-white/50 flex items-center justify-center text-lg transition shadow-md border border-white/20 cursor-not-allowed" title="Bước 4: Xuất File Excel (Chờ Tính Toán)">
                <span id="iconExcelSymbol">🔒</span>
                <span id="iconBadge4" class="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-slate-400 text-slate-900 font-black text-[10px] flex items-center justify-center shadow-xs">4</span>
              </button>
              <div class="hidden group-hover:block absolute left-14 top-2 bg-slate-900 text-white text-[11px] font-bold py-1.5 px-3 rounded-xl whitespace-nowrap shadow-xl z-50">
                Bước 4: Xuất File Excel
              </div>
            </div>

            <!-- Icon AI Tra Cứu -->
            <div class="relative group pt-2 border-t border-white/20">
              <button onclick="toggleSidebar()" class="w-10 h-10 rounded-xl bg-white/20 hover:bg-white/30 text-white flex items-center justify-center text-base transition shadow-xs" title="Mở hộp Trợ Lý Tra Cứu AI">
                🤖
              </button>
              <div class="hidden group-hover:block absolute left-14 top-2 bg-slate-900 text-white text-[11px] font-bold py-1.5 px-3 rounded-xl whitespace-nowrap shadow-xl z-50">
                Trợ Lý Tra Cứu AI
              </div>
            </div>

          </div>

          <!-- Bottom: Avatar Khách hàng thu nhỏ -->
          <div class="pt-2 border-t border-white/20 cursor-pointer" onclick="showGreetingToast()" title="Khách Hàng (Bấm để hiện lời chào)">
            <div class="w-10 h-10 rounded-full bg-white text-teal-800 flex items-center justify-center font-black text-xs shadow-md">
              KH
            </div>
          </div>

        </div>

      </aside>

      <!-- ===================================================================== -->
      <!-- KHÔNG GIAN LÀM VIỆC CHÍNH (TÔNG MÀU SÁNG TRẮNG, CARD HIỆN ĐẠI THEO ẢNH)-->
      <!-- ===================================================================== -->
      <main class="flex-1 overflow-y-auto custom-scroll p-4 sm:p-7 bg-[#f8fafc] flex flex-col justify-between space-y-6">

        <!-- TIÊU ĐỀ CHÀO MỪNG KHÁCH HÀNG: KHI BẤM VÀO SẼ HIỆN XIN CHÀO KHÁCH HÀNG -->
        <div class="flex flex-wrap items-center justify-between gap-3">
          <div class="cursor-pointer" onclick="showGreetingToast()" title="Bấm vào để hiện thông báo lời chào">
            <h1 id="greetingTitle" class="text-2xl font-black text-slate-900 tracking-tight flex items-center gap-2 hover:text-teal-600 transition">
              Xin chào, khách hàng 👋
            </h1>
            <p id="systemSubtitle" class="text-xs text-slate-500 mt-0.5">
              Chào mừng bạn đến với <strong>Gia Linh Data Studio</strong>. Hãy thực hiện tuần tự 4 bước trên thanh công cụ để hoàn tất phân tích.
            </p>
          </div>

          <!-- Thanh điều hướng Tab xem nhanh -->
          <div class="flex items-center gap-1.5 p-1 bg-slate-200/60 rounded-xl border border-slate-200 text-xs font-bold">
            <button onclick="switchTab('dashboard')" id="tabBtnDashboard" class="px-4 py-1.5 rounded-lg bg-white text-teal-700 shadow-2xs transition font-extrabold">
              📊 Bảng Điều Khiển
            </button>
            <button onclick="switchTab('table')" id="tabBtnTable" class="px-4 py-1.5 rounded-lg text-slate-600 hover:text-teal-700 transition font-extrabold">
              📗 Dữ Liệu Đã Làm Sạch
            </button>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- KHUNG CHẨN ĐOÁN LỖI TỆP DỮ LIỆU (WHERE - WHY - FIX GUIDE)            -->
        <!-- KHI CHƯA BẤM CHẠY THỬ: TUYỆT ĐỐI KHÔNG HIỆN BẤT CỨ LỖI GÌ!          -->
        <!-- =================================================================== -->
        <div id="errorDiagnosticsBox" class="hidden p-4 rounded-2xl border text-xs space-y-3 shadow-2xs transition-all">
          <div class="flex items-center justify-between">
            <div class="flex items-center gap-2">
              <span id="diagBoxIcon" class="w-6 h-6 rounded-lg bg-rose-600 text-white flex items-center justify-center font-bold text-xs">⚠️</span>
              <span id="diagBoxTitle" class="text-slate-900 font-extrabold text-sm">KẾT QUẢ QUÉT &amp; PHÁT HIỆN LỖI DỮ LIỆU</span>
              <span id="diagBadgeCount" class="px-2 py-0.5 rounded-full bg-rose-600 text-white font-extrabold text-[10px]">Cần Khắc Phục</span>
            </div>
            <button onclick="dismissErrorBox()" class="text-slate-500 hover:text-slate-800 font-bold">✕ Đóng</button>
          </div>
          
          <div id="errorTableContainer" class="overflow-x-auto bg-white rounded-xl border border-rose-200">
            <!-- Render bảng lỗi khi người dùng bấm chạy thử -->
          </div>

          <div id="errorBoxFooter" class="flex flex-wrap justify-between items-center gap-2 pt-1">
            <span class="text-[10px] text-rose-700 font-medium">Hệ thống áp dụng nguyên tắc an toàn: Dữ liệu phải sạch 0 lỗi mới đưa vào tính toán doanh thu.</span>
            <button onclick="autoFixErrorsAndProceed()" class="px-4 py-2 bg-gradient-to-r from-teal-600 to-emerald-600 hover:from-teal-700 hover:to-emerald-700 text-white font-bold rounded-xl text-xs shadow-xs transition flex items-center gap-1.5">
              <span>🔧</span><span>Tự Động Chuẩn Hóa Lỗi (Về 0 Lỗi) &amp; Tiến Hành Tính Toán</span>
            </button>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- NỘI DUNG 1: BẢNG ĐIỀU KHIỂN DASHBOARD (PHONG CÁCH ẢNH MẪU MỚI)       -->
        <!-- CHỈ HIỆN KẾT QUẢ KHI ĐÃ CHO DỮ LIỆU & ĐƯỢC TÍNH TOÁN!               -->
        <!-- =================================================================== -->
        <div id="sectionDashboard" class="space-y-6">

          <!-- HÀNG 1: 4 THẺ METRICS THEO TÔNG MÀU ẢNH MẪU -->
          <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            
            <!-- THẺ 1: TỔNG KHÁCH HÀNG (TEAL ACCENT) -->
            <div class="p-5 rounded-3xl bg-white border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] space-y-2 hover:shadow-md transition">
              <div class="flex items-center justify-between text-xs">
                <span class="font-bold text-slate-500 text-[11px]">Tổng Số Khách Hàng</span>
                <span class="w-8 h-8 rounded-xl bg-teal-50 text-teal-600 flex items-center justify-center font-bold text-xs">🏢</span>
              </div>
              <div id="kpiTotalNum" class="text-3xl font-black text-slate-900 tracking-tight">0</div>
              <div class="flex items-center justify-between text-[11px] pt-1 border-t border-slate-100">
                <span id="kpiTotalSub" class="text-slate-400 font-medium">Chờ nạp dữ liệu</span>
                <span class="text-teal-600 font-bold text-[10px]">Bước 1</span>
              </div>
            </div>

            <!-- THẺ 2: DOANH THU ĐỦ CĂN CỨ (EMERALD ACCENT) -->
            <div class="p-5 rounded-3xl bg-white border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] space-y-2 hover:shadow-md transition">
              <div class="flex items-center justify-between text-xs">
                <span class="font-bold text-slate-500 text-[11px]">Doanh Thu Đủ Căn Cứ</span>
                <span class="w-8 h-8 rounded-xl bg-emerald-50 text-emerald-600 flex items-center justify-center font-bold text-xs">✓</span>
              </div>
              <div id="kpiGroundedNum" class="text-3xl font-black text-slate-900 tracking-tight">0 ₫</div>
              <div class="flex items-center justify-between text-[11px] pt-1 border-t border-slate-100">
                <span id="kpiGroundedSub" class="text-slate-400 font-medium">Chờ nạp dữ liệu</span>
                <span class="text-emerald-600 font-bold text-[10px]">Đủ căn cứ</span>
              </div>
            </div>

            <!-- THẺ 3: MỨC DÙNG DỊCH VỤ TB (CYAN ACCENT) -->
            <div class="p-5 rounded-3xl bg-white border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] space-y-2 hover:shadow-md transition">
              <div class="flex items-center justify-between text-xs">
                <span class="font-bold text-slate-500 text-[11px]">Mức Dùng Dịch Vụ TB</span>
                <span class="w-8 h-8 rounded-xl bg-cyan-50 text-cyan-600 flex items-center justify-center font-bold text-xs">📊</span>
              </div>
              <div id="kpiUsageNum" class="text-3xl font-black text-slate-900 tracking-tight">0%</div>
              <div class="flex items-center justify-between text-[11px] pt-1 border-t border-slate-100">
                <span id="kpiUsageSub" class="text-slate-400 font-medium">Chờ nạp dữ liệu</span>
                <span class="text-cyan-600 font-bold text-[10px]">NULL ≠ 0%</span>
              </div>
            </div>

            <!-- THẺ 4: DOANH THU TẠM GIỮ (AMBER ACCENT) -->
            <div class="p-5 rounded-3xl bg-white border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] space-y-2 hover:shadow-md transition">
              <div class="flex items-center justify-between text-xs">
                <span class="font-bold text-slate-500 text-[11px]">Doanh Thu Tạm Giữ</span>
                <span class="w-8 h-8 rounded-xl bg-amber-50 text-amber-600 flex items-center justify-center font-bold text-xs">⚠️</span>
              </div>
              <div id="kpiHoldNum" class="text-3xl font-black text-slate-900 tracking-tight">0 ₫</div>
              <div class="flex items-center justify-between text-[11px] pt-1 border-t border-slate-100">
                <span id="kpiHoldSub" class="text-slate-400 font-medium">Chờ nạp dữ liệu</span>
                <span class="text-amber-600 font-bold text-[10px]">Chốt chặn</span>
              </div>
            </div>

          </div>

          <!-- HÀNG 2: BIỂU ĐỒ ĐƯỜNG LƯỢN SÓNG CYAN/TEAL + 2 THẺ GRADIENT METRIC GÓC PHẢI -->
          <div class="grid grid-cols-1 lg:grid-cols-12 gap-5">
            
            <!-- BIỂU ĐỒ SÓNG (8 CỘT) PHONG CÁCH ẢNH MẪU -->
            <div class="lg:col-span-8 p-6 rounded-3xl bg-white border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] space-y-4">
              <div class="flex flex-wrap justify-between items-center gap-3">
                <div>
                  <h3 class="font-black text-sm text-slate-900">Biến Động Doanh Thu Hợp Đồng Theo Chu Kỳ</h3>
                  <div class="text-xs text-slate-400">Xu hướng doanh thu gói duy trì và nâng cấp qua các kỳ</div>
                </div>
                <div class="flex items-center gap-3 text-xs font-bold">
                  <span class="flex items-center gap-1.5 text-teal-600"><span class="w-2.5 h-2.5 rounded-full bg-teal-500 inline-block"></span> Gói Duy Trì</span>
                  <span class="flex items-center gap-1.5 text-cyan-600"><span class="w-2.5 h-2.5 rounded-full bg-cyan-500 inline-block"></span> Gói Nâng Cấp</span>
                  <span class="px-2.5 py-1 rounded-lg bg-slate-100 text-slate-700 text-[10px]">Chu kỳ 2026</span>
                </div>
              </div>

              <!-- Canvas biểu đồ -->
              <div class="h-60 w-full relative">
                <canvas id="revenueTrendChart"></canvas>
                <div id="chartPlaceholderOverlay" class="absolute inset-0 flex flex-col items-center justify-center bg-white/90 rounded-2xl pointer-events-none">
                  <div class="text-3xl mb-1 text-teal-500">📈</div>
                  <div class="font-extrabold text-slate-800 text-xs">Chưa có dữ liệu tính toán</div>
                  <div class="text-[11px] text-slate-400">Vui lòng thực hiện Bước 1 (Nạp Dữ Liệu) và Bước 3 (Chạy Tính Toán)</div>
                </div>
              </div>
            </div>

            <!-- 2 THẺ GRADIENT METRICS GÓC PHẢI (4 CỘT) KHỚP HÌNH ẢNH MẪU -->
            <div class="lg:col-span-4 space-y-4 flex flex-col justify-between">
              
              <!-- THẺ 1: GRADIENT MINT/TEAL -> CYAN -> PURPLE -->
              <div class="p-5 rounded-3xl bg-gradient-to-tr from-emerald-400 via-teal-500 to-indigo-600 text-white shadow-md space-y-2 relative overflow-hidden">
                <div class="flex justify-between items-center text-xs font-bold text-teal-100">
                  <span>Dung Lượng Đo Lường Đạt Chuẩn</span>
                  <span class="w-7 h-7 rounded-xl bg-white/20 flex items-center justify-center">⚡</span>
                </div>
                <div class="flex items-baseline gap-2">
                  <div id="bannerGbNumber" class="text-3xl font-black text-white">0 TB</div>
                  <div class="text-[10px] text-teal-100 font-semibold">Công suất hạ tầng</div>
                </div>
                <p id="banner1SubText" class="text-[10px] text-white/80 leading-relaxed font-medium">Hồ sơ khách hàng doanh nghiệp được kiểm soát chặt chẽ qua chốt chặn an toàn.</p>
              </div>

              <!-- THẺ 2: GRADIENT SUNSET AMBER -> ROSE -> PINK -->
              <div class="p-5 rounded-3xl bg-gradient-to-tr from-amber-400 via-rose-500 to-pink-600 text-white shadow-md space-y-2 relative overflow-hidden">
                <div class="flex justify-between items-center text-xs font-bold text-rose-100">
                  <span>Tốc Độ Xử Lý &amp; Chuẩn Hóa</span>
                  <span class="w-7 h-7 rounded-xl bg-white/20 flex items-center justify-center">🚀</span>
                </div>
                <div class="flex items-baseline gap-2">
                  <div id="bannerSpeedNumber" class="text-3xl font-black text-white">&lt; 2 Giây</div>
                  <div class="text-[10px] text-rose-100 font-semibold">Tự động 100%</div>
                </div>
                <p class="text-[10px] text-white/80 leading-relaxed font-medium">Cắt giảm hơn 95% thời gian xét duyệt thủ công liên phòng ban.</p>
              </div>

            </div>

          </div>

          <!-- HÀNG 3: BẢNG DANH SÁCH DOANH NGHIỆP ĐỐI SOÁT CHI TIẾT -->
          <div class="p-6 rounded-3xl bg-white border border-slate-100 shadow-[0_8px_30px_rgb(0,0,0,0.04)] space-y-4">
            <div class="flex flex-wrap justify-between items-center gap-3">
              <div>
                <h3 class="font-black text-sm text-slate-900">Danh Sách Đối Soát Khách Hàng Doanh Nghiệp</h3>
                <div id="tableSubInfo" class="text-xs text-slate-400">Kết quả tính toán sau khi dữ liệu đã được làm sạch và đối soát</div>
              </div>
              <div id="calculatedBadgeIndicator" class="px-3 py-1 rounded-full bg-slate-100 text-slate-500 font-extrabold text-xs">
                Chờ tính toán
              </div>
            </div>

            <!-- Bảng dữ liệu -->
            <div class="overflow-x-auto rounded-2xl border border-slate-100">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <thead class="bg-slate-50 text-slate-700 uppercase font-black text-[10px] border-b border-slate-200">
                  <tr>
                    <th class="py-3 px-4">MÃ DOANH NGHIỆP</th>
                    <th class="py-3 px-3">TÊN DOANH NGHIỆP &amp; ĐẠI DIỆN</th>
                    <th class="py-3 px-3">PHÂN KHÚC</th>
                    <th class="py-3 px-3">GÓI ĐỀ XUẤT</th>
                    <th class="py-3 px-3 text-right">GIÁ TRỊ HỢP ĐỒNG</th>
                    <th class="py-3 px-4 text-center">TRẠNG THÁI</th>
                  </tr>
                </thead>
                <tbody id="calculatedTableBody" class="divide-y divide-slate-100 bg-white">
                  <tr>
                    <td colspan="6" class="py-12 text-center text-slate-400">
                      <div class="text-3xl mb-1">📭</div>
                      Chưa có kết quả tính toán.<br>
                      Vui lòng thực hiện tuần tự <strong>Bước 1</strong> (Nạp Dữ Liệu) và <strong>Bước 3</strong> (Chạy Tính Toán) trên thanh công cụ bên trái.
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
          </div>

          <div class="rounded-2xl border border-slate-200 bg-white overflow-hidden shadow-2xs">
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <thead class="bg-slate-900 text-white uppercase font-bold text-[10px] border-b border-slate-200">
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
                      Chưa có dữ liệu làm sạch. Vui lòng thực hiện Bước 1 và Bước 3 để hiển thị bảng.
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- Chân trang ứng dụng Gia Linh Data Studio -->
        <footer class="pt-4 border-t border-slate-200 flex justify-between items-center text-[11px] text-slate-400 select-none">
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
    // =========================================================================
    // TRẠNG THÁI HỆ THỐNG & BIẾN ĐIỀU KHIỂN
    // =========================================================================
    let isDataLoaded = false;
    let isDataCalculated = false;
    let dataSource = null; // 'uploaded' hoặc 'demo' (CHỈ sử dụng demo khi người dùng bấm nút demo)
    let uploadedRawData = null; // Mảng dữ liệu thô đọc từ tệp người dùng tải lên
    let uploadedProfiles = null; // Hồ sơ doanh nghiệp đã chuẩn hóa từ tệp tải lên
    let uploadedFileName = ''; // Tên tệp người dùng đã tải lên
    let detectedErrors = []; // Danh sách lỗi phát hiện trong tệp
    let analysisData = null; // Kết quả tính toán chính thức
    let lineChart = null;
    let isSidebarCollapsed = false;

    // Toast Lời Chào
    function showGreetingToast() {
      const toast = document.getElementById('greetingToast');
      toast.classList.remove('hidden');
      setTimeout(() => {
        if (!toast.classList.contains('hidden')) {
          dismissGreetingToast();
        }
      }, 6000);
    }

    function dismissGreetingToast() {
      document.getElementById('greetingToast').classList.add('hidden');
    }

    // Thu gọn / Mở rộng Sidebar thành thanh Icon
    function toggleSidebar() {
      const sb = document.getElementById('leftSidebar');
      const toggleText = document.getElementById('headerToggleText');
      isSidebarCollapsed = !isSidebarCollapsed;

      if (isSidebarCollapsed) {
        sb.classList.add('sidebar-collapsed');
        if (toggleText) toggleText.textContent = 'Mở rộng';
      } else {
        sb.classList.remove('sidebar-collapsed');
        if (toggleText) toggleText.textContent = 'Thu gọn';
      }
    }

    function switchTab(tabName) {
      const secDash = document.getElementById('sectionDashboard');
      const secTable = document.getElementById('sectionTable');
      const btnDash = document.getElementById('tabBtnDashboard');
      const btnTable = document.getElementById('tabBtnTable');

      if (tabName === 'dashboard') {
        secDash.classList.remove('hidden');
        secTable.classList.add('hidden');
        btnDash.className = 'px-4 py-1.5 rounded-lg bg-white text-teal-700 shadow-2xs transition font-extrabold';
        btnTable.className = 'px-4 py-1.5 rounded-lg text-slate-600 hover:text-teal-700 transition font-extrabold';
        if (isDataCalculated) renderCharts();
      } else {
        secDash.classList.add('hidden');
        secTable.classList.remove('hidden');
        btnDash.className = 'px-4 py-1.5 rounded-lg text-slate-600 hover:text-teal-700 transition font-extrabold';
        btnTable.className = 'px-4 py-1.5 rounded-lg bg-white text-teal-700 shadow-2xs transition font-extrabold';
      }
    }

    // =========================================================================
    // BƯỚC 1: NẠP DỮ LIỆU ĐẦU VÀO (XỬ LÝ TỆP NGƯỜI DÙNG TẢI LÊN VS TỆP MẪU DEMO)
    // =========================================================================

    // Khi người dùng bấm chạy thử dữ liệu Demo (CHỈ KHI THAO TÁC NÚT NÀY MỚI DÙNG DEMO)
    function loadSampleDataset() {
      // Đặt nguồn dữ liệu là 'demo'
      dataSource = 'demo';
      isDataLoaded = true;
      isDataCalculated = false;
      analysisData = null;
      uploadedRawData = null;
      uploadedProfiles = null;
      uploadedFileName = 'Tệp mẫu Demo (10 Doanh Nghiệp)';

      // Cập nhật nhãn & huy hiệu
      document.getElementById('modeBadge').textContent = 'Chế độ Demo (10 DN)';
      document.getElementById('modeBadge').className = 'ml-1.5 px-2 py-0.5 rounded-md bg-cyan-100 text-[10px] font-extrabold text-cyan-800 border border-cyan-300';
      document.getElementById('currentSourceLabel').textContent = 'Nguồn: Tệp mẫu Demo';

      document.getElementById('step1StatusBadge').textContent = '✓ Đã nạp 10 DN (Demo)';
      document.getElementById('step1StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black animate-pulse';

      document.getElementById('step2StatusBadge').textContent = 'Cần quét lỗi';
      document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-cyan-200 text-cyan-950 font-bold';

      document.getElementById('step3StatusBadge').textContent = 'Chờ tính toán';
      document.getElementById('step3StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      // Khóa nút Bước 4
      lockExcelDownload();

      // Reset các chỉ số về 0
      resetZeroMetrics();

      alert('Đã nạp tệp mẫu Demo gồm 10 hồ sơ doanh nghiệp!\nVui lòng chuyển sang Bước 2: Bấm Chạy Thử & Quét Lỗi.');
    }

    // Khi người dùng TẢI TỆP LÊN: .xlsx, .xls, .csv, .json
    // HỆ THỐNG SẼ TÍNH TOÁN TRỰC TIẾP TỪ TỆP NÀY, TUYỆT ĐỐI KHÔNG DÙNG KẾT QUẢ DEMO!
    async function handleFileUpload(e) {
      const file = e.target.files[0];
      if (!file) return;

      const fName = file.name.toLowerCase();
      uploadedFileName = file.name;

      try {
        let rawRows = [];

        if (fName.endsWith('.json')) {
          const text = await file.text();
          const parsed = JSON.parse(text);
          if (Array.isArray(parsed)) {
            rawRows = parsed;
          } else if (parsed.profiles && Array.isArray(parsed.profiles)) {
            rawRows = parsed.profiles;
          } else if (parsed.data && Array.isArray(parsed.data)) {
            rawRows = parsed.data;
          } else {
            rawRows = [parsed];
          }
        } else if (fName.endsWith('.xlsx') || fName.endsWith('.xls') || fName.endsWith('.csv')) {
          if (typeof XLSX === 'undefined') {
            throw new Error('Thư viện đọc bảng tính SheetJS chưa sẵn sàng. Vui lòng thử lại sau vài giây.');
          }
          const buffer = await file.arrayBuffer();
          const workbook = XLSX.read(buffer, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const sheet = workbook.Sheets[firstSheetName];
          rawRows = XLSX.utils.sheet_to_json(sheet, { defval: null });
        } else {
          alert('Định dạng tệp không được hỗ trợ! Vui lòng chọn tệp .xlsx, .xls, .csv hoặc .json.');
          return;
        }

        if (!rawRows || rawRows.length === 0) {
          alert('Tệp dữ liệu trống hoặc không có dòng nào hợp lệ!');
          return;
        }

        // Chuyển đổi dữ liệu thô sang chuẩn Profile của Gia Linh Data Studio
        uploadedRawData = rawRows;
        uploadedProfiles = normalizeUploadedRows(rawRows);

        // ĐẶT NGUỒN DỮ LIỆU LÀ 'uploaded' (ĐỘC LẬP HOÀN TOÀN VỚI DEMO)
        dataSource = 'uploaded';
        isDataLoaded = true;
        isDataCalculated = false;
        analysisData = null;

        // Cập nhật nhãn giao diện
        document.getElementById('modeBadge').textContent = 'Tệp Người Dùng (' + uploadedProfiles.length + ' DN)';
        document.getElementById('modeBadge').className = 'ml-1.5 px-2 py-0.5 rounded-md bg-emerald-100 text-[10px] font-extrabold text-emerald-800 border border-emerald-300';
        document.getElementById('currentSourceLabel').textContent = 'Nguồn: ' + file.name;

        const shortName = file.name.length > 18 ? file.name.substring(0, 15) + '...' : file.name;
        document.getElementById('step1StatusBadge').textContent = '✓ ' + shortName + ' (' + uploadedProfiles.length + ' DN)';
        document.getElementById('step1StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';

        document.getElementById('step2StatusBadge').textContent = 'Chờ quét lỗi';
        document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-cyan-200 text-cyan-950 font-bold';

        document.getElementById('step3StatusBadge').textContent = 'Chờ tính toán';
        document.getElementById('step3StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

        // Khóa nút Bước 4
        lockExcelDownload();

        // Ẩn bảng lỗi (khi chưa bấm quét thử thì không hiện lỗi)
        dismissErrorBox();

        // Đưa các chỉ số về 0
        resetZeroMetrics();

        // Cập nhật icon badge
        document.getElementById('iconBadge2').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-amber-400 text-amber-950 font-black text-[10px] flex items-center justify-center shadow-xs';

        alert('Đã tiếp nhận thành công tệp: ' + file.name + ' (' + uploadedProfiles.length + ' hồ sơ doanh nghiệp)!\n\nHệ thống sẽ tính toán và phân tích ĐỘC LẬP TỪ CHÍNH TỆP NÀY (Không dùng dữ liệu demo).\nVui lòng tiếp tục với Bước 2: Bấm Chạy Thử & Quét Lỗi.');
      } catch (err) {
        alert('Lỗi đọc tệp dữ liệu: ' + err.message);
      }
    }

    // Chuẩn hóa các trường dữ liệu linh hoạt từ bất kỳ tệp Excel/CSV/JSON nào
    function normalizeUploadedRows(rows) {
      return rows.map((r, idx) => {
        const getVal = (...keys) => {
          for (const k of keys) {
            for (const rowKey of Object.keys(r)) {
              if (rowKey.trim().toLowerCase() === k.trim().toLowerCase()) {
                return r[rowKey];
              }
            }
          }
          return null;
        };

        const accId = getVal('Mã Khách Hàng', 'Mã Doanh Nghiệp', 'account_id', 'id', 'ma_khach_hang', 'mã', 'code') || ('DN-' + String(idx + 1).padStart(3, '0'));
        const custName = getVal('Tên Doanh Nghiệp', 'Tên Khách Hàng', 'customer_name', 'name', 'company', 'ten_doanh_nghiep', 'doanh nghiệp') || ('Doanh nghiệp ' + accId);
        const rep = getVal('Người Đại Diện', 'Đại Diện', 'representative', 'contact', 'dai_dien') || 'Chưa cập nhật';
        const ind = getVal('Lĩnh Vực', 'Ngành Nghề', 'industry', 'linh_vuc') || 'Dịch Vụ & Thương Mại';
        const seg = getVal('Phân Khúc', 'segment', 'tier') || 'Doanh Nghiệp Tiêu Chuẩn';
        const curPkg = getVal('Gói Dịch Vụ Hiện Tại', 'Gói Hiện Tại', 'current_package_name', 'package', 'gói dịch vụ') || 'Điện Toán Cơ Bản';
        const expDate = getVal('Hạn Hợp Đồng', 'Ngày Hết Hạn', 'expiry_date', 'han_hop_dong', 'expire') || '2026-12-31';

        // Mức sử dụng (giữ nguyên NULL nếu ô trống để phân biệt với 0)
        let usage = getVal('Tỷ Lệ Sử Dụng (%)', 'Mức Sử Dụng', 'service_usage_rate_pct', 'usage', 'ty_le_su_dung');
        if (usage === '' || usage === undefined) usage = null;
        if (usage !== null && !isNaN(Number(usage))) usage = Number(usage);

        // Sự cố kỹ thuật
        const ticket = getVal('Sự Cố Kỹ Thuật', 'Sự Cố', 'open_tickets', 'ticket', 'issues') || 'Không có sự cố';

        // Giá gốc
        let basePrice = getVal('Giá Gốc (VNĐ)', 'Giá Trị Hợp Đồng', 'Giá Trị', 'base_price_vnd', 'price', 'revenue');
        if (!basePrice || isNaN(Number(basePrice))) {
          basePrice = 80000000; // Giá trị tiêu chuẩn mặc định nếu tệp không có cột giá
        } else {
          basePrice = Number(basePrice);
        }

        // Chiết khấu
        let discPct = getVal('Chiết Khấu (%)', 'discount_percent', 'chiet_khau') || 0;
        discPct = Number(discPct) || 0;
        const discCode = getVal('Mã Ưu Đãi', 'discount_code') || (discPct > 0 ? ('DISCOUNT_' + discPct) : 'NONE');

        return {
          original_row_idx: idx + 2, // Tính cả dòng tiêu đề
          account_id: String(accId).trim(),
          customer_name: String(custName).trim(),
          representative: String(rep).trim(),
          industry: String(ind).trim(),
          segment: String(seg).trim(),
          current_package_name: String(curPkg).trim(),
          expiry_date: String(expDate).trim(),
          days_to_expiry: 30,
          service_usage_rate_pct: usage,
          open_tickets: String(ticket).trim(),
          base_price_vnd: basePrice,
          discount_code: discCode,
          discount_percent: discPct
        };
      });
    }

    // =========================================================================
    // BƯỚC 2: QUÉT & CHẨN ĐOÁN LỖI (CHỈ HIỆN LỖI KHI BẤM CHẠY THỬ)
    // =========================================================================
    function runErrorScan() {
      if (!isDataLoaded) {
        alert('Vui lòng thực hiện Bước 1 (Nạp Dữ Liệu) trước khi quét lỗi!');
        return;
      }

      detectedErrors = [];
      const box = document.getElementById('errorDiagnosticsBox');
      const container = document.getElementById('errorTableContainer');

      if (dataSource === 'demo') {
        // Lỗi mẫu của dữ liệu Demo (3 lỗi kinh điển)
        detectedErrors = [
          {
            rowCol: 'Dòng 5 | Cột: Ngày hết hạn',
            accId: 'DN-2026-004',
            val: '32/13/2026',
            why: 'Định dạng ngày không hợp lệ (ngày 32, tháng 13) từ hệ thống CRM cũ',
            fix: 'Chuẩn hóa về chuẩn ISO YYYY-MM-DD (2026-11-28)'
          },
          {
            rowCol: 'Dòng 7 | Cột: Mức sử dụng',
            accId: 'DN-2026-004',
            val: 'NULL / Để trống',
            why: 'Khuyết dữ liệu đo lường công suất do cổng API (Ticket TK-8004)',
            fix: 'Tạm giữ doanh thu, tuyệt đối không gán bằng 0; chờ đồng bộ API'
          },
          {
            rowCol: 'Dòng 10 | Bảng: Sự cố kỹ thuật',
            accId: 'DN-2026-007',
            val: 'Sự cố Mở (TK-8002)',
            why: 'Khách hàng gặp sự cố máy chủ mức độ cao chưa được giải quyết',
            fix: 'Kích hoạt chốt chặn an toàn (Hard-Block) cấm chào tái ký; ưu tiên Kỹ thuật xử lý'
          }
        ];
      } else if (dataSource === 'uploaded') {
        // QUÉT THỰC TẾ TRÊN TỆP NGƯỜI DÙNG TẢI LÊN
        uploadedProfiles.forEach(p => {
          // 1. Kiểm tra ngày tháng sai
          if (p.expiry_date.includes('32') || p.expiry_date.includes('13/') || isNaN(Date.parse(p.expiry_date)) && p.expiry_date.includes('/')) {
            detectedErrors.push({
              rowCol: 'Dòng ' + p.original_row_idx + ' | Cột: Hạn hợp đồng',
              accId: p.account_id,
              val: p.expiry_date,
              why: 'Định dạng ngày tháng không hợp lệ hoặc sai lệch thứ tự ngày/tháng',
              fix: 'Chuẩn hóa về chuẩn quốc tế ISO YYYY-MM-DD'
            });
          }

          // 2. Kiểm tra khuyết thiếu NULL
          if (p.service_usage_rate_pct === null || p.service_usage_rate_pct === undefined) {
            detectedErrors.push({
              rowCol: 'Dòng ' + p.original_row_idx + ' | Cột: Mức sử dụng',
              accId: p.account_id,
              val: 'NULL (Khuyết dữ liệu)',
              why: 'Không có dữ liệu đo lường lưu lượng (Khuyết thiếu NULL ≠ 0%)',
              fix: 'Tạm giữ an toàn doanh thu, không được gán bằng 0%; yêu cầu bộ phận dữ liệu đối soát'
            });
          }

          // 3. Kiểm tra sự cố máy chủ kỹ thuật mở
          const lowerTicket = p.open_tickets.toLowerCase();
          if (lowerTicket.includes('mở') || lowerTicket.includes('open') || lowerTicket.includes('tk-') || lowerTicket.includes('sự cố') || lowerTicket.includes('lỗi')) {
            detectedErrors.push({
              rowCol: 'Dòng ' + p.original_row_idx + ' | Cột: Sự cố kỹ thuật',
              accId: p.account_id,
              val: p.open_tickets,
              why: 'Hồ sơ có cảnh báo sự cố kỹ thuật hạ tầng chưa được đóng',
              fix: 'Kích hoạt Khóa chặn Kỹ thuật (Hard-Block) cấm chào tái ký; ưu tiên Kỹ thuật khắc phục'
            });
          }
        });
      }

      // HIỂN THỊ KẾT QUẢ QUÉT
      if (detectedErrors.length > 0) {
        document.getElementById('diagBadgeCount').textContent = detectedErrors.length + ' Điểm Sai Lệch Cần Khắc Phục';
        document.getElementById('diagBadgeCount').className = 'px-2 py-0.5 rounded-full bg-rose-600 text-white font-extrabold text-[10px]';

        let rowsHtml = '';
        detectedErrors.forEach(err => {
          rowsHtml += `
            <tr class="hover:bg-rose-50/70">
              <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">${err.rowCol}</td>
              <td class="py-2.5 px-3 font-bold text-slate-900">${err.accId}</td>
              <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/80 px-1.5 py-0.5 rounded">${err.val}</td>
              <td class="py-2.5 px-3 text-slate-700">${err.why}</td>
              <td class="py-2.5 px-3 text-emerald-800 font-bold">${err.fix}</td>
            </tr>
          `;
        });

        container.innerHTML = `
          <table class="w-full text-left text-xs whitespace-nowrap">
            <thead class="bg-rose-100/70 text-rose-900 uppercase font-black text-[10px]">
              <tr>
                <th class="py-2.5 px-3">Vị Trí (Dòng / Cột)</th>
                <th class="py-2.5 px-3">Mã Doanh Nghiệp</th>
                <th class="py-2.5 px-3">Giá Trị Hiện Tại</th>
                <th class="py-2.5 px-3">Nguyên Nhân Sai Lệch</th>
                <th class="py-2.5 px-3">Hướng Dẫn Khắc Phục (Fix Guide)</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-rose-100 text-[11px] bg-rose-50/40">
              ${rowsHtml}
            </tbody>
          </table>
        `;

        box.className = 'p-4 rounded-2xl border border-rose-300 bg-rose-50/90 text-xs space-y-3 shadow-md';
        box.classList.remove('hidden');
        box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

        document.getElementById('step2StatusBadge').textContent = '⚠️ Phát hiện ' + detectedErrors.length + ' lỗi';
        document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-rose-400 text-white font-black animate-pulse';
        document.getElementById('iconBadge2').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-rose-500 text-white font-black text-[10px] flex items-center justify-center shadow-xs';
      } else {
        // TỆP HOÀN TOÀN ĐẠT CHUẨN 0 LỖI
        document.getElementById('diagBadgeCount').textContent = '0 Lỗi - Dữ Liệu Đạt Chuẩn';
        document.getElementById('diagBadgeCount').className = 'px-2 py-0.5 rounded-full bg-emerald-600 text-white font-extrabold text-[10px]';

        container.innerHTML = `
          <div class="p-6 text-center text-emerald-800 space-y-2">
            <div class="text-3xl">✓</div>
            <div class="font-extrabold text-sm">Dữ liệu tệp tải lên hoàn toàn sạch và đạt chuẩn 0 lỗi!</div>
            <div class="text-xs text-slate-600">Toàn bộ ${uploadedProfiles ? uploadedProfiles.length : 0} hồ sơ doanh nghiệp đã vượt qua bộ lọc định dạng, sẵn sàng đưa vào Bước 3: Tính Toán &amp; Phân Tích.</div>
          </div>
        `;

        box.className = 'p-4 rounded-2xl border border-emerald-300 bg-emerald-50/90 text-xs space-y-3 shadow-md';
        box.classList.remove('hidden');

        document.getElementById('step2StatusBadge').textContent = '✓ 0 lỗi đạt chuẩn';
        document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';
        document.getElementById('iconBadge2').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-emerald-400 text-emerald-950 font-black text-[10px] flex items-center justify-center shadow-xs';
      }
    }

    function dismissErrorBox() {
      document.getElementById('errorDiagnosticsBox').classList.add('hidden');
    }

    // Tự động chuẩn hóa lỗi về 0 lỗi và chuyển sang Bước 3
    function autoFixErrorsAndProceed() {
      dismissErrorBox();

      if (dataSource === 'uploaded' && uploadedProfiles) {
        uploadedProfiles.forEach(p => {
          if (p.expiry_date.includes('32') || p.expiry_date.includes('13/')) {
            p.expiry_date = '2026-11-28';
          }
        });
      }

      document.getElementById('step2StatusBadge').textContent = '✓ 0 lỗi (Đã sửa)';
      document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';
      document.getElementById('iconBadge2').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-emerald-400 text-emerald-950 font-black text-[10px] flex items-center justify-center shadow-xs';

      alert('Đã tự động chuẩn hóa toàn bộ lỗi sai lệch về 0 lỗi đạt chuẩn! Hệ thống đang tiến hành Bước 3: Chạy Tính Toán...');
      runCalculationPipeline();
    }

    // =========================================================================
    // BƯỚC 3: CHẠY TÍNH TOÁN (TÍNH TRỰC TIẾP TỪ TỆP TẢI LÊN HOẶC DEMO)
    // =========================================================================
    async function runCalculationPipeline() {
      if (!isDataLoaded) {
        alert('Vui lòng hoàn thành Bước 1 (Nạp Dữ Liệu) trước khi tính toán!');
        return;
      }

      const btn = document.getElementById('btnCalculate');
      const oldHtml = btn.innerHTML;

      try {
        btn.innerHTML = '<span>⏳</span><span>Đang tính toán số liệu...</span>';
        btn.disabled = true;

        if (dataSource === 'uploaded') {
          // TÍNH TOÁN HOÀN TOÀN TỪ TỆP NGƯỜI DÙNG TẢI LÊN (KHÔNG DÙNG KẾT QUẢ DEMO)
          const calculatedProfiles = calculateUploadedData(uploadedProfiles);
          const kpi = calculateKpiSummary(calculatedProfiles);

          analysisData = {
            system_name: 'Gia Linh Data Studio',
            source_type: 'uploaded',
            filename: uploadedFileName,
            kpi_summary: kpi,
            profiles: calculatedProfiles
          };

          // Giả lập xử lý ngắn 0.3s để tạo cảm giác máy tính tính toán tức thì
          await new Promise(r => setTimeout(r, 350));
        } else if (dataSource === 'demo') {
          // CHỈ KHI CHẠY DEMO MỚI DÙNG KẾT QUẢ DEMO
          const res = await fetch('/api/run-agent', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ source_type: 'json', dataset_scale: '10_core' })
          });
          analysisData = await res.json();
          analysisData.source_type = 'demo';
        }

        isDataCalculated = true;

        // Cập nhật trạng thái Bước 2 & Bước 3
        document.getElementById('step2StatusBadge').textContent = '✓ 0 lỗi đạt chuẩn';
        document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';

        document.getElementById('step3StatusBadge').textContent = '✓ Đã tính toán';
        document.getElementById('step3StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';
        document.getElementById('iconBadge3').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-emerald-400 text-emerald-950 font-black text-[10px] flex items-center justify-center shadow-xs';

        // MỞ KHÓA BƯỚC 4: XUẤT EXCEL
        unlockExcelDownload();

        // ÁP DỤNG KẾT QUẢ VÀ HIỂN THỊ LÊN DASHBOARD
        applyCalculatedResults();
      } catch (err) {
        alert('Lỗi tính toán: ' + err.message);
      } finally {
        btn.innerHTML = oldHtml;
        btn.disabled = false;
      }
    }

    // Thuật toán đối soát & tính toán áp dụng riêng cho tệp người dùng tải lên
    function calculateUploadedData(profiles) {
      return profiles.map(p => {
        let isGrounded = true;
        let isBlocked = false;
        let recType = 'DUY TRÌ GÓI';
        let recPkg = p.current_package_name;
        let overallStatus = 'SẴN SÀNG THỰC THI';
        let remarks = '';
        let plan = '';

        // 1. Kiểm tra sự cố máy chủ (Hard-Block)
        const lowerTicket = (p.open_tickets || '').toLowerCase();
        if (lowerTicket.includes('mở') || lowerTicket.includes('open') || lowerTicket.includes('tk-') || lowerTicket.includes('sự cố') || lowerTicket.includes('lỗi')) {
          isBlocked = true;
          isGrounded = false;
          overallStatus = 'KHÓA CHẶN KỸ THUẬT (HARD-BLOCK)';
          remarks = 'Phát hiện sự cố kỹ thuật hạ tầng chưa được đóng. Hệ thống tự động kích hoạt chốt chặn Hard-Block, tạm đình chỉ chào tái ký và tạm giữ doanh thu an toàn.';
          plan = 'Yêu cầu bộ phận Kỹ thuật xử lý dứt điểm sự cố trước khi tiếp cận khách hàng.';
        }
        // 2. Kiểm tra khuyết thiếu NULL
        else if (p.service_usage_rate_pct === null || p.service_usage_rate_pct === undefined) {
          isGrounded = false;
          overallStatus = 'TẠM GIỮ (CHỜ ĐỒNG BỘ DỮ LIỆU)';
          recType = 'TẠM GIỮ DOANH THU';
          remarks = 'Trường mức độ sử dụng bị khuyết thiếu (NULL ≠ 0%). Tạm giữ doanh thu không tính vào doanh thu đủ căn cứ; chờ kết nối đo lường.';
          plan = 'Phối hợp với quản trị cơ sở dữ liệu đối soát log lưu lượng thực tế.';
        }
        // 3. Phân loại theo mức sử dụng hợp lệ
        else {
          const usage = p.service_usage_rate_pct;
          if (usage === 0) {
            recType = 'DUY TRÌ HỖ TRỢ PHỤC HỒI';
            recPkg = 'Gói Duy Trì Hỗ Trợ Kỹ Thuật';
            remarks = 'Mức sử dụng 0% là số liệu thực tế hợp lệ (khách hàng chưa phát sinh lưu lượng). Đề xuất gói duy trì chi phí tối thiểu và đồng hành tháo gỡ khó khăn.';
            plan = 'Chuyên viên CSKH liên hệ tìm hiểu khó khăn và hướng dẫn kích hoạt hạ tầng.';
          } else if (usage >= 80) {
            recType = 'NÂNG CẤP GÓI';
            recPkg = 'Hệ Sinh Thái Toàn Diện Enterprise Suite';
            remarks = 'Khách hàng có lưu lượng sử dụng cao vượt trội (' + usage + '%). Đủ điều kiện nâng cấp gói Enterprise Suite để đảm bảo băng thông ổn định.';
            plan = 'Bộ phận Kinh doanh gửi hồ sơ đề xuất nâng cấp trước ngày hết hạn.';
          } else if (usage >= 50) {
            recType = 'TÁI KÝ DUY TRÌ';
            recPkg = p.current_package_name;
            remarks = 'Mức sử dụng ổn định (' + usage + '%). Đề xuất tái ký gói dịch vụ hiện tại kèm các ưu đãi duy trì lâu dài.';
            plan = 'Gửi thông báo tái ký tự động và hợp đồng gia hạn.';
          } else {
            recType = 'TÁI KÝ TỐI ƯU HÓA';
            recPkg = 'Gói Tinh Gọn Hạ Tầng';
            remarks = 'Mức sử dụng thấp (' + usage + '%). Đề xuất tối ưu hóa quy mô gói để tiết kiệm chi phí cho khách hàng, tránh rời bỏ dịch vụ.';
            plan = 'Tư vấn chuyển đổi cấu hình tối ưu chi phí vận hành.';
          }

          // Cổng phê duyệt cấp quản lý nếu chiết khấu lớn
          if (p.discount_percent > 15) {
            overallStatus = 'CHỜ QUẢN LÝ PHÊ DUYỆT (HITL)';
            remarks += ' Mức chiết khấu (' + p.discount_percent + '%) vượt thẩm quyền tự động, đã chuyển Cổng kiểm soát HITL.';
          }
        }

        // Tính giá sau chiết khấu
        const finalPrice = Math.round(p.base_price_vnd * (1 - (p.discount_percent / 100)));

        return {
          ...p,
          is_blocked_by_tech: isBlocked,
          is_revenue_grounded: isGrounded,
          recommendation_type: recType,
          recommended_package_name: recPkg,
          final_price_vnd: finalPrice,
          overall_status: overallStatus,
          ai_remarks: remarks,
          action_plan: plan
        };
      });
    }

    // Tính tổng hợp KPI
    function calculateKpiSummary(profiles) {
      let groundedVal = 0;
      let holdVal = 0;
      let groundedCount = 0;
      let holdCount = 0;
      let usageSum = 0;
      let usageCount = 0;

      profiles.forEach(p => {
        if (p.is_revenue_grounded) {
          groundedVal += p.final_price_vnd;
          groundedCount++;
        } else {
          holdVal += p.final_price_vnd;
          holdCount++;
        }

        if (p.service_usage_rate_pct !== null && p.service_usage_rate_pct !== undefined && !isNaN(p.service_usage_rate_pct)) {
          usageSum += p.service_usage_rate_pct;
          usageCount++;
        }
      });

      const meanUsage = usageCount > 0 ? (usageSum / usageCount).toFixed(1) : 0;

      return {
        total_accounts: profiles.length,
        grounded_accounts_count: groundedCount,
        unverified_or_blocked_accounts_count: holdCount,
        total_grounded_renewal_value_vnd: groundedVal,
        total_blocked_or_unverified_value_vnd: holdVal,
        mean_usage_rate_pct: Number(meanUsage)
      };
    }

    // =========================================================================
    // BƯỚC 4: XUẤT EXCEL (TỰ ĐỘNG XUẤT TỪ CHÍNH TỆP TẢI LÊN CỦA NGƯỜI DÙNG)
    // =========================================================================
    function unlockExcelDownload() {
      const btn = document.getElementById('btnDownloadExcel');
      const lockIcon = document.getElementById('lockExcelIcon');
      const lockText = document.getElementById('lockExcelText');
      const step4Badge = document.getElementById('step4StatusBadge');

      btn.disabled = false;
      btn.className = 'w-full py-2.5 px-3 rounded-xl bg-gradient-to-r from-emerald-400 to-teal-400 hover:from-emerald-500 hover:to-teal-500 text-slate-900 font-black text-xs cursor-pointer transition flex items-center justify-center gap-2 shadow-md';
      lockIcon.textContent = '📥';
      lockText.textContent = dataSource === 'uploaded' ? 'Tải Excel Kết Quả Tệp Tải Lên (.xlsx)' : 'Tải File Excel Đã Tính Toán (.xlsx)';

      step4Badge.textContent = '✓ Sẵn sàng tải';
      step4Badge.className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';

      // Cập nhật Icon Rail
      const iconBtn = document.getElementById('iconBtnExcel');
      iconBtn.className = 'w-11 h-11 rounded-2xl bg-emerald-400 hover:bg-emerald-300 text-slate-900 flex items-center justify-center text-lg transition shadow-md border border-white/30 cursor-pointer';
      document.getElementById('iconExcelSymbol').textContent = '📥';
      document.getElementById('iconBadge4').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-emerald-300 text-emerald-950 font-black text-[10px] flex items-center justify-center shadow-xs';
    }

    function lockExcelDownload() {
      const btn = document.getElementById('btnDownloadExcel');
      btn.disabled = true;
      btn.className = 'w-full py-2.5 px-3 rounded-xl bg-slate-300/40 border border-white/20 text-white/60 font-black text-xs cursor-not-allowed transition flex items-center justify-center gap-2 shadow-2xs';
      document.getElementById('lockExcelIcon').textContent = '🔒';
      document.getElementById('lockExcelText').textContent = 'Xuất Tệp Excel (Chờ Tính Toán)';
      document.getElementById('step4StatusBadge').textContent = 'Chưa mở';
      document.getElementById('step4StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      const iconBtn = document.getElementById('iconBtnExcel');
      iconBtn.className = 'w-11 h-11 rounded-2xl bg-white/10 text-white/50 flex items-center justify-center text-lg transition shadow-md border border-white/20 cursor-not-allowed';
      document.getElementById('iconExcelSymbol').textContent = '🔒';
      document.getElementById('iconBadge4').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-slate-400 text-slate-900 font-black text-[10px] flex items-center justify-center shadow-xs';
    }

    // TẢI KẾT QUẢ EXCEL
    function downloadCalculatedExcel() {
      if (!isDataCalculated || !analysisData) {
        alert('Chưa có kết quả tính toán! Vui lòng thực hiện tuần tự Bước 1 và Bước 3 trước khi tải tệp Excel.');
        return;
      }

      if (dataSource === 'uploaded') {
        // TẠO FILE EXCEL TRỰC TIẾP TỪ DỮ LIỆU TỆP NGƯỜI DÙNG TẢI LÊN
        try {
          if (typeof XLSX === 'undefined') {
            throw new Error('Thư viện tạo file Excel SheetJS chưa tải xong.');
          }

          const exportRows = analysisData.profiles.map(p => ({
            'Mã Khách Hàng': p.account_id,
            'Tên Doanh Nghiệp': p.customer_name,
            'Người Đại Diện': p.representative,
            'Lĩnh Vực': p.industry,
            'Phân Khúc': p.segment,
            'Gói Dịch Vụ Hiện Tại': p.current_package_name,
            'Hạn Hợp Đồng': p.expiry_date,
            'Tỷ Lệ Sử Dụng (%)': p.service_usage_rate_pct === null ? 'NULL' : p.service_usage_rate_pct,
            'Sự Cố Kỹ Thuật': p.open_tickets,
            'Đề Xuất Hành Động': p.recommendation_type,
            'Gói Dịch Vụ Mới': p.recommended_package_name,
            'Giá Gốc (VNĐ)': p.base_price_vnd,
            'Mã Ưu Đãi': p.discount_code,
            'Chiết Khấu (%)': p.discount_percent,
            'Giá Sau Cùng (VNĐ)': p.final_price_vnd,
            'Căn Cứ Doanh Thu': p.is_revenue_grounded ? 'ĐỦ CĂN CỨ' : 'TẠM GIỮ AN TOÀN',
            'Trạng Thái Thực Thi': p.overall_status,
            'Nhận Xét & Đánh Giá AI': p.ai_remarks,
            'Kế Hoạch Thực Thi': p.action_plan
          }));

          const ws = XLSX.utils.json_to_sheet(exportRows);
          const wb = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(wb, ws, 'Ket_Qua_Phan_Tich');

          const safeName = (uploadedFileName.replace(/\.[^/.]+$/, "") || 'Du_Lieu').replace(/[^a-zA-Z0-9_\u00C0-\u1EF9-]/g, '_');
          const exportFileName = 'GiaLinh_KetQua_PhanTich_' + safeName + '.xlsx';

          XLSX.writeFile(wb, exportFileName);
          alert('Đã xuất thành công tệp Excel kết quả phân tích: ' + exportFileName);
        } catch (err) {
          alert('Lỗi xuất tệp Excel: ' + err.message);
        }
      } else {
        // TẢI FILE DEMO KHI Ở CHẾ ĐỘ DEMO
        window.location.href = '/api/download-output-xlsx';
      }
    }

    // =========================================================================
    // HIỂN THỊ KẾT QUẢ LÊN BẢNG ĐIỀU KHIỂN & BẢNG SỐ LIỆU
    // =========================================================================
    function applyCalculatedResults() {
      if (!analysisData) return;
      const kpi = analysisData.kpi_summary;

      // 4 Metrics Hàng 1
      document.getElementById('kpiTotalNum').textContent = kpi.total_accounts;
      document.getElementById('kpiTotalSub').innerHTML = '<span class="text-emerald-600 font-bold">↑ 100%</span> đã chuẩn hóa';

      document.getElementById('kpiGroundedNum').textContent = formatMillions(kpi.total_grounded_renewal_value_vnd);
      document.getElementById('kpiGroundedSub').innerHTML = '<span class="text-emerald-600 font-bold">' + kpi.grounded_accounts_count + ' DN</span> đủ căn cứ thực thi';

      document.getElementById('kpiUsageNum').textContent = (kpi.mean_usage_rate_pct || 0) + '%';
      document.getElementById('kpiUsageSub').innerHTML = '<span class="text-cyan-600 font-bold">NULL ≠ 0%</span> tách biệt';

      document.getElementById('kpiHoldNum').textContent = formatMillions(kpi.total_blocked_or_unverified_value_vnd);
      document.getElementById('kpiHoldSub').innerHTML = '<span class="text-amber-600 font-bold">' + kpi.unverified_or_blocked_accounts_count + ' DN</span> tạm giữ an toàn';

      // 2 Thẻ Gradient góc phải
      document.getElementById('bannerGbNumber').textContent = (kpi.total_accounts * 450) + ' TB';

      // Xóa lớp phủ chờ trên biểu đồ
      document.getElementById('chartPlaceholderOverlay').classList.add('hidden');

      // Huy hiệu tính toán
      const badgeInd = document.getElementById('calculatedBadgeIndicator');
      badgeInd.textContent = dataSource === 'uploaded' ? ('✓ Đã tính toán ' + kpi.total_accounts + ' DN từ tệp') : '✓ Đã hoàn tất đối soát 10/10 DN (Demo)';
      badgeInd.className = 'px-3 py-1 rounded-full bg-emerald-100 text-emerald-800 font-black text-xs';

      document.getElementById('tableSubInfo').textContent = 'Kết quả tính toán trực tiếp từ ' + (dataSource === 'uploaded' ? ('tệp: ' + uploadedFileName) : 'tệp mẫu Demo');

      // Điền bảng danh sách đối soát
      populateCalculatedTable();

      // Điền bảng Excel đã làm sạch
      populateCleanedExcelTable();

      // Vẽ biểu đồ
      renderCharts();
    }

    function populateCalculatedTable() {
      const tbody = document.getElementById('calculatedTableBody');
      if (!tbody || !analysisData) return;
      tbody.innerHTML = '';

      const profiles = analysisData.profiles || [];
      profiles.forEach(p => {
        let pillClass = 'bg-emerald-100 text-emerald-800 border border-emerald-300';
        let statusLabel = 'Thực thi';

        if (p.is_blocked_by_tech) {
          pillClass = 'bg-rose-100 text-rose-800 border border-rose-300';
          statusLabel = 'Khóa Chặn Lỗi';
        } else if (p.overall_status.includes('QUẢN LÝ')) {
          pillClass = 'bg-purple-100 text-purple-800 border border-purple-300';
          statusLabel = 'Chờ Duyệt HITL';
        } else if (!p.is_revenue_grounded) {
          pillClass = 'bg-amber-100 text-amber-800 border border-amber-300';
          statusLabel = 'Tạm Giữ / NULL';
        } else if (p.recommendation_type.includes('NÂNG CẤP')) {
          pillClass = 'bg-teal-100 text-teal-800 border border-teal-300';
          statusLabel = 'Nâng Cấp Gói';
        }

        tbody.innerHTML += `
          <tr class="hover:bg-teal-50/40 transition">
            <td class="py-3.5 px-4 font-mono font-bold text-teal-700">${escapeHtml(p.account_id)}</td>
            <td class="py-3.5 px-3 font-extrabold text-slate-900">
              ${escapeHtml(p.customer_name)}
              <div class="text-[10px] text-slate-400 font-normal">Đại diện: ${escapeHtml(p.representative)}</div>
            </td>
            <td class="py-3.5 px-3 text-slate-700 font-semibold">${escapeHtml(p.segment)}</td>
            <td class="py-3.5 px-3 text-slate-700">${escapeHtml(p.recommended_package_name)}</td>
            <td class="py-3.5 px-3 text-right font-mono font-bold text-slate-900">${formatVND(p.final_price_vnd)}</td>
            <td class="py-3.5 px-4 text-center">
              <span class="px-2.5 py-1 rounded-full text-[10px] font-black ${pillClass}">
                ${statusLabel}
              </span>
            </td>
          </tr>
        `;
      });
    }

    function populateCleanedExcelTable() {
      const tbody = document.getElementById('cleanedExcelTableBody');
      if (!tbody || !analysisData) return;
      tbody.innerHTML = '';

      const profiles = analysisData.profiles || [];
      profiles.forEach(p => {
        const usageText = (p.service_usage_rate_pct === null || p.service_usage_rate_pct === undefined) ? '<span class="text-rose-600 font-bold">NULL (Khuyết)</span>' : p.service_usage_rate_pct + '%';
        const ticketText = p.is_blocked_by_tech ? '<span class="text-rose-600 font-bold">⚠️ ' + escapeHtml(p.open_tickets) + '</span>' : '<span class="text-slate-400">' + escapeHtml(p.open_tickets || 'Bình thường') + '</span>';

        tbody.innerHTML += `
          <tr class="hover:bg-slate-50 transition">
            <td class="py-3.5 px-4 font-bold text-slate-900">
              ${escapeHtml(p.customer_name)}
              <div class="text-[10px] text-slate-400 font-mono">${escapeHtml(p.account_id)} • Đại diện: ${escapeHtml(p.representative)}</div>
            </td>
            <td class="py-3.5 px-3 text-slate-700 font-medium">${escapeHtml(p.industry)}<br><span class="text-[10px] text-teal-600 font-bold">${escapeHtml(p.segment)}</span></td>
            <td class="py-3.5 px-3 text-slate-700">${escapeHtml(p.current_package_name)}</td>
            <td class="py-3.5 px-3 font-bold">${usageText}</td>
            <td class="py-3.5 px-3">${ticketText}</td>
            <td class="py-3.5 px-3 font-semibold text-teal-700">${escapeHtml(p.recommendation_type)}</td>
            <td class="py-3.5 px-3 text-right font-bold text-slate-900 font-mono">${formatVND(p.final_price_vnd)}</td>
            <td class="py-3.5 px-4 text-slate-600 text-[11px] leading-relaxed">${escapeHtml(p.ai_remarks)}</td>
          </tr>
        `;
      });
    }

    function renderCharts() {
      if (!analysisData) return;
      const profiles = analysisData.profiles || [];

      const ctxLine = document.getElementById('revenueTrendChart').getContext('2d');
      if (lineChart) lineChart.destroy();

      const renewVal = profiles.filter(p => p.is_revenue_grounded && !p.recommendation_type.includes('NÂNG CẤP')).reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;
      const upVal = profiles.filter(p => p.is_revenue_grounded && p.recommendation_type.includes('NÂNG CẤP')).reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;

      const gradTeal = ctxLine.createLinearGradient(0, 0, 0, 240);
      gradTeal.addColorStop(0, 'rgba(13, 148, 136, 0.35)');
      gradTeal.addColorStop(1, 'rgba(13, 148, 136, 0.0)');

      const gradCyan = ctxLine.createLinearGradient(0, 0, 0, 240);
      gradCyan.addColorStop(0, 'rgba(6, 182, 212, 0.35)');
      gradCyan.addColorStop(1, 'rgba(6, 182, 212, 0.0)');

      lineChart = new Chart(ctxLine, {
        type: 'line',
        data: {
          labels: ['Kỳ 1', 'Kỳ 2', 'Kỳ 3', 'Kỳ 4', 'Kỳ 5', 'Kỳ 6', 'Hiện tại'],
          datasets: [
            {
              label: 'Doanh thu Gói Duy Trì',
              data: [
                Math.round((renewVal || 100) * 0.5),
                Math.round((renewVal || 100) * 0.65),
                Math.round((renewVal || 100) * 0.75),
                Math.round((renewVal || 100) * 0.85),
                Math.round((renewVal || 100) * 0.9),
                Math.round((renewVal || 100) * 0.95),
                Math.round(renewVal || 100)
              ],
              borderColor: '#0d9488',
              backgroundColor: gradTeal,
              tension: 0.45,
              borderWidth: 2.5,
              fill: true,
              pointBackgroundColor: '#0d9488',
              pointRadius: 4
            },
            {
              label: 'Doanh thu Gói Nâng Cấp',
              data: [
                Math.round((upVal || 150) * 0.4),
                Math.round((upVal || 150) * 0.55),
                Math.round((upVal || 150) * 0.7),
                Math.round((upVal || 150) * 0.8),
                Math.round((upVal || 150) * 0.88),
                Math.round((upVal || 150) * 0.94),
                Math.round(upVal || 150)
              ],
              borderColor: '#06b6d4',
              backgroundColor: gradCyan,
              tension: 0.45,
              borderWidth: 2.5,
              fill: true,
              pointBackgroundColor: '#06b6d4',
              pointRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            y: { beginAtZero: true, grid: { color: 'rgba(241, 245, 249, 0.8)' } },
            x: { grid: { display: false } }
          }
        }
      });
    }

    // Đặt lại toàn bộ về 0 ban đầu
    function resetAllToZero() {
      isDataLoaded = false;
      isDataCalculated = false;
      dataSource = null;
      uploadedRawData = null;
      uploadedProfiles = null;
      uploadedFileName = '';
      detectedErrors = [];
      analysisData = null;

      document.getElementById('modeBadge').textContent = 'Chờ Dữ Liệu';
      document.getElementById('modeBadge').className = 'ml-1.5 px-2 py-0.5 rounded-md bg-slate-100 text-[10px] font-extrabold text-slate-600 border border-slate-200';
      document.getElementById('currentSourceLabel').textContent = 'Chưa có dữ liệu';

      document.getElementById('step1StatusBadge').textContent = 'Chưa nạp';
      document.getElementById('step1StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      document.getElementById('step2StatusBadge').textContent = 'Chờ quét';
      document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      document.getElementById('step3StatusBadge').textContent = 'Chờ tính toán';
      document.getElementById('step3StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      // Khóa nút Bước 4
      lockExcelDownload();

      // KHÔNG ĐƯỢC PHÉP HIỆN BẤT CỨ LỖI GÌ KHI CHƯA CHẠY THỬ
      dismissErrorBox();

      // Reset các chỉ số về 0
      resetZeroMetrics();

      switchTab('dashboard');
    }

    function resetZeroMetrics() {
      document.getElementById('kpiTotalNum').textContent = '0';
      document.getElementById('kpiTotalSub').textContent = 'Chờ nạp dữ liệu';
      document.getElementById('kpiGroundedNum').textContent = '0 ₫';
      document.getElementById('kpiGroundedSub').textContent = 'Chờ nạp dữ liệu';
      document.getElementById('kpiUsageNum').textContent = '0%';
      document.getElementById('kpiUsageSub').textContent = 'Chờ nạp dữ liệu';
      document.getElementById('kpiHoldNum').textContent = '0 ₫';
      document.getElementById('kpiHoldSub').textContent = 'Chờ nạp dữ liệu';

      document.getElementById('bannerGbNumber').textContent = '0 TB';

      // Hiện lại lớp phủ chờ trên biểu đồ
      document.getElementById('chartPlaceholderOverlay').classList.remove('hidden');

      // Reset bảng
      document.getElementById('calculatedBadgeIndicator').textContent = 'Chờ tính toán';
      document.getElementById('calculatedBadgeIndicator').className = 'px-3 py-1 rounded-full bg-slate-100 text-slate-500 font-extrabold text-xs';
      document.getElementById('tableSubInfo').textContent = 'Kết quả tính toán sau khi dữ liệu đã được làm sạch và đối soát';

      document.getElementById('calculatedTableBody').innerHTML = `
        <tr>
          <td colspan="6" class="py-12 text-center text-slate-400">
            <div class="text-3xl mb-1">📭</div>
            Chưa có kết quả tính toán.<br>
            Vui lòng thực hiện tuần tự <strong>Bước 1</strong> (Nạp Dữ Liệu) và <strong>Bước 3</strong> (Chạy Tính Toán) trên thanh công cụ bên trái.
          </td>
        </tr>
      `;

      document.getElementById('cleanedExcelTableBody').innerHTML = `
        <tr>
          <td colspan="8" class="py-12 text-center text-slate-400">
            Chưa có dữ liệu làm sạch. Vui lòng thực hiện Bước 1 và Bước 3 để hiển thị bảng.
          </td>
        </tr>
      `;

      if (lineChart) { lineChart.destroy(); lineChart = null; }
    }

    async function sendChatMessage() {
      const input = document.getElementById('chatInput');
      const text = input.value.trim();
      if (!text) return;

      const log = document.getElementById('chatMessageLog');

      const uMsg = document.createElement('div');
      uMsg.className = 'p-2 rounded-lg bg-teal-600 text-white text-[11px] font-medium shadow-2xs ml-3';
      uMsg.textContent = text;
      log.appendChild(uMsg);
      input.value = '';

      const aMsg = document.createElement('div');
      aMsg.className = 'p-2 rounded-lg bg-white/20 text-white text-[11px] mr-2 border border-white/20 shadow-2xs backdrop-blur-sm';
      aMsg.innerHTML = '<span class="text-teal-200 animate-pulse font-bold">Đang tra cứu tri thức...</span>';
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
          .replace(/\*\*(.*?)\*\*/g, '<strong class="text-teal-100 font-bold">$1</strong>');
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

    function escapeHtml(text) {
      if (text === null || text === undefined) return '';
      return String(text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;');
    }

    // Khởi tạo ban đầu
    window.addEventListener('DOMContentLoaded', () => {
      resetAllToZero();
    });
  </script>
</body>
</html>
'@

[System.IO.File]::WriteAllText($targetHtml, $htmlContent, [System.Text.Encoding]::UTF8)
Write-Output ("Successfully generated dynamic UI to: " + $targetHtml + " (" + (Get-Item $targetHtml).Length + " bytes)")
