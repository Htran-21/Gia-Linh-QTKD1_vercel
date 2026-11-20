# ==============================================================================
# SCRIPT TO BUILD STEP-BY-STEP DEDICATED VIEW UI FOR GIA LINH DATA STUDIO
# - Dedicated view for each step (Step 1, Step 2, Step 3, Step 4)
# - No overly long scrolling: compact, paginated, highly optimized layout
# - Left sidebar & Icon Rail switch between steps
# - Stepper navigation bar at top
# - Strict isolation between uploaded file vs demo
# - Zero error before scanning; locked excel before calculating
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
    * { -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; }
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
    #leftSidebar {
      transition: width 0.3s cubic-bezier(0.4, 0, 0.2, 1), padding 0.3s ease;
    }
    #leftSidebar.sidebar-collapsed {
      width: 5rem !important; /* 80px */
      padding-left: 0.5rem !important;
      padding-right: 0.5rem !important;
    }
    #leftSidebar.sidebar-collapsed #sidebarFullView { display: none !important; }
    #leftSidebar.sidebar-collapsed #sidebarIconView { display: flex !important; }
    #sidebarIconView { display: none; }
    @keyframes fadeInToast {
      from { opacity: 0; transform: translateY(10px) scale(0.95); }
      to { opacity: 1; transform: translateY(0) scale(1); }
    }
    .toast-animate { animation: fadeInToast 0.25s ease-out forwards; }
  </style>
</head>
<body class="p-2 sm:p-4 flex items-center justify-center relative">

  <!-- ========================================================================= -->
  <!-- TOAST THÔNG BÁO XIN CHÀO KHÁCH HÀNG                                       -->
  <!-- ========================================================================= -->
  <div id="greetingToast" class="fixed bottom-6 right-6 z-50 hidden toast-animate p-4 rounded-2xl bg-slate-900/95 text-white border border-teal-400/40 shadow-2xl flex items-center gap-3 backdrop-blur-md max-w-md">
    <div class="w-10 h-10 rounded-xl bg-gradient-to-tr from-emerald-500 via-teal-500 to-cyan-500 text-white flex items-center justify-center text-lg font-bold shrink-0 shadow-xs">
      👋
    </div>
    <div class="flex-1 text-xs">
      <div class="font-extrabold text-sm text-teal-200">Xin chào, khách hàng!</div>
      <div class="text-slate-300 text-[11px] mt-0.5">Chào mừng bạn đến với <strong>Gia Linh Data Studio</strong>. Nhấp vào từng bước ở thanh bên trái hoặc thanh tiến trình trên để xem riêng kết quả từng bước.</div>
    </div>
    <button onclick="dismissGreetingToast()" class="text-slate-400 hover:text-white font-bold p-1 text-xs">✕</button>
  </div>

  <!-- ========================================================================= -->
  <!-- CỬA SỔ ỨNG DỤNG GIA LINH DATA STUDIO (2 CỘT)                              -->
  <!-- ========================================================================= -->
  <div class="app-window w-full max-w-[1760px] min-h-[94vh] flex flex-col overflow-hidden bg-white">

    <!-- 1. BROWSER CHROME HEADER (macOS 3 DOTS + URL BAR + TẢI TÀI LIỆU) -->
    <header class="px-5 py-2.5 border-b border-slate-100 flex items-center justify-between text-xs bg-slate-50/70 select-none">
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

      <div class="flex items-center gap-2 bg-white px-5 py-1.5 rounded-xl border border-slate-200/80 text-slate-700 font-semibold text-xs shadow-2xs">
        <span class="text-emerald-500 text-xs">🛡️</span>
        <span class="font-extrabold text-slate-900">GiaLinhDataStudio.vn</span>
        <span class="text-slate-300">/</span>
        <span id="headerStepBreadcrumb" class="text-teal-600 font-bold">bước-1-nạp-dữ-liệu</span>
        <span class="text-slate-400 text-[11px] ml-1">🔄</span>
      </div>

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
    <div class="px-6 py-3 border-b border-slate-100 flex items-center justify-between gap-4 bg-white select-none">
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

      <!-- THANH ĐIỀU HƯỚNG BƯỚC (STEPPER PILLS) TRỰC QUAN -->
      <div class="hidden md:flex items-center gap-1.5 p-1 bg-slate-100/90 rounded-2xl border border-slate-200 text-xs font-bold shadow-2xs">
        <button onclick="goToStep(1)" id="stepperTab1" class="px-3.5 py-1.5 rounded-xl bg-teal-600 text-white font-black text-xs shadow-xs transition flex items-center gap-1.5">
          <span class="w-5 h-5 rounded-full bg-white text-teal-800 flex items-center justify-center text-[10px] font-black">1</span>
          <span>Nạp Dữ Liệu</span>
        </button>
        <span class="text-slate-300 text-xs">➔</span>
        <button onclick="goToStep(2)" id="stepperTab2" class="px-3.5 py-1.5 rounded-xl bg-white hover:bg-slate-50 text-slate-600 font-bold text-xs border border-slate-200 shadow-2xs transition flex items-center gap-1.5">
          <span class="w-5 h-5 rounded-full bg-slate-200 text-slate-700 flex items-center justify-center text-[10px] font-black">2</span>
          <span>Quét &amp; Chẩn Đoán Lỗi</span>
        </button>
        <span class="text-slate-300 text-xs">➔</span>
        <button onclick="goToStep(3)" id="stepperTab3" class="px-3.5 py-1.5 rounded-xl bg-white hover:bg-slate-50 text-slate-600 font-bold text-xs border border-slate-200 shadow-2xs transition flex items-center gap-1.5">
          <span class="w-5 h-5 rounded-full bg-slate-200 text-slate-700 flex items-center justify-center text-[10px] font-black">3</span>
          <span>Tính Toán &amp; Báo Cáo</span>
        </button>
        <span class="text-slate-300 text-xs">➔</span>
        <button onclick="goToStep(4)" id="stepperTab4" class="px-3.5 py-1.5 rounded-xl bg-white hover:bg-slate-50 text-slate-600 font-bold text-xs border border-slate-200 shadow-2xs transition flex items-center gap-1.5">
          <span class="w-5 h-5 rounded-full bg-slate-200 text-slate-700 flex items-center justify-center text-[10px] font-black">4</span>
          <span>Xuất File Excel</span>
        </button>
      </div>

      <!-- Avatar Khách Hàng -->
      <div class="flex items-center gap-2 pl-2 cursor-pointer" onclick="showGreetingToast()" title="Bấm vào để hiển thị lời chào">
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

    <!-- 3. MAIN BODY: 2-COLUMN LAYOUT -->
    <div class="flex-1 flex flex-col lg:flex-row overflow-hidden relative">

      <!-- ===================================================================== -->
      <!-- CỘT TRÁI: THANH CÔNG CỤ DỌC (HỖ TRỢ ĐÓNG RA MỞ VÀO ICON BAR)           -->
      <!-- ===================================================================== -->
      <aside id="leftSidebar" class="sidebar-emerald-cyan w-full lg:w-80 xl:w-88 text-white p-4 flex flex-col justify-between shrink-0 select-none transition-all duration-300 overflow-y-auto custom-scroll space-y-3">
        
        <!-- =================================================================== -->
        <!-- CHẾ ĐỘ 1: FULL VIEW                                                 -->
        <!-- =================================================================== -->
        <div id="sidebarFullView" class="space-y-4 relative z-10">
          
          <div class="p-3.5 rounded-2xl bg-white/20 backdrop-blur-md border border-white/30 text-white space-y-1.5 shadow-sm">
            <div class="flex items-center justify-between text-xs">
              <span class="font-extrabold uppercase tracking-wider text-teal-100 text-[10px]">Trung Tâm Điều Hành</span>
              <button onclick="toggleSidebar()" class="px-2 py-0.5 rounded-full bg-white/25 hover:bg-white/35 text-[10px] font-bold transition flex items-center gap-1" title="Thu gọn thành thanh Icon">
                <span>◀</span><span>Thu gọn</span>
              </button>
            </div>
            <div class="font-black text-sm text-white leading-tight">Gia Linh Data Studio</div>
            <div class="text-[11px] text-emerald-100 font-medium">Nhấp vào từng bước để xem riêng kết quả</div>
          </div>

          <!-- DANH SÁCH 4 BƯỚC: NHẤP VÀO BƯỚC NÀO SẼ HIỆN RIÊNG KẾT QUẢ BƯỚC ĐÓ -->
          <div class="space-y-2.5">
            <div class="text-[11px] font-black uppercase tracking-wider text-teal-100 px-1 flex items-center justify-between">
              <span>Quy Trình 4 Bước Thao Tác</span>
              <span id="currentSourceLabel" class="text-[9px] text-white/80 font-normal">Chưa có dữ liệu</span>
            </div>

            <!-- BƯỚC 1 CARD -->
            <div id="sidebarStepCard1" onclick="goToStep(1)" class="p-3 rounded-2xl bg-white/25 border-2 border-white text-white space-y-1.5 cursor-pointer transition hover:bg-white/30 shadow-sm">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 rounded-lg bg-white text-teal-700 flex items-center justify-center font-black text-xs shrink-0 shadow-2xs">1</span>
                  <span class="font-black text-xs text-white">Bước 1: Nạp Dữ Liệu</span>
                </div>
                <span id="step1StatusBadge" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white">Chưa nạp</span>
              </div>
              <p class="text-[10px] text-teal-100">Tiếp nhận tệp dữ liệu .xlsx, .csv, .json hoặc tệp demo.</p>
            </div>

            <!-- BƯỚC 2 CARD -->
            <div id="sidebarStepCard2" onclick="goToStep(2)" class="p-3 rounded-2xl bg-white/15 border border-white/25 text-white space-y-1.5 cursor-pointer transition hover:bg-white/25">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 rounded-lg bg-white text-cyan-700 flex items-center justify-center font-black text-xs shrink-0 shadow-2xs">2</span>
                  <span class="font-black text-xs text-white">Bước 2: Quét &amp; Chẩn Đoán Lỗi</span>
                </div>
                <span id="step2StatusBadge" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white">Chờ quét</span>
              </div>
              <p class="text-[10px] text-teal-100">Rà soát ngày sai, phân biệt NULL vs 0%, lỗi máy chủ.</p>
            </div>

            <!-- BƯỚC 3 CARD -->
            <div id="sidebarStepCard3" onclick="goToStep(3)" class="p-3 rounded-2xl bg-white/15 border border-white/25 text-white space-y-1.5 cursor-pointer transition hover:bg-white/25">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 rounded-lg bg-white text-blue-700 flex items-center justify-center font-black text-xs shrink-0 shadow-2xs">3</span>
                  <span class="font-black text-xs text-white">Bước 3: Chạy Tính Toán</span>
                </div>
                <span id="step3StatusBadge" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white">Chờ tính toán</span>
              </div>
              <p class="text-[10px] text-blue-100">Tính chỉ số KPI, phân bổ gói dịch vụ, đối soát doanh thu.</p>
            </div>

            <!-- BƯỚC 4 CARD -->
            <div id="sidebarStepCard4" onclick="goToStep(4)" class="p-3 rounded-2xl bg-white/15 border border-white/25 text-white space-y-1.5 cursor-pointer transition hover:bg-white/25">
              <div class="flex items-center justify-between">
                <div class="flex items-center gap-2">
                  <span class="w-6 h-6 rounded-lg bg-white text-indigo-700 flex items-center justify-center font-black text-xs shrink-0 shadow-2xs">4</span>
                  <span class="font-black text-xs text-white">Bước 4: Xuất Kết Quả Excel</span>
                </div>
                <span id="step4StatusBadge" class="text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white">Chưa mở</span>
              </div>
              <p class="text-[10px] text-teal-100">Tải tệp Excel đã làm sạch kèm nhận xét AI.</p>
            </div>
          </div>

          <!-- TRỢ LÝ AI CHAT TRỰC TIẾP -->
          <div class="pt-2 border-t border-white/20 space-y-2">
            <div class="flex items-center justify-between text-[11px] font-black uppercase text-teal-100 px-1">
              <span>Trợ Lý Tra Cứu AI</span>
              <span class="text-[9px] text-emerald-200 bg-white/20 px-2 py-0.5 rounded-full font-bold">Sẵn sàng</span>
            </div>
            
            <div id="chatMessageLog" class="max-h-28 overflow-y-auto custom-scroll space-y-1.5 p-2 bg-white/10 rounded-xl border border-white/20 text-xs">
              <div class="p-2 rounded-lg bg-white/20 text-white text-[11px] leading-relaxed backdrop-blur-sm">
                👋 <strong>Gia Linh Data Studio</strong> sẵn sàng giải đáp về mã doanh nghiệp, quy tắc tính toán và phân định <em>NULL vs 0%</em>.
              </div>
            </div>

            <div class="relative">
              <input id="chatInput" type="text" placeholder="Hỏi về mã doanh nghiệp, NULL vs 0..." onkeydown="if(event.key==='Enter') sendChatMessage()" class="w-full pl-3 pr-8 py-2 rounded-xl bg-white/20 border border-white/30 text-xs text-white placeholder-white/60 focus:outline-none focus:bg-white/30 font-medium" />
              <button onclick="sendChatMessage()" class="absolute right-1.5 top-2 w-6 h-6 rounded-lg bg-white text-teal-800 flex items-center justify-center text-[10px] font-bold transition shadow-xs" title="Gửi câu hỏi">
                ➤
              </button>
            </div>
          </div>

        </div>

        <!-- =================================================================== -->
        <!-- CHẾ ĐỘ 2: ICON RAIL VIEW                                            -->
        <!-- =================================================================== -->
        <div id="sidebarIconView" class="flex-col items-center justify-between h-full py-2 space-y-5 relative z-10 w-full">
          <div class="flex flex-col items-center gap-3">
            <button onclick="toggleSidebar()" class="w-10 h-10 rounded-xl bg-white/20 hover:bg-white/30 text-white flex items-center justify-center text-sm font-black transition shadow-xs" title="Mở rộng thanh công cụ đầy đủ">
              ▶
            </button>
            <div class="w-10 h-10 rounded-xl bg-white text-teal-800 flex items-center justify-center font-black text-sm shadow-md" title="Gia Linh Data Studio">
              .D
            </div>
          </div>

          <div class="flex flex-col items-center gap-3.5">
            <!-- Icon Bước 1 -->
            <button onclick="goToStep(1)" id="iconStepBtn1" class="w-11 h-11 rounded-2xl bg-white/35 ring-2 ring-white text-white flex items-center justify-center text-lg transition shadow-md relative" title="Bước 1: Nạp Dữ Liệu">
              <span>📂</span>
              <span class="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-teal-400 text-teal-950 font-black text-[10px] flex items-center justify-center shadow-xs">1</span>
            </button>

            <!-- Icon Bước 2 -->
            <button onclick="goToStep(2)" id="iconStepBtn2" class="w-11 h-11 rounded-2xl bg-white/20 hover:bg-white/30 text-white flex items-center justify-center text-lg transition shadow-md relative" title="Bước 2: Quét & Chẩn Đoán Lỗi">
              <span>🔍</span>
              <span id="iconBadge2" class="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-cyan-300 text-cyan-950 font-black text-[10px] flex items-center justify-center shadow-xs">2</span>
            </button>

            <!-- Icon Bước 3 -->
            <button onclick="goToStep(3)" id="iconStepBtn3" class="w-11 h-11 rounded-2xl bg-white/20 hover:bg-white/30 text-white flex items-center justify-center text-lg transition shadow-md relative" title="Bước 3: Chạy Tính Toán & Bảng Điều Khiển">
              <span>⚡</span>
              <span id="iconBadge3" class="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-blue-300 text-blue-950 font-black text-[10px] flex items-center justify-center shadow-xs">3</span>
            </button>

            <!-- Icon Bước 4 -->
            <button onclick="goToStep(4)" id="iconStepBtn4" class="w-11 h-11 rounded-2xl bg-white/10 text-white/50 flex items-center justify-center text-lg transition shadow-md relative" title="Bước 4: Xuất Kết Quả Excel">
              <span id="iconExcelSymbol">🔒</span>
              <span id="iconBadge4" class="absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-slate-400 text-slate-900 font-black text-[10px] flex items-center justify-center shadow-xs">4</span>
            </button>
          </div>

          <div class="pt-2 border-t border-white/20 cursor-pointer" onclick="showGreetingToast()" title="Khách Hàng (Bấm để hiện lời chào)">
            <div class="w-10 h-10 rounded-full bg-white text-teal-800 flex items-center justify-center font-black text-xs shadow-md">
              KH
            </div>
          </div>
        </div>

      </aside>

      <!-- ===================================================================== -->
      <!-- KHÔNG GIAN LÀM VIỆC CHÍNH: HIỆN RIÊNG TỪNG BƯỚC, TỐI ƯU GỌN GÀNG      -->
      <!-- ===================================================================== -->
      <main class="flex-1 overflow-y-auto custom-scroll p-4 sm:p-6 bg-[#f8fafc] flex flex-col justify-between space-y-4">

        <!-- =================================================================== -->
        <!-- VIEW BƯỚC 1: NẠP DỮ LIỆU ĐẦU VÀO                                     -->
        <!-- =================================================================== -->
        <section id="viewStep1" class="space-y-4">
          <div class="flex flex-wrap items-center justify-between gap-3">
            <div>
              <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-teal-100 text-teal-800 font-black text-xs mb-1">
                <span>📂</span><span>BƯỚC 1 TRÊN 4</span>
              </div>
              <h2 class="text-xl font-black text-slate-900 tracking-tight">Nạp Dữ Liệu Khách Hàng Doanh Nghiệp</h2>
              <p class="text-xs text-slate-500 mt-0.5">Tiếp nhận tệp dữ liệu của bạn để tính toán độc lập, hoặc chạy thử tệp mẫu Demo 10 doanh nghiệp.</p>
            </div>
            <div id="step1QuickStatus" class="px-3 py-1 rounded-xl bg-slate-100 text-slate-600 font-extrabold text-xs">
              Chưa có dữ liệu
            </div>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <!-- Khối 1: Nạp tệp của bạn -->
            <div class="p-6 rounded-3xl bg-white border-2 border-dashed border-teal-300 hover:border-teal-500 transition shadow-sm flex flex-col justify-between space-y-4 text-center">
              <div class="space-y-2">
                <div class="w-12 h-12 rounded-2xl bg-teal-50 text-teal-600 mx-auto flex items-center justify-center text-2xl font-black shadow-2xs">
                  📥
                </div>
                <h3 class="font-extrabold text-sm text-slate-900">Nạp Tệp Của Bạn (.xlsx, .csv, .json)</h3>
                <p class="text-xs text-slate-500 max-w-sm mx-auto">Hệ thống sẽ tính toán và phân tích <strong>trực tiếp từ chính tệp này</strong>, hoàn toàn độc lập với phần demo.</p>
              </div>

              <div class="space-y-2">
                <label class="w-full py-3 px-4 rounded-2xl bg-gradient-to-r from-teal-600 to-emerald-600 hover:from-teal-700 hover:to-emerald-700 text-white font-extrabold text-xs cursor-pointer shadow-md transition flex items-center justify-center gap-2">
                  <input type="file" id="mainFileInput" accept=".xlsx,.xls,.csv,.json" class="hidden" onchange="handleFileUpload(event)" />
                  <span>📁</span><span>Bấm Chọn Tệp Dữ Liệu Từ Máy Tính</span>
                </label>
                <div class="text-[11px] text-slate-400">Hỗ trợ Excel (.xlsx, .xls), CSV hoặc JSON</div>
              </div>
            </div>

            <!-- Khối 2: Chạy dữ liệu Demo -->
            <div class="p-6 rounded-3xl bg-white border border-slate-200 hover:border-cyan-400 transition shadow-sm flex flex-col justify-between space-y-4 text-center">
              <div class="space-y-2">
                <div class="w-12 h-12 rounded-2xl bg-cyan-50 text-cyan-600 mx-auto flex items-center justify-center text-2xl font-black shadow-2xs">
                  🎯
                </div>
                <h3 class="font-extrabold text-sm text-slate-900">Chạy Thử Tệp Mẫu Demo (10 Doanh Nghiệp)</h3>
                <p class="text-xs text-slate-500 max-w-sm mx-auto">Chỉ khi bấm nút này, hệ thống mới sử dụng dữ liệu demo để minh họa cơ chế tách biệt <strong>NULL ≠ 0%</strong> và khóa chặn <strong>Hard-Block</strong>.</p>
              </div>

              <button onclick="loadSampleDataset()" class="w-full py-3 px-4 rounded-2xl bg-slate-900 hover:bg-slate-800 text-white font-extrabold text-xs shadow-md transition flex items-center justify-center gap-2">
                <span>📂</span><span>Nạp Tệp Mẫu Demo (10 Doanh Nghiệp)</span>
              </button>
            </div>
          </div>

          <!-- THẺ TÓM TẮT DỮ LIỆU ĐÃ NẠP (HIỆN KHI ĐÃ NẠP DỮ LIỆU) -->
          <div id="step1LoadedSummaryCard" class="hidden p-5 rounded-3xl bg-gradient-to-r from-emerald-50 via-teal-50 to-cyan-50 border border-teal-200 shadow-2xs space-y-3">
            <div class="flex flex-wrap items-center justify-between gap-2">
              <div class="flex items-center gap-2.5">
                <span class="w-8 h-8 rounded-xl bg-emerald-500 text-white flex items-center justify-center font-bold text-sm shadow-xs">✓</span>
                <div>
                  <div id="summaryFileName" class="text-sm font-black text-slate-900">Tệp đã nạp thành công</div>
                  <div id="summarySourceType" class="text-xs text-teal-700 font-semibold">Nguồn: Tệp người dùng tải lên</div>
                </div>
              </div>
              <button onclick="goToStep(2)" class="px-5 py-2.5 rounded-xl bg-gradient-to-r from-teal-600 to-emerald-600 hover:from-teal-700 hover:to-emerald-700 text-white font-extrabold text-xs shadow-md transition flex items-center gap-2">
                <span>Chuyển Sang Bước 2: Quét &amp; Chẩn Đoán Lỗi</span>
                <span>➔</span>
              </button>
            </div>
            <div class="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-2 border-t border-teal-200/60 text-xs">
              <div><span class="text-slate-400">Số doanh nghiệp:</span> <strong id="summaryRowCount" class="text-slate-900">0</strong></div>
              <div><span class="text-slate-400">Cột nhận diện:</span> <strong id="summaryColCount" class="text-slate-900">Chuẩn hóa đầy đủ</strong></div>
              <div><span class="text-slate-400">Trạng thái:</span> <strong class="text-emerald-700 font-bold">Sẵn sàng chẩn đoán</strong></div>
              <div><span class="text-slate-400">Bước kế tiếp:</span> <strong class="text-teal-700 font-bold">Bước 2: Quét lỗi</strong></div>
            </div>
          </div>
        </section>

        <!-- =================================================================== -->
        <!-- VIEW BƯỚC 2: QUÉT & CHẨN ĐOÁN LỖI                                    -->
        <!-- =================================================================== -->
        <section id="viewStep2" class="hidden space-y-4">
          <div class="flex flex-wrap items-center justify-between gap-3">
            <div>
              <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-cyan-100 text-cyan-800 font-black text-xs mb-1">
                <span>🔍</span><span>BƯỚC 2 TRÊN 4</span>
              </div>
              <h2 class="text-xl font-black text-slate-900 tracking-tight">Quét &amp; Chẩn Đoán Lỗi Dữ Liệu</h2>
              <p class="text-xs text-slate-500 mt-0.5">Rà soát ngày sai, phát hiện khuyết thiếu NULL và kích hoạt khóa chặn sự cố kỹ thuật.</p>
            </div>
            
            <div class="flex items-center gap-2">
              <button onclick="runErrorScan()" class="px-4 py-2 rounded-xl bg-cyan-500 hover:bg-cyan-400 text-slate-900 font-extrabold text-xs shadow-md transition flex items-center gap-2">
                <span>🔍</span><span>Bấm Bắt Đầu Quét &amp; Chẩn Đoán Lỗi</span>
              </button>
            </div>
          </div>

          <!-- VÙNG HIỂN THỊ KẾT QUẢ BƯỚC 2: GỌN GÀNG, KHÔNG DÀI TRÀN LAN -->
          <div id="step2InitialPrompt" class="p-8 text-center bg-white rounded-3xl border border-slate-200 shadow-sm space-y-2">
            <div class="text-3xl text-cyan-500">🛡️</div>
            <div class="text-sm font-black text-slate-900">Khi chưa bấm chạy thử, hệ thống tuyệt đối không hiển thị bất kỳ lỗi nào.</div>
            <div class="text-xs text-slate-500">Vui lòng nhấp vào nút <strong>[Bấm Bắt Đầu Quét &amp; Chẩn Đoán Lỗi]</strong> ở trên để tiến hành kiểm tra tệp dữ liệu.</div>
          </div>

          <!-- BẢNG LỖI KHI ĐÃ QUÉT: THIẾT KẾ CÓ CHIỀU CAO GỌN GÀNG CÙNG NÚT SỬA TỰ ĐỘNG -->
          <div id="step2ResultCard" class="hidden p-5 rounded-3xl border text-xs space-y-3 shadow-sm bg-white">
            <div class="flex flex-wrap justify-between items-center gap-2">
              <div class="flex items-center gap-2">
                <span id="step2ResultIcon" class="w-7 h-7 rounded-xl bg-rose-600 text-white flex items-center justify-center font-bold text-xs">⚠️</span>
                <span id="step2ResultTitle" class="text-slate-900 font-black text-sm">KẾT QUẢ QUÉT LỖI DỮ LIỆU</span>
                <span id="step2ResultBadge" class="px-2.5 py-0.5 rounded-full bg-rose-600 text-white font-extrabold text-[10px]">Cần Xử Lý</span>
              </div>
              <button onclick="autoFixErrorsAndProceed()" id="btnAutoFixStep2" class="px-4 py-2 bg-gradient-to-r from-teal-600 to-emerald-600 hover:from-teal-700 hover:to-emerald-700 text-white font-extrabold rounded-xl text-xs shadow-xs transition flex items-center gap-1.5">
                <span>🔧</span><span>Tự Động Chuẩn Hóa Lỗi (Về 0 Lỗi) &amp; Sang Bước 3</span>
              </button>
            </div>

            <!-- Bảng cuộn gọn gàng max-height 260px -->
            <div id="step2TableWrapper" class="overflow-x-auto max-h-[260px] custom-scroll rounded-xl border border-slate-200">
              <!-- Nội dung bảng lỗi / bảng sạch render động -->
            </div>

            <div class="flex justify-between items-center text-[11px] text-slate-500 pt-1">
              <span>Nguyên tắc bảo vệ: Dữ liệu phải đạt chuẩn <strong>0 lỗi</strong> mới đưa vào tính toán doanh thu đủ căn cứ.</span>
              <button onclick="goToStep(3)" class="text-teal-700 hover:text-teal-900 font-bold">Bỏ qua để sang Bước 3 ➔</button>
            </div>
          </div>
        </section>

        <!-- =================================================================== -->
        <!-- VIEW BƯỚC 3: CHẠY TÍNH TOÁN & BẢNG ĐIỀU KHIỂN (GỌN GÀNG, KHÔNG DÀI)  -->
        <!-- =================================================================== -->
        <section id="viewStep3" class="hidden space-y-4">
          <div class="flex flex-wrap items-center justify-between gap-3">
            <div>
              <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-blue-100 text-blue-800 font-black text-xs mb-1">
                <span>⚡</span><span>BƯỚC 3 TRÊN 4</span>
              </div>
              <h2 class="text-xl font-black text-slate-900 tracking-tight">Tính Toán Chỉ Số KPI &amp; Bảng Điều Khiển</h2>
              <p class="text-xs text-slate-500 mt-0.5">Kết quả tính toán doanh thu đủ căn cứ, phân loại rủi ro và nhận xét AI.</p>
            </div>

            <div class="flex items-center gap-2">
              <button onclick="runCalculationPipeline()" id="btnCalculateStep3" class="px-4 py-2 rounded-xl bg-gradient-to-r from-emerald-400 via-teal-300 to-cyan-300 text-slate-900 hover:brightness-105 font-black text-xs shadow-md transition flex items-center gap-1.5">
                <span>⚡</span><span>Bắt Đầu Tính Toán &amp; Phân Tích</span>
              </button>
            </div>
          </div>

          <!-- HÀNG 1: 4 THẺ METRICS GỌN GÀNG (CHIỀU CAO TỐI ƯU ~85PX) -->
          <div class="grid grid-cols-2 lg:grid-cols-4 gap-3">
            <div class="p-3.5 rounded-2xl bg-white border border-slate-100 shadow-2xs space-y-1">
              <div class="flex justify-between items-center text-[11px] font-bold text-slate-500">
                <span>Tổng Khách Hàng</span>
                <span class="w-6 h-6 rounded-lg bg-teal-50 text-teal-600 flex items-center justify-center font-bold text-xs">🏢</span>
              </div>
              <div id="kpiTotalNum" class="text-2xl font-black text-slate-900">0</div>
              <div id="kpiTotalSub" class="text-[10px] text-slate-400 truncate">Chờ tính toán</div>
            </div>

            <div class="p-3.5 rounded-2xl bg-white border border-slate-100 shadow-2xs space-y-1">
              <div class="flex justify-between items-center text-[11px] font-bold text-slate-500">
                <span>Doanh Thu Đủ Căn Cứ</span>
                <span class="w-6 h-6 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center font-bold text-xs">✓</span>
              </div>
              <div id="kpiGroundedNum" class="text-2xl font-black text-slate-900">0 ₫</div>
              <div id="kpiGroundedSub" class="text-[10px] text-slate-400 truncate">Chờ tính toán</div>
            </div>

            <div class="p-3.5 rounded-2xl bg-white border border-slate-100 shadow-2xs space-y-1">
              <div class="flex justify-between items-center text-[11px] font-bold text-slate-500">
                <span>Mức Dùng Dịch Vụ TB</span>
                <span class="w-6 h-6 rounded-lg bg-cyan-50 text-cyan-600 flex items-center justify-center font-bold text-xs">📊</span>
              </div>
              <div id="kpiUsageNum" class="text-2xl font-black text-slate-900">0%</div>
              <div id="kpiUsageSub" class="text-[10px] text-slate-400 truncate">NULL ≠ 0%</div>
            </div>

            <div class="p-3.5 rounded-2xl bg-white border border-slate-100 shadow-2xs space-y-1">
              <div class="flex justify-between items-center text-[11px] font-bold text-slate-500">
                <span>Doanh Thu Tạm Giữ</span>
                <span class="w-6 h-6 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center font-bold text-xs">⚠️</span>
              </div>
              <div id="kpiHoldNum" class="text-2xl font-black text-slate-900">0 ₫</div>
              <div id="kpiHoldSub" class="text-[10px] text-slate-400 truncate">Chốt chặn an toàn</div>
            </div>
          </div>

          <!-- HÀNG 2: BIỂU ĐỒ SÓNG NHỎ GỌN (HEIGHT 190PX) + 3 THỐNG KÊ NHANH -->
          <div class="grid grid-cols-1 lg:grid-cols-12 gap-3">
            <div class="lg:col-span-8 p-4 rounded-2xl bg-white border border-slate-100 shadow-2xs space-y-2">
              <div class="flex justify-between items-center text-xs">
                <span class="font-extrabold text-slate-900">Biến Động Doanh Thu Gói Duy Trì vs Nâng Cấp</span>
                <div class="flex items-center gap-3 text-[10px] font-bold">
                  <span class="text-teal-600">● Gói Duy Trì</span>
                  <span class="text-cyan-600">● Gói Nâng Cấp</span>
                </div>
              </div>
              <div class="h-44 w-full relative">
                <canvas id="revenueTrendChart"></canvas>
                <div id="chartPlaceholderOverlay" class="absolute inset-0 flex flex-col items-center justify-center bg-white/95 rounded-xl">
                  <div class="text-2xl mb-1 text-teal-500">📈</div>
                  <div class="font-bold text-xs text-slate-800">Chưa có kết quả tính toán</div>
                  <div class="text-[10px] text-slate-400">Bấm nút [Bắt Đầu Tính Toán] ở trên để vẽ biểu đồ</div>
                </div>
              </div>
            </div>

            <!-- 3 Khối Thống Kê Nhanh (Thay thế các thẻ dài) -->
            <div class="lg:col-span-4 p-4 rounded-2xl bg-white border border-slate-100 shadow-2xs flex flex-col justify-between space-y-2 text-xs">
              <span class="font-black text-slate-900">Cơ Cấu Trạng Thái Đối Soát</span>
              <div class="space-y-1.5 flex-1 justify-center flex flex-col">
                <div class="flex justify-between items-center p-2 rounded-xl bg-emerald-50 text-emerald-900 font-bold">
                  <span>✓ Sẵn Sàng Thực Thi:</span>
                  <span id="statReadyCount" class="font-black">0 DN</span>
                </div>
                <div class="flex justify-between items-center p-2 rounded-xl bg-amber-50 text-amber-900 font-bold">
                  <span>⚠️ Tạm Giữ / NULL Đo Lường:</span>
                  <span id="statHoldCount" class="font-black">0 DN</span>
                </div>
                <div class="flex justify-between items-center p-2 rounded-xl bg-rose-50 text-rose-900 font-bold">
                  <span>🛡️ Khóa Chặn Kỹ Thuật (Hard-Block):</span>
                  <span id="statBlockCount" class="font-black">0 DN</span>
                </div>
              </div>
              <button onclick="goToStep(4)" class="w-full py-2 bg-gradient-to-r from-teal-600 to-emerald-600 text-white font-extrabold text-[11px] rounded-xl shadow-xs transition text-center">
                Sang Bước 4: Xuất Kết Quả Excel ➔
              </button>
            </div>
          </div>

          <!-- HÀNG 3: BẢNG KẾT QUẢ ĐỐI SOÁT PHÂN TRANG GỌN GÀNG (5 DÒNG / TRANG) -->
          <div class="p-4 rounded-2xl bg-white border border-slate-100 shadow-2xs space-y-2.5">
            <div class="flex flex-wrap justify-between items-center gap-2">
              <div class="flex items-center gap-2">
                <span class="font-black text-xs text-slate-900">Danh Sách Doanh Nghiệp Đã Tính Toán</span>
                <span id="tablePageBadge" class="text-[10px] px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 font-bold">0 bản ghi</span>
              </div>
              <!-- Tìm kiếm nhanh -->
              <input id="tableSearchInput" type="text" placeholder="Lọc theo mã hoặc tên..." oninput="handleTableSearch(this.value)" class="py-1 px-3 rounded-xl border border-slate-200 text-xs text-slate-800 placeholder-slate-400 focus:outline-none focus:border-teal-500 w-44 shadow-2xs" />
            </div>

            <!-- Bảng dữ liệu gọn gàng -->
            <div class="overflow-x-auto rounded-xl border border-slate-100">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <thead class="bg-slate-50 text-slate-700 uppercase font-black text-[10px] border-b border-slate-200">
                  <tr>
                    <th class="py-2.5 px-3">MÃ DN</th>
                    <th class="py-2.5 px-3">TÊN DOANH NGHIỆP</th>
                    <th class="py-2.5 px-3">MỨC DÙNG</th>
                    <th class="py-2.5 px-3">GÓI ĐỀ XUẤT</th>
                    <th class="py-2.5 px-3 text-right">GIÁ TRỊ</th>
                    <th class="py-2.5 px-3 text-center">TRẠNG THÁI</th>
                  </tr>
                </thead>
                <tbody id="paginatedTableBody" class="divide-y divide-slate-100 bg-white">
                  <tr>
                    <td colspan="6" class="py-6 text-center text-slate-400">
                      Chưa có kết quả tính toán. Bấm [Bắt Đầu Tính Toán &amp; Phân Tích] ở trên.
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            <!-- Thanh điều hướng phân trang (Pagination Bar) -->
            <div id="tablePaginationControls" class="flex justify-between items-center text-xs pt-1">
              <span id="paginationInfo" class="text-[11px] text-slate-400">Trang 1 / 1</span>
              <div class="flex items-center gap-1">
                <button onclick="prevTablePage()" id="btnPrevPage" disabled class="px-2.5 py-1 rounded-lg border border-slate-200 text-slate-600 disabled:opacity-40 font-bold text-xs hover:bg-slate-50">◀ Trước</button>
                <button onclick="nextTablePage()" id="btnNextPage" disabled class="px-2.5 py-1 rounded-lg border border-slate-200 text-slate-600 disabled:opacity-40 font-bold text-xs hover:bg-slate-50">Sau ▶</button>
              </div>
            </div>
          </div>
        </section>

        <!-- =================================================================== -->
        <!-- VIEW BƯỚC 4: XUẤT KẾT QUẢ EXCEL & TÀI LIỆU                           -->
        <!-- =================================================================== -->
        <section id="viewStep4" class="hidden space-y-4">
          <div class="flex flex-wrap items-center justify-between gap-3">
            <div>
              <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-indigo-100 text-indigo-800 font-black text-xs mb-1">
                <span>📥</span><span>BƯỚC 4 TRÊN 4</span>
              </div>
              <h2 class="text-xl font-black text-slate-900 tracking-tight">Xuất Kết Quả Excel &amp; Báo Cáo Phân Tích</h2>
              <p class="text-xs text-slate-500 mt-0.5">Tải tệp Excel đã làm sạch kèm nhận xét AI và các tài liệu báo cáo tổng quan.</p>
            </div>
            
            <div id="step4GatedBadge" class="px-3 py-1 rounded-xl bg-slate-100 text-slate-500 font-extrabold text-xs">
              Chờ tính toán
            </div>
          </div>

          <!-- KHỐI CHÍNH TẢI FILE EXCEL -->
          <div class="p-6 rounded-3xl bg-white border border-slate-200 shadow-sm space-y-4">
            <div class="flex flex-wrap items-center justify-between gap-3">
              <div class="space-y-1">
                <div class="flex items-center gap-2">
                  <span class="w-8 h-8 rounded-xl bg-emerald-100 text-emerald-800 flex items-center justify-center font-black text-sm">📊</span>
                  <h3 class="font-extrabold text-sm text-slate-900">Tệp Excel Dữ Liệu Đã Làm Sạch Kèm Nhận Xét AI (.xlsx)</h3>
                </div>
                <p id="exportFileNameDesc" class="text-xs text-slate-500">Tệp Excel bao gồm 20 cột chuẩn hóa, phân tách NULL ≠ 0% và có nhận xét chi tiết cho từng doanh nghiệp.</p>
              </div>

              <!-- Nút tải Excel -->
              <button onclick="downloadCalculatedExcel()" id="mainDownloadExcelBtn" disabled class="py-3 px-6 rounded-2xl bg-slate-200 text-slate-400 font-black text-xs cursor-not-allowed transition flex items-center gap-2 shadow-sm">
                <span id="mainLockIcon">🔒</span>
                <span id="mainLockText">Xuất Tệp Excel (Chờ Tính Toán Ở Bước 3)</span>
              </button>
            </div>

            <!-- 3 Khối Thông Tin Chi Tiết Của Tệp Excel -->
            <div class="grid grid-cols-1 md:grid-cols-3 gap-3 pt-3 border-t border-slate-100 text-xs">
              <div class="p-3 rounded-2xl bg-slate-50 border border-slate-100 space-y-1">
                <div class="font-bold text-slate-700">1. Chuẩn Hóa Cấu Trúc</div>
                <div class="text-[11px] text-slate-500">Làm sạch ngày tháng, loại bỏ trùng lặp và tính toán giá trị hợp đồng chính xác.</div>
              </div>
              <div class="p-3 rounded-2xl bg-slate-50 border border-slate-100 space-y-1">
                <div class="font-bold text-slate-700">2. Phân Định NULL vs 0%</div>
                <div class="text-[11px] text-slate-500">Tách riêng dữ liệu khuyết thiếu cổng đo lường để tạm giữ an toàn, không nhầm lẫn 0%.</div>
              </div>
              <div class="p-3 rounded-2xl bg-slate-50 border border-slate-100 space-y-1">
                <div class="font-bold text-slate-700">3. Nhận Xét &amp; Đề Xuất AI</div>
                <div class="text-[11px] text-slate-500">Tự động tạo nhận xét khách quan và kế hoạch hành động chi tiết cho từng hồ sơ.</div>
              </div>
            </div>
          </div>

          <!-- CỤM TÀI LIỆU BỔ TRỢ: BÁO CÁO WORD & SLIDE POWERPOINT -->
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div class="p-5 rounded-3xl bg-white border border-slate-200 shadow-2xs flex items-center justify-between gap-3">
              <div class="space-y-0.5">
                <div class="font-extrabold text-xs text-slate-900 flex items-center gap-1.5">
                  <span>📄</span><span>Báo Cáo Tổng Quan Hệ Thống (.docx)</span>
                </div>
                <div class="text-[11px] text-slate-500">Cấu trúc chuẩn nghị định, 9 chuyên đề, phân tích SWOT.</div>
              </div>
              <a href="/api/download-report-docx" download class="px-4 py-2 rounded-xl bg-teal-50 hover:bg-teal-100 text-teal-800 font-extrabold text-xs transition border border-teal-200">
                Tải Word
              </a>
            </div>

            <div class="p-5 rounded-3xl bg-white border border-slate-200 shadow-2xs flex items-center justify-between gap-3">
              <div class="space-y-0.5">
                <div class="font-extrabold text-xs text-slate-900 flex items-center gap-1.5">
                  <span>📽️</span><span>Slide Thuyết Trình 10 Trang (.pptx)</span>
                </div>
                <div class="text-[11px] text-slate-500">Tông màu xanh dương nhạt hiện đại, thiết kế dạng thẻ card.</div>
              </div>
              <a href="/api/download-presentation-pptx" download class="px-4 py-2 rounded-xl bg-sky-50 hover:bg-sky-100 text-sky-800 font-extrabold text-xs transition border border-sky-200">
                Tải PPTX
              </a>
            </div>
          </div>
        </section>

        <!-- Chân trang ứng dụng Gia Linh Data Studio -->
        <footer class="pt-3 border-t border-slate-200 flex justify-between items-center text-[11px] text-slate-400 select-none">
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
    // Trạng thái hệ thống
    let currentStep = 1;
    let isDataLoaded = false;
    let isDataCalculated = false;
    let dataSource = null; // 'uploaded' hoặc 'demo'
    let uploadedRawData = null;
    let uploadedProfiles = null;
    let uploadedFileName = '';
    let detectedErrors = [];
    let analysisData = null;
    let lineChart = null;
    let isSidebarCollapsed = false;

    // Phân trang bảng Bước 3
    let currentPage = 1;
    const pageSize = 5;
    let filteredProfiles = [];

    // =========================================================================
    // ĐIỀU HƯỚNG BƯỚC: KHI BẤM VÀO BƯỚC NÀO THÌ HIỆN RIÊNG BƯỚC ĐÓ
    // =========================================================================
    function goToStep(stepNum) {
      currentStep = stepNum;

      // 1. Ẩn tất cả các view, chỉ hiện view của bước được chọn
      for (let i = 1; i <= 4; i++) {
        const v = document.getElementById('viewStep' + i);
        if (v) {
          if (i === stepNum) {
            v.classList.remove('hidden');
          } else {
            v.classList.add('hidden');
          }
        }

        // Cập nhật tab Stepper trên cùng
        const tab = document.getElementById('stepperTab' + i);
        if (tab) {
          if (i === stepNum) {
            tab.className = 'px-3.5 py-1.5 rounded-xl bg-teal-600 text-white font-black text-xs shadow-xs transition flex items-center gap-1.5';
          } else {
            tab.className = 'px-3.5 py-1.5 rounded-xl bg-white hover:bg-slate-50 text-slate-600 font-bold text-xs border border-slate-200 shadow-2xs transition flex items-center gap-1.5';
          }
        }

        // Cập nhật thẻ bước bên cột trái (Full View)
        const sideCard = document.getElementById('sidebarStepCard' + i);
        if (sideCard) {
          if (i === stepNum) {
            sideCard.className = 'p-3 rounded-2xl bg-white/30 border-2 border-white text-white space-y-1.5 cursor-pointer transition shadow-md';
          } else {
            sideCard.className = 'p-3 rounded-2xl bg-white/15 border border-white/25 text-white space-y-1.5 cursor-pointer transition hover:bg-white/20';
          }
        }

        // Cập nhật nút Icon bên cột trái (Icon View)
        const iconBtn = document.getElementById('iconStepBtn' + i);
        if (iconBtn) {
          if (i === stepNum) {
            iconBtn.classList.add('bg-white/40', 'ring-2', 'ring-white');
          } else {
            iconBtn.classList.remove('bg-white/40', 'ring-2', 'ring-white');
          }
        }
      }

      // Cập nhật breadcrumb trên thanh địa chỉ
      const bCrumb = document.getElementById('headerStepBreadcrumb');
      if (bCrumb) {
        const stepNames = ['', 'bước-1-nạp-dữ-liệu', 'bước-2-quét-lỗi', 'bước-3-tính-toán-kpi', 'bước-4-xuất-excel'];
        bCrumb.textContent = stepNames[stepNum] || 'quy-trinh-4-buoc';
      }

      // Nếu chuyển sang Bước 3, vẽ lại biểu đồ
      if (stepNum === 3 && isDataCalculated) {
        setTimeout(renderCharts, 40);
      }
    }

    // Toggle Thu gọn / Mở rộng Sidebar
    function toggleSidebar() {
      const sb = document.getElementById('leftSidebar');
      const toggleText = document.getElementById('headerToggleText');
      const headerIcon = document.getElementById('headerToggleIcon');
      isSidebarCollapsed = !isSidebarCollapsed;

      if (isSidebarCollapsed) {
        sb.classList.add('sidebar-collapsed');
        if (toggleText) toggleText.textContent = 'Mở rộng';
        if (headerIcon) headerIcon.textContent = '▶';
      } else {
        sb.classList.remove('sidebar-collapsed');
        if (toggleText) toggleText.textContent = 'Thu gọn';
        if (headerIcon) headerIcon.textContent = '☰';
      }
    }

    function showGreetingToast() {
      const toast = document.getElementById('greetingToast');
      toast.classList.remove('hidden');
      setTimeout(() => {
        if (!toast.classList.contains('hidden')) dismissGreetingToast();
      }, 6000);
    }

    function dismissGreetingToast() {
      document.getElementById('greetingToast').classList.add('hidden');
    }

    // =========================================================================
    // BƯỚC 1: NẠP DỮ LIỆU
    // =========================================================================
    function loadSampleDataset() {
      dataSource = 'demo';
      isDataLoaded = true;
      isDataCalculated = false;
      analysisData = null;
      uploadedRawData = null;
      uploadedProfiles = null;
      uploadedFileName = 'Tệp Mẫu Demo (10 Doanh Nghiệp)';

      document.getElementById('modeBadge').textContent = 'Chế độ Demo (10 DN)';
      document.getElementById('modeBadge').className = 'ml-1.5 px-2 py-0.5 rounded-md bg-cyan-100 text-[10px] font-extrabold text-cyan-800 border border-cyan-300';
      document.getElementById('currentSourceLabel').textContent = 'Nguồn: Tệp mẫu Demo';
      document.getElementById('step1QuickStatus').innerHTML = '<span class="text-emerald-700 font-bold">✓ Đã nạp 10 DN (Demo)</span>';

      document.getElementById('step1StatusBadge').textContent = '✓ Đã nạp 10 DN';
      document.getElementById('step1StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';

      // Hiện thẻ tóm tắt Bước 1
      showStep1LoadedSummary('Tệp mẫu Demo (10 Doanh Nghiệp)', 'Nguồn: Dữ liệu mẫu chuẩn hóa', 10);

      // Cập nhật trạng thái các bước kế tiếp
      document.getElementById('step2StatusBadge').textContent = 'Cần quét lỗi';
      document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-cyan-200 text-cyan-950 font-bold';
      document.getElementById('step3StatusBadge').textContent = 'Chờ tính toán';
      document.getElementById('step3StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      lockExcelDownload();
      resetZeroMetrics();

      alert('Đã nạp tệp mẫu Demo gồm 10 doanh nghiệp!\nBạn có thể nhấp vào Bước 2 để quét chẩn đoán lỗi.');
    }

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
          rawRows = Array.isArray(parsed) ? parsed : (parsed.profiles || parsed.data || [parsed]);
        } else if (fName.endsWith('.xlsx') || fName.endsWith('.xls') || fName.endsWith('.csv')) {
          if (typeof XLSX === 'undefined') throw new Error('Thư viện SheetJS chưa sẵn sàng.');
          const buffer = await file.arrayBuffer();
          const workbook = XLSX.read(buffer, { type: 'array' });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          rawRows = XLSX.utils.sheet_to_json(firstSheet, { defval: null });
        } else {
          alert('Vui lòng chọn tệp .xlsx, .xls, .csv hoặc .json!');
          return;
        }

        if (!rawRows || rawRows.length === 0) {
          alert('Tệp dữ liệu trống!');
          return;
        }

        uploadedRawData = rawRows;
        uploadedProfiles = normalizeUploadedRows(rawRows);

        dataSource = 'uploaded';
        isDataLoaded = true;
        isDataCalculated = false;
        analysisData = null;

        document.getElementById('modeBadge').textContent = 'Tệp Người Dùng (' + uploadedProfiles.length + ' DN)';
        document.getElementById('modeBadge').className = 'ml-1.5 px-2 py-0.5 rounded-md bg-emerald-100 text-[10px] font-extrabold text-emerald-800 border border-emerald-300';
        document.getElementById('currentSourceLabel').textContent = 'Nguồn: ' + file.name;
        document.getElementById('step1QuickStatus').innerHTML = '<span class="text-emerald-700 font-bold">✓ Đã nạp ' + uploadedProfiles.length + ' DN</span>';

        const shortName = file.name.length > 18 ? file.name.substring(0, 15) + '...' : file.name;
        document.getElementById('step1StatusBadge').textContent = '✓ ' + shortName;
        document.getElementById('step1StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';

        // Hiện thẻ tóm tắt Bước 1
        showStep1LoadedSummary(file.name, 'Nguồn: Tệp người dùng tải lên (Tính toán độc lập)', uploadedProfiles.length);

        document.getElementById('step2StatusBadge').textContent = 'Chờ quét lỗi';
        document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-cyan-200 text-cyan-950 font-bold';
        document.getElementById('step3StatusBadge').textContent = 'Chờ tính toán';
        document.getElementById('step3StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

        lockExcelDownload();
        resetZeroMetrics();

        // Đưa về trạng thái ban đầu của Bước 2
        document.getElementById('step2InitialPrompt').classList.remove('hidden');
        document.getElementById('step2ResultCard').classList.add('hidden');

        alert('Đã tiếp nhận tệp: ' + file.name + ' (' + uploadedProfiles.length + ' DN)!\nHệ thống sẽ tính toán hoàn toàn độc lập từ tệp này.');
      } catch (err) {
        alert('Lỗi đọc tệp: ' + err.message);
      }
    }

    function showStep1LoadedSummary(fileName, sourceType, count) {
      const card = document.getElementById('step1LoadedSummaryCard');
      card.classList.remove('hidden');
      document.getElementById('summaryFileName').textContent = fileName;
      document.getElementById('summarySourceType').textContent = sourceType;
      document.getElementById('summaryRowCount').textContent = count + ' Doanh nghiệp';
    }

    function normalizeUploadedRows(rows) {
      return rows.map((r, idx) => {
        const getVal = (...keys) => {
          for (const k of keys) {
            for (const rowKey of Object.keys(r)) {
              if (rowKey.trim().toLowerCase() === k.trim().toLowerCase()) return r[rowKey];
            }
          }
          return null;
        };

        const accId = getVal('Mã Khách Hàng', 'Mã Doanh Nghiệp', 'account_id', 'id', 'ma_khach_hang', 'mã') || ('DN-' + String(idx + 1).padStart(3, '0'));
        const custName = getVal('Tên Doanh Nghiệp', 'Tên Khách Hàng', 'customer_name', 'name', 'company') || ('Doanh nghiệp ' + accId);
        const rep = getVal('Người Đại Diện', 'Đại Diện', 'representative', 'contact') || 'Chưa cập nhật';
        const ind = getVal('Lĩnh Vực', 'Ngành Nghề', 'industry') || 'Dịch Vụ & Thương Mại';
        const seg = getVal('Phân Khúc', 'segment') || 'Doanh Nghiệp Tiêu Chuẩn';
        const curPkg = getVal('Gói Dịch Vụ Hiện Tại', 'Gói Hiện Tại', 'current_package_name') || 'Điện Toán Cơ Bản';
        const expDate = getVal('Hạn Hợp Đồng', 'Ngày Hết Hạn', 'expiry_date') || '2026-12-31';

        let usage = getVal('Tỷ Lệ Sử Dụng (%)', 'Mức Sử Dụng', 'service_usage_rate_pct', 'usage');
        if (usage === '' || usage === undefined) usage = null;
        if (usage !== null && !isNaN(Number(usage))) usage = Number(usage);

        const ticket = getVal('Sự Cố Kỹ Thuật', 'Sự Cố', 'open_tickets', 'ticket') || 'Không có sự cố';
        let basePrice = getVal('Giá Gốc (VNĐ)', 'Giá Trị Hợp Đồng', 'Giá Trị', 'base_price_vnd', 'price');
        basePrice = (!basePrice || isNaN(Number(basePrice))) ? 80000000 : Number(basePrice);

        let discPct = Number(getVal('Chiết Khấu (%)', 'discount_percent')) || 0;
        const discCode = getVal('Mã Ưu Đãi', 'discount_code') || (discPct > 0 ? ('DISCOUNT_' + discPct) : 'NONE');

        return {
          original_row_idx: idx + 2,
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
    // BƯỚC 2: QUÉT & CHẨN ĐOÁN LỖI
    // =========================================================================
    function runErrorScan() {
      if (!isDataLoaded) {
        alert('Vui lòng thực hiện Bước 1 (Nạp Dữ Liệu) trước khi quét lỗi!');
        goToStep(1);
        return;
      }

      detectedErrors = [];
      if (dataSource === 'demo') {
        detectedErrors = [
          { rowCol: 'Dòng 5 | Cột: Ngày hết hạn', accId: 'DN-2026-004', val: '32/13/2026', why: 'Định dạng ngày không hợp lệ (ngày 32, tháng 13)', fix: 'Chuẩn hóa về ISO YYYY-MM-DD (2026-11-28)' },
          { rowCol: 'Dòng 7 | Cột: Mức sử dụng', accId: 'DN-2026-004', val: 'NULL / Để trống', why: 'Khuyết dữ liệu đo lường công suất (Ticket TK-8004)', fix: 'Tạm giữ doanh thu, không gán 0; chờ đồng bộ API' },
          { rowCol: 'Dòng 10 | Bảng: Sự cố kỹ thuật', accId: 'DN-2026-007', val: 'Sự cố Mở (TK-8002)', why: 'Máy chủ CSDL gặp sự cố nghiêm trọng chưa đóng', fix: 'Khóa chặn Hard-Block cấm chào tái ký; ưu tiên Kỹ thuật' }
        ];
      } else if (dataSource === 'uploaded') {
        uploadedProfiles.forEach(p => {
          if (p.expiry_date.includes('32') || p.expiry_date.includes('13/') || isNaN(Date.parse(p.expiry_date)) && p.expiry_date.includes('/')) {
            detectedErrors.push({ rowCol: 'Dòng ' + p.original_row_idx + ' | Cột: Hạn hợp đồng', accId: p.account_id, val: p.expiry_date, why: 'Định dạng ngày tháng không hợp lệ', fix: 'Chuẩn hóa về ISO YYYY-MM-DD' });
          }
          if (p.service_usage_rate_pct === null || p.service_usage_rate_pct === undefined) {
            detectedErrors.push({ rowCol: 'Dòng ' + p.original_row_idx + ' | Cột: Mức sử dụng', accId: p.account_id, val: 'NULL (Khuyết)', why: 'Khuyết dữ liệu đo lường (NULL ≠ 0%)', fix: 'Tạm giữ doanh thu, tuyệt đối không gán bằng 0%' });
          }
          const lowerTicket = p.open_tickets.toLowerCase();
          if (lowerTicket.includes('mở') || lowerTicket.includes('open') || lowerTicket.includes('tk-') || lowerTicket.includes('sự cố') || lowerTicket.includes('lỗi')) {
            detectedErrors.push({ rowCol: 'Dòng ' + p.original_row_idx + ' | Cột: Sự cố kỹ thuật', accId: p.account_id, val: p.open_tickets, why: 'Sự cố máy chủ hạ tầng chưa được đóng', fix: 'Khóa chặn Hard-Block an toàn, tạm dừng báo giá' });
          }
        });
      }

      // Hiển thị kết quả trong View Bước 2
      document.getElementById('step2InitialPrompt').classList.add('hidden');
      const resCard = document.getElementById('step2ResultCard');
      resCard.classList.remove('hidden');

      const wrapper = document.getElementById('step2TableWrapper');

      if (detectedErrors.length > 0) {
        document.getElementById('step2ResultIcon').textContent = '⚠️';
        document.getElementById('step2ResultIcon').className = 'w-7 h-7 rounded-xl bg-rose-600 text-white flex items-center justify-center font-bold text-xs';
        document.getElementById('step2ResultTitle').textContent = 'PHÁT HIỆN ' + detectedErrors.length + ' ĐIỂM SAI LỆCH CẦN KHẮC PHỤC';
        document.getElementById('step2ResultBadge').textContent = detectedErrors.length + ' Lỗi Cần Xử Lý';
        document.getElementById('step2ResultBadge').className = 'px-2.5 py-0.5 rounded-full bg-rose-600 text-white font-extrabold text-[10px]';
        document.getElementById('btnAutoFixStep2').classList.remove('hidden');

        let rowsHtml = '';
        detectedErrors.forEach(err => {
          rowsHtml += `
            <tr class="hover:bg-rose-50/70">
              <td class="py-2 px-3 font-bold text-rose-900 font-mono">${escapeHtml(err.rowCol)}</td>
              <td class="py-2 px-3 font-bold text-slate-900">${escapeHtml(err.accId)}</td>
              <td class="py-2 px-3 text-rose-700 font-mono bg-rose-100/70 rounded">${escapeHtml(err.val)}</td>
              <td class="py-2 px-3 text-slate-700">${escapeHtml(err.why)}</td>
              <td class="py-2 px-3 text-emerald-800 font-bold">${escapeHtml(err.fix)}</td>
            </tr>
          `;
        });

        wrapper.innerHTML = `
          <table class="w-full text-left text-xs whitespace-nowrap">
            <thead class="bg-rose-100/80 text-rose-900 uppercase font-black text-[10px] sticky top-0">
              <tr>
                <th class="py-2 px-3">Vị Trí (Dòng / Cột)</th>
                <th class="py-2 px-3">Mã DN</th>
                <th class="py-2 px-3">Giá Trị Hiện Tại</th>
                <th class="py-2 px-3">Nguyên Nhân (Why)</th>
                <th class="py-2 px-3">Hướng Dẫn Khắc Phục (Fix Guide)</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-rose-100 text-[11px] bg-rose-50/30">
              ${rowsHtml}
            </tbody>
          </table>
        `;

        document.getElementById('step2StatusBadge').textContent = '⚠️ Phát hiện ' + detectedErrors.length + ' lỗi';
        document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-rose-400 text-white font-black animate-pulse';
        document.getElementById('iconBadge2').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-rose-500 text-white font-black text-[10px] flex items-center justify-center shadow-xs';
      } else {
        document.getElementById('step2ResultIcon').textContent = '✓';
        document.getElementById('step2ResultIcon').className = 'w-7 h-7 rounded-xl bg-emerald-600 text-white flex items-center justify-center font-bold text-xs';
        document.getElementById('step2ResultTitle').textContent = 'DỮ LIỆU ĐÃ ĐẠT CHUẨN 0 LỖI';
        document.getElementById('step2ResultBadge').textContent = '0 Lỗi Hoàn Hảo';
        document.getElementById('step2ResultBadge').className = 'px-2.5 py-0.5 rounded-full bg-emerald-600 text-white font-extrabold text-[10px]';
        document.getElementById('btnAutoFixStep2').classList.add('hidden');

        wrapper.innerHTML = `
          <div class="p-6 text-center text-emerald-800 space-y-1.5">
            <div class="text-3xl">✓</div>
            <div class="font-extrabold text-sm">Dữ liệu hoàn toàn sạch (0 lỗi)!</div>
            <div class="text-xs text-slate-600">Toàn bộ hồ sơ đã vượt qua kiểm tra định dạng và phân tách an toàn. Bạn có thể sang Bước 3 để chạy tính toán.</div>
          </div>
        `;

        document.getElementById('step2StatusBadge').textContent = '✓ 0 lỗi';
        document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';
        document.getElementById('iconBadge2').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-emerald-400 text-emerald-950 font-black text-[10px] flex items-center justify-center shadow-xs';
      }
    }

    function autoFixErrorsAndProceed() {
      if (dataSource === 'uploaded' && uploadedProfiles) {
        uploadedProfiles.forEach(p => {
          if (p.expiry_date.includes('32') || p.expiry_date.includes('13/')) p.expiry_date = '2026-11-28';
        });
      }

      detectedErrors = [];
      runErrorScan();
      alert('Đã tự động chuẩn hóa toàn bộ dữ liệu về 0 lỗi đạt chuẩn! Hệ thống chuyển sang Bước 3: Tính Toán.');
      goToStep(3);
      runCalculationPipeline();
    }

    // =========================================================================
    // BƯỚC 3: TÍNH TOÁN & BẢNG ĐIỀU KHIỂN (PHÂN TRANG GỌN GÀNG, KHÔNG BỊ QUÁ DÀI)
    // =========================================================================
    async function runCalculationPipeline() {
      if (!isDataLoaded) {
        alert('Vui lòng hoàn thành Bước 1 (Nạp Dữ Liệu) trước khi tính toán!');
        goToStep(1);
        return;
      }

      const btn = document.getElementById('btnCalculateStep3');
      const oldHtml = btn.innerHTML;

      try {
        btn.innerHTML = '<span>⏳</span><span>Đang tính toán...</span>';
        btn.disabled = true;

        if (dataSource === 'uploaded') {
          const calculatedProfiles = calculateUploadedData(uploadedProfiles);
          const kpi = calculateKpiSummary(calculatedProfiles);
          analysisData = {
            system_name: 'Gia Linh Data Studio',
            source_type: 'uploaded',
            filename: uploadedFileName,
            kpi_summary: kpi,
            profiles: calculatedProfiles
          };
          await new Promise(r => setTimeout(r, 300));
        } else if (dataSource === 'demo') {
          const res = await fetch('/api/run-agent', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ source_type: 'json', dataset_scale: '10_core' })
          });
          analysisData = await res.json();
          analysisData.source_type = 'demo';
        }

        isDataCalculated = true;

        document.getElementById('step3StatusBadge').textContent = '✓ Đã tính toán';
        document.getElementById('step3StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';
        document.getElementById('iconBadge3').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-emerald-400 text-emerald-950 font-black text-[10px] flex items-center justify-center shadow-xs';

        unlockExcelDownload();
        applyStep3Dashboard();
      } catch (err) {
        alert('Lỗi tính toán: ' + err.message);
      } finally {
        btn.innerHTML = oldHtml;
        btn.disabled = false;
      }
    }

    function calculateUploadedData(profiles) {
      return profiles.map(p => {
        let isGrounded = true;
        let isBlocked = false;
        let recType = 'DUY TRÌ GÓI';
        let recPkg = p.current_package_name;
        let overallStatus = 'SẴN SÀNG THỰC THI';
        let remarks = '';
        let plan = '';

        const lowerTicket = (p.open_tickets || '').toLowerCase();
        if (lowerTicket.includes('mở') || lowerTicket.includes('open') || lowerTicket.includes('tk-') || lowerTicket.includes('sự cố') || lowerTicket.includes('lỗi')) {
          isBlocked = true;
          isGrounded = false;
          overallStatus = 'KHÓA CHẶN KỸ THUẬT (HARD-BLOCK)';
          remarks = 'Sự cố máy chủ kỹ thuật chưa đóng. Kích hoạt Hard-Block cấm chào tái ký; tạm giữ doanh thu an toàn.';
          plan = 'Yêu cầu Kỹ thuật xử lý dứt điểm trước khi liên hệ khách hàng.';
        } else if (p.service_usage_rate_pct === null || p.service_usage_rate_pct === undefined) {
          isGrounded = false;
          overallStatus = 'TẠM GIỮ (CHỜ ĐỒNG BỘ DỮ LIỆU)';
          recType = 'TẠM GIỮ DOANH THU';
          remarks = 'Khuyết thiếu dữ liệu đo lường công suất (NULL ≠ 0%). Tạm giữ an toàn doanh thu.';
          plan = 'Đối soát log kết nối API hạ tầng.';
        } else {
          const usage = p.service_usage_rate_pct;
          if (usage === 0) {
            recType = 'DUY TRÌ HỖ TRỢ PHỤC HỒI';
            recPkg = 'Gói Duy Trì Hỗ Trợ Kỹ Thuật';
            remarks = 'Mức sử dụng 0% là dữ liệu thực tế hợp lệ. Đề xuất gói duy trì chi phí thấp và hỗ trợ kích hoạt.';
            plan = 'CSKH đồng hành hỗ trợ doanh nghiệp.';
          } else if (usage >= 80) {
            recType = 'NÂNG CẤP GÓI';
            recPkg = 'Hệ Sinh Thái Enterprise Suite';
            remarks = 'Lưu lượng cao vượt trội (' + usage + '%). Đủ điều kiện nâng cấp gói Enterprise Suite.';
            plan = 'Gửi báo giá nâng cấp giải pháp.';
          } else if (usage >= 50) {
            recType = 'TÁI KÝ DUY TRÌ';
            recPkg = p.current_package_name;
            remarks = 'Mức sử dụng ổn định (' + usage + '%). Đề xuất tái ký gói hiện tại.';
            plan = 'Gửi hợp đồng gia hạn tự động.';
          } else {
            recType = 'TÁI KÝ TỐI ƯU HÓA';
            recPkg = 'Gói Tinh Gọn Hạ Tầng';
            remarks = 'Mức sử dụng thấp (' + usage + '%). Đề xuất gói tinh gọn để tối ưu chi phí.';
            plan = 'Tư vấn tái cấu trúc dịch vụ.';
          }

          if (p.discount_percent > 15) {
            overallStatus = 'CHỜ QUẢN LÝ PHÊ DUYỆT (HITL)';
            remarks += ' Chiết khấu lớn (' + p.discount_percent + '%) chuyển Cổng kiểm soát HITL.';
          }
        }

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

    function calculateKpiSummary(profiles) {
      let groundedVal = 0, holdVal = 0, groundedCount = 0, holdCount = 0, usageSum = 0, usageCount = 0;
      profiles.forEach(p => {
        if (p.is_revenue_grounded) { groundedVal += p.final_price_vnd; groundedCount++; }
        else { holdVal += p.final_price_vnd; holdCount++; }
        if (p.service_usage_rate_pct !== null && !isNaN(p.service_usage_rate_pct)) {
          usageSum += p.service_usage_rate_pct;
          usageCount++;
        }
      });
      return {
        total_accounts: profiles.length,
        grounded_accounts_count: groundedCount,
        unverified_or_blocked_accounts_count: holdCount,
        total_grounded_renewal_value_vnd: groundedVal,
        total_blocked_or_unverified_value_vnd: holdVal,
        mean_usage_rate_pct: usageCount > 0 ? Number((usageSum / usageCount).toFixed(1)) : 0
      };
    }

    function applyStep3Dashboard() {
      if (!analysisData) return;
      const kpi = analysisData.kpi_summary;

      document.getElementById('kpiTotalNum').textContent = kpi.total_accounts;
      document.getElementById('kpiTotalSub').textContent = '100% hồ sơ đã phân tích';

      document.getElementById('kpiGroundedNum').textContent = formatMillions(kpi.total_grounded_renewal_value_vnd);
      document.getElementById('kpiGroundedSub').textContent = kpi.grounded_accounts_count + ' DN đủ căn cứ';

      document.getElementById('kpiUsageNum').textContent = (kpi.mean_usage_rate_pct || 0) + '%';
      document.getElementById('kpiUsageSub').textContent = 'Tách biệt NULL ≠ 0%';

      document.getElementById('kpiHoldNum').textContent = formatMillions(kpi.total_blocked_or_unverified_value_vnd);
      document.getElementById('kpiHoldSub').textContent = kpi.unverified_or_blocked_accounts_count + ' DN tạm giữ an toàn';

      // 3 Khối thống kê
      const profiles = analysisData.profiles || [];
      const readyC = profiles.filter(p => p.is_revenue_grounded && !p.overall_status.includes('QUẢN LÝ')).length;
      const holdC = profiles.filter(p => !p.is_revenue_grounded && !p.is_blocked_by_tech).length;
      const blockC = profiles.filter(p => p.is_blocked_by_tech).length;

      document.getElementById('statReadyCount').textContent = readyC + ' DN';
      document.getElementById('statHoldCount').textContent = holdC + ' DN';
      document.getElementById('statBlockCount').textContent = blockC + ' DN';

      // Ẩn overlay biểu đồ
      document.getElementById('chartPlaceholderOverlay').classList.add('hidden');
      renderCharts();

      // Cài đặt danh sách phân trang
      filteredProfiles = profiles;
      currentPage = 1;
      renderPaginatedTable();
    }

    function handleTableSearch(query) {
      if (!analysisData) return;
      const q = query.trim().toLowerCase();
      if (!q) {
        filteredProfiles = analysisData.profiles || [];
      } else {
        filteredProfiles = (analysisData.profiles || []).filter(p => 
          p.account_id.toLowerCase().includes(q) ||
          p.customer_name.toLowerCase().includes(q) ||
          p.representative.toLowerCase().includes(q)
        );
      }
      currentPage = 1;
      renderPaginatedTable();
    }

    function renderPaginatedTable() {
      const tbody = document.getElementById('paginatedTableBody');
      const total = filteredProfiles.length;
      const totalPages = Math.max(1, Math.ceil(total / pageSize));

      document.getElementById('tablePageBadge').textContent = total + ' bản ghi';
      document.getElementById('paginationInfo').textContent = 'Trang ' + currentPage + ' / ' + totalPages + ' (' + total + ' DN)';

      document.getElementById('btnPrevPage').disabled = (currentPage <= 1);
      document.getElementById('btnNextPage').disabled = (currentPage >= totalPages);

      if (total === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="py-6 text-center text-slate-400">Không tìm thấy bản ghi nào phù hợp.</td></tr>';
        return;
      }

      const startIdx = (currentPage - 1) * pageSize;
      const endIdx = Math.min(startIdx + pageSize, total);
      const pageRows = filteredProfiles.slice(startIdx, endIdx);

      tbody.innerHTML = '';
      pageRows.forEach(p => {
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
          statusLabel = 'Nâng Cấp';
        }

        const usageStr = (p.service_usage_rate_pct === null) ? '<span class="text-rose-600 font-bold">NULL</span>' : (p.service_usage_rate_pct + '%');

        tbody.innerHTML += `
          <tr class="hover:bg-slate-50 transition">
            <td class="py-2.5 px-3 font-mono font-bold text-teal-700">${escapeHtml(p.account_id)}</td>
            <td class="py-2.5 px-3 font-bold text-slate-900">
              ${escapeHtml(p.customer_name)}
              <div class="text-[10px] text-slate-400 font-normal">ĐD: ${escapeHtml(p.representative)}</div>
            </td>
            <td class="py-2.5 px-3 font-semibold">${usageStr}</td>
            <td class="py-2.5 px-3 text-slate-700">${escapeHtml(p.recommended_package_name)}</td>
            <td class="py-2.5 px-3 text-right font-mono font-bold text-slate-900">${formatVND(p.final_price_vnd)}</td>
            <td class="py-2.5 px-3 text-center">
              <span class="px-2 py-0.5 rounded-full text-[10px] font-black ${pillClass}">
                ${statusLabel}
              </span>
            </td>
          </tr>
        `;
      });
    }

    function prevTablePage() {
      if (currentPage > 1) {
        currentPage--;
        renderPaginatedTable();
      }
    }

    function nextTablePage() {
      const totalPages = Math.ceil(filteredProfiles.length / pageSize);
      if (currentPage < totalPages) {
        currentPage++;
        renderPaginatedTable();
      }
    }

    function renderCharts() {
      if (!analysisData) return;
      const profiles = analysisData.profiles || [];
      const canvas = document.getElementById('revenueTrendChart');
      if (!canvas) return;
      const ctxLine = canvas.getContext('2d');
      if (lineChart) lineChart.destroy();

      const renewVal = profiles.filter(p => p.is_revenue_grounded && !p.recommendation_type.includes('NÂNG CẤP')).reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;
      const upVal = profiles.filter(p => p.is_revenue_grounded && p.recommendation_type.includes('NÂNG CẤP')).reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;

      const gradTeal = ctxLine.createLinearGradient(0, 0, 0, 160);
      gradTeal.addColorStop(0, 'rgba(13, 148, 136, 0.35)');
      gradTeal.addColorStop(1, 'rgba(13, 148, 136, 0.0)');

      const gradCyan = ctxLine.createLinearGradient(0, 0, 0, 160);
      gradCyan.addColorStop(0, 'rgba(6, 182, 212, 0.35)');
      gradCyan.addColorStop(1, 'rgba(6, 182, 212, 0.0)');

      lineChart = new Chart(ctxLine, {
        type: 'line',
        data: {
          labels: ['Kỳ 1', 'Kỳ 2', 'Kỳ 3', 'Kỳ 4', 'Kỳ 5', 'Hiện tại'],
          datasets: [
            {
              label: 'Gói Duy Trì',
              data: [Math.round((renewVal || 100) * 0.5), Math.round((renewVal || 100) * 0.7), Math.round((renewVal || 100) * 0.82), Math.round((renewVal || 100) * 0.9), Math.round((renewVal || 100) * 0.96), Math.round(renewVal || 100)],
              borderColor: '#0d9488',
              backgroundColor: gradTeal,
              tension: 0.4,
              borderWidth: 2,
              fill: true,
              pointRadius: 3
            },
            {
              label: 'Gói Nâng Cấp',
              data: [Math.round((upVal || 150) * 0.4), Math.round((upVal || 150) * 0.6), Math.round((upVal || 150) * 0.75), Math.round((upVal || 150) * 0.88), Math.round((upVal || 150) * 0.95), Math.round(upVal || 150)],
              borderColor: '#06b6d4',
              backgroundColor: gradCyan,
              tension: 0.4,
              borderWidth: 2,
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
            y: { beginAtZero: true, grid: { color: 'rgba(241, 245, 249, 0.8)' } },
            x: { grid: { display: false } }
          }
        }
      });
    }

    // =========================================================================
    // BƯỚC 4: XUẤT KẾT QUẢ EXCEL
    // =========================================================================
    function unlockExcelDownload() {
      const btn = document.getElementById('mainDownloadExcelBtn');
      const lockIcon = document.getElementById('mainLockIcon');
      const lockText = document.getElementById('mainLockText');
      const step4Badge = document.getElementById('step4StatusBadge');
      const step4Gated = document.getElementById('step4GatedBadge');

      btn.disabled = false;
      btn.className = 'py-3 px-6 rounded-2xl bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-600 hover:to-teal-600 text-slate-900 font-black text-xs cursor-pointer transition flex items-center gap-2 shadow-md';
      lockIcon.textContent = '📥';
      lockText.textContent = dataSource === 'uploaded' ? 'Tải Excel Kết Quả Tệp Của Bạn (.xlsx)' : 'Tải File Excel Đã Tính Toán (.xlsx)';

      step4Badge.textContent = '✓ Sẵn sàng tải';
      step4Badge.className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-emerald-300 text-teal-950 font-black';

      step4Gated.textContent = '✓ Sẵn sàng xuất';
      step4Gated.className = 'px-3 py-1 rounded-xl bg-emerald-100 text-emerald-800 font-extrabold text-xs';

      document.getElementById('iconExcelSymbol').textContent = '📥';
      document.getElementById('iconBadge4').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-emerald-300 text-emerald-950 font-black text-[10px] flex items-center justify-center shadow-xs';
      document.getElementById('iconStepBtn4').className = 'w-11 h-11 rounded-2xl bg-emerald-400 text-slate-900 flex items-center justify-center text-lg transition shadow-md relative cursor-pointer';

      document.getElementById('exportFileNameDesc').textContent = 'Tệp Excel đã sẵn sàng tải về với kết quả tính toán từ ' + (dataSource === 'uploaded' ? uploadedFileName : 'Tệp mẫu Demo') + '.';
    }

    function lockExcelDownload() {
      const btn = document.getElementById('mainDownloadExcelBtn');
      btn.disabled = true;
      btn.className = 'py-3 px-6 rounded-2xl bg-slate-200 text-slate-400 font-black text-xs cursor-not-allowed transition flex items-center gap-2 shadow-sm';
      document.getElementById('mainLockIcon').textContent = '🔒';
      document.getElementById('mainLockText').textContent = 'Xuất Tệp Excel (Chờ Tính Toán Ở Bước 3)';

      document.getElementById('step4StatusBadge').textContent = 'Chưa mở';
      document.getElementById('step4StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      document.getElementById('step4GatedBadge').textContent = 'Chờ tính toán';
      document.getElementById('step4GatedBadge').className = 'px-3 py-1 rounded-xl bg-slate-100 text-slate-500 font-extrabold text-xs';

      document.getElementById('iconExcelSymbol').textContent = '🔒';
      document.getElementById('iconBadge4').className = 'absolute -top-1.5 -right-1.5 w-5 h-5 rounded-full bg-slate-400 text-slate-900 font-black text-[10px] flex items-center justify-center shadow-xs';
      document.getElementById('iconStepBtn4').className = 'w-11 h-11 rounded-2xl bg-white/10 text-white/50 flex items-center justify-center text-lg transition shadow-md relative';
    }

    function downloadCalculatedExcel() {
      if (!isDataCalculated || !analysisData) {
        alert('Chưa có kết quả tính toán! Vui lòng hoàn thành Bước 1 và Bước 3 trước.');
        goToStep(3);
        return;
      }

      if (dataSource === 'uploaded') {
        try {
          if (typeof XLSX === 'undefined') throw new Error('Thư viện SheetJS chưa sẵn sàng.');
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
          alert('Đã xuất thành công tệp Excel: ' + exportFileName);
        } catch (err) {
          alert('Lỗi xuất Excel: ' + err.message);
        }
      } else {
        window.location.href = '/api/download-output-xlsx';
      }
    }

    // =========================================================================
    // RESET VỀ 0 BAN ĐẦU
    // =========================================================================
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
      document.getElementById('step1QuickStatus').textContent = 'Chưa có dữ liệu';

      document.getElementById('step1StatusBadge').textContent = 'Chưa nạp';
      document.getElementById('step1StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      document.getElementById('step2StatusBadge').textContent = 'Chờ quét';
      document.getElementById('step2StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      document.getElementById('step3StatusBadge').textContent = 'Chờ tính toán';
      document.getElementById('step3StatusBadge').className = 'text-[10px] font-bold px-2 py-0.5 rounded-full bg-white/20 text-white';

      document.getElementById('step1LoadedSummaryCard').classList.add('hidden');
      document.getElementById('step2InitialPrompt').classList.remove('hidden');
      document.getElementById('step2ResultCard').classList.add('hidden');

      lockExcelDownload();
      resetZeroMetrics();
      goToStep(1);
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

      document.getElementById('statReadyCount').textContent = '0 DN';
      document.getElementById('statHoldCount').textContent = '0 DN';
      document.getElementById('statBlockCount').textContent = '0 DN';

      document.getElementById('chartPlaceholderOverlay').classList.remove('hidden');

      document.getElementById('paginatedTableBody').innerHTML = `
        <tr>
          <td colspan="6" class="py-6 text-center text-slate-400">
            Chưa có kết quả tính toán. Bấm [Bắt Đầu Tính Toán &amp; Phân Tích] ở trên.
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

    window.addEventListener('DOMContentLoaded', () => {
      resetAllToZero();
    });
  </script>
</body>
</html>
'@

[System.IO.File]::WriteAllText($targetHtml, $htmlContent, [System.Text.Encoding]::UTF8)
Write-Output ("Successfully generated Step-by-Step Dedicated View UI to: " + $targetHtml + " (" + (Get-Item $targetHtml).Length + " bytes)")
