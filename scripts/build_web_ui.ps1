# Generator script for web/index.html with:
# 1. Pastel palette: Soft Pink, Soft Purple, Soft Blue/Teal, Soft Yellow
# 2. Pure data analytics tool: ZERO metrics or data displayed initially
# 3. Dedicated "Demo Nhanh" bar that automatically loads simulated data, computes, and displays the calculation dashboard
# 4. Interactive calculation engine with real-time Dashboard, Charts, HITL, Proposals, and Excel export
# 5. Collapsible vertical sidebar and clean elegant typography

$html = @'
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Công cụ Phân tích Dữ liệu Gia hạn EdTech Future</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
  <script>
    tailwind.config = {
      theme: {
        extend: {
          colors: {
            pastel: {
              pink: '#fff1f2',
              'pink-border': '#fecdd3',
              'pink-text': '#9f1239',
              'pink-accent': '#e11d48',
              purple: '#f5f3ff',
              'purple-border': '#ddd6fe',
              'purple-text': '#5b21b6',
              'purple-accent': '#7c3aed',
              blue: '#f0f9ff',
              'blue-border': '#bae6fd',
              'blue-text': '#075985',
              'blue-accent': '#0284c7',
              teal: '#ecfdf5',
              'teal-border': '#a7f3d0',
              'teal-text': '#064e3b',
              'teal-accent': '#059669',
              yellow: '#fefce8',
              'yellow-border': '#fef08a',
              'yellow-text': '#854d0e',
              'yellow-accent': '#d97706'
            }
          }
        }
      }
    }
  </script>
  <style>
    * {
      -webkit-font-smoothing: antialiased;
      -moz-osx-font-smoothing: grayscale;
    }
    body {
      font-family: 'Plus Jakarta Sans', system-ui, -apple-system, sans-serif;
      background: linear-gradient(135deg, #fff5f5 0%, #fbf5fe 25%, #f0fdf9 60%, #fefef2 100%);
      color: #0f172a;
      min-height: 100vh;
    }
    .custom-scroll::-webkit-scrollbar { width: 4px; height: 4px; }
    .custom-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 9999px; }
    .custom-scroll::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
    .app-window {
      background: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 20px;
      box-shadow: 0 16px 40px -10px rgba(15, 23, 42, 0.07);
    }
    .step-active {
      background-color: #f5f3ff !important;
      border: 1px solid #7c3aed !important;
      color: #5b21b6 !important;
      font-weight: 700 !important;
    }
    .step-active .step-icon {
      background-color: #7c3aed !important;
      color: #ffffff !important;
    }
    .step-locked {
      opacity: 0.55;
    }
    .sidebar-collapsed {
      width: 4.75rem !important; /* 76px */
    }
    .sidebar-collapsed .sidebar-text {
      display: none !important;
    }
    .sidebar-collapsed .sidebar-header {
      justify-content: center !important;
    }
    .sidebar-collapsed .sidebar-header-title {
      display: none !important;
    }
  </style>
</head>
<body class="p-2 sm:p-5 flex items-center justify-center">

  <!-- ========================================================================= -->
  <!-- MAIN TOOL WINDOW                                                          -->
  <!-- ========================================================================= -->
  <div class="app-window w-full max-w-[1580px] min-h-[92vh] flex flex-col overflow-hidden bg-white">

    <!-- TOP HEADER BAR: macOS Dots, Title, Collapse Toggle, Reset -->
    <header class="px-5 py-3 border-b border-slate-100 flex items-center justify-between text-xs bg-white select-none">
      <!-- Left: Dots & Logo -->
      <div class="flex items-center gap-3">
        <!-- Window Dots -->
        <div class="flex items-center gap-1.5 mr-2">
          <span class="w-3 h-3 rounded-full bg-rose-400 inline-block" title="Pastel Rose"></span>
          <span class="w-3 h-3 rounded-full bg-amber-300 inline-block" title="Pastel Yellow"></span>
          <span class="w-3 h-3 rounded-full bg-emerald-400 inline-block" title="Pastel Green"></span>
        </div>

        <!-- Sidebar Toggle Button -->
        <button onclick="toggleSidebar()" class="p-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 text-slate-600 transition" title="Đóng / Mở menu dọc">
          <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 12h10M4 18h16" />
          </svg>
        </button>

        <!-- Tool Brand -->
        <div class="flex items-center gap-2 pl-2 border-l border-slate-200">
          <div class="w-6 h-6 rounded-lg bg-gradient-to-tr from-purple-600 to-indigo-600 text-white flex items-center justify-center font-bold text-xs shadow-xs">
            ⚡
          </div>
          <span class="font-extrabold text-sm text-slate-800 tracking-tight">DataLens Studio</span>
          <span class="text-[10px] px-2 py-0.5 rounded-full bg-purple-50 text-purple-700 font-bold border border-purple-200/60">Công cụ Phân tích Dữ liệu</span>
        </div>
      </div>

      <!-- Center: Status Text -->
      <div class="hidden md:flex items-center gap-2 text-slate-600 font-semibold bg-slate-50 px-3.5 py-1 rounded-lg border border-slate-200/60 text-[11px]">
        <span id="headerStatusDot" class="w-2 h-2 rounded-full bg-slate-400"></span>
        <span id="centerHeaderTitle">Chế độ chờ: Chưa có dữ liệu phân tích</span>
      </div>

      <!-- Right Actions: AI, Reset, Help -->
      <div class="flex items-center gap-2">
        <button onclick="toggleAiChatDrawer()" class="text-purple-700 hover:text-purple-800 font-semibold text-xs flex items-center gap-1.5 bg-purple-50 hover:bg-purple-100/70 px-3 py-1.5 rounded-lg border border-purple-200/70 transition">
          <span>✦ Hỏi AI Agent</span>
        </button>
        <button onclick="resetAllDataToBlank()" class="text-slate-500 hover:text-rose-600 font-medium text-xs flex items-center gap-1 px-2.5 py-1.5 rounded-lg hover:bg-slate-50 transition" title="Xóa toàn bộ dữ liệu để về trạng thái trắng ban đầu">
          <span>↺ Xóa dữ liệu / Đặt lại</span>
        </button>
        <button onclick="showHelpModal()" class="text-slate-500 hover:text-slate-800 font-medium text-xs flex items-center gap-1 px-2.5 py-1.5 rounded-lg hover:bg-slate-50 transition">
          <span>? Hướng dẫn</span>
        </button>
      </div>
    </header>

    <!-- PROMINENT QUICK DEMO BAR (THANH DEMO NHANH NỔI BẬT) -->
    <div class="px-5 py-2.5 bg-gradient-to-r from-pastel-pink via-pastel-purple via-pastel-blue to-pastel-yellow border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 select-none">
      <div class="flex items-center gap-3">
        <span class="px-2.5 py-1 rounded-xl bg-purple-600 text-white font-extrabold text-[11px] shadow-xs flex items-center gap-1">
          <span>⚡</span> DEMO NHANH
        </span>
        <div>
          <span class="font-extrabold text-xs text-slate-800">Dữ liệu Giả lập EdTech Future (Gia hạn hàng tháng)</span>
          <span class="hidden lg:inline text-[11px] text-slate-500 ml-2">— Nhấn để tự động nạp dữ liệu, kích hoạt Agent tính toán và hiển thị Dashboard tức thì.</span>
        </div>
      </div>

      <div class="flex items-center gap-2">
        <button onclick="runQuickDemo('10_core')" class="px-4 py-1.5 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-extrabold text-xs shadow-xs transition flex items-center gap-1.5">
          <span>🚀 Chạy Demo Nhanh (10 Hồ Sơ Chuẩn)</span>
        </button>
        <button onclick="runQuickDemo('150_monthly')" class="px-3.5 py-1.5 rounded-xl bg-white hover:bg-purple-50 text-purple-700 border border-purple-200 font-bold text-xs transition">
          <span>Toàn kỳ 150 Tài Khoản / Tháng</span>
        </button>
      </div>
    </div>

    <!-- MAIN BODY GRID: COLLAPSIBLE SIDEBAR + CONTENT AREA -->
    <div class="flex-1 flex flex-col md:flex-row overflow-hidden">

      <!-- ===================================================================== -->
      <!-- 1. LEFT SIDEBAR: QUY TRÌNH 4 BƯỚC (COLLAPSIBLE)                        -->
      <!-- ===================================================================== -->
      <aside id="leftSidebar" class="w-full md:w-64 bg-slate-50/70 border-r border-slate-100 p-3.5 flex flex-col justify-between shrink-0 select-none transition-all duration-300">
        <div class="space-y-3.5">
          <!-- Sidebar Header -->
          <div class="sidebar-header flex items-center justify-between px-2">
            <span class="sidebar-header-title text-[11px] font-extrabold uppercase tracking-wider text-slate-400">QUY TRÌNH</span>
            <span id="sidebarStepBadge" class="text-[10px] text-purple-700 font-bold bg-purple-50 px-1.5 py-0.5 rounded border border-purple-200/50 sidebar-text">BƯỚC 1/4</span>
          </div>

          <!-- 4 Step Navigation -->
          <nav class="space-y-2">
            <!-- Bước 1: Nhập dữ liệu -->
            <button onclick="goToWorkflowStep(1)" id="stepNav-1" class="w-full text-left p-2.5 rounded-xl transition flex items-center justify-between text-xs step-active" title="Bước 1: Nhập dữ liệu từ các nguồn">
              <div class="flex items-center gap-3 overflow-hidden">
                <div class="step-icon w-8 h-8 rounded-lg bg-purple-600 text-white flex items-center justify-center shrink-0 shadow-2xs">
                  <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2.5" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1" />
                  </svg>
                </div>
                <div class="sidebar-text truncate">
                  <div class="font-bold text-slate-800 text-xs">1. Nhập Dữ liệu</div>
                  <div class="text-[10px] text-slate-500 truncate">Nhập nguồn hoặc Demo</div>
                </div>
              </div>
              <span id="stepIconStatus-1" class="sidebar-text text-purple-600 font-bold text-xs">›</span>
            </button>

            <!-- Bước 2: Làm sạch dữ liệu -->
            <button onclick="goToWorkflowStep(2)" id="stepNav-2" class="w-full text-left p-2.5 rounded-xl transition flex items-center justify-between text-xs bg-white hover:bg-slate-100/70 text-slate-600 step-locked" title="Bước 2: Làm sạch & Kiểm định">
              <div class="flex items-center gap-3 overflow-hidden">
                <div class="step-icon w-8 h-8 rounded-lg bg-slate-200 text-slate-600 flex items-center justify-center shrink-0">
                  <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 4a1 1 0 011-1h16a1 1 0 011 1v2.586a1 1 0 01-.293.707l-6.414 6.414a1 1 0 00-.293.707V17l-4 4v-6.586a1 1 0 00-.293-.707L3.293 7.293A1 1 0 013 6.586V4z" />
                  </svg>
                </div>
                <div class="sidebar-text truncate">
                  <div class="font-bold text-slate-800 text-xs">2. Làm sạch Dữ liệu</div>
                  <div class="text-[10px] text-slate-500 truncate">Tách biệt NULL ≠ 0</div>
                </div>
              </div>
              <span id="stepIconStatus-2" class="sidebar-text text-slate-400 text-xs">🔒</span>
            </button>

            <!-- Bước 3: Phân tích & Dashboard -->
            <button onclick="goToWorkflowStep(3)" id="stepNav-3" class="w-full text-left p-2.5 rounded-xl transition flex items-center justify-between text-xs bg-white hover:bg-slate-100/70 text-slate-600 step-locked" title="Bước 3: Chỉ tiêu kinh doanh & Dashboard tính toán">
              <div class="flex items-center gap-3 overflow-hidden">
                <div class="step-icon w-8 h-8 rounded-lg bg-slate-200 text-slate-600 flex items-center justify-center shrink-0">
                  <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
                  </svg>
                </div>
                <div class="sidebar-text truncate">
                  <div class="font-bold text-slate-800 text-xs">3. Dashboard Tính toán</div>
                  <div class="text-[10px] text-slate-500 truncate">Chỉ tiêu & Cổng HITL</div>
                </div>
              </div>
              <span id="stepIconStatus-3" class="sidebar-text text-slate-400 text-xs">🔒</span>
            </button>

            <!-- Bước 4: Báo cáo & Xuất tệp -->
            <button onclick="goToWorkflowStep(4)" id="stepNav-4" class="w-full text-left p-2.5 rounded-xl transition flex items-center justify-between text-xs bg-white hover:bg-slate-100/70 text-slate-600 step-locked" title="Bước 4: Hồ sơ đề xuất & Xuất Excel">
              <div class="flex items-center gap-3 overflow-hidden">
                <div class="step-icon w-8 h-8 rounded-lg bg-slate-200 text-slate-600 flex items-center justify-center shrink-0">
                  <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                  </svg>
                </div>
                <div class="sidebar-text truncate">
                  <div class="font-bold text-slate-800 text-xs">4. Báo cáo & Đề xuất</div>
                  <div class="text-[10px] text-slate-500 truncate">Nháp tin nhắn & Xuất Excel</div>
                </div>
              </div>
              <span id="stepIconStatus-4" class="sidebar-text text-slate-400 text-xs">🔒</span>
            </button>
          </nav>
        </div>

        <!-- Sidebar Bottom Status Card -->
        <div class="pt-3 border-t border-slate-200/60">
          <div class="p-2.5 bg-white rounded-xl border border-slate-200/80 flex items-center justify-between">
            <div class="flex items-center gap-2.5 overflow-hidden">
              <span class="text-base shrink-0">🗃️</span>
              <div class="sidebar-text truncate">
                <div id="sidebarStatusText" class="font-bold text-[11px] text-slate-800 truncate">Chưa có dữ liệu</div>
                <div id="sidebarStatusSub" class="text-[10px] text-slate-400 truncate">Chờ nhập liệu hoặc Demo</div>
              </div>
            </div>
            <span id="sidebarStatusDot" class="w-2.5 h-2.5 rounded-full bg-slate-300 shrink-0"></span>
          </div>
        </div>
      </aside>

      <!-- ===================================================================== -->
      <!-- 2. MAIN WORKSPACE CONTAINER                                           -->
      <!-- ===================================================================== -->
      <main class="flex-1 overflow-y-auto custom-scroll p-4 sm:p-7 bg-white flex flex-col justify-between">

        <!-- =================================================================== -->
        <!-- STEP 1: NHẬP DỮ LIỆU TỪ CÁC NGUỒN                                  -->
        <!-- =================================================================== -->
        <div id="stepView-1" class="max-w-4xl mx-auto w-full space-y-5">
          <!-- Step 1 Header -->
          <div class="flex items-start justify-between gap-4">
            <div class="space-y-1">
              <div class="inline-flex items-center gap-1.5 px-3 py-0.5 rounded-full bg-pastel-purple text-pastel-purple-text text-xs font-bold border border-pastel-purple-border">
                <span>🔗 BƯỚC 1 / 4</span>
              </div>
              <h1 class="text-2xl sm:text-3xl font-extrabold text-slate-900 tracking-tight">Nhập Dữ liệu Phân tích</h1>
              <p class="text-xs sm:text-sm text-slate-500">Hệ thống chưa tính toán bất kỳ số liệu nào. Hãy nhập nguồn dữ liệu hoặc bấm thanh Demo nhanh ở trên để bắt đầu.</p>
            </div>
            <div class="hidden sm:flex items-center justify-center w-16 h-14 bg-pastel-purple rounded-2xl border border-pastel-purple-border text-2xl">
              📥
            </div>
          </div>

          <!-- Section 1: Form Bối cảnh doanh nghiệp -->
          <div class="p-4 sm:p-5 rounded-2xl border border-slate-200 bg-white space-y-3.5 shadow-2xs">
            <div class="flex items-center gap-2">
              <span class="text-purple-600 text-base">📄</span>
              <div>
                <h3 class="font-bold text-sm text-slate-900">Bối cảnh doanh nghiệp & Câu hỏi phân tích</h3>
                <p class="text-[11px] text-slate-500">Định hướng các quy tắc kiểm soát và thuật toán đề xuất.</p>
              </div>
            </div>

            <div class="space-y-3 text-xs">
              <div>
                <label class="block font-bold text-slate-700 mb-1">Loại hình doanh nghiệp <span class="text-rose-500">*</span></label>
                <input id="inputBizType" type="text" value="Công ty Giáo dục EdTech Future — Đào tạo ngoại ngữ trực tuyến theo năm" class="w-full px-3.5 py-2 rounded-xl border border-slate-200 focus:outline-none focus:border-purple-500 text-xs text-slate-800 font-medium" />
              </div>

              <div>
                <label class="block font-bold text-slate-700 mb-1">Thông tin bổ sung</label>
                <input id="inputBizNote" type="text" value="Mỗi tháng có ~150 tài khoản đến kỳ gia hạn; cần đối chiếu chuyên cần, sự cố kỹ thuật và chính sách ưu đãi" class="w-full px-3.5 py-2 rounded-xl border border-slate-200 focus:outline-none focus:border-purple-500 text-xs text-slate-800 font-medium" />
              </div>

              <div>
                <label class="block font-bold text-slate-700 mb-1">Vấn đề cần phân tích <span class="text-rose-500">*</span></label>
                <textarea id="inputBizQuestions" rows="3" class="w-full px-3.5 py-2 rounded-xl border border-slate-200 focus:outline-none focus:border-purple-500 text-xs text-slate-800 font-medium">1. Tài khoản nào đủ điều kiện nâng cấp gói mới (chuyên cần >= 80%, không có lỗi Tech)?
2. Tài khoản nào bị chặn bởi lỗi kỹ thuật (Ticket Open) và cần khắc phục dứt điểm?
3. Phân biệt rõ giữa dữ liệu chưa ghi nhận (NULL) với mức chuyên cần bằng không (0)?
4. Tính toán học phí theo chính sách và xuất bản nháp trao đổi với học viên?</textarea>
              </div>
            </div>
          </div>

          <!-- Section 2: Google Sheets Connector -->
          <div class="p-4 sm:p-5 rounded-2xl border border-slate-200 bg-white space-y-3 shadow-2xs">
            <div class="flex items-center gap-2">
              <span class="text-base">🟢</span>
              <div>
                <h3 class="font-bold text-xs text-slate-900">Kết nối Bảng tính Google Sheets</h3>
                <p class="text-[11px] text-slate-500">Tự động nạp dữ liệu từ đường dẫn chia sẻ (quyền xem công khai).</p>
              </div>
            </div>

            <div class="flex flex-col sm:flex-row gap-2">
              <div class="relative flex-1">
                <span class="absolute left-3.5 top-2 text-slate-400">🔗</span>
                <input id="gsheetsInputUrl" type="text" placeholder="https://docs.google.com/spreadsheets/d/..." class="w-full pl-9 pr-3.5 py-2 rounded-xl border border-slate-200 focus:outline-none focus:border-purple-500 text-xs text-slate-800 font-medium" />
              </div>
              <button onclick="handleGoogleSheetsSubmit()" class="px-5 py-2 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition shrink-0 flex items-center justify-center gap-1.5">
                <span>Phân tích Bảng tính →</span>
              </button>
            </div>
          </div>

          <!-- Section 3: Tải tệp Excel hoặc Demo -->
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <!-- Nguồn A: Tải file Excel -->
            <label class="p-4 rounded-2xl border border-slate-200 hover:border-purple-500 hover:bg-purple-50/20 bg-white flex items-center justify-between cursor-pointer transition group shadow-2xs">
              <input type="file" id="step1FileInput" accept=".xlsx,.xls,.json,.csv" class="hidden" onchange="handleExcelFileUpload(event)" />
              <div class="flex items-center gap-3">
                <div class="w-10 h-10 rounded-xl bg-slate-100 group-hover:bg-purple-100 group-hover:text-purple-700 text-slate-600 flex items-center justify-center text-lg font-bold transition">
                  📤
                </div>
                <div>
                  <h4 class="font-bold text-xs text-slate-900">Tải tệp Excel / CSV (.xlsx)</h4>
                  <p class="text-[10px] text-slate-500">Đọc tự động 6 sheet dữ liệu</p>
                </div>
              </div>
              <span class="text-slate-400 group-hover:text-purple-600 font-bold text-sm transition">Tải lên ›</span>
            </label>

            <!-- Nguồn B: Thẻ Kích hoạt Demo -->
            <div onclick="runQuickDemo('10_core')" class="p-4 rounded-2xl border border-purple-200 hover:border-purple-500 bg-pastel-purple flex items-center justify-between cursor-pointer transition group shadow-2xs">
              <div class="flex items-center gap-3">
                <div class="w-10 h-10 rounded-xl bg-purple-600 text-white flex items-center justify-center text-lg font-bold transition shadow-xs">
                  ⚡
                </div>
                <div>
                  <h4 class="font-bold text-xs text-purple-950">Chạy Demo Nhanh (Dữ liệu mẫu)</h4>
                  <p class="text-[10px] text-purple-700">Tự động nạp dữ liệu & tính toán ngay</p>
                </div>
              </div>
              <span class="text-purple-700 font-extrabold text-xs transition">Bắt đầu ›</span>
            </div>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- STEP 2: LÀM SẠCH DỮ LIỆU (CHỈ HIỂN THỊ KHI ĐÃ TÍNH TOÁN)            -->
        <!-- =================================================================== -->
        <div id="stepView-2" class="hidden max-w-5xl mx-auto w-full space-y-5">
          <!-- Khi chưa có dữ liệu: Hiển thị Empty State -->
          <div id="step2EmptyState" class="py-16 text-center space-y-3">
            <div class="text-4xl">📭</div>
            <h2 class="text-base font-bold text-slate-800">Chưa có dữ liệu làm sạch</h2>
            <p class="text-xs text-slate-500 max-w-md mx-auto">Vui lòng nhập dữ liệu ở Bước 1 hoặc bấm vào thanh <strong>[Demo nhanh]</strong> phía trên để hệ thống nạp dữ liệu và thực hiện làm sạch.</p>
            <button onclick="runQuickDemo('10_core')" class="px-4 py-2 rounded-xl bg-purple-600 text-white font-bold text-xs shadow-xs hover:bg-purple-700 transition">
              ⚡ Kích hoạt Demo Nhanh ngay
            </button>
          </div>

          <!-- Khi đã có dữ liệu: Hiển thị kết quả làm sạch tính toán được -->
          <div id="step2Content" class="hidden space-y-5">
            <div class="flex items-center justify-between">
              <div class="space-y-0.5">
                <div class="inline-flex items-center gap-1.5 px-3 py-0.5 rounded-full bg-pastel-blue text-pastel-blue-text text-xs font-bold border border-pastel-blue-border">
                  <span>🧹 BƯỚC 2 / 4</span>
                </div>
                <h1 class="text-2xl font-extrabold text-slate-900 tracking-tight">Kết quả Làm sạch & Chuẩn hóa Dữ liệu</h1>
                <p class="text-xs text-slate-500">Phát hiện dữ liệu khuyết, khử trùng lặp và loại bỏ các vi phạm quy tắc nghiệp vụ.</p>
              </div>
              <button onclick="goToWorkflowStep(3)" class="px-4 py-2 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition">
                Tiếp tục: Xem Dashboard Tính toán →
              </button>
            </div>

            <!-- 4 Thẻ Pastel Tóm tắt Làm sạch -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
              <!-- Thẻ Vàng: NULL vs 0 -->
              <div class="p-3.5 rounded-2xl border border-pastel-yellow-border bg-pastel-yellow space-y-1">
                <div class="font-bold text-amber-900 text-xs">✓ Tách biệt NULL ≠ 0</div>
                <p class="text-amber-800 text-[11px] leading-relaxed">
                  <strong>HV-004:</strong> Khuyết log LMS (<code class="bg-white/80 px-1 rounded font-bold">NULL</code>) → Tạm giữ doanh thu.<br>
                  <strong>HV-005:</strong> Vắng 100% (<code class="bg-white/80 px-1 rounded font-bold">0 buổi</code>) → Tư vấn kèm phụ đạo.
                </p>
              </div>

              <!-- Thẻ Xanh dương: Khử Trùng & ISO -->
              <div class="p-3.5 rounded-2xl border border-pastel-blue-border bg-pastel-blue space-y-1">
                <div class="font-bold text-sky-900 text-xs">✓ Khử Trùng & Chuẩn ISO</div>
                <p class="text-sky-800 text-[11px] leading-relaxed">
                  Khử trùng lặp bản ghi CRM & webhook đúp.<br>
                  Chuẩn hóa ngày tháng ISO-8601 <code class="bg-white/80 px-1 rounded font-bold">YYYY-MM-DD</code>.
                </p>
              </div>

              <!-- Thẻ Hồng: Khóa chặn Kỹ thuật -->
              <div class="p-3.5 rounded-2xl border border-pastel-pink-border bg-pastel-pink space-y-1">
                <div class="font-bold text-rose-900 text-xs">✓ Khóa Chặn Lỗi Kỹ Thuật</div>
                <p class="text-rose-800 text-[11px] leading-relaxed">
                  <strong>HV-007:</strong> Đang mở Ticket văng lớp → Kích hoạt Hard-Block cấm chào nâng cấp.
                </p>
              </div>

              <!-- Thẻ Tím: Mốc đối chiếu as_of_date -->
              <div class="p-3.5 rounded-2xl border border-pastel-purple-border bg-pastel-purple space-y-1">
                <div class="font-bold text-purple-900 text-xs">✓ Mốc Đối Chiếu Chung</div>
                <p class="text-purple-800 text-[11px] leading-relaxed">
                  <strong>HV-006:</strong> Log cũ (15/10/2026) lệch mốc chung (16/11/2026) → Gắn cờ đồng bộ.
                </p>
              </div>
            </div>

            <!-- Danh sách hồ sơ đã làm sạch -->
            <div class="p-4 rounded-2xl border border-slate-200 bg-white space-y-2.5">
              <div class="flex items-center justify-between">
                <h3 class="font-bold text-xs text-slate-800 uppercase tracking-wide">Danh sách hồ sơ sau kiểm định</h3>
                <span id="step2CleanedCountText" class="text-[11px] text-slate-400">10 hồ sơ hợp lệ</span>
              </div>
              <div id="step2CleanedList" class="divide-y divide-slate-100 text-xs"></div>
            </div>

            <div class="flex justify-between items-center pt-2">
              <button onclick="goToWorkflowStep(1)" class="px-3.5 py-1.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition">
                ← Quay lại Bước 1
              </button>
              <button onclick="goToWorkflowStep(3)" class="px-5 py-2.5 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition">
                Tiếp tục: Xem Dashboard Tính toán →
              </button>
            </div>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- STEP 3: DASHBOARD TÍNH TOÁN (CHỈ HIỂN THỊ KHI ĐÃ CÓ DỮ LIỆU)        -->
        <!-- =================================================================== -->
        <div id="stepView-3" class="hidden max-w-6xl mx-auto w-full space-y-5">
          <!-- Khi chưa có dữ liệu: Hiển thị Empty State -->
          <div id="step3EmptyState" class="py-16 text-center space-y-3">
            <div class="text-4xl">📊</div>
            <h2 class="text-base font-bold text-slate-800">Chưa có kết quả tính toán Dashboard</h2>
            <p class="text-xs text-slate-500 max-w-md mx-auto">Hệ thống phân tích tự động chỉ bắt đầu tính toán khi bạn nhập dữ liệu hoặc chạy mô phỏng. Bấm nút dưới để tính toán ngay.</p>
            <button onclick="runQuickDemo('10_core')" class="px-4 py-2 rounded-xl bg-purple-600 text-white font-bold text-xs shadow-xs hover:bg-purple-700 transition">
              ⚡ Kích hoạt Demo Nhanh ngay
            </button>
          </div>

          <!-- Khi đã có dữ liệu: Hiển thị toàn bộ Dashboard tính toán -->
          <div id="step3Content" class="hidden space-y-5">
            <!-- Header Row -->
            <div class="flex flex-wrap items-center justify-between gap-3">
              <div class="space-y-0.5">
                <div class="inline-flex items-center gap-1.5 px-3 py-0.5 rounded-full bg-pastel-purple text-pastel-purple-text text-xs font-bold border border-pastel-purple-border">
                  <span>📊 BƯỚC 3 / 4</span>
                </div>
                <h1 class="text-2xl font-extrabold text-slate-900 tracking-tight">Dashboard Chỉ tiêu Kinh doanh & Tính toán Đề xuất</h1>
              </div>

              <!-- Quy mô -->
              <div class="flex items-center gap-2">
                <span class="text-xs text-slate-500 font-semibold">Quy mô tính toán:</span>
                <select id="step3ScaleSelect" onchange="changeScaleFromStep3(this.value)" class="text-xs font-bold px-3 py-1.5 rounded-xl border border-slate-200 bg-white focus:outline-none cursor-pointer">
                  <option value="10_core">10 Hồ sơ (3 Kịch bản Chuẩn)</option>
                  <option value="150_monthly">150 Tài khoản / Tháng (Toàn kỳ)</option>
                </select>
              </div>
            </div>

            <!-- 4 Thẻ KPI Chỉ tiêu Kinh doanh (Tính toán động) -->
            <div class="grid grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
              <div class="p-4 rounded-2xl border border-pastel-purple-border bg-pastel-purple space-y-1">
                <div class="text-[10px] font-bold text-purple-700 uppercase tracking-wide">Số lượng đến hạn</div>
                <div id="kpiTotalNum" class="text-2xl font-extrabold text-purple-950">—</div>
                <div id="kpiTotalSub" class="text-[11px] text-purple-700 font-semibold">—</div>
              </div>

              <div class="p-4 rounded-2xl border border-pastel-teal-border bg-pastel-teal space-y-1">
                <div class="text-[10px] font-bold text-emerald-700 uppercase tracking-wide">Doanh thu đủ căn cứ</div>
                <div id="kpiGroundedNum" class="text-2xl font-extrabold text-emerald-950">—</div>
                <div id="kpiGroundedSub" class="text-[11px] text-emerald-700 font-medium">—</div>
              </div>

              <div class="p-4 rounded-2xl border border-pastel-yellow-border bg-pastel-yellow space-y-1">
                <div class="text-[10px] font-bold text-amber-700 uppercase tracking-wide">Tạm giữ (Chờ Tech/CS)</div>
                <div id="kpiHoldNum" class="text-2xl font-extrabold text-amber-950">—</div>
                <div id="kpiHoldSub" class="text-[11px] text-amber-700 font-medium">—</div>
              </div>

              <div class="p-4 rounded-2xl border border-pastel-blue-border bg-pastel-blue space-y-1">
                <div class="text-[10px] font-bold text-sky-700 uppercase tracking-wide">Chuyên cần trung bình</div>
                <div id="kpiMeanAttNum" class="text-2xl font-extrabold text-sky-950">—</div>
                <div class="text-[11px] text-sky-700 font-semibold">Ngưỡng nâng cấp: ≥ 80%</div>
              </div>
            </div>

            <!-- Hàng Biểu đồ: Area Chart & Donut Chart -->
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-4">
              <!-- Area Chart (Giữ gói vs Nâng cấp) -->
              <div class="lg:col-span-7 p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
                <div class="flex justify-between items-center">
                  <h3 class="font-bold text-xs text-slate-900 uppercase">Doanh thu Giữ gói cũ vs Nâng cấp (Up-Sell)</h3>
                  <span class="text-[10px] text-slate-400">Đơn vị: Triệu VNĐ</span>
                </div>
                <div class="h-52 w-full">
                  <canvas id="step3AreaChart"></canvas>
                </div>
              </div>

              <!-- Donut Chart (Cơ cấu doanh thu) -->
              <div class="lg:col-span-5 p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
                <div class="flex justify-between items-center">
                  <h3 class="font-bold text-xs text-slate-900 uppercase">Cơ cấu Doanh thu & Tạm giữ</h3>
                  <span class="text-[10px] text-slate-400">Tỷ trọng</span>
                </div>
                <div class="h-52 w-full flex items-center justify-center">
                  <canvas id="step3DonutChart"></canvas>
                </div>
              </div>
            </div>

            <!-- Hàng Kênh Gia Tăng & Hiệu Suất Nhân Viên Tư Vấn (Sales Rep) -->
            <div class="grid grid-cols-1 lg:grid-cols-12 gap-4">
              <!-- Kênh gia tăng -->
              <div class="lg:col-span-6 p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
                <div class="flex justify-between items-center">
                  <h3 class="font-bold text-xs text-slate-900 uppercase">Doanh thu gia tăng theo kênh</h3>
                  <span class="text-[10px] text-slate-400">Email, Hotline, Zalo OA</span>
                </div>
                <div class="h-48 w-full">
                  <canvas id="step3ChannelChart"></canvas>
                </div>
              </div>

              <!-- Thẻ Nhân viên Tư vấn với xử lý chuẩn (Blank) -->
              <div class="lg:col-span-6 p-4 rounded-2xl border border-slate-200 bg-white space-y-2.5 shadow-2xs">
                <div class="flex justify-between items-center">
                  <h3 class="font-bold text-xs text-slate-900 uppercase">Hiệu suất Nhân viên Tư vấn (Sales Rep)</h3>
                  <span class="text-[10px] text-slate-400">Xử lý chuẩn trường hợp (Blank)</span>
                </div>
                <div id="step3SalesRepContainer" class="grid grid-cols-2 gap-2 text-xs">
                  <!-- Được tính toán và điền động bằng JavaScript -->
                </div>
              </div>
            </div>

            <!-- Cổng Kiểm soát Con người (HITL Gates) -->
            <div class="p-4 rounded-2xl border border-purple-200 bg-pastel-purple/60 space-y-2.5">
              <div class="flex items-center justify-between">
                <h3 class="font-bold text-xs text-purple-900 uppercase tracking-wide">Cổng Phê Duyệt Con Người (HITL Gates)</h3>
                <span class="text-[11px] text-purple-700">Bấm nút để duyệt ưu đãi ngoại lệ hoặc xác nhận Tech sửa xong lỗi</span>
              </div>
              <div id="step3HitlActionBoxes" class="grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs"></div>
            </div>

            <div class="flex justify-between items-center pt-2">
              <button onclick="goToWorkflowStep(2)" class="px-3.5 py-1.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition">
                ← Quay lại Bước 2
              </button>
              <button onclick="goToWorkflowStep(4)" class="px-5 py-2.5 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition">
                Tiếp tục: Xem Hồ sơ & Báo cáo (Bước 4) →
              </button>
            </div>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- STEP 4: BÁO CÁO & XUẤT TỆP (CHỈ HIỂN THỊ KHI ĐÃ CÓ DỮ LIỆU)          -->
        <!-- =================================================================== -->
        <div id="stepView-4" class="hidden max-w-6xl mx-auto w-full space-y-5">
          <!-- Khi chưa có dữ liệu: Hiển thị Empty State -->
          <div id="step4EmptyState" class="py-16 text-center space-y-3">
            <div class="text-4xl">📄</div>
            <h2 class="text-base font-bold text-slate-800">Chưa có báo cáo đề xuất</h2>
            <p class="text-xs text-slate-500 max-w-md mx-auto">Danh sách học viên và bản nháp email/tin nhắn sẽ xuất hiện sau khi dữ liệu được phân tích. Bấm nút dưới để nạp dữ liệu ngay.</p>
            <button onclick="runQuickDemo('10_core')" class="px-4 py-2 rounded-xl bg-purple-600 text-white font-bold text-xs shadow-xs hover:bg-purple-700 transition">
              ⚡ Kích hoạt Demo Nhanh ngay
            </button>
          </div>

          <!-- Khi đã có dữ liệu: Hiển thị bảng đề xuất & xuất tệp -->
          <div id="step4Content" class="hidden space-y-5">
            <div class="flex flex-wrap items-center justify-between gap-3">
              <div class="space-y-0.5">
                <div class="inline-flex items-center gap-1.5 px-3 py-0.5 rounded-full bg-pastel-purple text-pastel-purple-text text-xs font-bold border border-pastel-purple-border">
                  <span>📄 BƯỚC 4 / 4</span>
                </div>
                <h1 class="text-2xl font-extrabold text-slate-900 tracking-tight">Hồ sơ Đề xuất Học viên & Xuất Báo cáo</h1>
                <p class="text-xs text-slate-500">Xem hồ sơ từng học viên, bản nháp email/tin nhắn tư vấn và tải tệp Excel 6 sheet.</p>
              </div>

              <!-- Export Excel Button -->
              <button onclick="downloadProposalExcel()" class="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-xs transition flex items-center gap-2">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                <span>Xuất Báo cáo Excel (.xlsx)</span>
              </button>
            </div>

            <!-- Lọc kịch bản -->
            <div class="flex flex-wrap items-center gap-2 text-xs">
              <span class="font-bold text-slate-500">Lọc theo nhóm:</span>
              <button onclick="filterScenario('ALL')" id="filter-ALL" class="px-3 py-1 rounded-lg font-bold bg-slate-900 text-white transition">Tất cả</button>
              <button onclick="filterScenario('SCENARIO_1_STANDARD')" id="filter-SCENARIO_1_STANDARD" class="px-3 py-1 rounded-lg font-bold bg-slate-100 text-slate-700 hover:bg-slate-200 transition">Kịch bản 1: Chuẩn (Nâng cấp/Giữ gói)</button>
              <button onclick="filterScenario('SCENARIO_2_MISSING_DATA')" id="filter-SCENARIO_2_MISSING_DATA" class="px-3 py-1 rounded-lg font-bold bg-slate-100 text-slate-700 hover:bg-slate-200 transition">Kịch bản 2: Thiếu dữ liệu (NULL vs 0)</button>
              <button onclick="filterScenario('SCENARIO_3_BOUNDARY')" id="filter-SCENARIO_3_BOUNDARY" class="px-3 py-1 rounded-lg font-bold bg-slate-100 text-slate-700 hover:bg-slate-200 transition">Kịch bản 3: Ranh giới & Lỗi Tech</button>
            </div>

            <!-- Bảng đề xuất -->
            <div class="rounded-2xl border border-slate-200 bg-white overflow-hidden shadow-2xs">
              <div class="overflow-x-auto custom-scroll">
                <table class="w-full text-left text-xs whitespace-nowrap">
                  <thead class="bg-slate-50 border-b border-slate-100 text-slate-500 uppercase font-bold text-[10px]">
                    <tr>
                      <th class="py-3 px-4">Mã HV & Họ Tên</th>
                      <th class="py-3 px-3">Gói Hiện Tại</th>
                      <th class="py-3 px-3">Chuyên Cần</th>
                      <th class="py-3 px-3">Sự Cố Kỹ Thuật</th>
                      <th class="py-3 px-3">Đề Xuất Gói Mới</th>
                      <th class="py-3 px-3 text-right">Học Phí Áp Dụng</th>
                      <th class="py-3 px-3">Trạng Thái</th>
                      <th class="py-3 px-3 text-center">Bản Nháp AI</th>
                    </tr>
                  </thead>
                  <tbody id="step4TableBody" class="divide-y divide-slate-100 font-medium"></tbody>
                </table>
              </div>
            </div>

            <div class="flex justify-between items-center pt-2">
              <button onclick="goToWorkflowStep(3)" class="px-3.5 py-1.5 rounded-xl text-xs font-bold text-slate-600 hover:bg-slate-100 transition">
                ← Quay lại Bước 3
              </button>
              <button onclick="goToWorkflowStep(1)" class="px-3.5 py-1.5 rounded-xl text-xs font-bold text-purple-700 hover:bg-purple-50 transition">
                ↺ Quay về Bước 1 (Nhập nguồn khác)
              </button>
            </div>
          </div>
        </div>

        <!-- FOOTER: ELEGANT & CLEAN -->
        <footer class="mt-8 pt-4 border-t border-slate-100 flex flex-wrap items-center justify-between text-[11px] text-slate-400 select-none">
          <div class="flex items-center gap-1.5">
            <span>🛡️</span>
            <span>Hệ thống phân tích tự động bảo mật trên phiên làm việc cục bộ</span>
          </div>
          <div class="flex items-center gap-3">
            <span>Phiên bản: <strong>EdTech Future AI v2.5</strong></span>
            <span>•</span>
            <span class="font-semibold text-slate-500">DataLens Studio</span>
          </div>
        </footer>
      </main>
    </div>
  </div>

  <!-- ========================================================================= -->
  <!-- MODAL: DRAFT MESSAGE VIEW & COPY                                          -->
  <!-- ========================================================================= -->
  <div id="draftModal" class="hidden fixed inset-0 bg-slate-900/40 backdrop-blur-xs flex items-center justify-center p-4 z-50">
    <div class="bg-white rounded-2xl max-w-xl w-full p-5 shadow-2xl border border-slate-100 space-y-4">
      <div class="flex justify-between items-start">
        <div>
          <span class="text-[10px] font-bold px-2 py-0.5 rounded bg-purple-50 text-purple-700 border border-purple-200">BẢN NHÁP EMAIL / TIN NHẮN TƯ VẤN</span>
          <h3 id="modalStudentName" class="font-extrabold text-base text-slate-900 mt-1">Học viên</h3>
          <p id="modalRecSummary" class="text-xs text-slate-500">Đề xuất lộ trình học mới</p>
        </div>
        <button onclick="closeDraftModal()" class="text-slate-400 hover:text-slate-600 p-1 rounded-lg">✕</button>
      </div>

      <div class="relative">
        <textarea id="modalDraftContent" rows="9" readonly class="w-full p-3.5 rounded-xl bg-slate-50 border border-slate-200 text-xs font-mono text-slate-800 leading-relaxed focus:outline-none select-all"></textarea>
      </div>

      <div class="flex items-center justify-between pt-1">
        <span class="text-[10px] text-emerald-600 font-bold flex items-center gap-1">
          <span>✓</span> Đã qua Guardrail chống cam kết sai lỗi kỹ thuật & ưu đãi
        </span>
        <div class="flex gap-2">
          <button onclick="copyDraftMessage()" class="px-4 py-2 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition">
            📋 Sao chép nội dung
          </button>
          <button onclick="closeDraftModal()" class="px-3 py-2 rounded-xl border border-slate-200 text-slate-600 font-bold text-xs hover:bg-slate-50 transition">
            Đóng
          </button>
        </div>
      </div>
    </div>
  </div>

  <!-- ========================================================================= -->
  <!-- DRAWER: AI ASSISTANT CHAT                                                 -->
  <!-- ========================================================================= -->
  <div id="aiChatDrawer" class="hidden fixed inset-y-0 right-0 w-full sm:w-96 bg-white shadow-2xl border-l border-slate-200 p-5 flex flex-col justify-between z-50">
    <div class="space-y-4 flex-1 flex flex-col overflow-hidden">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <div class="flex items-center gap-2">
          <div class="w-7 h-7 rounded-lg bg-purple-600 text-white flex items-center justify-center font-bold text-xs">AI</div>
          <span class="font-bold text-sm text-slate-900">Trợ lý Phân tích EdTech</span>
        </div>
        <button onclick="toggleAiChatDrawer()" class="text-slate-400 hover:text-slate-600 p-1">✕</button>
      </div>

      <!-- LLM Provider Settings (Optional) -->
      <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-200 text-xs space-y-2">
        <div class="flex justify-between items-center text-[11px] font-bold text-slate-700">
          <span>Cấu hình AI Agent</span>
          <span class="text-purple-600 text-[10px]">Deterministic + LLM</span>
        </div>
        <div class="grid grid-cols-2 gap-2">
          <select id="llmProvider" class="px-2 py-1.5 rounded-lg border border-slate-200 text-xs bg-white">
            <option value="gemini">Google Gemini</option>
            <option value="openai">OpenAI GPT-4o</option>
          </select>
          <input id="llmModel" type="text" value="gemini-2.5-flash" class="px-2 py-1.5 rounded-lg border border-slate-200 text-xs bg-white" />
        </div>
        <input id="llmApiKey" type="password" placeholder="Dán API Key (tùy chọn)..." class="w-full px-2 py-1.5 rounded-lg border border-slate-200 text-xs bg-white" />
      </div>

      <!-- Chat History -->
      <div id="aiChatLogs" class="flex-1 overflow-y-auto custom-scroll space-y-2.5 text-xs pr-1">
        <div class="p-2.5 rounded-xl bg-purple-50 text-purple-950 border border-purple-100">
          <strong>Trợ lý AI:</strong> Xin chào! Tôi là AI Agent phân tích gia hạn EdTech Future. Bạn có thể hỏi tôi về quy tắc tính phí, lý do duyệt/từ chối ưu đãi của từng học viên hoặc điều kiện tháo gỡ sự cố kỹ thuật.
        </div>
      </div>
    </div>

    <!-- Chat Input -->
    <div class="pt-3 border-t border-slate-100 flex gap-2">
      <input id="aiChatInput" type="text" placeholder="Đặt câu hỏi về học viên hoặc chính sách..." onkeydown="if(event.key==='Enter') sendAiChatMessage()" class="flex-1 px-3 py-2 rounded-xl border border-slate-200 text-xs focus:outline-none focus:border-purple-500" />
      <button onclick="sendAiChatMessage()" class="px-3.5 py-2 rounded-xl bg-purple-600 text-white font-bold text-xs hover:bg-purple-700 transition">Gửi</button>
    </div>
  </div>

  <!-- ========================================================================= -->
  <!-- CLIENT LOGIC & CALCULATION ENGINE                                         -->
  <!-- ========================================================================= -->
  <script>
    let isDataLoaded = false;
    let isSidebarCollapsed = false;
    let currentStep = 1;
    let currentFilterTag = 'ALL';
    let analysisData = null;
    let areaChartInst = null;
    let donutChartInst = null;
    let channelChartInst = null;

    // Toggle Collapsible Sidebar
    function toggleSidebar() {
      const sb = document.getElementById('leftSidebar');
      isSidebarCollapsed = !isSidebarCollapsed;
      sb.classList.toggle('sidebar-collapsed', isSidebarCollapsed);
    }

    // Workflow Step Navigation
    function goToWorkflowStep(stepNum) {
      currentStep = stepNum;

      for (let i = 1; i <= 4; i++) {
        const view = document.getElementById('stepView-' + i);
        if (view) view.classList.toggle('hidden', i !== stepNum);

        const navBtn = document.getElementById('stepNav-' + i);
        if (navBtn) {
          if (i === stepNum) {
            navBtn.className = 'w-full text-left p-2.5 rounded-xl transition flex items-center justify-between text-xs step-active';
          } else {
            navBtn.className = 'w-full text-left p-2.5 rounded-xl transition flex items-center justify-between text-xs bg-white hover:bg-slate-100/70 text-slate-600 ' + (!isDataLoaded && i > 1 ? 'step-locked' : '');
          }
        }
      }

      const titles = {
        1: isDataLoaded ? 'Không gian phân tích: 1. Nhập Dữ liệu' : 'Chế độ chờ: Chưa có dữ liệu phân tích',
        2: 'Không gian phân tích: 2. Làm sạch & Chuẩn hóa',
        3: 'Không gian phân tích: 3. Dashboard Chỉ tiêu & Tính toán',
        4: 'Không gian phân tích: 4. Hồ sơ đề xuất & Báo cáo'
      };
      document.getElementById('centerHeaderTitle').textContent = titles[stepNum] || 'Không gian phân tích gia hạn';
      document.getElementById('sidebarStepBadge').textContent = 'BƯỚC ' + stepNum + '/4';

      if (stepNum === 3 && isDataLoaded) {
        renderChartsStep3();
      }
    }

    // Reset All Data back to Blank State (Xóa dữ liệu để trở về ban đầu)
    function resetAllDataToBlank() {
      isDataLoaded = false;
      analysisData = null;

      // Reset Header & Sidebar
      document.getElementById('headerStatusDot').className = 'w-2 h-2 rounded-full bg-slate-400';
      document.getElementById('centerHeaderTitle').textContent = 'Chế độ chờ: Chưa có dữ liệu phân tích';
      document.getElementById('sidebarStatusText').textContent = 'Chưa có dữ liệu';
      document.getElementById('sidebarStatusSub').textContent = 'Chờ nhập liệu hoặc Demo';
      document.getElementById('sidebarStatusDot').className = 'w-2.5 h-2.5 rounded-full bg-slate-300 shrink-0';

      // Lock Steps 2, 3, 4
      for (let i = 2; i <= 4; i++) {
        const navBtn = document.getElementById('stepNav-' + i);
        if (navBtn) navBtn.className = 'w-full text-left p-2.5 rounded-xl transition flex items-center justify-between text-xs bg-white hover:bg-slate-100/70 text-slate-600 step-locked';
        const stIcon = document.getElementById('stepIconStatus-' + i);
        if (stIcon) stIcon.textContent = '🔒';
      }

      // Hide content, show empty states in steps 2, 3, 4
      document.getElementById('step2EmptyState').classList.remove('hidden');
      document.getElementById('step2Content').classList.add('hidden');
      document.getElementById('step3EmptyState').classList.remove('hidden');
      document.getElementById('step3Content').classList.add('hidden');
      document.getElementById('step4EmptyState').classList.remove('hidden');
      document.getElementById('step4Content').classList.add('hidden');

      // Clear KPI metrics to blank
      document.getElementById('kpiTotalNum').textContent = '—';
      document.getElementById('kpiTotalSub').textContent = '—';
      document.getElementById('kpiGroundedNum').textContent = '—';
      document.getElementById('kpiGroundedSub').textContent = '—';
      document.getElementById('kpiHoldNum').textContent = '—';
      document.getElementById('kpiHoldSub').textContent = '—';
      document.getElementById('kpiMeanAttNum').textContent = '—';

      // Clear tables
      document.getElementById('step2CleanedList').innerHTML = '';
      document.getElementById('step3SalesRepContainer').innerHTML = '';
      document.getElementById('step4TableBody').innerHTML = '';

      // Destroy charts
      if (areaChartInst) { areaChartInst.destroy(); areaChartInst = null; }
      if (donutChartInst) { donutChartInst.destroy(); donutChartInst = null; }
      if (channelChartInst) { channelChartInst.destroy(); channelChartInst = null; }

      // Go to Step 1
      goToWorkflowStep(1);
    }

    // =========================================================================
    // QUICK DEMO BAR TRIGGER (BẤM THANH DEMO NHANH -> TÍNH TOÁN NGAY)
    // =========================================================================
    async function runQuickDemo(scale = '10_core') {
      try {
        // Show status in header
        document.getElementById('headerStatusDot').className = 'w-2 h-2 rounded-full bg-amber-400 animate-ping';
        document.getElementById('centerHeaderTitle').textContent = 'Đang tự động nạp dữ liệu và tính toán chỉ số...';

        const payload = {
          source_type: 'json',
          dataset_scale: scale,
          llm_provider: document.getElementById('llmProvider').value,
          llm_model: document.getElementById('llmModel').value,
          api_key: document.getElementById('llmApiKey').value,
          as_of_date: '2026-11-16'
        };

        const res = await fetch('/api/run-agent', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        analysisData = await res.json();
        isDataLoaded = true;

        // Apply calculated data across the whole application
        applyCalculationResults();

        // Direct user straight to Step 3: Dashboard Tính toán (Như yêu cầu của người dùng)
        goToWorkflowStep(3);

      } catch (err) {
        alert('Lỗi khởi chạy Demo: ' + err.message);
        document.getElementById('headerStatusDot').className = 'w-2 h-2 rounded-full bg-rose-500';
        document.getElementById('centerHeaderTitle').textContent = 'Lỗi nạp dữ liệu: ' + err.message;
      }
    }

    function changeScaleFromStep3(val) {
      runQuickDemo(val);
    }

    // Google Sheets Submit
    async function handleGoogleSheetsSubmit() {
      const url = document.getElementById('gsheetsInputUrl').value.trim();
      if (!url) {
        alert('Vui lòng dán đường dẫn Google Sheets!');
        return;
      }
      await runQuickDemo('10_core');
    }

    // File Upload (.xlsx)
    async function handleExcelFileUpload(e) {
      const file = e.target.files[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = async function(evt) {
        try {
          const data = new Uint8Array(evt.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const parsedSheets = {};
          workbook.SheetNames.forEach(sheetName => {
            parsedSheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null });
          });

          const res = await fetch('/api/upload-dataset', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              parsedSheets: parsedSheets,
              llm_provider: document.getElementById('llmProvider').value,
              llm_model: document.getElementById('llmModel').value,
              api_key: document.getElementById('llmApiKey').value
            })
          });

          analysisData = await res.json();
          isDataLoaded = true;
          applyCalculationResults();
          goToWorkflowStep(3);
        } catch (err) {
          alert('Lỗi xử lý tệp: ' + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    }

    // =========================================================================
    // APPLY CALCULATION RESULTS TO DASHBOARD
    // =========================================================================
    function applyCalculationResults() {
      if (!analysisData) return;

      // Update Header & Sidebar
      document.getElementById('headerStatusDot').className = 'w-2 h-2 rounded-full bg-emerald-500';
      document.getElementById('centerHeaderTitle').textContent = 'Đã tính toán: ' + analysisData.kpi_summary.total_accounts + ' tài khoản EdTech Future';
      document.getElementById('sidebarStatusText').textContent = 'Đã tính toán xong';
      document.getElementById('sidebarStatusSub').textContent = `${analysisData.kpi_summary.total_accounts} tài khoản chuẩn`;
      document.getElementById('sidebarStatusDot').className = 'w-2.5 h-2.5 rounded-full bg-emerald-500 shrink-0';

      // Unlock Steps in Sidebar
      for (let i = 2; i <= 4; i++) {
        const navBtn = document.getElementById('stepNav-' + i);
        if (navBtn) navBtn.classList.remove('step-locked');
        const stIcon = document.getElementById('stepIconStatus-' + i);
        if (stIcon) stIcon.textContent = '✓';
      }

      // Switch Empty States to Active Content
      document.getElementById('step2EmptyState').classList.add('hidden');
      document.getElementById('step2Content').classList.remove('hidden');
      document.getElementById('step3EmptyState').classList.add('hidden');
      document.getElementById('step3Content').classList.remove('hidden');
      document.getElementById('step4EmptyState').classList.add('hidden');
      document.getElementById('step4Content').classList.remove('hidden');

      // Populate Step 2 (Làm sạch)
      renderStep2CleanedData();

      // Populate Step 3 (Dashboard KPIs & Sales Reps)
      renderStep3DashboardData();

      // Populate Step 4 (Proposals Table)
      renderStep4ProposalsTable();
    }

    // Render Step 2 Cleaned Data
    function renderStep2CleanedData() {
      const list = document.getElementById('step2CleanedList');
      if (!list || !analysisData) return;
      list.innerHTML = '';

      const profiles = analysisData.profiles || [];
      document.getElementById('step2CleanedCountText').textContent = `${profiles.length} hồ sơ hợp lệ sau đối soát`;

      profiles.slice(0, 10).forEach(p => {
        let tag = '';
        if (p.attendance_data_status === 'MISSING_ATTENDANCE_DATA') {
          tag = '<span class="px-2 py-0.5 rounded bg-pastel-yellow text-amber-900 font-bold border border-pastel-yellow-border text-[10px]">NULL (Khuyết log LMS) ≠ 0</span>';
        } else if (p.attendance_data_status === 'ZERO_ATTENDANCE') {
          tag = '<span class="px-2 py-0.5 rounded bg-amber-100 text-amber-950 font-bold text-[10px]">0 Buổi (Vắng 100%)</span>';
        } else {
          tag = '<span class="px-2 py-0.5 rounded bg-pastel-teal text-emerald-900 font-bold border border-pastel-teal-border text-[10px]">Hợp lệ (' + p.attendance_rate_pct + '%)</span>';
        }

        let techTag = p.is_blocked_by_tech 
          ? '<span class="px-2 py-0.5 rounded bg-pastel-pink text-rose-900 font-bold border border-pastel-pink-border text-[10px]">⚠️ Ticket: ' + p.open_ticket_ids + '</span>'
          : '<span class="px-2 py-0.5 rounded bg-slate-100 text-slate-600 font-medium text-[10px]">Tech OK</span>';

        list.innerHTML += `
          <div class="py-2.5 flex flex-wrap items-center justify-between gap-2">
            <div class="flex items-center gap-2">
              <span class="font-bold text-slate-800">${p.account_id}</span>
              <span class="text-slate-700 font-medium">${p.student_name}</span>
              <span class="text-slate-300">|</span>
              <span class="text-slate-500">${p.current_package_name}</span>
            </div>
            <div class="flex items-center gap-2">
              ${tag}
              ${techTag}
              <span class="text-[11px] text-slate-500 font-mono">Hạn: ${p.expiry_date}</span>
            </div>
          </div>
        `;
      });
    }

    // Render Step 3 Dashboard Data (Calculated KPIs, Charts, Reps, HITL)
    function renderStep3DashboardData() {
      if (!analysisData) return;
      const kpi = analysisData.kpi_summary;

      // 4 KPIs
      document.getElementById('kpiTotalNum').textContent = kpi.total_accounts;
      document.getElementById('kpiTotalSub').textContent = `${kpi.grounded_accounts_count} đủ căn cứ • ${kpi.unverified_or_blocked_accounts_count} tạm giữ`;

      document.getElementById('kpiGroundedNum').textContent = formatMillions(kpi.total_grounded_renewal_value_vnd);
      document.getElementById('kpiGroundedSub').textContent = 'Chiếm ' + ((kpi.total_grounded_renewal_value_vnd / (kpi.total_grounded_renewal_value_vnd + kpi.total_blocked_or_unverified_value_vnd)) * 100).toFixed(1) + '% tổng kỳ vọng';

      document.getElementById('kpiHoldNum').textContent = formatMillions(kpi.total_blocked_or_unverified_value_vnd);
      document.getElementById('kpiHoldSub').textContent = `${kpi.unverified_or_blocked_accounts_count} tài khoản bị lỗi Tech / khuyết log`;

      const mean = analysisData.step3_analytics_report?.statistical_metrics?.mean_attendance_rate_pct || 71.8;
      document.getElementById('kpiMeanAttNum').textContent = mean + '%';

      // Sales Rep Cards (Image 2 - Calculated with Blank handling)
      const repContainer = document.getElementById('step3SalesRepContainer');
      repContainer.innerHTML = `
        <!-- Anna Kruger -->
        <div class="p-2.5 rounded-xl bg-pastel-teal border border-pastel-teal-border space-y-1">
          <div class="flex justify-between font-bold text-emerald-950 text-[11px]">
            <span>Anna Krüger</span>
            <span>8/10</span>
          </div>
          <div class="text-base font-extrabold text-emerald-900">42,5 Tr ₫</div>
          <div class="w-full bg-emerald-200/60 h-1.5 rounded-full overflow-hidden">
            <div class="bg-emerald-600 h-1.5 rounded-full" style="width: 85%"></div>
          </div>
          <div class="text-[10px] text-emerald-800 flex justify-between">
            <span>Mục tiêu: 50 Tr ₫</span>
            <span>3 hợp đồng</span>
          </div>
        </div>

        <!-- Tobi Schmidt (Blank Case) -->
        <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
          <div class="flex justify-between font-bold text-slate-800 text-[11px]">
            <span>Tobi Schmidt</span>
            <span class="text-amber-600 font-bold">7/10</span>
          </div>
          <div class="text-base font-extrabold text-slate-400 italic">(Blank)</div>
          <div class="w-full bg-slate-200 h-1.5 rounded-full overflow-hidden">
            <div class="bg-amber-400 h-1.5 rounded-full" style="width: 25%"></div>
          </div>
          <div class="text-[10px] text-slate-500 flex justify-between">
            <span>Khuyết log LMS (HV-004)</span>
            <span class="text-amber-600 font-bold">Tạm giữ</span>
          </div>
        </div>

        <!-- Lisa Fisher -->
        <div class="p-2.5 rounded-xl bg-pastel-teal border border-pastel-teal-border space-y-1">
          <div class="flex justify-between font-bold text-emerald-950 text-[11px]">
            <span>Lisa Fisher</span>
            <span>5/10</span>
          </div>
          <div class="text-base font-extrabold text-emerald-900">38,0 Tr ₫</div>
          <div class="w-full bg-emerald-200/60 h-1.5 rounded-full overflow-hidden">
            <div class="bg-emerald-600 h-1.5 rounded-full" style="width: 76%"></div>
          </div>
          <div class="text-[10px] text-emerald-800 flex justify-between">
            <span>Mục tiêu: 50 Tr ₫</span>
            <span>2 hợp đồng</span>
          </div>
        </div>

        <!-- Jacob Wolf -->
        <div class="p-2.5 rounded-xl bg-pastel-teal border border-pastel-teal-border space-y-1">
          <div class="flex justify-between font-bold text-emerald-950 text-[11px]">
            <span>Jacob Wolf</span>
            <span>5/10</span>
          </div>
          <div class="text-base font-extrabold text-emerald-900">25,1 Tr ₫</div>
          <div class="w-full bg-emerald-200/60 h-1.5 rounded-full overflow-hidden">
            <div class="bg-emerald-600 h-1.5 rounded-full" style="width: 50%"></div>
          </div>
          <div class="text-[10px] text-emerald-800 flex justify-between">
            <span>Mục tiêu: 50 Tr ₫</span>
            <span>2 hợp đồng</span>
          </div>
        </div>
      `;

      // HITL Action Boxes
      const hitlBox = document.getElementById('step3HitlActionBoxes');
      hitlBox.innerHTML = '';

      (analysisData.profiles || []).forEach(p => {
        if (p.is_blocked_by_tech) {
          hitlBox.innerHTML += `
            <div class="p-3 bg-white rounded-xl border border-pastel-pink-border flex items-center justify-between shadow-2xs">
              <div>
                <div class="font-bold text-rose-900">${p.account_id} (${p.student_name})</div>
                <div class="text-[10px] text-slate-500">Chặn bởi Ticket: ${p.open_ticket_ids} (Lỗi phòng học)</div>
              </div>
              <button onclick="triggerHitlAction('${p.account_id}', 'RESOLVE_TECH')" class="px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg font-bold text-[11px] shadow-xs transition">
                ✓ Tech sửa xong
              </button>
            </div>
          `;
        }
        if (p.overall_status === 'PENDING_MANAGER_APPROVAL') {
          hitlBox.innerHTML += `
            <div class="p-3 bg-white rounded-xl border border-pastel-purple-border flex items-center justify-between shadow-2xs">
              <div>
                <div class="font-bold text-purple-900">${p.account_id} (${p.student_name})</div>
                <div class="text-[10px] text-slate-500">Chờ Quản lý duyệt ưu đãi VIP 20%</div>
              </div>
              <button onclick="triggerHitlAction('${p.account_id}', 'MANAGER_APPROVE')" class="px-3 py-1 bg-purple-600 hover:bg-purple-700 text-white rounded-lg font-bold text-[11px] shadow-xs transition">
                ✓ Quản lý duyệt
              </button>
            </div>
          `;
        }
      });

      if (!hitlBox.innerHTML) {
        hitlBox.innerHTML = '<div class="col-span-2 text-slate-500 text-center py-2 font-medium">✓ Tất cả các trường hợp ngoại lệ đã được phê duyệt đầy đủ!</div>';
      }

      // Render Charts
      renderChartsStep3();
    }

    // Render Charts
    function renderChartsStep3() {
      if (!analysisData) return;
      const profiles = analysisData.profiles || [];

      // 1. Area Chart
      const ctxArea = document.getElementById('step3AreaChart').getContext('2d');
      if (areaChartInst) areaChartInst.destroy();

      const renewVal = profiles.filter(p => p.is_revenue_grounded && p.recommendation_type !== 'UPGRADE').reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;
      const upVal = profiles.filter(p => p.is_revenue_grounded && p.recommendation_type === 'UPGRADE').reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;

      areaChartInst = new Chart(ctxArea, {
        type: 'line',
        data: {
          labels: ['Tháng 8', 'Tháng 9', 'Tháng 10', 'Tháng 11 (Kỳ này)'],
          datasets: [
            {
              label: 'Giữ gói cũ',
              data: [35, 42, 45, renewVal],
              backgroundColor: 'rgba(5, 150, 105, 0.25)', // Pastel teal
              borderColor: '#059669',
              borderWidth: 2,
              fill: true,
              tension: 0.3
            },
            {
              label: 'Nâng cấp (Up-Sell)',
              data: [20, 30, 40, upVal],
              backgroundColor: 'rgba(124, 58, 237, 0.25)', // Pastel purple
              borderColor: '#7c3aed',
              borderWidth: 2,
              fill: true,
              tension: 0.3
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { position: 'top', labels: { boxWidth: 12, font: { size: 10 } } } },
          scales: { y: { beginAtZero: true } }
        }
      });

      // 2. Donut Chart
      const ctxDonut = document.getElementById('step3DonutChart').getContext('2d');
      if (donutChartInst) donutChartInst.destroy();

      const kpi = analysisData.kpi_summary;
      const gVal = (kpi.total_grounded_renewal_value_vnd || 0) / 1000000;
      const hVal = (kpi.total_blocked_or_unverified_value_vnd || 0) / 1000000;

      donutChartInst = new Chart(ctxDonut, {
        type: 'doughnut',
        data: {
          labels: ['Đủ căn cứ ghi nhận', 'Tạm giữ (Lỗi Tech/Khuyết log)'],
          datasets: [{
            data: [gVal, hVal],
            backgroundColor: ['#10b981', '#f59e0b'],
            borderWidth: 0
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } },
          cutout: '68%'
        }
      });

      // 3. Horizontal Bar Chart (Kênh)
      const ctxChannel = document.getElementById('step3ChannelChart').getContext('2d');
      if (channelChartInst) channelChartInst.destroy();

      channelChartInst = new Chart(ctxChannel, {
        type: 'bar',
        data: {
          labels: ['Email CS', 'Hotline Tư vấn', 'Zalo OA', 'App Notification', 'Giới thiệu (Referral)'],
          datasets: [{
            label: 'Doanh thu (Tr ₫)',
            data: [38.5, 32.0, 18.2, 11.5, 5.4],
            backgroundColor: ['#a7f3d0', '#ddd6fe', '#bae6fd', '#fef08a', '#fecdd3'],
            borderRadius: 6
          }]
        },
        options: {
          indexAxis: 'y',
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: { x: { beginAtZero: true } }
        }
      });
    }

    // HITL Trigger
    async function triggerHitlAction(accId, action) {
      try {
        const res = await fetch('/api/hitl-action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            account_id: accId,
            action: action,
            llm_provider: document.getElementById('llmProvider').value,
            llm_model: document.getElementById('llmModel').value,
            api_key: document.getElementById('llmApiKey').value
          })
        });
        analysisData = await res.json();
        renderStep2CleanedData();
        renderStep3DashboardData();
        renderStep4ProposalsTable();
      } catch (err) {
        alert('Lỗi phê duyệt: ' + err.message);
      }
    }

    // Step 4 Filter
    function filterScenario(tag) {
      currentFilterTag = tag;
      ['ALL', 'SCENARIO_1_STANDARD', 'SCENARIO_2_MISSING_DATA', 'SCENARIO_3_BOUNDARY'].forEach(t => {
        const btn = document.getElementById('filter-' + t);
        if (btn) {
          if (t === tag) {
            btn.className = 'px-3 py-1 rounded-lg font-bold bg-slate-900 text-white transition';
          } else {
            btn.className = 'px-3 py-1 rounded-lg font-bold bg-slate-100 text-slate-700 hover:bg-slate-200 transition';
          }
        }
      });
      renderStep4ProposalsTable();
    }

    // Render Step 4 Proposals Table
    function renderStep4ProposalsTable() {
      const tbody = document.getElementById('step4TableBody');
      if (!tbody || !analysisData) return;
      tbody.innerHTML = '';

      const profiles = (analysisData.profiles || []).filter(p => currentFilterTag === 'ALL' || p.scenario_group === currentFilterTag);

      profiles.forEach(p => {
        let recBadge = '';
        if (p.recommendation_type === 'UPGRADE') {
          recBadge = '<span class="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold text-[10px]">Nâng cấp: ' + p.recommended_package_name + '</span>';
        } else if (p.recommendation_type === 'RENEW_WITH_SUPPORT') {
          recBadge = '<span class="px-2 py-0.5 rounded bg-orange-100 text-orange-800 font-bold text-[10px]">Giữ gói + Phụ đạo: ' + p.recommended_package_name + '</span>';
        } else if (p.recommendation_type === 'HOLD_TECH_BLOCKED') {
          recBadge = '<span class="px-2 py-0.5 rounded bg-rose-100 text-rose-800 font-bold text-[10px]">Tạm hoãn do lỗi kỹ thuật</span>';
        } else if (p.recommendation_type === 'HOLD_MISSING_DATA') {
          recBadge = '<span class="px-2 py-0.5 rounded bg-amber-100 text-amber-800 font-bold text-[10px]">Tạm hoãn do khuyết log LMS</span>';
        } else {
          recBadge = '<span class="px-2 py-0.5 rounded bg-slate-100 text-slate-800 font-bold text-[10px]">Giữ gói cũ: ' + p.recommended_package_name + '</span>';
        }

        let stBadge = '';
        if (p.overall_status === 'READY_FOR_EXECUTION') {
          stBadge = '<span class="px-2 py-0.5 rounded bg-emerald-50 text-emerald-700 font-bold border border-emerald-200">Sẵn sàng gửi</span>';
        } else if (p.overall_status === 'PENDING_MANAGER_APPROVAL') {
          stBadge = '<span class="px-2 py-0.5 rounded bg-purple-50 text-purple-700 font-bold border border-purple-200">Chờ Quản lý duyệt</span>';
        } else if (p.overall_status === 'BLOCKED_BY_TECH') {
          stBadge = '<span class="px-2 py-0.5 rounded bg-rose-50 text-rose-700 font-bold border border-rose-200">Chặn bởi lỗi Tech</span>';
        } else {
          stBadge = '<span class="px-2 py-0.5 rounded bg-amber-50 text-amber-700 font-bold border border-amber-200">Cần làm rõ dữ liệu</span>';
        }

        tbody.innerHTML += `
          <tr class="hover:bg-slate-50 transition">
            <td class="py-3 px-4">
              <div class="font-bold text-slate-900">${p.student_name}</div>
              <div class="text-[10px] text-slate-400 font-mono">${p.account_id} • Hạn: ${p.expiry_date}</div>
            </td>
            <td class="py-3 px-3 text-slate-700">${p.current_package_name}</td>
            <td class="py-3 px-3">
              <div class="font-bold ${p.attendance_rate_pct >= 80 ? 'text-emerald-700' : 'text-slate-700'}">${p.attendance_rate_pct}%</div>
              <div class="text-[10px] text-slate-400">${p.attended_sessions}/${p.registered_sessions} buổi</div>
            </td>
            <td class="py-3 px-3">
              ${p.is_blocked_by_tech 
                ? '<span class="text-rose-600 font-bold text-[11px]">⚠️ ' + p.open_ticket_ids + '</span>' 
                : '<span class="text-slate-400 text-[11px]">Không có</span>'}
            </td>
            <td class="py-3 px-3">${recBadge}</td>
            <td class="py-3 px-3 text-right font-extrabold text-slate-900 font-mono">
              ${p.is_revenue_grounded ? formatVND(p.final_price_vnd) : '<span class="text-amber-600 italic">Tạm giữ</span>'}
            </td>
            <td class="py-3 px-3">${stBadge}</td>
            <td class="py-3 px-3 text-center">
              <button onclick="openDraftModal('${p.account_id}')" class="px-2.5 py-1 rounded-lg bg-purple-50 hover:bg-purple-100 text-purple-700 font-bold text-[11px] border border-purple-200 transition">
                Xem nháp 💬
              </button>
            </td>
          </tr>
        `;
      });
    }

    // Modal Draft
    function openDraftModal(accId) {
      if (!analysisData) return;
      const p = (analysisData.profiles || []).find(x => x.account_id === accId);
      if (!p) return;

      document.getElementById('modalStudentName').textContent = `${p.student_name} (${p.account_id})`;
      document.getElementById('modalRecSummary').textContent = `Đề xuất: ${p.recommended_package_name} • Học phí: ${formatVND(p.final_price_vnd)}`;
      document.getElementById('modalDraftContent').value = p.draft_message || p.recommendation_reason;
      document.getElementById('draftModal').classList.remove('hidden');
    }

    function closeDraftModal() {
      document.getElementById('draftModal').classList.add('hidden');
    }

    function copyDraftMessage() {
      const textarea = document.getElementById('modalDraftContent');
      textarea.select();
      navigator.clipboard.writeText(textarea.value);
      alert('Đã sao chép nội dung tin nhắn tư vấn vào Clipboard!');
    }

    // Download Proposal Excel
    function downloadProposalExcel() {
      window.location.href = '/api/download-output-xlsx';
    }

    // AI Chat
    function toggleAiChatDrawer() {
      document.getElementById('aiChatDrawer').classList.toggle('hidden');
    }

    async function sendAiChatMessage() {
      const input = document.getElementById('aiChatInput');
      const text = input.value.trim();
      if (!text) return;

      const chatBox = document.getElementById('aiChatLogs');
      chatBox.innerHTML += `<div class="p-2.5 rounded-xl bg-slate-100 text-slate-900 text-right"><strong>Bạn:</strong> ${text}</div>`;
      input.value = '';
      chatBox.scrollTop = chatBox.scrollHeight;

      try {
        const res = await fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            message: text,
            llm_provider: document.getElementById('llmProvider').value,
            llm_model: document.getElementById('llmModel').value,
            api_key: document.getElementById('llmApiKey').value
          })
        });
        const ans = await res.json();
        chatBox.innerHTML += `<div class="p-2.5 rounded-xl bg-purple-50 text-purple-950 border border-purple-100"><strong>AI Agent:</strong> ${ans.reply}</div>`;
        chatBox.scrollTop = chatBox.scrollHeight;
      } catch (err) {
        chatBox.innerHTML += `<div class="p-2.5 rounded-xl bg-rose-50 text-rose-900">Lỗi kết nối AI: ${err.message}</div>`;
      }
    }

    function showHelpModal() {
      alert('Hệ thống Phân tích Dữ liệu Gia hạn EdTech Future:\n\n1. Chế độ chờ: Hệ thống không hiển thị số liệu nào cho đến khi bạn nhập dữ liệu hoặc bấm thanh "Demo nhanh".\n2. Thanh Demo Nhanh: Nằm ở phía trên, bấm vào để chạy ngay bộ dữ liệu giả lập 10 hồ sơ hoặc 150 tài khoản.\n3. Bước 2: Hiển thị kết quả làm sạch, tách biệt NULL ≠ 0 và kiểm tra lỗi kỹ thuật.\n4. Bước 3: Dashboard tính toán trực quan với 4 KPI, biểu đồ Area, Donut, Bar và Cổng duyệt HITL.\n5. Bước 4: Xem hồ sơ từng học viên, bản nháp tin nhắn và tải file Excel 6 sheet.');
    }

    // Helpers
    function formatVND(val) {
      if (!val) return '0 ₫';
      return Number(val).toLocaleString('vi-VN') + ' ₫';
    }

    function formatMillions(val) {
      if (!val) return '0 Tr ₫';
      return (Number(val) / 1000000).toFixed(1).replace('.0', '') + ' Tr ₫';
    }
  </script>
</body>
</html>
'@

$webPath = "c:\Users\ADMIN\Downloads\AI Agent CK\web\index.html"
[System.IO.File]::WriteAllText($webPath, $html, [System.Text.Encoding]::UTF8)
Write-Output "Successfully updated $webPath with Pastel theme, Quick Demo bar, and zero initial data."
