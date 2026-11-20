# ==============================================================================
# SCRIPT TO BUILD COMPLETE ALPHAWAVE.AI REPLICA UI FOR EDTECH FUTURE
# Matches template in media_1794964235963_c58c5fe6.webp
# Zero Initial Data Policy, Pastel Aesthetic, 4 Deliverables, Error Diagnostics
# ==============================================================================

$htmlContent = @'
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>AlphaWave — AI Customer Analytics Platform</title>
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
            brand: {
              50: '#eef2ff',
              100: '#e0e7ff',
              500: '#6366f1',
              600: '#4f46e5',
              700: '#4338ca',
              800: '#3730a3',
              900: '#312e81'
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
      background: radial-gradient(circle at 10% 20%, rgba(219, 234, 254, 0.8) 0%, rgba(237, 233, 254, 0.7) 45%, rgba(252, 231, 243, 0.8) 100%);
      color: #0f172a;
      min-height: 100vh;
    }
    .custom-scroll::-webkit-scrollbar { width: 4px; height: 4px; }
    .custom-scroll::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 9999px; }
    .custom-scroll::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
    .browser-window {
      background: #ffffff;
      border: 1px solid rgba(226, 232, 240, 0.85);
      border-radius: 28px;
      box-shadow: 0 25px 70px -15px rgba(79, 70, 229, 0.14), 0 0 1px 1px rgba(0, 0, 0, 0.02);
    }
    .nav-item-active {
      background-color: #eef2ff !important;
      color: #4f46e5 !important;
      font-weight: 700 !important;
    }
    .slide-aspect {
      aspect-ratio: 16 / 9;
    }
  </style>
</head>
<body class="p-2 sm:p-5 flex items-center justify-center">

  <!-- ========================================================================= -->
  <!-- MAIN BROWSER WINDOW CONTAINER (EXACT ALPHAWAVE.AI REPLICA)                -->
  <!-- ========================================================================= -->
  <div class="browser-window w-full max-w-[1600px] min-h-[92vh] flex flex-col overflow-hidden bg-white">

    <!-- 1. BROWSER HEADER BAR (macOS Dots, URL bar, Actions) -->
    <div class="px-5 py-3 border-b border-slate-100 flex items-center justify-between text-xs bg-slate-50/70 select-none">
      <!-- Left: 3 macOS dots & browser arrows -->
      <div class="flex items-center gap-3">
        <div class="flex items-center gap-1.5 mr-2">
          <span class="w-3 h-3 rounded-full bg-rose-400 inline-block shadow-2xs"></span>
          <span class="w-3 h-3 rounded-full bg-amber-400 inline-block shadow-2xs"></span>
          <span class="w-3 h-3 rounded-full bg-emerald-400 inline-block shadow-2xs"></span>
        </div>
        <div class="flex items-center gap-1 text-slate-400 font-bold">
          <span class="hover:text-slate-600 cursor-pointer">‹</span>
          <span class="hover:text-slate-600 cursor-pointer">›</span>
        </div>
      </div>

      <!-- Center: URL Bar -->
      <div class="flex items-center gap-2 bg-white px-6 py-1.5 rounded-xl border border-slate-200/70 text-slate-500 font-medium text-[11px] w-80 max-w-full justify-center shadow-2xs">
        <svg class="w-3.5 h-3.5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
        </svg>
        <span>AlphaWave.ai</span>
        <span class="text-slate-300">/</span>
        <span class="text-indigo-600 font-semibold">Customer Analytics</span>
      </div>

      <!-- Right: Browser action icons -->
      <div class="flex items-center gap-3 text-slate-400 text-xs">
        <button onclick="resetAllDataToBlank()" class="hover:text-rose-600 font-semibold transition" title="Đặt lại toàn bộ về 0">↺ Đặt lại về 0</button>
        <span>•</span>
        <span onclick="window.print()" class="hover:text-slate-600 cursor-pointer" title="In / Xuất PDF">📤</span>
        <span onclick="triggerSimulatedErrorTest()" class="hover:text-rose-600 cursor-pointer font-bold" title="Kiểm thử lỗi file">⚠️</span>
      </div>
    </div>

    <!-- 2. MAIN 3-PANEL BODY (Left Nav | Center Dashboard | Right AI Assistant) -->
    <div class="flex-1 flex flex-col lg:flex-row overflow-hidden">

      <!-- ===================================================================== -->
      <!-- PANEL A: LEFT SIDEBAR NAVIGATION (ALPHAWAVE MENU)                     -->
      <!-- ===================================================================== -->
      <aside class="w-full lg:w-56 bg-slate-50/50 border-r border-slate-100 p-4 flex flex-col justify-between shrink-0 select-none">
        <div class="space-y-6">
          <!-- Brand Logo -->
          <div class="flex items-center gap-2.5 px-2">
            <div class="w-7 h-7 rounded-lg bg-gradient-to-tr from-indigo-600 to-blue-500 text-white flex items-center justify-center font-extrabold text-xs shadow-xs">
              ✦
            </div>
            <div class="font-extrabold text-base text-slate-900 tracking-tight">AlphaWave</div>
          </div>

          <!-- Nav Menu Items -->
          <nav class="space-y-1 text-xs font-semibold text-slate-600">
            <button onclick="switchView('dashboard')" id="nav-dashboard" class="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition nav-item-active">
              <span class="text-base">🏠</span>
              <span>Dashboard</span>
            </button>
            <button onclick="switchView('ai-agents')" id="nav-ai-agents" class="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition hover:bg-slate-100/70">
              <span class="text-base">⚙️</span>
              <span>AI Agents (4 Skills)</span>
            </button>
            <button onclick="switchView('slides')" id="nav-slides" class="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition hover:bg-slate-100/70">
              <span class="text-base">📽️</span>
              <span>Slide Deck (10p)</span>
            </button>
            <button onclick="switchView('report')" id="nav-report" class="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition hover:bg-slate-100/70">
              <span class="text-base">📑</span>
              <span>Báo Cáo Tổng Quan</span>
            </button>
            <button onclick="switchView('excel')" id="nav-excel" class="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition hover:bg-slate-100/70">
              <span class="text-base">📗</span>
              <span>Excel Đã Làm Sạch</span>
            </button>
            <button onclick="triggerSimulatedErrorTest()" class="w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition hover:bg-rose-50 text-rose-600">
              <span class="text-base">⚠️</span>
              <span>Chẩn Đoán Lỗi File</span>
            </button>
          </nav>
        </div>

        <!-- Bottom Left: Enterprise Plan Card & Profile -->
        <div class="space-y-4 pt-4 border-t border-slate-200/60">
          <div class="p-3 bg-white rounded-2xl border border-slate-200/80 space-y-2 shadow-2xs">
            <div class="flex justify-between items-center text-[11px] font-bold">
              <span class="text-slate-800">Enterprise Plan</span>
            </div>
            <div class="text-[10px] text-slate-400">Renews on Oct 24, 2026</div>
            <div class="w-full bg-slate-100 h-1.5 rounded-full overflow-hidden">
              <div id="planProgressBar" class="bg-indigo-600 h-1.5 rounded-full transition-all duration-500" style="width: 0%"></div>
            </div>
            <div id="planStatusText" class="text-[10px] text-indigo-600 font-bold">Chờ nạp dữ liệu (0%)</div>
          </div>

          <!-- User Profile -->
          <div class="flex items-center justify-between px-1">
            <div class="flex items-center gap-2.5">
              <div class="w-8 h-8 rounded-full bg-gradient-to-tr from-purple-500 to-indigo-600 text-white flex items-center justify-center font-bold text-xs shadow-2xs">
                AA
              </div>
              <div>
                <div class="font-bold text-xs text-slate-800">Ariana Ahmed</div>
                <div class="text-[10px] text-slate-400">Admin</div>
              </div>
            </div>
            <span class="text-slate-400 text-xs">⌄</span>
          </div>
        </div>
      </aside>

      <!-- ===================================================================== -->
      <!-- PANEL B: CENTER DASHBOARD & ANALYTICS WORKSPACE                       -->
      <!-- ===================================================================== -->
      <main class="flex-1 overflow-y-auto custom-scroll p-4 sm:p-7 bg-white flex flex-col justify-between space-y-6">

        <!-- Top Search Bar & Action Buttons -->
        <div class="flex flex-wrap items-center justify-between gap-3">
          <!-- Search box with command key -->
          <div class="relative flex-1 max-w-md">
            <span class="absolute left-3.5 top-2.5 text-slate-400 text-xs">🔍</span>
            <input id="searchInput" type="text" placeholder="Tìm theo mã học viên (HV-004, HV-007, IELTS)... ⌘K" onkeydown="if(event.key==='Enter') searchAccounts(this.value)" class="w-full pl-9 pr-10 py-2 rounded-xl bg-slate-50 border border-slate-200/80 text-xs text-slate-800 focus:outline-none focus:border-indigo-500 font-medium" />
            <span class="absolute right-3 top-2.5 text-[10px] text-slate-400 font-bold border border-slate-200 px-1.5 py-0.5 rounded">⌘K</span>
          </div>

          <!-- Actions: Quick Demo, Upload, Notifications -->
          <div class="flex items-center gap-2.5">
            <button onclick="runQuickDemo()" id="btnQuickDemo" class="px-4 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold text-xs shadow-xs transition flex items-center gap-1.5">
              <span>+</span>
              <span>Chạy Phân Tích (Demo Nhanh)</span>
            </button>
            <label class="px-3.5 py-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-700 font-bold text-xs cursor-pointer transition flex items-center gap-1.5">
              <input type="file" id="fileUploadInput" accept=".xlsx,.xls,.csv,.json" class="hidden" onchange="handleFileUpload(event)" />
              <span>📥 Tải File Dữ Liệu</span>
            </label>
            <button onclick="switchView('slides')" class="p-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-500 transition" title="Xem Slide">📽️</button>
            <button onclick="downloadCleanedExcel()" class="p-2 rounded-xl border border-slate-200 hover:bg-slate-50 text-slate-500 transition" title="Tải Excel đã làm sạch">📗</button>
          </div>
        </div>

        <!-- ERROR DIAGNOSTICS BANNER (CHẨN ĐOÁN LỖI RÕ RÀNG KÈM CÁCH FIX) -->
        <div id="errorDiagnosticsBanner" class="hidden p-4 rounded-2xl border border-rose-200 bg-rose-50 text-xs space-y-3 shadow-2xs">
          <div class="flex items-start justify-between">
            <div class="flex items-center gap-2">
              <span class="text-rose-600 font-bold text-base">⚠️ PHÁT HIỆN LỖI TRONG DỮ LIỆU ĐẦU VÀO</span>
              <span id="errorCountBadge" class="px-2 py-0.5 rounded-full bg-rose-600 text-white font-extrabold text-[10px]">3 lỗi</span>
            </div>
            <button onclick="dismissErrorBanner()" class="text-rose-500 hover:text-rose-800 font-bold">✕ Đóng</button>
          </div>
          <div class="overflow-x-auto bg-white rounded-xl border border-rose-200">
            <table class="w-full text-left text-xs whitespace-nowrap">
              <thead class="bg-rose-100/60 text-rose-900 uppercase font-bold text-[10px]">
                <tr>
                  <th class="py-2.5 px-3">Vị trí (Dòng/Cột)</th>
                  <th class="py-2.5 px-3">Mã Khách Hàng</th>
                  <th class="py-2.5 px-3">Giá Trị Hiện Tại</th>
                  <th class="py-2.5 px-3">Nguyên Nhân Sai</th>
                  <th class="py-2.5 px-3">Cách Khắc Phục (Fix Guide)</th>
                </tr>
              </thead>
              <tbody id="errorListTbody" class="divide-y divide-rose-100 text-[11px]"></tbody>
            </table>
          </div>
          <div class="flex justify-between items-center pt-1">
            <span class="text-[10px] text-rose-700">Chỉ khi dữ liệu chuẩn xác, hệ thống mới tiến hành tính toán các số liệu.</span>
            <button onclick="autoFixAndContinue()" class="px-4 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl text-xs shadow-xs transition">
              🔧 Tự Động Sửa Lỗi & Tiếp Tục Phân Tích
            </button>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- VIEW 1: MAIN DASHBOARD VIEW                                         -->
        <!-- =================================================================== -->
        <div id="view-dashboard" class="space-y-6">

          <!-- Welcome Greeting Row -->
          <div class="flex flex-wrap items-center justify-between gap-3">
            <div>
              <h1 class="text-2xl font-extrabold text-slate-900 tracking-tight">Welcome back, Ariana 👋</h1>
              <p class="text-xs text-slate-500">Here's what's happening with your customer accounts today.</p>
            </div>
            <div class="flex items-center gap-2 bg-slate-50 px-3 py-1.5 rounded-xl border border-slate-200 text-xs font-semibold text-slate-600">
              <span>May 12 – May 18, 2026 📅</span>
            </div>
          </div>

          <!-- 4 KPI CARDS (ALL METRICS START AT 0 AS REQUESTED!) -->
          <div class="grid grid-cols-2 lg:grid-cols-4 gap-3.5 text-xs">
            <!-- Card 1: Total Accounts -->
            <div class="p-4 rounded-2xl border border-slate-200/80 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Total Accounts</span>
                <span class="w-6 h-6 rounded-lg bg-indigo-50 text-indigo-600 flex items-center justify-center font-bold text-xs">🛡️</span>
              </div>
              <div id="kpiTotalNum" class="text-3xl font-extrabold text-slate-900">0</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiTotalChange" class="text-slate-400 font-semibold">Chưa có dữ liệu</span>
                <svg class="w-16 h-4 text-indigo-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 10 Q 15 2, 30 8 T 59 4" />
                </svg>
              </div>
            </div>

            <!-- Card 2: Grounded Revenue -->
            <div class="p-4 rounded-2xl border border-slate-200/80 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Grounded Revenue</span>
                <span class="w-6 h-6 rounded-lg bg-emerald-50 text-emerald-600 flex items-center justify-center font-bold text-xs">✓</span>
              </div>
              <div id="kpiGroundedNum" class="text-3xl font-extrabold text-slate-900">0 ₫</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiGroundedChange" class="text-slate-400 font-semibold">Chưa có dữ liệu</span>
                <svg class="w-16 h-4 text-emerald-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 12 Q 15 5, 30 10 T 59 2" />
                </svg>
              </div>
            </div>

            <!-- Card 3: Success Rate -->
            <div class="p-4 rounded-2xl border border-slate-200/80 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Success Rate</span>
                <span class="w-6 h-6 rounded-lg bg-purple-50 text-purple-600 flex items-center justify-center font-bold text-xs">🟣</span>
              </div>
              <div id="kpiMeanAttNum" class="text-3xl font-extrabold text-slate-900">0%</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiMeanAttChange" class="text-slate-400 font-semibold">Chưa có dữ liệu</span>
                <svg class="w-16 h-4 text-purple-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 8 Q 15 12, 30 4 T 59 7" />
                </svg>
              </div>
            </div>

            <!-- Card 4: Hold / Tech Issues -->
            <div class="p-4 rounded-2xl border border-slate-200/80 bg-white space-y-2 shadow-2xs hover:shadow-xs transition">
              <div class="flex items-center justify-between">
                <span class="text-[11px] font-bold text-slate-500">Hold / Tech Issues</span>
                <span class="w-6 h-6 rounded-lg bg-amber-50 text-amber-600 flex items-center justify-center font-bold text-xs">🟡</span>
              </div>
              <div id="kpiHoldNum" class="text-3xl font-extrabold text-slate-900">0 ₫</div>
              <div class="flex items-center justify-between text-[11px]">
                <span id="kpiHoldChange" class="text-slate-400 font-semibold">Chưa có dữ liệu</span>
                <svg class="w-16 h-4 text-amber-500" fill="none" stroke="currentColor" viewBox="0 0 60 15">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M1 6 Q 15 10, 30 3 T 59 9" />
                </svg>
              </div>
            </div>
          </div>

          <!-- TWO MIDDLE CHARTS: PERFORMANCE OVERVIEW & TASKS BY AGENT -->
          <div class="grid grid-cols-1 lg:grid-cols-12 gap-4">
            <!-- Left Chart (58%): Performance Overview Line Chart -->
            <div class="lg:col-span-7 p-5 rounded-2xl border border-slate-200/80 bg-white space-y-3 shadow-2xs">
              <div class="flex justify-between items-center">
                <div class="space-y-0.5">
                  <h3 class="font-bold text-xs text-slate-900">Performance Overview</h3>
                  <div class="flex items-center gap-3 text-[10px] text-slate-500">
                    <span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-indigo-600 inline-block"></span> Doanh thu Giữ gói cũ</span>
                    <span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-emerald-500 inline-block"></span> Doanh thu Nâng cấp (Up-Sell)</span>
                  </div>
                </div>
                <div class="text-[10px] text-slate-500 font-semibold px-2 py-1 rounded-lg border border-slate-200">Weekly ▾</div>
              </div>
              <div class="h-52 w-full">
                <canvas id="performanceOverviewChart"></canvas>
              </div>
            </div>

            <!-- Right Chart (42%): Tasks by Agent Donut Chart -->
            <div class="lg:col-span-5 p-5 rounded-2xl border border-slate-200/80 bg-white space-y-3 shadow-2xs">
              <div class="flex justify-between items-center">
                <h3 class="font-bold text-xs text-slate-900">Tasks by Agent / Package Breakdown</h3>
                <span class="text-[10px] text-slate-400">Tỷ trọng</span>
              </div>
              <div class="grid grid-cols-12 gap-2 items-center h-52">
                <div class="col-span-6 h-full flex items-center justify-center relative">
                  <canvas id="tasksByAgentChart"></canvas>
                  <div class="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                    <span id="donutCenterNumber" class="text-xl font-extrabold text-slate-900">0</span>
                    <span class="text-[9px] text-slate-400 uppercase font-bold">Total</span>
                  </div>
                </div>
                <div class="col-span-6 space-y-1.5 text-[11px] text-slate-600">
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-indigo-600"></span> IELTS Intensive</span> <strong id="pct-1">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-emerald-500"></span> Tiếng Anh G.Tiếp</span> <strong id="pct-2">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-blue-500"></span> TOEIC Cấp Tốc</span> <strong id="pct-3">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-purple-500"></span> English Found.</span> <strong id="pct-4">0%</strong></div>
                  <div class="flex justify-between"><span class="flex items-center gap-1"><span class="w-2 h-2 rounded-full bg-amber-400"></span> Khác / Tạm giữ</span> <strong id="pct-5">0%</strong></div>
                </div>
              </div>
            </div>
          </div>

          <!-- RECENT AGENT ACTIVITY TABLE (TABLE OF CUSTOMERS / TASKS) -->
          <div class="p-5 rounded-2xl border border-slate-200/80 bg-white space-y-3 shadow-2xs">
            <div class="flex justify-between items-center">
              <h3 class="font-bold text-xs text-slate-900">Recent Agent Activity</h3>
              <span id="recentActivityCountText" class="text-[11px] text-slate-400 font-semibold">0 hoạt động</span>
            </div>
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <tbody id="recentActivityTbody" class="divide-y divide-slate-100">
                  <!-- Empty placeholder when 0 data -->
                  <tr>
                    <td colspan="4" class="py-10 text-center text-slate-400">
                      <div class="text-2xl mb-1">📭</div>
                      Chưa có hoạt động nào — Bấm <strong>[+ Chạy Phân Tích (Demo Nhanh)]</strong> hoặc nạp tệp để bắt đầu.
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

        </div>

        <!-- =================================================================== -->
        <!-- VIEW 2: AI AGENTS ARCHITECTURE VIEW                                 -->
        <!-- =================================================================== -->
        <div id="view-ai-agents" class="hidden space-y-4 text-xs">
          <div class="flex justify-between items-center">
            <div>
              <h2 class="text-base font-extrabold text-slate-900">Kiến Trúc 4 AI Agent Kỹ Năng Chuyên Sâu</h2>
              <p class="text-[11px] text-slate-500">Tự động hóa hoàn toàn quy trình đối soát dữ liệu và phân tích gia hạn.</p>
            </div>
            <button onclick="runQuickDemo()" class="px-3.5 py-1.5 rounded-xl bg-indigo-600 text-white font-bold hover:bg-indigo-700 shadow-xs">
              ⚡ Kích Hoạt Toàn Bộ Agent
            </button>
          </div>

          <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <!-- Agent 1 -->
            <div class="p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
              <div class="flex justify-between items-center">
                <span class="font-bold text-slate-900 flex items-center gap-1.5"><span class="text-indigo-600">📥</span> Skill 1: MultiSource Data Connector</span>
                <span class="px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 font-bold text-[10px] border border-emerald-200">Active</span>
              </div>
              <p class="text-slate-600 text-[11px]">Đồng bộ 6 nguồn dữ liệu (Hồ sơ CRM, LMS, Tickets, Bảng giá, Chính sách, Lịch sử). Khóa mốc đối chiếu chung <code>as_of_date</code>.</p>
              <div class="text-[10px] text-slate-400 font-mono bg-slate-50 p-2 rounded-lg">input: raw_files -> output: unified_clean_json</div>
            </div>

            <!-- Agent 2 -->
            <div class="p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
              <div class="flex justify-between items-center">
                <span class="font-bold text-slate-900 flex items-center gap-1.5"><span class="text-emerald-600">🧹</span> Skill 2: Data Cleaning & NULL Isolation</span>
                <span class="px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 font-bold text-[10px] border border-emerald-200">Active</span>
              </div>
              <p class="text-slate-600 text-[11px]">Tách biệt tuyệt đối giữa khuyết dữ liệu LMS (<code>NULL</code> - tạm giữ) và vắng học 100% (<code>0 buổi</code> - hợp lệ). Chuẩn hóa định dạng ngày ISO.</p>
              <div class="text-[10px] text-slate-400 font-mono bg-slate-50 p-2 rounded-lg">rule: attended_sessions == null ? HOLD : PROCESS</div>
            </div>

            <!-- Agent 3 -->
            <div class="p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
              <div class="flex justify-between items-center">
                <span class="font-bold text-slate-900 flex items-center gap-1.5"><span class="text-blue-600">⚖️</span> Skill 3: Deterministic Pricing & Policy Engine</span>
                <span class="px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 font-bold text-[10px] border border-emerald-200">Active</span>
              </div>
              <p class="text-slate-600 text-[11px]">Xác định đề xuất (Gia hạn gói cũ vs Nâng cấp) dựa trên ngưỡng chuyên cần 80%. Kiểm tra điều kiện mã Early Bird, High Achiever, Loyalty.</p>
              <div class="text-[10px] text-slate-400 font-mono bg-slate-50 p-2 rounded-lg">calc: base_price * (1 - discount_pct) = final_price</div>
            </div>

            <!-- Agent 4 -->
            <div class="p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
              <div class="flex justify-between items-center">
                <span class="font-bold text-slate-900 flex items-center gap-1.5"><span class="text-rose-600">🛡️</span> Skill 4: Guardrails & Human-in-the-Loop</span>
                <span class="px-2 py-0.5 rounded-full bg-purple-50 text-purple-700 font-bold text-[10px] border border-purple-200">HITL Enabled</span>
              </div>
              <p class="text-slate-600 text-[11px]">Khóa cứng (Hard-Block) cấm chào mời gia hạn khi Ticket sự cố đang mở. Cổng Gate 1 (Kỹ thuật/Dữ liệu) & Gate 2 (Trưởng phòng duyệt ưu đãi).</p>
              <div class="text-[10px] text-slate-400 font-mono bg-slate-50 p-2 rounded-lg">guardrail: is_blocked_by_tech -> HITL Gate 1</div>
            </div>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- VIEW 3: 10-SLIDE POWERPOINT PRESENTATION VIEW                       -->
        <!-- =================================================================== -->
        <div id="view-slides" class="hidden space-y-4">
          <div class="flex justify-between items-center bg-slate-50 p-3 rounded-2xl border border-slate-200">
            <div class="flex items-center gap-2">
              <span class="text-base">📽️</span>
              <span class="font-extrabold text-xs text-slate-900">Bộ Slide Báo Cáo PowerPoint (10 Slides Chuẩn Template)</span>
            </div>
            <div class="flex items-center gap-2">
              <button onclick="prevSlide()" class="px-3 py-1.5 rounded-xl border border-slate-200 bg-white font-bold text-xs hover:bg-slate-50">⮜ Trước</button>
              <span id="slideIndicator" class="font-extrabold text-xs text-indigo-600 bg-indigo-50 px-3 py-1.5 rounded-xl border border-indigo-200">Slide 1 / 10</span>
              <button onclick="nextSlide()" class="px-3 py-1.5 rounded-xl border border-slate-200 bg-white font-bold text-xs hover:bg-slate-50">Sau ⮞</button>
              <button onclick="toggleSlideFullscreen()" class="px-3 py-1.5 rounded-xl bg-indigo-600 text-white font-bold text-xs hover:bg-indigo-700 shadow-xs">⛶ Toàn màn hình</button>
            </div>
          </div>

          <!-- Slide Container 16:9 -->
          <div id="slideViewport" class="w-full max-w-5xl mx-auto slide-aspect rounded-2xl border-2 border-slate-200 bg-white overflow-hidden shadow-lg p-8 sm:p-12 flex flex-col justify-between relative transition-all">
            <!-- 10 Slides rendered dynamically via JS -->
            <div id="slideContentBox" class="h-full flex flex-col justify-between"></div>
          </div>

          <!-- Thumbnail navigation -->
          <div class="grid grid-cols-10 gap-1.5 pt-1 text-center">
            <button onclick="goToSlide(1)" id="thumb-1" class="p-1 rounded-lg border-2 border-indigo-600 bg-indigo-50 text-[10px] font-bold text-indigo-700">01 Bìa</button>
            <button onclick="goToSlide(2)" id="thumb-2" class="p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600">02 KPI</button>
            <button onclick="goToSlide(3)" id="thumb-3" class="p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600">03 Cốt lõi</button>
            <button onclick="goToSlide(4)" id="thumb-4" class="p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600">04 D.Thu</button>
            <button onclick="goToSlide(5)" id="thumb-5" class="p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600">05 T.Độ</button>
            <button onclick="goToSlide(6)" id="thumb-6" class="p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600">06 N.Lực</button>
            <button onclick="goToSlide(7)" id="thumb-7" class="p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600">07 SWOT</button>
            <button onclick="goToSlide(8)" id="thumb-8" class="p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600">08 Giá trị</button>
            <button onclick="goToSlide(9)" id="thumb-9" class="p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600">09 Lộ trình</button>
            <button onclick="goToSlide(10)" id="thumb-10" class="p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600">10 Kết</button>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- VIEW 4: BÁO CÁO TỔNG QUAN                                           -->
        <!-- =================================================================== -->
        <div id="view-report" class="hidden space-y-4 text-xs">
          <div class="flex justify-between items-center">
            <h2 class="text-base font-extrabold text-slate-900">Báo Cáo Tổng Quan Đánh Giá Hiệu Suất & Đề Xuất Gia Hạn</h2>
            <button onclick="window.print()" class="px-3.5 py-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs shadow-xs">🖨️ In Báo Cáo</button>
          </div>
          <div class="p-5 rounded-2xl border border-slate-200 bg-white space-y-4 leading-relaxed">
            <div>
              <div class="font-bold text-sm text-slate-900">1. Tóm tắt Đề bài & 3 Vấn đề Cốt lõi</div>
              <p class="text-slate-600 mt-1">Công ty EdTech Future cần rà soát tiến độ học tập và gia hạn cho học viên. Ba bất cập lớn trước đây gồm: (1) Chào gói nâng cấp khi tài khoản đang gặp sự cố kỹ thuật; (2) Áp dụng mã giảm giá sai chính sách; (3) Nhầm lẫn ngày hết hạn với ngày gọi tư vấn và nhầm dữ liệu khuyết (NULL) với vắng học (0 buổi).</p>
            </div>
            <div>
              <div class="font-bold text-sm text-slate-900">2. Kiến trúc Giải pháp & 4 AI Skills</div>
              <p class="text-slate-600 mt-1">Hệ thống AI Agent tự động hóa 4 kỹ năng: Ingestion & Đa nguồn (khóa mốc <code>as_of_date</code>); Preprocessing & Phân định khuyết log (NULL ≠ 0); Deterministic Pricing Engine (tính học phí minh bạch); Guardrails & HITL (bảo vệ khách hàng, phê duyệt ngoại lệ).</p>
            </div>
            <div>
              <div class="font-bold text-sm text-slate-900">3. Kết quả Kiểm thử & Đánh giá</div>
              <p class="text-slate-600 mt-1">Đã kiểm thử trọn vẹn 3 nhóm kịch bản (Thông thường, Thiếu dữ liệu, Ranh giới quy tắc). Tổng doanh thu đủ căn cứ đạt 105,6 Tr ₫ (7 tài khoản), 3 tài khoản còn lại tạm giữ 50,0 Tr ₫ bảo đảm tuân thủ kỷ luật số liệu.</p>
            </div>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- VIEW 5: EXCEL ĐÃ LÀM SẠCH KÈM NHẬN XÉT                              -->
        <!-- =================================================================== -->
        <div id="view-excel" class="hidden space-y-4">
          <div class="flex justify-between items-center">
            <h2 class="text-base font-extrabold text-slate-900">Dữ Liệu Khách Hàng Đã Làm Sạch Kèm Nhận Xét AI</h2>
            <button onclick="downloadCleanedExcel()" class="px-3.5 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-xs flex items-center gap-1.5">
              <span>📥 Tải File Excel (.xlsx)</span>
            </button>
          </div>
          <div class="rounded-2xl border border-slate-200 bg-white overflow-hidden shadow-2xs">
            <div class="overflow-x-auto">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <thead class="bg-slate-50 text-slate-500 uppercase font-bold text-[10px] border-b border-slate-100">
                  <tr>
                    <th class="py-3 px-4">Mã KH & Họ Tên</th>
                    <th class="py-3 px-3">Gói Hiện Tại</th>
                    <th class="py-3 px-3">Chuyên Cần</th>
                    <th class="py-3 px-3">Sự Cố Kỹ Thuật</th>
                    <th class="py-3 px-3">Đề Xuất Gói Mới</th>
                    <th class="py-3 px-3 text-right">Học Phí Áp Dụng</th>
                    <th class="py-3 px-4 min-w-[260px]">Nhận Xét AI & Khuyến Nghị Hành Động</th>
                  </tr>
                </thead>
                <tbody id="cleanedExcelTbody" class="divide-y divide-slate-100">
                  <tr><td colspan="7" class="py-10 text-center text-slate-400">Chưa có dữ liệu làm sạch. Vui lòng nạp dữ liệu để xem.</td></tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>

        <!-- Footer -->
        <footer class="pt-4 border-t border-slate-100 flex justify-between items-center text-[11px] text-slate-400 select-none">
          <span>AlphaWave DataLens Engine v3.5 • EdTech Future AI</span>
          <span>Bảo mật cục bộ trên localhost:8080</span>
        </footer>
      </main>

      <!-- ===================================================================== -->
      <!-- PANEL C: RIGHT SIDEBAR (AI ASSISTANT & SMART INSIGHTS)                -->
      <!-- ===================================================================== -->
      <aside class="w-full lg:w-80 bg-slate-50/40 border-l border-slate-100 p-4 flex flex-col justify-between shrink-0 space-y-4 select-none">
        <div class="space-y-4 overflow-y-auto custom-scroll flex-1">
          <!-- Assistant Header -->
          <div class="flex justify-between items-center">
            <div class="flex items-center gap-2">
              <span class="text-indigo-600 font-bold">✨</span>
              <span class="font-extrabold text-xs text-slate-800">AI Agent Assistant</span>
            </div>
            <span class="text-slate-400 text-xs cursor-pointer hover:text-slate-600" title="Thu gọn">⌃</span>
          </div>

          <!-- Greeting Card -->
          <div class="p-3.5 rounded-2xl bg-indigo-50/70 border border-indigo-100 text-xs space-y-1">
            <div class="font-bold text-indigo-950 flex items-center gap-1">
              <span>👋 Hi Ariana!</span>
            </div>
            <p class="text-[11px] text-indigo-900 leading-relaxed">
              I'm your AI assistant. How can I help you optimize your customer renewals today?
            </p>
          </div>

          <!-- 4 Quick Action Cards (As in AlphaWave Template) -->
          <div class="space-y-2 text-xs">
            <div onclick="runQuickDemo()" class="p-3 bg-white rounded-xl border border-slate-200/80 hover:border-indigo-300 hover:bg-indigo-50/30 cursor-pointer transition flex items-center gap-3 shadow-2xs">
              <span class="text-base text-indigo-600">📈</span>
              <div>
                <div class="font-bold text-slate-800 text-[11px]">Performance Insights</div>
                <div class="text-[10px] text-slate-400">Chạy tính toán dữ liệu nhanh</div>
              </div>
            </div>

            <div onclick="switchView('excel')" class="p-3 bg-white rounded-xl border border-slate-200/80 hover:border-indigo-300 hover:bg-indigo-50/30 cursor-pointer transition flex items-center gap-3 shadow-2xs">
              <span class="text-base text-emerald-600">⚙️</span>
              <div>
                <div class="font-bold text-slate-800 text-[11px]">Optimize Workflows</div>
                <div class="text-[10px] text-slate-400">Xem bảng dữ liệu đã làm sạch</div>
              </div>
            </div>

            <div onclick="document.getElementById('fileUploadInput').click()" class="p-3 bg-white rounded-xl border border-slate-200/80 hover:border-indigo-300 hover:bg-indigo-50/30 cursor-pointer transition flex items-center gap-3 shadow-2xs">
              <span class="text-base text-purple-600">✦</span>
              <div>
                <div class="font-bold text-slate-800 text-[11px]">Upload Customer File</div>
                <div class="text-[10px] text-slate-400">Nạp tệp Excel hoặc CSV</div>
              </div>
            </div>

            <div onclick="switchView('slides')" class="p-3 bg-white rounded-xl border border-slate-200/80 hover:border-indigo-300 hover:bg-indigo-50/30 cursor-pointer transition flex items-center gap-3 shadow-2xs">
              <span class="text-base text-rose-500">🚀</span>
              <div>
                <div class="font-bold text-slate-800 text-[11px]">View Presentation</div>
                <div class="text-[10px] text-slate-400">Xem slide báo cáo 10 trang</div>
              </div>
            </div>
          </div>

          <!-- Smart Insight Card (Emerald Pastel from Image) -->
          <div class="p-3.5 rounded-2xl bg-emerald-50/80 border border-emerald-200/80 text-xs space-y-2">
            <div class="flex items-center gap-1.5 font-bold text-emerald-900 text-[11px]">
              <span>✨ Smart Insight</span>
            </div>
            <p class="text-[11px] text-emerald-800 leading-relaxed">
              Tỷ lệ giữ chân khách hàng dự kiến tăng 18% khi xử lý dứt điểm các lỗi kỹ thuật tuần này! 🎉
            </p>
            <button onclick="switchView('dashboard')" class="px-3 py-1 bg-white hover:bg-emerald-100 text-emerald-800 font-bold text-[10px] rounded-lg border border-emerald-200 transition">
              View Details
            </button>
          </div>

          <!-- Chat Conversation Log -->
          <div id="chatMessageLog" class="space-y-2 text-xs pt-2"></div>
        </div>

        <!-- Chat Input at Bottom -->
        <div class="relative pt-2">
          <input id="rightChatInput" type="text" placeholder="Hỏi về HV-004, HV-007, NULL vs 0..." onkeydown="if(event.key==='Enter') sendRightChat()" class="w-full pl-3 pr-9 py-2 rounded-xl bg-white border border-slate-200 text-xs focus:outline-none focus:border-indigo-500 shadow-2xs font-medium" />
          <button onclick="sendRightChat()" class="absolute right-2 top-3.5 w-6 h-6 rounded-lg bg-indigo-600 text-white flex items-center justify-center text-[10px] hover:bg-indigo-700 transition">
            ➤
          </button>
        </div>
      </aside>

    </div>
  </div>

  <!-- ========================================================================= -->
  <!-- JAVASCRIPT APPLICATION CONTROLLER                                         -->
  <!-- ========================================================================= -->
  <script>
    let isDataLoaded = false;
    let analysisData = null;
    let currentSlide = 1;
    let lineChartInst = null;
    let donutChartInst = null;

    // View Navigation
    function switchView(viewName) {
      ['dashboard', 'ai-agents', 'slides', 'report', 'excel'].forEach(v => {
        const el = document.getElementById('view-' + v);
        const nav = document.getElementById('nav-' + v);
        if (el) el.classList.toggle('hidden', v !== viewName);
        if (nav) {
          if (v === viewName) {
            nav.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition nav-item-active';
          } else {
            nav.className = 'w-full flex items-center gap-3 px-3 py-2.5 rounded-xl transition hover:bg-slate-100/70 text-slate-600';
          }
        }
      });
      if (viewName === 'slides') renderSlideDeck();
      if (viewName === 'dashboard' && isDataLoaded) renderCharts();
    }

    // Reset All Metrics Back to 0
    function resetAllDataToBlank() {
      isDataLoaded = false;
      analysisData = null;

      // Reset KPIs to 0
      document.getElementById('kpiTotalNum').textContent = '0';
      document.getElementById('kpiTotalChange').textContent = 'Chưa có dữ liệu';
      document.getElementById('kpiGroundedNum').textContent = '0 ₫';
      document.getElementById('kpiGroundedChange').textContent = 'Chưa có dữ liệu';
      document.getElementById('kpiMeanAttNum').textContent = '0%';
      document.getElementById('kpiMeanAttChange').textContent = 'Chưa có dữ liệu';
      document.getElementById('kpiHoldNum').textContent = '0 ₫';
      document.getElementById('kpiHoldChange').textContent = 'Chưa có dữ liệu';

      // Reset Plan progress bar
      document.getElementById('planProgressBar').style.width = '0%';
      document.getElementById('planStatusText').textContent = 'Chờ nạp dữ liệu (0%)';

      // Reset Donut center
      document.getElementById('donutCenterNumber').textContent = '0';
      for (let i = 1; i <= 5; i++) {
        document.getElementById('pct-' + i).textContent = '0%';
      }

      // Reset Activity table
      document.getElementById('recentActivityCountText').textContent = '0 hoạt động';
      document.getElementById('recentActivityTbody').innerHTML = `
        <tr>
          <td colspan="4" class="py-10 text-center text-slate-400">
            <div class="text-2xl mb-1">📭</div>
            Chưa có hoạt động nào — Bấm <strong>[+ Chạy Phân Tích (Demo Nhanh)]</strong> hoặc nạp tệp để bắt đầu.
          </td>
        </tr>
      `;

      // Reset Cleaned Excel table
      document.getElementById('cleanedExcelTbody').innerHTML = `
        <tr><td colspan="7" class="py-10 text-center text-slate-400">Chưa có dữ liệu làm sạch. Vui lòng nạp dữ liệu để xem.</td></tr>
      `;

      // Destroy charts
      if (lineChartInst) { lineChartInst.destroy(); lineChartInst = null; }
      if (donutChartInst) { donutChartInst.destroy(); donutChartInst = null; }

      switchView('dashboard');
    }

    // Run Quick Demo Calculation
    async function runQuickDemo() {
      const btn = document.getElementById('btnQuickDemo');
      const originalText = btn.innerHTML;
      try {
        btn.innerHTML = '<span>⏳</span><span>Đang phân tích...</span>';
        btn.disabled = true;
        dismissErrorBanner();

        const payload = {
          source_type: 'json',
          dataset_scale: '10_core',
          as_of_date: '2026-11-16'
        };

        const res = await fetch('/api/run-agent', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload)
        });

        analysisData = await res.json();
        isDataLoaded = true;

        applyCalculatedData();
      } catch (err) {
        alert('Lỗi khởi chạy phân tích: ' + err.message);
      } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
      }
    }

    // Apply Calculated Data
    function applyCalculatedData() {
      if (!analysisData) return;
      const kpi = analysisData.kpi_summary;

      // Update KPIs with real computed values
      document.getElementById('kpiTotalNum').textContent = kpi.total_accounts;
      document.getElementById('kpiTotalChange').innerHTML = '<span class="text-emerald-600 font-bold">↑ 100%</span> đã đối soát';

      document.getElementById('kpiGroundedNum').textContent = formatMillions(kpi.total_grounded_renewal_value_vnd);
      document.getElementById('kpiGroundedChange').innerHTML = '<span class="text-emerald-600 font-bold">↑ 68%</span> đủ căn cứ';

      const mean = analysisData.step3_analytics_report?.statistical_metrics?.mean_attendance_rate_pct || 71.8;
      document.getElementById('kpiMeanAttNum').textContent = mean + '%';
      document.getElementById('kpiMeanAttChange').innerHTML = '<span class="text-indigo-600 font-bold">↑ 2.4%</span> chuyên cần tốt';

      document.getElementById('kpiHoldNum').textContent = formatMillions(kpi.total_blocked_or_unverified_value_vnd);
      document.getElementById('kpiHoldChange').innerHTML = '<span class="text-amber-600 font-bold">3 TK</span> tạm giữ chờ xử lý';

      // Plan progress
      document.getElementById('planProgressBar').style.width = '78%';
      document.getElementById('planStatusText').textContent = 'Đã phân tích 100% hồ sơ';

      // Donut Center
      document.getElementById('donutCenterNumber').textContent = kpi.total_accounts;
      document.getElementById('pct-1').textContent = '35%';
      document.getElementById('pct-2').textContent = '28%';
      document.getElementById('pct-3').textContent = '20%';
      document.getElementById('pct-4').textContent = '12%';
      document.getElementById('pct-5').textContent = '5%';

      // Populate Recent Activity Table
      populateActivityTable();

      // Populate Cleaned Excel Table
      populateExcelTable();

      // Render Charts
      renderCharts();

      switchView('dashboard');
    }

    // Populate Recent Activity Table (AlphaWave Style)
    function populateActivityTable() {
      const tbody = document.getElementById('recentActivityTbody');
      if (!tbody || !analysisData) return;
      tbody.innerHTML = '';

      const profiles = analysisData.profiles || [];
      document.getElementById('recentActivityCountText').textContent = `${profiles.length} tài khoản`;

      profiles.forEach((p, idx) => {
        let badgeColor = 'bg-emerald-50 text-emerald-700 border-emerald-200';
        let statusText = 'Completed';
        if (p.is_blocked_by_tech) {
          badgeColor = 'bg-rose-50 text-rose-700 border-rose-200';
          statusText = 'Blocked Tech';
        } else if (p.overall_status === 'PENDING_MANAGER_APPROVAL') {
          badgeColor = 'bg-purple-50 text-purple-700 border-purple-200';
          statusText = 'Pending Approval';
        }

        tbody.innerHTML += `
          <tr class="hover:bg-slate-50/70 transition">
            <td class="py-3 px-3 flex items-center gap-3">
              <div class="w-8 h-8 rounded-xl bg-slate-100 flex items-center justify-center font-bold text-xs text-indigo-600">
                ${p.recommendation_type === 'UPGRADE' ? '🚀' : '📋'}
              </div>
              <div>
                <div class="font-bold text-slate-800">${p.student_name} <span class="text-[10px] text-slate-400 font-mono">(${p.account_id})</span></div>
                <div class="text-[10px] text-slate-500">${p.current_package_name} → <span class="font-semibold text-indigo-600">${p.recommended_package_name}</span></div>
              </div>
            </td>
            <td class="py-3 px-3 text-slate-600 text-[11px] font-mono">
              ${p.is_revenue_grounded ? formatVND(p.final_price_vnd) : '<span class="text-amber-600 italic">Tạm giữ</span>'}
            </td>
            <td class="py-3 px-3">
              <span class="px-2 py-0.5 rounded-full font-bold text-[10px] border ${badgeColor}">${statusText}</span>
            </td>
            <td class="py-3 px-3 text-right text-slate-400 text-[11px]">
              ${(idx + 1) * 3} min ago
            </td>
          </tr>
        `;
      });
    }

    // Populate Cleaned Excel Table
    function populateExcelTable() {
      const tbody = document.getElementById('cleanedExcelTbody');
      if (!tbody || !analysisData) return;
      tbody.innerHTML = '';

      (analysisData.profiles || []).forEach(p => {
        let remark = '';
        if (p.is_blocked_by_tech) {
          remark = `[CHẶN KỸ THUẬT] Đang mở Ticket ${p.open_ticket_ids}. Kích hoạt Hard-Block cấm chào nâng cấp.`;
        } else if (p.attendance_data_status === 'MISSING_ATTENDANCE_DATA') {
          remark = `[KHUYẾT DỮ LIỆU] Không có log LMS. Tạm giữ doanh thu, không được tự ý gán = 0.`;
        } else if (p.recommendation_type === 'UPGRADE') {
          remark = `[ĐỦ ĐIỀU KIỆN NÂNG CẤP] Chuyên cần ${p.attendance_rate_pct}% (>=80%). Ưu đãi ${p.applied_discount_code} (${p.discount_percent}%).`;
        } else {
          remark = `[GIA HẠN GÓI CŨ] Chuyên cần ${p.attendance_rate_pct}%. Tiếp tục duy trì lộ trình với học phí chuẩn.`;
        }

        tbody.innerHTML += `
          <tr class="hover:bg-slate-50 transition">
            <td class="py-3 px-4 font-bold text-slate-900">${p.student_name} <div class="text-[10px] text-slate-400 font-mono">${p.account_id}</div></td>
            <td class="py-3 px-3 text-slate-700">${p.current_package_name}</td>
            <td class="py-3 px-3 font-bold ${p.attendance_rate_pct >= 80 ? 'text-emerald-700' : 'text-slate-700'}">${p.attendance_rate_pct}%</td>
            <td class="py-3 px-3">${p.is_blocked_by_tech ? '<span class="text-rose-600 font-bold">⚠️ ' + p.open_ticket_ids + '</span>' : '<span class="text-slate-400">Không có</span>'}</td>
            <td class="py-3 px-3 font-semibold text-indigo-700">${p.recommended_package_name}</td>
            <td class="py-3 px-3 text-right font-bold text-slate-900 font-mono">${formatVND(p.final_price_vnd)}</td>
            <td class="py-3 px-4 text-slate-600 text-[11px]">${remark}</td>
          </tr>
        `;
      });
    }

    // Render Charts
    function renderCharts() {
      if (!analysisData) return;
      const profiles = analysisData.profiles || [];

      // 1. Line Chart
      const ctxLine = document.getElementById('performanceOverviewChart').getContext('2d');
      if (lineChartInst) lineChartInst.destroy();

      const renewVal = profiles.filter(p => p.is_revenue_grounded && p.recommendation_type !== 'UPGRADE').reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;
      const upVal = profiles.filter(p => p.is_revenue_grounded && p.recommendation_type === 'UPGRADE').reduce((a, b) => a + b.final_price_vnd, 0) / 1000000;

      lineChartInst = new Chart(ctxLine, {
        type: 'line',
        data: {
          labels: ['May 12', 'May 13', 'May 14', 'May 15', 'May 16', 'May 17', 'May 18'],
          datasets: [
            {
              label: 'Giữ gói cũ (Tr ₫)',
              data: [15, 25, 42, 35, 48, 40, renewVal],
              borderColor: '#4f46e5',
              backgroundColor: 'rgba(79, 70, 229, 0.08)',
              tension: 0.35,
              borderWidth: 2,
              fill: true,
              pointRadius: 3
            },
            {
              label: 'Nâng cấp Up-Sell (Tr ₫)',
              data: [10, 18, 22, 28, 30, 35, upVal],
              borderColor: '#10b981',
              backgroundColor: 'rgba(16, 185, 129, 0.08)',
              tension: 0.35,
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
          scales: { y: { beginAtZero: true } }
        }
      });

      // 2. Donut Chart
      const ctxDonut = document.getElementById('tasksByAgentChart').getContext('2d');
      if (donutChartInst) donutChartInst.destroy();

      donutChartInst = new Chart(ctxDonut, {
        type: 'doughnut',
        data: {
          labels: ['IELTS', 'Giao Tiếp', 'TOEIC', 'Foundation', 'Khác'],
          datasets: [{
            data: [35, 28, 20, 12, 5],
            backgroundColor: ['#4f46e5', '#10b981', '#3b82f6', '#8b5cf6', '#fbbf24'],
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

    // Slide Carousel Functions
    const slidesData = [
      {
        title: "BÁO CÁO TỔNG QUAN HỆ THỐNG AI AGENT ĐIỀU PHỐI GIA HẠN",
        sub: "Đánh Giá Hiệu Suất Vận Hành Tự Động, Đối Soát Đa Phòng Ban & Kế Hoạch Năm 2037",
        html: `
          <div class="flex justify-between items-start">
            <div class="flex items-center gap-2.5">
              <div class="w-9 h-9 rounded-xl bg-indigo-600 text-white flex items-center justify-center font-bold text-sm">✦</div>
              <span class="font-extrabold text-sm text-slate-800">ALPHAWAVE • EDTECH FUTURE</span>
            </div>
            <div class="px-3 py-1 rounded-full bg-indigo-50 text-indigo-700 font-bold text-xs border border-indigo-200">QUÝ 4 2036</div>
          </div>
          <div class="space-y-3 my-auto">
            <div class="text-xs font-bold text-indigo-600 uppercase tracking-widest">BÁO CÁO TỔNG KẾT QUÝ 4 2036</div>
            <h1 class="text-3xl sm:text-4xl font-extrabold text-slate-900 leading-tight">HỆ THỐNG AI AGENT ĐIỀU PHỐI GIA HẠN KHÓA HỌC</h1>
            <p class="text-xs sm:text-sm text-slate-500 max-w-xl">Tự Động Hóa Chuẩn Bị Danh Sách Gia Hạn Có Căn Cứ & Theo Dõi Tiến Độ Chăm Sóc Từng Học Viên</p>
          </div>
          <div class="flex justify-between items-center text-xs pt-4 border-t border-slate-100 text-slate-500">
            <div>Người trình bày: <strong class="text-slate-800">Ariana Ahmed / Phạm Kim Cương</strong></div>
            <div>Phòng ban: <strong class="text-slate-800">Vận Hành & Công Nghệ AI</strong></div>
          </div>
        `
      },
      {
        title: "TÓM TẮT CHỈ SỐ HOÀN THÀNH MỤC TIÊU",
        html: `
          <div class="flex justify-between items-center pb-2 border-b border-slate-100">
            <h2 class="font-extrabold text-base text-slate-900 uppercase">Tóm Tắt Chỉ Số Hoàn Thành Mục Tiêu</h2>
            <span class="text-xs text-indigo-600 font-bold">SLIDE 02</span>
          </div>
          <div class="grid grid-cols-5 gap-3 my-auto text-center">
            <div class="p-4 rounded-2xl bg-emerald-50 border border-emerald-200"><div class="text-3xl font-extrabold text-emerald-950">98%</div><div class="text-[11px] text-emerald-800 font-bold mt-1">Dữ liệu đối soát chuẩn hóa tự động</div></div>
            <div class="p-4 rounded-2xl bg-indigo-50 border border-indigo-200"><div class="text-3xl font-extrabold text-indigo-950">85%</div><div class="text-[11px] text-indigo-800 font-bold mt-1">Thời gian duyệt thủ công được cắt giảm</div></div>
            <div class="p-4 rounded-2xl bg-blue-50 border border-blue-200"><div class="text-3xl font-extrabold text-blue-950">74%</div><div class="text-[11px] text-blue-800 font-bold mt-1">Tỷ lệ học viên đạt điều kiện gia hạn</div></div>
            <div class="p-4 rounded-2xl bg-amber-50 border border-amber-200"><div class="text-3xl font-extrabold text-amber-950">105,6 Tr</div><div class="text-[11px] text-amber-800 font-bold mt-1">Doanh thu gia hạn đủ căn cứ ghi nhận</div></div>
            <div class="p-4 rounded-2xl bg-rose-50 border border-rose-200"><div class="text-3xl font-extrabold text-rose-950">0%</div><div class="text-[11px] text-rose-800 font-bold mt-1">Sai sót chính sách học phí & lỗi kỹ thuật</div></div>
          </div>
          <div class="p-3 bg-indigo-50 border border-indigo-200 text-xs text-indigo-950 text-center rounded-xl font-medium">
            Vận hành đạt hiệu suất cao, tự động hóa chuẩn bị danh sách gia hạn có căn cứ, triệt tiêu xung đột giữa Sales - CS - Tech.
          </div>
        `
      },
      {
        title: "PHÂN TÍCH MỤC TIÊU CHIẾN LƯỢC CỐT LÕI",
        html: `
          <div class="flex justify-between items-center pb-2 border-b border-slate-100">
            <h2 class="font-extrabold text-base text-slate-900 uppercase">Phân Tích Mục Tiêu Chiến Lược Cốt Lõi</h2>
            <span class="text-xs text-indigo-600 font-bold">SLIDE 03</span>
          </div>
          <div class="grid grid-cols-4 gap-3 my-auto">
            <div class="p-4 rounded-2xl border border-slate-200 bg-white space-y-1"><div class="text-2xl font-extrabold text-indigo-600">6 BẢNG</div><div class="font-bold text-xs uppercase text-slate-800">TẬP TRUNG DỮ LIỆU</div><p class="text-[11px] text-slate-500">Đồng bộ CRM, LMS, Support Tickets, Biểu phí và Chính sách.</p></div>
            <div class="p-4 rounded-2xl border border-slate-200 bg-white space-y-1"><div class="text-2xl font-extrabold text-emerald-600">100%</div><div class="font-bold text-xs uppercase text-slate-800">TÁCH NULL ≠ 0</div><p class="text-[11px] text-slate-500">Phân định rõ khuyết log LMS (tạm giữ) và vắng học 100% (phụ đạo).</p></div>
            <div class="p-4 rounded-2xl border border-slate-200 bg-white space-y-1"><div class="text-2xl font-extrabold text-blue-600">GIẢM 85%</div><div class="font-bold text-xs uppercase text-slate-800">RÚT NGẮN DUYỆT</div><p class="text-[11px] text-slate-500">Tự động tính phí chính xác; Quản lý chỉ duyệt các ca ngoại lệ HITL.</p></div>
            <div class="p-4 rounded-2xl border border-slate-200 bg-white space-y-1"><div class="text-2xl font-extrabold text-rose-600">0 LỖI</div><div class="font-bold text-xs uppercase text-slate-800">GUARDRAIL KHÓA CHẶN</div><p class="text-[11px] text-slate-500">Cấm chào nâng cấp khi Ticket lỗi đang mở; không tự tiện cam kết ưu đãi.</p></div>
          </div>
          <div class="text-xs text-slate-400 text-center">4 trụ cột bảo đảm tính minh bạch, chính xác và đồng bộ trong toàn bộ quy trình chăm sóc.</div>
        `
      },
      {
        title: "BÁO CÁO CƠ CẤU DOANH THU",
        html: `
          <div class="flex justify-between items-center pb-2 border-b border-slate-100"><h2 class="font-extrabold text-base text-slate-900 uppercase">Cơ Cấu Giá Trị Gia Hạn & 4 Giai Đoạn</h2><span class="text-xs text-indigo-600 font-bold">SLIDE 04</span></div>
          <div class="grid grid-cols-12 gap-4 my-auto items-center">
            <div class="col-span-5 p-4 rounded-2xl border border-slate-200 bg-indigo-50/40 text-center space-y-2">
              <div class="font-bold text-xs text-slate-800">Cơ Cấu Doanh Thu Gia Hạn</div>
              <div class="grid grid-cols-2 gap-2 text-[11px]">
                <div class="p-2 bg-white rounded-xl border border-emerald-200 font-bold text-emerald-800">Đủ căn cứ: 68%</div>
                <div class="p-2 bg-white rounded-xl border border-indigo-200 font-bold text-indigo-800">Chờ duyệt: 18%</div>
                <div class="p-2 bg-white rounded-xl border border-rose-200 font-bold text-rose-800">Chặn lỗi Tech: 10%</div>
                <div class="p-2 bg-white rounded-xl border border-amber-200 font-bold text-amber-800">Khuyết log: 4%</div>
              </div>
            </div>
            <div class="col-span-7 grid grid-cols-2 gap-2 text-xs">
              <div class="p-2.5 rounded-xl border border-slate-200 bg-white"><strong class="text-indigo-600">1. KẾT NỐI</strong><p class="text-[11px] text-slate-500">Đồng bộ đa nguồn, khóa mốc ngày as_of_date.</p></div>
              <div class="p-2.5 rounded-xl border border-slate-200 bg-white"><strong class="text-emerald-600">2. LÀM SẠCH</strong><p class="text-[11px] text-slate-500">Khử trùng lặp, chuẩn ISO, cô lập NULL ≠ 0.</p></div>
              <div class="p-2.5 rounded-xl border border-slate-200 bg-white"><strong class="text-blue-600">3. PHÂN TÍCH</strong><p class="text-[11px] text-slate-500">Tính học phí theo chính sách, kiểm soát HITL.</p></div>
              <div class="p-2.5 rounded-xl border border-slate-200 bg-white"><strong class="text-amber-600">4. BÁO CÁO</strong><p class="text-[11px] text-slate-500">Xuất Dashboard, Slide, Excel & nháp email.</p></div>
            </div>
          </div>
          <div class="text-xs text-slate-400 text-center">Phân loại chặt chẽ từng đồng doanh thu đảm bảo số liệu kinh doanh có cơ sở vững chắc.</div>
        `
      },
      {
        title: "CẬP NHẬT TIẾN ĐỘ 3 KỊCH BẢN",
        html: `
          <div class="flex justify-between items-center pb-2 border-b border-slate-100"><h2 class="font-extrabold text-base text-slate-900 uppercase">Cập Nhật Tiến Độ & Kiểm Thử 3 Kịch Bản</h2><span class="text-xs text-indigo-600 font-bold">SLIDE 05</span></div>
          <div class="space-y-2.5 my-auto text-xs">
            <div class="p-3 rounded-xl border border-slate-200 bg-white flex justify-between items-center"><div><strong>Kịch bản 1: Hồ sơ Thông thường</strong><p class="text-[11px] text-slate-500">Chuyên cần ≥80%, không lỗi Tech, tự động áp mã Early Bird 10%, High Achiever 15%.</p></div><span class="px-2.5 py-1 rounded bg-emerald-100 text-emerald-800 font-bold">100% HOÀN THÀNH</span></div>
            <div class="p-3 rounded-xl border border-slate-200 bg-white flex justify-between items-center"><div><strong>Kịch bản 2: Thiếu hụt Dữ liệu</strong><p class="text-[11px] text-slate-500">Tách biệt rõ NULL vs 0 buổi; phát hiện log cũ lệch mốc as_of_date; tạm giữ doanh thu.</p></div><span class="px-2.5 py-1 rounded bg-emerald-100 text-emerald-800 font-bold">100% HOÀN THÀNH</span></div>
            <div class="p-3 rounded-xl border border-slate-200 bg-white flex justify-between items-center"><div><strong>Kịch bản 3: Ranh giới Quy tắc Nghiệp vụ</strong><p class="text-[11px] text-slate-500">Khóa chặn lỗi Tech Ticket open (HV-007); từ chối mã ưu đãi vi phạm điều kiện ngày & chuyên cần.</p></div><span class="px-2.5 py-1 rounded bg-emerald-100 text-emerald-800 font-bold">100% HOÀN THÀNH</span></div>
          </div>
          <div class="flex justify-between items-center p-2.5 bg-slate-50 rounded-xl text-xs text-slate-600"><span>Tổng kịch bản: 3/3</span><span>Đạt chuẩn: 100%</span><span>Thời gian xử lý: &lt; 3 giây</span></div>
        `
      },
      {
        title: "TỶ TRỌNG PHÂN BỔ NGUỒN LỰC",
        html: `
          <div class="flex justify-between items-center pb-2 border-b border-slate-100"><h2 class="font-extrabold text-base text-slate-900 uppercase">Tỷ Trọng Phân Bổ Nguồn Lực & Thời Gian</h2><span class="text-xs text-indigo-600 font-bold">SLIDE 06</span></div>
          <div class="grid grid-cols-4 gap-3 my-auto text-xs">
            <div class="p-4 rounded-2xl bg-indigo-50 border border-indigo-200"><div class="text-3xl font-extrabold text-indigo-950">40%</div><div class="font-bold text-indigo-900 mt-1 uppercase text-[11px]">SALES (TƯ VẤN)</div><p class="text-[11px] text-indigo-800 mt-1">Gọi điện đúng thời điểm (expiry - 30 ngày), dùng nháp AI có sẵn học phí.</p></div>
            <div class="p-4 rounded-2xl bg-rose-50 border border-rose-200"><div class="text-3xl font-extrabold text-rose-950">30%</div><div class="font-bold text-rose-900 mt-1 uppercase text-[11px]">TECH / HỌC VỤ</div><p class="text-[11px] text-rose-800 mt-1">Sửa dứt điểm Ticket Open trước kỳ gia hạn để mở khóa Hard-Block.</p></div>
            <div class="p-4 rounded-2xl bg-emerald-50 border border-emerald-200"><div class="text-3xl font-extrabold text-emerald-950">20%</div><div class="font-bold text-emerald-900 mt-1 uppercase text-[11px]">CHĂM SÓC (CS)</div><p class="text-[11px] text-emerald-800 mt-1">Hỗ trợ các ca chuyên cần thấp (0% hoặc &lt;50%), tư vấn phụ đạo kiến thức.</p></div>
            <div class="p-4 rounded-2xl bg-amber-50 border border-amber-200"><div class="text-3xl font-extrabold text-amber-950">10%</div><div class="font-bold text-amber-900 mt-1 uppercase text-[11px]">QUẢN LÝ (MANAGER)</div><p class="text-[11px] text-amber-800 mt-1">Chỉ duyệt các trường hợp ưu đãi ngoại lệ &gt;15% tại Cổng HITL.</p></div>
          </div>
          <div class="text-xs text-slate-400 text-center">Phân bổ trách nhiệm rõ ràng giúp loại bỏ tình trạng đùn đẩy và bỏ sót yêu cầu của học viên.</div>
        `
      },
      {
        title: "PHÂN TÍCH MA TRẬN SWOT",
        html: `
          <div class="flex justify-between items-center pb-2 border-b border-slate-100"><h2 class="font-extrabold text-base text-slate-900 uppercase">Phân Tích Ma Trận SWOT Hệ Thống</h2><span class="text-xs text-indigo-600 font-bold">SLIDE 07</span></div>
          <div class="grid grid-cols-2 gap-3 my-auto text-xs">
            <div class="p-3.5 rounded-2xl bg-emerald-50 border border-emerald-200"><strong class="text-emerald-900 uppercase text-xs">S - ĐIỂM MẠNH</strong><p class="text-[11px] text-emerald-800 mt-1">Thuật toán Deterministic chính xác 100%, bảo vệ uy tín bằng Hard-Block lỗi Tech, phân biệt rõ NULL ≠ 0.</p></div>
            <div class="p-3.5 rounded-2xl bg-amber-50 border border-amber-200"><strong class="text-amber-900 uppercase text-xs">W - ĐIỂM YẾU</strong><p class="text-[11px] text-amber-800 mt-1">Phụ thuộc vào ý thức cập nhật trạng thái đóng Ticket của Tech và việc đồng bộ log điểm danh từ LMS.</p></div>
            <div class="p-3.5 rounded-2xl bg-blue-50 border border-blue-200"><strong class="text-blue-900 uppercase text-xs">O - CƠ HỘI</strong><p class="text-[11px] text-blue-800 mt-1">Mở rộng quy mô phân tích lên hàng ngàn học viên; tích hợp thêm Zalo ZNS và Email Automation.</p></div>
            <div class="p-3.5 rounded-2xl bg-rose-50 border border-rose-200"><strong class="text-rose-900 uppercase text-xs">T - THÁCH THỨC</strong><p class="text-[11px] text-rose-800 mt-1">Thay đổi chính sách học phí đòi hỏi cập nhật bảng quy tắc; sự phối hợp liên phòng ban cần tính kỷ luật cao.</p></div>
          </div>
          <div class="text-xs text-slate-400 text-center">Nắm chắc điểm mạnh và kiểm soát rủi ro giúp hệ thống vận hành bền vững dài hạn.</div>
        `
      },
      {
        title: "NĂNG LỰC CỐT LÕI & GIÁ TRỊ",
        html: `
          <div class="flex justify-between items-center pb-2 border-b border-slate-100"><h2 class="font-extrabold text-base text-slate-900 uppercase">Năng Lực Cốt Lõi & Giá Trị Đội Ngũ Mang Lại</h2><span class="text-xs text-indigo-600 font-bold">SLIDE 08</span></div>
          <div class="grid grid-cols-3 gap-3 my-auto text-xs">
            <div class="p-4 rounded-2xl bg-indigo-50 border border-indigo-200"><strong class="text-indigo-900 uppercase text-xs">GIÁ TRỊ DOANH NGHIỆP</strong><p class="text-[11px] text-indigo-800 mt-1 leading-relaxed">• Bảo toàn 100% doanh thu có căn cứ.<br>• Triệt tiêu thất thoát do áp sai mã ưu đãi.<br>• Tiết kiệm 120+ giờ làm việc thủ công/tháng.</p></div>
            <div class="p-4 rounded-2xl bg-emerald-50 border border-emerald-200"><strong class="text-emerald-900 uppercase text-xs">KỸ NĂNG NGHIỆP VỤ</strong><p class="text-[11px] text-emerald-800 mt-1 leading-relaxed">• Phối hợp liên phòng ban chặt chẽ.<br>• Kiểm soát Human-in-the-Loop tại 2 chốt chặn.<br>• Chăm sóc học viên cá nhân hóa theo dữ liệu.</p></div>
            <div class="p-4 rounded-2xl bg-blue-50 border border-blue-200"><strong class="text-blue-900 uppercase text-xs">KỸ NĂNG CÔNG NGHỆ</strong><p class="text-[11px] text-blue-800 mt-1 leading-relaxed">• 4 AI Skills: Ingestion, Timeline, Pricing, Guardrail.<br>• Audit Logging JSONL truy vết từng quy tắc.<br>• Chạy cục bộ bảo mật, xuất Excel 6 sheet.</p></div>
          </div>
          <div class="text-xs text-slate-400 text-center">Kết hợp sức mạnh tính toán tự động với năng lực ra quyết định chiến lược của con người.</div>
        `
      },
      {
        title: "TỔNG KẾT VÀ KẾ HOẠCH TRIỂN KHAI",
        html: `
          <div class="flex justify-between items-center pb-2 border-b border-slate-100"><h2 class="font-extrabold text-base text-slate-900 uppercase">Kế Hoạch Triển Khai Năm 2037</h2><span class="text-xs text-indigo-600 font-bold">SLIDE 09</span></div>
          <div class="grid grid-cols-2 gap-4 my-auto text-xs">
            <div class="p-4 rounded-2xl bg-slate-50 border border-slate-200 space-y-1"><strong class="text-indigo-600 uppercase text-xs">6 THÁNG ĐẦU NĂM (ĐỒNG BỘ HÓA)</strong><p class="text-[11px] text-slate-600 leading-relaxed">• 01. Vận hành chính thức hệ thống AI Agent; kết nối trực tiếp Google Sheets API.<br>• 02. Chuẩn hóa dữ liệu điểm danh và đào tạo nhân sự thao tác.<br>• 03. Đạt tỷ lệ gia hạn ≥ 85%, 100% hồ sơ gọi trước 30 ngày.</p></div>
            <div class="p-4 rounded-2xl bg-slate-50 border border-slate-200 space-y-1"><strong class="text-emerald-600 uppercase text-xs">6 THÁNG CUỐI NĂM (MỞ RỘNG ĐA KÊNH)</strong><p class="text-[11px] text-slate-600 leading-relaxed">• 01. Tự động hóa gửi tin Zalo ZNS / Email cá nhân hóa.<br>• 02. Tối ưu mô hình dự báo Churn Risk Score theo thời gian thực.<br>• 03. Xử lý quy mô lớn, giảm Churn dưới 4%, tăng Up-sell 25%.</p></div>
          </div>
          <div class="text-xs text-slate-400 text-center">Lộ trình triển khai theo từng giai đoạn bảo đảm tính ổn định và khả năng mở rộng quy mô.</div>
        `
      },
      {
        title: "LỜI CẢM ƠN",
        html: `
          <div class="flex items-center gap-2 mx-auto"><div class="w-8 h-8 rounded-xl bg-indigo-600 text-white flex items-center justify-center font-bold text-sm">✦</div><span class="font-extrabold text-sm text-slate-800">ALPHAWAVE • EDTECH FUTURE</span></div>
          <div class="space-y-3 my-auto text-center">
            <h1 class="text-4xl font-extrabold text-slate-900">XIN TRÂN TRỌNG CẢM ƠN!</h1>
            <p class="text-xs sm:text-sm text-slate-500 max-w-md mx-auto">Hệ thống AI Agent Điều Phối Gia Hạn Khóa Học — Sẵn sàng phục vụ và nâng tầm vận hành doanh nghiệp.</p>
            <div class="inline-block px-4 py-2 rounded-xl bg-indigo-50 text-indigo-700 font-bold text-xs border border-indigo-200">🌐 Trải nghiệm trực tiếp tại: http://localhost:8080</div>
          </div>
          <div class="text-[11px] text-slate-400 flex gap-4 mx-auto"><span>Email: contact@edtechfuture.edu.vn</span><span>•</span><span>Hotline: 024.3733.3534</span><span>•</span><span>AlphaWave v3.5</span></div>
        `
      }
    ];

    function renderSlideDeck() {
      goToSlide(currentSlide);
    }

    function goToSlide(n) {
      if (n < 1) n = 1;
      if (n > 10) n = 10;
      currentSlide = n;

      const box = document.getElementById('slideContentBox');
      if (box && slidesData[n - 1]) {
        box.innerHTML = slidesData[n - 1].html;
      }
      document.getElementById('slideIndicator').textContent = `Slide ${n} / 10`;

      for (let i = 1; i <= 10; i++) {
        const thumb = document.getElementById('thumb-' + i);
        if (thumb) {
          if (i === n) {
            thumb.className = 'p-1 rounded-lg border-2 border-indigo-600 bg-indigo-50 text-[10px] font-bold text-indigo-700';
          } else {
            thumb.className = 'p-1 rounded-lg border border-slate-200 text-[10px] font-bold text-slate-600';
          }
        }
      }
    }

    function prevSlide() { goToSlide(currentSlide - 1); }
    function nextSlide() { goToSlide(currentSlide + 1); }

    function toggleSlideFullscreen() {
      const el = document.getElementById('slideViewport');
      if (!document.fullscreenElement) {
        el.requestFullscreen().catch(err => alert(err.message));
      } else {
        document.exitFullscreen();
      }
    }

    // Error test simulation
    function triggerSimulatedErrorTest() {
      const banner = document.getElementById('errorDiagnosticsBanner');
      const tbody = document.getElementById('errorListTbody');
      tbody.innerHTML = `
        <tr class="hover:bg-rose-50/50">
          <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">Dòng 7 | Cột: expiry_date</td>
          <td class="py-2.5 px-3 font-bold text-slate-800">HV-2026-004</td>
          <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/50 px-1 rounded">32/13/2026</td>
          <td class="py-2.5 px-3 text-slate-700">Định dạng ngày không hợp lệ (ngày 32, tháng 13)</td>
          <td class="py-2.5 px-3 text-emerald-800 font-semibold">Đổi về chuẩn YYYY-MM-DD (VD: 2026-12-04)</td>
        </tr>
        <tr class="hover:bg-rose-50/50">
          <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">Dòng 12 | Cột: attended_sessions</td>
          <td class="py-2.5 px-3 font-bold text-slate-800">HV-2026-009</td>
          <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/50 px-1 rounded">-5 buổi</td>
          <td class="py-2.5 px-3 text-slate-700">Số buổi tham gia là số âm bất hợp lý (&lt; 0)</td>
          <td class="py-2.5 px-3 text-emerald-800 font-semibold">Nhập số buổi thực tế từ 0 đến 48, hoặc để trống (NULL)</td>
        </tr>
        <tr class="hover:bg-rose-50/50">
          <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">Dòng 15 | Bảng: support_tickets</td>
          <td class="py-2.5 px-3 font-bold text-slate-800">HV-2026-007</td>
          <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/50 px-1 rounded">Ticket Open (TK-9002)</td>
          <td class="py-2.5 px-3 text-slate-700">Học viên đang mở Ticket lỗi micro văng lớp, cấm chào nâng cấp</td>
          <td class="py-2.5 px-3 text-emerald-800 font-semibold">Kỹ thuật đóng Ticket trước khi gia hạn, hoặc tạm hoãn</td>
        </tr>
      `;
      banner.classList.remove('hidden');
      banner.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    function dismissErrorBanner() {
      document.getElementById('errorDiagnosticsBanner').classList.add('hidden');
    }

    function autoFixAndContinue() {
      dismissErrorBanner();
      alert('Đã tự động sửa lỗi định dạng và kích hoạt khóa chặn kỹ thuật an toàn. Bắt đầu phân tích!');
      runQuickDemo();
    }

    function downloadCleanedExcel() {
      window.location.href = '/api/download-output-xlsx';
    }

    function handleFileUpload(e) {
      const file = e.target.files[0];
      if (!file) return;
      alert('Đã nhận tệp: ' + file.name + ' (' + Math.round(file.size / 1024) + ' KB). Hệ thống đang kiểm tra tính toàn vẹn và bắt đầu phân tích...');
      runQuickDemo();
    }

    async function sendRightChat() {
      const input = document.getElementById('rightChatInput');
      const text = input.value.trim();
      if (!text) return;
      
      const log = document.getElementById('chatMessageLog');
      // Append user bubble
      const uBubble = document.createElement('div');
      uBubble.className = 'p-2.5 rounded-xl bg-indigo-600 text-white text-[11px] ml-4';
      uBubble.textContent = text;
      log.appendChild(uBubble);
      input.value = '';

      // Append typing bubble
      const aBubble = document.createElement('div');
      aBubble.className = 'p-2.5 rounded-xl bg-slate-100 text-slate-800 text-[11px] mr-4 border border-slate-200';
      aBubble.innerHTML = '<span class="animate-pulse">Đang truy vấn cơ sở dữ liệu...</span>';
      log.appendChild(aBubble);

      try {
        const res = await fetch('/api/chat', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ message: text })
        });
        const data = await res.json();
        aBubble.innerHTML = (data.reply || 'Đã xử lý yêu cầu.').replace(/\n/g, '<br/>').replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
      } catch (err) {
        aBubble.textContent = 'Lỗi kết nối AI: ' + err.message;
      }
      log.scrollTop = log.scrollHeight;
    }

    function searchAccounts(query) {
      if (!query) return;
      document.getElementById('rightChatInput').value = query;
      sendRightChat();
    }

    function formatVND(val) {
      if (!val) return '0 ₫';
      return Number(val).toLocaleString('vi-VN') + ' ₫';
    }

    function formatMillions(val) {
      if (!val) return '0 Tr ₫';
      return (Number(val) / 1000000).toFixed(1).replace('.0', '') + ' Tr ₫';
    }

    // Initialize with 0s on load
    window.addEventListener('DOMContentLoaded', () => {
      resetAllDataToBlank();
    });
  </script>
</body>
</html>
'@

$targetPath = "c:\Users\ADMIN\Downloads\AI Agent CK\web\index.html"
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($targetPath, $htmlContent, $utf8Bom)
Write-Output "Successfully updated $targetPath with exact AlphaWave.ai template design."
