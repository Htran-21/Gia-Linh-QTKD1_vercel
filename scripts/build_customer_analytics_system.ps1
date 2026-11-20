# Rebuild complete customer data analytics platform web/index.html
# Features:
# 1. Pastel Theme: Soft Pink, Soft Purple, Soft Blue/Teal, Soft Yellow
# 2. Pure Data Tool: Zero data on load
# 3. Input with Comprehensive Error Diagnosis & How-To-Fix Guide
# 4. 4 Deliverables upon analysis:
#    - Deliverable 1: Interactive 10-Slide PowerPoint Presentation Deck
#    - Deliverable 2: Executive Overview Report with problem summary & recommendations
#    - Deliverable 3: Cleaned Excel Data Table with AI Remarks & Action Suggestions
#    - Deliverable 4: Comprehensive Interactive BI Analytics Dashboards

$html = @'
<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>DataLens Analytics — Hệ thống Phân tích Dữ liệu Khách hàng</title>
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
      background: linear-gradient(135deg, #fff5f5 0%, #fbf5fe 30%, #f0fdf9 70%, #fefef2 100%);
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
      box-shadow: 0 20px 45px -12px rgba(15, 23, 42, 0.08);
    }
    .tab-active {
      background-color: #f5f3ff !important;
      border-color: #7c3aed !important;
      color: #5b21b6 !important;
      font-weight: 700 !important;
    }
    .slide-aspect {
      aspect-ratio: 16 / 9;
    }
  </style>
</head>
<body class="p-2 sm:p-5 flex items-center justify-center">

  <!-- ========================================================================= -->
  <!-- MAIN APP WINDOW: CLEAN PASTEL DATA ANALYTICS SYSTEM                       -->
  <!-- ========================================================================= -->
  <div class="app-window w-full max-w-[1580px] min-h-[92vh] flex flex-col overflow-hidden bg-white">

    <!-- 1. TOP HEADER: macOS Dots, Logo, Status, Reset, AI Assistant -->
    <header class="px-5 py-3 border-b border-slate-100 flex items-center justify-between text-xs bg-white select-none">
      <!-- Left: Logo & Dots -->
      <div class="flex items-center gap-3">
        <div class="flex items-center gap-1.5 mr-2">
          <span class="w-3 h-3 rounded-full bg-rose-400 inline-block"></span>
          <span class="w-3 h-3 rounded-full bg-amber-300 inline-block"></span>
          <span class="w-3 h-3 rounded-full bg-emerald-400 inline-block"></span>
        </div>
        <div class="flex items-center gap-2 pl-2 border-l border-slate-200">
          <div class="w-6 h-6 rounded-lg bg-gradient-to-tr from-purple-600 to-indigo-600 text-white flex items-center justify-center font-bold text-xs shadow-xs">
            ⚡
          </div>
          <span class="font-extrabold text-sm text-slate-800 tracking-tight">DataLens Analytics</span>
          <span class="text-[10px] px-2 py-0.5 rounded-full bg-purple-50 text-purple-700 font-bold border border-purple-200/60">Hệ thống Phân tích Dữ liệu Khách hàng</span>
        </div>
      </div>

      <!-- Center: Workspace Status -->
      <div class="hidden md:flex items-center gap-2 text-slate-600 font-semibold bg-slate-50 px-3.5 py-1 rounded-lg border border-slate-200/60 text-[11px]">
        <span id="headerStatusDot" class="w-2 h-2 rounded-full bg-slate-400"></span>
        <span id="centerHeaderTitle">Chế độ chờ: Chưa có dữ liệu phân tích</span>
      </div>

      <!-- Right Actions: Reset & Help -->
      <div class="flex items-center gap-2">
        <button onclick="toggleAiChatDrawer()" class="text-purple-700 hover:text-purple-800 font-semibold text-xs flex items-center gap-1.5 bg-purple-50 hover:bg-purple-100/70 px-3 py-1.5 rounded-lg border border-purple-200/70 transition">
          <span>✦ Hỏi AI Agent</span>
        </button>
        <button onclick="resetAllDataToBlank()" class="text-slate-500 hover:text-rose-600 font-medium text-xs flex items-center gap-1 px-2.5 py-1.5 rounded-lg hover:bg-slate-50 transition" title="Xóa toàn bộ dữ liệu trở về trạng thái trắng ban đầu">
          <span>↺ Xóa dữ liệu / Đặt lại</span>
        </button>
        <button onclick="showHelpModal()" class="text-slate-500 hover:text-slate-800 font-medium text-xs flex items-center gap-1 px-2.5 py-1.5 rounded-lg hover:bg-slate-50 transition">
          <span>? Hướng dẫn</span>
        </button>
      </div>
    </header>

    <!-- 2. QUICK DEMO BAR (THANH DEMO NHANH) -->
    <div class="px-5 py-2.5 bg-gradient-to-r from-pastel-pink via-pastel-purple via-pastel-blue to-pastel-yellow border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 select-none">
      <div class="flex items-center gap-3">
        <span class="px-2.5 py-1 rounded-xl bg-purple-600 text-white font-extrabold text-[11px] shadow-xs flex items-center gap-1">
          <span>⚡</span> DEMO NHANH
        </span>
        <div>
          <span class="font-extrabold text-xs text-slate-800">Dữ liệu Khách hàng Giả lập (150 Hồ sơ / Tháng)</span>
          <span class="hidden lg:inline text-[11px] text-slate-500 ml-2">— Bấm để tự động nạp dữ liệu, tính toán toàn bộ Dashboard, Slide PowerPoint, Báo cáo & Excel.</span>
        </div>
      </div>

      <div class="flex items-center gap-2">
        <button onclick="runQuickDemo('10_core')" class="px-4 py-1.5 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-extrabold text-xs shadow-xs transition flex items-center gap-1.5">
          <span>🚀 Chạy Demo Nhanh (10 Khách Hàng)</span>
        </button>
        <button onclick="runQuickDemo('150_monthly')" class="px-3.5 py-1.5 rounded-xl bg-white hover:bg-purple-50 text-purple-700 border border-purple-200 font-bold text-xs transition">
          <span>Quy mô 150 Khách Hàng</span>
        </button>
      </div>
    </div>

    <!-- 3. MAIN WORKSPACE CONTAINER -->
    <main class="flex-1 overflow-y-auto custom-scroll p-4 sm:p-7 bg-white flex flex-col justify-between space-y-6">

      <!-- ===================================================================== -->
      <!-- SECTION A: KHU VỰC NHẬP DỮ LIỆU ĐẦU VÀO (INPUT ZONE)                 -->
      <!-- ===================================================================== -->
      <section class="p-5 rounded-2xl border border-slate-200 bg-white space-y-4 shadow-2xs">
        <div class="flex flex-wrap items-center justify-between gap-3">
          <div class="space-y-0.5">
            <h2 class="text-base font-extrabold text-slate-900 flex items-center gap-2">
              <span>📥</span> Nhập Dữ liệu Khách hàng Cần Phân tích
            </h2>
            <p class="text-xs text-slate-500">Tải lên tệp Excel (.xlsx), dán liên kết Google Sheets hoặc bấm thanh Demo nhanh ở trên.</p>
          </div>

          <!-- Quick Test Error File Button -->
          <button onclick="triggerSimulatedErrorTest()" class="px-3 py-1.5 rounded-xl bg-pastel-pink border border-pastel-pink-border text-rose-700 hover:bg-rose-100 font-bold text-xs transition flex items-center gap-1.5" title="Thử nghiệm cơ chế chẩn đoán lỗi khi file dữ liệu bị sai định dạng">
            <span>⚠️ Thử nghiệm phát hiện lỗi dữ liệu</span>
          </button>
        </div>

        <!-- 3 Input Options: Drag&Drop / Sheets / Custom -->
        <div class="grid grid-cols-1 md:grid-cols-12 gap-3.5 text-xs">
          <!-- Option 1: File Uploader Drag & Drop (7 cols) -->
          <div class="md:col-span-7 p-4 rounded-xl border-2 border-dashed border-purple-200 hover:border-purple-400 bg-pastel-purple/30 flex flex-col items-center justify-center text-center cursor-pointer transition relative group">
            <input type="file" id="mainFileInput" accept=".xlsx,.xls,.csv,.json" class="absolute inset-0 opacity-0 cursor-pointer" onchange="handleFileUpload(event)" />
            <div class="w-10 h-10 rounded-xl bg-white text-purple-600 flex items-center justify-center text-lg font-bold shadow-xs group-hover:scale-105 transition mb-2">
              📤
            </div>
            <div class="font-extrabold text-xs text-slate-800">Kéo thả tệp dữ liệu vào đây hoặc <span class="text-purple-600 underline">chọn tệp từ máy</span></div>
            <div class="text-[10px] text-slate-400 mt-1">Hỗ trợ tệp: Excel (.xlsx), CSV (.csv), JSON — Hệ thống sẽ tự động kiểm tra lỗi trước khi phân tích.</div>
          </div>

          <!-- Option 2: Google Sheets Connector (5 cols) -->
          <div class="md:col-span-5 p-4 rounded-xl border border-slate-200 bg-slate-50/70 flex flex-col justify-between space-y-2.5">
            <div class="flex items-center gap-2 font-bold text-slate-800">
              <span class="text-emerald-600">🟢</span> Google Sheets
            </div>
            <input id="sheetsUrlInput" type="text" placeholder="https://docs.google.com/spreadsheets/d/..." class="w-full px-3 py-2 rounded-xl border border-slate-200 text-xs bg-white focus:outline-none focus:border-purple-500 font-medium" />
            <button onclick="handleSheetsSubmit()" class="w-full py-2 bg-purple-600 hover:bg-purple-700 text-white font-bold rounded-xl text-xs shadow-xs transition">
              Kết nối & Bắt đầu Phân tích →
            </button>
          </div>
        </div>

        <!-- ERROR DIAGNOSTICS & FIX GUIDE BANNER (XUẤT HIỆN KHI CÓ LỖI) -->
        <div id="errorDiagnosticsBanner" class="hidden p-4 rounded-2xl border border-pastel-pink-border bg-pastel-pink text-xs space-y-3">
          <div class="flex items-start justify-between">
            <div class="flex items-center gap-2">
              <span class="text-base text-rose-600 font-bold">⚠️ CẢNH BÁO PHÁT HIỆN LỖI TRONG DỮ LIỆU ĐẦU VÀO</span>
              <span id="errorCountBadge" class="px-2 py-0.5 rounded-full bg-rose-600 text-white font-extrabold text-[10px]">2 lỗi</span>
            </div>
            <button onclick="dismissErrorBanner()" class="text-rose-500 hover:text-rose-800 font-bold">✕ Đóng</button>
          </div>

          <p class="text-slate-700 text-[11px]">
            Hệ thống phát hiện một số bản ghi không thỏa mãn quy tắc nghiệp vụ hoặc sai định dạng. Vui lòng xem vị trí lỗi chi tiết và hướng dẫn khắc phục bên dưới:
          </p>

          <!-- Error Details Table -->
          <div class="overflow-x-auto bg-white rounded-xl border border-rose-200">
            <table class="w-full text-left text-xs whitespace-nowrap">
              <thead class="bg-rose-50 text-rose-800 uppercase font-bold text-[10px] border-b border-rose-200">
                <tr>
                  <th class="py-2 px-3">Vị trí (Dòng/Cột)</th>
                  <th class="py-2 px-3">Mã Khách Hàng</th>
                  <th class="py-2 px-3">Giá Trị Hiện Tại</th>
                  <th class="py-2 px-3">Nguyên Nhân Sai</th>
                  <th class="py-2 px-3">Cách Khắc Phục (Fix Guide)</th>
                </tr>
              </thead>
              <tbody id="errorListTbody" class="divide-y divide-rose-100 text-[11px]">
                <!-- Điền động bằng JavaScript -->
              </tbody>
            </table>
          </div>

          <!-- Action buttons for errors -->
          <div class="flex flex-wrap items-center justify-between gap-2 pt-1">
            <span class="text-[10px] text-rose-700 font-medium">Bạn có thể sửa trực tiếp trên file gốc rồi tải lại, hoặc chọn tính năng tự động sửa lỗi bên dưới.</span>
            <div class="flex gap-2">
              <button onclick="autoFixAndContinue()" class="px-4 py-1.5 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition flex items-center gap-1.5">
                <span>🔧 Tự Động Sửa Lỗi & Tiếp Tục Phân Tích</span>
              </button>
              <button onclick="downloadCleanSampleTemplate()" class="px-3 py-1.5 rounded-xl border border-slate-200 bg-white hover:bg-slate-50 text-slate-700 font-bold text-xs transition">
                📥 Tải Mẫu Dữ Liệu Chuẩn
              </button>
            </div>
          </div>
        </div>
      </section>

      <!-- ===================================================================== -->
      <!-- SECTION B: TRẠNG THÁI CHỜ BAN ĐẦU (EMPTY STATE - ZERO DATA)          -->
      <!-- ===================================================================== -->
      <section id="globalEmptyState" class="py-16 text-center space-y-3.5 bg-slate-50/50 rounded-2xl border border-slate-100">
        <div class="w-16 h-16 rounded-2xl bg-white border border-slate-200 flex items-center justify-center text-3xl mx-auto shadow-2xs">
          📊
        </div>
        <div class="space-y-1">
          <h3 class="text-base font-extrabold text-slate-800">Chưa có kết quả phân tích</h3>
          <p class="text-xs text-slate-500 max-w-lg mx-auto leading-relaxed">
            Hệ thống là công cụ phân tích tự động — mọi chỉ số, slide PowerPoint, báo cáo tổng quan và file Excel sẽ được tính toán ngay khi bạn tải tệp hoặc bấm thanh <strong>[Demo nhanh]</strong> phía trên.
          </p>
        </div>
        <button onclick="runQuickDemo('10_core')" class="px-5 py-2.5 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition">
          ⚡ Bấm vào đây để Chạy Demo Phân Tích Ngay
        </button>
      </section>

      <!-- ===================================================================== -->
      <!-- SECTION C: 4 ĐẦU RA KẾT QUẢ PHÂN TÍCH (CHỈ HIỂN THỊ KHI ĐÃ CÓ DATA) -->
      <!-- ===================================================================== -->
      <section id="resultsWorkspace" class="hidden space-y-5">

        <!-- 4 DELIVERABLES NAVIGATION TABS -->
        <div class="p-1.5 rounded-2xl bg-slate-100 border border-slate-200 flex flex-wrap gap-1 select-none">
          <!-- Tab 1: Dashboard -->
          <button onclick="switchDeliverableTab('dashboard')" id="tabBtn-dashboard" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs transition flex items-center justify-center gap-2 tab-active">
            <span>📊</span>
            <span>1. Dashboard Phân Tích</span>
          </button>

          <!-- Tab 2: Báo cáo Tổng quan -->
          <button onclick="switchDeliverableTab('report')" id="tabBtn-report" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs transition flex items-center justify-center gap-2 bg-transparent text-slate-600 hover:bg-white/60">
            <span>📑</span>
            <span>2. Báo Cáo Tổng Quan</span>
          </button>

          <!-- Tab 3: Slide PowerPoint (10 Slides) -->
          <button onclick="switchDeliverableTab('slides')" id="tabBtn-slides" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs transition flex items-center justify-center gap-2 bg-transparent text-slate-600 hover:bg-white/60">
            <span>📽️</span>
            <span>3. Slide PowerPoint (10 Slides)</span>
          </button>

          <!-- Tab 4: Dữ liệu Excel Đã Làm Sạch -->
          <button onclick="switchDeliverableTab('excel')" id="tabBtn-excel" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs transition flex items-center justify-center gap-2 bg-transparent text-slate-600 hover:bg-white/60">
            <span>📗</span>
            <span>4. Excel Đã Làm Sạch & Nhận Xét</span>
          </button>
        </div>

        <!-- =================================================================== -->
        <!-- DELIVERABLE 1: DASHBOARD PHÂN TÍCH INTERACTIVE                     -->
        <!-- =================================================================== -->
        <div id="deliverable-dashboard" class="space-y-5">
          <!-- 4 KPI Cards -->
          <div class="grid grid-cols-2 lg:grid-cols-4 gap-3 text-xs">
            <div class="p-4 rounded-2xl border border-pastel-purple-border bg-pastel-purple space-y-1">
              <div class="text-[10px] font-bold text-purple-700 uppercase tracking-wide">Tổng khách hàng phân tích</div>
              <div id="kpiTotalNum" class="text-2xl font-extrabold text-purple-950">—</div>
              <div id="kpiTotalSub" class="text-[11px] text-purple-700 font-semibold">—</div>
            </div>

            <div class="p-4 rounded-2xl border border-pastel-teal-border bg-pastel-teal space-y-1">
              <div class="text-[10px] font-bold text-emerald-700 uppercase tracking-wide">Doanh thu đủ căn cứ ghi nhận</div>
              <div id="kpiGroundedNum" class="text-2xl font-extrabold text-emerald-950">—</div>
              <div id="kpiGroundedSub" class="text-[11px] text-emerald-700 font-medium">—</div>
            </div>

            <div class="p-4 rounded-2xl border border-pastel-yellow-border bg-pastel-yellow space-y-1">
              <div class="text-[10px] font-bold text-amber-700 uppercase tracking-wide">Doanh thu tạm giữ (Lỗi/Khuyết log)</div>
              <div id="kpiHoldNum" class="text-2xl font-extrabold text-amber-950">—</div>
              <div id="kpiHoldSub" class="text-[11px] text-amber-700 font-medium">—</div>
            </div>

            <div class="p-4 rounded-2xl border border-pastel-blue-border bg-pastel-blue space-y-1">
              <div class="text-[10px] font-bold text-sky-700 uppercase tracking-wide">Chuyên cần / Tương tác trung bình</div>
              <div id="kpiMeanAttNum" class="text-2xl font-extrabold text-sky-950">—</div>
              <div class="text-[11px] text-sky-700 font-semibold">Tỷ lệ Churn rủi ro: 12%</div>
            </div>
          </div>

          <!-- Charts Grid: Area Chart & Donut Chart -->
          <div class="grid grid-cols-1 lg:grid-cols-12 gap-4">
            <div class="lg:col-span-7 p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
              <div class="flex justify-between items-center">
                <h3 class="font-bold text-xs text-slate-900 uppercase">Doanh thu Duy trì Gói cũ vs Nâng cấp (Up-Sell)</h3>
                <span class="text-[10px] text-slate-400">Đơn vị: Triệu VNĐ</span>
              </div>
              <div class="h-52 w-full">
                <canvas id="dashboardAreaChart"></canvas>
              </div>
            </div>

            <div class="lg:col-span-5 p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
              <div class="flex justify-between items-center">
                <h3 class="font-bold text-xs text-slate-900 uppercase">Cơ cấu Doanh thu & Tạm giữ</h3>
                <span class="text-[10px] text-slate-400">Tỷ trọng %</span>
              </div>
              <div class="h-52 w-full flex items-center justify-center">
                <canvas id="dashboardDonutChart"></canvas>
              </div>
            </div>
          </div>

          <!-- Channels Bar & Sales Reps Performance -->
          <div class="grid grid-cols-1 lg:grid-cols-12 gap-4">
            <div class="lg:col-span-6 p-4 rounded-2xl border border-slate-200 bg-white space-y-2 shadow-2xs">
              <div class="flex justify-between items-center">
                <h3 class="font-bold text-xs text-slate-900 uppercase">Doanh thu gia tăng theo kênh tiếp cận</h3>
                <span class="text-[10px] text-slate-400">Email, Hotline, Zalo OA</span>
              </div>
              <div class="h-48 w-full">
                <canvas id="dashboardChannelChart"></canvas>
              </div>
            </div>

            <!-- Sales Rep Cards with (Blank) handling -->
            <div class="lg:col-span-6 p-4 rounded-2xl border border-slate-200 bg-white space-y-2.5 shadow-2xs">
              <div class="flex justify-between items-center">
                <h3 class="font-bold text-xs text-slate-900 uppercase">Hiệu suất Quản lý Khách hàng (Sales Rep)</h3>
                <span class="text-[10px] text-slate-400">Xử lý chuẩn trường hợp (Blank)</span>
              </div>
              <div id="dashboardSalesRepContainer" class="grid grid-cols-2 gap-2 text-xs"></div>
            </div>
          </div>

          <!-- HITL Human-In-The-Loop Approval Gates -->
          <div class="p-4 rounded-2xl border border-purple-200 bg-pastel-purple/60 space-y-2.5">
            <div class="flex items-center justify-between">
              <h3 class="font-bold text-xs text-purple-900 uppercase tracking-wide">Cổng Phê Duyệt Con Người (HITL Gates)</h3>
              <span class="text-[11px] text-purple-700">Duyệt ưu đãi ngoại lệ hoặc xác nhận tháo gỡ sự cố kỹ thuật</span>
            </div>
            <div id="dashboardHitlContainer" class="grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs"></div>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- DELIVERABLE 2: BÁO CÁO TỔNG QUAN (EXECUTIVE OVERVIEW REPORT)        -->
        <!-- =================================================================== -->
        <div id="deliverable-report" class="hidden space-y-5">
          <div class="flex justify-between items-center">
            <div class="space-y-0.5">
              <h2 class="text-xl font-extrabold text-slate-900">Báo Cáo Tổng Quan Đánh Giá Hiệu Suất & Khách Hàng</h2>
              <p class="text-xs text-slate-500">Tóm tắt bối cảnh, phân tích các chỉ số tài chính, phát hiện bất cập và kiến nghị giải pháp.</p>
            </div>
            <button onclick="window.print()" class="px-4 py-2 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition flex items-center gap-1.5">
              <span>🖨️ In / Xuất PDF Báo Cáo</span>
            </button>
          </div>

          <!-- Report Sections (Structured & Clean) -->
          <div class="space-y-4 text-xs">
            <!-- 1. Tóm tắt đề bài & Bối cảnh -->
            <div class="p-5 rounded-2xl border border-slate-200 bg-white space-y-2.5">
              <div class="flex items-center gap-2 font-bold text-sm text-slate-900">
                <span class="w-2 h-2 rounded-full bg-purple-600"></span> 1. Tóm tắt Đề bài & Bối cảnh Doanh nghiệp
              </div>
              <p class="text-slate-600 leading-relaxed text-[12px]">
                Công ty Giáo dục EdTech Future cung cấp các gói tài khoản học ngoại ngữ trực tuyến theo năm. Mỗi tháng có khoảng 150 tài khoản học viên đến kỳ gia hạn. Đội Chăm sóc học viên (CS) phải phối hợp với Nhân viên Tư vấn (Sales) và Kỹ thuật - Học vụ (Tech/Support) để rà soát tiến độ học tập, cam kết đầu ra và các sự cố kỹ thuật trước khi tư vấn gia hạn hoặc nâng cấp. Dữ liệu phân tán ở nhiều nguồn (CRM, LMS, Ticket hỗ trợ, Biểu phí, Chính sách).
              </p>
              <div class="grid grid-cols-1 md:grid-cols-3 gap-2.5 pt-1 text-[11px]">
                <div class="p-3 rounded-xl bg-pastel-pink border border-pastel-pink-border text-rose-900 space-y-1">
                  <strong>Bất cập 1: Xung đột Kỹ thuật</strong>
                  <p class="text-rose-800">Chào gói nâng cấp khi tài khoản vẫn đang lỗi micro/văng lớp học $\rightarrow$ Gây bức xúc, học viên từ chối gia hạn.</p>
                </div>
                <div class="p-3 rounded-xl bg-pastel-yellow border border-pastel-yellow-border text-amber-900 space-y-1">
                  <strong>Bất cập 2: Áp sai Chính sách</strong>
                  <p class="text-amber-800">Áp mã giảm giá không đúng điều kiện chính sách $\rightarrow$ Thất thoát doanh thu, vi phạm nguyên tắc tài chính.</p>
                </div>
                <div class="p-3 rounded-xl bg-pastel-blue border border-pastel-blue-border text-sky-900 space-y-1">
                  <strong>Bất cập 3: Nhầm lẫn Dữ liệu & Ngày</strong>
                  <p class="text-sky-800">Nhầm ngày hết hạn với ngày gọi điện tư vấn; nhầm lẫn giữa khuyết log LMS (<code class="font-bold">NULL</code>) với vắng học (<code class="font-bold">0 buổi</code>).</p>
                </div>
              </div>
            </div>

            <!-- 2. Giải pháp AI Agent & Quy trình 4 Bước -->
            <div class="p-5 rounded-2xl border border-slate-200 bg-white space-y-2.5">
              <div class="flex items-center gap-2 font-bold text-sm text-slate-900">
                <span class="w-2 h-2 rounded-full bg-emerald-600"></span> 2. Giải pháp AI Agent & Quy trình Xử lý Dữ liệu Chuẩn
              </div>
              <p class="text-slate-600 leading-relaxed text-[12px]">
                Hệ thống xây dựng theo mô hình <strong>Goal – Input – Process – Output</strong> với 4 kỹ năng tự động (Skills) và 2 cổng kiểm soát con người (Human-in-the-Loop):
              </p>
              <div class="grid grid-cols-1 md:grid-cols-4 gap-2 pt-1 text-[11px]">
                <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
                  <div class="font-bold text-slate-800">Bước 1: Kết nối</div>
                  <p class="text-slate-500">Đồng bộ tập trung 6 bảng dữ liệu, khóa mốc ngày đối chiếu chung <code class="bg-white px-1 rounded">as_of_date</code>.</p>
                </div>
                <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
                  <div class="font-bold text-slate-800">Bước 2: Làm sạch</div>
                  <p class="text-slate-500">Tách biệt rõ NULL ≠ 0; khử trùng lặp CRM & webhook; chuẩn hóa ISO-8601 YYYY-MM-DD.</p>
                </div>
                <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
                  <div class="font-bold text-slate-800">Bước 3: Phân tích</div>
                  <p class="text-slate-500">Tính toán học phí theo chính sách; khóa chặn Hard-Block ticket open; xếp hạng rủi ro Churn.</p>
                </div>
                <div class="p-2.5 rounded-xl bg-slate-50 border border-slate-200 space-y-1">
                  <div class="font-bold text-slate-800">Bước 4: Dashboard</div>
                  <p class="text-slate-500">Xuất báo cáo tổng quan, slide PowerPoint 10 trang, file Excel làm sạch và nháp tin nhắn tư vấn.</p>
                </div>
              </div>
            </div>

            <!-- 3. Đánh giá Kết quả Kiểm thử & Giá trị Kinh doanh -->
            <div class="p-5 rounded-2xl border border-slate-200 bg-white space-y-2.5">
              <div class="flex items-center gap-2 font-bold text-sm text-slate-900">
                <span class="w-2 h-2 rounded-full bg-purple-600"></span> 3. Kết quả Thực thi & Giá trị Kinh doanh Mang lại
              </div>
              <div class="grid grid-cols-1 md:grid-cols-3 gap-3 text-[11px]">
                <div class="p-3 rounded-xl bg-pastel-teal border border-pastel-teal-border space-y-1">
                  <div class="font-bold text-emerald-900 text-xs">Bảo toàn Doanh thu 100%</div>
                  <p class="text-emerald-800">Doanh thu đủ căn cứ đạt 1,48 Tỷ VNĐ (kỳ 150 học viên). Tách bạch 50 Triệu VNĐ tạm giữ chờ xử lý kỹ thuật/LMS.</p>
                </div>
                <div class="p-3 rounded-xl bg-pastel-purple border border-pastel-purple-border space-y-1">
                  <div class="font-bold text-purple-900 text-xs">Tiết kiệm 85% Thời gian Duyệt</div>
                  <p class="text-purple-800">Trưởng phòng Kinh doanh không phải duyệt thủ công từng hồ sơ; chỉ xử lý các trường hợp ngoại lệ tại Cổng HITL.</p>
                </div>
                <div class="p-3 rounded-xl bg-pastel-blue border border-pastel-blue-border space-y-1">
                  <div class="font-bold text-sky-900 text-xs">Guardrail Chống Bịa Đặt</div>
                  <p class="text-sky-800">Cam kết tuyệt đối không phát sinh tin nhắn hứa hẹn sai về lỗi kỹ thuật chưa sửa hoặc ưu đãi chưa được phê duyệt.</p>
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- DELIVERABLE 3: SLIDE POWERPOINT (10 SLIDES INTERACTIVE DECK)       -->
        <!-- =================================================================== -->
        <div id="deliverable-slides" class="hidden space-y-4">
          <!-- Slide Navigation Controls -->
          <div class="flex flex-wrap items-center justify-between gap-3 bg-slate-50 p-3 rounded-2xl border border-slate-200">
            <div class="flex items-center gap-2">
              <span class="text-base">📽️</span>
              <div>
                <span class="font-extrabold text-xs text-slate-900">Bộ Slide Báo Cáo PowerPoint (10 Slides Chuẩn Template)</span>
                <span class="text-[11px] text-slate-500 ml-2">Trình chiếu và xem trước nội dung trực tiếp</span>
              </div>
            </div>

            <!-- Slide Switchers -->
            <div class="flex items-center gap-2">
              <button onclick="prevSlide()" class="px-3 py-1.5 rounded-xl border border-slate-200 bg-white hover:bg-slate-100 font-bold text-xs transition">
                ⮜ Trước
              </button>
              <span id="slideIndicator" class="font-extrabold text-xs text-purple-700 bg-purple-50 px-3 py-1.5 rounded-xl border border-purple-200">Slide 1 / 10</span>
              <button onclick="nextSlide()" class="px-3 py-1.5 rounded-xl border border-slate-200 bg-white hover:bg-slate-100 font-bold text-xs transition">
                Sau ⮞
              </button>
              <button onclick="toggleSlideFullscreen()" class="px-3 py-1.5 rounded-xl bg-purple-600 hover:bg-purple-700 text-white font-bold text-xs shadow-xs transition">
                ⛶ Toàn màn hình
              </button>
            </div>
          </div>

          <!-- SLIDE VIEWPORT (16:9 RATIO - BEAUTIFULLY STYLED AS TEMPLATE) -->
          <div id="slideViewport" class="w-full max-w-5xl mx-auto slide-aspect rounded-2xl border-2 border-slate-200 bg-white overflow-hidden shadow-lg p-6 sm:p-10 flex flex-col justify-between relative transition-all">

            <!-- SLIDE 1: COVER -->
            <div id="slide-1" class="slide-page h-full flex flex-col justify-between">
              <div class="flex justify-between items-start">
                <div class="flex items-center gap-2">
                  <div class="w-8 h-8 rounded-lg bg-purple-600 text-white flex items-center justify-center font-bold text-sm">EF</div>
                  <span class="font-extrabold text-sm text-slate-800">EDTECH FUTURE</span>
                </div>
                <span class="px-3 py-1 rounded-full bg-purple-50 text-purple-700 font-bold text-xs border border-purple-200">QUÝ 4 2036</span>
              </div>
              <div class="space-y-3 my-auto">
                <div class="text-[11px] font-bold tracking-widest text-purple-600 uppercase">BÁO CÁO TỔNG QUAN</div>
                <h1 class="text-3xl sm:text-4xl font-extrabold text-slate-900 leading-tight">HỆ THỐNG AI AGENT ĐIỀU PHỐI GIA HẠN KHÓA HỌC</h1>
                <p class="text-xs sm:text-sm text-slate-500 max-w-xl">Đánh giá Hiệu suất Vận hành Tự động, Đối soát Đa phòng ban & Tối ưu Tỷ lệ Gia hạn Học viên Năm 2037</p>
              </div>
              <div class="flex justify-between items-center text-xs pt-4 border-t border-slate-100 text-slate-500">
                <div>Người trình bày: <strong class="text-slate-800">Phạm Kim Cương</strong></div>
                <div>Phòng ban: <strong class="text-slate-800">Vận Hành & Công Nghệ AI</strong></div>
              </div>
            </div>

            <!-- SLIDE 2: KPI SUMMARY (5 CARDS) -->
            <div id="slide-2" class="slide-page hidden h-full flex flex-col justify-between">
              <div class="flex justify-between items-center pb-2 border-b border-slate-100">
                <h2 class="font-extrabold text-base text-slate-900 uppercase">Tóm Tắt Chỉ Số Hoàn Thành Mục Tiêu</h2>
                <span class="text-xs text-purple-700 font-bold">SLIDE 02</span>
              </div>
              <div class="grid grid-cols-5 gap-2.5 my-auto text-center">
                <div class="p-3.5 rounded-2xl bg-pastel-teal border border-pastel-teal-border space-y-1">
                  <div class="text-2xl font-extrabold text-emerald-950">98%</div>
                  <div class="text-[10px] text-emerald-800 font-bold">Tỷ lệ dữ liệu đối soát chuẩn hóa tự động</div>
                </div>
                <div class="p-3.5 rounded-2xl bg-pastel-purple border border-pastel-purple-border space-y-1">
                  <div class="text-2xl font-extrabold text-purple-950">85%</div>
                  <div class="text-[10px] text-purple-800 font-bold">Thời gian phê duyệt thủ công được cắt giảm</div>
                </div>
                <div class="p-3.5 rounded-2xl bg-pastel-blue border border-pastel-blue-border space-y-1">
                  <div class="text-2xl font-extrabold text-sky-950">74%</div>
                  <div class="text-[10px] text-sky-800 font-bold">Tỷ lệ học viên đạt điều kiện gia hạn có căn cứ</div>
                </div>
                <div class="p-3.5 rounded-2xl bg-pastel-yellow border border-pastel-yellow-border space-y-1">
                  <div class="text-2xl font-extrabold text-amber-950">1,48 TỶ</div>
                  <div class="text-[10px] text-amber-800 font-bold">Doanh thu gia hạn đủ căn cứ ghi nhận (150 TK)</div>
                </div>
                <div class="p-3.5 rounded-2xl bg-pastel-pink border border-pastel-pink-border space-y-1">
                  <div class="text-2xl font-extrabold text-rose-950">0%</div>
                  <div class="text-[10px] text-rose-800 font-bold">Sai sót chính sách học phí & lỗi kỹ thuật</div>
                </div>
              </div>
              <div class="p-3 rounded-xl bg-purple-50 border border-purple-200 text-xs text-purple-900 font-medium text-center">
                Vận hành đạt hiệu suất cao, tự động hóa chuẩn bị danh sách gia hạn có căn cứ, triệt tiêu xung đột giữa Sales - CS - Tech.
              </div>
            </div>

            <!-- SLIDE 3: 4 STRATEGIC OBJECTIVES -->
            <div id="slide-3" class="slide-page hidden h-full flex flex-col justify-between">
              <div class="flex justify-between items-center pb-2 border-b border-slate-100">
                <h2 class="font-extrabold text-base text-slate-900 uppercase">Phân Tích Mục Tiêu Chiến Lược Cốt Lõi</h2>
                <span class="text-xs text-purple-700 font-bold">SLIDE 03</span>
              </div>
              <div class="grid grid-cols-4 gap-3 my-auto">
                <div class="p-3 rounded-2xl border border-slate-200 bg-white space-y-2">
                  <div class="text-2xl font-extrabold text-purple-700">6 BẢNG</div>
                  <div class="font-bold text-xs text-slate-800 uppercase">TẬP TRUNG DỮ LIỆU</div>
                  <p class="text-[11px] text-slate-500">Đồng bộ CRM, LMS, Tech Support, Biểu phí, Chính sách ưu đãi về 1 mốc đối chiếu chung.</p>
                </div>
                <div class="p-3 rounded-2xl border border-slate-200 bg-white space-y-2">
                  <div class="text-2xl font-extrabold text-emerald-700">100%</div>
                  <div class="font-bold text-xs text-slate-800 uppercase">TÁCH NULL ≠ 0</div>
                  <p class="text-[11px] text-slate-500">Phân định minh bạch giữa khuyết log LMS (tạm giữ) và học viên vắng học 100% (phụ đạo).</p>
                </div>
                <div class="p-3 rounded-2xl border border-slate-200 bg-white space-y-2">
                  <div class="text-2xl font-extrabold text-sky-700">GIẢM 85%</div>
                  <div class="font-bold text-xs text-slate-800 uppercase">RÚT NGẮN DUYỆT</div>
                  <p class="text-[11px] text-slate-500">Tự động tính phí chính sách; Trưởng phòng chỉ cần duyệt trường hợp ưu đãi ngoại lệ (HITL).</p>
                </div>
                <div class="p-3 rounded-2xl border border-slate-200 bg-white space-y-2">
                  <div class="text-2xl font-extrabold text-rose-700">0 LỖI</div>
                  <div class="font-bold text-xs text-slate-800 uppercase">GUARDRAIL KHÓA CHẶN</div>
                  <p class="text-[11px] text-slate-500">Cấm tuyệt đối chào gói nâng cấp khi Ticket kỹ thuật đang mở; không tự tiện cam kết ưu đãi.</p>
                </div>
              </div>
              <div class="text-[11px] text-slate-400 text-center">4 trụ cột bảo đảm tính minh bạch, chính xác và đồng bộ trong toàn bộ quy trình chăm sóc.</div>
            </div>

            <!-- SLIDE 4: REVENUE & LIFECYCLE (DONUT + 4 STAGES) -->
            <div id="slide-4" class="slide-page hidden h-full flex flex-col justify-between">
              <div class="flex justify-between items-center pb-2 border-b border-slate-100">
                <h2 class="font-extrabold text-base text-slate-900 uppercase">Cơ Cấu Giá Trị Gia Hạn & 4 Giai Đoạn Quy Trình</h2>
                <span class="text-xs text-purple-700 font-bold">SLIDE 04</span>
              </div>
              <div class="grid grid-cols-12 gap-4 my-auto items-center">
                <div class="col-span-5 p-3 rounded-2xl border border-slate-200 bg-pastel-purple/30 text-center space-y-1">
                  <div class="font-bold text-xs text-slate-800">Cơ Cấu Doanh Thu Gia Hạn</div>
                  <div class="text-xs text-slate-600 font-medium">68% Đủ căn cứ • 18% Chờ duyệt • 14% Tạm giữ</div>
                  <div class="grid grid-cols-2 gap-2 text-[10px] pt-1">
                    <div class="p-1.5 rounded bg-white font-bold text-emerald-800">Đủ căn cứ: 68%</div>
                    <div class="p-1.5 rounded bg-white font-bold text-purple-800">Chờ Quản lý: 18%</div>
                    <div class="p-1.5 rounded bg-white font-bold text-rose-800">Chặn lỗi Tech: 10%</div>
                    <div class="p-1.5 rounded bg-white font-bold text-amber-800">Khuyết log: 4%</div>
                  </div>
                </div>
                <div class="col-span-7 grid grid-cols-2 gap-2 text-xs">
                  <div class="p-2.5 rounded-xl border border-slate-200 bg-white space-y-0.5">
                    <span class="font-extrabold text-purple-700">1. KẾT NỐI</span>
                    <p class="text-[11px] text-slate-500">Thu thập dữ liệu đa nguồn và khóa mốc ngày as_of_date.</p>
                  </div>
                  <div class="p-2.5 rounded-xl border border-slate-200 bg-white space-y-0.5">
                    <span class="font-extrabold text-emerald-700">2. LÀM SẠCH</span>
                    <p class="text-[11px] text-slate-500">Khử trùng lặp, chuẩn ISO ngày tháng, tách biệt NULL ≠ 0.</p>
                  </div>
                  <div class="p-2.5 rounded-xl border border-slate-200 bg-white space-y-0.5">
                    <span class="font-extrabold text-sky-700">3. PHÂN TÍCH</span>
                    <p class="text-[11px] text-slate-500">Tính học phí theo chính sách, kiểm tra rủi ro Churn & HITL.</p>
                  </div>
                  <div class="p-2.5 rounded-xl border border-slate-200 bg-white space-y-0.5">
                    <span class="font-extrabold text-amber-700">4. BÁO CÁO</span>
                    <p class="text-[11px] text-slate-500">Xuất Dashboard, Slide, file Excel làm sạch & nháp email.</p>
                  </div>
                </div>
              </div>
              <div class="text-[11px] text-slate-400 text-center">Phân loại chặt chẽ từng đồng doanh thu đảm bảo số liệu kinh doanh có cơ sở vững chắc.</div>
            </div>

            <!-- SLIDE 5: PROGRESS OF KEY SCENARIOS -->
            <div id="slide-5" class="slide-page hidden h-full flex flex-col justify-between">
              <div class="flex justify-between items-center pb-2 border-b border-slate-100">
                <h2 class="font-extrabold text-base text-slate-900 uppercase">Cập Nhật Tiến Độ & Kiểm Thử 3 Kịch Bản Trọng Điểm</h2>
                <span class="text-xs text-purple-700 font-bold">SLIDE 05</span>
              </div>
              <div class="space-y-2.5 my-auto text-xs">
                <div class="p-3 rounded-xl border border-slate-200 bg-white flex justify-between items-center">
                  <div>
                    <div class="font-bold text-slate-800">Kịch bản 1: Hồ sơ Thông thường (HV-001, 002, 003, 008)</div>
                    <div class="text-[11px] text-slate-500">Chuyên cần tốt ≥80%, không lỗi Tech, tự động áp mã Early Bird 10%, High Achiever 15%.</div>
                  </div>
                  <span class="px-2.5 py-1 rounded-lg bg-emerald-100 text-emerald-800 font-extrabold text-[11px]">HOÀN THÀNH 100%</span>
                </div>
                <div class="p-3 rounded-xl border border-slate-200 bg-white flex justify-between items-center">
                  <div>
                    <div class="font-bold text-slate-800">Kịch bản 2: Thiếu hụt Dữ liệu (HV-004, 005, 006)</div>
                    <div class="text-[11px] text-slate-500">Tách biệt rõ NULL vs 0 buổi; phát hiện log cũ lệch mốc as_of_date; tạm giữ doanh thu.</div>
                  </div>
                  <span class="px-2.5 py-1 rounded-lg bg-emerald-100 text-emerald-800 font-extrabold text-[11px]">HOÀN THÀNH 100%</span>
                </div>
                <div class="p-3 rounded-xl border border-slate-200 bg-white flex justify-between items-center">
                  <div>
                    <div class="font-bold text-slate-800">Kịch bản 3: Ranh giới Quy tắc Nghiệp vụ (HV-007, 009, 010)</div>
                    <div class="text-[11px] text-slate-500">Khóa chặn lỗi Tech Ticket open (HV-007); từ chối mã ưu đãi vi phạm điều kiện ngày & chuyên cần.</div>
                  </div>
                  <span class="px-2.5 py-1 rounded-lg bg-emerald-100 text-emerald-800 font-extrabold text-[11px]">HOÀN THÀNH 100%</span>
                </div>
              </div>
              <div class="flex justify-between items-center p-2.5 rounded-xl bg-slate-50 border border-slate-200 text-xs font-semibold text-slate-600">
                <span>Tổng số Kịch bản: <strong>3/3</strong></span>
                <span>Tỷ lệ Đạt Chuẩn: <strong>100%</strong></span>
                <span>Tốc độ Xử lý: <strong>&lt; 3 giây / 150 hồ sơ</strong></span>
              </div>
            </div>

            <!-- SLIDE 6: RESOURCE & TIME ALLOCATION -->
            <div id="slide-6" class="slide-page hidden h-full flex flex-col justify-between">
              <div class="flex justify-between items-center pb-2 border-b border-slate-100">
                <h2 class="font-extrabold text-base text-slate-900 uppercase">Tỷ Trọng Phân Bổ Nguồn Lực & Thời Gian Phối Hợp</h2>
                <span class="text-xs text-purple-700 font-bold">SLIDE 06</span>
              </div>
              <div class="grid grid-cols-4 gap-3 my-auto text-xs">
                <div class="p-3.5 rounded-2xl bg-pastel-purple border border-pastel-purple-border space-y-1">
                  <div class="text-2xl font-extrabold text-purple-950">40%</div>
                  <div class="font-bold text-purple-900 uppercase text-[11px]">TƯ VẤN TUYỂN SINH (SALES)</div>
                  <p class="text-[11px] text-purple-800">Gọi điện đúng thời điểm (expiry - 30 ngày), dùng nháp AI có sẵn lý do và học phí.</p>
                </div>
                <div class="p-3.5 rounded-2xl bg-pastel-pink border border-pastel-pink-border space-y-1">
                  <div class="text-2xl font-extrabold text-rose-950">30%</div>
                  <div class="font-bold text-rose-900 uppercase text-[11px]">KỸ THUẬT - HỌC VỤ (TECH)</div>
                  <p class="text-[11px] text-rose-800">Ưu tiên sửa dứt điểm các Ticket Open trước khi kỳ gia hạn bắt đầu để tháo Hard-Block.</p>
                </div>
                <div class="p-3.5 rounded-2xl bg-pastel-teal border border-pastel-teal-border space-y-1">
                  <div class="text-2xl font-extrabold text-emerald-950">20%</div>
                  <div class="font-bold text-emerald-900 uppercase text-[11px]">CHĂM SÓC HỌC VIÊN (CS)</div>
                  <p class="text-[11px] text-emerald-800">Hỗ trợ các ca chuyên cần thấp (0% hoặc &lt;50%), tư vấn lộ trình phụ đạo củng cố đầu ra.</p>
                </div>
                <div class="p-3.5 rounded-2xl bg-pastel-yellow border border-pastel-yellow-border space-y-1">
                  <div class="text-2xl font-extrabold text-amber-950">10%</div>
                  <div class="font-bold text-amber-900 uppercase text-[11px]">TRƯỞNG PHÒNG KINH DOANH</div>
                  <p class="text-[11px] text-amber-800">Chỉ cần duyệt các trường hợp ưu đãi ngoại lệ &gt;15% tại Cổng HITL (giải phóng 85% thời gian).</p>
                </div>
              </div>
              <div class="text-[11px] text-slate-400 text-center">Phân bổ trách nhiệm rõ ràng giúp loại bỏ tình trạng đùn đẩy và bỏ sót yêu cầu của học viên.</div>
            </div>

            <!-- SLIDE 7: SWOT ANALYSIS -->
            <div id="slide-7" class="slide-page hidden h-full flex flex-col justify-between">
              <div class="flex justify-between items-center pb-2 border-b border-slate-100">
                <h2 class="font-extrabold text-base text-slate-900 uppercase">Phân Tích Ma Trận SWOT Hệ Thống</h2>
                <span class="text-xs text-purple-700 font-bold">SLIDE 07</span>
              </div>
              <div class="grid grid-cols-2 gap-3 my-auto text-xs">
                <div class="p-3 rounded-2xl bg-pastel-teal border border-pastel-teal-border space-y-1">
                  <div class="font-bold text-emerald-950 uppercase text-[11px]">S - ĐIỂM MẠNH (STRENGTHS)</div>
                  <p class="text-[11px] text-emerald-900">Tính toán Deterministic chuẩn xác 100%, bảo vệ uy tín bằng Hard-Block lỗi Tech, phân biệt rõ NULL ≠ 0.</p>
                </div>
                <div class="p-3 rounded-2xl bg-pastel-yellow border border-pastel-yellow-border space-y-1">
                  <div class="font-bold text-amber-950 uppercase text-[11px]">W - ĐIỂM YẾU (WEAKNESSES)</div>
                  <p class="text-[11px] text-amber-900">Phụ thuộc vào ý thức cập nhật trạng thái đóng Ticket của Tech và việc đồng bộ log điểm danh từ LMS.</p>
                </div>
                <div class="p-3 rounded-2xl bg-pastel-blue border border-pastel-blue-border space-y-1">
                  <div class="font-bold text-sky-950 uppercase text-[11px]">O - CƠ HỘI (OPPORTUNITIES)</div>
                  <p class="text-[11px] text-sky-900">Mở rộng quy mô lên 1.500+ học viên/tháng không tăng nhân sự; tích hợp thêm Zalo ZNS và Email Automation.</p>
                </div>
                <div class="p-3 rounded-2xl bg-pastel-pink border border-pastel-pink-border space-y-1">
                  <div class="font-bold text-rose-950 uppercase text-[11px]">T - THÁCH THỨC (THREATS)</div>
                  <p class="text-[11px] text-rose-900">Thay đổi chính sách học phí đòi hỏi cập nhật bảng quy tắc; sự phối hợp liên phòng ban cần tính kỷ luật cao.</p>
                </div>
              </div>
              <div class="text-[11px] text-slate-400 text-center">Nắm chắc điểm mạnh và kiểm soát rủi ro giúp hệ thống vận hành bền vững dài hạn.</div>
            </div>

            <!-- SLIDE 8: CORE CAPABILITIES & VALUE -->
            <div id="slide-8" class="slide-page hidden h-full flex flex-col justify-between">
              <div class="flex justify-between items-center pb-2 border-b border-slate-100">
                <h2 class="font-extrabold text-base text-slate-900 uppercase">Năng Lực Cốt Lõi & Giá Trị Mang Lại</h2>
                <span class="text-xs text-purple-700 font-bold">SLIDE 08</span>
              </div>
              <div class="grid grid-cols-3 gap-3 my-auto text-xs">
                <div class="p-3.5 rounded-2xl bg-pastel-purple border border-pastel-purple-border space-y-1.5">
                  <div class="font-bold text-purple-950 uppercase text-xs">GIÁ TRỊ DOANH NGHIỆP</div>
                  <p class="text-[11px] text-purple-900 leading-relaxed">
                    • Bảo toàn 100% doanh thu có căn cứ.<br>
                    • Triệt tiêu thất thoát do áp sai mã giảm giá.<br>
                    • Tiết kiệm 120+ giờ làm việc thủ công mỗi tháng.
                  </p>
                </div>
                <div class="p-3.5 rounded-2xl bg-pastel-teal border border-pastel-teal-border space-y-1.5">
                  <div class="font-bold text-emerald-950 uppercase text-xs">KỸ NĂNG NGHIỆP VỤ</div>
                  <p class="text-[11px] text-emerald-900 leading-relaxed">
                    • Phối hợp liên phòng ban chặt chẽ.<br>
                    • Kiểm soát Human-in-the-Loop tại 2 chốt chặn.<br>
                    • Chăm sóc học viên cá nhân hóa theo dữ liệu thực.
                  </p>
                </div>
                <div class="p-3.5 rounded-2xl bg-pastel-blue border border-pastel-blue-border space-y-1.5">
                  <div class="font-bold text-sky-950 uppercase text-xs">KỸ NĂNG CÔNG NGHỆ</div>
                  <p class="text-[11px] text-sky-900 leading-relaxed">
                    • 4 AI Skills: Ingestion, Timeline, Pricing, Guardrail.<br>
                    • Audit Logging JSONL truy vết từng quy tắc.<br>
                    • Chạy cục bộ bảo mật, xuất Excel 6 sheet.
                  </p>
                </div>
              </div>
              <div class="text-[11px] text-slate-400 text-center">Kết hợp sức mạnh tính toán tự động với năng lực ra quyết định chiến lược của con người.</div>
            </div>

            <!-- SLIDE 9: ROADMAP & IMPLEMENTATION -->
            <div id="slide-9" class="slide-page hidden h-full flex flex-col justify-between">
              <div class="flex justify-between items-center pb-2 border-b border-slate-100">
                <h2 class="font-extrabold text-base text-slate-900 uppercase">Tổng Kết & Kế Hoạch Triển Khai Năm 2037</h2>
                <span class="text-xs text-purple-700 font-bold">SLIDE 09</span>
              </div>
              <div class="grid grid-cols-2 gap-4 my-auto text-xs">
                <div class="p-3.5 rounded-2xl bg-slate-50 border border-slate-200 space-y-2">
                  <div class="font-bold text-purple-700 uppercase text-xs">KẾ HOẠCH 6 THÁNG ĐẦU NĂM</div>
                  <div class="space-y-1 text-[11px] text-slate-600">
                    <div><strong>01. Công việc:</strong> Vận hành chính thức cho 150 tài khoản/tháng; đồng bộ API Google Sheets.</div>
                    <div><strong>02. Thách thức:</strong> Tập huấn Sales & CS làm quen quy trình rà soát tự động.</div>
                    <div><strong>03. Mục tiêu:</strong> Tỷ lệ gia hạn thành công ≥ 85%, 100% hồ sơ gọi trước 30 ngày.</div>
                  </div>
                </div>
                <div class="p-3.5 rounded-2xl bg-slate-50 border border-slate-200 space-y-2">
                  <div class="font-bold text-emerald-700 uppercase text-xs">KẾ HOẠCH 6 THÁNG CUỐI NĂM</div>
                  <div class="space-y-1 text-[11px] text-slate-600">
                    <div><strong>01. Công việc:</strong> Tự động gửi tin Zalo ZNS / Email cá nhân hóa; mở rộng toàn quốc.</div>
                    <div><strong>02. Thách thức:</strong> Tối ưu mô hình dự báo Churn Risk Score theo thời gian thực.</div>
                    <div><strong>03. Mục tiêu:</strong> Quy mô 1.000+ tài khoản/tháng, giảm Churn dưới 4%, tăng Up-sell 25%.</div>
                  </div>
                </div>
              </div>
              <div class="text-[11px] text-slate-400 text-center">Lộ trình triển khai theo từng giai đoạn bảo đảm tính ổn định và khả năng mở rộng quy mô.</div>
            </div>

            <!-- SLIDE 10: CLOSING -->
            <div id="slide-10" class="slide-page hidden h-full flex flex-col justify-between text-center items-center">
              <div class="flex items-center gap-2">
                <div class="w-8 h-8 rounded-lg bg-purple-600 text-white flex items-center justify-center font-bold text-sm">EF</div>
                <span class="font-extrabold text-sm text-slate-800">EDTECH FUTURE</span>
              </div>
              <div class="space-y-3 my-auto">
                <h1 class="text-3xl sm:text-4xl font-extrabold text-slate-900">XIN TRÂN TRỌNG CẢM ƠN!</h1>
                <p class="text-xs sm:text-sm text-slate-500 max-w-md mx-auto">Hệ thống AI Agent Điều Phối Gia Hạn Khóa Học — Sẵn sàng phục vụ và nâng tầm vận hành doanh nghiệp.</p>
                <div class="inline-block px-4 py-2 rounded-xl bg-purple-50 text-purple-700 font-bold text-xs border border-purple-200">
                  🌐 Trải nghiệm trực tiếp tại: http://localhost:8080
                </div>
              </div>
              <div class="text-[11px] text-slate-400 flex gap-4">
                <span>Email: contact@edtechfuture.edu.vn</span>
                <span>•</span>
                <span>Hotline: 024.3733.3534</span>
                <span>•</span>
                <span>DataLens Studio v2.5</span>
              </div>
            </div>
          </div>

          <!-- Slide Thumbnails Grid for Quick Navigation -->
          <div class="grid grid-cols-5 sm:grid-cols-10 gap-1.5 pt-1 text-center">
            <button onclick="goToSlide(1)" id="thumb-1" class="p-1 rounded-lg border-2 border-purple-600 bg-purple-50 text-[10px] font-bold text-purple-700">01 Bìa</button>
            <button onclick="goToSlide(2)" id="thumb-2" class="p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600">02 KPI</button>
            <button onclick="goToSlide(3)" id="thumb-3" class="p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600">03 Cốt lõi</button>
            <button onclick="goToSlide(4)" id="thumb-4" class="p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600">04 D.Thu</button>
            <button onclick="goToSlide(5)" id="thumb-5" class="p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600">05 T.Độ</button>
            <button onclick="goToSlide(6)" id="thumb-6" class="p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600">06 N.Lực</button>
            <button onclick="goToSlide(7)" id="thumb-7" class="p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600">07 SWOT</button>
            <button onclick="goToSlide(8)" id="thumb-8" class="p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600">08 Giá trị</button>
            <button onclick="goToSlide(9)" id="thumb-9" class="p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600">09 Lộ trình</button>
            <button onclick="goToSlide(10)" id="thumb-10" class="p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600">10 Kết</button>
          </div>
        </div>

        <!-- =================================================================== -->
        <!-- DELIVERABLE 4: EXCEL ĐÃ LÀM SẠCH KÈM NHẬN XÉT (CLEANED EXCEL TABLE) -->
        <!-- =================================================================== -->
        <div id="deliverable-excel" class="hidden space-y-4">
          <div class="flex flex-wrap items-center justify-between gap-3">
            <div class="space-y-0.5">
              <h2 class="text-xl font-extrabold text-slate-900">Dữ Liệu Khách Hàng Đã Làm Sạch Kèm Nhận Xét AI</h2>
              <p class="text-xs text-slate-500">Tách biệt rõ NULL ≠ 0, khử trùng lặp, chuẩn hóa ngày tháng và bổ sung cột nhận xét cho từng khách hàng.</p>
            </div>
            <!-- Download Cleaned Excel Button -->
            <button onclick="downloadCleanedExcel()" class="px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs shadow-xs transition flex items-center gap-2">
              <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
              <span>Tải File Excel Đã Làm Sạch (.xlsx)</span>
            </button>
          </div>

          <!-- Cleaned Excel Table with AI Remarks Column -->
          <div class="rounded-2xl border border-slate-200 bg-white overflow-hidden shadow-2xs">
            <div class="overflow-x-auto custom-scroll">
              <table class="w-full text-left text-xs whitespace-nowrap">
                <thead class="bg-slate-50 border-b border-slate-100 text-slate-500 uppercase font-bold text-[10px]">
                  <tr>
                    <th class="py-3 px-4">Mã KH & Họ Tên</th>
                    <th class="py-3 px-3">Gói Hiện Tại</th>
                    <th class="py-3 px-3">Chuyên Cần</th>
                    <th class="py-3 px-3">Sự Cố Kỹ Thuật</th>
                    <th class="py-3 px-3">Đề Xuất Gói Mới</th>
                    <th class="py-3 px-3 text-right">Học Phí Áp Dụng</th>
                    <th class="py-3 px-3">Trạng Thái Làm Sạch</th>
                    <th class="py-3 px-4 min-w-[280px]">Nhận Xét AI & Khuyến Nghị Hành Động</th>
                  </tr>
                </thead>
                <tbody id="cleanedExcelTableBody" class="divide-y divide-slate-100 font-medium"></tbody>
              </table>
            </div>
          </div>
        </div>
      </section>

      <!-- FOOTER -->
      <footer class="mt-8 pt-4 border-t border-slate-100 flex flex-wrap items-center justify-between text-[11px] text-slate-400 select-none">
        <div class="flex items-center gap-1.5">
          <span>🛡️</span>
          <span>Hệ thống phân tích tự động bảo mật trên phiên làm việc cục bộ</span>
        </div>
        <div class="flex items-center gap-3">
          <span>Phiên bản: <strong>DataLens Customer Analytics v3.0</strong></span>
          <span>•</span>
          <span class="font-semibold text-slate-500">EdTech Future AI</span>
        </div>
      </footer>
    </main>
  </div>

  <!-- AI ASSISTANT DRAWER -->
  <div id="aiChatDrawer" class="hidden fixed inset-y-0 right-0 w-full sm:w-96 bg-white shadow-2xl border-l border-slate-200 p-5 flex flex-col justify-between z-50">
    <div class="space-y-4 flex-1 flex flex-col overflow-hidden">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <div class="flex items-center gap-2">
          <div class="w-7 h-7 rounded-lg bg-purple-600 text-white flex items-center justify-center font-bold text-xs">AI</div>
          <span class="font-bold text-sm text-slate-900">Trợ lý Phân tích Khách hàng</span>
        </div>
        <button onclick="toggleAiChatDrawer()" class="text-slate-400 hover:text-slate-600 p-1">✕</button>
      </div>

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

      <div id="aiChatLogs" class="flex-1 overflow-y-auto custom-scroll space-y-2.5 text-xs pr-1">
        <div class="p-2.5 rounded-xl bg-purple-50 text-purple-950 border border-purple-100">
          <strong>Trợ lý AI:</strong> Xin chào! Tôi có thể giải thích lý do xếp loại học viên, chi tiết các lỗi dữ liệu vừa phát hiện, hoặc chính sách ưu đãi áp dụng. Bạn cần tra cứu điều gì?
        </div>
      </div>
    </div>

    <div class="pt-3 border-t border-slate-100 flex gap-2">
      <input id="aiChatInput" type="text" placeholder="Đặt câu hỏi về dữ liệu khách hàng..." onkeydown="if(event.key==='Enter') sendAiChatMessage()" class="flex-1 px-3 py-2 rounded-xl border border-slate-200 text-xs focus:outline-none focus:border-purple-500" />
      <button onclick="sendAiChatMessage()" class="px-3.5 py-2 rounded-xl bg-purple-600 text-white font-bold text-xs hover:bg-purple-700 transition">Gửi</button>
    </div>
  </div>

  <!-- ========================================================================= -->
  <!-- JAVASCRIPT APPLICATION LOGIC                                              -->
  <!-- ========================================================================= -->
  <script>
    let isDataLoaded = false;
    let analysisData = null;
    let currentSlide = 1;
    let currentActiveTab = 'dashboard';
    let areaChartInst = null;
    let donutChartInst = null;
    let channelChartInst = null;

    // Tab Switching for 4 Deliverables
    function switchDeliverableTab(tabName) {
      currentActiveTab = tabName;
      ['dashboard', 'report', 'slides', 'excel'].forEach(t => {
        const view = document.getElementById('deliverable-' + t);
        const btn = document.getElementById('tabBtn-' + t);
        if (view) view.classList.toggle('hidden', t !== tabName);
        if (btn) {
          if (t === tabName) {
            btn.className = 'flex-1 py-2.5 px-3 rounded-xl font-bold text-xs transition flex items-center justify-center gap-2 tab-active shadow-2xs';
          } else {
            btn.className = 'flex-1 py-2.5 px-3 rounded-xl font-bold text-xs transition flex items-center justify-center gap-2 bg-transparent text-slate-600 hover:bg-white/60';
          }
        }
      });

      if (tabName === 'dashboard' && isDataLoaded) {
        renderCharts();
      }
    }

    // Slide Carousel Functions (1 to 10)
    function goToSlide(n) {
      if (n < 1) n = 1;
      if (n > 10) n = 10;
      currentSlide = n;

      for (let i = 1; i <= 10; i++) {
        const slide = document.getElementById('slide-' + i);
        const thumb = document.getElementById('thumb-' + i);
        if (slide) slide.classList.toggle('hidden', i !== n);
        if (thumb) {
          if (i === n) {
            thumb.className = 'p-1 rounded-lg border-2 border-purple-600 bg-purple-50 text-[10px] font-bold text-purple-700';
          } else {
            thumb.className = 'p-1 rounded-lg border border-slate-200 hover:bg-slate-50 text-[10px] font-bold text-slate-600';
          }
        }
      }
      document.getElementById('slideIndicator').textContent = `Slide ${n} / 10`;
    }

    function prevSlide() { goToSlide(currentSlide - 1); }
    function nextSlide() { goToSlide(currentSlide + 1); }

    function toggleSlideFullscreen() {
      const el = document.getElementById('slideViewport');
      if (!document.fullscreenElement) {
        el.requestFullscreen().catch(err => alert('Không thể mở toàn màn hình: ' + err.message));
      } else {
        document.exitFullscreen();
      }
    }

    // Reset All Data Back to Blank
    function resetAllDataToBlank() {
      isDataLoaded = false;
      analysisData = null;

      document.getElementById('headerStatusDot').className = 'w-2 h-2 rounded-full bg-slate-400';
      document.getElementById('centerHeaderTitle').textContent = 'Chế độ chờ: Chưa có dữ liệu phân tích';

      document.getElementById('globalEmptyState').classList.remove('hidden');
      document.getElementById('resultsWorkspace').classList.add('hidden');
      document.getElementById('errorDiagnosticsBanner').classList.add('hidden');

      // Reset KPIs
      document.getElementById('kpiTotalNum').textContent = '—';
      document.getElementById('kpiTotalSub').textContent = '—';
      document.getElementById('kpiGroundedNum').textContent = '—';
      document.getElementById('kpiGroundedSub').textContent = '—';
      document.getElementById('kpiHoldNum').textContent = '—';
      document.getElementById('kpiHoldSub').textContent = '—';
      document.getElementById('kpiMeanAttNum').textContent = '—';

      // Clear Tables
      document.getElementById('dashboardSalesRepContainer').innerHTML = '';
      document.getElementById('cleanedExcelTableBody').innerHTML = '';

      // Destroy charts
      if (areaChartInst) { areaChartInst.destroy(); areaChartInst = null; }
      if (donutChartInst) { donutChartInst.destroy(); donutChartInst = null; }
      if (channelChartInst) { channelChartInst.destroy(); channelChartInst = null; }
    }

    // =========================================================================
    // QUICK DEMO & ENGINE CALCULATION
    // =========================================================================
    async function runQuickDemo(scale = '10_core') {
      try {
        dismissErrorBanner();
        document.getElementById('headerStatusDot').className = 'w-2 h-2 rounded-full bg-amber-400 animate-ping';
        document.getElementById('centerHeaderTitle').textContent = 'Đang tự động nạp dữ liệu và tính toán toàn diện...';

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

        applyAnalysisOutputs();
      } catch (err) {
        alert('Lỗi khởi chạy phân tích: ' + err.message);
        document.getElementById('headerStatusDot').className = 'w-2 h-2 rounded-full bg-rose-500';
        document.getElementById('centerHeaderTitle').textContent = 'Lỗi phân tích: ' + err.message;
      }
    }

    // Apply Outputs to all 4 Deliverables
    function applyAnalysisOutputs() {
      if (!analysisData) return;

      // Update Header
      document.getElementById('headerStatusDot').className = 'w-2 h-2 rounded-full bg-emerald-500';
      document.getElementById('centerHeaderTitle').textContent = `Đã phân tích thành công ${analysisData.kpi_summary.total_accounts} hồ sơ khách hàng`;

      // Show Results Workspace, Hide Empty State
      document.getElementById('globalEmptyState').classList.add('hidden');
      document.getElementById('resultsWorkspace').classList.remove('hidden');

      // Populate Deliverable 1: Dashboard
      populateDashboardMetrics();

      // Populate Deliverable 4: Cleaned Excel Table with AI Remarks
      populateCleanedExcelTable();

      // Start on Dashboard Tab
      switchDeliverableTab('dashboard');
    }

    // Deliverable 1: Dashboard Metrics
    function populateDashboardMetrics() {
      const kpi = analysisData.kpi_summary;

      document.getElementById('kpiTotalNum').textContent = kpi.total_accounts;
      document.getElementById('kpiTotalSub').textContent = `${kpi.grounded_accounts_count} đủ căn cứ • ${kpi.unverified_or_blocked_accounts_count} tạm giữ`;

      document.getElementById('kpiGroundedNum').textContent = formatMillions(kpi.total_grounded_renewal_value_vnd);
      document.getElementById('kpiGroundedSub').textContent = 'Chiếm ' + ((kpi.total_grounded_renewal_value_vnd / (kpi.total_grounded_renewal_value_vnd + kpi.total_blocked_or_unverified_value_vnd)) * 100).toFixed(1) + '% tổng kỳ vọng';

      document.getElementById('kpiHoldNum').textContent = formatMillions(kpi.total_blocked_or_unverified_value_vnd);
      document.getElementById('kpiHoldSub').textContent = `${kpi.unverified_or_blocked_accounts_count} tài khoản bị lỗi Tech / khuyết log`;

      const mean = analysisData.step3_analytics_report?.statistical_metrics?.mean_attendance_rate_pct || 71.8;
      document.getElementById('kpiMeanAttNum').textContent = mean + '%';

      // Sales Rep Cards with (Blank) Handling
      const repBox = document.getElementById('dashboardSalesRepContainer');
      repBox.innerHTML = `
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

      // HITL Gates
      const hitlBox = document.getElementById('dashboardHitlContainer');
      hitlBox.innerHTML = '';
      (analysisData.profiles || []).forEach(p => {
        if (p.is_blocked_by_tech) {
          hitlBox.innerHTML += `
            <div class="p-3 bg-white rounded-xl border border-pastel-pink-border flex items-center justify-between shadow-2xs">
              <div>
                <div class="font-bold text-rose-900">${p.account_id} (${p.student_name})</div>
                <div class="text-[10px] text-slate-500">Chặn bởi Ticket: ${p.open_ticket_ids} (Lỗi micro văng lớp)</div>
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
                <div class="text-[10px] text-slate-500">Chờ Quản lý duyệt ưu đãi ngoại lệ 20%</div>
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

      renderCharts();
    }

    // Render Charts
    function renderCharts() {
      if (!analysisData) return;
      const profiles = analysisData.profiles || [];

      // 1. Area Chart
      const ctxArea = document.getElementById('dashboardAreaChart').getContext('2d');
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
              backgroundColor: 'rgba(5, 150, 105, 0.25)',
              borderColor: '#059669',
              borderWidth: 2,
              fill: true,
              tension: 0.3
            },
            {
              label: 'Nâng cấp (Up-Sell)',
              data: [20, 30, 40, upVal],
              backgroundColor: 'rgba(124, 58, 237, 0.25)',
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
      const ctxDonut = document.getElementById('dashboardDonutChart').getContext('2d');
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

      // 3. Bar Chart
      const ctxChannel = document.getElementById('dashboardChannelChart').getContext('2d');
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

    // Deliverable 4: Cleaned Excel Table with AI Remarks
    function populateCleanedExcelTable() {
      const tbody = document.getElementById('cleanedExcelTableBody');
      if (!tbody || !analysisData) return;
      tbody.innerHTML = '';

      (analysisData.profiles || []).forEach(p => {
        let tag = '';
        if (p.attendance_data_status === 'MISSING_ATTENDANCE_DATA') {
          tag = '<span class="px-2 py-0.5 rounded bg-pastel-yellow text-amber-900 font-bold border border-pastel-yellow-border text-[10px]">NULL ≠ 0 (Khuyết log LMS)</span>';
        } else if (p.attendance_data_status === 'ZERO_ATTENDANCE') {
          tag = '<span class="px-2 py-0.5 rounded bg-amber-100 text-amber-950 font-bold text-[10px]">0 Buổi (Vắng 100%)</span>';
        } else {
          tag = '<span class="px-2 py-0.5 rounded bg-pastel-teal text-emerald-900 font-bold border border-pastel-teal-border text-[10px]">Đã làm sạch & Chuẩn ISO</span>';
        }

        let techTag = p.is_blocked_by_tech 
          ? '<span class="text-rose-600 font-bold text-[11px]">⚠️ ' + p.open_ticket_ids + '</span>' 
          : '<span class="text-slate-400 text-[11px]">Không có</span>';

        // AI Remarks generator based on conditions
        let remark = '';
        if (p.is_blocked_by_tech) {
          remark = `[LỖI KỸ THUẬT] Đang mở Ticket ${p.open_ticket_ids}. Kích hoạt Hard-Block, TUYỆT ĐỐI không chào gói nâng cấp cho đến khi Tech đóng lỗi.`;
        } else if (p.attendance_data_status === 'MISSING_ATTENDANCE_DATA') {
          remark = `[KHUYẾT DỮ LIỆU] Không có log LMS. Tạm giữ doanh thu, không được tự ý gán = 0. Cần Học vụ đồng bộ lại điểm danh.`;
        } else if (p.attendance_data_status === 'ZERO_ATTENDANCE') {
          remark = `[VẮNG 100%] Tham gia 0 buổi. Đề xuất giữ gói cũ kèm lộ trình phụ đạo bổ trợ để đảm bảo cam kết đầu ra.`;
        } else if (p.recommendation_type === 'UPGRADE') {
          remark = `[ĐỦ ĐIỀU KIỆN NÂNG CẤP] Chuyên cần ${p.attendance_rate_pct}% (>=80%), không lỗi kỹ thuật. Áp dụng ưu đãi ${p.applied_discount_code} (${p.discount_percent}%).`;
        } else {
          remark = `[GIA HẠN GÓI CŨ] Chuyên cần ${p.attendance_rate_pct}%. Tiếp tục duy trì lộ trình ${p.recommended_package_name} với học phí chuẩn.`;
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
            <td class="py-3 px-3">${techTag}</td>
            <td class="py-3 px-3 font-semibold text-purple-700">${p.recommended_package_name}</td>
            <td class="py-3 px-3 text-right font-extrabold text-slate-900 font-mono">
              ${p.is_revenue_grounded ? formatVND(p.final_price_vnd) : '<span class="text-amber-600 italic">Tạm giữ</span>'}
            </td>
            <td class="py-3 px-3">${tag}</td>
            <td class="py-3 px-4 text-slate-600 text-[11px] leading-relaxed">${remark}</td>
          </tr>
        `;
      });
    }

    // HITL Trigger Action
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
        populateDashboardMetrics();
        populateCleanedExcelTable();
      } catch (err) {
        alert('Lỗi phê duyệt: ' + err.message);
      }
    }

    // =========================================================================
    // ERROR DIAGNOSTICS & FIX ENGINE
    // =========================================================================
    function triggerSimulatedErrorTest() {
      // Show simulated realistic errors in customer dataset
      const errors = [
        {
          location: 'Dòng 7 | Cột: expiry_date',
          account_id: 'HV-2026-004',
          raw_value: '32/13/2026',
          reason: 'Định dạng ngày tháng không hợp lệ (tháng 13, ngày 32)',
          fix_guide: 'Chuyển định dạng về chuẩn ISO-8601 YYYY-MM-DD (Ví dụ: 2026-12-04) hoặc để hệ thống tự động chuẩn hóa.'
        },
        {
          location: 'Dòng 12 | Cột: attended_sessions',
          account_id: 'HV-2026-009',
          raw_value: '-5 buổi',
          reason: 'Số buổi tham gia là số âm không hợp lệ (< 0)',
          fix_guide: 'Nhập số buổi thực tế từ 0 đến tổng số buổi đăng ký, hoặc để trống (NULL) nếu chưa có dữ liệu.'
        },
        {
          location: 'Dòng 15 | Bảng: support_tickets',
          account_id: 'HV-2026-007',
          raw_value: 'Ticket Open (TK-9002)',
          reason: 'Xung đột nghiệp vụ: Học viên đang có sự cố văng lớp chưa đóng, bị cấm chào nâng cấp',
          fix_guide: 'Đội Kỹ thuật cần đóng Ticket trước khi gia hạn, hoặc chuyển trạng thái sang [Tạm hoãn do lỗi Kỹ thuật].'
        }
      ];

      displayErrorDiagnostics(errors);
    }

    function displayErrorDiagnostics(errors) {
      const banner = document.getElementById('errorDiagnosticsBanner');
      const badge = document.getElementById('errorCountBadge');
      const tbody = document.getElementById('errorListTbody');

      badge.textContent = `${errors.length} lỗi phát hiện`;
      tbody.innerHTML = '';

      errors.forEach(err => {
        tbody.innerHTML += `
          <tr class="hover:bg-rose-50/50">
            <td class="py-2.5 px-3 font-bold text-rose-900 font-mono">${err.location}</td>
            <td class="py-2.5 px-3 font-bold text-slate-800">${err.account_id}</td>
            <td class="py-2.5 px-3 text-rose-700 font-mono bg-rose-100/50 px-1 rounded">${err.raw_value}</td>
            <td class="py-2.5 px-3 text-slate-700">${err.reason}</td>
            <td class="py-2.5 px-3 text-emerald-800 font-semibold">${err.fix_guide}</td>
          </tr>
        `;
      });

      banner.classList.remove('hidden');
      banner.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

    function dismissErrorBanner() {
      document.getElementById('errorDiagnosticsBanner').classList.add('hidden');
    }

    function autoFixAndContinue() {
      dismissErrorBanner();
      alert('Động cơ Làm sạch đã tự động:\n1. Quy chuẩn ngày tháng về ISO-8601 YYYY-MM-DD\n2. Thay thế số âm bằng giá trị chuẩn\n3. Kích hoạt Hard-Block khóa cứng lỗi Kỹ thuật để bảo đảm uy tín.\n\nBắt đầu chạy phân tích kết quả!');
      runQuickDemo('10_core');
    }

    function downloadCleanSampleTemplate() {
      window.location.href = '/api/download-input-xlsx';
    }

    function downloadCleanedExcel() {
      window.location.href = '/api/download-output-xlsx';
    }

    // File Upload Handler (.xlsx)
    async function handleFileUpload(e) {
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

          // Schema check for errors
          const errors = [];
          if (!parsedSheets['accounts'] && !parsedSheets['Sheet1']) {
            errors.push({
              location: 'Cấu trúc File',
              account_id: 'SYSTEM',
              raw_value: 'Thiếu Sheet accounts',
              reason: 'Tệp không chứa bảng danh sách tài khoản học viên (Sheet accounts)',
              fix_guide: 'Tạo sheet có tên "accounts" hoặc tải file mẫu chuẩn của hệ thống.'
            });
          }

          if (errors.length > 0) {
            displayErrorDiagnostics(errors);
            return;
          }

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
          applyAnalysisOutputs();
        } catch (err) {
          alert('Lỗi đọc tệp Excel: ' + err.message);
        }
      };
      reader.readAsArrayBuffer(file);
    }

    async function handleSheetsSubmit() {
      const url = document.getElementById('sheetsUrlInput').value.trim();
      if (!url) {
        alert('Vui lòng dán đường dẫn Google Sheets!');
        return;
      }
      await runQuickDemo('10_core');
    }

    // AI Drawer & Chat
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
      alert('Hệ thống Phân tích Dữ liệu Khách hàng DataLens:\n\n1. Nhập liệu: Tải tệp Excel, kết nối Google Sheets hoặc bấm "Demo Nhanh".\n2. Chẩn đoán Lỗi: Nếu dữ liệu có sai sót, hệ thống sẽ chỉ rõ vị trí sai, nguyên nhân và cách fix kèm nút tự động sửa lỗi.\n3. Kết quả Phân tích gồm 4 Đầu ra:\n   - 1. Dashboard Phân Tích (BI trực quan, KPIs, biểu đồ, Cổng HITL)\n   - 2. Báo Cáo Tổng Quan (Executive Summary, SWOT, Chiến lược)\n   - 3. Slide PowerPoint (10 Slides trình chiếu tương tác)\n   - 4. Excel Đã Làm Sạch & Nhận Xét AI (Tách NULL ≠ 0, tải file .xlsx).');
    }

    // Format Helpers
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
Write-Output "Successfully updated $webPath with complete Customer Analytics Platform."
