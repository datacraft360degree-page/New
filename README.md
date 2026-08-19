

<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Business Portal - Web Application Management System</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- SheetJS for Exporting to Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- html2canvas for Generating JPEG Receipts -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
  <!-- FontAwesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" />
  <style>
    /* -------------------------------------------------- */
    /* ROOT AND CORE SYSTEM BASE VARIABLES                */
    /* -------------------------------------------------- */
    :root {
      --primary-color: #2563eb;
      --primary-hover: #1d4ed8;
      --secondary-color: #0f172a;
      --bg-base: #f2f4f7;
      --border-color: #e2e8f0;
      --card-bg: #ffffff;
      --text-main: #1e293b;
      --text-muted: #64748b;
      --success-color: #10b981;
      --warning-color: #f59e0b;
      --danger-color: #ef4444;
      --transition-speed: 0.25s;
    }

    /* -------------------------------------------------- */
    /* GLOBAL SCROLLBAR & TYPOGRAPHY ADJUSTMENTS         */
    /* -------------------------------------------------- */
    html {
      scroll-behavior: smooth;
    }

    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      background-color: var(--bg-base);
      color: var(--text-main);
      margin: 0;
      padding: 0;
      box-sizing: border-box;
      overflow-x: hidden;
    }

    *, ::before, ::after {
      box-sizing: border-box;
    }

    ::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }

    ::-webkit-scrollbar-track {
      background: #F2F4F7;
      border-radius: 10px;
    }

    ::-webkit-scrollbar-thumb {
      background: #CBD5E1;
      border-radius: 10px;
    }

    ::-webkit-scrollbar-thumb:hover {
      background: #94A3B8;
    }

    /* -------------------------------------------------- */
    /* PRINTABLE INVOICE SPECIFIC OVERRIDES               */
    /* -------------------------------------------------- */
    @media print {
      body * {
        visibility: hidden;
      }
      #printable-invoice, #printable-invoice * {
        visibility: visible;
      }
      #printable-invoice {
        position: absolute;
        left: 0;
        top: 0;
        width: 100%;
        margin: 0;
        padding: 15px;
      }
      .no-print {
        display: none !important;
      }
    }

    /* -------------------------------------------------- */
    /* CUSTOM UI COMPONENTS AND CARDS                    */
    /* -------------------------------------------------- */
    .excel-comment-box::before {
      content: '';
      position: absolute;
      top: -8px;
      left: 16px;
      border-width: 0 8px 8px 8px;
      border-style: solid;
      border-color: transparent transparent #1E293B transparent;
    }

    .nav-control-fab {
      opacity: 0.35;
      transition: all var(--transition-speed) ease-in-out;
    }

    .nav-control-fab:hover {
      opacity: 1.0;
      transform: translateY(-2px);
    }

    .custom-card-shadow {
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
    }

    .custom-pill-active {
      background-color: #ffffff;
      color: #2563eb;
      box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
      font-weight: 700;
    }

    /* -------------------------------------------------- */
    /* UI MODAL ANIMATION CLASSES                         */
    /* -------------------------------------------------- */
    .modal-fadeIn {
      animation: fadeIn 0.2s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }

    @keyframes fadeIn {
      from {
        opacity: 0;
        transform: scale(0.96);
      }
      to {
        opacity: 1;
        transform: scale(1);
      }
    }

    /* -------------------------------------------------- */
    /* DASHBOARD CARD & STATS HOVER EFFECTS              */
    /* -------------------------------------------------- */
    .dash-stat-card {
      transition: transform var(--transition-speed) ease, box-shadow var(--transition-speed) ease;
    }

    .dash-stat-card:hover {
      transform: translateY(-2px);
      box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -2px rgba(0, 0, 0, 0.025);
    }

    /* -------------------------------------------------- */
    /* TABLE CUSTOM HIGHLIGHTS AND ROW STYLES             */
    /* -------------------------------------------------- */
    .table-row-hover {
      transition: background-color 0.15s ease;
    }

    .table-row-hover:hover {
      background-color: #f8fafc;
    }

    .status-badge-live {
      background-color: #fef3c7;
      color: #92400e;
      border: 1px solid #fde68a;
    }

    .status-badge-upcoming {
      background-color: #dbeafe;
      color: #1e40af;
      border: 1px solid #bfdbfe;
    }

    .status-badge-closed {
      background-color: #d1fae5;
      color: #065f46;
      border: 1px solid #a7f3d0;
    }

    .status-badge-inactive {
      background-color: #f1f5f9;
      color: #475569;
      border: 1px solid #e2e8f0;
    }
  </style>
</head>
<body class="text-slate-800 font-sans min-h-screen flex flex-col relative antialiased text-xs" onclick="closeCommentBox()">

  <!-- FLOATING HOVER PAGE NAVIGATION CONTROLS (UP/DOWN/LEFT/RIGHT) -->
  <div class="fixed bottom-6 left-6 z-50 flex items-center gap-1.5 bg-slate-900/80 backdrop-blur-md p-2 rounded-2xl border border-slate-700/60 shadow-2xl no-print nav-control-fab text-white">
    <button onclick="scrollPage('up')" title="Page Up" class="w-8 h-8 rounded-xl bg-slate-800 hover:bg-blue-600 flex items-center justify-center text-xs transition active:scale-95 border border-slate-700">
      <i class="fa-solid fa-arrow-up"></i>
    </button>
    <button onclick="scrollPage('down')" title="Page Down" class="w-8 h-8 rounded-xl bg-slate-800 hover:bg-blue-600 flex items-center justify-center text-xs transition active:scale-95 border border-slate-700">
      <i class="fa-solid fa-arrow-down"></i>
    </button>
    <div class="h-5 w-[1px] bg-slate-700 mx-0.5"></div>
    <button onclick="scrollPage('left')" title="Page Left" class="w-8 h-8 rounded-xl bg-slate-800 hover:bg-blue-600 flex items-center justify-center text-xs transition active:scale-95 border border-slate-700">
      <i class="fa-solid fa-arrow-left"></i>
    </button>
    <button onclick="scrollPage('right')" title="Page Right" class="w-8 h-8 rounded-xl bg-slate-800 hover:bg-blue-600 flex items-center justify-center text-xs transition active:scale-95 border border-slate-700">
      <i class="fa-solid fa-arrow-right"></i>
    </button>
  </div>

  <!-- LOGIN MODAL OVERLAY -->
  <div id="login-overlay" class="fixed inset-0 z-[60] bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-6 space-y-4 text-left modal-fadeIn">
      <div class="text-center space-y-1">
        <div class="bg-blue-50 text-blue-600 w-12 h-12 rounded-2xl flex items-center justify-center mx-auto text-xl shadow-sm">
          <i class="fa-solid fa-lock"></i>
        </div>
        <p class="text-[11px] text-slate-500">Please enter your credentials to access the system</p>
      </div>

      <form onsubmit="handleLogin(event)" class="space-y-3">
        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">User ID</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-user"></i>
            </span>
            <input type="text" id="login-userid" required="" placeholder="Enter User ID" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition" />
          </div>
        </div>

        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">Password</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-key"></i>
            </span>
            <input type="password" id="login-password" required="" placeholder="Enter Password" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition" />
          </div>
        </div>

        <div id="login-error" class="hidden bg-rose-50 border border-rose-100 text-rose-600 text-[10px] p-2 rounded-xl text-center font-medium">
          Invalid User ID or Password!
        </div>

        <button type="submit" class="w-full bg-blue-600 hover:bg-blue-700 active:scale-98 text-white font-bold py-2.5 rounded-2xl shadow-sm transition text-xs flex items-center justify-center gap-1.5">
          <i class="fa-solid fa-right-to-bracket"></i> Login
        </button>
      </form>
    </div>
  </div>

  <!-- LOGIN ALERT MESSAGE MODAL (POPUP ON SUCCESSFUL LOGIN) -->
  <div id="login-alert-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-blue-100 max-w-md w-full p-6 space-y-4 text-left modal-fadeIn">
      <div class="flex items-center gap-3 border-b border-slate-100 pb-3">
        <div class="bg-amber-50 text-amber-600 w-10 h-10 rounded-2xl flex items-center justify-center text-lg shadow-sm">
          <i class="fa-solid fa-circle-exclamation"></i>
        </div>
        <div>
          <h3 class="text-sm font-bold text-slate-900">Important System Guidelines</h3>
          <p class="text-[10px] text-slate-500">Please read these instructions before continuing</p>
        </div>
      </div>
      <div class="space-y-2.5 text-[11px] text-slate-700 font-medium">
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>1. Google Chrome/Microsoft Edge is the best viewing browser for this Portal.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>2. Take data backup every day or every week.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>3. Do not force close 'The Portal'; always close it using the 'Logout' option.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>4. Do not 'Login' with multiple devices/multiple browsers at the same time.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>5. Data save/Data fetch takes a moment, so please hold on ⏳.</span></p>
      </div>
      <button onclick="closeLoginAlertModal()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 rounded-2xl shadow-sm transition text-xs mt-2">
        I Understand, Continue
      </button>
    </div>
  </div>

  <!-- MASTER DATA ACCESS PASSWORD MODAL -->
  <div id="master-auth-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-left modal-fadeIn">
      <div class="text-center space-y-1">
        <div class="bg-rose-50 text-rose-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-lg shadow-sm">
          <i class="fa-solid fa-shield-halved"></i>
        </div>
        <h3 class="text-xs font-bold text-slate-900 mt-1">Master Data Protected</h3>
        <p class="text-[10px] text-slate-500">Enter master password to access configuration and deletion tools.</p>
      </div>

      <form onsubmit="handleMasterAuth(event)" class="space-y-2.5">
        <div>
          <label class="block text-[10px] font-semibold text-slate-700 mb-1">Master Password</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-key"></i>
            </span>
            <input type="password" id="master-password-input" required="" placeholder="Enter Master Password" class="w-full bg-slate-100 border border-transparent focus:border-rose-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition" />
          </div>
        </div>

        <div id="master-auth-error" class="hidden bg-rose-50 border border-rose-100 text-rose-600 text-[10px] p-1.5 rounded-xl text-center font-medium">
          Incorrect Master Password!
        </div>

        <div class="flex space-x-2 pt-1">
          <button type="button" onclick="closeMasterAuthModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">
            Cancel
          </button>
          <button type="submit" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px] flex items-center justify-center gap-1">
            <i class="fa-solid fa-unlock text-[10px]"></i> Unlock
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- MASTER DATA PERMANENT DELETION RECONFIRMATION POPUP MODAL -->
  <div id="master-delete-confirm-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-rose-100 max-w-sm w-full p-5 space-y-3 text-center modal-fadeIn">
      <div class="bg-rose-50 text-rose-600 w-12 h-12 rounded-2xl flex items-center justify-center mx-auto text-xl shadow-sm">
        <i class="fa-solid fa-triangle-exclamation"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-900">Confirm Permanent Deletion</h3>
        <p id="master-delete-modal-msg" class="text-[11px] text-slate-600 mt-1">Are you sure you want to permanently delete this data from Master Tab? This action cannot be undone.</p>
      </div>
      <div class="flex space-x-2 pt-2">
        <button type="button" onclick="closeMasterDeleteModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">
          Cancel
        </button>
        <button type="button" onclick="confirmMasterDeletion()" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px] flex items-center justify-center gap-1">
          <i class="fa-solid fa-trash-can text-[10px]"></i> Delete Permanently
        </button>
      </div>
    </div>
  </div>

  <!-- MANUAL LOGOUT CONFIRM MODAL -->
  <div id="logout-confirm-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-center modal-fadeIn">
      <div class="bg-rose-50 text-rose-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-lg">
        <i class="fa-solid fa-right-from-bracket"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-900">Confirm Logout</h3>
        <p class="text-[10px] text-slate-500 mt-1">Are you sure you want to log out? All your latest changes will be saved securely before exiting.</p>
      </div>
      <div class="flex space-x-2 pt-2">
        <button type="button" onclick="cancelLogout()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">Cancel</button>
        <button type="button" onclick="processLogoutWithSave()" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px]">Logout</button>
      </div>
    </div>
  </div>

  <!-- SAVING LOCK MODAL (SAND TIMER) -->
  <div id="saving-lock-modal" class="hidden fixed inset-0 z-[100] bg-slate-900/60 backdrop-blur-md flex items-center justify-center p-4 no-print cursor-wait">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-6 text-center space-y-4 modal-fadeIn">
      <div class="text-blue-600 text-5xl animate-bounce">
        ⏳
      </div>
      <div>
        <h3 class="text-lg font-black text-slate-900">Saving & Logging Out...</h3>
        <p class="text-xs text-rose-600 mt-2 font-bold uppercase">Do not close window or shutdown!</p>
        <p class="text-[10px] text-slate-500 mt-1">Please wait while we secure your data.</p>
      </div>
    </div>
  </div>

  <!-- SESSION AUTO LOGOUT WARNING MODAL -->
  <div id="logout-warning-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-center modal-fadeIn">
      <div class="bg-amber-50 text-amber-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-lg">
        <i class="fa-solid fa-hourglass-half"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-900">Inactivity Timeout Warning</h3>
        <p class="text-[10px] text-slate-500 mt-1">You will be logged out automatically in <strong id="logout-countdown-seconds" class="text-rose-600">60</strong> seconds due to inactivity. Data will be saved securely.</p>
      </div>
      <button onclick="resetInactivityTimer()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 rounded-xl text-[11px] transition shadow-sm">
        Stay Logged In
      </button>
    </div>
  </div>

  <!-- EXPORT TO EXCEL DATE RANGE MODAL -->
  <div id="export-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-5 space-y-4 text-left modal-fadeIn">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
          <i class="fa-solid fa-file-excel text-emerald-600"></i> Export to Excel
        </h3>
        <button onclick="closeExportModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>
      <div class="space-y-3 text-[11px]">
        <p class="text-slate-500">Select a specific period to download booking details. Available from 1st Aug 2026 to 31st Dec 2085.</p>
        <div>
          <label class="block font-semibold text-slate-600 mb-0.5">Start Date</label>
          <input type="date" id="export-start-date" min="2026-08-01" max="2085-12-31" onchange="validateExportDates()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-emerald-500 font-medium" />
        </div>
        <div>
          <label class="block font-semibold text-slate-600 mb-0.5">End Date</label>
          <input type="date" id="export-end-date" min="2026-08-01" max="2085-12-31" onchange="validateExportDates()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-emerald-500 font-medium" />
        </div>
      </div>
      <div class="flex space-x-2 pt-2 border-t border-slate-100">
        <button type="button" onclick="closeExportModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">
          Cancel
        </button>
        <button type="button" onclick="processExport()" class="w-1/2 bg-emerald-600 hover:bg-emerald-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px] flex items-center justify-center gap-1">
          <i class="fa-solid fa-download text-[10px]"></i> Download
        </button>
      </div>
    </div>
  </div>

  <!-- Excel Comment Box Popout -->
  <div id="excel-comment-box" onclick="event.stopPropagation()" class="excel-comment-box hidden absolute z-50 bg-slate-900 text-white text-[11px] rounded-2xl p-3 shadow-2xl border border-slate-800 space-y-2 w-64 transition-all duration-150">
    <div class="font-bold text-blue-400 border-b border-slate-800 pb-1.5 flex justify-between items-center text-[10px]">
      <span class="flex items-center gap-1.5">
        <i class="fa-solid fa-comment-dots text-blue-400"></i>
        <span id="comm-date-header">Date Overview</span>
      </span>
      <button onclick="closeCommentBox()" class="text-slate-400 hover:text-white px-1 py-0.5 rounded text-[10px]">
        <i class="fa-solid fa-xmark"></i>
      </button>
    </div>
    <div id="comm-booking-list" class="space-y-1.5 max-h-56 overflow-y-auto pr-0.5"></div>
  </div>

  <!-- Header Navigation -->
  <header class="bg-white/80 backdrop-blur-md border-b border-slate-200/60 text-slate-900 sticky top-0 z-40 no-print">
    <div class="max-w-7xl mx-auto px-4 py-2.5 flex flex-col md:flex-row justify-between items-center gap-2.5">
      <div class="flex items-center space-x-2.5">
        <div class="bg-blue-600 p-2 rounded-2xl text-white shadow-sm">
          <i class="fa-solid fa-hotel text-sm"></i>
        </div>
        <div>
          <p class="text-[10px] text-slate-500 mt-0.5">Management &amp; Booking Control System</p>
        </div>
      </div>
      
      <!-- One UI Pill Navigation -->
      <nav class="flex space-x-1 bg-slate-100 p-1 rounded-full text-[11px] font-medium">
        <button onclick="switchTab('dashboard')" id="btn-dashboard" class="tab-btn px-3 py-1 rounded-full transition-all active-tab bg-white text-blue-600 shadow-sm font-bold">Dashboard</button>
        <button onclick="switchTab('booking')" id="btn-booking" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900">Booking Details</button>
        <button onclick="switchTab('master')" id="btn-master" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900 flex items-center gap-1">
          <i class="fa-solid fa-lock text-[9px] text-amber-500"></i> Master Data
        </button>
        <button onclick="switchTab('calendar')" id="btn-calendar" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900">Calendar</button>
      </nav>

      <!-- Action Buttons -->
      <div class="flex items-center space-x-1.5">
        <button onclick="openAlertModal()" title="View Alerts" class="relative bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-bell text-[10px]"></i> Alerts
          <span id="alert-badge" class="hidden absolute -top-1 -right-1 bg-rose-600 text-white text-[9px] font-black px-1.5 py-0.2 rounded-full border border-white animate-bounce">0</span>
        </button>
        <button onclick="saveChanges()" class="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-brands fa-google text-[10px]"></i> Save
        </button>
        
        <div id="wipe-layer-1-modal" class="hidden fixed inset-0 z-[60] bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 no-print">
          <div class="bg-white rounded-3xl shadow-2xl border border-rose-100 max-w-sm w-full p-6 text-center space-y-4">
            <div class="bg-rose-100 text-rose-600 w-16 h-16 rounded-full flex items-center justify-center mx-auto text-3xl shadow-sm">
              <i class="fa-solid fa-triangle-exclamation"></i>
            </div>
            <div>
              <h3 class="text-base font-black text-slate-900">Initiate Data Wipe?</h3>
              <p class="text-xs text-slate-600 mt-2">You are about to delete ALL data from the Google Sheet. This affects bookings, rooms, and agents.</p>
            </div>
            <div class="flex space-x-3 pt-2">
              <button type="button" onclick="closeWipeModals()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-800 font-bold py-2.5 rounded-xl text-xs transition">Cancel</button>
              <button type="button" onclick="proceedToWipeLayer2()" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2.5 rounded-xl shadow-sm transition text-xs">Proceed</button>
            </div>
          </div>
        </div>

        <div id="wipe-layer-2-modal" class="hidden fixed inset-0 z-[70] bg-rose-900/80 backdrop-blur-md flex items-center justify-center p-4 no-print">
          <div class="bg-black rounded-3xl shadow-2xl border border-rose-600 max-w-sm w-full p-6 text-center space-y-4">
            <div class="text-rose-500 w-16 h-16 rounded-full flex items-center justify-center mx-auto text-4xl animate-pulse">
              <i class="fa-solid fa-skull-crossbones"></i>
            </div>
            <div>
              <h3 class="text-lg font-black text-white uppercase tracking-widest">Final Warning</h3>
              <p class="text-xs text-rose-200 mt-2">This action is <strong class="text-white">IRREVERSIBLE</strong>. All records will be permanently deleted from the database. Are you absolutely sure?</p>
            </div>
            <div class="flex space-x-3 pt-2">
              <button type="button" onclick="closeWipeModals()" class="w-1/2 bg-slate-800 hover:bg-slate-700 text-white font-bold py-2.5 rounded-xl text-xs transition">Cancel</button>
              <button type="button" id="btn-final-wipe" onclick="executeGoogleSheetWipe()" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2.5 rounded-xl shadow-lg shadow-rose-900/50 transition text-xs">ERASE ALL DATA</button>
            </div>
          </div>
        </div>

        <button onclick="requestDataWipe()" class="bg-rose-600 hover:bg-rose-700 text-white border border-rose-800 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition shadow-sm">
          <i class="fa-solid fa-skull-crossbones text-[8px]"></i> Wipe Data
        </button>

        <button onclick="openExportModal()" class="bg-blue-50 hover:bg-blue-100 text-blue-700 border border-blue-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-file-excel text-[10px]"></i> Export
        </button>
        <button onclick="logoutUser()" title="Logout" class="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-right-from-bracket text-[10px]"></i> Logout
        </button>
      </div>
    </div>
  </header>

  <!-- Notification Toast -->
  <div id="toast" class="hidden fixed bottom-6 right-6 bg-slate-900/90 backdrop-blur-md text-white px-4 py-2.5 rounded-2xl shadow-xl z-50 flex items-center gap-2.5 no-print border border-slate-800 text-[11px]">
    <i class="fa-solid fa-circle-check text-emerald-400 text-base"></i>
    <span id="toast-message" class="font-medium">Changes Auto save successfully!</span>
  </div>

  <!-- Main Content Area -->
  <main class="max-w-7xl mx-auto px-4 py-4 flex-1 w-full no-print space-y-4">

    <!-- DASHBOARD TAB -->
    <section id="tab-dashboard" class="tab-content space-y-4">
      <div class="bg-gradient-to-r from-blue-600 to-indigo-600 rounded-3xl p-5 text-white shadow-sm flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
        <div>
          <h2 class="text-base font-bold tracking-tight">Hi Aniruddha, Welcome to dashboard 🏠</h2>
          <p class="text-blue-100 text-[10px] mt-0.5">Quickly view, schedule, and manage room allocations and orders.</p>
        </div>
        <div class="flex items-center bg-white/10 backdrop-blur-md px-3 py-1.5 rounded-2xl border border-white/20 space-x-2">
          <label for="dash-year-select" class="text-[10px] font-bold text-blue-50 uppercase flex items-center gap-1">
            <i class="fa-solid fa-filter text-amber-300"></i> Filter Year:
          </label>
          <select id="dash-year-select" onchange="handleDashboardYearChange(this.value)" class="bg-white text-blue-900 text-[11px] font-bold rounded-xl px-2.5 py-1 focus:outline-none cursor-pointer shadow-sm">
          </select>
        </div>
      </div>

      <!-- Summary Filter Banner Indicator -->
      <div class="flex items-center justify-between bg-white px-4 py-2 rounded-2xl border border-slate-200/60 shadow-sm">
        <span class="text-[11px] font-semibold text-slate-600 flex items-center gap-2">
          <i class="fa-solid fa-chart-line text-blue-600"></i>
          Showing Summary For: <strong id="dash-filter-label" class="text-blue-600 font-bold">Consolidated (All Years)</strong>
        </span>
        <button onclick="handleDashboardYearChange('CURRENT')" class="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-3 py-1 rounded-full transition border border-slate-200">
          Reset to Current Year
        </button>
      </div>

      <!-- Financial Metrics Grid -->
      <div class="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex items-center justify-between dash-stat-card">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Total Bookings</p>
            <p id="dash-total-bookings" class="text-xl font-black text-slate-900 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-blue-50 text-blue-600 rounded-2xl"><i class="fa-solid fa-bookmark text-base"></i></div>
        </div>
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex items-center justify-between dash-stat-card">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Booking Amount</p>
            <p id="dash-total-amount" class="text-xl font-black text-slate-900 mt-0.5">₹0</p>
          </div>
          <div class="p-3 bg-indigo-50 text-indigo-600 rounded-2xl"><i class="fa-solid fa-receipt text-base"></i></div>
        </div>
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex items-center justify-between dash-stat-card">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Amount Received</p>
            <p id="dash-advanced" class="text-xl font-black text-emerald-600 mt-0.5">₹0</p>
          </div>
          <div class="p-3 bg-emerald-50 text-emerald-600 rounded-2xl"><i class="fa-solid fa-wallet text-base"></i></div>
        </div>
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex items-center justify-between dash-stat-card">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Total Due Amount</p>
            <p id="dash-due" class="text-xl font-black text-rose-600 mt-0.5">₹0</p>
          </div>
          <div class="p-3 bg-rose-50 text-rose-600 rounded-2xl"><i class="fa-solid fa-hand-holding-dollar text-base"></i></div>
        </div>
      </div>

      <!-- Year-Wise Status Breakdown Counts -->
      <div class="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-amber-200/80 flex items-center justify-between dash-stat-card">
          <div>
            <p class="text-[9px] uppercase font-bold text-amber-600 tracking-wider">Live Booking Count</p>
            <p id="dash-live-count" class="text-xl font-black text-amber-600 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-amber-50 text-amber-600 rounded-2xl"><i class="fa-solid fa-bed text-base"></i></div>
        </div>
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-blue-200/80 flex items-center justify-between dash-stat-card">
          <div>
            <p class="text-[9px] uppercase font-bold text-blue-600 tracking-wider">Upcoming Booking Count</p>
            <p id="dash-upcoming-count" class="text-xl font-black text-blue-600 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-blue-50 text-blue-600 rounded-2xl"><i class="fa-solid fa-calendar-plus text-base"></i></div>
        </div>
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-emerald-200/80 flex items-center justify-between dash-stat-card">
          <div>
            <p class="text-[9px] uppercase font-bold text-emerald-600 tracking-wider">Closed Booking Count</p>
            <p id="dash-closed-count" class="text-xl font-black text-emerald-600 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-emerald-50 text-emerald-600 rounded-2xl"><i class="fa-solid fa-circle-check text-base"></i></div>
        </div>
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-300 flex items-center justify-between dash-stat-card">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-500 tracking-wider">Inactive Booking Count</p>
            <p id="dash-inactive-count" class="text-xl font-black text-slate-700 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-slate-100 text-slate-600 rounded-2xl"><i class="fa-solid fa-ban text-base"></i></div>
        </div>
      </div>

      <!-- Active Years Directory Table Hidden -->
      <div class="hidden bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="mb-3 flex justify-between items-center">
          <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
            <i class="fa-solid fa-calendar-days text-blue-600"></i> Active Years Directory (2026 – 2085)
          </h3>
          <span class="text-[10px] text-slate-400 font-medium">Click any year to filter dashboard &amp; open year calendar</span>
        </div>
        <div id="years-grid" class="grid grid-cols-6 sm:grid-cols-10 md:grid-cols-12 gap-2"></div>
      </div>
    </section>

    <!-- BOOKING DETAILS TAB -->
    <section id="tab-booking" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-3 mb-4 pb-3 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-address-card text-blue-600"></i> Guest Information &amp; Reservation Directory
            </h2>
            <div class="flex items-center gap-3 mt-2 text-[10px]">
              <span class="flex items-center gap-1.5 font-semibold text-amber-800">
                <span class="relative flex h-2.5 w-2.5">
                  <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
                  <span class="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500"></span>
                </span> Live Booking
              </span>
              <span class="flex items-center gap-1.5 font-semibold text-blue-800">
                <span class="w-2.5 h-2.5 bg-blue-500 rounded-full inline-block"></span> Upcoming Booking
              </span>
              <span class="flex items-center gap-1.5 font-semibold text-emerald-800">
                <span class="w-2.5 h-2.5 bg-emerald-500 rounded-full inline-block"></span> Closed Booking
              </span>
              <span class="flex items-center gap-1.5 font-semibold text-slate-700">
                <span class="w-2 h-2 bg-rose-600 rounded-full inline-block"></span> Inactive Booking
              </span>
            </div>
          </div>
          
          <div class="flex items-center space-x-2 w-full md:w-auto">
            <div class="flex items-center bg-slate-100 border border-slate-200 rounded-2xl px-2 py-1 space-x-1.5">
              <label for="booking-date-search" class="text-[10px] font-bold text-slate-500 uppercase flex items-center gap-1 pl-1">
                <i class="fa-solid fa-calendar-day text-blue-600"></i> Search Date:
              </label>
              <input type="date" id="booking-date-search" onchange="searchBookingByDate()" class="bg-white text-[11px] border border-slate-200 rounded-xl px-2 py-0.5 focus:outline-none focus:ring-2 focus:ring-blue-500 font-bold text-blue-600 cursor-pointer" />
              <button onclick="clearDateSearchBooking()" class="text-slate-400 hover:text-slate-600 px-1 text-[10px]" title="Reset Filter">
                <i class="fa-solid fa-rotate-left"></i> Reset
              </button>
            </div>

            <button onclick="openBookingModal()" class="bg-blue-600 hover:bg-blue-700 text-white px-3.5 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1.5 transition shadow-sm whitespace-nowrap">
              <i class="fa-solid fa-plus text-[10px]"></i> Add Booking
            </button>
          </div>
        </div>

        <!-- Bookings Table View -->
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200/80 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                <th class="py-2.5 px-3">Booking ID</th>
                <th class="py-2.5 px-3">Guest Name</th>
                <th class="py-2.5 px-3">Contact No</th>
                <th class="py-2.5 px-3">ID No</th>
                <th class="py-2.5 px-3">Attached ID</th>
                <th class="py-2.5 px-3">Room</th>
                <th class="py-2.5 px-3">Capacity</th>
                <th class="py-2.5 px-3">Agent Info</th>
                <th class="py-2.5 px-3 min-w-[150px]">Stay Window</th>
                <th class="py-2.5 px-3">Tariff &amp; Extras</th>
                <th class="py-2.5 px-3">Payment/Adv</th>
                <th class="py-2.5 px-3">Due</th>
                <th class="py-2.5 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="bookings-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- MASTER DATA TAB -->
    <section id="tab-master" class="tab-content hidden space-y-4">
      
      <!-- Room Capacity Table -->
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="flex justify-between items-center mb-3 pb-2 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-door-open text-blue-600"></i> Room Capacity Configuration
            </h2>
            <p class="text-[10px] text-slate-500 mt-0.5">Default rooms 1 to 5. Click Add Room to append new rooms anytime.</p>
          </div>
          <button type="button" onclick="addRoomCapacityRow()" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-full text-[11px] font-medium flex items-center gap-1.5 transition shadow-sm cursor-pointer">
            <i class="fa-solid fa-plus text-[10px]"></i> Add Room
          </button>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200/80 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                <th class="py-2.5 px-3">Room No</th>
                <th class="py-2.5 px-3">Room Capacity (Person)</th>
                <th class="py-2.5 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="room-capacity-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
          </table>
        </div>
      </div>

      <!-- Agent Information Table -->
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="flex justify-between items-center mb-3 pb-2 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-users-gear text-blue-600"></i> Master Agent Directory
            </h2>
            <p class="text-[10px] text-slate-500 mt-0.5">Manage Agents linked with room allocations.</p>
          </div>
          <button type="button" onclick="addAgentRow()" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-full text-[11px] font-medium flex items-center gap-1.5 transition shadow-sm cursor-pointer">
            <i class="fa-solid fa-plus text-[10px]"></i> Add Agent Entry
          </button>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200/80 text-[10px] font-bold text-slate-400 uppercase tracking-wider">
                <th class="py-2.5 px-3">Agent Name</th>
                <th class="py-2.5 px-3">Agent Contact</th>
                <th class="py-2.5 px-3">Linked Room No</th>
                <th class="py-2.5 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="agent-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
          </table>
        </div>
      </div>

      <!-- BOOKING ID TYPING SEARCH & DELETION CONTROL -->
      <div class="bg-white rounded-3xl shadow-sm border border-rose-200/80 p-4 space-y-3">
        <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 border-b border-slate-100 pb-2.5">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-trash-can text-rose-600"></i> Booking Deletion Manager
            </h2>
            <p class="text-[10px] text-slate-500 mt-0.5">Type a Booking ID directly to safely locate and remove it from the system.</p>
          </div>
          
          <div class="flex items-center bg-slate-100 border border-slate-200 rounded-2xl px-2 py-1 space-x-1.5">
            <label for="master-booking-search-input" class="text-[10px] font-bold text-slate-600 uppercase flex items-center gap-1 pl-1">
              <i class="fa-solid fa-magnifying-glass text-blue-600"></i> Type Booking ID:
            </label>
            <input type="text" id="master-booking-search-input" oninput="searchMasterBookingById()" placeholder="e.g. BKG-2026-0000001" class="bg-white text-[11px] border border-slate-200 rounded-xl px-2 py-0.5 focus:outline-none focus:ring-2 focus:ring-blue-500 font-mono font-bold text-blue-600 uppercase w-48" />
            <button onclick="clearMasterBookingSearch()" class="text-slate-400 hover:text-slate-600 px-1 text-[10px]" title="Clear Search">
              <i class="fa-solid fa-xmark"></i>
            </button>
          </div>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-rose-50/60 border-b border-rose-100 text-[10px] font-bold text-rose-800 uppercase tracking-wider">
                <th class="py-2.5 px-3">Booking ID</th>
                <th class="py-2.5 px-3">Guest Name</th>
                <th class="py-2.5 px-3">Room No</th>
                <th class="py-2.5 px-3">Stay Window</th>
                <th class="py-2.5 px-3">Total Amount</th>
                <th class="py-2.5 px-3">Due Amount</th>
                <th class="py-2.5 px-3 text-center">Delete Linked Booking</th>
              </tr>
            </thead>
            <tbody id="master-delete-tbody" class="divide-y divide-slate-100 text-[11px]">
              <tr>
                <td colspan="7" class="text-center py-4 text-slate-400">Please type a Booking ID into the search field above to view and delete details.</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- CALENDAR TAB -->
    <section id="tab-calendar" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="flex justify-between items-center mb-4">
          <div>
            <h2 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-regular fa-calendar-check text-blue-600"></i> Year Overview Calendar
            </h2>
          </div>
          
          <div class="flex items-center bg-slate-100 border border-slate-200 px-3 py-1.5 rounded-2xl space-x-2">
            <label for="cal-year-select" class="text-[10px] font-bold text-slate-600 uppercase flex items-center gap-1">
              <i class="fa-solid fa-filter text-blue-600"></i> Filter Year:
            </label>
            <select id="cal-year-select" onchange="renderCalendar(parseInt(this.value))" class="bg-white text-blue-900 text-[11px] font-bold rounded-xl px-2.5 py-1 focus:outline-none focus:ring-2 focus:ring-blue-500 cursor-pointer shadow-sm"></select>
          </div>
        </div>

        <div id="calendar-container" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-3"></div>
      </div>
    </section>

  </main>

  <!-- POPUP MODAL: CHECK-OUT ALERT LIST -->
  <div id="alert-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-lg w-full flex flex-col max-h-[85vh] overflow-hidden modal-fadeIn">
      <div class="bg-amber-500 p-4 text-white flex justify-between items-center">
        <div class="flex items-center space-x-2">
          <i class="fa-solid fa-bell text-base"></i>
          <h3 class="text-xs font-bold">Due Payment Alert</h3>
        </div>
        <button onclick="closeAlertModal()" class="text-amber-100 hover:text-white px-1 text-base">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>

      <div id="alert-list-container" class="p-4 overflow-y-auto space-y-2 flex-1 text-[11px]"></div>

      <div class="bg-slate-50 border-t border-slate-100 p-3 flex justify-between items-center text-[11px]">
        <span id="alert-list-count-text" class="text-slate-500 font-medium">0 active warnings found</span>
        <button onclick="closeAlertModal()" class="px-4 py-1.5 bg-slate-900 text-white rounded-xl font-semibold text-[10px]">Dismiss</button>
      </div>
    </div>
  </div>

  <!-- COMPACT ADD / EDIT BOOKING MODAL -->
  <div id="booking-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-3 overflow-y-auto no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-3xl w-full p-5 space-y-3 my-4 max-h-[90vh] overflow-y-auto modal-fadeIn">
      <div class="flex justify-between items-center pb-2.5 border-b border-slate-100">
        <div>
          <h3 id="modal-title" class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
            <i class="fa-solid fa-calendar-plus text-blue-600"></i> Add New Booking
          </h3>
        </div>
        <button onclick="closeBookingModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <form id="booking-form" onsubmit="handleSaveBooking(event)" class="space-y-3 text-[11px]">
        <input type="hidden" id="modal-booking-id" />

        <!-- GUEST DETAILS -->
        <div id="sec-guest-info" class="bg-slate-50 p-3 rounded-2xl border border-slate-200/60 space-y-2.5 transition-all">
          <h4 class="text-[9px] font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
            <i class="fa-solid fa-user-tag text-blue-600"></i> Guest Information
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Guest Name <span class="text-rose-500">*</span></label>
              <input type="text" id="cust-name" required="" pattern="[A-Za-z\s]+" oninput="this.value = formatTitleCase(this.value.replace(/[^A-Za-z\s]/g, ''))" title="Please enter Guest Name using characters only" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Address</label>
              <input type="text" id="cust-address" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">City</label>
              <input type="text" id="cust-city" placeholder="City" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">State</label>
              <input type="text" id="cust-state" oninput="this.value = formatTitleCase(this.value); handleStateChange(this.value)" placeholder="State" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Country</label>
              <input type="text" id="cust-country" placeholder="Country" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Pin/Zip Code</label>
              <input type="text" id="cust-zip" placeholder="Pin/Zip Code" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">ID Number</label>
              <input type="text" id="cust-id" maxlength="16" pattern="[A-Za-z0-9\s]*" oninput="this.value = this.value.replace(/[^A-Za-z0-9\s]/g, '')" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Contact No</label>
              <div class="flex gap-1">
                <input type="text" id="cust-country-code" value="+91" placeholder="+91" class="w-1/3 bg-white border border-slate-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-center text-blue-700" />
                <input type="text" id="cust-contact" maxlength="10" pattern="[0-9]*" oninput="this.value = this.value.replace(/[^0-9]/g, '').slice(0, 10)" placeholder="Mobile No" class="w-2/3 bg-white border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500" />
              </div>
            </div>
            <div class="sm:col-span-2">
              <label class="block font-semibold text-slate-600 mb-0.5 flex justify-between items-center">
                <span>Attached ID Proof <span class="text-[9px] text-blue-600 font-normal">(PDF, 10KB - 900KB)</span></span>
                <button type="button" id="cust-id-file-remove" onclick="removeAttachedIdProof()" class="hidden text-rose-500 hover:text-rose-700 text-[9px] font-bold">Remove</button>
              </label>
              <div class="flex items-center gap-1.5">
                <input type="file" id="cust-id-file" accept="application/pdf" onchange="handleIdProofUpload(event)" class="w-full text-[10px] text-slate-500 file:mr-2 file:py-1 file:px-2.5 file:rounded-xl file:border-0 file:text-[10px] file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 cursor-pointer bg-white border border-slate-200 rounded-xl py-1" />
                <input type="hidden" id="cust-id-file-base64" />
                <input type="hidden" id="cust-id-file-name" />
              </div>
              <p id="cust-id-file-status" class="text-[9px] text-slate-400 mt-0.5 italic">No PDF document attached.</p>
            </div>
          </div>
        </div>

        <!-- Room & Stay Schedule Box -->
        <div id="sec-room-dates" class="bg-slate-50 p-3 rounded-2xl border border-slate-200/60 space-y-2.5 transition-all">
          <h4 class="text-[9px] font-bold uppercase tracking-wider text-slate-400 flex items-center gap-1.5">
            <i class="fa-solid fa-bed text-blue-600"></i> Room Selection &amp; Stay Dates
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Room No(s)</label>
              <div class="relative" id="room-dropdown-container">
                <button type="button" onclick="toggleRoomDropdown()" id="room-dropdown-btn" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 focus:outline-none focus:border-blue-500 font-bold text-blue-600 text-left flex justify-between items-center" style="height: 34px;">
                  <span id="room-dropdown-text" class="truncate pr-2">Select Rooms...</span>
                  <i class="fa-solid fa-chevron-down text-slate-400"></i>
                </button>
                <div id="room-checkboxes" class="hidden absolute z-50 w-full mt-1 bg-white border border-slate-200 rounded-xl shadow-xl max-h-48 overflow-y-auto p-2 space-y-1">
                  <!-- Generated Checkboxes Go Here -->
                </div>
              </div>
            </div>
            
            <div class="flex flex-col gap-2">
              <div>
                <label class="block font-semibold text-slate-600 mb-0.5">Agent Info</label>
                <select id="cust-agent" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-slate-700"></select>
              </div>

              <div>
                <label class="block font-semibold text-slate-600 mb-0.5">Total Capacity</label>
                <input type="number" id="cust-capacity" min="1" value="1" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-slate-700" />
              </div>
            </div>

            <!-- EXTRA PERSON(S) COUNT FIELD -->
            <div>
              <label class="block font-semibold text-amber-700 mb-0.5 flex items-center gap-1">
                <i class="fa-solid fa-user-plus text-amber-600"></i> Add Extra Person(s)
              </label>
              <input type="number" id="cust-extra-persons" min="0" value="0" placeholder="0" oninput="calculateModalBilling()" class="w-full bg-amber-50 border border-amber-300 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-amber-500 font-bold text-amber-900" />
            </div>

            <!-- ADDITIONAL PERSON CUSTOM CHECK-IN & CHECK-OUT WINDOW -->
            <div id="sec-extra-person-time-wrapper" class="sm:col-span-4 hidden bg-amber-50/70 p-2.5 rounded-2xl border border-amber-200/80 space-y-2">
              <label class="block font-bold text-amber-900 mb-1 flex items-center gap-1">
                <i class="fa-solid fa-clock-rotate-left text-amber-600"></i> Additional Person Stay Window (Custom Dates Required)
              </label>
              <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
                <div>
                  <label class="block font-semibold text-amber-800 text-[10px] mb-0.5">Extra Person Check-In</label>
                  <div class="flex gap-1">
                    <input type="date" id="cust-extra-person-date" onchange="handleExtraPersonDatesChange()" class="w-2/3 bg-white border border-amber-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
                    <input type="time" id="cust-extra-person-time" onchange="handleExtraPersonDatesChange()" class="w-1/3 bg-white border border-amber-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
                  </div>
                </div>
                <div>
                  <label class="block font-semibold text-amber-800 text-[10px] mb-0.5">Extra Person Check-Out</label>
                  <div class="flex gap-1">
                    <input type="date" id="cust-extra-person-out-date" onchange="handleExtraPersonDatesChange()" class="w-2/3 bg-white border border-amber-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
                    <input type="time" id="cust-extra-person-out-time" onchange="handleExtraPersonDatesChange()" class="w-1/3 bg-white border border-amber-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
                  </div>
                </div>
              </div>
            </div>

            <div class="sm:col-span-4 grid grid-cols-2 gap-2 pt-1 border-t border-slate-200/60">
              <div>
                <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-solid fa-plane-arrival text-emerald-600 mr-1"></i> Check In</label>
                <div class="flex gap-1">
                  <input type="date" id="cust-checkin-date" onchange="handleStayDatesChange()" required="" class="w-2/3 bg-white border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
                  <input type="time" id="cust-checkin-time" onchange="handleStayDatesChange()" required="" class="w-1/3 bg-white border border-slate-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
                </div>
              </div>

              <div>
                <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-solid fa-plane-departure text-rose-500 mr-1"></i> Check Out</label>
                <div class="flex gap-1">
                  <input type="date" id="cust-checkout-date" onchange="handleStayDatesChange()" required="" class="w-2/3 bg-white border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
                  <input type="time" id="cust-checkout-time" onchange="handleStayDatesChange()" required="" class="w-1/3 bg-white border border-slate-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
                </div>
              </div>
            </div>

            <!-- Extended Check Out Checkbox & Fields -->
            <div id="sec-extended-checkout-wrapper" class="sm:col-span-4 pt-2 border-t border-slate-200/60">
              <div class="flex items-center gap-2 mb-1">
                <input type="checkbox" id="cust-has-extended-checkout" onchange="toggleExtendedCheckoutFields(this.checked)" class="w-4 h-4 text-blue-600 rounded-md border-slate-300 focus:ring-blue-500 cursor-pointer" />
                <label for="cust-has-extended-checkout" id="lbl-has-extended-checkout" class="font-bold text-slate-700 cursor-pointer flex items-center gap-1 select-none text-[11px]">
                  <i class="fa-solid fa-clock-rotate-left text-blue-600"></i> Extended Check-out Date &amp; Time <span id="ext-checkout-timer-notice" class="text-[9px] text-amber-700 font-normal ml-1 hidden">(Active post check-out)</span>
                </label>
              </div>

              <div id="extended-checkout-container" class="hidden bg-blue-50/70 p-2.5 rounded-2xl border border-blue-200/80 mt-1.5">
                <label class="block font-bold text-blue-900 mb-1 flex items-center gap-1">
                  <i class="fa-solid fa-calendar-plus text-blue-600"></i> New Check-out Date &amp; Time
                </label>
                <div class="flex gap-1.5">
                  <input type="date" id="cust-ext-checkout-date" onchange="handleStayDatesChange()" class="w-2/3 bg-white border border-blue-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500 font-semibold text-blue-900" />
                  <input type="time" id="cust-ext-checkout-time" onchange="handleStayDatesChange()" class="w-1/3 bg-white border border-blue-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-blue-500 font-semibold text-blue-900" />
                </div>
              </div>
            </div>

            <!-- Meal Plan Inclusions Checkbox -->
            <div class="sm:col-span-4 pt-2 border-t border-slate-200/60">
              <div class="flex items-center gap-2">
                <input type="checkbox" id="cust-include-meals" checked="" class="w-4 h-4 text-blue-600 rounded-md border-slate-300 focus:ring-blue-500 cursor-pointer" />
                <label for="cust-include-meals" class="font-bold text-slate-700 cursor-pointer flex items-center gap-1 select-none text-[11px]">
                  <i class="fa-solid fa-utensils text-emerald-600"></i> Include Meal (*Include Breakfast, Lunch, Evening snack &amp; Dinner)
                </label>
              </div>
            </div>

          </div>
        </div>

        <!-- EXTRA FOOD SECTION WITH DATE & TIME -->
        <div id="sec-extra-food" class="bg-amber-50/60 p-3 rounded-2xl border border-amber-200/80 space-y-2.5 transition-all">
          <div class="flex justify-between items-center">
            <h4 class="text-[9px] font-bold uppercase tracking-wider text-amber-800 flex items-center gap-1.5">
              <i class="fa-solid fa-utensils text-amber-600"></i> Extra Food / Drink Orders List
            </h4>
            <button type="button" id="btn-add-food-order" onclick="addFoodOrderItem()" class="bg-amber-600 hover:bg-amber-700 text-white px-2.5 py-1 rounded-full text-[10px] font-semibold flex items-center gap-1 transition shadow-sm">
              <i class="fa-solid fa-plus text-[9px]"></i> Add Food Order
            </button>
          </div>
          
          <div id="food-orders-container" class="space-y-2 max-h-40 overflow-y-auto pr-1"></div>
        </div>

        <!-- CAB FARE SECTION -->
        <div id="sec-cab-fare" class="bg-indigo-50/40 p-3 rounded-2xl border border-indigo-200/80 space-y-2.5 transition-all">
          <div class="flex justify-between items-center">
            <h4 class="text-[9px] font-bold uppercase tracking-wider text-indigo-700 flex items-center gap-1.5">
              <i class="fa-solid fa-taxi text-indigo-600"></i> Cab Fare Details
            </h4>
            <button type="button" id="btn-add-cab-trip" onclick="addCabTripRow()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded-full text-[10px] font-semibold flex items-center gap-1 transition shadow-sm">
              <i class="fa-solid fa-plus text-[9px]"></i> Add Cab Trip
            </button>
          </div>
          <div id="cab-trips-container" class="space-y-2 max-h-40 overflow-y-auto pr-1 mt-2"></div>
        </div>

        <!-- Billing Calculation Box -->
        <div id="sec-billing-summary" class="bg-blue-50/40 p-3 rounded-2xl border border-blue-100 space-y-2.5 transition-all">
          <h4 class="text-[9px] font-bold uppercase tracking-wider text-blue-700 flex items-center gap-1.5">
            <i class="fa-solid fa-calculator text-blue-600"></i> Billing Summary
          </h4>
          
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-2 pb-2 border-b border-blue-200">
             <div>
                <label class="block font-semibold text-slate-600 mb-0.5">Extra Person (₹)</label>
                <input type="number" id="cust-extra-total" readonly="" class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
             </div>
             <div>
                <label class="block font-semibold text-slate-600 mb-0.5">Cab Fare (₹)</label>
                <input type="number" id="cust-cab-total" readonly="" class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
             </div>
          </div>
          
          <div class="grid grid-cols-2 sm:grid-cols-6 gap-2">
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Days</label>
              <input type="number" id="cust-days" readonly="" class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Price/Day (₹)</label>
              <input type="number" id="cust-price" value="1200" oninput="calculateModalBilling()" class="w-full bg-white font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Total (₹)</label>
              <input type="number" id="cust-total" readonly="" class="w-full bg-slate-200/60 text-blue-700 font-bold border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Advance (₹)</label>
              <input type="number" id="cust-advance" value="0" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-semibold text-emerald-600" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Due (₹)</label>
              <input type="number" id="cust-due" readonly="" class="w-full bg-slate-200/60 text-rose-700 font-bold border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5 text-[10px] text-emerald-700">Clear Bill (₹)</label>
              <input type="number" id="cust-clear-bill" value="0" placeholder="0" oninput="handleClearBillPayment(this.value)" class="w-full bg-emerald-50 border border-emerald-300 font-bold text-emerald-800 rounded-xl px-2 py-1.5 focus:outline-none focus:border-emerald-500" title="Put payment amount to clear due bill" />
            </div>
          </div>
        </div>

        <div class="flex justify-end space-x-2 pt-1">
          <button type="button" onclick="closeBookingModal()" class="px-4 py-1.5 bg-slate-100 text-slate-700 rounded-xl font-semibold transition hover:bg-slate-200">Cancel</button>
          <button type="submit" id="btn-save-booking" class="px-5 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl font-semibold shadow-sm transition">Save Booking</button>
        </div>
      </form>
    </div>
  </div>

   <!-- FIXED & PRINTABLE INVOICE / BOOKING RECEIPT MODAL -->
  <div id="invoice-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-3 sm:p-6 overflow-y-auto">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-xl w-full p-5 sm:p-6 space-y-4 relative my-auto max-h-[92vh] overflow-y-auto modal-fadeIn" id="printable-invoice">
      
      <!-- Read-Only Notice Bar -->
      <div id="inv-readonly-notice" class="hidden bg-slate-900 text-amber-300 text-[10px] font-bold px-3.5 py-2 rounded-2xl flex items-center justify-between border border-slate-800 no-print">
        <span class="flex items-center gap-1.5">
          <i class="fa-solid fa-lock text-amber-400"></i> Read-Only View Mode (Editing Disabled)
        </span>
        <span class="text-[9px] text-slate-400 font-normal">System Protected</span>
      </div>

      <div class="flex justify-between items-start border-b border-slate-200 pb-3">
        <div>
          <h2 class="text-base sm:text-lg font-black text-blue-600 uppercase tracking-wide">Aniruddha Homestay</h2>
          <p class="text-[10px] text-slate-500 mt-0.5">Sittong, Village in West Bengal</p>
          <p class="text-[10px] text-slate-500">Phone: +91 9804396541 | Email: info@businessportal.com</p>
        </div>
        <div class="text-right">
          <div id="e-invoice-section">
            <span id="inv-badge" class="inline-block bg-blue-50 text-blue-700 text-[9px] font-bold px-2.5 py-0.5 rounded-full uppercase mb-1 border border-blue-100">e-Invoice</span>
            <p id="inv-id-container" class="text-[10px] text-slate-500">Invoice ID: <strong id="inv-id" class="text-slate-800 font-mono">INV-2026-0000001</strong></p>
          </div>
          <p class="text-[10px] text-slate-500">Booking ID: <strong id="inv-booking-id" class="text-blue-600 font-mono">BKG-2026-0000001</strong></p>
          <p class="text-[10px] text-slate-500">Issued On: <strong id="inv-date" class="text-slate-800"></strong></p>
        </div>
      </div>

      <div class="grid grid-cols-2 gap-3 bg-slate-50 p-3 rounded-2xl border border-slate-100 text-[10px] sm:text-[11px]">
        <div class="space-y-0.5">
          <h4 class="font-bold text-slate-400 uppercase text-[9px] tracking-wider mb-1">Guest Information</h4>
          <p class="text-slate-800 font-bold" id="inv-guest-name">-</p>
          <p class="text-slate-600 leading-tight" id="inv-guest-address">Address: -</p>
          <p class="text-slate-600" id="inv-guest-contact">Contact: -</p>
          <p class="text-slate-600" id="inv-guest-id">ID No: -</p>
        </div>
        <div class="space-y-0.5">
          <h4 class="font-bold text-slate-400 uppercase text-[9px] tracking-wider mb-1">Reservation Info</h4>
          <p class="text-slate-800 font-bold" id="inv-room">Room No: -</p>
          <p class="text-slate-600" id="inv-checkin">Check-in: -</p>
          <div id="inv-checkout-container" class="space-y-0.5">
            <p class="text-slate-600 font-medium" id="inv-checkout">Check-out: -</p>
            <p class="text-blue-600 font-bold hidden" id="inv-ext-checkout"></p>
          </div>
        </div>
      </div>

      <div class="overflow-x-auto border border-slate-100 rounded-2xl">
        <table class="w-full text-left text-[10px] sm:text-[11px]">
          <thead>
            <tr class="bg-blue-50/70 text-blue-900 border-b border-blue-100">
              <th class="p-2 sm:p-2.5 font-bold">Description</th>
              <th class="p-2 sm:p-2.5 text-center font-bold">Qty / Duration</th>
              <th class="p-2 sm:p-2.5 text-right font-bold">Rate/Day</th>
              <th class="p-2 sm:p-2.5 text-right font-bold">Total Amount</th>
            </tr>
          </thead>
          <tbody id="inv-items-tbody" class="divide-y divide-slate-100"></tbody>
        </table>
      </div>

      <div class="flex justify-end pt-1">
        <div class="w-1/2 space-y-1 text-[10px] sm:text-[11px]">
          <div class="flex justify-between text-slate-600">
            <span>Total Amount:</span>
            <strong id="inv-sum-total" class="text-slate-800">₹0</strong>
          </div>
          <div class="flex justify-between text-emerald-600">
            <span>Advance Payment:</span>
            <strong id="inv-sum-advance">₹0</strong>
          </div>
          <div class="flex justify-between text-rose-600 font-bold border-t border-slate-200 pt-1">
            <span>Balance Due:</span>
            <span id="inv-sum-due">₹0</span>
          </div>
          <div id="inv-clear-due-row" class="hidden flex justify-between text-emerald-700 font-bold border-t border-slate-100 pt-1">
            <span>Clear Due:</span>
            <strong id="inv-sum-clear-due">₹0</strong>
          </div>
        </div>
      </div>

      <div class="pt-4 border-t border-slate-200/80 flex justify-between items-end text-[10px] text-slate-400">
        <div>
          <p class="font-bold text-slate-600">Thank you for staying with us!</p>
          <p>For inquiries, please contact hotel management.</p>
        </div>
        <div class="text-center border-t border-slate-300 pt-1 w-28">
          <p class="font-semibold text-slate-600">Authorized Signature</p>
        </div>
      </div>

      <div class="flex flex-wrap justify-end space-x-2 gap-y-2 pt-2 no-print border-t border-slate-100">
        <button type="button" onclick="closeInvoiceModal()" class="px-4 py-1.5 bg-slate-100 text-slate-700 rounded-xl font-semibold transition hover:bg-slate-200">Close</button>
        <button type="button" id="inv-whatsapp-btn" onclick="sendReceiptViaWhatsApp()" class="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-semibold shadow-sm flex items-center gap-1.5 transition cursor-pointer">
          <i class="fa-brands fa-whatsapp text-sm"></i> Send receipt via WhatsApp
        </button>
        <button type="button" onclick="window.print()" id="inv-print-btn" class="px-4 py-1.5 bg-blue-600 text-white rounded-xl font-semibold shadow-sm flex items-center gap-1 transition hover:bg-blue-700 cursor-pointer">
          <i class="fa-solid fa-print"></i> Print Invoice
        </button>
      </div>
    </div>
  </div>

  <!-- Javascript Logic Modules -->
  <script>
    // --------------------------------------------------
    // SYSTEM CONSTANTS & HELPER UTILITIES
    // --------------------------------------------------
    const MAX_SHEET_ROWS = 10000000;
    const GAS_API_URL = "https://script.google.com/macros/s/AKfycbx1y0ZQcPX9v66ddWlU8B5xCnOpgGvld39iY3EVNzKQ9tcNcod2onajvq0fM2p6pqExqQ/exec";
    const ONE_HOUR_MS = 1 * 60 * 60 * 1000;
    const INACTIVITY_LIMIT_MS = 10 * 60 * 1000;
    const WARNING_BUFFER_MS = 1 * 60 * 1000;

    const DEFAULT_USER_ID = "Admin";
    const DEFAULT_PASSWORD = "Aadmin123";

    let activeModalBooking = null;
    let isLoggedIn = false;
    let isMasterUnlocked = false;
    let inactivityTimer = null;
    let warningTimer = null;
    let countdownInterval = null;

    let pendingMasterDeleteType = null;
    let pendingMasterDeleteTarget = null;

    const currentRealYear = new Date().getFullYear();
    const defaultAppYear = currentRealYear >= 2026 && currentRealYear <= 2085 ? currentRealYear : 2026;

    let state = {
      yearlyCounters: { [defaultAppYear]: 0 },
      bookings: [],
      roomsCapacity: [
        { roomNo: 1, capacity: 4 },
        { roomNo: 2, capacity: 2 },
        { roomNo: 3, capacity: 4 },
        { roomNo: 4, capacity: 4 },
        { roomNo: 5, capacity: 4 }
      ],
      masterAgents: [
        { agentName: "Self", phone: "Direct", roomNo: "All Rooms" }
      ],
      selectedYear: defaultAppYear,
      dashSelectedYear: defaultAppYear
    };

    // Smooth Page Scrolling Helper
    function scrollPage(direction) {
      const scrollStep = window.innerHeight * 0.7;
      const horizontalStep = window.innerWidth * 0.5;

      switch(direction) {
        case 'up':
          window.scrollBy({ top: -scrollStep, behavior: 'smooth' });
          break;
        case 'down':
          window.scrollBy({ top: scrollStep, behavior: 'smooth' });
          break;
        case 'left':
          window.scrollBy({ left: -horizontalStep, behavior: 'smooth' });
          break;
        case 'right':
          window.scrollBy({ left: horizontalStep, behavior: 'smooth' });
          break;
      }
    }

    function isTrue(val) {
      return val === true || val === 'true' || val === 'TRUE' || val === 1 || val === '1';
    }

    function isInactiveBooking(b) {
      return isTrue(b && b.inactive);
    }

    function getBookingRooms(b) {
      if (!b || b.roomNo === undefined || b.roomNo === null) return [];
      if (Array.isArray(b.roomNo)) return b.roomNo.map(r => String(r).trim());
      return String(b.roomNo).split(/[,|]/).map(s => s.trim());
    }

    function parseDateMs(dtStr) {
      if (!dtStr) return NaN;
      if (typeof dtStr === 'number') return dtStr;
      let sanitized = String(dtStr).trim().replace(' ', 'T');
      let d = new Date(sanitized);
      if (isNaN(d.getTime())) {
        d = new Date(dtStr);
      }
      return d.getTime();
    }

    function extractISTDateParts(dtStr) {
      if (!dtStr) return { date: '', time: '' };
      let d = new Date(typeof dtStr === 'string' ? dtStr.replace(' ', 'T') : dtStr);
      if (isNaN(d.getTime())) {
        const parts = String(dtStr).replace(' ', 'T').split('T');
        return { date: parts[0] || '', time: parts[1] ? parts[1].substring(0, 5) : '' };
      }
      const utcMs = d.getTime();
      const istDate = new Date(utcMs + (330 * 60000));
      
      const yyyy = istDate.getUTCFullYear();
      const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
      const dd = String(istDate.getUTCDate()).padStart(2, '0');
      const hh = String(istDate.getUTCHours()).padStart(2, '0');
      const min = String(istDate.getUTCMinutes()).padStart(2, '0');
      
      return { date: `${yyyy}-${mm}-${dd}`, time: `${hh}:${min}` };
    }

    function toLocalISOString(date) {
      if (!(date instanceof Date) || isNaN(date)) return '';
      const utcMs = date.getTime();
      const istDate = new Date(utcMs + (330 * 60000));
      
      const yyyy = istDate.getUTCFullYear();
      const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
      const dd = String(istDate.getUTCDate()).padStart(2, '0');
      const hh = String(istDate.getUTCHours()).padStart(2, '0');
      const min = String(istDate.getUTCMinutes()).padStart(2, '0');
      return `${yyyy}-${mm}-${dd}T${hh}:${min}:00+05:30`;
    }

    function format24hDate(dtStr) {
      if (!dtStr) return '';
      const d = new Date(typeof dtStr === 'string' ? dtStr.replace(' ', 'T') : dtStr);
      if (isNaN(d.getTime())) return String(dtStr);
      
      const utcMs = d.getTime();
      const istDate = new Date(utcMs + (330 * 60000));
      
      const dd = String(istDate.getUTCDate()).padStart(2, '0');
      const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
      const yy = String(istDate.getUTCFullYear()).slice(-2);
      const hh = String(istDate.getUTCHours()).padStart(2, '0');
      const min = String(istDate.getUTCMinutes()).padStart(2, '0');
      return `${dd}/${mm}/${yy} ${hh}:${min}`;
    }

    function parseJSONField(fieldData) {
      if (!fieldData) return [];
      if (Array.isArray(fieldData)) return fieldData;
      if (typeof fieldData === 'string' && fieldData.length > 5) {
        try { return JSON.parse(fieldData); } catch (e) {}
      }
      return [];
    }

    function checkSheetRowLimits() {
      const currentRowCount = state.bookings.length;
      if (currentRowCount >= (MAX_SHEET_ROWS - 20)) {
        alert("CRITICAL NOTIFICATION: Google Sheet has reached its row limit (less than 20 rows left). Saving has been stopped.");
        return false;
      } else if (currentRowCount >= (MAX_SHEET_ROWS - 100)) {
        alert("WARNING: Google Sheet is approaching its row limit.");
      }
      return true;
    }

    // --------------------------------------------------
    // AUTHENTICATION AND SYSTEM TIMERS
    // --------------------------------------------------
    function checkAuthStatus() {
      const sessionAuth = sessionStorage.getItem('app_authenticated');
      if (sessionAuth === 'true') {
        isLoggedIn = true;
        document.getElementById('login-overlay').classList.add('hidden');
        startInactivityMonitoring();
      } else {
        isLoggedIn = false;
        document.getElementById('login-overlay').classList.remove('hidden');
      }
    }

    function handleLogin(e) {
      e.preventDefault();
      const user = document.getElementById('login-userid').value.trim();
      const pass = document.getElementById('login-password').value.trim();

      if (user === DEFAULT_USER_ID && pass === DEFAULT_PASSWORD) {
        isLoggedIn = true;
        sessionStorage.setItem('app_authenticated', 'true');
        document.getElementById('login-overlay').classList.add('hidden');
        document.getElementById('login-error').classList.add('hidden');
        startInactivityMonitoring();
        document.getElementById('login-alert-modal').classList.remove('hidden');
      } else {
        document.getElementById('login-error').classList.remove('hidden');
      }
    }

    function logoutUser(isAuto = false) {
      if (isAuto) {
        processLogoutWithSave();
      } else {
        document.getElementById('logout-confirm-modal').classList.remove('hidden');
      }
    }

    function cancelLogout() {
      document.getElementById('logout-confirm-modal').classList.add('hidden');
      resetInactivityTimer();
    }

    async function processLogoutWithSave() {
      document.getElementById('logout-warning-modal').classList.add('hidden');
      document.getElementById('logout-confirm-modal').classList.add('hidden');
      
      if (isLoggedIn) {
        document.getElementById('saving-lock-modal').classList.remove('hidden');
        try {
          await saveChanges(true, true);
        } catch (e) {
          console.error("Save on logout error", e);
        }
      }
      
      isLoggedIn = false;
      isMasterUnlocked = false;
      sessionStorage.removeItem('app_authenticated');
      stopInactivityMonitoring();
      window.location.reload();
    }

    function startInactivityMonitoring() {
      stopInactivityMonitoring();
      const activityEvents = ['mousemove', 'keydown', 'mousedown', 'touchstart', 'scroll'];
      activityEvents.forEach(evt => {
        window.addEventListener(evt, resetInactivityTimer);
      });
      resetInactivityTimer();
    }

    function stopInactivityMonitoring() {
      if (inactivityTimer) clearTimeout(inactivityTimer);
      if (warningTimer) clearTimeout(warningTimer);
      if (countdownInterval) clearInterval(countdownInterval);
      const activityEvents = ['mousemove', 'keydown', 'mousedown', 'touchstart', 'scroll'];
      activityEvents.forEach(evt => {
        window.removeEventListener(evt, resetInactivityTimer);
      });
    }

    function resetInactivityTimer() {
      if (!isLoggedIn) return;
      if (inactivityTimer) clearTimeout(inactivityTimer);
      if (warningTimer) clearTimeout(warningTimer);
      if (countdownInterval) clearInterval(countdownInterval);

      document.getElementById('logout-warning-modal').classList.add('hidden');
      warningTimer = setTimeout(showInactivityWarning, INACTIVITY_LIMIT_MS - WARNING_BUFFER_MS);
      inactivityTimer = setTimeout(() => logoutUser(true), INACTIVITY_LIMIT_MS);
    }

    function showInactivityWarning() {
      if (!isLoggedIn) return;
      let secondsLeft = 60;
      document.getElementById('logout-countdown-seconds').innerText = secondsLeft;
      document.getElementById('logout-warning-modal').classList.remove('hidden');

      countdownInterval = setInterval(() => {
        secondsLeft--;
        if (secondsLeft >= 0) {
          document.getElementById('logout-countdown-seconds').innerText = secondsLeft;
        } else {
          clearInterval(countdownInterval);
        }
      }, 1000);
    }

    // --------------------------------------------------
    // MASTER CONTROLS & WIPING MODULES
    // --------------------------------------------------
    function openMasterAuthModal() {
      document.getElementById('master-password-input').value = '';
      document.getElementById('master-auth-error').classList.add('hidden');
      document.getElementById('master-auth-modal').classList.remove('hidden');
    }

    function closeMasterAuthModal() {
      document.getElementById('master-auth-modal').classList.add('hidden');
    }

    function handleMasterAuth(e) {
      e.preventDefault();
      const enteredPass = document.getElementById('master-password-input').value.trim();
      if (enteredPass === DEFAULT_PASSWORD) {
        isMasterUnlocked = true;
        closeMasterAuthModal();
        performSwitchTab('master');
      } else {
        document.getElementById('master-auth-error').classList.remove('hidden');
      }
    }

    function requestDataWipe() {
      if (!isMasterUnlocked) {
        alert("You must unlock Master Data access first to perform a data wipe.");
        openMasterAuthModal();
        return;
      }
      document.getElementById('wipe-layer-1-modal').classList.remove('hidden');
    }

    function proceedToWipeLayer2() {
      document.getElementById('wipe-layer-1-modal').classList.add('hidden');
      document.getElementById('wipe-layer-2-modal').classList.remove('hidden');
    }

    function closeWipeModals() {
      document.getElementById('wipe-layer-1-modal').classList.add('hidden');
      document.getElementById('wipe-layer-2-modal').classList.add('hidden');
    }

    async function executeGoogleSheetWipe() {
      const btn = document.getElementById('btn-final-wipe');
      btn.innerText = "WIPING DATA...";
      btn.disabled = true;

      try {
        const payload = { action: "wipeData" };
        const response = await fetch(GAS_API_URL, {
          method: "POST",
          body: JSON.stringify(payload)
        });

        const textResult = await response.text();
        try { JSON.parse(textResult); } catch(e) {}

        state.bookings = [];
        state.yearlyCounters = {};
        state.roomsCapacity = [
          { roomNo: 1, capacity: 4 },
          { roomNo: 2, capacity: 2 },
          { roomNo: 3, capacity: 4 },
          { roomNo: 4, capacity: 4 },
          { roomNo: 5, capacity: 4 }
        ];
        state.masterAgents = [{ agentName: "Self", phone: "Direct", roomNo: "All Rooms" }];
        
        refreshAllUI();
        closeWipeModals();
        alert("Database has been completely wiped.");
      } catch (error) {
        console.error("Wipe error:", error);
        alert("Failed to wipe database. Please check your connection.");
      } finally {
        btn.innerText = "ERASE ALL DATA";
        btn.disabled = false;
      }
    }

    // --------------------------------------------------
    // FORMATTING UTILITIES & DISPLAY CONVERTERS
    // --------------------------------------------------
    function formatTitleCase(text) {
      if (!text) return '';
      return String(text).replace(/\w\S*/g, function(txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
      });
    }

    function formatDateTime(dtStr) {
      if (!dtStr) return '-';
      const d = new Date(typeof dtStr === 'string' ? dtStr.replace(' ', 'T') : dtStr);
      if (isNaN(d.getTime())) return String(dtStr).replace('T', ' ');
      
      const utcMs = d.getTime();
      const istDate = new Date(utcMs + (330 * 60000));
      
      const day = String(istDate.getUTCDate()).padStart(2, '0');
      const month = String(istDate.getUTCMonth() + 1).padStart(2, '0');
      const year = istDate.getUTCFullYear();
      const hours = String(istDate.getUTCHours()).padStart(2, '0');
      const minutes = String(istDate.getUTCMinutes()).padStart(2, '0');
      return `${day}-${month}-${year} ${hours}:${minutes}`;
    }

    function formatDate(d) {
      if (!d) return '-';
      const dateObj = typeof d === 'string' ? new Date(d.replace(' ', 'T')) : d;
      if (isNaN(dateObj.getTime())) return d;
      const day = String(dateObj.getDate()).padStart(2, '0');
      const month = String(dateObj.getMonth() + 1).padStart(2, '0');
      const year = dateObj.getFullYear();
      return `${day}-${month}-${year}`;
    }

    function getEffectiveCheckoutTime(b) {
      if (!b) return 0;
      if (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) {
        return parseDateMs(b.extendedCheckOut);
      }
      return parseDateMs(b.checkOut);
    }

    function isRoomInMaster(roomNo) {
      if (!state.roomsCapacity) return true;
      let rooms = getBookingRooms({ roomNo });
      return rooms.every(r => state.roomsCapacity.some(m => String(m.roomNo) === String(r)));
    }

    // --------------------------------------------------
    // DASHBOARD VIEWS & CARD CALCULATIONS
    // --------------------------------------------------
    function populateDashboardYearDropdown() {
      const yearSelect = document.getElementById('dash-year-select');
      if (!yearSelect) return;
      yearSelect.innerHTML = '';

      const optConsolidated = document.createElement('option');
      optConsolidated.value = "ALL";
      optConsolidated.text = "All Years (Consolidated)";
      yearSelect.appendChild(optConsolidated);

      for (let y = 2026; y <= 2085; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.text = y === defaultAppYear ? `${y} (Current Year)` : `Year ${y}`;
        yearSelect.appendChild(opt);
      }

      yearSelect.value = defaultAppYear;
      state.dashSelectedYear = defaultAppYear;
    }

    function handleDashboardYearChange(val) {
      if (val === 'CURRENT') val = defaultAppYear;
      const select = document.getElementById('dash-year-select');
      if (select) select.value = val;

      if (val === 'ALL') {
        state.dashSelectedYear = 'ALL';
      } else {
        state.dashSelectedYear = parseInt(val);
      }
      initDashboard();
    }

    function initDashboard() {
      const grid = document.getElementById('years-grid');
      grid.innerHTML = '';
      
      for (let y = 2026; y <= 2085; y++) {
        const item = document.createElement('div');
        const isSelectedYear = state.dashSelectedYear !== 'ALL' && parseInt(state.dashSelectedYear) === y;
        const isCurrentRealYear = y === defaultAppYear;

        item.className = `text-center py-1.5 px-1 rounded-2xl text-[10px] font-bold cursor-pointer transition ${
          isSelectedYear
            ? 'bg-blue-600 text-white shadow-xs' 
            : (isCurrentRealYear ? 'bg-amber-100 text-amber-900 font-extrabold hover:bg-amber-200' : 'bg-slate-100 text-slate-600 hover:bg-blue-50 hover:text-blue-600')
        }`;
        
        item.innerText = y;
        item.onclick = () => selectDashboardYear(y);
        grid.appendChild(item);
      }
      updateDashboardCards();
    }

    function updateDashboardCards() {
      const selectedFilter = state.dashSelectedYear;
      const label = document.getElementById('dash-filter-label');

      let filteredBookings = [];
      if (selectedFilter === 'ALL' || !selectedFilter) {
        filteredBookings = state.bookings.filter(b => !isInactiveBooking(b));
        if (label) label.innerText = "Consolidated Summary (All Years)";
      } else {
        const targetYear = parseInt(selectedFilter);
        filteredBookings = state.bookings.filter(b => {
          if (isInactiveBooking(b) || !b.checkIn) return false;
          const yr = new Date(b.checkIn.replace(' ', 'T')).getFullYear();
          return yr === targetYear;
        });

        if (label) {
          label.innerText = targetYear === defaultAppYear 
            ? `Year ${targetYear} (Current Year)` 
            : `Year ${targetYear}`;
        }
      }

      const totalBookings = filteredBookings.length;
      const totalAmt = filteredBookings.reduce((sum, b) => sum + (b.totalAmount || 0), 0);
      const totalAdv = filteredBookings.reduce((sum, b) => sum + (b.initialAdv || 0) + (b.clearedDue || 0), 0);
      const totalDue = filteredBookings.reduce((sum, b) => sum + (b.totalDue || 0), 0);

      document.getElementById('dash-total-bookings').innerText = totalBookings;
      document.getElementById('dash-total-amount').innerText = `₹${totalAmt.toLocaleString('en-IN')}`;
      document.getElementById('dash-advanced').innerText = `₹${totalAdv.toLocaleString('en-IN')}`;
      document.getElementById('dash-due').innerText = `₹${totalDue.toLocaleString('en-IN')}`;

      // Calculate status breakdown counts for live, upcoming, closed, and inactive
      const now = new Date().getTime();
      let liveCount = 0;
      let upcomingCount = 0;
      let closedCount = 0;
      let inactiveCount = 0;

      let allScopedBookings = state.bookings;
      if (selectedFilter !== 'ALL' && selectedFilter) {
        const targetYear = parseInt(selectedFilter);
        allScopedBookings = state.bookings.filter(b => {
          if (!b.checkIn) return false;
          const yr = new Date(b.checkIn.replace(' ', 'T')).getFullYear();
          return yr === targetYear;
        });
      }

      allScopedBookings.forEach(b => {
        if (isInactiveBooking(b)) {
          inactiveCount++;
        } else {
          const cIn = parseDateMs(b.checkIn);
          const cOut = getEffectiveCheckoutTime(b);

          if (now > cOut) {
            closedCount++;
          } else if (now >= cIn && now <= cOut) {
            liveCount++;
          } else {
            upcomingCount++;
          }
        }
      });

      document.getElementById('dash-live-count').innerText = liveCount;
      document.getElementById('dash-upcoming-count').innerText = upcomingCount;
      document.getElementById('dash-closed-count').innerText = closedCount;
      document.getElementById('dash-inactive-count').innerText = inactiveCount;
    }

    function selectDashboardYear(year) {
      handleDashboardYearChange(year);
      renderCalendar(year);
      switchTab('calendar');
    }

    // --------------------------------------------------
    // BOOKINGS TABLE AND DIRECTORY RENDERERS
    // --------------------------------------------------
    function renderBookingsTable(dateFilter = "") {
      const tbody = document.getElementById('bookings-tbody');
      tbody.innerHTML = '';

      let listToRender = [...state.bookings];
      if (dateFilter) {
        listToRender = listToRender.filter(b => {
          if (!b.checkIn || !b.checkOut) return false;
          const bIn = String(b.checkIn).replace(' ', 'T').split('T')[0];
          const bOutVal = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
          const bOut = String(bOutVal).replace(' ', 'T').split('T')[0];
          return (dateFilter >= bIn && dateFilter <= bOut);
        });
      }

      const now = new Date().getTime();
      const getStatusPriority = (b) => {
        const isMasterValid = isRoomInMaster(b.roomNo);
        if (!isMasterValid || isInactiveBooking(b)) return 4;
        const cIn = parseDateMs(b.checkIn);
        const cOut = getEffectiveCheckoutTime(b);

        if (now > cOut) return 3;
        else if (now >= cIn && now <= cOut) return 1;
        else return 2;
      };

      listToRender.sort((a, b) => {
        const priorityA = getStatusPriority(a);
        const priorityB = getStatusPriority(b);
        if (priorityA !== priorityB) return priorityA - priorityB;
        return (b.bookingCode || '').localeCompare(a.bookingCode || '');
      });

      if (listToRender.length === 0) {
        tbody.innerHTML = `<tr><td colspan="13" class="text-center py-6 text-slate-400">No bookings found for the selected date.</td></tr>`;
        return;
      }

      listToRender.forEach((b) => {
        const isMasterValid = isRoomInMaster(b.roomNo);
        const checkInFmt = formatDateTime(b.checkIn);
        const effectiveOut = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
        const checkOutFmt = formatDateTime(effectiveOut);

        const checkInTime = parseDateMs(b.checkIn);
        const checkOutTime = getEffectiveCheckoutTime(b);

        const isClosed = now > checkOutTime;
        const isInactive = isInactiveBooking(b);
        const roomsDisplay = getBookingRooms(b).join(', ');

        let statusBgClass = "hover:bg-slate-50";
        let statusDotHtml = "";

        if (!isMasterValid) {
          statusBgClass = "bg-rose-100/70 hover:bg-rose-200/60 text-rose-900";
        } else if (isInactive) {
          statusBgClass = "bg-slate-100 hover:bg-slate-200 text-slate-500 opacity-75";
        } else if (isClosed) {
          statusDotHtml = `<span class="w-2.5 h-2.5 bg-emerald-500 rounded-full inline-block flex-shrink-0" title="Closed Booking"></span>`;
        } else if (now >= checkInTime && now <= checkOutTime) {
          statusDotHtml = `
            <span class="relative flex h-2.5 w-2.5 flex-shrink-0" title="Live Booking">
              <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
              <span class="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500"></span>
            </span>
          `;
        } else if (now < checkInTime) {
          statusDotHtml = `<span class="w-2.5 h-2.5 bg-blue-500 rounded-full inline-block flex-shrink-0" title="Upcoming Booking"></span>`;
        }

        let foodSummaryHtml = '';
        const parseFood = parseJSONField(b.foodOrders);
        if (parseFood.length > 0) {
          const totalFoodCharge = parseFood.reduce((acc, fo) => acc + (fo.foodCharge || 0), 0);
          if (totalFoodCharge > 0) {
            foodSummaryHtml = `<div class="text-[9px] ${!isMasterValid ? 'text-rose-950 font-bold' : 'text-amber-800 font-semibold'}"><i class="fa-solid fa-utensils text-[8px] mr-0.5"></i>Food (${parseFood.length}): +₹${totalFoodCharge}</div>`;
          }
        }
        
        let cabSummaryHtml = '';
        const parseCab = parseJSONField(b.cabTrips);
        const totalCab = parseCab.reduce((acc, t) => acc + (t.rate || 0), 0);
        if (totalCab > 0) {
          cabSummaryHtml = `<div class="text-[9px] ${!isMasterValid ? 'text-rose-950 font-bold' : 'text-indigo-800 font-semibold'}"><i class="fa-solid fa-taxi text-[8px] mr-0.5"></i>Cab: +₹${totalCab}</div>`;
        }

        const printOnClick = `printInvoice('${b.id}')`;
        let actionButtonsHtml = `
          <div class="flex items-center justify-center space-x-1">
            <button onclick="openBookingModal('${b.id}')" class="text-blue-600 hover:text-blue-800 p-1 text-sm" title="Edit Booking Details">
              <i class="fa-solid fa-pen-to-square"></i>
            </button>              
            <button onclick="${printOnClick}" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1 rounded-full text-[11px] font-bold transition shadow-xs">Print</button>
          </div>
        `;

        let idProofCellHtml = `<span class="text-slate-400 italic text-[10px]">None</span>`;
        if (b.idProofBase64) {
          idProofCellHtml = `
            <button onclick="openPdfAttachment('${b.idProofBase64}')" class="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-2 py-0.5 rounded-full text-[10px] font-bold flex items-center gap-1 transition">
              <i class="fa-solid fa-file-pdf text-rose-600"></i> View PDF
            </button>
          `;
        }

        const tableCap = parseInt(b.capacity) || 1;
        const tableCapLabel = tableCap === 1 ? 'Person' : 'Persons';
        const extraPersonsText = (b.extraPersons && b.extraPersons > 0) ? `<span class="text-amber-700 font-bold block text-[9px]">(+${b.extraPersons} Extra)</span>` : '';
        const contactDisplay = b.contactNo ? `${b.countryCode || '+91'} ${b.contactNo}`.trim() : '-';
        const totalReceived = (b.initialAdv || 0) + (b.clearedDue || 0);

        const tr = document.createElement('tr');
        tr.className = `${statusBgClass} transition border-b border-slate-100`;
        tr.innerHTML = `
          <td class="py-2.5 px-3">
            <div class="flex items-center gap-1.5">
              ${statusDotHtml}
              ${isInactive ? '<span class="w-2.5 h-2.5 bg-rose-600 rounded-full inline-block flex-shrink-0" title="Inactive Booking"></span>' : ''}
              <span class="bg-blue-50 border border-blue-200 text-blue-700 font-mono font-bold px-2 py-0.5 rounded-full text-[9px] block w-max">${b.bookingCode}</span>
            </div>
            ${isInactive ? '<span class="bg-slate-600 text-white font-bold px-1.5 py-0.2 rounded-full text-[8px] uppercase block mt-0.5 w-max">Inactive</span>' : (!isMasterValid ? '<span class="bg-rose-700 text-white font-bold px-1.5 py-0.2 rounded-full text-[8px] uppercase block mt-0.5 w-max">Master Removed</span>' : '')}
          </td>
          <td class="py-2.5 px-3 font-bold ${!isMasterValid ? 'text-rose-950' : 'text-slate-800'}">${b.name}</td>
          <td class="py-2.5 px-3 font-medium whitespace-nowrap">${contactDisplay}</td>
          <td class="py-2.5 px-3 font-mono text-[10px]">${b.idNo || '-'}</td>
          <td class="py-2.5 px-3">${idProofCellHtml}</td>
          <td class="py-2.5 px-3"><span class="bg-blue-50 text-blue-700 font-bold px-2 py-0.5 rounded-full text-[10px] break-all">Room ${roomsDisplay}</span></td>
          <td class="py-2.5 px-3 font-bold text-slate-700">${tableCap} ${tableCapLabel} ${extraPersonsText}</td>
          <td class="py-2.5 px-3 ${!isMasterValid ? 'text-rose-900' : 'text-slate-600'} text-[10px]">${b.agentInfo || '-'}</td>
          <td class="py-2.5 px-3 text-[10px]">
            <div class="font-semibold ${!isMasterValid ? 'text-rose-950' : 'text-slate-700'}">${checkInFmt}</div>
            <div class="${!isMasterValid ? 'text-rose-900' : 'text-slate-500'} text-[9px]">to ${checkOutFmt} ${isTrue(b.hasExtendedCheckout) ? '<span class="text-blue-600 font-bold">(Ext)</span>' : ''}</div>
          </td>
          <td class="py-2.5 px-3">
            <div class="font-bold ${!isMasterValid ? 'text-rose-950' : 'text-slate-800'}">₹${b.perDayPrice}/day (${b.noOfDays}d)</div>
            ${foodSummaryHtml}
            ${cabSummaryHtml}
          </td>
          <td class="py-2.5 px-3 font-bold ${!isMasterValid ? 'text-rose-950' : 'text-slate-800'}">
            ₹${b.totalAmount}
            <span class="block text-[9px] text-emerald-600 font-medium">Adv: ₹${totalReceived}</span>
          </td>
          <td class="py-2.5 px-3 font-bold ${b.totalDue > 0 ? 'text-rose-600' : 'text-emerald-600'}">₹${b.totalDue}</td>
          <td class="py-2.5 px-3 text-center">
            ${actionButtonsHtml}
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    // --------------------------------------------------
    // MASTER DATA MANAGEMENT FUNCTIONS
    // --------------------------------------------------
    function renderRoomCapacityTable() {
      const tbody = document.getElementById('room-capacity-tbody');
      tbody.innerHTML = '';
      if (!state.roomsCapacity || state.roomsCapacity.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" class="text-center py-4 text-slate-400">No room capacity data available.</td></tr>`;
        return;
      }

      state.roomsCapacity.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2.5 px-3">
            <input type="number" value="${r.roomNo}" min="1" oninput="state.roomsCapacity[${idx}].roomNo = parseInt(this.value) || 1" onchange="populateRoomDropdown(); renderBookingsTable(); saveChanges(false, true)" class="w-24 bg-transparent font-bold text-blue-600 focus:bg-white focus:border focus:border-blue-300 rounded-xl px-2 py-1">
          </td>
          <td class="py-2.5 px-3">
            <input type="number" value="${r.capacity}" min="1" oninput="state.roomsCapacity[${idx}].capacity = parseInt(this.value) || 1" onchange="saveChanges(false, true)" class="w-24 bg-transparent font-semibold text-slate-800 focus:bg-white focus:border focus:border-blue-300 rounded-xl px-2 py-1">
          </td>
          <td class="py-2.5 px-3 text-center">
            <button type="button" onclick="removeRoomCapacityRow(${idx})" class="text-rose-500 hover:text-rose-700 p-1 text-xs" title="Delete Room Entry">
              <i class="fa-solid fa-trash-can"></i>
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function renderMasterAgentTable() {
      const tbody = document.getElementById('agent-tbody');
      tbody.innerHTML = '';

      if (!state.masterAgents || state.masterAgents.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="text-center py-4 text-slate-400">No agent data available.</td></tr>`;
        return;
      }

      state.masterAgents.forEach((a, idx) => {
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2.5 px-3">
            <input type="text" value="${a.agentName}" oninput="state.masterAgents[${idx}].agentName = formatTitleCase(this.value)" onchange="populateAgentDropdown(); saveChanges(false, true)" class="w-full bg-transparent font-semibold text-slate-800 focus:bg-white focus:border focus:border-blue-300 rounded-xl px-2 py-1">
          </td>
          <td class="py-2.5 px-3">
            <input type="text" value="${a.phone}" oninput="state.masterAgents[${idx}].phone = this.value" onchange="populateAgentDropdown(); saveChanges(false, true)" class="w-full bg-transparent text-slate-600 focus:bg-white focus:border focus:border-blue-300 rounded-xl px-2 py-1">
          </td>
          <td class="py-2.5 px-3 font-bold text-blue-600">
            ${a.roomNo || 'All Rooms'}
          </td>
          <td class="py-2.5 px-3 text-center">
            <button type="button" onclick="removeAgentRow(${idx})" class="text-rose-500 hover:text-rose-700 p-1 text-xs" title="Delete Agent Entry">
              <i class="fa-solid fa-trash-can"></i>
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function addRoomCapacityRow() {
      if (!state.roomsCapacity) state.roomsCapacity = [];
      const existingRoomNos = state.roomsCapacity.map(r => parseInt(r.roomNo) || 0);
      let nextRoom = 1;
      if (existingRoomNos.length > 0) nextRoom = Math.max(...existingRoomNos) + 1;

      state.roomsCapacity.push({ roomNo: nextRoom, capacity: 4 });
      renderRoomCapacityTable();
      populateRoomDropdown();
      saveChanges(false, false);
    }

    function removeRoomCapacityRow(index) {
      openMasterDeleteModal('room', index);
    }

    function addAgentRow() {
      if (!state.masterAgents) state.masterAgents = [];
      const nextNum = state.masterAgents.length;
      state.masterAgents.push({
        agentName: `Agent ${nextNum + 1}`,
        phone: "1234567890",
        roomNo: "All Rooms"
      });
      renderMasterAgentTable();
      populateAgentDropdown();
      saveChanges(false, false);
    }

    function removeAgentRow(index) {
      openMasterDeleteModal('agent', index);
    }

    // --------------------------------------------------
    // CALENDAR VIEW MODULES
    // --------------------------------------------------
    function populateCalendarYearDropdown() {
      const yearSelect = document.getElementById('cal-year-select');
      if (!yearSelect) return;
      yearSelect.innerHTML = '';
      for (let y = 2026; y <= 2085; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.text = y === defaultAppYear ? `${y} (Current Year)` : `Year ${y}`;
        if (y === defaultAppYear) opt.selected = true;
        yearSelect.appendChild(opt);
      }
    }

    function renderCalendar(year) {
      state.selectedYear = year;
      const calSelect = document.getElementById('cal-year-select');
      if (calSelect) calSelect.value = year;

      const container = document.getElementById('calendar-container');
      container.innerHTML = '';
      const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

      months.forEach((monthName, monthIndex) => {
        const monthBox = document.createElement('div');
        monthBox.className = "bg-slate-50/80 rounded-3xl p-3 border border-slate-200/60 shadow-xs flex flex-col justify-between";

        const title = document.createElement('h4');
        title.className = "font-bold text-slate-800 text-[11px] mb-2 pb-1 border-b border-slate-200/60 flex justify-between items-center px-1";
        title.innerHTML = `<span>${monthName}</span> <span class="text-[9px] text-blue-600 font-mono font-normal">${year}</span>`;
        monthBox.appendChild(title);

        const grid = document.createElement('div');
        grid.className = "grid grid-cols-7 gap-1 text-center text-[9px] font-medium text-slate-500 mb-1";
        
        ['S', 'M', 'T', 'W', 'T', 'F', 'S'].forEach(d => {
          const dh = document.createElement('div');
          dh.innerText = d;
          dh.className = "font-bold text-slate-400";
          grid.appendChild(dh);
        });

        const firstDay = new Date(year, monthIndex, 1).getDay();
        const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

        for (let i = 0; i < firstDay; i++) {
          const empty = document.createElement('div');
          grid.appendChild(empty);
        }

        const todayObj = new Date();
        const isCurrentYearAndMonth = todayObj.getFullYear() === year && todayObj.getMonth() === monthIndex;

        for (let day = 1; day <= daysInMonth; day++) {
          const cell = document.createElement('div');
          const dateStr = `${year}-${String(monthIndex + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
          
          const matchingBookings = state.bookings.filter(b => {
            if (isInactiveBooking(b) || !b.checkIn || !b.checkOut) return false;
            if (!isRoomInMaster(b.roomNo)) return false;

            const bIn = String(b.checkIn).replace(' ', 'T').split('T')[0];
            const bOutVal = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
            const bOut = String(bOutVal).replace(' ', 'T').split('T')[0];

            return (dateStr >= bIn && dateStr <= bOut);
          });

          const isBooked = matchingBookings.length > 0;
          const isToday = isCurrentYearAndMonth && todayObj.getDate() === day;

          cell.className = `py-1 rounded-xl text-[10px] font-bold cursor-pointer transition relative flex items-center justify-center ${
            isToday 
              ? 'ring-2 ring-blue-600 ring-offset-1 z-10' 
              : ''
          } ${
            isBooked 
              ? 'bg-amber-400 text-slate-900 hover:bg-amber-500 shadow-xs' 
              : 'bg-white text-slate-700 hover:bg-blue-50 hover:text-blue-600 border border-slate-100'
          }`;

          cell.innerText = day;
          if (isBooked) {
            cell.onclick = (e) => {
              e.stopPropagation();
              showExcelCommentBox(e, dateStr, matchingBookings);
            };
          }
          grid.appendChild(cell);
        }
        monthBox.appendChild(grid);
        container.appendChild(monthBox);
      });
    }

    function showExcelCommentBox(e, dateStr, bookings) {
      const box = document.getElementById('excel-comment-box');
      const dateHeader = document.getElementById('comm-date-header');
      const listContainer = document.getElementById('comm-booking-list');

      dateHeader.innerText = formatDate(dateStr);
      listContainer.innerHTML = '';
      const now = new Date().getTime();

      bookings.forEach(b => {
        const item = document.createElement('div');
        item.className = "bg-slate-800 p-2.5 rounded-2xl border border-slate-700 space-y-1 hover:border-blue-400 transition cursor-pointer";
        item.onclick = () => {
          closeCommentBox();
          openBookingModal(b.id);
        };

        const cInMs = parseDateMs(b.checkIn);
        const cOutMs = getEffectiveCheckoutTime(b);
        const effectiveCheckout = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
        const roomsDisplay = getBookingRooms(b).join(', ');

        let statusText = "Upcoming";
        let statusColorClass = "text-blue-400 font-bold";

        if (now > cOutMs) {
          statusText = "Closed";
          statusColorClass = "text-emerald-400 font-bold";
        } else if (now >= cInMs && now <= cOutMs) {
          statusText = "Live";
          statusColorClass = "text-amber-400 font-bold";
        }

        item.innerHTML = `
          <div class="flex justify-between items-center">
            <span class="font-bold text-blue-400 text-[11px]">${b.name}</span>
            <span class="bg-blue-900/60 text-blue-200 px-2 py-0.5 rounded-full text-[9px] font-mono">Room ${roomsDisplay}</span>
          </div>
          <div class="text-[9px] text-slate-300">
            Check-In: ${formatDateTime(b.checkIn)}<br>
            Check-Out: ${formatDateTime(effectiveCheckout)}<br>
            Status: <span class="${statusColorClass}">${statusText}</span>
          </div>
          <div class="flex justify-between items-center text-[9px] pt-1 border-t border-slate-700/60">
            <span class="text-emerald-400 font-semibold">Total: ₹${b.totalAmount}</span>
            <span class="${b.totalDue > 0 ? 'text-rose-400 font-bold' : 'text-emerald-400 font-bold'}">
              ${b.totalDue > 0 ? `Due: ₹${b.totalDue}` : 'Paid'}
            </span>
          </div>
        `;
        listContainer.appendChild(item);
      });

      const rect = e.target.getBoundingClientRect();
      const scrollY = window.scrollY || window.pageYOffset;
      const scrollX = window.scrollX || window.pageXOffset;

      box.style.top = `${rect.bottom + scrollY + 5}px`;
      let leftPos = rect.left + scrollX - 20;
      if (leftPos + 260 > window.innerWidth) {
        leftPos = window.innerWidth - 270;
      }
      box.style.left = `${Math.max(10, leftPos)}px`;
      box.classList.remove('hidden');
    }

    function closeCommentBox() {
      const box = document.getElementById('excel-comment-box');
      if (box) box.classList.add('hidden');
    }

    // --------------------------------------------------
    // NAVIGATION TAB SWITCHING & EVENT LISTENERS
    // --------------------------------------------------
    function switchTab(tabId) {
      if (tabId === 'master' && !isMasterUnlocked) {
        openMasterAuthModal();
        return;
      }
      performSwitchTab(tabId);
    }

    function performSwitchTab(tabId) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active-tab', 'bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
        btn.classList.add('text-slate-600', 'hover:text-slate-900');
      });

      document.getElementById(`tab-${tabId}`).classList.remove('hidden');
      const activeBtn = document.getElementById(`btn-${tabId}`);
      activeBtn.classList.add('active-tab', 'bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
      closeCommentBox();

      if (tabId === 'dashboard') {
        handleDashboardYearChange(defaultAppYear);
      }
    }

    function refreshAllUI() {
      if (!state.roomsCapacity || state.roomsCapacity.length === 0) {
        state.roomsCapacity = [
          { roomNo: 1, capacity: 4 },
          { roomNo: 2, capacity: 2 },
          { roomNo: 3, capacity: 4 },
          { roomNo: 4, capacity: 4 },
          { roomNo: 5, capacity: 4 }
        ];
      }
      if (!state.masterAgents || state.masterAgents.length === 0) {
        state.masterAgents = [{ agentName: "Self", phone: "Direct", roomNo: "All Rooms" }];
      }
      
      state.selectedYear = defaultAppYear;
      state.dashSelectedYear = defaultAppYear;
      
      if (!state.yearlyCounters || Object.keys(state.yearlyCounters).length === 0) {
        state.yearlyCounters = { [defaultAppYear]: state.bookings.length || 0 };
      }

      populateRoomDropdown();
      populateAgentDropdown();
      searchMasterBookingById();
      renderBookingsTable();
      renderRoomCapacityTable();
      renderMasterAgentTable();
      renderCalendar(defaultAppYear);
      updateDashboardCards();
    }

    async function loadSavedData() {
      const toast = document.getElementById('toast');
      const msg = document.getElementById('toast-message');
      msg.innerText = 'Syncing database...';
      toast.classList.remove('hidden');

      try {
        const response = await fetch(GAS_API_URL + "?action=fetchData");
        const textData = await response.text();
        
        let sheetData;
        try { sheetData = JSON.parse(textData); } catch(err) {
          throw new Error("Invalid response format from server.");
        }
        
        if (sheetData && sheetData.bookings) {
          state = sheetData;
          refreshAllUI(); 
          msg.innerText = 'Database synced successfully!';
        }
      } catch (error) {
        console.error("Error loading data:", error);
        msg.innerText = 'Failed to connect to Database.';
      }
      checkSheetRowLimits();
      setTimeout(() => toast.classList.add('hidden'), 2000);
    }

    async function saveChanges(isAutoSave = false, quiet = false) {
      if (!checkSheetRowLimits()) return;
      if (!quiet) {
        const toast = document.getElementById('toast');
        const msg = document.getElementById('toast-message');
        msg.innerText = isAutoSave ? 'Auto-saving to cloud...' : 'Saving to cloud storage...';
        toast.classList.remove('hidden');
      }

      try {
        const payload = { action: "saveData", state: state };
        const response = await fetch(GAS_API_URL, {
          method: "POST",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify(payload)
        });

        const textResult = await response.text();
        let result;
        try { result = JSON.parse(textResult); } catch(e) {
          throw new Error("Invalid response format.");
        }

        if (result.status === "success") {
          if (!quiet) {
            const msg = document.getElementById('toast-message');
            msg.innerText = isAutoSave ? 'Changes Auto saved successfully!' : 'Data synced with Cloud Storage!';
            setTimeout(() => document.getElementById('toast').classList.add('hidden'), 3000);
          }
        } else {
          throw new Error(result.message || "Server Error");
        }
      } catch (error) {
        console.error("Error saving:", error);
        if (!quiet) {
          alert("Saving Error: " + error.message);
          document.getElementById('toast').classList.add('hidden');
        }
      }
    }

    // --------------------------------------------------
    // APPLICATION INITIALIZATION ON DOM LOADED
    // --------------------------------------------------
    document.addEventListener("DOMContentLoaded", () => {
      checkAuthStatus();
      loadSavedData();
      populateDashboardYearDropdown();
      initDashboard();
      populateCalendarYearDropdown();

      document.addEventListener('click', function(e) {
        const container = document.getElementById('room-dropdown-container');
        if (container && !container.contains(e.target)) {
          const boxes = document.getElementById('room-checkboxes');
          if (boxes) boxes.classList.add('hidden');
        }
      });

      setInterval(checkUpcomingCheckoutsWithDue, 60000);
      setInterval(() => saveChanges(true, true), 300000);
    });
  </script>
</body>
</html>

