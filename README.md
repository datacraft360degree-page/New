<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Business Portal - Web Application</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- SheetJS for Exporting to Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- html2canvas for Generating JPEG Receipts -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js"></script>
  <!-- FontAwesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" />
  <style>
    /* Samsung One UI Smooth Styling & Compact Scrollbar */
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      background-color: #F2F4F7;
    }
    
    ::-webkit-scrollbar {
      width: 5px;
      height: 5px;
    }
    ::-webkit-scrollbar-track {
      background: #F2F4F7;
    }
    ::-webkit-scrollbar-thumb {
      background: #CBD5E1;
      border-radius: 10px;
    }
    ::-webkit-scrollbar-thumb:hover {
      background: #94A3B8;
    }

    /* Print-specific Styles */
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

    /* One UI Comment Box Arrow */
    .excel-comment-box::before {
      content: '';
      position: absolute;
      top: -8px;
      left: 16px;
      border-width: 0 8px 8px 8px;
      border-style: solid;
      border-color: transparent transparent #1E293B transparent;
    }
  </style>
</head>
<body class="text-slate-800 font-sans min-h-screen flex flex-col relative antialiased text-xs" onclick="closeCommentBox()">

  <!-- LOGIN MODAL OVERLAY -->
  <div id="login-overlay" class="fixed inset-0 z-[60] bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-6 space-y-4 text-left">
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
            <input type="text" id="login-userid" required placeholder="Enter User ID" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition" />
          </div>
        </div>

        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">Password</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-key"></i>
            </span>
            <input type="password" id="login-password" required placeholder="Enter Password" class="w-full bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition" />
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

  <!-- LOGIN ALERT MESSAGE MODAL -->
  <div id="login-alert-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-blue-100 max-w-md w-full p-6 space-y-4 text-left">
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
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>1. Google Chrome / Microsoft Edge is the best view browser for this Portal.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>2. Take data backups regularly every day or every week.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>3. Do not force close 'The Portal'; always exit using the 'Logout' option.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>4. Do not login with multiple devices, browsers, or tabs concurrently to prevent data corruption.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>5. Data synchronization takes a brief moment; please allow operations to finish. ⏳</span></p>
      </div>
      <button onclick="closeLoginAlertModal()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-2.5 rounded-2xl shadow-sm transition text-xs mt-2">
        I Understand, Continue
      </button>
    </div>
  </div>

  <!-- MASTER DATA ACCESS PASSWORD MODAL -->
  <div id="master-auth-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-left">
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
            <input type="password" id="master-password-input" required placeholder="Enter Master Password" class="w-full bg-slate-100 border border-transparent focus:border-rose-500 rounded-2xl pl-9 pr-3 py-2 focus:outline-none focus:bg-white text-xs transition" />
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
    <div class="bg-white rounded-3xl shadow-2xl border border-rose-100 max-w-sm w-full p-5 space-y-3 text-center">
      <div class="bg-rose-50 text-rose-600 w-12 h-12 rounded-2xl flex items-center justify-center mx-auto text-xl shadow-sm">
        <i class="fa-solid fa-triangle-exclamation"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-900">Confirm Permanent Deletion</h3>
        <p id="master-delete-modal-msg" class="text-[11px] text-slate-600 mt-1">Are you sure you want to permanently delete this record? This action cannot be undone.</p>
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
    <div class="bg-white rounded-3xl shadow-xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-center">
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

  <!-- SAVING LOCK MODAL -->
  <div id="saving-lock-modal" class="hidden fixed inset-0 z-[100] bg-slate-900/60 backdrop-blur-md flex items-center justify-center p-4 no-print cursor-wait">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-6 text-center space-y-4">
      <div class="text-blue-600 text-5xl animate-bounce">
        ⏳
      </div>
      <div>
        <h3 class="text-lg font-black text-slate-900">Saving Data...</h3>
        <p class="text-xs text-rose-600 mt-2 font-bold uppercase">Do not close window or shutdown!</p>
        <p class="text-[10px] text-slate-500 mt-1">Please wait while we secure your records.</p>
      </div>
    </div>
  </div>

  <!-- SESSION AUTO LOGOUT WARNING MODAL -->
  <div id="logout-warning-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-xl border border-slate-100 max-w-xs w-full p-5 space-y-3 text-center">
      <div class="bg-amber-50 text-amber-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-lg">
        <i class="fa-solid fa-hourglass-half"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-900">Inactivity Timeout Warning</h3>
        <p class="text-[10px] text-slate-500 mt-1">You will be logged out automatically in <strong id="logout-countdown-seconds" class="text-rose-600">60</strong> seconds due to inactivity.</p>
      </div>
      <button onclick="resetInactivityTimer()" class="w-full bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 rounded-xl text-[11px] transition shadow-sm">
        Stay Logged In
      </button>
    </div>
  </div>

  <!-- EXPORT TO EXCEL DATE RANGE MODAL -->
  <div id="export-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-5 space-y-4 text-left">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
          <i class="fa-solid fa-file-excel text-emerald-600"></i> Export to Excel
        </h3>
        <button onclick="closeExportModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>
      <div class="space-y-3 text-[11px]">
        <p class="text-slate-500">Select a specific period to download booking details.</p>
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
      
      <!-- One UI Navigation Tabs -->
      <nav class="flex space-x-1 bg-slate-100 p-1 rounded-full text-[11px] font-medium">
        <button onclick="switchTab('dashboard')" id="btn-dashboard" class="tab-btn px-3 py-1 rounded-full transition-all active-tab bg-white text-blue-600 shadow-sm font-bold">Dashboard</button>
        <button onclick="switchTab('booking')" id="btn-booking" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900">Booking Details</button>
        <button onclick="switchTab('master')" id="btn-master" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900 flex items-center gap-1">
          <i class="fa-solid fa-lock text-[9px] text-amber-500"></i> Master Data
        </button>
        <button onclick="switchTab('calendar')" id="btn-calendar" class="tab-btn px-3 py-1 rounded-full transition-all text-slate-600 hover:text-slate-900">Calendar</button>
      </nav>

      <!-- Header Quick Action Buttons -->
      <div class="flex items-center space-x-1.5">
        <button onclick="openAlertModal()" title="View Alerts" class="relative bg-amber-50 hover:bg-amber-100 text-amber-700 border border-amber-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-bell text-[10px]"></i> Alerts
          <span id="alert-badge" class="hidden absolute -top-1 -right-1 bg-rose-600 text-white text-[9px] font-black px-1.5 py-0.2 rounded-full border border-white animate-bounce">0</span>
        </button>
        <button onclick="saveChanges()" class="bg-emerald-50 hover:bg-emerald-100 text-emerald-700 border border-emerald-200 px-3 py-1.5 rounded-full text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-brands fa-google text-[10px]"></i> Save
        </button>
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
    <span id="toast-message" class="font-medium">Changes saved successfully!</span>
  </div>

  <!-- Main Content Area -->
  <main class="max-w-7xl mx-auto px-4 py-4 flex-1 w-full no-print space-y-4">

    <!-- 1. DASHBOARD TAB -->
    <section id="tab-dashboard" class="tab-content space-y-4">
      <!-- Welcome Banner -->
      <div class="bg-gradient-to-r from-blue-600 to-indigo-600 rounded-3xl p-5 text-white shadow-sm flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
        <div>
          <h2 class="text-base font-bold tracking-tight">Hi Aniruddha, Welcome to Dashboard 🏠</h2>
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

      <!-- Excel Slicer Style Filter Bar -->
      <div class="bg-white p-3 rounded-3xl border border-slate-200/60 shadow-sm space-y-2">
        <div class="flex flex-wrap items-center justify-between gap-2 border-b border-slate-100 pb-2">
          <span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
            <i class="fa-solid fa-sliders text-blue-600"></i> Dynamic Slicer Filter
          </span>
          <div id="dash-slicer-indicator" class="text-[10px] font-semibold text-blue-600">
            Slicer: <strong id="dash-slicer-label">All Categories Active</strong>
          </div>
        </div>
        
        <div class="flex flex-wrap items-center gap-1.5" id="slicer-pill-container">
          <button type="button" onclick="handleDashboardSlicerChange('ALL')" id="slicer-btn-ALL" class="slicer-btn bg-blue-600 text-white font-bold px-3 py-1 rounded-xl text-[10px] shadow-sm transition border border-blue-600">
            <i class="fa-solid fa-border-all mr-1"></i> All Bookings
          </button>
          <button type="button" onclick="handleDashboardSlicerChange('LIVE')" id="slicer-btn-LIVE" class="slicer-btn bg-amber-50 hover:bg-amber-100 text-amber-800 font-semibold px-3 py-1 rounded-xl text-[10px] transition border border-amber-200">
            <i class="fa-solid fa-circle-play mr-1 text-amber-500"></i> Live Only
          </button>
          <button type="button" onclick="handleDashboardSlicerChange('UPCOMING')" id="slicer-btn-UPCOMING" class="slicer-btn bg-blue-50 hover:bg-blue-100 text-blue-800 font-semibold px-3 py-1 rounded-xl text-[10px] transition border border-blue-200">
            <i class="fa-solid fa-clock-rotate-left mr-1 text-blue-500"></i> Upcoming Only
          </button>
          <button type="button" onclick="handleDashboardSlicerChange('CLOSED')" id="slicer-btn-CLOSED" class="slicer-btn bg-emerald-50 hover:bg-emerald-100 text-emerald-800 font-semibold px-3 py-1 rounded-xl text-[10px] transition border border-emerald-200">
            <i class="fa-solid fa-circle-check mr-1 text-emerald-500"></i> Closed Only
          </button>
          <button type="button" onclick="handleDashboardSlicerChange('INACTIVE')" id="slicer-btn-INACTIVE" class="slicer-btn bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold px-3 py-1 rounded-xl text-[10px] transition border border-slate-200">
            <i class="fa-solid fa-ban mr-1 text-rose-500"></i> Inactive Only
          </button>
        </div>
      </div>

      <!-- 4 Status Breakdown Boxes -->
      <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <!-- Live Bookings -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-amber-200/80 space-y-2 relative overflow-hidden">
          <div class="flex items-center justify-between border-b border-amber-100 pb-2">
            <div class="flex items-center gap-1.5">
              <span class="relative flex h-2.5 w-2.5">
                <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
                <span class="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500"></span>
              </span>
              <p class="text-[10px] font-bold text-amber-900 uppercase tracking-wider">Live Bookings</p>
            </div>
            <span id="dash-live-count" class="bg-amber-100 text-amber-800 font-black text-xs px-2.5 py-0.5 rounded-full">0</span>
          </div>
          <div class="space-y-1 text-[11px]">
            <div class="flex justify-between items-center text-slate-600">
              <span>Booking Amount:</span>
              <strong id="dash-live-amount" class="text-slate-900 font-bold">₹0</strong>
            </div>
            <div class="flex justify-between items-center text-emerald-600">
              <span>Amount Received:</span>
              <strong id="dash-live-advance" class="font-bold">₹0</strong>
            </div>
            <div class="flex justify-between items-center text-rose-600">
              <span>Total Due:</span>
              <strong id="dash-live-due" class="font-bold">₹0</strong>
            </div>
          </div>
        </div>

        <!-- Upcoming Bookings -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-blue-200/80 space-y-2 relative overflow-hidden">
          <div class="flex items-center justify-between border-b border-blue-100 pb-2">
            <div class="flex items-center gap-1.5">
              <span class="w-2.5 h-2.5 bg-blue-500 rounded-full inline-block"></span>
              <p class="text-[10px] font-bold text-blue-900 uppercase tracking-wider">Upcoming Bookings</p>
            </div>
            <span id="dash-upcoming-count" class="bg-blue-100 text-blue-800 font-black text-xs px-2.5 py-0.5 rounded-full">0</span>
          </div>
          <div class="space-y-1 text-[11px]">
            <div class="flex justify-between items-center text-slate-600">
              <span>Booking Amount:</span>
              <strong id="dash-upcoming-amount" class="text-slate-900 font-bold">₹0</strong>
            </div>
            <div class="flex justify-between items-center text-emerald-600">
              <span>Amount Received:</span>
              <strong id="dash-upcoming-advance" class="font-bold">₹0</strong>
            </div>
            <div class="flex justify-between items-center text-rose-600">
              <span>Total Due:</span>
              <strong id="dash-upcoming-due" class="font-bold">₹0</strong>
            </div>
          </div>
        </div>

        <!-- Closed Bookings -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-emerald-200/80 space-y-2 relative overflow-hidden">
          <div class="flex items-center justify-between border-b border-emerald-100 pb-2">
            <div class="flex items-center gap-1.5">
              <span class="w-2.5 h-2.5 bg-emerald-500 rounded-full inline-block"></span>
              <p class="text-[10px] font-bold text-emerald-900 uppercase tracking-wider">Closed Bookings</p>
            </div>
            <span id="dash-closed-count" class="bg-emerald-100 text-emerald-800 font-black text-xs px-2.5 py-0.5 rounded-full">0</span>
          </div>
          <div class="space-y-1 text-[11px]">
            <div class="flex justify-between items-center text-slate-600">
              <span>Booking Amount:</span>
              <strong id="dash-closed-amount" class="text-slate-900 font-bold">₹0</strong>
            </div>
            <div class="flex justify-between items-center text-emerald-600">
              <span>Amount Received:</span>
              <strong id="dash-closed-advance" class="font-bold">₹0</strong>
            </div>
            <div class="flex justify-between items-center text-rose-600">
              <span>Total Due:</span>
              <strong id="dash-closed-due" class="font-bold">₹0</strong>
            </div>
          </div>
        </div>

        <!-- Inactive Bookings -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-300 space-y-2 relative overflow-hidden">
          <div class="flex items-center justify-between border-b border-slate-200 pb-2">
            <div class="flex items-center gap-1.5">
              <span class="w-2.5 h-2.5 bg-rose-600 rounded-full inline-block"></span>
              <p class="text-[10px] font-bold text-slate-800 uppercase tracking-wider">Inactive Bookings</p>
            </div>
            <span id="dash-inactive-count" class="bg-slate-200 text-slate-800 font-black text-xs px-2.5 py-0.5 rounded-full">0</span>
          </div>
          <div class="space-y-1 text-[11px]">
            <div class="flex justify-between items-center text-slate-600">
              <span>Booking Amount:</span>
              <strong id="dash-inactive-amount" class="text-slate-900 font-bold">₹0</strong>
            </div>
            <div class="flex justify-between items-center text-emerald-600">
              <span>Amount Received:</span>
              <strong id="dash-inactive-advance" class="font-bold">₹0</strong>
            </div>
            <div class="flex justify-between items-center text-rose-600">
              <span>Total Due:</span>
              <strong id="dash-inactive-due" class="font-bold">₹0</strong>
            </div>
          </div>
        </div>
      </div>

      <!-- Summary Banner -->
      <div class="flex items-center justify-between bg-white px-4 py-2 rounded-2xl border border-slate-200/60 shadow-sm">
        <span class="text-[11px] font-semibold text-slate-600 flex items-center gap-2">
          <i class="fa-solid fa-chart-line text-blue-600"></i>
          Showing Summary For: <strong id="dash-filter-label" class="text-blue-600 font-bold">Consolidated (All Years)</strong>
        </span>
        <button onclick="handleDashboardYearChange('CURRENT')" class="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-3 py-1 rounded-full transition border border-slate-200">
          Reset to Current Year
        </button>
      </div>

      <!-- General Statistics Cards -->
      <div class="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Total Bookings</p>
            <p id="dash-total-bookings" class="text-xl font-black text-slate-900 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-blue-50 text-blue-600 rounded-2xl"><i class="fa-solid fa-bookmark text-base"></i></div>
        </div>
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Booking Amount</p>
            <p id="dash-total-amount" class="text-xl font-black text-slate-900 mt-0.5">₹0</p>
          </div>
          <div class="p-3 bg-indigo-50 text-indigo-600 rounded-2xl"><i class="fa-solid fa-receipt text-base"></i></div>
        </div>
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Amount Received</p>
            <p id="dash-advanced" class="text-xl font-black text-emerald-600 mt-0.5">₹0</p>
          </div>
          <div class="p-3 bg-emerald-50 text-emerald-600 rounded-2xl"><i class="fa-solid fa-wallet text-base"></i></div>
        </div>
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Total Due Amount</p>
            <p id="dash-due" class="text-xl font-black text-rose-600 mt-0.5">₹0</p>
          </div>
          <div class="p-3 bg-rose-50 text-rose-600 rounded-2xl"><i class="fa-solid fa-hand-holding-dollar text-base"></i></div>
        </div>
      </div>

      <!-- Active Years Directory Table -->
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4">
        <div class="mb-3 flex justify-between items-center">
          <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
            <i class="fa-solid fa-calendar-days text-blue-600"></i> Active Years Directory (2026 – 2085)
          </h3>
          <span class="text-[10px] text-slate-400 font-medium">Click any year to filter dashboard &amp; view calendar</span>
        </div>
        <div id="years-grid" class="grid grid-cols-6 sm:grid-cols-10 md:grid-cols-12 gap-2"></div>
      </div>
    </section>

    <!-- 2. BOOKING DETAILS TAB -->
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
          
          <div class="flex flex-wrap items-center space-x-2 w-full md:w-auto">
            <!-- Search by Date -->
            <div class="flex items-center bg-slate-100 border border-slate-200 rounded-2xl px-2 py-1 space-x-1.5">
              <label for="booking-date-search" class="text-[10px] font-bold text-slate-500 uppercase flex items-center gap-1 pl-1">
                <i class="fa-solid fa-calendar-day text-blue-600"></i> Date:
              </label>
              <input type="date" id="booking-date-search" onchange="searchBookingByDate()" class="bg-white text-[11px] border border-slate-200 rounded-xl px-2 py-0.5 focus:outline-none focus:ring-2 focus:ring-blue-500 font-bold text-blue-600 cursor-pointer" />
              <button onclick="clearDateSearchBooking()" class="text-slate-400 hover:text-slate-600 px-1 text-[10px]" title="Reset Date Filter">
                <i class="fa-solid fa-rotate-left"></i>
              </button>
            </div>

            <!-- Global Search Box -->
            <div class="relative">
              <input type="text" id="booking-text-search" onkeyup="filterBookingsTable()" placeholder="Search guest, ID, room..." class="bg-slate-100 border border-slate-200 rounded-2xl pl-7 pr-3 py-1.5 text-[11px] focus:outline-none focus:bg-white focus:ring-2 focus:ring-blue-500 w-44" />
              <i class="fa-solid fa-magnifying-glass absolute left-2.5 top-2.5 text-slate-400 text-[10px]"></i>
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

    <!-- 3. MASTER DATA TAB -->
    <section id="tab-master" class="tab-content hidden space-y-4">
      <div class="bg-amber-50/80 border border-amber-200 p-4 rounded-3xl flex items-center justify-between">
        <div class="flex items-center gap-2.5">
          <div class="bg-amber-500 text-white p-2 rounded-2xl text-sm"><i class="fa-solid fa-lock"></i></div>
          <div>
            <h3 class="text-xs font-bold text-amber-900">Protected Master Configuration Tab</h3>
            <p class="text-[10px] text-amber-700">Add or manage room types, agent lists, and portal settings.</p>
          </div>
        </div>
      </div>

      <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
        <!-- Room Management Card -->
        <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4 space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2.5">
            <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-bed text-blue-600"></i> Room Category &amp; Inventory
            </h3>
            <button onclick="openAddRoomModal()" class="bg-blue-50 text-blue-600 hover:bg-blue-100 font-bold px-2.5 py-1 rounded-xl text-[10px] transition">
              + Add Room
            </button>
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-left">
              <thead>
                <tr class="text-[9px] font-bold text-slate-400 uppercase border-b border-slate-100">
                  <th class="py-1.5 px-2">Room No/Name</th>
                  <th class="py-1.5 px-2">Type</th>
                  <th class="py-1.5 px-2">Capacity</th>
                  <th class="py-1.5 px-2">Base Price</th>
                  <th class="py-1.5 px-2 text-right">Action</th>
                </tr>
              </thead>
              <tbody id="rooms-tbody" class="divide-y divide-slate-100"></tbody>
            </table>
          </div>
        </div>

        <!-- Agent Management Card -->
        <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4 space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2.5">
            <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
              <i class="fa-solid fa-user-tie text-blue-600"></i> Booking Agents Directory
            </h3>
            <button onclick="openAddAgentModal()" class="bg-blue-50 text-blue-600 hover:bg-blue-100 font-bold px-2.5 py-1 rounded-xl text-[10px] transition">
              + Add Agent
            </button>
          </div>
          <div class="overflow-x-auto">
            <table class="w-full text-left">
              <thead>
                <tr class="text-[9px] font-bold text-slate-400 uppercase border-b border-slate-100">
                  <th class="py-1.5 px-2">Agent Name</th>
                  <th class="py-1.5 px-2">Company / Channel</th>
                  <th class="py-1.5 px-2">Contact</th>
                  <th class="py-1.5 px-2 text-right">Action</th>
                </tr>
              </thead>
              <tbody id="agents-tbody" class="divide-y divide-slate-100"></tbody>
            </table>
          </div>
        </div>
      </div>
    </section>

    <!-- 4. CALENDAR TAB -->
    <section id="tab-calendar" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4 space-y-4">
        <!-- Calendar Controls -->
        <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 border-b border-slate-100 pb-3">
          <div class="flex items-center space-x-2">
            <button onclick="navigateMonth(-1)" class="p-1.5 rounded-xl hover:bg-slate-100 text-slate-600 text-xs transition">
              <i class="fa-solid fa-chevron-left"></i>
            </button>
            <h2 id="calendar-month-year" class="text-sm font-bold text-slate-900 min-w-[140px] text-center">August 2026</h2>
            <button onclick="navigateMonth(1)" class="p-1.5 rounded-xl hover:bg-slate-100 text-slate-600 text-xs transition">
              <i class="fa-solid fa-chevron-right"></i>
            </button>
          </div>

          <div class="flex items-center space-x-2">
            <label for="calendar-year-select" class="text-[10px] font-bold text-slate-500 uppercase">Jump To Year:</label>
            <select id="calendar-year-select" onchange="jumpCalendarYear(this.value)" class="bg-slate-100 border border-slate-200 rounded-xl px-2.5 py-1 text-xs font-bold text-slate-800 focus:outline-none focus:bg-white">
            </select>
          </div>
        </div>

        <!-- Days Header -->
        <div class="grid grid-cols-7 text-center font-bold text-[10px] text-slate-400 uppercase tracking-wider py-1">
          <span>Sun</span><span>Mon</span><span>Tue</span><span>Wed</span><span>Thu</span><span>Fri</span><span>Sat</span>
        </div>

        <!-- Month Grid -->
        <div id="calendar-days-grid" class="grid grid-cols-7 gap-1.5 text-center"></div>
      </div>
    </section>

  </main>
 <!-- BOOKING ADD / EDIT MODAL -->
  <div id="booking-modal" class="hidden fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-md flex items-center justify-center p-4 no-print overflow-y-auto">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-2xl w-full p-6 space-y-4 my-8 text-left">
      <div class="flex justify-between items-center pb-3 border-b border-slate-100">
        <h3 id="booking-modal-title" class="text-sm font-bold text-slate-900 flex items-center gap-2">
          <i class="fa-solid fa-calendar-plus text-blue-600"></i> New Reservation Record
        </h3>
        <button onclick="closeBookingModal()" class="text-slate-400 hover:text-slate-600 p-1 text-base transition">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>

      <form id="booking-form" onsubmit="handleBookingFormSubmit(event)" class="space-y-4">
        <input type="hidden" id="booking-form-id" />

        <!-- Guest Details Section -->
        <div class="bg-slate-50 p-3.5 rounded-2xl border border-slate-200/60 space-y-3">
          <p class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Guest Information</p>
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Full Name <span class="text-rose-500">*</span></label>
              <input type="text" id="b-guest-name" required placeholder="e.g. Rahul Sharma" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Contact Number <span class="text-rose-500">*</span></label>
              <input type="tel" id="b-contact" required placeholder="e.g. +91 9876543210" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Government ID Number</label>
              <input type="text" id="b-id-number" placeholder="Aadhaar / Passport / DL" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Attach ID Proof (Image / PDF)</label>
              <input type="file" id="b-id-file" accept="image/*,application/pdf" onchange="handleIDFileUpload(event)" class="w-full text-[10px] text-slate-500 file:mr-2 file:py-1 file:px-2.5 file:rounded-xl file:border-0 file:text-[10px] file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100" />
              <input type="hidden" id="b-id-file-data" />
            </div>
          </div>
        </div>

        <!-- Room & Stay Details Section -->
        <div class="bg-slate-50 p-3.5 rounded-2xl border border-slate-200/60 space-y-3">
          <p class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Stay &amp; Inventory Allocation</p>
          <div class="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Select Room <span class="text-rose-500">*</span></label>
              <select id="b-room-select" required onchange="autoFillRoomRate()" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium">
                <option value="">-- Choose Room --</option>
              </select>
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Guest Capacity</label>
              <input type="number" id="b-capacity" min="1" max="20" placeholder="Number of guests" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Booking Agent</label>
              <select id="b-agent-select" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium">
                <option value="Direct">Direct / Walk-In</option>
              </select>
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Check-In Date <span class="text-rose-500">*</span></label>
              <input type="date" id="b-checkin" required min="2026-08-01" max="2085-12-31" onchange="calculateBookingTotals()" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Check-Out Date <span class="text-rose-500">*</span></label>
              <input type="date" id="b-checkout" required min="2026-08-01" max="2085-12-31" onchange="calculateBookingTotals()" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Status Classification</label>
              <select id="b-status" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-blue-600">
                <option value="LIVE">Live</option>
                <option value="UPCOMING">Upcoming</option>
                <option value="CLOSED">Closed</option>
                <option value="INACTIVE">Inactive / Cancelled</option>
              </select>
            </div>
          </div>
        </div>

        <!-- Billing & Financial Calculation Section -->
        <div class="bg-slate-50 p-3.5 rounded-2xl border border-slate-200/60 space-y-3">
          <p class="text-[10px] font-bold text-slate-400 uppercase tracking-wider">Financial Breakdown</p>
          <div class="grid grid-cols-1 sm:grid-cols-4 gap-3">
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Room Tariff (₹/Night)</label>
              <input type="number" id="b-tariff" min="0" value="0" oninput="calculateBookingTotals()" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Extra Charges (₹)</label>
              <input type="number" id="b-extras" min="0" value="0" oninput="calculateBookingTotals()" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Advance Paid (₹)</label>
              <input type="number" id="b-advance" min="0" value="0" oninput="calculateBookingTotals()" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
            </div>
            <div>
              <label class="block font-semibold text-slate-700 mb-1">Balance Due (₹)</label>
              <input type="number" id="b-due" readonly value="0" class="w-full bg-slate-100 border border-slate-200 text-rose-600 font-bold rounded-xl px-3 py-1.5 cursor-not-allowed" />
            </div>
          </div>
        </div>

        <!-- Remarks & Notes -->
        <div>
          <label class="block font-semibold text-slate-700 mb-1">Special Instructions / Notes</label>
          <textarea id="b-notes" rows="2" placeholder="e.g. Early check-in requested, extra bed required" class="w-full bg-white border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium text-[11px]"></textarea>
        </div>

        <div class="flex space-x-2 pt-2 border-t border-slate-100">
          <button type="button" onclick="closeBookingModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">
            Cancel
          </button>
          <button type="submit" class="w-1/2 bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px] flex items-center justify-center gap-1.5">
            <i class="fa-solid fa-floppy-disk"></i> Save Reservation
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- MASTER ROOM ADD / EDIT MODAL -->
  <div id="room-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-5 space-y-3 text-left">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <h3 id="room-modal-title" class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
          <i class="fa-solid fa-bed text-blue-600"></i> Configure Room Category
        </h3>
        <button onclick="closeRoomModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <form id="room-form" onsubmit="handleRoomFormSubmit(event)" class="space-y-3">
        <input type="hidden" id="room-form-id" />
        <div>
          <label class="block font-semibold text-slate-700 mb-0.5">Room Number / Title <span class="text-rose-500">*</span></label>
          <input type="text" id="r-name" required placeholder="e.g. Room 101 or Suite A" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
        </div>
        <div>
          <label class="block font-semibold text-slate-700 mb-0.5">Room Type</label>
          <input type="text" id="r-type" placeholder="e.g. Deluxe Double / Executive Suite" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
        </div>
        <div class="grid grid-cols-2 gap-2">
          <div>
            <label class="block font-semibold text-slate-700 mb-0.5">Max Capacity</label>
            <input type="number" id="r-capacity" min="1" value="2" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
          </div>
          <div>
            <label class="block font-semibold text-slate-700 mb-0.5">Base Rate (₹)</label>
            <input type="number" id="r-price" min="0" value="1500" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
          </div>
        </div>
        <div class="flex space-x-2 pt-2 border-t border-slate-100">
          <button type="button" onclick="closeRoomModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">Cancel</button>
          <button type="submit" class="w-1/2 bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px]">Save Room</button>
        </div>
      </form>
    </div>
  </div>

  <!-- MASTER AGENT ADD / EDIT MODAL -->
  <div id="agent-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-sm w-full p-5 space-y-3 text-left">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <h3 id="agent-modal-title" class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
          <i class="fa-solid fa-user-tie text-blue-600"></i> Configure Booking Agent
        </h3>
        <button onclick="closeAgentModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <form id="agent-form" onsubmit="handleAgentFormSubmit(event)" class="space-y-3">
        <input type="hidden" id="agent-form-id" />
        <div>
          <label class="block font-semibold text-slate-700 mb-0.5">Agent Name <span class="text-rose-500">*</span></label>
          <input type="text" id="a-name" required placeholder="e.g. Ramesh Kumar" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
        </div>
        <div>
          <label class="block font-semibold text-slate-700 mb-0.5">Agency / Channel</label>
          <input type="text" id="a-company" placeholder="e.g. MakeMyTrip / Booking.com / Local" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
        </div>
        <div>
          <label class="block font-semibold text-slate-700 mb-0.5">Contact Number</label>
          <input type="tel" id="a-contact" placeholder="e.g. +91 9876500000" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
        </div>
        <div class="flex space-x-2 pt-2 border-t border-slate-100">
          <button type="button" onclick="closeAgentModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">Cancel</button>
          <button type="submit" class="w-1/2 bg-blue-600 hover:bg-blue-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px]">Save Agent</button>
        </div>
      </form>
    </div>
  </div>

  <!-- ALERTS / NOTIFICATIONS MODAL -->
  <div id="alerts-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-md w-full p-5 space-y-4 text-left">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
          <i class="fa-solid fa-bell text-amber-500"></i> Reservation Alerts &amp; Notifications
        </h3>
        <button onclick="closeAlertModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <div id="alerts-container" class="space-y-2 max-h-80 overflow-y-auto pr-1">
        <!-- Dynamic Alerts List Loaded via JS -->
      </div>

      <button onclick="closeAlertModal()" class="w-full bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">
        Close
      </button>
    </div>
  </div>

  <!-- WIPE DATA RECONFIRMATION PASSWORD MODAL -->
  <div id="wipe-data-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-rose-200 max-w-sm w-full p-6 space-y-4 text-center">
      <div class="bg-rose-100 text-rose-600 w-12 h-12 rounded-2xl flex items-center justify-center mx-auto text-xl shadow-sm">
        <i class="fa-solid fa-triangle-exclamation"></i>
      </div>
      <div>
        <h3 class="text-sm font-bold text-slate-900">Emergency Factory Data Wipe</h3>
        <p class="text-[11px] text-slate-500 mt-1">This will permanently erase all local booking data, master records, and reset portal configurations.</p>
      </div>

      <form onsubmit="handleDataWipeSubmit(event)" class="space-y-3 text-left">
        <div>
          <label class="block text-[10px] font-semibold text-slate-700 mb-1">Enter Security Wipe Password</label>
          <input type="password" id="wipe-password-input" required placeholder="Enter wipe password" class="w-full bg-slate-50 border border-slate-200 rounded-xl px-3 py-2 text-xs focus:outline-none focus:border-rose-500 font-bold" />
        </div>
        <div id="wipe-error-msg" class="hidden text-[10px] text-rose-600 font-semibold bg-rose-50 p-2 rounded-xl text-center">
          Incorrect security password!
        </div>
        <div class="flex space-x-2 pt-1">
          <button type="button" onclick="closeDataWipeModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-2 rounded-xl text-[11px] transition">
            Cancel
          </button>
          <button type="submit" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-2 rounded-xl shadow-sm transition text-[11px]">
            Wipe Everything
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- ATTACHED ID PREVIEW MODAL -->
  <div id="id-preview-modal" class="hidden fixed inset-0 z-[70] bg-slate-900/70 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-lg w-full p-5 space-y-3 text-center">
      <div class="flex justify-between items-center border-b border-slate-100 pb-2">
        <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
          <i class="fa-solid fa-id-card text-blue-600"></i> Attached Identification Document
        </h3>
        <button onclick="closeIDPreviewModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>
      <div id="id-preview-content" class="max-h-[70vh] overflow-auto flex items-center justify-center bg-slate-100 rounded-2xl p-2">
        <!-- Dynamic Image or PDF Link -->
      </div>
      <div class="flex justify-end pt-2">
        <button onclick="closeIDPreviewModal()" class="bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold px-4 py-1.5 rounded-xl text-[11px] transition">
          Close Preview
        </button>
      </div>
    </div>
  </div>

  <!-- PRINTABLE INVOICE / RECEIPT CONTAINER (VISIBLE ONLY DURING PRINT OR JPEG CAPTURE) -->
  <div id="printable-invoice" class="hidden bg-white p-8 max-w-2xl mx-auto rounded-3xl border border-slate-200 space-y-6 text-slate-800">
    <div class="flex justify-between items-start border-b border-slate-200 pb-4">
      <div>
        <h1 class="text-xl font-black text-blue-600 uppercase tracking-tight">Booking Receipt</h1>
        <p class="text-[11px] text-slate-500">Official Guest Stay Invoice</p>
      </div>
      <div class="text-right text-[10px] text-slate-500 space-y-0.5">
        <p><strong class="text-slate-900">Receipt No:</strong> <span id="inv-id">#0000</span></p>
        <p><strong class="text-slate-900">Date:</strong> <span id="inv-date">--</span></p>
      </div>
    </div>

    <!-- Guest & Stay Specs -->
    <div class="grid grid-cols-2 gap-4 text-[11px] bg-slate-50 p-4 rounded-2xl border border-slate-100">
      <div class="space-y-1">
        <p class="text-[9px] font-bold text-slate-400 uppercase">Guest Information</p>
        <p class="font-bold text-slate-900 text-xs" id="inv-guest-name">--</p>
        <p class="text-slate-600" id="inv-guest-contact">--</p>
        <p class="text-slate-600" id="inv-guest-id">--</p>
      </div>
      <div class="space-y-1">
        <p class="text-[9px] font-bold text-slate-400 uppercase">Stay Details</p>
        <p class="text-slate-700"><strong>Room:</strong> <span id="inv-room">--</span></p>
        <p class="text-slate-700"><strong>Check-In:</strong> <span id="inv-checkin">--</span></p>
        <p class="text-slate-700"><strong>Check-Out:</strong> <span id="inv-checkout">--</span></p>
      </div>
    </div>

    <!-- Financial Table -->
    <table class="w-full text-left text-[11px] border-collapse">
      <thead>
        <tr class="border-b border-slate-200 text-[10px] font-bold text-slate-400 uppercase">
          <th class="py-2">Description</th>
          <th class="py-2 text-right">Amount (₹)</th>
        </tr>
      </thead>
      <tbody class="divide-y divide-slate-100">
        <tr>
          <td class="py-2 text-slate-700">Room Tariff Charges</td>
          <td class="py-2 text-right font-medium" id="inv-tariff-total">₹0</td>
        </tr>
        <tr>
          <td class="py-2 text-slate-700">Extra Services &amp; Amenities</td>
          <td class="py-2 text-right font-medium" id="inv-extras">₹0</td>
        </tr>
        <tr class="font-bold text-slate-900 bg-slate-50">
          <td class="py-2.5 px-2">Total Amount Payable</td>
          <td class="py-2.5 px-2 text-right text-xs" id="inv-grand-total">₹0</td>
        </tr>
        <tr class="text-emerald-700 font-semibold">
          <td class="py-2">Advance Amount Paid</td>
          <td class="py-2 text-right" id="inv-advance">₹0</td>
        </tr>
        <tr class="text-rose-600 font-bold">
          <td class="py-2">Net Due Balance</td>
          <td class="py-2 text-right text-xs" id="inv-due">₹0</td>
        </tr>
      </tbody>
    </table>

    <div class="border-t border-slate-200 pt-4 flex justify-between items-center text-[10px] text-slate-400">
      <p>Thank you for staying with us!</p>
      <p>Computer Generated Receipt</p>
    </div>
  </div> 
  <script>
  /* ==========================================================================
     GLOBAL STATE & SYSTEM VARIABLES
     ========================================================================== */
  const GOOGLE_SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE';
  
  // Credentials & Security Keys
  const AUTH_USER_ID = "admin";
  const AUTH_PASSWORD = "123";
  const MASTER_PASSWORD = "123";
  const DATA_WIPE_PASSWORD = "123";

  // System Core Data Arrays
  let bookings = [];
  let rooms = [];
  let agents = [];

  // Inactivity & Session Timer State Variables
  let inactivityTimer = null;
  let warningTimer = null;
  let countdownInterval = null;
  const INACTIVITY_TIMEOUT_MS = 14 * 60 * 1000; // 14 Minutes
  const WARNING_DURATION_SEC = 60; // 1 Minute Warning

  // Active View & Filter State
  let currentTab = 'dashboard';
  let activeDashboardYear = 'CONSOLIDATED';
  let activeDashboardSlicer = 'ALL';
  let pendingMasterDeleteCallback = null;

  /* ==========================================================================
     APPLICATION INITIALIZATION & LOCAL STORAGE ENGINE
     ========================================================================== */
  window.addEventListener('DOMContentLoaded', () => {
    populateYearDropdowns();
    loadLocalDatabase();
    checkAuthStatus();
  });

  function loadLocalDatabase() {
    const savedBookings = localStorage.getItem('portal_bookings');
    const savedRooms = localStorage.getItem('portal_rooms');
    const savedAgents = localStorage.getItem('portal_agents');

    bookings = savedBookings ? JSON.parse(savedBookings) : getInitialDefaultBookings();
    rooms = savedRooms ? JSON.parse(savedRooms) : getInitialDefaultRooms();
    agents = savedAgents ? JSON.parse(savedAgents) : getInitialDefaultAgents();

    syncDatabaseToLocalStorage();
  }

  function syncDatabaseToLocalStorage() {
    localStorage.setItem('portal_bookings', JSON.stringify(bookings));
    localStorage.setItem('portal_rooms', JSON.stringify(rooms));
    localStorage.setItem('portal_agents', JSON.stringify(agents));
  }

  function getInitialDefaultRooms() {
    return [
      { id: 'RM-101', name: 'Room 101', type: 'Deluxe Double', capacity: 2, price: 1800 },
      { id: 'RM-102', name: 'Room 102', type: 'Deluxe Double', capacity: 2, price: 1800 },
      { id: 'RM-201', name: 'Suite 201', type: 'Executive Suite', capacity: 4, price: 3500 },
      { id: 'RM-202', name: 'Suite 202', type: 'Executive Suite', capacity: 4, price: 3500 }
    ];
  }

  function getInitialDefaultAgents() {
    return [
      { id: 'AG-01', name: 'Direct / Walk-In', company: 'Internal', contact: 'N/A' },
      { id: 'AG-02', name: 'Ramesh Sharma', company: 'MakeMyTrip', contact: '+91 9876543210' },
      { id: 'AG-03', name: 'Anita Roy', company: 'Booking.com', contact: '+91 9812345678' }
    ];
  }

  function getInitialDefaultBookings() {
    return [
      {
        id: 'BK-2026-001',
        guestName: 'Aniruddha Sen',
        contact: '+91 9800011122',
        idNumber: 'IND-1234-5678',
        idFileData: '',
        roomId: 'RM-201',
        capacity: 3,
        agent: 'Direct / Walk-In',
        checkIn: '2026-08-15',
        checkOut: '2026-08-20',
        tariff: 3500,
        extras: 500,
        advance: 5000,
        due: 13000,
        status: 'LIVE',
        notes: 'VVIP Guest. Prefers high floor.'
      }
    ];
  }

  /* ==========================================================================
     AUTHENTICATION & ACCESS CONTROL
     ========================================================================== */
  function checkAuthStatus() {
    const isAuthenticated = sessionStorage.getItem('portal_authenticated');
    const loginOverlay = document.getElementById('login-overlay');
    if (isAuthenticated === 'true') {
      loginOverlay.classList.add('hidden');
      startInactivityTimer();
      refreshActiveTabUI();
    } else {
      loginOverlay.classList.remove('hidden');
    }
  }

  function handleLogin(e) {
    e.preventDefault();
    const uid = document.getElementById('login-userid').value.trim();
    const pwd = document.getElementById('login-password').value.trim();
    const errEl = document.getElementById('login-error');

    if (uid === AUTH_USER_ID && pwd === AUTH_PASSWORD) {
      errEl.classList.add('hidden');
      sessionStorage.setItem('portal_authenticated', 'true');
      document.getElementById('login-overlay').classList.add('hidden');
      document.getElementById('login-alert-modal').classList.remove('hidden');
      startInactivityTimer();
      refreshActiveTabUI();
    } else {
      errEl.classList.remove('hidden');
    }
  }

  function closeLoginAlertModal() {
    document.getElementById('login-alert-modal').classList.add('hidden');
  }

  function logoutUser() {
    document.getElementById('logout-confirm-modal').classList.remove('hidden');
  }

  function cancelLogout() {
    document.getElementById('logout-confirm-modal').classList.add('hidden');
  }

  function processLogoutWithSave() {
    showSavingLock();
    syncDatabaseToLocalStorage();
    setTimeout(() => {
      sessionStorage.removeItem('portal_authenticated');
      sessionStorage.removeItem('master_authenticated');
      hideSavingLock();
      window.location.reload();
    }, 1200);
  }

  /* ==========================================================================
     PROTECTED MASTER DATA AUTHENTICATION
     ========================================================================== */
  function handleMasterAuth(e) {
    e.preventDefault();
    const pwd = document.getElementById('master-password-input').value.trim();
    const errEl = document.getElementById('master-auth-error');

    if (pwd === MASTER_PASSWORD) {
      errEl.classList.add('hidden');
      sessionStorage.setItem('master_authenticated', 'true');
      document.getElementById('master-auth-modal').classList.add('hidden');
      document.getElementById('master-password-input').value = '';
      switchTab('master');
    } else {
      errEl.classList.remove('hidden');
    }
  }

  function closeMasterAuthModal() {
    document.getElementById('master-auth-modal').classList.add('hidden');
  }

  function requestMasterDataDeletion(callback, message) {
    const isMasterAuth = sessionStorage.getItem('master_authenticated') === 'true';
    if (!isMasterAuth) {
      alert('Unauthorized! Master authentication required.');
      return;
    }
    pendingMasterDeleteCallback = callback;
    document.getElementById('master-delete-modal-msg').innerText = message || 'Are you sure you want to delete this master record?';
    document.getElementById('master-delete-confirm-modal').classList.remove('hidden');
  }

  function closeMasterDeleteModal() {
    document.getElementById('master-delete-confirm-modal').classList.add('hidden');
    pendingMasterDeleteCallback = null;
  }

  function confirmMasterDeletion() {
    if (typeof pendingMasterDeleteCallback === 'function') {
      pendingMasterDeleteCallback();
    }
    closeMasterDeleteModal();
  }

  /* ==========================================================================
     INACTIVITY & SESSION SECURITY TIMERS
     ========================================================================== */
  function startInactivityTimer() {
    clearInactivityTimers();
    document.addEventListener('mousemove', resetInactivityTimer);
    document.addEventListener('keydown', resetInactivityTimer);
    document.addEventListener('click', resetInactivityTimer);
    
    inactivityTimer = setTimeout(triggerInactivityWarning, INACTIVITY_TIMEOUT_MS);
  }

  function resetInactivityTimer() {
    const warningModal = document.getElementById('logout-warning-modal');
    if (!warningModal.classList.contains('hidden')) {
      warningModal.classList.add('hidden');
    }
    clearTimeout(inactivityTimer);
    clearInterval(countdownInterval);
    inactivityTimer = setTimeout(triggerInactivityWarning, INACTIVITY_TIMEOUT_MS);
  }

  function clearInactivityTimers() {
    if (inactivityTimer) clearTimeout(inactivityTimer);
    if (warningTimer) clearTimeout(warningTimer);
    if (countdownInterval) clearInterval(countdownInterval);
  }

  function triggerInactivityWarning() {
    let secondsLeft = WARNING_DURATION_SEC;
    const countdownEl = document.getElementById('logout-countdown-seconds');
    countdownEl.innerText = secondsLeft;
    document.getElementById('logout-warning-modal').classList.remove('hidden');

    countdownInterval = setInterval(() => {
      secondsLeft--;
      countdownEl.innerText = secondsLeft;
      if (secondsLeft <= 0) {
        clearInterval(countdownInterval);
        processLogoutWithSave();
      }
    }, 1000);
  }

  /* ==========================================================================
     UI NAVIGATION TABS & YEAR DIRECTORY
     ========================================================================== */
  function switchTab(tabName) {
    if (tabName === 'master' && sessionStorage.getItem('master_authenticated') !== 'true') {
      document.getElementById('master-auth-modal').classList.remove('hidden');
      return;
    }

    currentTab = tabName;
    document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('.tab-btn').forEach(btn => {
      btn.classList.remove('active-tab', 'bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
      btn.classList.add('text-slate-600');
    });

    const activeBtn = document.getElementById(`btn-${tabName}`);
    const activeSection = document.getElementById(`tab-${tabName}`);

    if (activeBtn && activeSection) {
      activeBtn.classList.add('active-tab', 'bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
      activeBtn.classList.remove('text-slate-600');
      activeSection.classList.remove('hidden');
    }

    refreshActiveTabUI();
  }

  function refreshActiveTabUI() {
    if (currentTab === 'dashboard') renderDashboard();
    if (currentTab === 'booking') renderBookingsTable();
    if (currentTab === 'master') renderMasterData();
    if (currentTab === 'calendar') renderCalendar();
    updateAlertBadge();
  }

  function populateYearDropdowns() {
    const dashSelect = document.getElementById('dash-year-select');
    const calSelect = document.getElementById('calendar-year-select');
    dashSelect.innerHTML = '<option value="CONSOLIDATED">Consolidated (All Years)</option>';
    calSelect.innerHTML = '';

    for (let yr = 2026; yr <= 2085; yr++) {
      dashSelect.innerHTML += `<option value="${yr}">${yr}</option>`;
      calSelect.innerHTML += `<option value="${yr}">${yr}</option>`;
    }

    const grid = document.getElementById('years-grid');
    grid.innerHTML = '';
    for (let yr = 2026; yr <= 2085; yr++) {
      grid.innerHTML += `
        <button onclick="selectYearFromDirectory(${yr})" class="bg-slate-50 hover:bg-blue-50 border border-slate-200/80 hover:border-blue-300 rounded-2xl py-1.5 text-center font-bold text-[11px] text-slate-700 hover:text-blue-600 transition">
          ${yr}
        </button>
      `;
    }
  }

  function selectYearFromDirectory(year) {
    document.getElementById('dash-year-select').value = year;
    handleDashboardYearChange(year);
  }

  /* ==========================================================================
     FINANCIAL CALCULATIONS & FORM VALIDATION ENGINE
     ========================================================================== */
  function autoFillRoomRate() {
    const roomId = document.getElementById('b-room-select').value;
    const selectedRoom = rooms.find(r => r.id === roomId);
    if (selectedRoom) {
      document.getElementById('b-tariff').value = selectedRoom.price;
      document.getElementById('b-capacity').value = selectedRoom.capacity;
      calculateBookingTotals();
    }
  }

  function calculateBookingTotals() {
    const checkInVal = document.getElementById('b-checkin').value;
    const checkOutVal = document.getElementById('b-checkout').value;
    const tariffVal = parseFloat(document.getElementById('b-tariff').value) || 0;
    const extrasVal = parseFloat(document.getElementById('b-extras').value) || 0;
    const advanceVal = parseFloat(document.getElementById('b-advance').value) || 0;

    let nights = 1;
    if (checkInVal && checkOutVal) {
      const d1 = new Date(checkInVal);
      const d2 = new Date(checkOutVal);
      const diffTime = d2 - d1;
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      nights = diffDays > 0 ? diffDays : 1;
    }

    const totalCost = (tariffVal * nights) + extrasVal;
    const dueAmount = totalCost - advanceVal;

    document.getElementById('b-due').value = dueAmount >= 0 ? dueAmount : 0;
  }

  function handleIDFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (file.size > 3 * 1024 * 1024) {
      alert('File size exceeds 3MB limit!');
      event.target.value = '';
      return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
      document.getElementById('b-id-file-data').value = e.target.result;
    };
    reader.readAsDataURL(file);
  }

  function showSavingLock() {
    document.getElementById('saving-lock-modal').classList.remove('hidden');
  }

  function hideSavingLock() {
    document.getElementById('saving-lock-modal').classList.add('hidden');
  }

  function showToast(msg) {
    const toast = document.getElementById('toast');
    document.getElementById('toast-message').innerText = msg || 'Action completed successfully!';
    toast.classList.remove('hidden');
    setTimeout(() => {
      toast.classList.add('hidden');
    }, 3000);
  }
</script>
<script>
  /* ==========================================================================
     BOOKING CRUD OPERATIONS & FORM HANDLERS
     ========================================================================== */
  function openBookingModal(bookingId = null) {
    const modal = document.getElementById('booking-modal');
    const form = document.getElementById('booking-form');
    const titleEl = document.getElementById('booking-modal-title');
    const roomSelect = document.getElementById('b-room-select');
    const agentSelect = document.getElementById('b-agent-select');

    form.reset();
    document.getElementById('booking-form-id').value = '';
    document.getElementById('b-id-file-data').value = '';

    // Populate Dynamic Room Select
    roomSelect.innerHTML = '<option value="">-- Choose Room --</option>';
    rooms.forEach(r => {
      roomSelect.innerHTML += `<option value="${r.id}">${r.name} (${r.type}) - ₹${r.price}</option>`;
    });

    // Populate Dynamic Agent Select
    agentSelect.innerHTML = '';
    agents.forEach(a => {
      agentSelect.innerHTML += `<option value="${a.name}">${a.name} ${a.company ? '(' + a.company + ')' : ''}</option>`;
    });

    if (bookingId) {
      const b = bookings.find(item => item.id === bookingId);
      if (b) {
        titleEl.innerHTML = `<i class="fa-solid fa-pen-to-square text-blue-600"></i> Edit Reservation Record (${b.id})`;
        document.getElementById('booking-form-id').value = b.id;
        document.getElementById('b-guest-name').value = b.guestName;
        document.getElementById('b-contact').value = b.contact;
        document.getElementById('b-id-number').value = b.idNumber || '';
        document.getElementById('b-id-file-data').value = b.idFileData || '';
        document.getElementById('b-room-select').value = b.roomId;
        document.getElementById('b-capacity').value = b.capacity || 1;
        document.getElementById('b-agent-select').value = b.agent || 'Direct / Walk-In';
        document.getElementById('b-checkin').value = b.checkIn;
        document.getElementById('b-checkout').value = b.checkOut;
        document.getElementById('b-status').value = b.status;
        document.getElementById('b-tariff').value = b.tariff;
        document.getElementById('b-extras').value = b.extras;
        document.getElementById('b-advance').value = b.advance;
        document.getElementById('b-due').value = b.due;
        document.getElementById('b-notes').value = b.notes || '';
      }
    } else {
      titleEl.innerHTML = `<i class="fa-solid fa-calendar-plus text-blue-600"></i> New Reservation Record`;
      const today = new Date().toISOString().split('T')[0];
      document.getElementById('b-checkin').value = today;
      document.getElementById('b-checkout').value = today;
    }

    modal.classList.remove('hidden');
  }

  function closeBookingModal() {
    document.getElementById('booking-modal').classList.add('hidden');
  }

  function handleBookingFormSubmit(e) {
    e.preventDefault();
    showSavingLock();

    const id = document.getElementById('booking-form-id').value;
    const bookingData = {
      id: id || `BK-${new Date().getFullYear()}-${String(bookings.length + 1).padStart(3, '0')}`,
      guestName: document.getElementById('b-guest-name').value.trim(),
      contact: document.getElementById('b-contact').value.trim(),
      idNumber: document.getElementById('b-id-number').value.trim(),
      idFileData: document.getElementById('b-id-file-data').value,
      roomId: document.getElementById('b-room-select').value,
      capacity: parseInt(document.getElementById('b-capacity').value) || 1,
      agent: document.getElementById('b-agent-select').value,
      checkIn: document.getElementById('b-checkin').value,
      checkOut: document.getElementById('b-checkout').value,
      status: document.getElementById('b-status').value,
      tariff: parseFloat(document.getElementById('b-tariff').value) || 0,
      extras: parseFloat(document.getElementById('b-extras').value) || 0,
      advance: parseFloat(document.getElementById('b-advance').value) || 0,
      due: parseFloat(document.getElementById('b-due').value) || 0,
      notes: document.getElementById('b-notes').value.trim()
    };

    if (id) {
      const idx = bookings.findIndex(b => b.id === id);
      if (idx !== -1) bookings[idx] = bookingData;
    } else {
      bookings.push(bookingData);
    }

    syncDatabaseToLocalStorage();

    setTimeout(() => {
      hideSavingLock();
      closeBookingModal();
      refreshActiveTabUI();
      showToast(id ? 'Reservation updated successfully!' : 'New booking created successfully!');
    }, 600);
  }

  function deleteBooking(bookingId) {
    if (confirm(`Are you sure you want to delete reservation ${bookingId}?`)) {
      showSavingLock();
      bookings = bookings.filter(b => b.id !== bookingId);
      syncDatabaseToLocalStorage();
      setTimeout(() => {
        hideSavingLock();
        refreshActiveTabUI();
        showToast('Reservation deleted!');
      }, 500);
    }
  }

  /* ==========================================================================
     BOOKING TABLE RENDERER & SEARCH FILTER
     ========================================================================== */
  function renderBookingsTable() {
    const tbody = document.getElementById('bookings-table-body');
    const searchVal = (document.getElementById('booking-search-input')?.value || '').toLowerCase();
    const filterStatus = document.getElementById('booking-filter-status')?.value || 'ALL';

    if (!tbody) return;

    let filtered = bookings.filter(b => {
      const matchesSearch = b.guestName.toLowerCase().includes(searchVal) ||
                            b.contact.toLowerCase().includes(searchVal) ||
                            b.id.toLowerCase().includes(searchVal);
      const matchesStatus = filterStatus === 'ALL' || b.status === filterStatus;
      return matchesSearch && matchesStatus;
    });

    if (filtered.length === 0) {
      tbody.innerHTML = `
        <tr>
          <td colspan="8" class="text-center py-8 text-slate-400 font-medium">
            <i class="fa-solid fa-folder-open text-2xl mb-2 block text-slate-300"></i> No reservation records found.
          </td>
        </tr>`;
      return;
    }

    tbody.innerHTML = filtered.map(b => {
      const room = rooms.find(r => r.id === b.roomId);
      const roomLabel = room ? room.name : b.roomId;

      let statusBadge = '';
      if (b.status === 'LIVE') statusBadge = '<span class="px-2 py-0.5 rounded-full text-[9px] font-bold bg-emerald-100 text-emerald-700">LIVE</span>';
      else if (b.status === 'UPCOMING') statusBadge = '<span class="px-2 py-0.5 rounded-full text-[9px] font-bold bg-blue-100 text-blue-700">UPCOMING</span>';
      else if (b.status === 'CLOSED') statusBadge = '<span class="px-2 py-0.5 rounded-full text-[9px] font-bold bg-slate-100 text-slate-600">CLOSED</span>';
      else statusBadge = '<span class="px-2 py-0.5 rounded-full text-[9px] font-bold bg-rose-100 text-rose-700">INACTIVE</span>';

      return `
        <tr class="hover:bg-slate-50/80 transition text-slate-700 border-b border-slate-100">
          <td class="py-3 px-3 font-bold text-blue-600">${b.id}</td>
          <td class="py-3 px-3">
            <p class="font-bold text-slate-900">${b.guestName}</p>
            <p class="text-[10px] text-slate-400">${b.contact}</p>
          </td>
          <td class="py-3 px-3">
            <span class="font-semibold text-slate-800">${roomLabel}</span>
            <p class="text-[10px] text-slate-400">${b.agent}</p>
          </td>
          <td class="py-3 px-3 text-[10px]">
            <p><strong>In:</strong> ${b.checkIn}</p>
            <p><strong>Out:</strong> ${b.checkOut}</p>
          </td>
          <td class="py-3 px-3 font-semibold">₹${b.tariff + b.extras}</td>
          <td class="py-3 px-3 font-bold ${b.due > 0 ? 'text-rose-600' : 'text-emerald-600'}">₹${b.due}</td>
          <td class="py-3 px-3">${statusBadge}</td>
          <td class="py-3 px-3 text-right space-x-1">
            ${b.idFileData ? `<button onclick="previewIDDocument('${b.id}')" title="View ID Proof" class="p-1.5 bg-blue-50 text-blue-600 hover:bg-blue-100 rounded-lg text-xs transition"><i class="fa-solid fa-id-card"></i></button>` : ''}
            <button onclick="printInvoice('${b.id}')" title="Print Receipt" class="p-1.5 bg-slate-100 text-slate-700 hover:bg-slate-200 rounded-lg text-xs transition"><i class="fa-solid fa-print"></i></button>
            <button onclick="openBookingModal('${b.id}')" title="Edit" class="p-1.5 bg-amber-50 text-amber-600 hover:bg-amber-100 rounded-lg text-xs transition"><i class="fa-solid fa-pen"></i></button>
            <button onclick="deleteBooking('${b.id}')" title="Delete" class="p-1.5 bg-rose-50 text-rose-600 hover:bg-rose-100 rounded-lg text-xs transition"><i class="fa-solid fa-trash"></i></button>
          </td>
        </tr>
      `;
    }).join('');
  }

  /* ==========================================================================
     MASTER ROOMS & AGENTS CONFIGURATION HANDLERS
     ========================================================================== */
  function openRoomModal(roomId = null) {
    const modal = document.getElementById('room-modal');
    const form = document.getElementById('room-form');
    const titleEl = document.getElementById('room-modal-title');
    form.reset();
    document.getElementById('room-form-id').value = '';

    if (roomId) {
      const r = rooms.find(item => item.id === roomId);
      if (r) {
        titleEl.innerText = `Edit Room (${r.id})`;
        document.getElementById('room-form-id').value = r.id;
        document.getElementById('r-name').value = r.name;
        document.getElementById('r-type').value = r.type;
        document.getElementById('r-capacity').value = r.capacity;
        document.getElementById('r-price').value = r.price;
      }
    } else {
      titleEl.innerText = 'Configure New Room';
    }
    modal.classList.remove('hidden');
  }

  function closeRoomModal() {
    document.getElementById('room-modal').classList.add('hidden');
  }

  function handleRoomFormSubmit(e) {
    e.preventDefault();
    const id = document.getElementById('room-form-id').value;
    const roomData = {
      id: id || `RM-${Math.floor(100 + Math.random() * 900)}`,
      name: document.getElementById('r-name').value.trim(),
      type: document.getElementById('r-type').value.trim() || 'Standard',
      capacity: parseInt(document.getElementById('r-capacity').value) || 2,
      price: parseFloat(document.getElementById('r-price').value) || 0
    };

    if (id) {
      const idx = rooms.findIndex(r => r.id === id);
      if (idx !== -1) rooms[idx] = roomData;
    } else {
      rooms.push(roomData);
    }

    syncDatabaseToLocalStorage();
    closeRoomModal();
    refreshActiveTabUI();
    showToast('Room settings saved!');
  }

  function deleteRoomRecord(roomId) {
    requestMasterDataDeletion(() => {
      rooms = rooms.filter(r => r.id !== roomId);
      syncDatabaseToLocalStorage();
      refreshActiveTabUI();
      showToast('Room master deleted!');
    }, `Are you sure you want to delete Room ${roomId}?`);
  }

  function openAgentModal(agentId = null) {
    const modal = document.getElementById('agent-modal');
    const form = document.getElementById('agent-form');
    const titleEl = document.getElementById('agent-modal-title');
    form.reset();
    document.getElementById('agent-form-id').value = '';

    if (agentId) {
      const a = agents.find(item => item.id === agentId);
      if (a) {
        titleEl.innerText = `Edit Agent (${a.id})`;
        document.getElementById('agent-form-id').value = a.id;
        document.getElementById('a-name').value = a.name;
        document.getElementById('a-company').value = a.company || '';
        document.getElementById('a-contact').value = a.contact || '';
      }
    } else {
      titleEl.innerText = 'Configure Booking Agent';
    }
    modal.classList.remove('hidden');
  }

  function closeAgentModal() {
    document.getElementById('agent-modal').classList.add('hidden');
  }

  function handleAgentFormSubmit(e) {
    e.preventDefault();
    const id = document.getElementById('agent-form-id').value;
    const agentData = {
      id: id || `AG-${String(agents.length + 1).padStart(2, '0')}`,
      name: document.getElementById('a-name').value.trim(),
      company: document.getElementById('a-company').value.trim(),
      contact: document.getElementById('a-contact').value.trim()
    };

    if (id) {
      const idx = agents.findIndex(a => a.id === id);
      if (idx !== -1) agents[idx] = agentData;
    } else {
      agents.push(agentData);
    }

    syncDatabaseToLocalStorage();
    closeAgentModal();
    refreshActiveTabUI();
    showToast('Agent master saved!');
  }

  function deleteAgentRecord(agentId) {
    requestMasterDataDeletion(() => {
      agents = agents.filter(a => a.id !== agentId);
      syncDatabaseToLocalStorage();
      refreshActiveTabUI();
      showToast('Agent master deleted!');
    }, `Are you sure you want to delete Agent ${agentId}?`);
  }

  function renderMasterData() {
    const roomList = document.getElementById('master-rooms-list');
    const agentList = document.getElementById('master-agents-list');

    if (roomList) {
      roomList.innerHTML = rooms.map(r => `
        <div class="flex items-center justify-between p-3 bg-slate-50 rounded-2xl border border-slate-200/80 hover:border-blue-300 transition">
          <div>
            <p class="font-bold text-slate-900 text-xs">${r.name} <span class="text-[10px] text-slate-400 font-normal">(${r.type})</span></p>
            <p class="text-[10px] text-slate-500">Cap: ${r.capacity} Guests | Base Rate: <strong class="text-blue-600">₹${r.price}</strong></p>
          </div>
          <div class="space-x-1">
            <button onclick="openRoomModal('${r.id}')" class="p-1.5 bg-amber-50 text-amber-600 rounded-xl text-xs"><i class="fa-solid fa-pen"></i></button>
            <button onclick="deleteRoomRecord('${r.id}')" class="p-1.5 bg-rose-50 text-rose-600 rounded-xl text-xs"><i class="fa-solid fa-trash"></i></button>
          </div>
        </div>
      `).join('');
    }

    if (agentList) {
      agentList.innerHTML = agents.map(a => `
        <div class="flex items-center justify-between p-3 bg-slate-50 rounded-2xl border border-slate-200/80 hover:border-blue-300 transition">
          <div>
            <p class="font-bold text-slate-900 text-xs">${a.name}</p>
            <p class="text-[10px] text-slate-500">${a.company ? a.company + ' | ' : ''}${a.contact || 'No contact'}</p>
          </div>
          <div class="space-x-1">
            <button onclick="openAgentModal('${a.id}')" class="p-1.5 bg-amber-50 text-amber-600 rounded-xl text-xs"><i class="fa-solid fa-pen"></i></button>
            <button onclick="deleteAgentRecord('${a.a.id}')" class="p-1.5 bg-rose-50 text-rose-600 rounded-xl text-xs"><i class="fa-solid fa-trash"></i></button>
          </div>
        </div>
      `).join('');
    }
  }

  /* ==========================================================================
     DASHBOARD ANALYTICS & SLICER ENGINE
     ========================================================================== */
  function handleDashboardYearChange(val) {
    activeDashboardYear = val;
    renderDashboard();
  }

  function handleDashboardSlicerChange(slicer) {
    activeDashboardSlicer = slicer;
    renderDashboard();
  }

  function renderDashboard() {
    let filtered = bookings;

    // Filter by Year Slicer
    if (activeDashboardYear !== 'CONSOLIDATED') {
      filtered = filtered.filter(b => b.checkIn.startsWith(activeDashboardYear));
    }

    // Filter by Quick Slicers
    const todayStr = new Date().toISOString().split('T')[0];
    if (activeDashboardSlicer === 'TODAY_IN') {
      filtered = filtered.filter(b => b.checkIn === todayStr);
    } else if (activeDashboardSlicer === 'TODAY_OUT') {
      filtered = filtered.filter(b => b.checkOut === todayStr);
    } else if (activeDashboardSlicer === 'PENDING_DUE') {
      filtered = filtered.filter(b => b.due > 0);
    }

    // Calculations
    const totalRevenue = filtered.reduce((acc, b) => acc + (b.tariff + b.extras), 0);
    const totalCollected = filtered.reduce((acc, b) => acc + b.advance, 0);
    const totalDue = filtered.reduce((acc, b) => acc + b.due, 0);
    const liveStays = filtered.filter(b => b.status === 'LIVE').length;

    // Update KPI Cards
    const totalRevEl = document.getElementById('dash-total-revenue');
    const totalCollEl = document.getElementById('dash-total-collected');
    const totalDueEl = document.getElementById('dash-total-due');
    const liveStaysEl = document.getElementById('dash-live-stays');

    if (totalRevEl) totalRevEl.innerText = `₹${totalRevenue.toLocaleString('en-IN')}`;
    if (totalCollEl) totalCollEl.innerText = `₹${totalCollected.toLocaleString('en-IN')}`;
    if (totalDueEl) totalDueEl.innerText = `₹${totalDue.toLocaleString('en-IN')}`;
    if (liveStaysEl) liveStaysEl.innerText = liveStays;
  }

  /* ==========================================================================
     AVAILABILITY CALENDAR MATRIX RENDERER
     ========================================================================== */
  function renderCalendar() {
    const grid = document.getElementById('calendar-matrix-grid');
    if (!grid) return;

    const year = document.getElementById('calendar-year-select')?.value || '2026';
    const month = document.getElementById('calendar-month-select')?.value || '08';

    const daysInMonth = new Date(parseInt(year), parseInt(month), 0).getDate();
    let html = `
      <div class="overflow-x-auto">
        <table class="w-full text-[10px] border-collapse">
          <thead>
            <tr class="bg-slate-100 text-slate-600 font-bold">
              <th class="p-2 border border-slate-200 text-left min-w-[100px]">Room</th>
              ${Array.from({ length: daysInMonth }, (_, i) => `<th class="p-1 border border-slate-200 text-center">${i + 1}</th>`).join('')}
            </tr>
          </thead>
          <tbody>
    `;

    rooms.forEach(room => {
      html += `<tr><td class="p-2 border border-slate-200 font-bold text-slate-800 bg-slate-50">${room.name}</td>`;

      for (let day = 1; day <= daysInMonth; day++) {
        const dayStr = `${year}-${month.padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const isBooked = bookings.some(b => b.roomId === room.id && b.checkIn <= dayStr && b.checkOut >= dayStr && b.status !== 'INACTIVE');

        if (isBooked) {
          html += `<td class="p-1 border border-slate-200 bg-rose-500 text-white text-center font-bold" title="Occupied"><i class="fa-solid fa-user-check text-[8px]"></i></td>`;
        } else {
          html += `<td class="p-1 border border-slate-200 bg-emerald-50 text-emerald-600 text-center font-medium hover:bg-emerald-100 transition">Free</td>`;
        }
      }

      html += `</tr>`;
    });

    html += `</tbody></table></div>`;
    grid.innerHTML = html;
  }

  /* ==========================================================================
     ALERTS & NOTIFICATION BADGE
     ========================================================================== */
  function updateAlertBadge() {
    const today = new Date().toISOString().split('T')[0];
    const pendingDueCount = bookings.filter(b => b.due > 0).length;
    const todayCheckInCount = bookings.filter(b => b.checkIn === today).length;

    const totalAlerts = pendingDueCount + todayCheckInCount;
    const badge = document.getElementById('alert-badge');
    if (badge) {
      if (totalAlerts > 0) {
        badge.innerText = totalAlerts;
        badge.classList.remove('hidden');
      } else {
        badge.classList.add('hidden');
      }
    }
  }

  function showAlertModal() {
    const container = document.getElementById('alerts-container');
    const modal = document.getElementById('alerts-modal');
    if (!container || !modal) return;

    const today = new Date().toISOString().split('T')[0];
    let alertsHtml = '';

    const todayIn = bookings.filter(b => b.checkIn === today);
    todayIn.forEach(b => {
      alertsHtml += `
        <div class="p-2.5 bg-blue-50 border border-blue-200 rounded-xl text-xs flex justify-between items-center">
          <div>
            <p class="font-bold text-blue-900"><i class="fa-solid fa-right-to-bracket text-blue-600"></i> Check-In Today</p>
            <p class="text-[10px] text-blue-700">${b.guestName} - Room ${b.roomId}</p>
          </div>
          <button onclick="openBookingModal('${b.id}')" class="px-2 py-1 bg-blue-600 text-white rounded-lg text-[10px] font-bold">View</button>
        </div>
      `;
    });

    const pendingDue = bookings.filter(b => b.due > 0);
    pendingDue.forEach(b => {
      alertsHtml += `
        <div class="p-2.5 bg-rose-50 border border-rose-200 rounded-xl text-xs flex justify-between items-center">
          <div>
            <p class="font-bold text-rose-900"><i class="fa-solid fa-circle-exclamation text-rose-600"></i> Balance Pending</p>
            <p class="text-[10px] text-rose-700">${b.guestName} - Due: ₹${b.due}</p>
          </div>
          <button onclick="openBookingModal('${b.id}')" class="px-2 py-1 bg-rose-600 text-white rounded-lg text-[10px] font-bold">Collect</button>
        </div>
      `;
    });

    container.innerHTML = alertsHtml || '<p class="text-center text-slate-400 py-4 text-xs">No pending alerts at this time.</p>';
    modal.classList.remove('hidden');
  }

  function closeAlertModal() {
    document.getElementById('alerts-modal').classList.add('hidden');
  }

  /* ==========================================================================
     ID PROOF PREVIEW & INVOICE PRINTING / JPEG EXPORT ENGINE
     ========================================================================== */
  function previewIDDocument(bookingId) {
    const b = bookings.find(item => item.id === bookingId);
    if (!b || !b.idFileData) {
      alert('No attached ID document found!');
      return;
    }

    const container = document.getElementById('id-preview-content');
    if (b.idFileData.startsWith('data:image')) {
      container.innerHTML = `<img src="${b.idFileData}" class="max-h-[60vh] rounded-xl shadow-sm object-contain" alt="ID Document" />`;
    } else {
      container.innerHTML = `<a href="${b.idFileData}" target="_blank" class="px-4 py-2 bg-blue-600 text-white font-bold text-xs rounded-xl shadow-sm">Download Attached Document</a>`;
    }

    document.getElementById('id-preview-modal').classList.remove('hidden');
  }

  function closeIDPreviewModal() {
    document.getElementById('id-preview-modal').classList.add('hidden');
  }

  function printInvoice(bookingId) {
    const b = bookings.find(item => item.id === bookingId);
    if (!b) return;

    document.getElementById('inv-id').innerText = b.id;
    document.getElementById('inv-date').innerText = new Date().toLocaleDateString('en-IN');
    document.getElementById('inv-guest-name').innerText = b.guestName;
    document.getElementById('inv-guest-contact').innerText = b.contact;
    document.getElementById('inv-guest-id').innerText = b.idNumber ? `ID: ${b.idNumber}` : '';
    document.getElementById('inv-room').innerText = b.roomId;
    document.getElementById('inv-checkin').innerText = b.checkIn;
    document.getElementById('inv-checkout').innerText = b.checkOut;
    document.getElementById('inv-tariff-total').innerText = `₹${b.tariff}`;
    document.getElementById('inv-extras').innerText = `₹${b.extras}`;
    document.getElementById('inv-grand-total').innerText = `₹${b.tariff + b.extras}`;
    document.getElementById('inv-advance').innerText = `₹${b.advance}`;
    document.getElementById('inv-due').innerText = `₹${b.due}`;

    const printContainer = document.getElementById('printable-invoice');
    printContainer.classList.remove('hidden');

    window.print();

    setTimeout(() => {
      printContainer.classList.add('hidden');
    }, 1000);
  }

  /* ==========================================================================
     DATA WIPE & GOOGLE SHEETS CLOUD SYNC
     ========================================================================== */
  function openDataWipeModal() {
    document.getElementById('wipe-password-input').value = '';
    document.getElementById('wipe-error-msg').classList.add('hidden');
    document.getElementById('wipe-data-modal').classList.remove('hidden');
  }

  function closeDataWipeModal() {
    document.getElementById('wipe-data-modal').classList.add('hidden');
  }

  function handleDataWipeSubmit(e) {
    e.preventDefault();
    const pwd = document.getElementById('wipe-password-input').value.trim();
    if (pwd === DATA_WIPE_PASSWORD) {
      localStorage.clear();
      sessionStorage.clear();
      alert('Factory Reset Complete! Portal will now reload.');
      window.location.reload();
    } else {
      document.getElementById('wipe-error-msg').classList.remove('hidden');
    }
  }

  function syncToGoogleSheets() {
    if (!GOOGLE_SCRIPT_URL || GOOGLE_SCRIPT_URL.includes('YOUR_GOOGLE_APPS_SCRIPT_URL_HERE')) {
      alert('Please configure your Google Apps Script URL in the source code first!');
      return;
    }

    showSavingLock();

    const payload = {
      bookings: bookings,
      rooms: rooms,
      agents: agents
    };

    fetch(GOOGLE_SCRIPT_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    })
    .then(() => {
      hideSavingLock();
      showToast('Cloud Data Sync Completed Successfully!');
    })
    .catch(err => {
      hideSavingLock();
      alert('Cloud Sync Failed: ' + err.message);
    });
  }
</script>
</body>
</html>
