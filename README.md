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

    /* Custom Slicer & Hover Navigation Control Styles */
    .slicer-btn {
      user-select: none;
      cursor: pointer;
    }

    .slicer-active-all { background-color: #2563EB !important; color: #FFFFFF !important; }
    .slicer-active-live { background-color: #D97706 !important; color: #FFFFFF !important; }
    .slicer-active-upcoming { background-color: #2563EB !important; color: #FFFFFF !important; }
    .slicer-active-closed { background-color: #059669 !important; color: #FFFFFF !important; }
    .slicer-active-inactive { background-color: #E11D48 !important; color: #FFFFFF !important; }
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
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>1. Google Chrome/Microsoft Edge is best view browser for this Portal.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>2. Take backup every day or every week.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>3. Do not force close 'The Portal'; always close it using the 'Logout' option.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>4. Do not 'Login' with multiple device/multiple browser/multiple browser tab at a same time to avoid data merge.</span></p>
        <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>5. Data save/Data fetch take little bit time so hold on ⏳.</span></p>
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
    <div class="bg-white rounded-3xl shadow-2xl border border-rose-100 max-w-sm w-full p-5 space-y-3 text-center">
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
        <h3 class="text-lg font-black text-slate-900">Saving & Logging Out...</h3>
        <p class="text-xs text-rose-600 mt-2 font-bold uppercase">Do not close window or shutdown!</p>
        <p class="text-[10px] text-slate-500 mt-1">Please wait while we secure your data.</p>
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
        <p class="text-[10px] text-slate-500 mt-1">You will be logged out automatically in <strong id="logout-countdown-seconds" class="text-rose-600">60</strong> seconds due to inactivity. Data will be saved securely.</p>
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

      <!-- 1. NEW 4 YEAR-WISE BOOKING COUNT BOXES -->
      <div class="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <div id="slicer-box-live" onclick="setDashboardStatusSlicer('live')" class="cursor-pointer bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 hover:border-amber-400 transition flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-amber-800 tracking-wider">Live Booking Count Total</p>
            <p id="dash-live-count" class="text-xl font-black text-amber-600 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-amber-50 text-amber-600 rounded-2xl"><i class="fa-solid fa-hotel text-base"></i></div>
        </div>
        <div id="slicer-box-upcoming" onclick="setDashboardStatusSlicer('upcoming')" class="cursor-pointer bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 hover:border-blue-400 transition flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-blue-800 tracking-wider">Upcoming Booking Count Total</p>
            <p id="dash-upcoming-count" class="text-xl font-black text-blue-600 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-blue-50 text-blue-600 rounded-2xl"><i class="fa-solid fa-calendar-check text-base"></i></div>
        </div>
        <div id="slicer-box-closed" onclick="setDashboardStatusSlicer('closed')" class="cursor-pointer bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 hover:border-emerald-400 transition flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-emerald-800 tracking-wider">Closed Booking Count Total</p>
            <p id="dash-closed-count" class="text-xl font-black text-emerald-600 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-emerald-50 text-emerald-600 rounded-2xl"><i class="fa-solid fa-circle-check text-base"></i></div>
        </div>
        <div id="slicer-box-inactive" onclick="setDashboardStatusSlicer('inactive')" class="cursor-pointer bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 hover:border-rose-400 transition flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-rose-800 tracking-wider">Inactive Booking Count Total</p>
            <p id="dash-inactive-count" class="text-xl font-black text-rose-600 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-rose-50 text-rose-600 rounded-2xl"><i class="fa-solid fa-ban text-base"></i></div>
        </div>
      </div>

      <!-- 2. EXCEL SLICER TAB CONTROL -->
      <div class="bg-white p-3 rounded-2xl border border-slate-200/60 flex flex-wrap items-center justify-between gap-2 shadow-xs">
        <div class="flex items-center space-x-1">
          <span class="text-[10px] font-bold text-slate-500 uppercase mr-2 flex items-center gap-1">
            <i class="fa-solid fa-filter text-emerald-600"></i> Status Slicer:
          </span>
          <button id="slicer-btn-all" onclick="setDashboardStatusSlicer('all')" class="slicer-btn px-3 py-1 rounded-xl text-[10px] font-bold transition bg-blue-600 text-white shadow-xs">All</button>
          <button id="slicer-btn-live" onclick="setDashboardStatusSlicer('live')" class="slicer-btn px-3 py-1 rounded-xl text-[10px] font-bold transition bg-slate-100 text-slate-600 hover:bg-slate-200">Live</button>
          <button id="slicer-btn-upcoming" onclick="setDashboardStatusSlicer('upcoming')" class="slicer-btn px-3 py-1 rounded-xl text-[10px] font-bold transition bg-slate-100 text-slate-600 hover:bg-slate-200">Upcoming</button>
          <button id="slicer-btn-closed" onclick="setDashboardStatusSlicer('closed')" class="slicer-btn px-3 py-1 rounded-xl text-[10px] font-bold transition bg-slate-100 text-slate-600 hover:bg-slate-200">Closed</button>
          <button id="slicer-btn-inactive" onclick="setDashboardStatusSlicer('inactive')" class="slicer-btn px-3 py-1 rounded-xl text-[10px] font-bold transition bg-slate-100 text-slate-600 hover:bg-slate-200">Inactive</button>
        </div>
        <span id="dash-slicer-active-label" class="text-[10px] font-semibold text-slate-400">Filtering: ALL Statuses</span>
      </div>

      <!-- Financial Monetary Summaries -->
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
      <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col md:flex-row justify-between items-stretch md:items-center gap-3">
        <div class="flex flex-wrap items-center gap-2">
          <button onclick="openBookingModal()" class="bg-blue-600 hover:bg-blue-700 active:scale-95 text-white font-bold px-4 py-2 rounded-2xl shadow-sm transition flex items-center gap-1.5 text-xs">
            <i class="fa-solid fa-plus text-xs"></i> Add New Booking
          </button>
          
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-3 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-magnifying-glass"></i>
            </span>
            <input type="text" id="booking-search" oninput="filterBookingTable()" placeholder="Search Guest Name, Room, Mobile, SL..." class="bg-slate-100 border border-transparent focus:border-blue-500 rounded-2xl pl-9 pr-3 py-1.5 focus:outline-none focus:bg-white text-xs w-64 transition font-medium" />
          </div>

          <div class="flex items-center space-x-1.5 bg-slate-100 p-1 rounded-2xl">
            <span class="text-[10px] font-bold text-slate-500 uppercase px-2">Filter:</span>
            <select id="booking-year-filter" onchange="filterBookingTable()" class="bg-white text-slate-800 text-[11px] font-bold rounded-xl px-2 py-1 focus:outline-none border border-slate-200">
            </select>
            <select id="booking-status-filter" onchange="filterBookingTable()" class="bg-white text-slate-800 text-[11px] font-bold rounded-xl px-2 py-1 focus:outline-none border border-slate-200">
              <option value="ALL">All Statuses</option>
              <option value="Live">Live</option>
              <option value="Upcoming">Upcoming</option>
              <option value="Closed">Closed</option>
              <option value="Inactive">Inactive</option>
            </select>
          </div>
        </div>

        <div class="flex items-center space-x-2 text-slate-500 text-[11px] font-medium self-end md:self-auto">
          <span>Total Records: <strong id="booking-count" class="text-slate-900 font-bold">0</strong></span>
        </div>
      </div>

      <!-- Booking Table Container -->
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 overflow-hidden relative">
        <div class="overflow-x-auto max-h-[600px] overflow-y-auto">
          <table class="w-full text-left border-collapse text-[11px]">
            <thead class="bg-slate-50 text-slate-500 uppercase text-[9px] font-bold tracking-wider sticky top-0 z-10 border-b border-slate-200/80">
              <tr>
                <th class="py-3 px-3">SL</th>
                <th class="py-3 px-3">Booking ID</th>
                <th class="py-3 px-3">Guest Name</th>
                <th class="py-3 px-3">Mobile No.</th>
                <th class="py-3 px-3">Room Allocated</th>
                <th class="py-3 px-3">Check-In</th>
                <th class="py-3 px-3">Check-Out</th>
                <th class="py-3 px-3">Agent</th>
                <th class="py-3 px-3 text-right">Total (₹)</th>
                <th class="py-3 px-3 text-right">Advance (₹)</th>
                <th class="py-3 px-3 text-right">Due (₹)</th>
                <th class="py-3 px-3 text-center">Status</th>
                <th class="py-3 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="booking-table-body" class="divide-y divide-slate-100 font-medium text-slate-700">
              <!-- Dynamic Rows inserted by JavaScript -->
            </tbody>
          </table>
        </div>
      </div>
    </section>

  </main>

  <!-- ADD / EDIT BOOKING MODAL -->
  <div id="booking-modal" class="hidden fixed inset-0 z-50 bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-2xl w-full p-6 space-y-4 max-h-[90vh] overflow-y-auto text-left">
      <div class="flex justify-between items-center pb-3 border-b border-slate-100">
        <h3 id="modal-title" class="text-sm font-bold text-slate-900 flex items-center gap-2">
          <i class="fa-solid fa-pen-to-square text-blue-600"></i> New Booking Entry
        </h3>
        <button onclick="closeBookingModal()" class="text-slate-400 hover:text-slate-600 p-1 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <form id="booking-form" onsubmit="handleBookingSubmit(event)" class="space-y-4 text-xs">
        <input type="hidden" id="edit-booking-sl" value="" />

        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div>
            <label class="block font-semibold text-slate-700 mb-1">Guest Full Name *</label>
            <input type="text" id="form-guest-name" required placeholder="Enter primary guest name" class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition" />
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Mobile / Phone Number *</label>
            <input type="text" id="form-guest-mobile" required placeholder="Enter contact number" class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition" />
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Allocation / Room No *</label>
            <select id="form-room" required class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition">
            </select>
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Booking Agent / Source *</label>
            <select id="form-agent" required class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition">
            </select>
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Check-In Date *</label>
            <input type="date" id="form-checkin" min="2026-08-01" max="2085-12-31" required onchange="calculateBookingFinancials()" class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition" />
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Check-Out Date *</label>
            <input type="date" id="form-checkout" min="2026-08-01" max="2085-12-31" required onchange="calculateBookingFinancials()" class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition" />
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Total Booking Amount (₹) *</label>
            <input type="number" id="form-total-amount" min="0" step="any" required oninput="calculateBookingFinancials()" placeholder="0.00" class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition font-bold text-slate-900" />
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Advance Received (₹)</label>
            <input type="number" id="form-advance-amount" min="0" step="any" oninput="calculateBookingFinancials()" placeholder="0.00" class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition font-bold text-emerald-600" />
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Balance Due (₹)</label>
            <input type="number" id="form-due-amount" readonly placeholder="0.00" class="w-full bg-slate-100 border border-slate-200 rounded-2xl px-3 py-2 font-bold text-rose-600 cursor-not-allowed" />
          </div>

          <div>
            <label class="block font-semibold text-slate-700 mb-1">Status Classification *</label>
            <select id="form-status" required class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition font-semibold">
              <option value="Live">Live</option>
              <option value="Upcoming">Upcoming</option>
              <option value="Closed">Closed</option>
              <option value="Inactive">Inactive</option>
            </select>
          </div>
        </div>

        <div>
          <label class="block font-semibold text-slate-700 mb-1">Special Notes / Remarks</label>
          <textarea id="form-remarks" rows="2" placeholder="Add optional details like ID proof, meal plan, extra bed..." class="w-full bg-slate-50 border border-slate-200 focus:border-blue-500 rounded-2xl px-3 py-2 focus:outline-none focus:bg-white transition"></textarea>
        </div>

        <div class="flex justify-end space-x-2 pt-3 border-t border-slate-100">
          <button type="button" onclick="closeBookingModal()" class="bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold px-4 py-2 rounded-2xl transition">Cancel</button>
          <button type="submit" class="bg-blue-600 hover:bg-blue-700 text-white font-bold px-5 py-2 rounded-2xl shadow-sm transition flex items-center gap-1.5">
            <i class="fa-solid fa-floppy-disk"></i> Save Booking Record
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- PRINTABLE INVOICE / RECEIPT TEMPLATE (VISIBLE IN PRINT OR GENERATED JPEG) -->
  <div id="printable-invoice" class="hidden bg-white p-8 max-w-2xl mx-auto border border-slate-200 rounded-3xl space-y-6 text-slate-800 text-xs">
    <div class="flex justify-between items-start border-b border-slate-200 pb-4">
      <div>
        <h1 class="text-xl font-black text-blue-600 uppercase tracking-wide">Hotel Business Portal</h1>
        <p class="text-[10px] text-slate-500">Official Booking Receipt &amp; Tax Invoice</p>
      </div>
      <div class="text-right">
        <p class="font-bold text-slate-900" id="inv-id">INV-0000</p>
        <p class="text-[10px] text-slate-500" id="inv-date">Date: DD/MM/YYYY</p>
      </div>
    </div>

    <div class="grid grid-cols-2 gap-4 bg-slate-50 p-4 rounded-2xl border border-slate-100">
      <div>
        <p class="text-[10px] uppercase font-bold text-slate-400">Guest Information</p>
        <p class="font-bold text-slate-900 text-sm mt-0.5" id="inv-guest-name">Guest Name</p>
        <p class="text-slate-600" id="inv-guest-mobile">Mobile: +91 0000000000</p>
      </div>
      <div class="text-right">
        <p class="text-[10px] uppercase font-bold text-slate-400">Booking Summary</p>
        <p class="font-semibold text-slate-800 mt-0.5">Room: <span id="inv-room" class="font-bold text-blue-600">Room 101</span></p>
        <p class="text-slate-600">Agent: <span id="inv-agent">Direct</span></p>
      </div>
    </div>

    <div class="border border-slate-200 rounded-2xl overflow-hidden">
      <table class="w-full text-left text-xs">
        <thead class="bg-slate-100 text-slate-600 font-bold uppercase text-[9px]">
          <tr>
            <th class="py-2.5 px-3">Check-In</th>
            <th class="py-2.5 px-3">Check-Out</th>
            <th class="py-2.5 px-3 text-right">Total Tariff</th>
            <th class="py-2.5 px-3 text-right">Advance Paid</th>
            <th class="py-2.5 px-3 text-right">Balance Due</th>
          </tr>
        </thead>
        <tbody class="divide-y divide-slate-100">
          <tr>
            <td class="py-3 px-3 font-medium" id="inv-checkin">01/08/2026</td>
            <td class="py-3 px-3 font-medium" id="inv-checkout">03/08/2026</td>
            <td class="py-3 px-3 text-right font-bold text-slate-900" id="inv-total">₹0.00</td>
            <td class="py-3 px-3 text-right font-bold text-emerald-600" id="inv-advance">₹0.00</td>
            <td class="py-3 px-3 text-right font-bold text-rose-600" id="inv-due">₹0.00</td>
          </tr>
        </tbody>
      </table>
    </div>

    <div class="flex justify-between items-end pt-4 border-t border-slate-100">
      <div>
        <p class="text-[10px] text-slate-400 italic">Thank you for staying with us!</p>
        <p class="text-[9px] text-slate-400">This is a computer-generated document.</p>
      </div>
      <div class="text-center">
        <div class="w-28 border-b border-slate-400 mb-1"></div>
        <p class="text-[10px] font-bold text-slate-600">Authorized Signatory</p>
      </div>
    </div>
  </div>





<!-- MASTER DATA TAB -->
    <section id="tab-master" class="tab-content hidden space-y-4">
      <!-- Locked State Overlay -->
      <div id="master-locked-view" class="bg-white rounded-3xl p-8 shadow-sm border border-slate-200/60 text-center space-y-3">
        <div class="bg-amber-50 text-amber-600 w-12 h-12 rounded-2xl flex items-center justify-center mx-auto text-xl shadow-sm">
          <i class="fa-solid fa-lock"></i>
        </div>
        <div>
          <h3 class="text-sm font-bold text-slate-900">Master Data Configuration is Locked</h3>
          <p class="text-[11px] text-slate-500 mt-0.5">Please enter the master password to view room lists, agent directories, or execute administrative data resets.</p>
        </div>
        <button onclick="openMasterAuthModal()" class="bg-rose-600 hover:bg-rose-700 text-white font-bold px-5 py-2 rounded-2xl text-xs transition shadow-sm flex items-center gap-1.5 mx-auto">
          <i class="fa-solid fa-key"></i> Unlock Master Settings
        </button>
      </div>

      <!-- Unlocked Settings Panel -->
      <div id="master-unlocked-view" class="hidden space-y-4">
        <div class="bg-emerald-50 border border-emerald-200 rounded-2xl p-3 flex justify-between items-center text-emerald-800 text-[11px]">
          <span class="font-bold flex items-center gap-1.5">
            <i class="fa-solid fa-shield-halved text-emerald-600"></i> Master Controls Unlocked &amp; Active
          </span>
          <button onclick="lockMasterData()" class="bg-emerald-700 hover:bg-emerald-800 text-white font-semibold px-3 py-1 rounded-xl text-[10px] transition">
            Lock Settings Now
          </button>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
          <!-- Room Management -->
          <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 space-y-3">
            <div class="flex justify-between items-center border-b border-slate-100 pb-2">
              <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                <i class="fa-solid fa-door-open text-blue-600"></i> Room / Allocation Directory
              </h3>
              <span id="master-room-count" class="text-[10px] font-bold text-slate-400">0 Rooms</span>
            </div>
            
            <form onsubmit="handleMasterAddRoom(event)" class="flex gap-2">
              <input type="text" id="master-room-input" required placeholder="Add new room (e.g., Room 105)" class="flex-1 bg-slate-50 border border-slate-200 rounded-2xl px-3 py-1.5 text-xs focus:outline-none focus:border-blue-500 font-medium" />
              <button type="submit" class="bg-blue-600 hover:bg-blue-700 text-white font-bold px-3 py-1.5 rounded-2xl text-xs transition shadow-sm">
                Add Room
              </button>
            </form>

            <div class="max-h-48 overflow-y-auto pr-1">
              <ul id="master-room-list" class="divide-y divide-slate-100 text-xs">
                <!-- Dynamic Room List -->
              </ul>
            </div>
          </div>

          <!-- Booking Agent / Source Directory -->
          <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 space-y-3">
            <div class="flex justify-between items-center border-b border-slate-100 pb-2">
              <h3 class="text-xs font-bold text-slate-900 flex items-center gap-1.5">
                <i class="fa-solid fa-user-tie text-indigo-600"></i> Agent / Source Directory
              </h3>
              <span id="master-agent-count" class="text-[10px] font-bold text-slate-400">0 Agents</span>
            </div>

            <form onsubmit="handleMasterAddAgent(event)" class="flex gap-2">
              <input type="text" id="master-agent-input" required placeholder="Add agent or source (e.g., MakeMyTrip)" class="flex-1 bg-slate-50 border border-slate-200 rounded-2xl px-3 py-1.5 text-xs focus:outline-none focus:border-indigo-500 font-medium" />
              <button type="submit" class="bg-indigo-600 hover:bg-indigo-700 text-white font-bold px-3 py-1.5 rounded-2xl text-xs transition shadow-sm">
                Add Agent
              </button>
            </form>

            <div class="max-h-48 overflow-y-auto pr-1">
              <ul id="master-agent-list" class="divide-y divide-slate-100 text-xs">
                <!-- Dynamic Agent List -->
              </ul>
            </div>
          </div>
        </div>
      </div>
    </section>

    <!-- CALENDAR VIEW TAB -->
    <section id="tab-calendar" class="tab-content hidden space-y-4">
      <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col md:flex-row justify-between items-stretch md:items-center gap-3">
        <div class="flex items-center space-x-2">
          <div class="bg-blue-50 text-blue-600 p-2 rounded-2xl">
            <i class="fa-solid fa-calendar-days text-base"></i>
          </div>
          <div>
            <h3 class="text-xs font-bold text-slate-900">Interactive Calendar Grid (2026 – 2085)</h3>
            <p class="text-[10px] text-slate-500">View month-by-month occupancy. Red indicator triangles show dates with existing bookings.</p>
          </div>
        </div>

        <div class="flex items-center space-x-2">
          <select id="cal-month-select" onchange="renderCalendarView()" class="bg-slate-100 border border-slate-200 font-bold text-slate-800 text-xs rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white transition">
            <option value="0">January</option>
            <option value="1">February</option>
            <option value="2">March</option>
            <option value="3">April</option>
            <option value="4">May</option>
            <option value="5">June</option>
            <option value="6">July</option>
            <option value="7">August</option>
            <option value="8">September</option>
            <option value="9">October</option>
            <option value="10">November</option>
            <option value="11">December</option>
          </select>

          <select id="cal-year-select" onchange="renderCalendarView()" class="bg-slate-100 border border-slate-200 font-bold text-slate-800 text-xs rounded-2xl px-3 py-1.5 focus:outline-none focus:bg-white transition">
          </select>

          <button onclick="jumpToTodayCalendar()" class="bg-blue-600 hover:bg-blue-700 text-white font-bold px-3 py-1.5 rounded-2xl text-xs transition shadow-sm">
            Today
          </button>
        </div>
      </div>

      <!-- Calendar Table Grid -->
      <div class="bg-white rounded-3xl shadow-sm border border-slate-200/60 p-4 overflow-x-auto">
        <div class="min-w-[700px]">
          <div class="grid grid-cols-7 gap-1 text-center font-bold text-[10px] text-slate-400 uppercase py-2 border-b border-slate-100">
            <div class="text-rose-500">Sun</div>
            <div>Mon</div>
            <div>Tue</div>
            <div>Wed</div>
            <div>Thu</div>
            <div>Fri</div>
            <div class="text-blue-500">Sat</div>
          </div>
          <div id="calendar-grid" class="grid grid-cols-7 gap-1 text-xs pt-2">
            <!-- Dynamic Days Rendered Here -->
          </div>
        </div>
      </div>
    </section>

  <!-- FLOATING HOVER NAVIGATION CONTROLS (Page Up / Down & Bar Left / Right) -->
  <div class="fixed bottom-6 right-6 z-30 flex flex-col items-center gap-2 no-print opacity-80 hover:opacity-100 transition-opacity">
    <div class="flex flex-col bg-slate-900/80 backdrop-blur-md text-white p-1 rounded-2xl shadow-2xl border border-slate-700">
      <button onclick="scrollPageVertical('up')" title="Page Up" class="p-2.5 hover:bg-slate-700/80 rounded-xl transition text-xs">
        <i class="fa-solid fa-chevron-up"></i>
      </button>
      <button onclick="scrollPageVertical('down')" title="Page Down" class="p-2.5 hover:bg-slate-700/80 rounded-xl transition text-xs border-t border-slate-800">
        <i class="fa-solid fa-chevron-down"></i>
      </button>
    </div>

    <div class="flex bg-slate-900/80 backdrop-blur-md text-white p-1 rounded-2xl shadow-2xl border border-slate-700">
      <button onclick="scrollPageHorizontal('left')" title="Bar Left" class="p-2.5 hover:bg-slate-700/80 rounded-xl transition text-xs">
        <i class="fa-solid fa-chevron-left"></i>
      </button>
      <button onclick="scrollPageHorizontal('right')" title="Bar Right" class="p-2.5 hover:bg-slate-700/80 rounded-xl transition text-xs border-l border-slate-800">
        <i class="fa-solid fa-chevron-right"></i>
      </button>
    </div>
  </div>



<!-- JAVASCRIPT LOGIC -->
  <script>
    // --- APP STATE ---
    let appData = {
      bookings: [],
      rooms: ['Room 101', 'Room 102', 'Room 103', 'Room 104', 'Room 105'],
      agents: ['Direct', 'MakeMyTrip', 'Goibibo', 'Agoda', 'Booking.com']
    };

    let activeDashboardSlicer = 'all';
    let isMasterUnlocked = false;
    let masterDeleteTarget = null;
    let inactivityTimer = null;
    let countdownInterval = null;
    let secondsRemaining = 60;

    const GOOGLE_SCRIPT_URL = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL'; // Replace with actual Web App URL

    // --- INITIALIZATION ---
    window.addEventListener('DOMContentLoaded', () => {
      populateYearDropdowns();
      loadDataFromStorage();
      resetInactivityTimer();
      setupInactivityListeners();
    });

    function populateYearDropdowns() {
      const currentYear = new Date().getFullYear();
      const startYear = 2026;
      const endYear = 2085;

      const dashSelect = document.getElementById('dash-year-select');
      const bookFilter = document.getElementById('booking-year-filter');
      const calSelect = document.getElementById('cal-year-select');

      dashSelect.innerHTML = '<option value="ALL">All Years (Consolidated)</option>';
      bookFilter.innerHTML = '<option value="ALL">All Years</option>';
      calSelect.innerHTML = '';

      for (let y = startYear; y <= endYear; y++) {
        dashSelect.innerHTML += `<option value="${y}">${y}</option>`;
        bookFilter.innerHTML += `<option value="${y}">${y}</option>`;
        calSelect.innerHTML += `<option value="${y}">${y}</option>`;
      }

      const activeYear = currentYear >= startYear && currentYear <= endYear ? currentYear : startYear;
      dashSelect.value = activeYear;
      bookFilter.value = 'ALL';
      calSelect.value = activeYear;
    }

    // --- TAB NAVIGATION ---
    function switchTab(tabId) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
        btn.classList.add('text-slate-600');
      });

      const activeTabEl = document.getElementById(`tab-${tabId}`);
      const activeBtnEl = document.getElementById(`btn-${tabId}`);

      if (activeTabEl) activeTabEl.classList.remove('hidden');
      if (activeBtnEl) {
        activeBtnEl.classList.add('bg-white', 'text-blue-600', 'shadow-sm', 'font-bold');
        activeBtnEl.classList.remove('text-slate-600');
      }

      if (tabId === 'dashboard') updateDashboardMetrics();
      if (tabId === 'booking') renderBookingTable();
      if (tabId === 'calendar') renderCalendarView();
    }

    // --- DASHBOARD METRICS & SLICER ---
    function handleDashboardYearChange(yearVal) {
      const dashSelect = document.getElementById('dash-year-select');
      if (yearVal === 'CURRENT') {
        const cur = new Date().getFullYear();
        dashSelect.value = (cur >= 2026 && cur <= 2085) ? cur : 2026;
      } else {
        dashSelect.value = yearVal;
      }
      updateDashboardMetrics();
    }

    function setDashboardStatusSlicer(status) {
      activeDashboardSlicer = status;
      const buttons = {
        all: document.getElementById('slicer-btn-all'),
        live: document.getElementById('slicer-btn-live'),
        upcoming: document.getElementById('slicer-btn-upcoming'),
        closed: document.getElementById('slicer-btn-closed'),
        inactive: document.getElementById('slicer-btn-inactive')
      };

      Object.keys(buttons).forEach(key => {
        if (buttons[key]) {
          buttons[key].className = 'slicer-btn px-3 py-1 rounded-xl text-[10px] font-bold transition bg-slate-100 text-slate-600 hover:bg-slate-200';
        }
      });

      const activeBtn = buttons[status];
      if (activeBtn) {
        activeBtn.className = `slicer-btn px-3 py-1 rounded-xl text-[10px] font-bold transition slicer-active-${status}`;
      }

      document.getElementById('dash-slicer-active-label').innerText = `Filtering: ${status.toUpperCase()} Statuses`;
      updateDashboardMetrics();
    }

    function updateDashboardMetrics() {
      const selectedYear = document.getElementById('dash-year-select').value;
      const filterLabel = document.getElementById('dash-filter-label');

      filterLabel.innerText = selectedYear === 'ALL' ? 'Consolidated (All Years)' : `Year ${selectedYear}`;

      let filtered = appData.bookings.filter(b => {
        if (selectedYear === 'ALL') return true;
        const bYear = new Date(b.checkIn).getFullYear().toString();
        return bYear === selectedYear;
      });

      // Calculate 4 Main Status Counts
      const liveCount = filtered.filter(b => b.status === 'Live').length;
      const upcomingCount = filtered.filter(b => b.status === 'Upcoming').length;
      const closedCount = filtered.filter(b => b.status === 'Closed').length;
      const inactiveCount = filtered.filter(b => b.status === 'Inactive').length;

      document.getElementById('dash-live-count').innerText = liveCount;
      document.getElementById('dash-upcoming-count').innerText = upcomingCount;
      document.getElementById('dash-closed-count').innerText = closedCount;
      document.getElementById('dash-inactive-count').innerText = inactiveCount;

      // Apply Excel Status Slicer Filter
      if (activeDashboardSlicer !== 'all') {
        const statusMap = {
          live: 'Live',
          upcoming: 'Upcoming',
          closed: 'Closed',
          inactive: 'Inactive'
        };
        filtered = filtered.filter(b => b.status === statusMap[activeDashboardSlicer]);
      }

      // Calculate Monetary Totals
      const totalBookings = filtered.length;
      const totalAmount = filtered.reduce((sum, b) => sum + (parseFloat(b.totalAmount) || 0), 0);
      const totalAdvance = filtered.reduce((sum, b) => sum + (parseFloat(b.advanceAmount) || 0), 0);
      const totalDue = filtered.reduce((sum, b) => sum + (parseFloat(b.dueAmount) || 0), 0);

      document.getElementById('dash-total-bookings').innerText = totalBookings;
      document.getElementById('dash-total-amount').innerText = `₹${totalAmount.toLocaleString('en-IN')}`;
      document.getElementById('dash-advanced').innerText = `₹${totalAdvance.toLocaleString('en-IN')}`;
      document.getElementById('dash-due').innerText = `₹${totalDue.toLocaleString('en-IN')}`;
    }

    // --- BOOKING MANAGEMENT ---
    function populateDropdowns() {
      const roomSelect = document.getElementById('form-room');
      const agentSelect = document.getElementById('form-agent');

      roomSelect.innerHTML = appData.rooms.map(r => `<option value="${r}">${r}</option>`).join('');
      agentSelect.innerHTML = appData.agents.map(a => `<option value="${a}">${a}</option>`).join('');
    }

    function openBookingModal(editSl = null) {
      populateDropdowns();
      const form = document.getElementById('booking-form');
      form.reset();

      if (editSl !== null) {
        const item = appData.bookings.find(b => b.sl == editSl);
        if (item) {
          document.getElementById('modal-title').innerText = 'Edit Booking Record';
          document.getElementById('edit-booking-sl').value = item.sl;
          document.getElementById('form-guest-name').value = item.guestName;
          document.getElementById('form-guest-mobile').value = item.guestMobile;
          document.getElementById('form-room').value = item.room;
          document.getElementById('form-agent').value = item.agent;
          document.getElementById('form-checkin').value = item.checkIn;
          document.getElementById('form-checkout').value = item.checkOut;
          document.getElementById('form-total-amount').value = item.totalAmount;
          document.getElementById('form-advance-amount').value = item.advanceAmount;
          document.getElementById('form-due-amount').value = item.dueAmount;
          document.getElementById('form-status').value = item.status;
          document.getElementById('form-remarks').value = item.remarks || '';
        }
      } else {
        document.getElementById('modal-title').innerText = 'New Booking Entry';
        document.getElementById('edit-booking-sl').value = '';
      }

      document.getElementById('booking-modal').classList.remove('hidden');
    }

    function closeBookingModal() {
      document.getElementById('booking-modal').classList.add('hidden');
    }

    function calculateBookingFinancials() {
      const total = parseFloat(document.getElementById('form-total-amount').value) || 0;
      const advance = parseFloat(document.getElementById('form-advance-amount').value) || 0;
      const due = Math.max(0, total - advance);
      document.getElementById('form-due-amount').value = due.toFixed(2);
    }

    function handleBookingSubmit(e) {
      e.preventDefault();
      const editSl = document.getElementById('edit-booking-sl').value;

      const record = {
        sl: editSl ? parseInt(editSl) : Date.now(),
        bookingId: editSl ? appData.bookings.find(b => b.sl == editSl).bookingId : 'BK-' + Math.floor(1000 + Math.random() * 9000),
        guestName: document.getElementById('form-guest-name').value.trim(),
        guestMobile: document.getElementById('form-guest-mobile').value.trim(),
        room: document.getElementById('form-room').value,
        agent: document.getElementById('form-agent').value,
        checkIn: document.getElementById('form-checkin').value,
        checkOut: document.getElementById('form-checkout').value,
        totalAmount: parseFloat(document.getElementById('form-total-amount').value) || 0,
        advanceAmount: parseFloat(document.getElementById('form-advance-amount').value) || 0,
        dueAmount: parseFloat(document.getElementById('form-due-amount').value) || 0,
        status: document.getElementById('form-status').value,
        remarks: document.getElementById('form-remarks').value.trim()
      };

      if (editSl) {
        const index = appData.bookings.findIndex(b => b.sl == editSl);
        if (index !== -1) appData.bookings[index] = record;
      } else {
        appData.bookings.push(record);
      }

      saveDataToStorage();
      closeBookingModal();
      renderBookingTable();
      updateDashboardMetrics();
      showToast('Booking record saved successfully!');
    }

    function deleteBooking(sl) {
      if (confirm('Are you sure you want to delete this booking record?')) {
        appData.bookings = appData.bookings.filter(b => b.sl != sl);
        saveDataToStorage();
        renderBookingTable();
        updateDashboardMetrics();
        showToast('Record deleted.');
      }
    }

    function renderBookingTable() {
      const tbody = document.getElementById('booking-table-body');
      const searchQuery = document.getElementById('booking-search').value.toLowerCase();
      const yearFilter = document.getElementById('booking-year-filter').value;
      const statusFilter = document.getElementById('booking-status-filter').value;

      let filtered = appData.bookings.filter(b => {
        const matchesSearch = b.guestName.toLowerCase().includes(searchQuery) ||
                              b.guestMobile.includes(searchQuery) ||
                              b.room.toLowerCase().includes(searchQuery) ||
                              b.bookingId.toLowerCase().includes(searchQuery);

        const bYear = new Date(b.checkIn).getFullYear().toString();
        const matchesYear = yearFilter === 'ALL' || bYear === yearFilter;
        const matchesStatus = statusFilter === 'ALL' || b.status === statusFilter;

        return matchesSearch && matchesYear && matchesStatus;
      });

      document.getElementById('booking-count').innerText = filtered.length;

      if (filtered.length === 0) {
        tbody.innerHTML = `<tr><td colspan="13" class="text-center py-8 text-slate-400">No booking records found.</td></tr>`;
        return;
      }

      const statusBadges = {
        Live: 'bg-amber-100 text-amber-700 border-amber-200',
        Upcoming: 'bg-blue-100 text-blue-700 border-blue-200',
        Closed: 'bg-emerald-100 text-emerald-700 border-emerald-200',
        Inactive: 'bg-rose-100 text-rose-700 border-rose-200'
      };

      tbody.innerHTML = filtered.map((b, idx) => `
        <tr class="hover:bg-slate-50 transition">
          <td class="py-2.5 px-3 font-semibold text-slate-400">${idx + 1}</td>
          <td class="py-2.5 px-3 font-bold text-blue-600">${b.bookingId}</td>
          <td class="py-2.5 px-3 font-bold text-slate-900">${b.guestName}</td>
          <td class="py-2.5 px-3">${b.guestMobile}</td>
          <td class="py-2.5 px-3 font-medium">${b.room}</td>
          <td class="py-2.5 px-3">${b.checkIn}</td>
          <td class="py-2.5 px-3">${b.checkOut}</td>
          <td class="py-2.5 px-3">${b.agent}</td>
          <td class="py-2.5 px-3 text-right font-bold text-slate-900">₹${b.totalAmount.toLocaleString('en-IN')}</td>
          <td class="py-2.5 px-3 text-right font-bold text-emerald-600">₹${b.advanceAmount.toLocaleString('en-IN')}</td>
          <td class="py-2.5 px-3 text-right font-bold text-rose-600">₹${b.dueAmount.toLocaleString('en-IN')}</td>
          <td class="py-2.5 px-3 text-center">
            <span class="px-2 py-0.5 rounded-full text-[10px] font-bold border ${statusBadges[b.status] || 'bg-slate-100'}">${b.status}</span>
          </td>
          <td class="py-2.5 px-3 text-center space-x-1">
            <button onclick="openBookingModal(${b.sl})" title="Edit" class="p-1 text-blue-600 hover:bg-blue-50 rounded-lg transition"><i class="fa-solid fa-pen-to-square"></i></button>
            <button onclick="printInvoice(${b.sl})" title="Print Invoice" class="p-1 text-emerald-600 hover:bg-emerald-50 rounded-lg transition"><i class="fa-solid fa-print"></i></button>
            <button onclick="deleteBooking(${b.sl})" title="Delete" class="p-1 text-rose-600 hover:bg-rose-50 rounded-lg transition"><i class="fa-solid fa-trash-can"></i></button>
          </td>
        </tr>
      `).join('');
    }

    function filterBookingTable() {
      renderBookingTable();
    }

    // --- CALENDAR VIEW ---
    function renderCalendarView() {
      const month = parseInt(document.getElementById('cal-month-select').value);
      const year = parseInt(document.getElementById('cal-year-select').value);
      const grid = document.getElementById('calendar-grid');

      grid.innerHTML = '';

      const firstDay = new Date(year, month, 1).getDay();
      const daysInMonth = new Date(year, month + 1, 0).getDate();

      for (let i = 0; i < firstDay; i++) {
        grid.innerHTML += `<div class="h-20 bg-slate-50/50 rounded-2xl border border-dashed border-slate-100"></div>`;
      }

      for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        
        // Check for bookings on this date
        const dayBookings = appData.bookings.filter(b => {
          return dateStr >= b.checkIn && dateStr <= b.checkOut;
        });

        const hasBooking = dayBookings.length > 0;

        grid.innerHTML += `
          <div onclick="showDateCommentBox(event, '${dateStr}')" class="h-20 bg-white border border-slate-200/80 rounded-2xl p-2 relative hover:border-blue-500 transition cursor-pointer flex flex-col justify-between group shadow-xs">
            <div class="flex justify-between items-center">
              <span class="font-bold text-slate-700 text-xs group-hover:text-blue-600">${day}</span>
              ${hasBooking ? `<span class="w-0 h-0 border-t-[8px] border-t-rose-500 border-l-[8px] border-l-transparent absolute top-1 right-1"></span>` : ''}
            </div>
            ${hasBooking ? `
              <div class="bg-blue-50 text-blue-700 text-[9px] font-bold px-1.5 py-0.5 rounded-lg border border-blue-100 truncate">
                ${dayBookings.length} Booking(s)
              </div>
            ` : '<span class="text-[9px] text-slate-300">Vacant</span>'}
          </div>
        `;
      }
    }

    function showDateCommentBox(event, dateStr) {
      event.stopPropagation();
      const box = document.getElementById('excel-comment-box');
      const header = document.getElementById('comm-date-header');
      const list = document.getElementById('comm-booking-list');

      header.innerText = `Overview: ${dateStr}`;

      const bookings = appData.bookings.filter(b => dateStr >= b.checkIn && dateStr <= b.checkOut);

      if (bookings.length === 0) {
        list.innerHTML = `<p class="text-slate-400 text-[10px] italic">No active bookings for this date.</p>`;
      } else {
        list.innerHTML = bookings.map(b => `
          <div class="bg-slate-800 p-2 rounded-xl border border-slate-700 text-[10px] space-y-0.5">
            <p class="font-bold text-amber-300">${b.guestName} (${b.room})</p>
            <p class="text-slate-300">Status: <strong class="text-white">${b.status}</strong></p>
            <p class="text-slate-400">In: ${b.checkIn} | Out: ${b.checkOut}</p>
          </div>
        `).join('');
      }

      box.style.top = `${event.clientY + window.scrollY + 10}px`;
      box.style.left = `${Math.min(event.clientX, window.innerWidth - 280)}px`;
      box.classList.remove('hidden');
    }

    function closeCommentBox() {
      document.getElementById('excel-comment-box').classList.add('hidden');
    }

    function jumpToTodayCalendar() {
      const today = new Date();
      const y = today.getFullYear();
      const m = today.getMonth();

      document.getElementById('cal-year-select').value = (y >= 2026 && y <= 2085) ? y : 2026;
      document.getElementById('cal-month-select').value = m;
      renderCalendarView();
    }

    // --- MASTER CONTROLS & AUTH ---
    function openMasterAuthModal() {
      document.getElementById('master-auth-modal').classList.remove('hidden');
    }

    function closeMasterAuthModal() {
      document.getElementById('master-auth-modal').classList.add('hidden');
      document.getElementById('master-auth-error').classList.add('hidden');
      document.getElementById('master-password-input').value = '';
    }

    function handleMasterAuth(e) {
      e.preventDefault();
      const pass = document.getElementById('master-password-input').value;
      if (pass === 'admin123') { // Replace with desired master password
        isMasterUnlocked = true;
        closeMasterAuthModal();
        document.getElementById('master-locked-view').classList.add('hidden');
        document.getElementById('master-unlocked-view').classList.remove('hidden');
        renderMasterLists();
      } else {
        document.getElementById('master-auth-error').classList.remove('hidden');
      }
    }

    function lockMasterData() {
      isMasterUnlocked = false;
      document.getElementById('master-unlocked-view').classList.add('hidden');
      document.getElementById('master-locked-view').classList.remove('hidden');
    }

    function renderMasterLists() {
      const roomList = document.getElementById('master-room-list');
      const agentList = document.getElementById('master-agent-list');

      document.getElementById('master-room-count').innerText = `${appData.rooms.length} Rooms`;
      document.getElementById('master-agent-count').innerText = `${appData.agents.length} Agents`;

      roomList.innerHTML = appData.rooms.map((r, i) => `
        <li class="py-2 flex justify-between items-center">
          <span class="font-bold text-slate-800">${r}</span>
          <button onclick="removeMasterItem('room', ${i})" class="text-rose-600 hover:text-rose-800"><i class="fa-solid fa-trash-can"></i></button>
        </li>
      `).join('');

      agentList.innerHTML = appData.agents.map((a, i) => `
        <li class="py-2 flex justify-between items-center">
          <span class="font-bold text-slate-800">${a}</span>
          <button onclick="removeMasterItem('agent', ${i})" class="text-indigo-600 hover:text-indigo-800"><i class="fa-solid fa-trash-can"></i></button>
        </li>
      `).join('');
    }

    function handleMasterAddRoom(e) {
      e.preventDefault();
      const val = document.getElementById('master-room-input').value.trim();
      if (val && !appData.rooms.includes(val)) {
        appData.rooms.push(val);
        document.getElementById('master-room-input').value = '';
        saveDataToStorage();
        renderMasterLists();
      }
    }

    function handleMasterAddAgent(e) {
      e.preventDefault();
      const val = document.getElementById('master-agent-input').value.trim();
      if (val && !appData.agents.includes(val)) {
        appData.agents.push(val);
        document.getElementById('master-agent-input').value = '';
        saveDataToStorage();
        renderMasterLists();
      }
    }

    function removeMasterItem(type, index) {
      if (type === 'room') appData.rooms.splice(index, 1);
      if (type === 'agent') appData.agents.splice(index, 1);
      saveDataToStorage();
      renderMasterLists();
    }

    // --- HOVER SCROLL CONTROLS ---
    function scrollPageVertical(dir) {
      const distance = 400;
      window.scrollBy({ top: dir === 'up' ? -distance : distance, behavior: 'smooth' });
    }

    function scrollPageHorizontal(dir) {
      const distance = 300;
      window.scrollBy({ left: dir === 'left' ? -distance : distance, behavior: 'smooth' });
    }

    // --- DATA PERSISTENCE & TOAST ---
    function saveDataToStorage() {
      localStorage.setItem('hotel_app_data', JSON.stringify(appData));
    }

    function loadDataFromStorage() {
      const saved = localStorage.getItem('hotel_app_data');
      if (saved) {
        try {
          appData = JSON.parse(saved);
        } catch (e) {
          console.error("Failed to parse local storage data", e);
        }
      }
      updateDashboardMetrics();
    }

    function showToast(msg) {
      const toast = document.getElementById('toast');
      document.getElementById('toast-message').innerText = msg;
      toast.classList.remove('hidden');
      setTimeout(() => toast.classList.add('hidden'), 3000);
    }

    function saveChanges() {
      saveDataToStorage();
      showToast('All changes saved locally!');
    }

    // --- PRINT INVOICE ---
    function printInvoice(sl) {
      const item = appData.bookings.find(b => b.sl == sl);
      if (!item) return;

      document.getElementById('inv-id').innerText = item.bookingId;
      document.getElementById('inv-date').innerText = `Date: ${new Date().toLocaleDateString()}`;
      document.getElementById('inv-guest-name').innerText = item.guestName;
      document.getElementById('inv-guest-mobile').innerText = `Mobile: ${item.guestMobile}`;
      document.getElementById('inv-room').innerText = item.room;
      document.getElementById('inv-agent').innerText = item.agent;
      document.getElementById('inv-checkin').innerText = item.checkIn;
      document.getElementById('inv-checkout').innerText = item.checkOut;
      document.getElementById('inv-total').innerText = `₹${item.totalAmount.toLocaleString('en-IN')}`;
      document.getElementById('inv-advance').innerText = `₹${item.advanceAmount.toLocaleString('en-IN')}`;
      document.getElementById('inv-due').innerText = `₹${item.dueAmount.toLocaleString('en-IN')}`;

      const printArea = document.getElementById('printable-invoice');
      printArea.classList.remove('hidden');
      window.print();
      printArea.classList.add('hidden');
    }

    // --- INACTIVITY & LOGOUT LOGIC ---
    function setupInactivityListeners() {
      ['mousemove', 'keydown', 'click', 'scroll'].forEach(evt => {
        window.addEventListener(evt, resetInactivityTimer);
      });
    }

    function resetInactivityTimer() {
      clearTimeout(inactivityTimer);
      clearInterval(countdownInterval);
      document.getElementById('logout-warning-modal').classList.add('hidden');
      secondsRemaining = 60;

      // Warn after 14 minutes of inactivity (15-min auto logout)
      inactivityTimer = setTimeout(() => {
        document.getElementById('logout-warning-modal').classList.remove('hidden');
        countdownInterval = setInterval(() => {
          secondsRemaining--;
          document.getElementById('logout-countdown-seconds').innerText = secondsRemaining;
          if (secondsRemaining <= 0) {
            clearInterval(countdownInterval);
            processLogoutWithSave();
          }
        }, 1000);
      }, 14 * 60 * 1000);
    }

    function logoutUser() {
      document.getElementById('logout-confirm-modal').classList.remove('hidden');
    }

    function cancelLogout() {
      document.getElementById('logout-confirm-modal').classList.add('hidden');
    }

    function processLogoutWithSave() {
      saveDataToStorage();
      document.getElementById('logout-confirm-modal').classList.add('hidden');
      document.getElementById('saving-lock-modal').classList.remove('hidden');

      setTimeout(() => {
        document.getElementById('saving-lock-modal').classList.add('hidden');
        document.getElementById('login-overlay').classList.remove('hidden');
      }, 1500);
    }

    function handleLogin(e) {
      e.preventDefault();
      document.getElementById('login-overlay').classList.add('hidden');
      document.getElementById('login-alert-modal').classList.remove('hidden');
    }

    function closeLoginAlertModal() {
      document.getElementById('login-alert-modal').classList.add('hidden');
    }

    // --- EXPORT TO EXCEL ---
    function openExportModal() {
      document.getElementById('export-modal').classList.remove('hidden');
    }

    function closeExportModal() {
      document.getElementById('export-modal').classList.add('hidden');
    }

    function processExport() {
      const start = document.getElementById('export-start-date').value;
      const end = document.getElementById('export-end-date').value;

      let exportData = appData.bookings;

      if (start && end) {
        exportData = exportData.filter(b => b.checkIn >= start && b.checkOut <= end);
      }

      const worksheet = XLSX.utils.json_to_sheet(exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Bookings");
      XLSX.writeFile(workbook, `Booking_Report_${start || 'All'}_to_${end || 'All'}.xlsx`);

      closeExportModal();
      showToast('Excel report generated successfully!');
    }

    // --- WIPE DATA LOGIC ---
    function requestDataWipe() {
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

    function executeGoogleSheetWipe() {
      appData.bookings = [];
      saveDataToStorage();
      closeWipeModals();
      updateDashboardMetrics();
      renderBookingTable();
      showToast('All system records permanently wiped.');
    }
  </script>
</body>
</html>
