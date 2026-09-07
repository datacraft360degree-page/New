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
    /* Floating Balloons & Ribbons Keyframe Animations */
    @keyframes floatUp {
      0% { transform: translateY(100vh) rotate(0deg); opacity: 0; }
      20% { opacity: 0.8; }
      100% { transform: translateY(-20vh) rotate(360deg); opacity: 0; }
    }

    @keyframes ribbonSway {
      0%, 100% { transform: translateY(0) rotate(-5deg); }
      50% { transform: translateY(-15px) rotate(5deg); }
    }

    .balloon {
      position: absolute;
      bottom: -100px;
      font-size: 2rem;
      animation: floatUp 8s linear infinite;
      z-index: 10;
      pointer-events: none;
    }

    .ribbon {
      position: absolute;
      font-size: 1.8rem;
      animation: ribbonSway 4s ease-in-out infinite;
      z-index: 10;
      pointer-events: none;
    }
  </style>
</head>
<body class="text-slate-800 font-sans min-h-screen flex flex-col relative antialiased text-xs" onclick="closeCommentBox()">

<!-- BIRTHDAY HURRAY ANIMATION MODAL (TRIGGERS ON 20TH AUGUST) -->
  <div id="birthday-hurray-modal" class="hidden fixed inset-0 z-[110] bg-slate-900/40 backdrop-blur-md flex items-center justify-center p-4 no-print overflow-hidden">

    <!-- Modal Card (Soft Light Palette) -->
    <div class="bg-gradient-to-br from-amber-100 via-pink-100 to-sky-100 rounded-3xl p-1 shadow-2xl max-w-md w-full relative z-20 border border-white/60">
      <div class="bg-white/90 backdrop-blur-md rounded-[22px] p-6 text-center space-y-4 relative overflow-hidden border border-slate-100">
        
        <div class="absolute -top-10 -right-10 w-32 h-32 bg-amber-200/50 rounded-full blur-2xl"></div>
        <div class="absolute -bottom-10 -left-10 w-32 h-32 bg-pink-200/50 rounded-full blur-2xl"></div>
        
        <div class="text-6xl animate-bounce relative z-10">
          🎉🥳🎂
        </div>
        
        <div class="space-y-2 relative z-10">
          <h2 class="text-xl font-black text-slate-800 tracking-tight">Happy Birthday Mr. Aniruddha Sir 🥳🎂</h2>
          <p class="text-sm font-semibold text-slate-600 bg-sky-50/80 border border-sky-100 py-3 px-4 rounded-2xl shadow-inner leading-relaxed">
            "Hi Aniruddha Sir, it's your birthday. Keep smiling always and stay healthy! This is a message from Mr. Kapil: wishing you a very Happy Birthday." 😊
          </p>
        </div>

        <button onclick="closeBirthdayModal()" class="w-full bg-gradient-to-r from-sky-400 via-rose-400 to-amber-400 hover:opacity-95 text-white font-black py-3 rounded-2xl shadow-md transition text-xs relative z-10 cursor-pointer active:scale-95 tracking-wide">
          Hurray! Let's Celebrate 🚀
        </button>
      </div>
    </div>
  </div>

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

  <!-- LOGIN ALERT MESSAGE MODAL (POPUP ON SUCCESSFUL LOGIN) -->
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
         <p class="flex items-start gap-2"><i class="fa-solid fa-check text-blue-500 mt-0.5"></i> <span>1. Google Chrome/Microsoft edge is best view browser for this Portal.</span></p>
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

  <!-- SAVING LOCK MODAL (SAND TIMER) -->
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
        <!-- Save Button -->
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
          <h2 class="text-base font-bold tracking-tight">Hi, Welcome to dashboard 🏠</h2>
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

      <!-- One UI Rounded Cards -->
    <!-- 6-Box Grid Layout (Line 1: 3 Boxes | Line 2: 3 Boxes) -->
      <div class="grid grid-cols-1 md:grid-cols-3 gap-3">

        <!-- BOX 1: TOTAL BOOKINGS (COUNT PIE CHART) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">Total Bookings</p>
              <p id="dash-total-bookings" class="text-lg font-black text-slate-900 leading-tight">0</p>
            </div>
            <div class="p-2.5 bg-amber-50 text-amber-600 rounded-2xl"><i class="fa-solid fa-bookmark text-sm"></i></div>
          </div>
          <div class="flex items-center space-x-3">
            <div id="pie-bookings" class="w-16 h-16 rounded-full shrink-0 shadow-inner border border-slate-100" style="background: conic-gradient(#E2E8F0 0% 100%);"></div>
            <div class="text-[11px] space-y-1 font-semibold text-slate-700 w-full">
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>Live:</span> <strong id="dash-count-live" class="text-amber-600 font-bold">0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-blue-500 inline-block"></span>Upcoming:</span> <strong id="dash-count-upcoming" class="text-blue-600 font-bold">0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>Closed:</span> <strong id="dash-count-closed" class="text-emerald-600 font-bold">0</strong></div>
            </div>
          </div>
        </div>

        <!-- BOX 2: BOOKING AMOUNT (AMOUNT PIE CHART) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">Booking Amount</p>
              <p id="dash-total-amount" class="text-lg font-black text-slate-900 leading-tight">₹0</p>
            </div>
            <div class="p-2.5 bg-blue-50 text-blue-600 rounded-2xl"><i class="fa-solid fa-receipt text-sm"></i></div>
          </div>
          <div class="flex items-center space-x-3">
            <div id="pie-amount" class="w-16 h-16 rounded-full shrink-0 shadow-inner border border-slate-100" style="background: conic-gradient(#E2E8F0 0% 100%);"></div>
            <div class="text-[11px] space-y-1 font-semibold text-slate-700 w-full">
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>Live:</span> <strong id="dash-amount-live" class="text-amber-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-blue-500 inline-block"></span>Upcoming:</span> <strong id="dash-amount-upcoming" class="text-blue-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>Closed:</span> <strong id="dash-amount-closed" class="text-emerald-600 font-bold">₹0</strong></div>
            </div>
          </div>
        </div>

        <!-- BOX 3: AMOUNT RECEIVED (ADVANCE PIE CHART) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">Amount Received</p>
              <p id="dash-advanced" class="text-lg font-black text-emerald-600 leading-tight">₹0</p>
            </div>
            <div class="p-2.5 bg-emerald-50 text-emerald-600 rounded-2xl"><i class="fa-solid fa-wallet text-sm"></i></div>
          </div>
          <div class="flex items-center space-x-3">
            <div id="pie-received" class="w-16 h-16 rounded-full shrink-0 shadow-inner border border-slate-100" style="background: conic-gradient(#E2E8F0 0% 100%);"></div>
            <div class="text-[11px] space-y-1 font-semibold text-slate-700 w-full">
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>Live:</span> <strong id="dash-recv-live" class="text-amber-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-blue-500 inline-block"></span>Upcoming:</span> <strong id="dash-recv-upcoming" class="text-blue-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>Closed:</span> <strong id="dash-recv-closed" class="text-emerald-600 font-bold">₹0</strong></div>
            </div>
          </div>
        </div>

        <!-- BOX 4: TOTAL DUE AMOUNT (DUE PIE CHART) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">Total Due Amount</p>
              <p id="dash-due" class="text-lg font-black text-rose-600 leading-tight">₹0</p>
            </div>
            <div class="p-2.5 bg-rose-50 text-rose-600 rounded-2xl"><i class="fa-solid fa-hand-holding-dollar text-sm"></i></div>
          </div>
          <div class="flex items-center space-x-3">
            <div id="pie-due" class="w-16 h-16 rounded-full shrink-0 shadow-inner border border-slate-100" style="background: conic-gradient(#E2E8F0 0% 100%);"></div>
            <div class="text-[11px] space-y-1 font-semibold text-slate-700 w-full">
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-amber-500 inline-block"></span>Live Due:</span> <strong id="dash-due-live" class="text-amber-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-blue-500 inline-block"></span>Upcoming Due:</span> <strong id="dash-due-upcoming" class="text-blue-600 font-bold">₹0</strong></div>
              <div class="flex items-center justify-between"><span class="flex items-center gap-1.5"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500 inline-block"></span>Closed Due:</span> <strong id="dash-due-closed" class="text-emerald-600 font-bold">₹0</strong></div>
            </div>
          </div>
        </div>

        <!-- BOX 5: INACTIVE BOOKINGS SUMMARY (RED THEME) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-3">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <div>
              <p class="text-[10px] uppercase font-bold text-red-500 tracking-wider break-words">Inactive Bookings</p>
              <p id="dash-inactive-count" class="text-lg font-black text-red-600 leading-tight">0 Bookings</p>
            </div>
            <div class="p-2.5 bg-red-50 text-red-600 rounded-2xl"><i class="fa-solid fa-ban text-sm"></i></div>
          </div>
          <div class="bg-red-50/50 p-3 rounded-2xl border border-red-100 text-[11px] space-y-1">
            <p class="text-slate-500 font-medium">Inactive Total Booking Amount:</p>
            <p id="dash-inactive-amount" class="text-lg font-black text-red-600">₹0</p>
          </div>
        </div>

        <!-- BOX 6: ALL LIVE BOOKING DETAILS (YELLOW/AMBER THEME) -->
        <div class="bg-white p-4 rounded-3xl shadow-sm border border-slate-200/60 flex flex-col justify-between space-y-2">
          <div class="flex justify-between items-center border-b border-slate-100 pb-2">
            <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider break-words">All Live Bookings Details</p>
            <span class="px-2 py-0.5 text-[9px] font-bold bg-amber-50 text-amber-600 rounded-full border border-amber-200 flex items-center gap-1">
              <span class="w-1.5 h-1.5 bg-amber-500 rounded-full animate-ping"></span> Live Today
            </span>
          </div>
          <div id="dash-today-live-container" class="overflow-x-auto max-h-32 overflow-y-auto">
            <p class="text-[11px] text-slate-400 italic py-2 text-center">No live bookings active today.</p>
          </div>
        </div>

      </div>

      <!-- Active years Directory Table Hidden -->
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
            <!-- Search by Date -->
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
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-lg w-full flex flex-col max-h-[85vh] overflow-hidden">
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
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-3xl w-full p-5 space-y-3 my-4 max-h-[90vh] overflow-y-auto">
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
            <!-- EDITABLE COUNTRY CODE & GUEST CONTACT NUMBER -->
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

       <!-- ROOM SELECTION & STAY DATES -->
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
    
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5">Agent Info</label>
      <select id="cust-agent" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-slate-700"></select>
    </div>

    <div>
      <label class="block font-semibold text-slate-600 mb-0.5">Main Guest Joined</label>
      <input type="number" id="cust-capacity" min="1" value="1" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-slate-700" />
    </div>

    <!-- EXTRA PERSON(S) COUNT FIELD -->
    <div>
      <label class="block font-semibold text-amber-700 mb-0.5 flex items-center gap-1">
        <i class="fa-solid fa-user-plus text-amber-600"></i> Add Extra Guest(s)
      </label>
      <input type="number" id="cust-extra-persons" min="0" value="0" placeholder="0" oninput="calculateModalBilling()" class="w-full bg-amber-50 border border-amber-300 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-amber-500 font-bold text-amber-900" />
    </div>

    <!-- NEW FIELD: ADD EXTRA ROOM(S) -->
    <div>
      <label class="block font-semibold text-blue-700 mb-0.5 flex items-center gap-1">
        <i class="fa-solid fa-door-open text-blue-600"></i> Add Extra Room No(s)
      </label>
      <input type="number" id="cust-extra-rooms" min="0" value="0" placeholder="0" oninput="calculateModalBilling()" class="w-full bg-blue-50 border border-blue-300 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-bold text-blue-900" />
    </div>

    <!-- NEW FIELD: NAME -->
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5">Extra Guest Name</label>
      <input type="text" id="cust-room-name" placeholder="Guest Name" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-medium" />
    </div>

    <!-- ADDITIONAL PERSON CUSTOM CHECK-IN & CHECK-OUT WINDOW -->
    <div id="sec-extra-person-time-wrapper" class="sm:col-span-4 hidden bg-amber-50/70 p-2.5 rounded-2xl border border-amber-200/80 space-y-2">
      <label class="block font-bold text-amber-900 mb-1 flex items-center gap-1">
        <i class="fa-solid fa-clock-rotate-left text-amber-600"></i> Additional Person Stay Window (Custom Dates Required)
      </label>
      <div class="grid grid-cols-1 sm:grid-cols-2 gap-2">
        <div>
          <label class="block font-semibold text-amber-800 text-[10px] mb-0.5">Extra Guest Check-In</label>
          <div class="flex gap-1">
            <input type="date" id="cust-extra-person-date" onchange="handleExtraPersonDatesChange()" class="w-2/3 bg-white border border-amber-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
            <input type="time" id="cust-extra-person-time" onchange="handleExtraPersonDatesChange()" class="w-1/3 bg-white border border-amber-200 rounded-xl px-1.5 py-1.5 focus:outline-none focus:border-amber-500 font-semibold text-amber-900" />
          </div>
        </div>
        <div>
          <label class="block font-semibold text-amber-800 text-[10px] mb-0.5">Extra Guest Check-Out</label>
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

      <!-- BILLING SUMMARY -->
<div id="sec-billing-summary" class="bg-blue-50/40 p-3 rounded-2xl border border-blue-100 space-y-2.5 transition-all">
  <h4 class="text-[9px] font-bold uppercase tracking-wider text-blue-700 flex items-center gap-1.5">
    <i class="fa-solid fa-calculator text-blue-600"></i> Billing Summary
  </h4>
  
  <!-- Row 1: 4 Columns (Extra items + Cab) -->
  <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-2 pb-2 border-b border-blue-200">
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5">Extra Guest Total Days</label>
      <input type="number" id="cust-extra-person-days" readonly="" class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
    </div>
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5">Extra Guest Price/Day</label>
      <input type="number" id="cust-extra-person-price" value="0" oninput="calculateModalBilling()" class="w-full bg-white font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500" />
    </div>
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5">Extra Guest Total Rate (₹)</label>
      <input type="number" id="cust-extra-total" readonly="" class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
    </div>
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5">Cab Fare Total (₹)</label>
      <input type="number" id="cust-cab-total" readonly="" class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
    </div>
  </div>

  <!-- Row 2: 4 Columns (Food + Main Guest items) -->
  <div class="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-2 pb-2 border-b border-blue-200">
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5">Extra Food/Drink Total (₹)</label>
      <input type="number" id="cust-food-total" readonly="" class="w-full bg-slate-200/60 font-bold text-amber-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
    </div>
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5">Main Guest Total Days</label>
      <input type="number" id="cust-days" readonly="" class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
    </div>
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5 leading-tight">Main Guest Price/Day (₹)</label>
      <input type="number" id="cust-price" value="1200" oninput="calculateModalBilling()" class="w-full bg-white font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 focus:outline-none focus:border-blue-500" />
    </div>
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5 leading-tight">Main Guest Total Rate (₹)</label>
      <input type="number" id="cust-main-person-rate" readonly="" class="w-full bg-slate-200/60 font-bold text-slate-700 border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" value="0" />
    </div>
  </div>

  <!-- Row 3: 4 Columns (Totals & Payments) -->
  <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5 leading-tight">Grand Total (₹)</label>
      <input type="number" id="cust-total" readonly="" class="w-full bg-slate-200/60 text-blue-700 font-bold border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
    </div>
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5 leading-tight">Initial Advance (₹)</label>
      <input type="number" id="cust-advance" value="0" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-200 rounded-xl px-2.5 py-1.5 focus:outline-none focus:border-blue-500 font-semibold text-emerald-600" />
    </div>
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5 leading-tight">Due (₹)</label>
      <input type="number" id="cust-due" readonly="" class="w-full bg-slate-200/60 text-rose-700 font-bold border border-slate-200 rounded-xl px-2 py-1.5 cursor-not-allowed" />
    </div>
    <div>
      <label class="block font-semibold text-slate-600 mb-0.5 leading-tight text-[10px] text-emerald-700">Clear Bill (₹)</label>
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
    <div class="bg-white rounded-3xl shadow-2xl border border-slate-100 max-w-xl w-full p-5 sm:p-6 space-y-4 relative my-auto max-h-[92vh] overflow-y-auto" id="printable-invoice">
      
      <!-- Read-Only Notice Bar -->
      <div id="inv-readonly-notice" class="hidden bg-slate-900 text-amber-300 text-[10px] font-bold px-3.5 py-2 rounded-2xl flex items-center justify-between border border-slate-800 no-print">
        <span class="flex items-center gap-1.5">
          <i class="fa-solid fa-lock text-amber-400"></i> Read-Only View Mode (Editing Disabled)
        </span>
        <span class="text-[9px] text-slate-400 font-normal">System Protected</span>
      </div>

      <div class="flex justify-between items-start border-b border-slate-200 pb-3">
        <div>
          <h2 class="text-base sm:text-lg font-black text-blue-600 uppercase tracking-wide">Sanoum Pema Homestay-by Anaristays</h2>
          <p class="text-[10px] text-slate-500 mt-0.5">Sittong, Village in West Bengal</p>
          <p class="text-[10px] text-slate-500">Phone: +91 9804396541 | Email: demo@gmail.com</p>
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

<script>
// Function to check and display the birthday animation on September 27th
function checkBirthdayTrigger() {
      const now = new Date();
      const currentMonth = now.getMonth() + 1; // Month 9 is September
      const currentDay = now.getDate();        // 27th day
      if (currentMonth === 9 && currentDay === 27) {
        const lastShownYear = sessionStorage.getItem('birthday_shown_year');
        const currentYear = now.getFullYear();
        if (lastShownYear !== String(currentYear)) {
          const bModal = document.getElementById('birthday-hurray-modal');
          if (bModal) {
            bModal.classList.remove('hidden');
            sessionStorage.setItem('birthday_shown_year', String(currentYear));
          }
        }
      }
    }
    function closeBirthdayModal() {
      const bModal = document.getElementById('birthday-hurray-modal');
      if (bModal) {
        bModal.classList.add('hidden');
      }
    }
    // System Constants & Robust Helpers
    const MAX_SHEET_ROWS = 10000000;
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
    
    // Extract IST local parts robustly for input fields ensuring UTC+05:30 offset
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
    // Convert a Date object to local ISO string adhering to IST (UTC+05:30)
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
    // Helper logic to format dates strictly into 24-hour style "dd/mm/yy hh:mm" for Excel exports (Enforcing IST)
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
    
    // Robust parsing for generic JSON array fields (Food Orders, Cab Trips)
    function parseJSONField(fieldData) {
      if (!fieldData) return [];
      if (Array.isArray(fieldData)) return fieldData;
      if (typeof fieldData === 'string' && fieldData.length > 2) {
        try { return JSON.parse(fieldData); } catch (e) {}
      }
      return [];
    }
    
    function checkSheetRowLimits() {
      const currentRowCount = state.bookings.length; 
      
      if (currentRowCount >= (MAX_SHEET_ROWS - 20)) {
        alert("CRITICAL NOTIFICATION: Google Sheet has reached its row limit (less than 20 rows left). Saving has been stopped. Please generate a new sheet and link it to continue.");
        return false;
      } else if (currentRowCount >= (MAX_SHEET_ROWS - 100)) {
        alert("WARNING: Google Sheet is approaching its 10 million row limit (less than 100 rows left). Please prepare a new sheet soon.");
      }
      return true;
    }
    // Double Layer Data Wipe
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
        try {
            JSON.parse(textResult);
        } catch(e) {
            console.error("Wipe format warning:", textResult);
        }
        // Reset local state
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
    const GAS_API_URL = "https://script.google.com/macros/s/AKfycbz6rME_OuYHucGBPCfCrV7EYjuE5YF0eqSeuqBjm42-HPXUYJzUSBu0mov9jCdM7zx5Ng/exec"; 
    
    const ONE_HOUR_MS = 1 * 60 * 60 * 1000;
    let activeModalBooking = null;
    window.addEventListener('beforeunload', function (e) {
      if (isLoggedIn) {
        e.preventDefault();
        e.returnValue = 'Please click "Save Changes" button to save the history.'; 
        return e.returnValue;
      }
    });
    function formatTitleCase(text) {
      if (!text) return '';
      return String(text).replace(/\w\S*/g, function(txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
      });
    }
    function handleStateChange(stateValue) {
      if (stateValue && stateValue.trim().toLowerCase() === 'west bengal') {
        const countryInput = document.getElementById('cust-country');
        if (countryInput) countryInput.value = 'India';
      }
    }
    function getEffectiveCheckoutTime(b) {
      if (!b) return 0;
      if (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) {
        return parseDateMs(b.extendedCheckOut);
      }
      return parseDateMs(b.checkOut);
    }
    function getModalFoodWindow() {
      const inDate = document.getElementById('cust-checkin-date')?.value;
      const inTime = document.getElementById('cust-checkin-time')?.value || '12:00';
      const hasExtCheckout = document.getElementById('cust-has-extended-checkout')?.checked;
      
      let outDate = document.getElementById('cust-checkout-date')?.value;
      let outTime = document.getElementById('cust-checkout-time')?.value || '11:00';
      if (hasExtCheckout) {
        const extDate = document.getElementById('cust-ext-checkout-date')?.value;
        const extTime = document.getElementById('cust-ext-checkout-time')?.value;
        if (extDate) outDate = extDate;
        if (extTime) outTime = extTime;
      }
      if (!inDate || !outDate) return null;
      const checkInDt = new Date(`${inDate}T${inTime}:00+05:30`);
      const checkOutDt = new Date(`${outDate}T${outTime}:00+05:30`);
      if (isNaN(checkInDt.getTime()) || isNaN(checkOutDt.getTime())) return null;
      const minFoodDt = new Date(checkInDt.getTime() + 15 * 60 * 1000);  
      const maxFoodDt = new Date(checkOutDt.getTime() - 30 * 60 * 1000); 
      return { checkInDt, checkOutDt, minFoodDt, maxFoodDt };
    }
    function validateFoodRowDateTime(inputElem) {
      const row = inputElem.closest('.food-order-row');
      if (!row) return;
      const fDate = row.querySelector('.cust-food-date').value;
      const fTime = row.querySelector('.cust-food-time').value || '00:00';
      if (!fDate) return;
      const foodWin = getModalFoodWindow();
      if (!foodWin) return;
      const selectedDt = new Date(`${fDate}T${fTime}:00+05:30`);
      if (selectedDt < foodWin.minFoodDt || selectedDt > foodWin.maxFoodDt) {
        const minStr = formatDateTime(foodWin.minFoodDt);
        const maxStr = formatDateTime(foodWin.maxFoodDt);
        alert(`⚠️ Extra Food Order time must be after 15 mins of Check-In (${minStr}) and at least 30 mins before Check-Out (${maxStr})!`);
        
        const targetDt = selectedDt < foodWin.minFoodDt ? foodWin.minFoodDt : foodWin.maxFoodDt;
        const utcMs = targetDt.getTime();
        const istDate = new Date(utcMs + (330 * 60000));
        
        const yyyy = istDate.getUTCFullYear();
        const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(istDate.getUTCDate()).padStart(2, '0');
        const hh = String(istDate.getUTCHours()).padStart(2, '0');
        const min = String(istDate.getUTCMinutes()).padStart(2, '0');
        row.querySelector('.cust-food-date').value = `${yyyy}-${mm}-${dd}`;
        row.querySelector('.cust-food-time').value = `${hh}:${min}`;
      }
    }
    function handleIdProofUpload(e) {
      const fileInput = e.target;
      const file = fileInput.files[0];
      const statusText = document.getElementById('cust-id-file-status');
      const base64Input = document.getElementById('cust-id-file-base64');
      const fileNameInput = document.getElementById('cust-id-file-name');
      const removeBtn = document.getElementById('cust-id-file-remove');
      if (!file) return;
      if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith('.pdf')) {
        alert("⚠️ Invalid file format! Only PDF files are allowed.");
        fileInput.value = '';
        return;
      }
      const minSize = 10 * 1024;  
      const maxSize = 900 * 1024; 
      if (file.size < minSize || file.size > maxSize) {
        const fileSizeKB = (file.size / 1024).toFixed(1);
        alert(`⚠️ Invalid file size (${fileSizeKB} KB)!\n\nThe attached ID proof PDF must be between 10 KB and 900 KB.`);
        fileInput.value = '';
        return;
      }
      const reader = new FileReader();
      reader.onload = function(evt) {
        base64Input.value = evt.target.result;
        fileNameInput.value = file.name;
        statusText.innerHTML = `<span class="text-emerald-600 font-semibold"><i class="fa-solid fa-circle-check"></i> Attached: ${file.name} (${(file.size / 1024).toFixed(1)} KB)</span>`;
        removeBtn.classList.remove('hidden');
      };
      reader.readAsDataURL(file);
    }
    function removeAttachedIdProof() {
      document.getElementById('cust-id-file').value = '';
      document.getElementById('cust-id-file-base64').value = '';
      document.getElementById('cust-id-file-name').value = '';
      document.getElementById('cust-id-file-status').innerText = 'No PDF document attached.';
      document.getElementById('cust-id-file-remove').classList.add('hidden');
    }
    function openPdfAttachment(base64Data) {
      if (!base64Data) {
        alert("No ID Proof attached!");
        return;
      }
      const win = window.open();
      if (win) {
        win.document.write(`<iframe src="${base64Data}" frameborder="0" style="border:0; top:0px; left:0px; bottom:0px; right:0px; width:100%; height:100%;" allowfullscreen></iframe>`);
      } else {
        alert("Please allow popups to view attached PDF document.");
      }
    }
    function toggleExtendedCheckoutFields(checked) {
      const container = document.getElementById('extended-checkout-container');
      if (!container) return;
      const normalOutDate = document.getElementById('cust-checkout-date').value;
      const normalOutTime = document.getElementById('cust-checkout-time').value;
      const extOutDate = document.getElementById('cust-ext-checkout-date');
      const extOutTime = document.getElementById('cust-ext-checkout-time');
      if (checked) {
        container.classList.remove('hidden');
        if (extOutDate) {
          extOutDate.min = normalOutDate;
          if (!extOutDate.value || extOutDate.value < normalOutDate) {
            extOutDate.value = normalOutDate;
          }
        }
        if (extOutTime && !extOutTime.value) extOutTime.value = normalOutTime || "12:00";
      } else {
        container.classList.add('hidden');
      }
      calculateModalBilling();
    }
    let isLoggedIn = false;
    let isMasterUnlocked = false; 
    let inactivityTimer = null;
    let warningTimer = null;
    let countdownInterval = null;
    const INACTIVITY_LIMIT_MS = 10 * 60 * 1000; 
    const WARNING_BUFFER_MS = 1 * 60 * 1000;   
    const DEFAULT_USER_ID = "Admin";
    const DEFAULT_PASSWORD = "Aadmin123";
    let pendingMasterDeleteType = null; 
    let pendingMasterDeleteTarget = null; 

    function openMasterDeleteModal(type, target) {
      pendingMasterDeleteType = type;
      pendingMasterDeleteTarget = target;
      const msgElem = document.getElementById('master-delete-modal-msg');
      if (type === 'booking') {
        const b = state.bookings.find(item => String(item.id) === String(target));
        const bCode = b ? b.bookingCode : 'this booking';
        msgElem.innerText = `Are you sure you want to permanently delete booking ${bCode} from the Master Tab? This action cannot be undone.`;
      } else if (type === 'room') {
        msgElem.innerText = `Are you sure you want to permanently delete this Room Capacity record? This action cannot be undone.`;
      } else if (type === 'agent') {
        msgElem.innerText = `Are you sure you want to permanently delete this Agent record? This action cannot be undone.`;
      }
      document.getElementById('master-delete-confirm-modal').classList.remove('hidden');
    }
    function closeMasterDeleteModal() {
      pendingMasterDeleteType = null;
      pendingMasterDeleteTarget = null;
      document.getElementById('master-delete-confirm-modal').classList.add('hidden');
    }
    function confirmMasterDeletion() {
      if (pendingMasterDeleteType === 'booking') {
        const id = pendingMasterDeleteTarget;
        const idx = state.bookings.findIndex(b => String(b.id) === String(id));
        if (idx !== -1) {
          state.bookings[idx].inactive = true;
        }
        searchMasterBookingById();
        renderBookingsTable();
        updateDashboardCards();
        renderCalendar(defaultAppYear);
        checkUpcomingCheckoutsWithDue();
        saveChanges(false, false);
      } else if (pendingMasterDeleteType === 'room') {
        const index = pendingMasterDeleteTarget;
        state.roomsCapacity.splice(index, 1);
        renderRoomCapacityTable();
        populateRoomDropdown();
        populateAgentDropdown();
        saveChanges(false, false);
      } else if (pendingMasterDeleteType === 'agent') {
        const index = pendingMasterDeleteTarget;
        state.masterAgents.splice(index, 1);
        renderMasterAgentTable();
        populateAgentDropdown();
        saveChanges(false, false);
      }
      closeMasterDeleteModal();
    }
    function closeLoginAlertModal() {
      document.getElementById('login-alert-modal').classList.add('hidden');
    }
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
        checkBirthdayTrigger();
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
    function formatDateTime(dtStr) {
      if (!dtStr) return '-';
      const d = new Date(typeof dtStr === 'string' ? dtStr.replace(' ', 'T') : dtStr);
      if (isNaN(d.getTime())) {
        const parts = String(dtStr).split('T');
        if (parts.length === 2) {
          const dateParts = parts[0].split('-');
          if (dateParts.length === 3) {
            return `${dateParts[2]}-${dateParts[1]}-${dateParts[0]} ${parts[1].substring(0, 5)}`;
          }
        }
        return String(dtStr).replace('T', ' ');
      }
      
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
    function isRoomInMaster(roomNo) {
      if (!state.roomsCapacity) return true;
      let rooms = getBookingRooms({ roomNo });
      return rooms.every(r => state.roomsCapacity.some(m => String(m.roomNo) === String(r)));
    }
    function openExportModal() {
      if (!state.bookings || state.bookings.length === 0) {
        alert("No booking records available to export!");
        return;
      }
      document.getElementById('export-start-date').value = '';
      document.getElementById('export-end-date').value = '';
      document.getElementById('export-modal').classList.remove('hidden');
    }
    function closeExportModal() {
      document.getElementById('export-modal').classList.add('hidden');
    }
    function validateExportDates() {
      const minDate = "2026-08-01";
      const maxDate = "2085-12-31";
      const startInput = document.getElementById('export-start-date');
      const endInput = document.getElementById('export-end-date');
      if (startInput.value && (startInput.value < minDate || startInput.value > maxDate)) {
        alert(`⚠️ Please select a Start Date between ${formatDate(minDate)} and ${formatDate(maxDate)}.`);
        startInput.value = "";
      }
      if (endInput.value && (endInput.value < minDate || endInput.value > maxDate)) {
        alert(`⚠️ Please select an End Date between ${formatDate(minDate)} and ${formatDate(maxDate)}.`);
        endInput.value = "";
      }
      if (startInput.value && endInput.value && startInput.value > endInput.value) {
        alert("⚠️ Start Date cannot be after End Date.");
        endInput.value = "";
      }
    }
    function processExport() {
      const startDateStr = document.getElementById('export-start-date').value;
      const endDateStr = document.getElementById('export-end-date').value;
      if (!startDateStr || !endDateStr) {
        alert("Please select both Start and End dates.");
        return;
      }
      exportToExcel(startDateStr, endDateStr);
    }
    function exportToExcel(startDateStr, endDateStr) {
      if (!state.bookings || state.bookings.length === 0) {
        alert("No booking records available to export!");
        return;
      }
      const filteredBookings = state.bookings.filter(b => {
        if (!b.checkIn) return false;
        const bIn = String(b.checkIn).replace(' ', 'T').split('T')[0];
        return (bIn >= startDateStr) && (bIn <= endDateStr);
      });
      if (filteredBookings.length === 0) {
        alert(`No booking records found with a Check-In date between ${formatDate(startDateStr)} and ${formatDate(endDateStr)}!`);
        return;
      }
      const now = new Date().getTime();
      const headers = [
        "Booking ID (System)", "Booking ID", "Invoice ID", "Booking Status", "Guest Name",
        "Address", "City", "State", "Country", "Pin/Zip Code", "ID Number",
        "Contact No", "Country Code", "Attached ID Proof(Yes/No)", "Room No(s)", "Agent Info",
        "Main Guest Joined(Show count)", "Extra Guest Joined(Show count)", "Extra Room No(s)",
        "Extra Guest Name", "Extra Guest Check-In", "Extra Guest Check-Out", "Check-In", "Check-Out",
        "Has Extended Check-Out(Yes/No)", "Extended Check-Out(Date & Time)", "Include Meals(Yes/No)",
        "Food Orders Details(JSON)", "Cab Trips Details(JSON)", "Extra Guest Total Days",
        "Extra Guest Price/Day (₹)", "Extra Guest Total Rate (₹)", "Cab Fare Total (₹)",
        "Extra Food/Drink Total (₹)", "Main Guest Total Days", "Main Guest Price/Day (₹)",
        "Main Guest Total Rate (₹)", "Grand Total (₹)", "Initial Advance (₹)", "Due (₹)", "Clear Bill (₹)"
      ];
      const rows = filteredBookings.map(b => {
        let bStatus = "Unknown";
        if (isInactiveBooking(b)) {
          bStatus = "Inactive";
        } else {
          const cIn = parseDateMs(b.checkIn);
          const cOut = getEffectiveCheckoutTime(b);
          if (now > cOut) bStatus = "Closed";
          else if (now >= cIn && now <= cOut) bStatus = "Live";
          else bStatus = "Upcoming";
        }
        const foodList = parseJSONField(b.foodOrders);
        const cabList = parseJSONField(b.cabTrips);
        
        const extraPrice = b.extraPersonPricePerDay !== undefined ? parseFloat(b.extraPersonPricePerDay) : parseFloat(b.perDayPrice || 0);
        const extraTotal = b.extraPersonTotalRate !== undefined ? parseFloat(b.extraPersonTotalRate) : ((parseFloat(b.extraPersons) || 0) * (parseFloat(b.extraPersonDays) || 0) * extraPrice);

        return [
          b.id || "",
          b.bookingCode || "",
          b.invoiceNo || "",
          bStatus,
          b.name || "",
          b.address || "",
          b.city || "",
          b.state || "",
          b.country || "",
          b.zipCode || "",
          b.idNo || "",
          b.contactNo || "",
          b.countryCode || "",
          b.idProofFileName ? "Yes" : "No",
          getBookingRooms(b).join(" | "),
          b.agentInfo || "",
          b.mainGuestCount || b.capacity || 1,
          b.extraPersons || 0,
          b.extraRoomNos || "",
          b.extraGuestName || "",
          format24hDate(b.extraPersonJoined),
          format24hDate(b.extraPersonOut),
          format24hDate(b.checkIn),
          format24hDate(b.checkOut),
          isTrue(b.hasExtendedCheckout) ? "Yes" : "No",
          format24hDate(b.extendedCheckOut),
          (b.includeMeals !== false && b.includeMeals !== 'false') ? "Yes" : "No",
          typeof b.foodOrders === 'string' ? b.foodOrders : JSON.stringify(foodList || []),
          typeof b.cabTrips === 'string' ? b.cabTrips : JSON.stringify(cabList || []),
          b.extraPersonDays || 0,
          extraPrice,
          extraTotal,
          b.cabFareTotal || 0,
          b.extraFoodTotal || 0,
          b.noOfDays || 0,
          b.perDayPrice || 0,
          b.mainGuestTotalRate || 0,
          b.totalAmount || 0,
          b.initialAdv || 0,
          b.totalDue || 0,
          b.clearedDue || 0
        ];
      });
      const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Bookings");
      XLSX.writeFile(workbook, `Booking_Report_${startDateStr}_to_${endDateStr}.xlsx`);
      closeExportModal();
    }
    function searchBookingByDate() {
      const dateVal = document.getElementById('booking-date-search').value;
      renderBookingsTable(dateVal);
    }
    function clearDateSearchBooking() {
      const input = document.getElementById('booking-date-search');
      if (input) input.value = "";
      renderBookingsTable();
    }
    function searchMasterBookingById() {
      const inputElem = document.getElementById('master-booking-search-input');
      if (!inputElem) return;
      const query = inputElem.value.trim().toUpperCase();
      const tbody = document.getElementById('master-delete-tbody');
      if (!tbody) return;
      tbody.innerHTML = '';
      if (!query) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-slate-400">Please type a Booking ID into the search field above to view and delete details.</td></tr>`;
        return;
      }
      const matchedBookings = state.bookings.filter(item => 
        !isInactiveBooking(item) && (item.bookingCode || '').toUpperCase().includes(query)
      );
      if (matchedBookings.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-rose-500 font-semibold">No active booking found matching "${query}".</td></tr>`;
        return;
      }
      matchedBookings.forEach(b => {
        const effectiveOutStr = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100";
        const roomsDisplay = getBookingRooms(b).join(', ');
        
        tr.innerHTML = `
          <td class="py-2.5 px-3 font-mono font-bold text-blue-600">${b.bookingCode}</td>
          <td class="py-2.5 px-3 font-bold text-slate-800">${b.name}</td>
          <td class="py-2.5 px-3"><span class="bg-blue-50 text-blue-700 font-bold px-2 py-0.5 rounded-full text-[10px]">Room ${roomsDisplay}</span></td>
          <td class="py-2.5 px-3 text-[10px] text-slate-600">${formatDateTime(b.checkIn)} to ${formatDateTime(effectiveOutStr)}</td>
          <td class="py-2.5 px-3 font-semibold text-slate-800">₹${b.totalAmount}</td>
          <td class="py-2.5 px-3 font-bold text-rose-600">₹${b.totalDue}</td>
          <td class="py-2.5 px-3 text-center">
            <button onclick="deleteBooking('${b.id}')" class="bg-rose-600 hover:bg-rose-700 text-white px-3 py-1 rounded-full text-[10px] font-semibold flex items-center gap-1 mx-auto transition shadow-xs">
              <i class="fa-solid fa-trash-can text-[9px]"></i> Delete Linked Booking
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }
    function clearMasterBookingSearch() {
      const input = document.getElementById('master-booking-search-input');
      if (input) input.value = '';
      searchMasterBookingById();
    }
    function generateIDsForYear(checkInDateStr) {
      let targetYear = defaultAppYear;
      if (checkInDateStr) {
        targetYear = new Date(checkInDateStr.replace(' ', 'T')).getFullYear() || defaultAppYear;
      }
      if (!state.yearlyCounters) state.yearlyCounters = {};
      if (!state.yearlyCounters[targetYear]) {
        const countForYear = state.bookings.filter(b => {
          return b.checkIn && new Date(b.checkIn.replace(' ', 'T')).getFullYear() === targetYear;
        }).length;
        state.yearlyCounters[targetYear] = countForYear;
      }
      state.yearlyCounters[targetYear] += 1;
      const seq = state.yearlyCounters[targetYear];
      const paddedSeq = String(seq).padStart(7, '0');
      return {
        bookingCode: `BKG-${targetYear}-${paddedSeq}`,
        invoiceNo: `INV-${targetYear}-${paddedSeq}`
      };
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
        try {
            sheetData = JSON.parse(textData);
        } catch(err) {
            console.error("JSON Error: Sync issue.", textData);
            throw new Error("Invalid response format from server (Sync Failed).");
        }
        
        if (sheetData && sheetData.bookings) {
          state = sheetData;
          refreshAllUI(); 
          msg.innerText = 'Database synced successfully!';
        }
      } catch (error) {
        console.error("Error loading data from Google Sheets:", error);
        msg.innerText = 'Failed to connect to Database.';
      }
      
      checkSheetRowLimits();
      setTimeout(() => toast.classList.add('hidden'), 2000);
    }
    
    function setMinBookingDates() {
      const checkInInput = document.getElementById('cust-checkin-date');
      const checkOutInput = document.getElementById('cust-checkout-date');
      const extDateInput = document.getElementById('cust-ext-checkout-date');
      
      if (checkInInput) checkInInput.removeAttribute('min');
      if (checkOutInput) checkOutInput.removeAttribute('min');
      if (checkOutInput && extDateInput) {
        extDateInput.min = checkOutInput.value;
      }
    }
    document.addEventListener("DOMContentLoaded", () => {
      checkAuthStatus();
      loadSavedData();
      setMinBookingDates();
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
      checkUpcomingCheckoutsWithDue();
      setInterval(checkUpcomingCheckoutsWithDue, 60000);
      setInterval(triggerPeriodicAutoSave, 300000);
    });
    function refreshDynamicUI() {
      if (document.getElementById('tab-booking') && !document.getElementById('tab-booking').classList.contains('hidden')) {
        renderBookingsTable(document.getElementById('booking-date-search').value);
      }
      if (document.getElementById('tab-dashboard') && !document.getElementById('tab-dashboard').classList.contains('hidden')) {
        updateDashboardCards();
      }
    }
    function triggerPeriodicAutoSave() {
      saveChanges(true, true);
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
        const payload = {
          action: "saveData",
          state: state
        };
        const response = await fetch(GAS_API_URL, {
          method: "POST",
          headers: {
            "Content-Type": "text/plain;charset=utf-8"
          },
          body: JSON.stringify(payload)
        });
        const textResult = await response.text();
        let result;
        try {
            result = JSON.parse(textResult);
        } catch(e) {
            console.error("Save JSON error", textResult);
            throw new Error("Invalid response format received from server.");
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
        console.error("Error saving to Google Sheets:", error);
        if (!quiet) {
          alert("Saving Error: " + error.message + "\n\nChecks:\n1. Ensure 'Who has access' is set to 'Anyone' in Web App deployment.\n2. Ensure URL in GAS_API_URL is correct.");
          document.getElementById('toast').classList.add('hidden');
        }
      }
    }
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
      if (val === 'CURRENT') {
        val = defaultAppYear;
      }
      
      const select = document.getElementById('dash-year-select');
      if (select) select.value = val;
      if (val === 'ALL') {
        state.dashSelectedYear = 'ALL';
      } else {
        state.dashSelectedYear = parseInt(val);
      }
      initDashboard();
    }
    function checkUpcomingCheckoutsWithDue() {
      const alertBookings = state.bookings.filter(b => {
        if (!isRoomInMaster(b.roomNo) || isInactiveBooking(b)) return false;
        
        const hasDue = (b.totalDue || 0) > 0;
        return hasDue;
      });
      const badge = document.getElementById('alert-badge');
      if (alertBookings.length > 0) {
        badge.innerText = alertBookings.length;
        badge.classList.remove('hidden');
      } else {
        badge.classList.add('hidden');
      }
      renderAlertModalList(alertBookings);
      refreshDynamicUI();
    }
    function renderAlertModalList(alertList) {
      const container = document.getElementById('alert-list-container');
      const textCount = document.getElementById('alert-list-count-text');
      container.innerHTML = '';
      textCount.innerText = `${alertList.length} active warnings found`;
      if (alertList.length === 0) {
        container.innerHTML = `
          <div class="text-center py-8 space-y-1">
            <div class="bg-emerald-50 text-emerald-600 w-10 h-10 rounded-2xl flex items-center justify-center mx-auto text-base">
              <i class="fa-solid fa-circle-check"></i>
            </div>
            <p class="font-bold text-slate-800">No Due Payment Alerts</p>
            <p class="text-slate-400 text-[10px]">All bookings have clear payments with no pending dues.</p>
          </div>
        `;
        return;
      }
      alertList.forEach((b, i) => {
        const effectiveOut = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
        const timeFormatted = formatDateTime(effectiveOut);
        const roomsDisplay = getBookingRooms(b).join(', ');
        const card = document.createElement('div');
        card.className = "bg-amber-50/60 border border-amber-200/80 rounded-2xl overflow-hidden shadow-xs";
        
        const alertMessageText = `Checkout: <strong>${timeFormatted}</strong> | Room ${roomsDisplay} | Guest: <strong>${b.name}</strong> | Total: ₹${b.totalAmount} | Due: ₹${b.totalDue}`;
        const alertBadgeHtml = `<span class="text-[11px] font-black text-rose-600 bg-rose-50 border border-rose-200 px-2.5 py-0.5 rounded-full">₹${b.totalDue.toLocaleString('en-IN')} Due</span>`;
        card.innerHTML = `
          <div class="p-3 flex justify-between items-center cursor-pointer hover:bg-amber-100/50 transition" onclick="toggleAlertDetails('alert-details-${i}')">
            <div class="flex items-center space-x-2.5">
              <span class="bg-amber-500 text-white p-2 rounded-xl text-[10px] font-bold shadow-xs"><i class="fa-solid fa-clock"></i></span>
              <div>
                <h4 class="font-bold text-slate-900 text-[11px] flex items-center gap-1.5">
                  ${b.name} <span class="bg-blue-50 text-blue-700 text-[9px] px-2 py-0.5 rounded-full font-mono">${b.bookingCode || 'N/A'}</span>
                  <span class="bg-slate-100 text-slate-700 text-[9px] px-2 py-0.5 rounded-full font-medium">Room ${roomsDisplay}</span>
                </h4>
                <p class="text-[10px] text-slate-600 mt-0.5">${alertMessageText}</p>
              </div>
            </div>
            <div class="flex items-center space-x-1.5">
              ${alertBadgeHtml}
              <i class="fa-solid fa-chevron-down text-slate-400 text-[10px]"></i>
            </div>
          </div>
          <div id="alert-details-${i}" class="hidden bg-white border-t border-amber-200/60 p-3 space-y-2 text-[10px]">
            <div class="grid grid-cols-2 gap-1 text-slate-600">
              <div>Total Charges: <strong>₹${b.totalAmount}</strong></div>
              <div>Advance Paid: <strong class="text-emerald-600">₹${b.initialAdv || 0}</strong></div>
            </div>
            <div class="flex justify-end pt-1 border-t border-slate-100">
              <button onclick="closeAlertModal(); openBookingModal('${b.id}')" class="bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded-full font-bold text-[10px] flex items-center gap-1 transition shadow-xs">
                <i class="fa-solid fa-wallet"></i> View / Edit Booking
              </button>
            </div>
          </div>
        `;
        container.appendChild(card);
      });
    }
    function toggleAlertDetails(elemId) {
      const detailsBox = document.getElementById(elemId);
      if (detailsBox) detailsBox.classList.toggle('hidden');
    }
    function openAlertModal() {
      checkUpcomingCheckoutsWithDue();
      document.getElementById('alert-modal').classList.remove('hidden');
    }
    function closeAlertModal() {
      document.getElementById('alert-modal').classList.add('hidden');
    }
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
    function toggleRoomDropdown() {
      document.getElementById('room-checkboxes').classList.toggle('hidden');
    }
    function populateRoomDropdown(selectedRoomNos = []) {
      const container = document.getElementById('room-checkboxes');
      if (!container) return;
      container.innerHTML = '';
      let selArr = [];
      if (Array.isArray(selectedRoomNos)) selArr = selectedRoomNos.map(String);
      else if (selectedRoomNos) selArr = String(selectedRoomNos).split(/[,|]/).map(s => s.trim());
      const allDiv = document.createElement('div');
      allDiv.className = "flex items-center gap-2 mb-1.5 pb-1.5 border-b border-slate-100";
      allDiv.innerHTML = `
        <input type="checkbox" id="room-all" value="ALL" onchange="handleRoomSelection(this)" class="w-3.5 h-3.5 text-blue-600 rounded border-slate-300 focus:ring-blue-500 cursor-pointer room-chk">
        <label for="room-all" class="text-[11px] font-bold text-slate-700 cursor-pointer flex-1">Select all rooms</label>
      `;
      container.appendChild(allDiv);
      state.roomsCapacity.forEach(m => {
        const isChecked = selArr.includes(String(m.roomNo)) ? 'checked' : '';
        const div = document.createElement('div');
        div.className = "flex items-center gap-2 py-1";
        div.innerHTML = `
          <input type="checkbox" id="room-${m.roomNo}" value="${m.roomNo}" ${isChecked} onchange="handleRoomSelection(this)" class="w-3.5 h-3.5 text-blue-600 rounded border-slate-300 focus:ring-blue-500 cursor-pointer room-chk item-chk">
          <label for="room-${m.roomNo}" class="text-[11px] font-bold text-slate-700 cursor-pointer flex-1">Room ${m.roomNo}</label>
        `;
        container.appendChild(div);
      });
      updateRoomDropdownText();
      autoCaptureRoomDetails();
    }
    function handleRoomSelection(chk) {
      const allChk = document.getElementById('room-all');
      const itemChks = document.querySelectorAll('.item-chk');
      if (chk.value === 'ALL') {
        itemChks.forEach(c => c.checked = chk.checked);
      } else {
        const allSelected = Array.from(itemChks).every(c => c.checked);
        if (allChk) allChk.checked = allSelected;
      }
      
      updateRoomDropdownText();
      autoCaptureRoomDetails();
    }
    function updateRoomDropdownText() {
      const itemChks = document.querySelectorAll('.item-chk');
      const checkedVals = Array.from(itemChks).filter(c => c.checked).map(c => c.value);
      const textSpan = document.getElementById('room-dropdown-text');
      
      if (checkedVals.length === 0) {
        textSpan.innerText = "Select Rooms...";
      } else if (checkedVals.length === itemChks.length) {
        textSpan.innerText = "All Rooms Selected";
        const allChk = document.getElementById('room-all');
        if(allChk) allChk.checked = true;
      } else {
        textSpan.innerText = checkedVals.map(r => `Room ${r}`).join(', ');
      }
    }
    
    function getSelectedRooms() {
      const itemChks = document.querySelectorAll('.item-chk');
      if(!itemChks.length) return [];
      const checkedVals = Array.from(itemChks).filter(c => c.checked).map(c => c.value);
      if (checkedVals.length === itemChks.length) return ["ALL"];
      return checkedVals;
    }
    function populateAgentDropdown(selectedAgentName = "") {
      const agentSelect = document.getElementById('cust-agent');
      if (!agentSelect) return;
      agentSelect.innerHTML = '';
      state.masterAgents.forEach(a => {
        const opt = document.createElement('option');
        opt.value = `${a.agentName} (${a.phone})`;
        opt.text = `${a.agentName} (${a.phone})`;
        if (selectedAgentName && opt.value.includes(selectedAgentName)) {
          opt.selected = true;
        }
        agentSelect.appendChild(opt);
      });
    }
    function autoCaptureRoomDetails() {
      let totalCap = 0;
      let allSelected = false;
      const allChk = document.getElementById('room-all');
      if (allChk && allChk.checked) {
         allSelected = true;
      }
      if (allSelected) {
         totalCap = state.roomsCapacity.reduce((sum, m) => sum + (m.capacity || 1), 0);
      } else {
         const itemChks = document.querySelectorAll('.item-chk');
         itemChks.forEach(chk => {
           if (chk.checked) {
             const matched = state.roomsCapacity.find(m => String(m.roomNo) === chk.value);
             if (matched) totalCap += (matched.capacity || 1);
           }
         });
      }
      
      document.getElementById('cust-capacity').value = totalCap > 0 ? totalCap : 1;
      calculateModalBilling();
    }
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
    function selectDashboardYear(year) {
      handleDashboardYearChange(year);
      renderCalendar(year);
      switchTab('calendar');
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
        if (isCurrentRealYear) {
          item.title = "Current Active Year";
        }
        item.onclick = () => selectDashboardYear(y);
        grid.appendChild(item);
      }
      updateDashboardCards();
    }
    function updatePieChart(elementId, liveVal = null, upcomingVal = null, closedVal = null) {
      const elem = document.getElementById(elementId);
      if (!elem) return;
      if (liveVal === null || upcomingVal === null || closedVal === null) {
        elem.style.background = 'conic-gradient(#e2e8f0 0% 100%)';
        return;
      }
      const live = Number(liveVal) || 0;
      const upcoming = Number(upcomingVal) || 0;
      const closed = Number(closedVal) || 0;
      const total = live + upcoming + closed;
      if (total <= 0) {
        elem.style.background = 'conic-gradient(#e2e8f0 0% 100%)';
        return;
      }
      const p1 = ((live / total) * 100).toFixed(2);
      const p2 = (((live + upcoming) / total) * 100).toFixed(2);
      elem.style.background = `conic-gradient(
        #f59e0b 0% ${p1}%,
        #3b82f6 ${p1}% ${p2}%,
        #10b981 ${p2}% 100%
      )`;
    }
    function updateDashboardLiveBookings(liveBookings) {
      const container = document.getElementById('dash-today-live-container');
      if (!container) return;
      if (!liveBookings || !Array.isArray(liveBookings) || liveBookings.length === 0) {
        container.innerHTML = '<p class="text-[11px] text-slate-400 italic py-2 text-center">No live bookings active today.</p>';
        return;
      }
      let html = `
        <table class="w-full text-left text-[10px] divide-y divide-slate-100">
          <thead class="bg-amber-50/60 text-amber-900 font-bold sticky top-0 bg-white">
            <tr>
              <th class="py-1 px-1.5">Guest Name</th>
              <th class="py-1 px-1.5">Room No</th>
              <th class="py-1 px-1.5">Total Amount</th>
              <th class="py-1 px-1.5">Due Amount</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-slate-100 font-medium text-slate-700">
      `;
      liveBookings.forEach(b => {
        const roomList = typeof getBookingRooms === 'function' ? getBookingRooms(b) : (b.rooms || []);
        const roomsStr = Array.isArray(roomList) && roomList.length > 0 ? roomList.join(', ') : '-';
        const rawOutDate = b.hasExtendedCheckout && b.extendedCheckOut ? b.extendedCheckOut : b.checkOut;
        const outStr = typeof format24hDate === 'function' ? format24hDate(rawOutDate) : (rawOutDate || '-');
        html += `
          <tr class="hover:bg-amber-50/40 transition">
            <td class="py-1 px-1.5 font-mono text-blue-600 font-bold">${b.name || b.id || 'N/A'}</td>
            <td class="py-1 px-1.5 font-bold text-slate-800">${roomsStr}</td>
            <td class="py-1 px-1.5"><span class="bg-amber-100 text-amber-800 font-bold px-1.5 py-0.5 rounded-md text-[9px]">₹ ${Number(b.totalAmount || 0).toLocaleString('en-IN')}</span></td>
            <td class="py-1 px-1.5 text-slate-600">₹ ${Number(b.totalDue || 0).toLocaleString('en-IN')}</td>
          </tr>
        `;
      });
      html += `</tbody></table>`;
      container.innerHTML = html;
    }
    function updateDashboardCards() {
      const selectedFilter = typeof state !== 'undefined' ? state.dashSelectedYear : null;
      const label = document.getElementById('dash-filter-label');
      const targetYear = (selectedFilter && selectedFilter !== 'ALL') ? parseInt(selectedFilter, 10) : null;
      if (label) {
        if (!targetYear) {
          label.innerText = "Consolidated Summary (All Years)";
        } else {
          const defaultYr = typeof defaultAppYear !== 'undefined' ? defaultAppYear : new Date().getFullYear();
          label.innerText = targetYear === defaultYr
            ? `Year ${targetYear} (Current Year)`
            : `Year ${targetYear}`;
        }
      }
      const now = Date.now();
      const liveBookingsList = []; 
      const stats = {
        live: { count: 0, amount: 0, adv: 0, due: 0 },
        upcoming: { count: 0, amount: 0, adv: 0, due: 0 },
        closed: { count: 0, amount: 0, adv: 0, due: 0 },
        inactive: { count: 0, amount: 0, adv: 0, due: 0 }
      };
      const parseDate = (dtStr) => {
        if (!dtStr) return NaN;
        return new Date(String(dtStr).replace(' ', 'T')).getTime();
      };
      const bookings = (typeof state !== 'undefined' && Array.isArray(state.bookings)) ? state.bookings : [];
      bookings.forEach(b => {
        const checkInMs = parseDate(b.checkIn);
        const checkOutMs = getEffectiveCheckoutTime(b);
        if (targetYear && !isNaN(checkInMs)) {
          const yr = new Date(checkInMs).getFullYear();
          if (yr !== targetYear) return;
        }
        const amt = Number(b.totalAmount || 0);
        const adv = Number(b.initialAdv || 0) + Number(b.clearedDue || 0);
        const due = Number(b.totalDue || 0);
        if (typeof isInactiveBooking === 'function' && isInactiveBooking(b)) {
          stats.inactive.count++;
          stats.inactive.amount += amt;
          stats.inactive.adv += adv;
          stats.inactive.due += due;
          return;
        }
        if (!isNaN(checkInMs) && !isNaN(checkOutMs)) {
          if (now >= checkInMs && now <= checkOutMs) {
            stats.live.count++;
            stats.live.amount += amt;
            stats.live.adv += adv;
            stats.live.due += due;
            liveBookingsList.push(b); 
          } else if (now < checkInMs) {
            stats.upcoming.count++;
            stats.upcoming.amount += amt;
            stats.upcoming.adv += adv;
            stats.upcoming.due += due;
          } else {
            stats.closed.count++;
            stats.closed.amount += amt;
            stats.closed.adv += adv;
            stats.closed.due += due;
          }
        }
      });
      const totalBookings = stats.live.count + stats.upcoming.count + stats.closed.count;
      const totalAmt = stats.live.amount + stats.upcoming.amount + stats.closed.amount;
      const totalAdv = stats.live.adv + stats.upcoming.adv + stats.closed.adv;
      const totalDue = stats.live.due + stats.upcoming.due + stats.closed.due;
      const setTxt = (id, text) => {
        const el = document.getElementById(id);
        if (el) el.innerText = text;
      };
      const fmt = num => `₹${Number(num || 0).toLocaleString('en-IN')}`;
      setTxt('dash-total-bookings', totalBookings);
      setTxt('dash-count-live', stats.live.count);
      setTxt('dash-count-upcoming', stats.upcoming.count);
      setTxt('dash-count-closed', stats.closed.count);
      updatePieChart('pie-bookings', stats.live.count, stats.upcoming.count, stats.closed.count, stats.inactive.count);
      setTxt('dash-total-amount', fmt(totalAmt));
      setTxt('dash-amount-live', fmt(stats.live.amount));
      setTxt('dash-amount-upcoming', fmt(stats.upcoming.amount));
      setTxt('dash-amount-closed', fmt(stats.closed.amount));
      updatePieChart('pie-amount', stats.live.amount, stats.upcoming.amount, stats.closed.amount, stats.inactive.amount);
      setTxt('dash-advanced', fmt(totalAdv));
      setTxt('dash-recv-live', fmt(stats.live.adv));
      setTxt('dash-recv-upcoming', fmt(stats.upcoming.adv));
      setTxt('dash-recv-closed', fmt(stats.closed.adv));
      updatePieChart('pie-received', stats.live.adv, stats.upcoming.adv, stats.closed.adv, stats.inactive.adv);
      setTxt('dash-due', fmt(totalDue));
      setTxt('dash-due-live', fmt(stats.live.due));
      setTxt('dash-due-upcoming', fmt(stats.upcoming.due));
      setTxt('dash-due-closed', fmt(stats.closed.due));
      updatePieChart('pie-due', stats.live.due, stats.upcoming.due, stats.closed.due, stats.inactive.due);
      setTxt('dash-inactive-count', `${stats.inactive.count} Bookings`);
      setTxt('dash-inactive-amount', fmt(stats.inactive.amount));
      updateDashboardLiveBookings(liveBookingsList);
    }
    
    function sendReceiptViaWhatsApp() {
      if (!activeModalBooking) {
        alert("⚠️ Booking information not found!");
        return;
      }
      const b = activeModalBooking;
      let rawCountryCode = b.countryCode ? String(b.countryCode).replace(/\D/g, '') : '91';
      let phone = b.contactNo ? String(b.contactNo).replace(/\D/g, '') : '';
      let initialAdvanceVal = b.initialAdv !== undefined ? parseFloat(b.initialAdv) : 0;
      
      let validationErrors = [];
      if (!phone || phone.length !== 10) {
        validationErrors.push("• Guest contact number must be exactly 10 digits.");
      }
      if (!(initialAdvanceVal > 0)) {
        validationErrors.push("• Advanced payment must be greater than 0.");
      }
      if (validationErrors.length > 0) {
        alert("⚠️ Cannot send via WhatsApp:\n\n" + validationErrors.join("\n"));
        return;
      }
      const fullPhoneNumber = rawCountryCode + phone;
      const upiId = "aniruddha.e@oksbi";
      const effectiveOut = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
      const roomsDisplay = getBookingRooms(b).join(', ');
      const messageText = `*Sanoum Pema Homestay-by Anaristays - Booking Receipt*\n\n` +
        `Dear *${b.name}*,\n` +
        `Thank you for booking with us! Here are your booking details:\n\n` +
        `*Reservation Details:*\n` +
        `• Booking ID: *${b.bookingCode}*\n` +
        `• Room No: ${roomsDisplay}\n` +
        `• Check-In: ${formatDateTime(b.checkIn)}\n` +
        `• Check-Out: ${formatDateTime(effectiveOut)}\n\n` +
        `*Billing Summary:*\n` +
        `• Total Amount: ₹${b.totalAmount}\n` +
        `• Advance Amount: ₹${initialAdvanceVal}\n` +
        `• Balance Due: ₹${b.totalDue}\n\n` +
        `*UPI Payment Details:*\n` +
        `• UPI ID: *${upiId}*\n\n` +
        `We look forward to hosting you! 🏠`;
      const encodedMessage = encodeURIComponent(messageText);
      const whatsappUrl = `https://api.whatsapp.com/send?phone=${fullPhoneNumber}&text=${encodedMessage}`;
      window.open(whatsappUrl, '_blank');
    }
    function printInvoice(bookingId) {
      const bIndex = state.bookings.findIndex(item => String(item.id) === String(bookingId));
      if (bIndex === -1) {
        alert("Booking details not found!");
        return;
      }
      const b = state.bookings[bIndex];
      activeModalBooking = b;
      
      const today = formatDate(new Date());
      const readOnlyNotice = document.getElementById('inv-readonly-notice');
      const invPrintBtn = document.getElementById('inv-print-btn');
      const waBtn = document.getElementById('inv-whatsapp-btn');
      const eInvoiceSection = document.getElementById('e-invoice-section');
      const invIdContainer = document.getElementById('inv-id-container');
      const now = new Date().getTime();
      const cIn = parseDateMs(b.checkIn);
      const cOut = getEffectiveCheckoutTime(b);
      const isClosed = now > cOut;
      const isInactive = isInactiveBooking(b);
      const isLive = now >= cIn && now <= cOut;
      const isUpcoming = now < cIn;
      if (waBtn) {
        if (isClosed || isInactive) {
          waBtn.classList.add('hidden');
        } else {
          waBtn.classList.remove('hidden');
          const contactDigits = b.contactNo ? String(b.contactNo).replace(/\D/g, '') : '';
          const hasValidContact = contactDigits.length === 10;
          const hasAdvanced = (parseFloat(b.initialAdv) || 0) > 0;
          if (hasValidContact && hasAdvanced) {
            waBtn.classList.remove('opacity-50', 'cursor-not-allowed');
            waBtn.title = "Send Receipt via WhatsApp";
          } else {
            waBtn.classList.add('opacity-50', 'cursor-not-allowed');
            waBtn.title = "Inactive: Advance payment must be > 0 and contact number must be 10 digits.";
          }
        }
      }
      const hasDue = (b.totalDue || 0) > 0;
      if (eInvoiceSection) {
        if (hasDue) {
          eInvoiceSection.classList.add('hidden');
        } else {
          eInvoiceSection.classList.remove('hidden');
        }
      }
      if (invIdContainer) {
        invIdContainer.querySelector('strong').innerText = b.invoiceNo || 'INV-2026-0000001';
      }
      if (invPrintBtn) {
        invPrintBtn.classList.remove('hidden');
        invPrintBtn.disabled = false;
        invPrintBtn.className = "px-4 py-1.5 bg-blue-600 text-white rounded-xl font-semibold shadow-sm flex items-center gap-1 transition hover:bg-blue-700 cursor-pointer";
        invPrintBtn.innerHTML = `<i class="fa-solid fa-print"></i> Print Invoice`;
      }
      document.getElementById('inv-booking-id').innerText = b.bookingCode || 'N/A';
      document.getElementById('inv-date').innerText = today;
      const fullLocation = [b.address, b.city, b.state, b.country, b.zipCode].filter(Boolean).map(formatTitleCase).join(', ');
      document.getElementById('inv-guest-name').innerText = formatTitleCase(b.name) || 'N/A';
      document.getElementById('inv-guest-address').innerText = `Address: ${fullLocation || 'N/A'}`;
      
      const fullGuestPhone = b.contactNo ? `${b.countryCode || '+91'} ${b.contactNo}`.trim() : '-';
      document.getElementById('inv-guest-contact').innerText = `Contact: ${fullGuestPhone}`;
      document.getElementById('inv-guest-id').innerText = `ID No: ${b.idNo || 'N/A'}`;
      const roomsDisplay = getBookingRooms(b).join(', ');
      document.getElementById('inv-room').innerText = `Room No: ${roomsDisplay}`;
      document.getElementById('inv-checkin').innerText = `Check-in: ${formatDateTime(b.checkIn)}`;
      document.getElementById('inv-checkout').innerText = `Check-out: ${formatDateTime(b.checkOut)}`;
      const extCheckoutElem = document.getElementById('inv-ext-checkout');
      if (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) {
        extCheckoutElem.innerHTML = `Extended Check-out: ${formatDateTime(b.extendedCheckOut)} <span class="text-[9px] text-blue-700 bg-blue-50 border border-blue-200 px-1.5 py-0.2 rounded-full font-bold ml-1">Extended</span>`;
        extCheckoutElem.classList.remove('hidden');
      } else {
        extCheckoutElem.innerText = '';
        extCheckoutElem.classList.add('hidden');
      }
      const tbody = document.getElementById('inv-items-tbody');
      tbody.innerHTML = '';
      const roomTotal = (b.noOfDays || 0) * (b.perDayPrice || 0) * (b.capacity || 1);
      const showMealsNote = b.includeMeals !== false && b.includeMeals !== 'false';
      const mealNotesStr = showMealsNote ? '<span class="text-[9px] text-slate-500 block font-normal">(*Include Breakfast,Lunch,Evening snack & Dinner)</span>' : '';
      const stayDaysCount = parseInt(b.noOfDays) || 0;
      const daysFormattedStr = stayDaysCount === 1 ? '1 Day' : `${stayDaysCount} Days`;
      const roomCapacityCount = parseInt(b.capacity) || 1;
      const roomCapacityLabel = roomCapacityCount === 1 ? 'Person' : 'Persons';
      const roomTr = document.createElement('tr');
      roomTr.innerHTML = `
        <td class="p-2.5 font-semibold text-slate-800">
          Room ${roomsDisplay} Accommodation (${roomCapacityCount} ${roomCapacityLabel}) ${isTrue(b.hasExtendedCheckout) ? '<span class="text-[9px] text-blue-600 block font-normal">(Includes extended stay duration)</span>' : ''}
          ${mealNotesStr}
        </td>
        <td class="p-2.5 text-center">${daysFormattedStr}</td>
        <td class="p-2.5 text-right">₹${(b.perDayPrice || 0).toLocaleString('en-IN')}</td>
        <td class="p-2.5 text-right font-semibold text-slate-800">₹${roomTotal.toLocaleString('en-IN')}</td>
      `;
      tbody.appendChild(roomTr);
      if (b.extraPersons && b.extraPersons > 0 && b.extraPersonDays > 0) {
        const extraPrice = b.extraPersonPricePerDay !== undefined ? parseFloat(b.extraPersonPricePerDay) : parseFloat(b.perDayPrice || 0);
        const extraPersonTotal = b.extraPersonTotalRate !== undefined ? parseFloat(b.extraPersonTotalRate) : (b.extraPersons * b.extraPersonDays * extraPrice);
        const extraJoinedFmt = b.extraPersonJoined ? formatDateTime(b.extraPersonJoined) : '';
        const extraOutFmt = b.extraPersonOut ? formatDateTime(b.extraPersonOut) : '';
        const extraDaysCount = parseInt(b.extraPersonDays) || 0;
        const extraDaysFormattedStr = extraDaysCount === 1 ? '1 Day' : `${extraDaysCount} Days`;
        const extraTr = document.createElement('tr');
        extraTr.innerHTML = `
          <td class="p-2.5 font-semibold text-amber-900">
            Additional Guest Accommodation (${b.extraPersons} ${b.extraPersons === 1 ? 'Person' : 'Persons'})
            <span class="text-[9px] text-amber-700 font-normal block">Stay: ${extraJoinedFmt} to ${extraOutFmt || 'Check-Out'}</span>
          </td>
          <td class="p-2.5 text-center">${extraDaysFormattedStr}</td>
          <td class="p-2.5 text-right">₹${extraPrice.toLocaleString('en-IN')}</td>
          <td class="p-2.5 text-right font-semibold text-amber-900">₹${extraPersonTotal.toLocaleString('en-IN')}</td>
        `;
        tbody.appendChild(extraTr);
      }
      const foodList = parseJSONField(b.foodOrders);
      if (foodList.length > 0) {
        foodList.forEach(fo => {
          if (fo.foodCharge > 0) {
            const foodTr = document.createElement('tr');
            const foodDateTimeFmt = fo.foodDateTime ? ` (${formatDateTime(fo.foodDateTime)})` : '';
            const plateCount = parseInt(fo.plates) || 1;
            const plateLabel = plateCount === 1 ? 'Plate' : 'Plates';
            foodTr.innerHTML = `
              <td class="p-2.5 font-semibold text-slate-800">Extra Food <span class="text-[9px] text-slate-500 font-normal block">${fo.foodDesc || 'Food Item'}${foodDateTimeFmt}</span></td>
              <td class="p-2.5 text-center">${plateCount} ${plateLabel}</td>
              <td class="p-2.5 text-right">₹${(fo.itemPrice || 0).toLocaleString('en-IN')}</td>
              <td class="p-2.5 text-right font-semibold text-slate-800">₹${(fo.foodCharge || 0).toLocaleString('en-IN')}</td>
            `;
            tbody.appendChild(foodTr);
          }
        });
      }
      
      const cabList = parseJSONField(b.cabTrips);
      if (cabList.length > 0) {
        cabList.forEach(trip => {
          if (trip.rate > 0) {
             const cabTr = document.createElement('tr');
             const dtFormat = trip.dateTime ? ` (${formatDateTime(trip.dateTime)})` : '';
             cabTr.innerHTML = `
              <td class="p-2.5 font-semibold text-indigo-900">
                Cab Fare - ${trip.tripName}${dtFormat}
                ${trip.remark ? `<span class="text-[9px] text-indigo-700 font-normal block">Remark: ${trip.remark}</span>` : ''}
              </td>
              <td class="p-2.5 text-center">1 Trip</td>
              <td class="p-2.5 text-right">₹${(trip.rate || 0).toLocaleString('en-IN')}</td>
              <td class="p-2.5 text-right font-semibold text-indigo-900">₹${(trip.rate || 0).toLocaleString('en-IN')}</td>
            `;
            tbody.appendChild(cabTr);
          }
        });
      }
      const initialAdv = b.initialAdv || 0;
      const clearDueAmt = b.clearedDue || 0;
      document.getElementById('inv-sum-total').innerText = `₹${(b.totalAmount || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-advance').innerText = `₹${(initialAdv || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-due').innerText = `₹${(b.totalDue || 0).toLocaleString('en-IN')}`;
      const clearDueRow = document.getElementById('inv-clear-due-row');
      if (clearDueAmt > 0) {
        document.getElementById('inv-sum-clear-due').innerText = `₹${clearDueAmt.toLocaleString('en-IN')}`;
        clearDueRow.classList.remove('hidden');
      } else {
        clearDueRow.classList.add('hidden');
      }
      document.getElementById('invoice-modal').classList.remove('hidden');
    }
    function closeInvoiceModal() {
      activeModalBooking = null;
      document.getElementById('invoice-modal').classList.add('hidden');
    }
    function addFoodOrderItem(desc = '', plates = 1, itemPrice = 0, charge = 0, dateStr = '', timeStr = '', disabled = false) {
      const foodWin = getModalFoodWindow();
      if (!foodWin && !disabled) {
        alert("⚠️ Please enter valid Check-In and Check-Out date & time first before adding extra food!");
        return;
      }
      if (foodWin && foodWin.minFoodDt >= foodWin.maxFoodDt && !disabled) {
        alert("⚠️ Invalid stay window! The duration between Check-In (+15m) and Check-Out (-30m) is too short to order food.");
        return;
      }
      if (!dateStr || !timeStr) {
        if (foodWin) {
          const now = new Date();
          let defaultDt = now;
          if (now < foodWin.minFoodDt || now > foodWin.maxFoodDt) {
            defaultDt = foodWin.minFoodDt;
          }
          
          const utcMs = defaultDt.getTime();
          const istDate = new Date(utcMs + (330 * 60000));
          
          const yyyy = istDate.getUTCFullYear();
          const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
          const dd = String(istDate.getUTCDate()).padStart(2, '0');
          const hh = String(istDate.getUTCHours()).padStart(2, '0');
          const min = String(istDate.getUTCMinutes()).padStart(2, '0');
          if (!dateStr) dateStr = `${yyyy}-${mm}-${dd}`;
          if (!timeStr) timeStr = `${hh}:${min}`;
        }
      }
      const container = document.getElementById('food-orders-container');
      const itemRow = document.createElement('div');
      itemRow.className = "food-order-row grid grid-cols-1 sm:grid-cols-12 gap-1.5 items-end bg-white p-2.5 rounded-2xl border border-amber-200/80 shadow-xs";
      
      const disabledAttr = disabled ? 'disabled' : '';
      const bgClass = disabled ? 'bg-slate-100 cursor-not-allowed text-slate-500' : 'bg-white';
      itemRow.innerHTML = `
        <div class="sm:col-span-3">
          <label class="block font-semibold text-slate-600 mb-0.5">Item Name</label>
          <input type="text" value="${desc}" ${disabledAttr} placeholder="e.g. Thali / Tea" class="cust-food-desc w-full ${bgClass} border border-slate-200 rounded-xl px-2.5 py-1 focus:outline-none focus:border-amber-500">
        </div>
        <div class="sm:col-span-3">
          <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-regular fa-clock text-amber-600 mr-1"></i> Date & Time</label>
          <div class="flex gap-1">
            <input type="date" value="${dateStr}" ${disabledAttr} onchange="validateFoodRowDateTime(this)" class="cust-food-date w-3/5 ${bgClass} border border-slate-200 rounded-xl px-1.5 py-1 focus:outline-none focus:border-amber-500 font-medium text-[10px]">
            <input type="time" value="${timeStr}" ${disabledAttr} onchange="validateFoodRowDateTime(this)" class="cust-food-time w-2/5 ${bgClass} border border-slate-200 rounded-xl px-1 py-1 focus:outline-none focus:border-amber-500 font-medium text-[10px]">
          </div>
        </div>
        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Price/Plate (₹)</label>
          <input type="number" value="${itemPrice}" min="0" ${disabledAttr} oninput="calculateFoodRowTotal(this)" class="cust-food-price w-full ${bgClass} border border-slate-200 rounded-xl px-2 py-1 focus:outline-none focus:border-amber-500">
        </div>
        
        <div class="sm:col-span-1">
          <label class="block font-semibold text-slate-600 mb-0.5">Plates</label>
          <input type="number" value="${plates}" min="1" ${disabledAttr} oninput="calculateFoodRowTotal(this)" class="cust-food-plates w-full ${bgClass} border border-slate-200 rounded-xl px-2 py-1 focus:outline-none focus:border-amber-500 font-bold">
        </div>
        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Total (₹)</label>
          <input type="number" value="${charge}" readonly class="cust-food-charge w-full bg-slate-100 font-bold text-amber-700 border border-slate-200 rounded-xl px-2 py-1 cursor-not-allowed">
        </div>
        <div class="sm:col-span-1 flex justify-end">
          <button type="button" onclick="removeFoodOrderItem(this)" ${disabledAttr} class="btn-remove-food-item text-rose-500 hover:text-rose-700 p-1.5 ${disabled ? 'hidden' : ''}" title="Remove Order">
            <i class="fa-solid fa-trash-can"></i>
          </button>
        </div>
      `;
      container.appendChild(itemRow);
      calculateModalBilling();
    }
    function calculateFoodRowTotal(inputElem) {
      const row = inputElem.closest('.food-order-row');
      if (!row) return;
      const price = parseFloat(row.querySelector('.cust-food-price').value) || 0;
      const plates = parseInt(row.querySelector('.cust-food-plates').value) || 0;
      const totalCharge = price * plates;
      row.querySelector('.cust-food-charge').value = totalCharge;
      calculateModalBilling();
    }
    function removeFoodOrderItem(btn) {
      const row = btn.closest('.food-order-row');
      if (row) {
        row.remove();
        calculateModalBilling();
      }
    }
    function addCabTripRow(rate = 0, dateStr = '', timeStr = '', remark = '', disabled = false) {
      const container = document.getElementById('cab-trips-container');
      const tripCount = container.children.length + 1;
      const itemRow = document.createElement('div');
      itemRow.className = "cab-trip-row grid grid-cols-1 sm:grid-cols-12 gap-1.5 items-end bg-white p-2.5 rounded-2xl border border-indigo-200/80 shadow-xs";
      const disabledAttr = disabled ? 'disabled' : '';
      const bgClass = disabled ? 'bg-slate-100 cursor-not-allowed text-slate-500' : 'bg-white';
      if (!dateStr) {
         const inDate = document.getElementById('cust-checkin-date')?.value;
         if (inDate) dateStr = inDate;
      }
      if (!timeStr) timeStr = '12:00';
      itemRow.innerHTML = `
        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Trip Name</label>
          <input type="text" value="Trip ${tripCount}" readonly class="w-full bg-slate-50 border border-slate-200 rounded-xl px-2.5 py-1 text-slate-500 font-bold cursor-not-allowed text-[10px]">
        </div>
        <div class="sm:col-span-4">
          <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-regular fa-clock text-indigo-600 mr-1"></i> Date & Time</label>
          <div class="flex gap-1">
            <input type="date" value="${dateStr}" ${disabledAttr} class="cust-cab-date w-3/5 ${bgClass} border border-slate-200 rounded-xl px-1.5 py-1 focus:outline-none focus:border-indigo-500 font-medium text-[10px]">
            <input type="time" value="${timeStr}" ${disabledAttr} class="cust-cab-time w-2/5 ${bgClass} border border-slate-200 rounded-xl px-1 py-1 focus:outline-none focus:border-indigo-500 font-medium text-[10px]">
          </div>
        </div>
        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Rate/Trip (₹)</label>
          <input type="number" value="${rate}" min="0" ${disabledAttr} oninput="calculateModalBilling()" class="cust-cab-rate w-full ${bgClass} border border-slate-200 rounded-xl px-2 py-1 focus:outline-none focus:border-indigo-500 font-bold text-indigo-700">
        </div>
        <div class="sm:col-span-3">
          <label class="block font-semibold text-slate-600 mb-0.5">Remark</label>
          <input type="text" value="${remark}" ${disabledAttr} placeholder="e.g. Airport drop" class="cust-cab-remark w-full ${bgClass} border border-slate-200 rounded-xl px-2.5 py-1 focus:outline-none focus:border-indigo-500 text-[10px]">
        </div>
        <div class="sm:col-span-1 flex justify-end">
          <button type="button" onclick="removeCabTripRow(this)" ${disabledAttr} class="text-rose-500 hover:text-rose-700 p-1.5 ${disabled ? 'hidden' : ''}" title="Remove Trip">
            <i class="fa-solid fa-trash-can"></i>
          </button>
        </div>
      `;
      container.appendChild(itemRow);
      calculateModalBilling();
    }
    function removeCabTripRow(btn) {
      const row = btn.closest('.cab-trip-row');
      if (row) {
        row.remove();
        calculateModalBilling();
      }
    }
    function setInputEnabled(elem, isEnabled) {
      if (!elem) return;
      elem.disabled = !isEnabled;
      if (!isEnabled) {
        elem.classList.add('bg-slate-100', 'cursor-not-allowed', 'text-slate-500');
        elem.classList.remove('bg-white', 'bg-amber-50');
      } else {
        elem.classList.remove('bg-slate-100', 'cursor-not-allowed', 'text-slate-500');
      }
    }
    function openBookingModal(bookingId = null) {
      const now = new Date().getTime();
      let isLiveBooking = false;
      let isClosedBooking = false;
      let isUpcomingBooking = false;
      let isPast730Days = false;
      let b = null;
      if (bookingId) {
        b = state.bookings.find(item => String(item.id) === String(bookingId));
        if (b) {
          if (isInactiveBooking(b)) {
            printInvoice(bookingId);
            return;
          }
          if (!isRoomInMaster(b.roomNo)) {
            alert("This booking details were deleted from Master Data and cannot be opened or edited.");
            return;
          }
          const effectiveOutTime = getEffectiveCheckoutTime(b);
          const checkInTime = parseDateMs(b.checkIn);
          if (now > effectiveOutTime) {
            isClosedBooking = true;
            if (now > effectiveOutTime + (730 * 24 * 60 * 60 * 1000)) {
               isPast730Days = true;
            }
          } else if (now >= checkInTime && now <= effectiveOutTime) {
            isLiveBooking = true;
          } else {
            isUpcomingBooking = true;
          }
        }
      } else {
        isUpcomingBooking = true;
      }
      document.getElementById('modal-booking-id').value = bookingId || '';
      setMinBookingDates();
      
      const form = document.getElementById('booking-form');
      form.reset();
      removeAttachedIdProof();
      
      const clearBillInput = document.getElementById('cust-clear-bill');
      if (clearBillInput) clearBillInput.value = 0;
      document.getElementById('food-orders-container').innerHTML = '';
      document.getElementById('cab-trips-container').innerHTML = '';
      populateAgentDropdown();
      setSectionEditability('sec-guest-info', !isClosedBooking);
      setSectionEditability('sec-room-dates', !isClosedBooking);
      const extChkBox = document.getElementById('cust-has-extended-checkout');
      const extDateInput = document.getElementById('cust-ext-checkout-date');
      const extTimeInput = document.getElementById('cust-ext-checkout-time');
      const timerNotice = document.getElementById('ext-checkout-timer-notice');
      let canToggleExtendedCheckout = false;
      if (b) {
        const initialCheckOutTime = parseDateMs(b.checkOut);
        const isWithin1HrPastCheckout = now > initialCheckOutTime && now <= (initialCheckOutTime + ONE_HOUR_MS);
        if (isLiveBooking || isWithin1HrPastCheckout) {
          canToggleExtendedCheckout = true;
        }
      }
      if (extChkBox) {
        extChkBox.disabled = !canToggleExtendedCheckout;
        if (canToggleExtendedCheckout) {
          if (timerNotice) {
            timerNotice.innerText = isLiveBooking ? "(Active for Live Booking)" : "(Active up to 1hr post check-out)";
            timerNotice.classList.remove('hidden', 'text-rose-600');
          }
        } else {
          if (timerNotice) {
            timerNotice.innerText = isClosedBooking ? "(Closed Booking - Inactive)" : "(Selectable only for Live Booking)";
            timerNotice.classList.remove('hidden');
            timerNotice.classList.add('text-rose-600');
          }
        }
      }
      const mealsChkBox = document.getElementById('cust-include-meals');
      if (mealsChkBox) {
        mealsChkBox.disabled = isClosedBooking;
      }
      const addFoodBtn = document.getElementById('btn-add-food-order');
      if (addFoodBtn) {
        addFoodBtn.disabled = isPast730Days;
        if (isPast730Days) {
          addFoodBtn.classList.add('opacity-50', 'cursor-not-allowed');
        } else {
          addFoodBtn.classList.remove('opacity-50', 'cursor-not-allowed');
        }
      }
      
      const addCabBtn = document.getElementById('btn-add-cab-trip');
      if (addCabBtn) {
        addCabBtn.disabled = isPast730Days;
        if (isPast730Days) {
          addCabBtn.classList.add('opacity-50', 'cursor-not-allowed');
        } else {
          addCabBtn.classList.remove('opacity-50', 'cursor-not-allowed');
        }
      }
      setSectionEditability('sec-cab-fare', !isPast730Days);
      setSectionEditability('sec-billing-summary', !isPast730Days);
      const btnSave = document.getElementById('btn-save-booking');
      if (btnSave) {
         if (isPast730Days) {
            btnSave.disabled = true;
            btnSave.classList.add('opacity-50', 'cursor-not-allowed', 'bg-slate-400');
            btnSave.classList.remove('bg-blue-600', 'hover:bg-blue-700');
         } else {
            btnSave.disabled = false;
            btnSave.classList.remove('opacity-50', 'cursor-not-allowed', 'bg-slate-400');
            btnSave.classList.add('bg-blue-600', 'hover:bg-blue-700');
         }
      }
      const extraPersonsInput = document.getElementById('cust-extra-persons');
      const extraPersonTimeWrapper = document.getElementById('sec-extra-person-time-wrapper');
      const extraPersonDateInput = document.getElementById('cust-extra-person-date');
      const extraPersonTimeInput = document.getElementById('cust-extra-person-time');
      const extraPersonOutDateInput = document.getElementById('cust-extra-person-out-date');
      const extraPersonOutTimeInput = document.getElementById('cust-extra-person-out-time');
      const canEditExtras = !isClosedBooking;
      if (extraPersonsInput) setInputEnabled(extraPersonsInput, canEditExtras);
      if (extraPersonDateInput) setInputEnabled(extraPersonDateInput, canEditExtras);
      if (extraPersonTimeInput) setInputEnabled(extraPersonTimeInput, canEditExtras);
      if (extraPersonOutDateInput) setInputEnabled(extraPersonOutDateInput, canEditExtras);
      if (extraPersonOutTimeInput) setInputEnabled(extraPersonOutTimeInput, canEditExtras);
      if (canEditExtras) {
        if (extraPersonTimeWrapper) extraPersonTimeWrapper.classList.remove('hidden');
      } else {
        if (extraPersonTimeWrapper && (!b || !b.extraPersons || b.extraPersons <= 0)) {
          extraPersonTimeWrapper.classList.add('hidden');
        } else if (extraPersonTimeWrapper) {
          extraPersonTimeWrapper.classList.remove('hidden');
        }
      }
      if (b) {
        document.getElementById('modal-title').innerText = isPast730Days ? 'Closed Booking (Read-Only)' : (isClosedBooking ? 'Closed Booking (Billing Active)' : 'Edit Booking Details');
        
        setInputEnabled(document.getElementById('cust-checkin-date'), isUpcomingBooking);
        setInputEnabled(document.getElementById('cust-checkin-time'), isUpcomingBooking);
        setInputEnabled(document.getElementById('cust-checkout-date'), isUpcomingBooking);
        setInputEnabled(document.getElementById('cust-checkout-time'), isUpcomingBooking);
        
        document.getElementById('modal-booking-id').value = b.id;
        document.getElementById('cust-name').value = formatTitleCase(b.name);
        document.getElementById('cust-address').value = formatTitleCase(b.address || '');
        document.getElementById('cust-city').value = formatTitleCase(b.city || '');
        document.getElementById('cust-state').value = formatTitleCase(b.state || '');
        document.getElementById('cust-country').value = formatTitleCase(b.country || '');
        document.getElementById('cust-zip').value = b.zipCode || '';
        document.getElementById('cust-id').value = b.idNo || '';
        document.getElementById('cust-country-code').value = b.countryCode || '+91';
        document.getElementById('cust-contact').value = b.contactNo || '';
        if (b.idProofBase64) {
          document.getElementById('cust-id-file-base64').value = b.idProofBase64;
          document.getElementById('cust-id-file-name').value = b.idProofFileName || 'Attached_ID_Proof.pdf';
          document.getElementById('cust-id-file-status').innerHTML = `<span class="text-emerald-600 font-semibold"><i class="fa-solid fa-circle-check"></i> Attached: ${b.idProofFileName || 'Attached_ID_Proof.pdf'}</span>`;
          document.getElementById('cust-id-file-remove').classList.remove('hidden');
        }
        
        populateRoomDropdown(b.roomNo);
        populateAgentDropdown(b.agentInfo);
        document.getElementById('cust-capacity').value = b.capacity || 1;
        if (extraPersonsInput) extraPersonsInput.value = b.extraPersons || 0;
        
        if (b.extraPersonJoined) {
          const parts = extractISTDateParts(b.extraPersonJoined);
          if (extraPersonDateInput) extraPersonDateInput.value = parts.date || '';
          if (extraPersonTimeInput) extraPersonTimeInput.value = parts.time || '';
        } else {
          if (extraPersonDateInput) extraPersonDateInput.value = '';
          if (extraPersonTimeInput) extraPersonTimeInput.value = '';
        }
        if (b.extraPersonOut) {
          const parts = extractISTDateParts(b.extraPersonOut);
          if (extraPersonOutDateInput) extraPersonOutDateInput.value = parts.date || '';
          if (extraPersonOutTimeInput) extraPersonOutTimeInput.value = parts.time || '';
        } else {
          if (extraPersonOutDateInput) extraPersonOutDateInput.value = '';
          if (extraPersonOutTimeInput) extraPersonOutTimeInput.value = '';
        }
        if (b.checkIn) {
          const parts = extractISTDateParts(b.checkIn);
          document.getElementById('cust-checkin-date').value = parts.date || '';
          document.getElementById('cust-checkin-time').value = parts.time || '';
        }
        if (b.checkOut) {
          const parts = extractISTDateParts(b.checkOut);
          document.getElementById('cust-checkout-date').value = parts.date || '';
          document.getElementById('cust-checkout-time').value = parts.time || '';
          if (extDateInput) extDateInput.min = parts.date || '';
        }
        extChkBox.checked = isTrue(b.hasExtendedCheckout);
        toggleExtendedCheckoutFields(extChkBox.checked);
        if (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) {
          const parts = extractISTDateParts(b.extendedCheckOut);
          extDateInput.value = parts.date || '';
          extTimeInput.value = parts.time || '';
        }
        if (mealsChkBox) {
          mealsChkBox.checked = b.includeMeals !== undefined ? (b.includeMeals !== false && b.includeMeals !== 'false') : true;
        }
        const foList = parseJSONField(b.foodOrders);
        foList.forEach(fo => {
          let fDate = '', fTime = '';
          if (fo.foodDateTime) {
            const parts = extractISTDateParts(fo.foodDateTime);
            fDate = parts.date || '';
            fTime = parts.time || '';
          }
          addFoodOrderItem(fo.foodDesc || '', fo.plates || 1, fo.itemPrice || 0, fo.foodCharge || 0, fDate, fTime, isClosedBooking);
        });
        
        const tripsList = parseJSONField(b.cabTrips);
        if (tripsList.length > 0) {
          tripsList.forEach(trip => {
            let cDate = '', cTime = '';
            if (trip.dateTime) {
               const parts = extractISTDateParts(trip.dateTime);
               cDate = parts.date;
               cTime = parts.time;
            } else {
               cDate = trip.dateStr || '';
               cTime = trip.timeStr || '';
            }
            addCabTripRow(trip.rate || 0, cDate, cTime, trip.remark || '', isClosedBooking);
          });
        }
        document.getElementById('cust-price').value = b.perDayPrice || 1200;
        
        const extraPriceElem = document.getElementById('cust-extra-price');
        if (extraPriceElem) {
          extraPriceElem.value = b.extraPersonPricePerDay !== undefined ? b.extraPersonPricePerDay : (b.perDayPrice || 1200);
        }

        const advanceElem = document.getElementById('cust-advance');
        const baseAdv = b.initialAdv || 0;
        advanceElem.value = baseAdv;
        advanceElem.setAttribute('data-initial-adv', baseAdv);
        if (b.clearedDue) {
          document.getElementById('cust-clear-bill').value = b.clearedDue;
        }
        calculateModalBilling();
      } else {
        document.getElementById('modal-title').innerText = 'Add New Booking';
        document.getElementById('modal-booking-id').value = '';
        
        setInputEnabled(document.getElementById('cust-checkin-date'), true);
        setInputEnabled(document.getElementById('cust-checkin-time'), true);
        setInputEnabled(document.getElementById('cust-checkout-date'), true);
        setInputEnabled(document.getElementById('cust-checkout-time'), true);
        
        const todayDt = new Date();
        const utcMs = todayDt.getTime();
        const istDate = new Date(utcMs + (330 * 60000));
        
        const yyyy = istDate.getUTCFullYear();
        const mm = String(istDate.getUTCMonth() + 1).padStart(2, '0');
        const dd = String(istDate.getUTCDate()).padStart(2, '0');
        const todayStr = `${yyyy}-${mm}-${dd}`;
        const tomorrowDt = new Date(todayDt);
        tomorrowDt.setDate(tomorrowDt.getDate() + 1);
        const t_utcMs = tomorrowDt.getTime();
        const t_istDate = new Date(t_utcMs + (330 * 60000));
        const t_yyyy = t_istDate.getUTCFullYear();
        const t_mm = String(t_istDate.getUTCMonth() + 1).padStart(2, '0');
        const t_dd = String(t_istDate.getUTCDate()).padStart(2, '0');
        const tomorrowStr = `${t_yyyy}-${t_mm}-${t_dd}`;
        const checkInElem = document.getElementById('cust-checkin-date');
        checkInElem.min = todayStr;
        checkInElem.value = todayStr;
        const checkOutElem = document.getElementById('cust-checkout-date');
        checkOutElem.min = todayStr;
        checkOutElem.value = tomorrowStr;
        populateRoomDropdown(state.roomsCapacity.length > 0 ? [state.roomsCapacity[0].roomNo] : []);
        document.getElementById('cust-country-code').value = "+91";
        document.getElementById('cust-checkin-time').value = "11:00";
        document.getElementById('cust-checkout-time').value = "11:00";
        if (extraPersonsInput) extraPersonsInput.value = 0;
        
        if (extraPersonDateInput) extraPersonDateInput.value = "";
        
        if (extraPersonTimeInput) extraPersonTimeInput.value = "11:00";
        if (extraPersonOutDateInput) extraPersonOutDateInput.value = "";
        if (extraPersonOutTimeInput) extraPersonOutTimeInput.value = "11:00";
        extChkBox.checked = false;
        extChkBox.disabled = true;
        
        if (timerNotice) {
          timerNotice.innerText = "(Selectable only for Live Booking)";
          timerNotice.classList.remove('hidden');
          timerNotice.classList.add('text-rose-600');
        }
        toggleExtendedCheckoutFields(false);
        if (mealsChkBox) {
          mealsChkBox.checked = true;
          mealsChkBox.disabled = false;
        }
        document.getElementById('cust-price').value = 1200;
        const extraPriceElem = document.getElementById('cust-extra-price');
        if (extraPriceElem) extraPriceElem.value = 1200;
        
        const advanceElem = document.getElementById('cust-advance');
        advanceElem.value = 0;
        advanceElem.setAttribute('data-initial-adv', 0);
        calculateModalBilling();
      }
      document.getElementById('booking-modal').classList.remove('hidden');
    }
    function setSectionEditability(sectionId, isEditable) {
      const container = document.getElementById(sectionId);
      if (!container) return;
      const inputs = container.querySelectorAll('input, select, button');
      inputs.forEach(el => {
        el.disabled = !isEditable;
        if (!isEditable) {
          el.classList.add('bg-slate-100', 'cursor-not-allowed', 'text-slate-500');
        } else {
          el.classList.remove('bg-slate-100', 'cursor-not-allowed', 'text-slate-500');
        }
      });
      if (!isEditable) {
        container.classList.add('opacity-75', 'bg-slate-100/60');
      } else {
        container.classList.remove('opacity-75', 'bg-slate-100/60');
      }
    }
    function closeBookingModal() {
      document.getElementById('booking-modal').classList.add('hidden');
    }
    function handleExtraPersonDatesChange() {
      const mainInDate = document.getElementById('cust-checkin-date')?.value;
      const mainInTime = document.getElementById('cust-checkin-time')?.value || '12:00';
      const mainOutDate = document.getElementById('cust-checkout-date')?.value;
      const mainOutTime = document.getElementById('cust-checkout-time')?.value || '11:00';
      const hasExt = document.getElementById('cust-has-extended-checkout')?.checked;
      
      let latestOutD = mainOutDate;
      let latestOutT = mainOutTime;
      if (hasExt) {
        const extD = document.getElementById('cust-ext-checkout-date')?.value;
        const extT = document.getElementById('cust-ext-checkout-time')?.value;
        if (extD) {
          latestOutD = extD;
          latestOutT = extT || '12:00';
        }
      }
      
      const epInDateElem = document.getElementById('cust-extra-person-date');
      const epInTimeElem = document.getElementById('cust-extra-person-time');
      const epOutDateElem = document.getElementById('cust-extra-person-out-date');
      const epOutTimeElem = document.getElementById('cust-extra-person-out-time');
      if (epInDateElem && epInDateElem.value && mainInDate) {
        const epInFull = new Date(`${epInDateElem.value}T${epInTimeElem.value || '12:00'}:00+05:30`);
        const mainInFull = new Date(`${mainInDate}T${mainInTime}:00+05:30`);
        if (epInFull < mainInFull) {
          alert(`⚠️ Additional person check-in cannot be earlier than the main check-in (${formatDateTime(`${mainInDate}T${mainInTime}`)}). Please correct it.`);
        }
      }
      if (epOutDateElem && epOutDateElem.value && latestOutD) {
        const epOutFull = new Date(`${epOutDateElem.value}T${epOutTimeElem.value || '11:00'}:00+05:30`);
        const mainOutFull = new Date(`${latestOutD}T${latestOutT}:00+05:30`);
        if (epOutFull > mainOutFull) {
          alert(`⚠️ Additional person check-out date cannot be later than the main/extended check-out date (${formatDateTime(`${latestOutD}T${latestOutT}`)}). Please correct it.`);
        }
      }
      
      calculateModalBilling();
    }
    function handleStayDatesChange() {
      const inDateInput = document.getElementById('cust-checkin-date');
      const outDateInput = document.getElementById('cust-checkout-date');
      const extDateInput = document.getElementById('cust-ext-checkout-date');
      if (inDateInput && outDateInput) {
        outDateInput.min = inDateInput.value;
        if (outDateInput.value && outDateInput.value < inDateInput.value) {
          outDateInput.value = inDateInput.value;
        }
      }
      if (outDateInput && extDateInput) {
        extDateInput.min = outDateInput.value;
        if (extDateInput.value && extDateInput.value < outDateInput.value) {
          alert("⚠️ Extended Check-Out date cannot be prior to the initial Check-Out date!");
          extDateInput.value = outDateInput.value;
        }
      }
      handleExtraPersonDatesChange(); 
    }
    function handleClearBillPayment(clearAmountVal) {
      const clearVal = parseFloat(clearAmountVal) || 0;
      const total = parseFloat(document.getElementById('cust-total').value) || 0;
      
      const advanceElem = document.getElementById('cust-advance');
      let initialAdvance = parseFloat(advanceElem.getAttribute('data-initial-adv'));
      if (isNaN(initialAdvance)) {
        initialAdvance = parseFloat(advanceElem.value) || 0;
        advanceElem.setAttribute('data-initial-adv', initialAdvance);
      }
      if (clearVal + initialAdvance > total) {
        alert("⚠️ Payment amount exceeds the remaining balance due!");
        document.getElementById('cust-clear-bill').value = 0;
        document.getElementById('cust-due').value = Math.max(0, total - initialAdvance);
        return;
      }
      const due = Math.max(0, total - initialAdvance - clearVal);
      document.getElementById('cust-due').value = due;
    }

    function calculateModalBilling() {
      const inDate = document.getElementById('cust-checkin-date')?.value;
      const inTime = document.getElementById('cust-checkin-time')?.value || '11:00';
      const outDate = document.getElementById('cust-checkout-date')?.value;
      const outTime = document.getElementById('cust-checkout-time')?.value || '11:00';
      const hasExt = document.getElementById('cust-has-extended-checkout')?.checked;
      
      let effectiveOutDate = outDate;
      let effectiveOutTime = outTime;
      if (hasExt) {
        const extDate = document.getElementById('cust-ext-checkout-date')?.value;
        const extTime = document.getElementById('cust-ext-checkout-time')?.value || '11:00';
        if (extDate) effectiveOutDate = extDate;
        if (extTime) effectiveOutTime = extTime;
      }

      let noOfDays = 0;
      if (inDate && effectiveOutDate) {
        const start = new Date(`${inDate}T${inTime}:00+05:30`).getTime();
        const end = new Date(`${effectiveOutDate}T${effectiveOutTime}:00+05:30`).getTime();
        if (end > start) {
          const diffHours = (end - start) / (1000 * 60 * 60);
          noOfDays = Math.max(1, Math.ceil(diffHours / 24));
        }
      }
      document.getElementById('cust-days').value = noOfDays;

      const perDayPrice = parseFloat(document.getElementById('cust-price')?.value) || 0;
      const roomCapacity = parseInt(document.getElementById('cust-capacity')?.value) || 1;
      const mainGuestTotalRate = noOfDays * perDayPrice * roomCapacity;

      // Calculate Extra Guest Details
      const extraPersons = parseInt(document.getElementById('cust-extra-persons')?.value) || 0;
      const extraPriceInput = document.getElementById('cust-extra-price');
      let extraPersonPricePerDay = extraPriceInput ? parseFloat(extraPriceInput.value) : perDayPrice;
      if (isNaN(extraPersonPricePerDay) || extraPersonPricePerDay <= 0) {
        extraPersonPricePerDay = perDayPrice;
        if (extraPriceInput) extraPriceInput.value = extraPersonPricePerDay;
      }

      const epInDate = document.getElementById('cust-extra-person-date')?.value || inDate;
      const epInTime = document.getElementById('cust-extra-person-time')?.value || inTime;
      const epOutDate = document.getElementById('cust-extra-person-out-date')?.value || effectiveOutDate;
      const epOutTime = document.getElementById('cust-extra-person-out-time')?.value || effectiveOutTime;

      let extraPersonDays = 0;
      if (extraPersons > 0 && epInDate && epOutDate) {
        const epStart = new Date(`${epInDate}T${epInTime}:00+05:30`).getTime();
        const epEnd = new Date(`${epOutDate}T${epOutTime}:00+05:30`).getTime();
        if (epEnd > epStart) {
          const diffHours = (epEnd - epStart) / (1000 * 60 * 60);
          extraPersonDays = Math.max(1, Math.ceil(diffHours / 24));
        }
      }
      
      const extraPersonTotalRate = extraPersons * extraPersonDays * extraPersonPricePerDay;

      const epDaysElem = document.getElementById('cust-extra-days');
      if (epDaysElem) epDaysElem.value = extraPersonDays;
      
      const epTotalElem = document.getElementById('cust-extra-total');
      if (epTotalElem) epTotalElem.value = extraPersonTotalRate;

      // Calculate Extra Food Total
      let extraFoodTotal = 0;
      document.querySelectorAll('#food-orders-container .cust-food-charge').forEach(input => {
        extraFoodTotal += parseFloat(input.value) || 0;
      });

      // Calculate Cab Fare Total
      let cabFareTotal = 0;
      document.querySelectorAll('#cab-trips-container .cust-cab-rate').forEach(input => {
        cabFareTotal += parseFloat(input.value) || 0;
      });

      const grandTotal = mainGuestTotalRate + extraPersonTotalRate + extraFoodTotal + cabFareTotal;
      document.getElementById('cust-total').value = grandTotal;

      const advanceElem = document.getElementById('cust-advance');
      let initialAdv = parseFloat(advanceElem.getAttribute('data-initial-adv'));
      if (isNaN(initialAdv)) {
        initialAdv = parseFloat(advanceElem.value) || 0;
      }
      const clearBill = parseFloat(document.getElementById('cust-clear-bill')?.value) || 0;

      const totalDue = Math.max(0, grandTotal - initialAdv - clearBill);
      document.getElementById('cust-due').value = totalDue;
    }

    async function saveBookingFromModal(e) {
      if (e) e.preventDefault();
      
      const bookingId = document.getElementById('modal-booking-id').value;
      const name = document.getElementById('cust-name').value.trim();
      const contactNo = document.getElementById('cust-contact').value.trim();
      
      if (!name) {
        alert("⚠️ Please enter Guest Name.");
        return;
      }
      
      const selectedRooms = getSelectedRooms();
      if (selectedRooms.length === 0) {
        alert("⚠️ Please select at least one Room No.");
        return;
      }

      const inDate = document.getElementById('cust-checkin-date').value;
      const inTime = document.getElementById('cust-checkin-time').value || '11:00';
      const outDate = document.getElementById('cust-checkout-date').value;
      const outTime = document.getElementById('cust-checkout-time').value || '11:00';

      if (!inDate || !outDate) {
        alert("⚠️ Please select Check-In and Check-Out dates.");
        return;
      }

      const checkInISO = `${inDate}T${inTime}:00+05:30`;
      const checkOutISO = `${outDate}T${outTime}:00+05:30`;

      const hasExt = document.getElementById('cust-has-extended-checkout').checked;
      let extISO = "";
      if (hasExt) {
        const extD = document.getElementById('cust-ext-checkout-date').value;
        const extT = document.getElementById('cust-ext-checkout-time').value || '11:00';
        if (extD) extISO = `${extD}T${extT}:00+05:30`;
      }

      const extraPersons = parseInt(document.getElementById('cust-extra-persons').value) || 0;
      let epJoinedISO = "", epOutISO = "";
      if (extraPersons > 0) {
        const epInD = document.getElementById('cust-extra-person-date').value || inDate;
        const epInT = document.getElementById('cust-extra-person-time').value || inTime;
        const epOutD = document.getElementById('cust-extra-person-out-date').value || (extISO ? extISO.split('T')[0] : outDate);
        const epOutT = document.getElementById('cust-extra-person-out-time').value || (extISO ? extISO.split('T')[1].substring(0,5) : outTime);
        
        epJoinedISO = `${epInD}T${epInT}:00+05:30`;
        epOutISO = `${epOutD}T${epOutT}:00+05:30`;
      }

      // Collect food order details
      const foodOrders = [];
      let extraFoodTotal = 0;
      document.querySelectorAll('#food-orders-container .food-order-row').forEach(row => {
        const desc = row.querySelector('.cust-food-desc').value.trim();
        const fDate = row.querySelector('.cust-food-date').value;
        const fTime = row.querySelector('.cust-food-time').value || '00:00';
        const price = parseFloat(row.querySelector('.cust-food-price').value) || 0;
        const plates = parseInt(row.querySelector('.cust-food-plates').value) || 1;
        const charge = parseFloat(row.querySelector('.cust-food-charge').value) || (price * plates);
        if (desc || charge > 0) {
          foodOrders.push({
            foodDesc: desc,
            foodDateTime: fDate ? `${fDate}T${fTime}:00+05:30` : '',
            itemPrice: price,
            plates: plates,
            foodCharge: charge
          });
          extraFoodTotal += charge;
        }
      });

      // Collect cab trip details
      const cabTrips = [];
      let cabFareTotal = 0;
      document.querySelectorAll('#cab-trips-container .cab-trip-row').forEach((row, idx) => {
        const cDate = row.querySelector('.cust-cab-date').value;
        const cTime = row.querySelector('.cust-cab-time').value || '12:00';
        const rate = parseFloat(row.querySelector('.cust-cab-rate').value) || 0;
        const remark = row.querySelector('.cust-cab-remark').value.trim();
        if (rate > 0 || remark) {
          cabTrips.push({
            tripName: `Trip ${idx + 1}`,
            dateTime: cDate ? `${cDate}T${cTime}:00+05:30` : '',
            rate: rate,
            remark: remark
          });
          cabFareTotal += rate;
        }
      });

      const perDayPrice = parseFloat(document.getElementById('cust-price').value) || 0;
      const extraPriceInput = document.getElementById('cust-extra-price');
      const extraPersonPricePerDay = extraPriceInput ? (parseFloat(extraPriceInput.value) || perDayPrice) : perDayPrice;
      
      const noOfDays = parseInt(document.getElementById('cust-days').value) || 0;
      const extraPersonDays = parseInt(document.getElementById('cust-extra-days')?.value) || 0;
      const extraPersonTotalRate = parseFloat(document.getElementById('cust-extra-total')?.value) || (extraPersons * extraPersonDays * extraPersonPricePerDay);
      
      const capacity = parseInt(document.getElementById('cust-capacity').value) || 1;
      const mainGuestTotalRate = noOfDays * perDayPrice * capacity;
      const totalAmount = parseFloat(document.getElementById('cust-total').value) || 0;
      
      const advanceElem = document.getElementById('cust-advance');
      let initialAdv = parseFloat(advanceElem.getAttribute('data-initial-adv'));
      if (isNaN(initialAdv)) {
        initialAdv = parseFloat(advanceElem.value) || 0;
      }
      
      const clearedDue = parseFloat(document.getElementById('cust-clear-bill').value) || 0;
      const totalDue = parseFloat(document.getElementById('cust-due').value) || 0;

      let bookingRecord = {};
      if (bookingId) {
        const existingIndex = state.bookings.findIndex(b => String(b.id) === String(bookingId));
        if (existingIndex !== -1) {
          bookingRecord = state.bookings[existingIndex];
        }
      } else {
        const ids = generateIDsForYear(inDate);
        bookingRecord = {
          id: 'B-' + Date.now(),
          bookingCode: ids.bookingCode,
          invoiceNo: ids.invoiceNo,
          inactive: false
        };
      }

      // Explicit Mapping of All Portal Fields for Storage & Cloud Sync
      bookingRecord.name = name;
      bookingRecord.address = document.getElementById('cust-address').value.trim();
      bookingRecord.city = document.getElementById('cust-city').value.trim();
      bookingRecord.state = document.getElementById('cust-state').value.trim();
      bookingRecord.country = document.getElementById('cust-country').value.trim();
      bookingRecord.zipCode = document.getElementById('cust-zip').value.trim();
      bookingRecord.idNo = document.getElementById('cust-id').value.trim();
      bookingRecord.countryCode = document.getElementById('cust-country-code').value;
      bookingRecord.contactNo = contactNo;
      bookingRecord.idProofBase64 = document.getElementById('cust-id-file-base64').value;
      bookingRecord.idProofFileName = document.getElementById('cust-id-file-name').value;
      bookingRecord.roomNo = selectedRooms.join(', ');
      bookingRecord.agentInfo = document.getElementById('cust-agent').value;
      bookingRecord.capacity = capacity;
      bookingRecord.mainGuestCount = capacity;
      bookingRecord.extraPersons = extraPersons;
      bookingRecord.extraPersonJoined = epJoinedISO;
      bookingRecord.extraPersonOut = epOutISO;
      bookingRecord.checkIn = checkInISO;
      bookingRecord.checkOut = checkOutISO;
      bookingRecord.hasExtendedCheckout = hasExt;
      bookingRecord.extendedCheckOut = extISO;
      bookingRecord.includeMeals = document.getElementById('cust-include-meals').checked;
      bookingRecord.foodOrders = JSON.stringify(foodOrders);
      bookingRecord.cabTrips = JSON.stringify(cabTrips);
      bookingRecord.noOfDays = noOfDays;
      bookingRecord.perDayPrice = perDayPrice;
      bookingRecord.mainGuestTotalRate = mainGuestTotalRate;
      bookingRecord.extraPersonDays = extraPersonDays;
      bookingRecord.extraPersonPricePerDay = extraPersonPricePerDay;
      bookingRecord.extraPersonTotalRate = extraPersonTotalRate;
      bookingRecord.extraFoodTotal = extraFoodTotal;
      bookingRecord.cabFareTotal = cabFareTotal;
      bookingRecord.totalAmount = totalAmount;
      bookingRecord.initialAdv = initialAdv;
      bookingRecord.clearedDue = clearedDue;
      bookingRecord.totalDue = totalDue;

      if (!bookingId) {
        state.bookings.push(bookingRecord);
      }

      closeBookingModal();
      refreshAllUI();
      checkUpcomingCheckoutsWithDue();
      await saveChanges(false, false);
    }

    function deleteBooking(bookingId) {
      openMasterDeleteModal('booking', bookingId);
    }

    function renderBookingsTable(filterDate = '') {
      const tbody = document.getElementById('bookings-tbody');
      if (!tbody) return;
      tbody.innerHTML = '';

      let list = state.bookings.filter(b => !isInactiveBooking(b));

      if (filterDate) {
        list = list.filter(b => {
          if (!b.checkIn) return false;
          const cInDate = b.checkIn.replace(' ', 'T').split('T')[0];
          const effOut = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
          const cOutDate = effOut ? effOut.replace(' ', 'T').split('T')[0] : cInDate;
          return filterDate >= cInDate && filterDate <= cOutDate;
        });
      }

      if (list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="text-center py-6 text-slate-400">No active bookings found.</td></tr>`;
        return;
      }

      const now = new Date().getTime();

      list.forEach(b => {
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100 text-[11px]";
        
        const cIn = parseDateMs(b.checkIn);
        const cOut = getEffectiveCheckoutTime(b);
        let statusBadge = '<span class="bg-blue-50 text-blue-700 px-2 py-0.5 rounded-full font-bold">Upcoming</span>';
        if (now > cOut) {
          statusBadge = '<span class="bg-emerald-50 text-emerald-700 px-2 py-0.5 rounded-full font-bold">Closed</span>';
        } else if (now >= cIn && now <= cOut) {
          statusBadge = '<span class="bg-amber-100 text-amber-800 px-2 py-0.5 rounded-full font-bold">Live</span>';
        }

        const roomsDisplay = getBookingRooms(b).join(', ');
        const effectiveOut = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;

        tr.innerHTML = `
          <td class="py-2.5 px-3 font-mono font-bold text-blue-600">${b.bookingCode || 'N/A'}</td>
          <td class="py-2.5 px-3 font-bold text-slate-800">${formatTitleCase(b.name)}</td>
          <td class="py-2.5 px-3"><span class="bg-slate-100 text-slate-800 font-bold px-2 py-0.5 rounded-full">Room ${roomsDisplay}</span></td>
          <td class="py-2.5 px-3 text-slate-600">${formatDateTime(b.checkIn)}</td>
          <td class="py-2.5 px-3 text-slate-600">${formatDateTime(effectiveOut)}</td>
          <td class="py-2.5 px-3 font-semibold text-slate-800">₹${(b.totalAmount || 0).toLocaleString('en-IN')}</td>
          <td class="py-2.5 px-3 font-bold ${b.totalDue > 0 ? 'text-rose-600' : 'text-emerald-600'}">₹${(b.totalDue || 0).toLocaleString('en-IN')}</td>
          <td class="py-2.5 px-3">${statusBadge}</td>
          <td class="py-2.5 px-3 text-center">
            <div class="flex items-center justify-center gap-1.5">
              <button onclick="openBookingModal('${b.id}')" class="bg-blue-50 hover:bg-blue-100 text-blue-600 p-1.5 rounded-xl transition" title="Edit Booking">
                <i class="fa-solid fa-pen-to-square"></i>
              </button>
              <button onclick="printInvoice('${b.id}')" class="bg-emerald-50 hover:bg-emerald-100 text-emerald-600 p-1.5 rounded-xl transition" title="Print Invoice">
                <i class="fa-solid fa-receipt"></i>
              </button>
            </div>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function renderRoomCapacityTable() {
      const tbody = document.getElementById('room-capacity-tbody');
      if (!tbody) return;
      tbody.innerHTML = '';
      if (!state.roomsCapacity || state.roomsCapacity.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" class="text-center py-4 text-slate-400">No Room Capacity records found.</td></tr>`;
        return;
      }
      state.roomsCapacity.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.className = "border-b border-slate-100 hover:bg-slate-50 transition text-[11px]";
        tr.innerHTML = `
          <td class="py-2 px-3 font-bold text-slate-800">Room ${r.roomNo}</td>
          <td class="py-2 px-3 text-slate-600 font-semibold">${r.capacity} Persons</td>
          <td class="py-2 px-3 text-center">
            <button onclick="openMasterDeleteModal('room', ${idx})" class="text-rose-600 hover:text-rose-800 p-1">
              <i class="fa-solid fa-trash-can"></i>
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function addRoomCapacityRow() {
      const roomNo = prompt("Enter Room Number:");
      const cap = prompt("Enter Max Capacity (Persons):");
      if (roomNo && cap) {
        state.roomsCapacity.push({ roomNo: parseInt(roomNo), capacity: parseInt(cap) });
        renderRoomCapacityTable();
        populateRoomDropdown();
        saveChanges(false, false);
      }
    }

    function renderMasterAgentTable() {
      const tbody = document.getElementById('master-agent-tbody');
      if (!tbody) return;
      tbody.innerHTML = '';
      if (!state.masterAgents || state.masterAgents.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="text-center py-4 text-slate-400">No Agent records found.</td></tr>`;
        return;
      }
      state.masterAgents.forEach((a, idx) => {
        const tr = document.createElement('tr');
        tr.className = "border-b border-slate-100 hover:bg-slate-50 transition text-[11px]";
        tr.innerHTML = `
          <td class="py-2 px-3 font-bold text-slate-800">${a.agentName}</td>
          <td class="py-2 px-3 text-slate-600">${a.phone}</td>
          <td class="py-2 px-3 text-slate-600">${a.roomNo}</td>
          <td class="py-2 px-3 text-center">
            <button onclick="openMasterDeleteModal('agent', ${idx})" class="text-rose-600 hover:text-rose-800 p-1">
              <i class="fa-solid fa-trash-can"></i>
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function addMasterAgentRow() {
      const name = prompt("Enter Agent Name:");
      const phone = prompt("Enter Contact Number:");
      const rooms = prompt("Enter Assigned Rooms (or 'All Rooms'):", "All Rooms");
      if (name) {
        state.masterAgents.push({ agentName: name, phone: phone || 'N/A', roomNo: rooms || 'All Rooms' });
        renderMasterAgentTable();
        populateAgentDropdown();
        saveChanges(false, false);
      }
    }

    function renderCalendar(year) {
      const calContainer = document.getElementById('calendar-container');
      if (!calContainer) return;
      calContainer.innerHTML = '';
      
      const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
      
      months.forEach((monthName, mIdx) => {
        const monthBox = document.createElement('div');
        monthBox.className = "bg-white border border-slate-200/80 rounded-2xl p-3 shadow-xs";
        
        let monthHeader = `<h3 class="font-bold text-slate-800 text-[12px] mb-2 text-center">${monthName} ${year}</h3>`;
        let daysHeader = `<div class="grid grid-cols-7 gap-1 text-center font-bold text-[9px] text-slate-400 mb-1">
          <div>Su</div><div>Mo</div><div>Tu</div><div>We</div><div>Th</div><div>Fr</div><div>Sa</div>
        </div>`;
        
        let daysGrid = `<div class="grid grid-cols-7 gap-1 text-center text-[10px]">`;
        const firstDay = new Date(year, mIdx, 1).getDay();
        const daysInMonth = new Date(year, mIdx + 1, 0).getDate();

        for (let i = 0; i < firstDay; i++) {
          daysGrid += `<div></div>`;
        }

        for (let day = 1; day <= daysInMonth; day++) {
          const dateStr = `${year}-${String(mIdx + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
          
          let dayClass = "p-1 rounded-xl cursor-pointer hover:bg-blue-50 text-slate-700 font-medium";
          const activeBookingsOnDate = state.bookings.filter(b => {
            if (isInactiveBooking(b) || !b.checkIn) return false;
            const cIn = b.checkIn.split('T')[0];
            const effOut = (isTrue(b.hasExtendedCheckout) && b.extendedCheckOut) ? b.extendedCheckOut : b.checkOut;
            const cOut = effOut ? effOut.split('T')[0] : cIn;
            return dateStr >= cIn && dateStr <= cOut;
          });

          if (activeBookingsOnDate.length > 0) {
            dayClass = "p-1 rounded-xl cursor-pointer bg-amber-500 text-white font-bold shadow-xs";
          }

          daysGrid += `<div class="${dayClass}" onclick="openCalendarDateBookings('${dateStr}')">${day}</div>`;
        }

        daysGrid += `</div>`;
        monthBox.innerHTML = monthHeader + daysHeader + daysGrid;
        calContainer.appendChild(monthBox);
      });
    }

    function openCalendarDateBookings(dateStr) {
      document.getElementById('booking-date-search').value = dateStr;
      switchTab('booking');
      renderBookingsTable(dateStr);
    }

    function handleCalendarYearChange(val) {
      const year = parseInt(val) || defaultAppYear;
      state.selectedYear = year;
      renderCalendar(year);
    }

    function closeCommentBox() {
      // Helper function for UI reset
    }
</script>
</body>
</html>
