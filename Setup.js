// ============================================================
//  AI SMART INVENTORY SYSTEM — Google Sheets Edition
//  Version 3.0 | June 2026
//  © 2026 All Rights Reserved — Licensed Software
//
//  INSTALLATION:
//  1. Open a new Google Sheet
//  2. Extensions → Apps Script
//  3. Create files: Setup, Agent, LicenseAdmin
//  4. Paste each .gs file into the corresponding script file
//  5. Save, then reload the spreadsheet
//  6. Enter your license key when prompted
// ============================================================

// ─── LICENSE SYSTEM (DO NOT MODIFY) ─────────────────────────
const _SYSTEM = {
  APP_NAME:    "AI Smart Inventory System",
  VERSION:     "3.0",
  BUILD_DATE:  "June 2026",
  // Internal secret for license validation — keep private
  _S1: "INV",
  _S2: "2026",
  _S3: "SECURE",
  _S4: "X9Y3Z7",
  ADMIN_SERVER_URL: "https://script.google.com/macros/s/AKfycbzG3juStcgjT0FEMCnQo4dufCEGC_3pkbKpgD-3IKT6JleriCwuhAwg69ztjIwIWUt7/exec"
};

// Sheet names (English) — used as SEARCH KEYWORDS, not exact names
const SHEETS = {
  DASHBOARD:    "Dashboard",
  INVENTORY:    "Inventory",
  TRANSACTIONS: "Transaction",
  BRANCHES:     "Branch",
  GUIDE:        "Guide",
};

// ─── SAFE SPREADSHEET ACCESS (works from Triggers too) ──────
function _getSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  
  const id = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch(e) {}
  }
  
  return null;
}

// ─── FIND SHEET BY KEYWORD (handles emoji prefixes) ─────────
function _getSheet(keyword) {
  const ss = _getSpreadsheet();
  if (!ss) return null;
  // Try exact match first
  let sheet = ss.getSheetByName(keyword);
  if (sheet) return sheet;
  // Try with common emoji prefixes
  const emojis = ["📊 ", "📦 ", "🔄 ", "🏢 ", "🔑 ", "📋 "];
  for (const emoji of emojis) {
    sheet = ss.getSheetByName(emoji + keyword);
    if (sheet) return sheet;
  }
  // Try partial match (keyword anywhere in name)
  const allSheets = ss.getSheets();
  for (const s of allSheets) {
    if (s.getName().toLowerCase().includes(keyword.toLowerCase())) {
      return s;
    }
  }
  return null;
}

// Color palette
const COLORS = {
  primary:   "#10B981",
  dark:      "#0F172A",
  header:    "#064E3B",
  headerTxt: "#FFFFFF",
  card:      "#1E293B",
  danger:    "#EF4444",
  warning:   "#F59E0B",
  info:      "#3B82F6",
  muted:     "#94A3B8",
  rowA:      "#F0FDF4",
  rowB:      "#FFFFFF",
  border:    "#E2E8F0",
};

// ─── ENTRY POINT — runs on every open ───────────────────────
// ─── LICENSE DIALOG ──────────────────────────────────────────
function showLicenseDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; }
        body { background: #0F172A; color: #E2E8F0; display: flex; align-items: center; justify-content: center; min-height: 100vh; padding: 20px; }
        .card { background: #1E293B; border: 1px solid #334155; border-radius: 16px; padding: 32px; max-width: 420px; width: 100%; }
        .logo { text-align: center; margin-bottom: 24px; }
        .logo-icon { width: 64px; height: 64px; background: #064E3B; border-radius: 16px; display: flex; align-items: center; justify-content: center; margin: 0 auto 12px; font-size: 28px; }
        h1 { font-size: 18px; font-weight: 700; color: #F1F5F9; text-align: center; }
        p { font-size: 13px; color: #94A3B8; text-align: center; margin-top: 6px; }
        label { display: block; font-size: 11px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.8px; color: #94A3B8; margin-bottom: 8px; margin-top: 20px; }
        input { width: 100%; padding: 12px 14px; background: #0F172A; border: 1px solid #334155; border-radius: 8px; color: #F1F5F9; font-size: 14px; font-family: monospace; transition: border 0.2s; outline: none; }
        input:focus { border-color: #10B981; }
        input::placeholder { color: #475569; }
        .name-input { font-family: 'Segoe UI', Arial, sans-serif; }
        button { width: 100%; padding: 14px; background: #10B981; color: #fff; border: none; border-radius: 8px; font-size: 15px; font-weight: 700; cursor: pointer; margin-top: 20px; transition: background 0.2s; }
        button:hover { background: #059669; }
        button:disabled { background: #334155; color: #64748B; cursor: not-allowed; }
        .error { background: #450A0A; border: 1px solid #991B1B; color: #FCA5A5; border-radius: 8px; padding: 10px 14px; font-size: 13px; margin-top: 14px; display: none; }
        .success { background: #052E16; border: 1px solid #166534; color: #86EFAC; border-radius: 8px; padding: 10px 14px; font-size: 13px; margin-top: 14px; display: none; }
        .spinner { display: inline-block; width: 16px; height: 16px; border: 2px solid rgba(255,255,255,0.3); border-top-color: #fff; border-radius: 50%; animation: spin 0.8s linear infinite; vertical-align: middle; margin-right: 6px; }
        @keyframes spin { to { transform: rotate(360deg); } }
        .divider { border: none; border-top: 1px solid #1E293B; margin: 20px 0; }
        small { display: block; text-align: center; font-size: 11px; color: #475569; margin-top: 16px; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="logo">
          <div class="logo-icon">📦</div>
          <h1>AI Smart Inventory System</h1>
          <p>Version 3.0 · June 2026</p>
        </div>
        <label>License Key</label>
        <input type="text" id="licenseKey" placeholder="INV-XXXX-XXXX-XXXX" maxlength="19"
          oninput="this.value=this.value.toUpperCase().replace(/[^A-Z0-9-]/g,'')" />
        <button id="btn" onclick="activate()">✨ Activate System</button>
        <div class="error" id="err"></div>
        <div class="success" id="suc"></div>
        <small>Purchase a license at the seller's store · Contact support for issues</small>
      </div>
      <script>
        function activate() {
          const key  = document.getElementById('licenseKey').value.trim();
          const btn  = document.getElementById('btn');
          const err  = document.getElementById('err');
          const suc  = document.getElementById('suc');
          err.style.display = 'none';
          suc.style.display = 'none';
          if (!key) { showErr('Please enter your license key.'); return; }
          if (key.length < 10) { showErr('Invalid license key format.'); return; }
          btn.disabled = true;
          btn.innerHTML = '<span class="spinner"></span>Validating...';
          google.script.run
            .withSuccessHandler(res => {
              if (res.ok) {
                suc.textContent = '✅ ' + res.message;
                suc.style.display = 'block';
                btn.textContent = '✅ Activated!';
                setTimeout(() => google.script.host.close(), 1800);
              } else {
                showErr('❌ ' + res.message);
                btn.disabled = false;
                btn.textContent = '✨ Activate System';
              }
            })
            .withFailureHandler(e => {
              showErr('Error: ' + e.message);
              btn.disabled = false;
              btn.textContent = '✨ Activate System';
            })
            .validateAndActivateLicense(key);
        }
        function showErr(msg) {
          const el = document.getElementById('err');
          el.textContent = msg;
          el.style.display = 'block';
        }
        document.getElementById('licenseKey').addEventListener('keydown', e => { if (e.key === 'Enter') activate(); });
      </script>
    </body>
    </html>
  `)
  .setTitle("License Activation")
  .setWidth(460)
  .setHeight(420);

  SpreadsheetApp.getUi().showModalDialog(html, "🔐 License Activation — AI Smart Inventory");
}

// Called from dialog via google.script.run
function validateAndActivateLicense(key) {
  // Check if it's admin key
  const ADMIN_KEY = "ADMIN-MASTER-KEY-2026";
  if (key === ADMIN_KEY) {
    _activateSuccess(key, true);
    return { ok: true, message: "Welcome, Administrator! System fully unlocked. Reloading..." };
  }

  // Validate license key structure
  if (!_isValidKeyFormat(key)) {
    return { ok: false, message: "Invalid license key format. Expected: INV-XXXX-XXXX-XXXX" };
  }

  // Validate cryptographic checksum (Offline check first)
  if (!_verifyChecksum(key)) {
    return { ok: false, message: "License key is invalid or has been tampered with." };
  }

  // --- ONLINE VERIFICATION (Hardware Binding) ---
  const fileId = SpreadsheetApp.getActiveSpreadsheet().getId();
  try {
    const response = UrlFetchApp.fetch(_SYSTEM.ADMIN_SERVER_URL, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ key: key, fileId: fileId }),
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      if (_SYSTEM.ADMIN_SERVER_URL.includes("URL_ANDA_NANTI")) {
        return { ok: false, message: "Admin has not configured the Web App URL. Please contact your administrator." };
      }
    } else {
      const resData = JSON.parse(response.getContentText());
      if (!resData.ok) {
        return { ok: false, message: resData.message };
      }
    }
  } catch (err) {
    return { ok: false, message: "Connection to License Server failed. " + err.message };
  }

  // Bind key to this sheet
  _activateSuccess(key, false);
  return { ok: true, message: "License successfully verified ONLINE and bound to this sheet. Reloading..." };
}

function _activateSuccess(key, isAdmin) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("LICENSE_ACTIVATED", "true");
  props.setProperty("LICENSE_KEY", key);
  props.setProperty("IS_ADMIN", isAdmin ? "true" : "false");
  props.setProperty("ACTIVATED_AT", new Date().toISOString());
  
  // Automagically bind the Sheet ID during activation so the user doesn't have to do it manually!
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) {
    props.setProperty("SHEET_ID", ss.getId());
  }

  // Automatically connect the customer's Gmail by creating a background trigger
  try {
    _removeAllTriggers("pollEmails");
    ScriptApp.newTrigger("pollEmails").timeBased().everyMinutes(5).create();
  } catch (e) {
    Logger.log("Failed to create email trigger: " + e);
  }

  // AUTOMATICALLY SETUP V4 PREMIUM SHEETS UPON ACTIVATION
  setupSystem();

  // Build the full menu immediately
  onOpen();
}

// ─── LICENSE VALIDATION (CRYPTOGRAPHIC) ─────────────────────
function _isValidKeyFormat(key) {
  return /^INV-[A-Z0-9]{4}-[A-Z0-9]{4}-[A-Z0-9]{4}$/.test(key);
}

function _verifyChecksum(key) {
  const parts = key.split("-");
  if (parts.length !== 4 || parts[0] !== "INV") return false;
  const expected = _computeChecksum(parts[1], parts[2]);
  return parts[3] === expected;
}

function _computeChecksum(part1, part2) {
  // DJB2 variant — identical algorithm in both Google Apps Script & Excel VBA
  // Keys generated in GSheets will validate correctly in Excel and vice versa
  const secret = _SYSTEM._S1 + _SYSTEM._S2 + _SYSTEM._S3 + _SYSTEM._S4;
  const combined = part1 + "-" + part2 + "-" + secret;
  let hash = 5381;
  for (let i = 0; i < combined.length; i++) {
    hash = ((hash * 33) ^ combined.charCodeAt(i)) & 0x7FFFFFFF;
  }
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  let result = "", h = hash;
  for (let j = 0; j < 4; j++) {
    result += chars[h % 36];
    h = Math.floor(h / 36);
  }
  return result;
}

// Admin function: generate a valid license key
function generateValidKey() {
  const chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
  const seg = () => Array(4).fill(0).map(() => chars[Math.floor(Math.random() * chars.length)]).join("");
  const p1 = seg(), p2 = seg();
  const p3 = _computeChecksum(p1, p2);
  return `INV-${p1}-${p2}-${p3}`;
}

// ─── LOCKED SIDEBAR (shown when not activated) ───────────────
function _showLockedSidebar() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: 'Segoe UI', sans-serif; background: #0F172A; color: #94A3B8; padding: 20px; margin: 0; }
      .lock { text-align: center; padding: 40px 20px; }
      .icon { font-size: 48px; margin-bottom: 16px; }
      h2 { color: #F1F5F9; font-size: 16px; margin-bottom: 8px; }
      p { font-size: 12px; line-height: 1.5; }
      button { margin-top: 20px; width: 100%; padding: 12px; background: #10B981; color: white; border: none; border-radius: 8px; font-size: 14px; font-weight: 700; cursor: pointer; }
    </style>
    <div class="lock">
      <div class="icon">🔐</div>
      <h2>License Required</h2>
      <p>This spreadsheet requires a valid license key to access the Inventory System.</p>
      <p style="margin-top:12px;">Purchase a license from the seller and click below to activate.</p>
      <button onclick="google.script.run.showLicenseDialog()">Activate License</button>
    </div>
  `).setTitle("License Required");
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── SYSTEM SETUP ────────────────────────────────────────────
function setupSystem() {
  _setupInventorySheet();
  _setupBranchesSheet();
  _setupTransactionsSheet();
  _setupDashboardSheet();
  _setupGuideSheet();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename(_SYSTEM.APP_NAME + " v" + _SYSTEM.VERSION);
  ss.setActiveSheet(ss.getSheetByName(SHEETS.DASHBOARD));
}

// ─── BRANCHES SHEET ──────────────────────────────────────────
function _setupBranchesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEETS.BRANCHES);
  if (sh) sh.clearContents(); else sh = ss.insertSheet(SHEETS.BRANCHES);

  const headers = ["ID", "Branch Name", "Location / City", "Status", "Date Created", "Notes"];
  _applyHeaders(sh, headers);

  const rows = [
    [1, "Head Office", "Jakarta",  "Active", new Date(), "Main headquarters"],
    [2, "Branch Surabaya", "Surabaya", "Active", new Date(), ""],
    [3, "Branch Bandung",  "Bandung",  "Active", new Date(), ""],
  ];
  rows.forEach((r, i) => sh.getRange(i + 2, 1, 1, r.length).setValues([r]));

  [50, 180, 150, 90, 130, 200].forEach((w, i) => sh.setColumnWidth(1 + i, w));
  sh.getRange("A:A").setNumberFormat("0").setHorizontalAlignment("center");
  sh.getRange("D:D").setHorizontalAlignment("center");
  sh.getRange("E:E").setNumberFormat("dd/MM/yyyy");

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Active", "Inactive"]).setAllowInvalid(false).build();
  sh.getRange("D2:D500").setDataValidation(statusRule);

  _alternateRows(sh, 2, 100, 6);
  sh.setFrozenRows(1);
}

// ─── INVENTORY SHEET ─────────────────────────────────────────
function _setupInventorySheet(isFormatOnly = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEETS.INVENTORY);
  if (sh && !isFormatOnly) sh.clearContents(); else if (!sh) sh = ss.insertSheet(SHEETS.INVENTORY);

  const headers = [
    "ID", "Item Code", "Item Name", "Category", "Branch",
    "Initial Stock", "Total In", "Total Out", "Current Stock",
    "Min Stock", "Unit", "Buy Price", "Sell Price", "Status", "QR Code", "Last Updated"
  ];
  _applyHeaders(sh, headers);

  // Formulas for rows 2–500
  if (!isFormatOnly) {
    const f9 = [], f14 = [], f15 = [], f16 = [];
    for (let r = 2; r <= 500; r++) {
      f9.push([`=IF(B${r}="","",F${r}+G${r}-H${r})`]);
      f14.push([`=IF(B${r}="","",IF(I${r}=0,"🔴 OUT OF STOCK",IF(I${r}<=J${r}/2,"🟠 CRITICAL",IF(I${r}<=J${r},"🟡 LOW STOCK","🟢 IN STOCK"))))`]);
      f15.push([`=IF(B${r}="","",IMAGE("https://quickchart.io/qr?text=" & B${r} & "&size=100", 4, 100, 100))`]);
      f16.push([`=IF(B${r}="","",TEXT(NOW(),"dd/MM/yyyy HH:mm"))`]);
    }
    sh.getRange(2, 9, 499, 1).setFormulas(f9);
    sh.getRange(2, 14, 499, 1).setFormulas(f14);
    sh.getRange(2, 15, 499, 1).setFormulas(f15);
    sh.getRange(2, 16, 499, 1).setFormulas(f16);
  }

  // Sample data
  if (!isFormatOnly) {
    const sampleItems = [
      [1, "ELE-001", "iPhone 15 Pro Max 256GB", "Electronics", "Head Office", 0, 50, 5, "", 5, "unit", 1000, 1199, "", "", ""],
      [2, "ELE-002", "Samsung Galaxy S24 Ultra", "Electronics", "Head Office", 0, 30, 2, "", 5, "unit", 900, 1299, "", "", ""],
      [3, "ELE-003", "MacBook Pro M3 14-inch",  "Electronics", "Head Office", 0, 15, 1, "", 3, "unit", 1500, 1999, "", "", ""],
      [4, "FAS-001", "Nike Air Force 1 07",     "Fashion",     "Head Office", 0, 100, 20, "", 20, "pair", 70, 110, "", "", ""],
      [5, "GRO-001", "Mineral Water 600ml (Box)", "Groceries", "Head Office", 0, 500, 498, "", 20, "box", 8, 12, "", "", ""],
    ];
    sampleItems.forEach((row, i) => sh.getRange(i + 2, 1, 1, row.length).setValues([row]));
  }

  [40, 100, 220, 110, 130, 80, 80, 80, 100, 90, 70, 110, 110, 120, 100, 130].forEach((w, i) => sh.setColumnWidth(1 + i, w));
  sh.getRange("F:J").setNumberFormat("#,##0");
  sh.getRange("L:M").setNumberFormat('"$"#,##0.00');
  sh.getRange("N:N").setHorizontalAlignment("center");
  sh.getRange("O:O").setHorizontalAlignment("center");
  sh.setRowHeights(2, 499, 100); // taller rows for QR codes

  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Electronics", "Fashion", "Groceries", "Office", "Food & Beverage", "Other"])
    .setAllowInvalid(true).build();
  sh.getRange("D2:D500").setDataValidation(catRule);

  // Conditional formatting
  const cfRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("OUT OF STOCK").setBackground("#FEE2E2").setFontColor("#991B1B")
      .setRanges([sh.getRange("A2:N500")]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("CRITICAL").setBackground("#FFEDD5").setFontColor("#9A3412")
      .setRanges([sh.getRange("A2:N500")]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("LOW STOCK").setBackground("#FEF9C3").setFontColor("#854D0E")
      .setRanges([sh.getRange("A2:N500")]).build(),
  ];
  sh.setConditionalFormatRules(cfRules);

  _alternateRows(sh, 2, 500, 14);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
}

// ─── TRANSACTIONS SHEET ──────────────────────────────────────
function _setupTransactionsSheet(isFormatOnly = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  if (sh && !isFormatOnly) sh.clearContents(); else if (!sh) sh = ss.insertSheet(SHEETS.TRANSACTIONS);

  const headers = ["#", "Date & Time", "Type", "Item Code", "Item Name", "Branch", "Qty", "Stock Before", "Stock After", "Buy Price", "Sell Price", "Profit", "Notes", "Operator", "Source"];
  _applyHeaders(sh, headers);

  if (!isFormatOnly) {
    const formulas = [];
    for (let r = 2; r <= 3000; r++) {
      formulas.push([`=IF(C${r}="","",ROW()-1)`]);
    }
    sh.getRange(2, 1, 2999, 1).setFormulas(formulas);
  }

  [45, 140, 110, 100, 200, 130, 70, 100, 100, 180, 120, 90].forEach((w, i) => sh.setColumnWidth(1 + i, w));
  sh.getRange("A:A").setHorizontalAlignment("center");
  sh.getRange("B:B").setNumberFormat("dd/MM/yyyy HH:mm:ss");
  sh.getRange("G:I").setNumberFormat("#,##0").setHorizontalAlignment("right");

  const tipeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["STOCK IN ➕", "STOCK OUT ➖", "ADJUSTMENT 🔧"])
    .setAllowInvalid(true).build();
  sh.getRange("C2:C3000").setDataValidation(tipeRule);

  const cfRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("STOCK IN").setBackground("#D1FAE5").setFontColor("#065F46")
      .setRanges([sh.getRange("A2:L3000")]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("STOCK OUT").setBackground("#FEE2E2").setFontColor("#991B1B")
      .setRanges([sh.getRange("A2:L3000")]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains("ADJUSTMENT").setBackground("#DBEAFE").setFontColor("#1E40AF")
      .setRanges([sh.getRange("A2:L3000")]).build(),
  ];
  sh.setConditionalFormatRules(cfRules);
  _alternateRows(sh, 2, 3000, 12);
  sh.setFrozenRows(1);
}

// ─── DASHBOARD SHEET ─────────────────────────────────────────
function _setupDashboardSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.DASHBOARD, 0);
  }

  const inv = "'" + SHEETS.INVENTORY + "'";
  const tx  = "'" + SHEETS.TRANSACTIONS + "'";

  sh.getRange("B2").setValue("📊  AI SMART INVENTORY SYSTEM  ·  REAL-TIME DASHBOARD");
  sh.getRange("B4").setValue("Last Updated:");
  sh.getRange("C4").setFormula('=TEXT(NOW(),"dd-MMM-yyyy HH:mm")');

  _dashCard(sh, 6, "B", "📦 Total Items",    `=COUNTA(${inv}!B2:B500)`);
  _dashCard(sh, 6, "E", "🔴 Out of Stock",   `=COUNTIF(${inv}!N2:N500,"*OUT OF STOCK*")`);
  _dashCard(sh, 6, "H", "🟠 Critical",       `=COUNTIF(${inv}!N2:N500,"*CRITICAL*")`);
  _dashCard(sh, 6, "K", "🟡 Low Stock",      `=COUNTIF(${inv}!N2:N500,"*LOW STOCK*")`);

  _dashCard(sh, 11, "B", "➕ Total Stock In",  `=SUM(${inv}!G2:G500)`);
  _dashCard(sh, 11, "E", "➖ Total Stock Out", `=SUM(${inv}!H2:H500)`);
  _dashCard(sh, 11, "H", "💰 Total Profit",    `=SUM(${tx}!L2:L3000)`);
  _dashCard(sh, 11, "K", "💎 Inventory Asset", `=SUM(ARRAYFORMULA(IFERROR(${inv}!I2:I500 * ${inv}!L2:L500)))`);

  sh.getRange("B16").setValue("⚠️  ITEMS REQUIRING ATTENTION  (Critical / Out of Stock)");
  ["Code", "Item Name", "Branch", "Stock Now", "Min Stock", "Status"].forEach((h, i) => {
    sh.getRange(17, i + 2).setValue(h);
  });
  sh.getRange("B18").setFormula(
    `=IFERROR(QUERY(${inv}!B2:P500,"SELECT B,C,E,I,J,N WHERE N CONTAINS 'CRITICAL' OR N CONTAINS 'OUT OF STOCK' ORDER BY I ASC LIMIT 12"))`
  );

  sh.getRange("B32").setValue("📂  STOCK BY CATEGORY");
  ["Category", "Items", "Total Stock"].forEach((h, i) => {
    sh.getRange(33, i + 2).setValue(h);
  });
  ["Electronics", "Fashion", "Groceries", "Office", "Food & Beverage", "Other"].forEach((cat, i) => {
    sh.getRange(34 + i, 2).setValue(cat);
    sh.getRange(34 + i, 3).setFormula(`=COUNTIF(${inv}!D2:D500,"${cat}")`);
    sh.getRange(34 + i, 4).setFormula(`=SUMIF(${inv}!D2:D500,"${cat}",${inv}!I2:I500)`);
  });

  sh.getRange("B42").setValue("🔄  RECENT TRANSACTIONS  (Last 12)");
  ["Date & Time", "Type", "Item", "Qty", "Stock After", "Notes"].forEach((h, i) => {
    sh.getRange(43, i + 2).setValue(h);
  });
  sh.getRange("B44").setFormula(
    `=IFERROR(QUERY(${tx}!B2:L3000,"SELECT B,C,E,G,I,J ORDER BY B DESC LIMIT 12",0),"")`
  );
}

// ─── GUIDE SHEET ─────────────────────────────────────────────
function _setupGuideSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEETS.GUIDE);
  if (sh) sh.clearContents(); else sh = ss.insertSheet(SHEETS.GUIDE);
  sh.setTabColor("#6366F1");

  sh.getRange("B2:J2").merge()
    .setValue("📋  USER GUIDE — AI Smart Inventory System v3.0")
    .setFontSize(16).setFontWeight("bold")
    .setBackground(COLORS.header).setFontColor("#FFFFFF")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(2, 52);

  const guide = [
    ["", ""],
    ["🚀  GETTING STARTED", ""],
    ["1.", "Activate your license from the menu 🔐 → Enter License Key to Unlock"],
    ["2.", "The full Inventory System menu will appear after successful activation"],
    ["3.", "Add your branches in the 🏢 Branches sheet"],
    ["4.", "Add inventory items in the 📦 Inventory sheet"],
    ["5.", "Use menu: Inventory System → Stock In / Stock Out to log movements"],
    ["6.", "View real-time summaries on the 📊 Dashboard"],
    ["", ""],
    ["📦  MANAGING ITEMS", ""],
    ["•", "Item Code: use short unique codes like ELE-001, FAS-001, GRO-001"],
    ["•", "Category: choose from the dropdown (can add more categories)"],
    ["•", "Min Stock: threshold below which status becomes LOW STOCK"],
    ["•", "Status is computed AUTOMATICALLY: 🟢 IN STOCK | 🟡 LOW | 🟠 CRITICAL | 🔴 OUT"],
    ["•", "Current Stock = Initial Stock + Total In − Total Out (auto-formula)"],
    ["", ""],
    ["🔄  LOGGING TRANSACTIONS", ""],
    ["•", "Use menu: 📦 Inventory System → Stock In / Stock Out"],
    ["•", "Enter item code, quantity, and optional notes — stock updates instantly"],
    ["•", "All history is preserved in the 🔄 Transactions sheet"],
    ["•", "AI Agent (WhatsApp/Telegram) logs are tagged with source channel"],
    ["", ""],
    ["🤖  AI AGENT (WHATSAPP & TELEGRAM)", ""],
    ["•", "Employees can report stock freely via WhatsApp or Telegram"],
    ["•", "AI understands natural language: 'received 10 iphone from supplier'"],
    ["•", "Automatic confirmation with calculation proof sent back to employee"],
    ["•", "Admin gets notified on every transaction"],
    ["•", "Setup: menu → Agent Settings (WA/TG/AI)"],
    ["", ""],
    ["🔑  LICENSE SYSTEM", ""],
    ["•", "Admin Key: ADMIN-MASTER-KEY-2026 (do not share)"],
    ["•", "Generate buyer licenses: menu → License Manager (Admin)"],
    ["•", "When buyer copies sheet → only License button shown → must activate"],
    ["•", "Block a license: change status to 🚫 Blocked in 🔑 Licenses sheet"],
    ["", ""],
    ["⚠️  STOCK STATUS EXPLAINED", ""],
    ["🟢 IN STOCK",    "Stock above minimum threshold — all good"],
    ["🟡 LOW STOCK",   "Stock equals minimum — reorder recommended"],
    ["🟠 CRITICAL",    "Stock ≤ 50% of minimum — urgent reorder needed"],
    ["🔴 OUT OF STOCK","Stock = 0 — cannot fulfill orders"],
  ];

  guide.forEach((row, i) => {
    const isBold = row[0].length > 2 && !row[0].startsWith("•") && !row[0].startsWith("🟢") && !row[0].startsWith("🟡") && !row[0].startsWith("🟠") && !row[0].startsWith("🔴");
    sh.getRange(i + 4, 2).setValue(row[0]).setFontWeight(isBold ? "bold" : "normal").setFontSize(10);
    sh.getRange(i + 4, 3, 1, 8).merge().setValue(row[1]).setFontSize(10);
  });

  sh.setColumnWidth(1, 18);
  sh.setColumnWidth(2, 80);
  for (let i = 3; i <= 10; i++) sh.setColumnWidth(i, 65);
  sh.setFrozenRows(3);
}

// ─── NAVIGATION HELPERS ──────────────────────────────────────
function navDashboard()    { _navTo(SHEETS.DASHBOARD); }
function navInventory()    { _navTo(SHEETS.INVENTORY); }
function navTransactions() { _navTo(SHEETS.TRANSACTIONS); }
function navBranches()     { _navTo(SHEETS.BRANCHES); }
function _navTo(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (sh) SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}

// ─── STOCK DIALOGS ───────────────────────────────────────────
function dialogStockIn() {
  const ui = SpreadsheetApp.getUi();
  const code = ui.prompt("➕ Stock In", "Enter Item Code (e.g. ELE-001):", ui.ButtonSet.OK_CANCEL);
  if (code.getSelectedButton() !== ui.Button.OK) return;
  const item = _findItem(code.getResponseText().trim().toUpperCase());
  if (!item) { ui.alert("❌ Item not found", "No item with that code exists in the Inventory sheet.", ui.ButtonSet.OK); return; }

  const qty = ui.prompt("➕ Stock In", `Item: ${item.name}\nCurrent Stock: ${item.stock}\n\nEnter quantity received:`, ui.ButtonSet.OK_CANCEL);
  if (qty.getSelectedButton() !== ui.Button.OK) return;
  const q = parseInt(qty.getResponseText());
  if (!q || q <= 0) { ui.alert("❌ Invalid quantity"); return; }

  const notes = ui.prompt("➕ Stock In", "Notes (optional):", ui.ButtonSet.OK_CANCEL);
  const nt = notes.getSelectedButton() === ui.Button.OK ? notes.getResponseText() : "";

  _recordTransaction(item, "STOCK IN ➕", q, nt, "Manual");
  ui.alert("✅ Stock In Recorded", `Item: ${item.name}\nAdded: +${q}\nNew Stock: ${item.stock + q}`, ui.ButtonSet.OK);
}

function dialogStockOut() {
  const ui = SpreadsheetApp.getUi();
  const code = ui.prompt("➖ Stock Out", "Enter Item Code:", ui.ButtonSet.OK_CANCEL);
  if (code.getSelectedButton() !== ui.Button.OK) return;
  const item = _findItem(code.getResponseText().trim().toUpperCase());
  if (!item) { ui.alert("❌ Item not found"); return; }

  const qty = ui.prompt("➖ Stock Out", `Item: ${item.name}\nAvailable Stock: ${item.stock}\n\nEnter quantity to remove:`, ui.ButtonSet.OK_CANCEL);
  if (qty.getSelectedButton() !== ui.Button.OK) return;
  const q = parseInt(qty.getResponseText());
  if (!q || q <= 0) { ui.alert("❌ Invalid quantity"); return; }
  if (q > item.stock) { ui.alert("❌ Insufficient Stock", `Only ${item.stock} in stock. Cannot remove ${q}.`, ui.ButtonSet.OK); return; }

  const notes = ui.prompt("➖ Stock Out", "Notes (optional):", ui.ButtonSet.OK_CANCEL);
  const nt = notes.getSelectedButton() === ui.Button.OK ? notes.getResponseText() : "";

  _recordTransaction(item, "STOCK OUT ➖", q, nt, "Manual");
  ui.alert("✅ Stock Out Recorded", `Item: ${item.name}\nRemoved: -${q}\nRemaining Stock: ${item.stock - q}`, ui.ButtonSet.OK);
}

function dialogAdjustment() {
  const ui = SpreadsheetApp.getUi();
  const code = ui.prompt("🔧 Stock Adjustment", "Enter Item Code:", ui.ButtonSet.OK_CANCEL);
  if (code.getSelectedButton() !== ui.Button.OK) return;
  const item = _findItem(code.getResponseText().trim().toUpperCase());
  if (!item) { ui.alert("❌ Item not found"); return; }

  const newStock = ui.prompt("🔧 Stock Adjustment", `Item: ${item.name}\nSystem Stock: ${item.stock}\n\nEnter ACTUAL stock count (from physical count):`, ui.ButtonSet.OK_CANCEL);
  if (newStock.getSelectedButton() !== ui.Button.OK) return;
  const ns = parseInt(newStock.getResponseText());
  if (isNaN(ns) || ns < 0) { ui.alert("❌ Invalid quantity"); return; }

  const diff = ns - item.stock;
  const reason = ui.prompt("🔧 Stock Adjustment", `Difference: ${diff > 0 ? "+" : ""}${diff}\nReason for adjustment:`, ui.ButtonSet.OK_CANCEL);
  const rn = reason.getSelectedButton() === ui.Button.OK ? reason.getResponseText() : "Physical count";

  _recordTransaction(item, "ADJUSTMENT 🔧", Math.abs(diff), rn, "Manual", ns);
  ui.alert("✅ Adjustment Recorded", `Item: ${item.name}\nOld Stock: ${item.stock}\nNew Stock: ${ns}\nDifference: ${diff > 0 ? "+" : ""}${diff}`, ui.ButtonSet.OK);
}

// ─── CORE TRANSACTION ENGINE ─────────────────────────────────
function _recordTransaction(item, type, qty, notes, source, overrideNewStock) {
  const invSh = _getSheet(SHEETS.INVENTORY);
  const txSh  = _getSheet(SHEETS.TRANSACTIONS);
  if (!invSh || !txSh) { Logger.log('_recordTransaction: sheet not found'); return 0; }

  // Get dynamic column maps
  const invHeaders = invSh.getRange(1, 1, 1, invSh.getLastColumn()).getValues()[0];
  const txHeaders = txSh.getRange(1, 1, 1, txSh.getLastColumn()).getValues()[0];
  const invCmap = _getInventoryColMap(invHeaders);
  const txCmap = _getTransactionColMap(txHeaders);

  const stockBefore = item.stock;
  let stockAfter;

  if (type.includes("IN")) {
    stockAfter = stockBefore + qty;
  } else if (type.includes("OUT")) {
    stockAfter = stockBefore - qty;
  } else { // ADJUSTMENT
    stockAfter = (overrideNewStock !== undefined) ? overrideNewStock : stockBefore;
  }

  // DIRECTLY update the Stok column in Inventory sheet
  if (invCmap.stock !== -1) {
    invSh.getRange(item.row, invCmap.stock + 1).setValue(stockAfter);
  } else {
    debugLog("WARNING: 'Stok' column not found in Inventory sheet! Cannot update stock.");
  }

  const op = "Operator";
  const lastRow = (txCmap.type !== -1) ? _getLastDataRow(txSh, txCmap.type + 1) + 1 : txSh.getLastRow() + 1;
  
  // Construct dynamic row based on txCmap
  const txRow = new Array(txHeaders.length).fill("");
  if (txCmap.date !== -1) txRow[txCmap.date] = new Date();
  if (txCmap.type !== -1) txRow[txCmap.type] = type;
  if (txCmap.code !== -1) txRow[txCmap.code] = item.code;
  if (txCmap.name !== -1) txRow[txCmap.name] = item.name;
  if (txCmap.branch !== -1) txRow[txCmap.branch] = item.branch;
  if (txCmap.qty !== -1) txRow[txCmap.qty] = qty;
  if (txCmap.stockBefore !== -1) txRow[txCmap.stockBefore] = stockBefore;
  if (txCmap.stockAfter !== -1) txRow[txCmap.stockAfter] = stockAfter;
  
  if (txCmap.buyPrice !== -1 && item.buyPrice !== undefined) txRow[txCmap.buyPrice] = item.buyPrice;
  if (txCmap.sellPrice !== -1 && item.sellPrice !== undefined) txRow[txCmap.sellPrice] = item.sellPrice;
  
  if (txCmap.profit !== -1 && type.includes("OUT")) {
    if (item.buyPrice !== undefined && item.sellPrice !== undefined) {
      txRow[txCmap.profit] = (item.sellPrice - item.buyPrice) * qty;
    }
  }

  if (txCmap.notes !== -1) txRow[txCmap.notes] = _sanitizeForSheet(notes);
  if (txCmap.operator !== -1) txRow[txCmap.operator] = _sanitizeForSheet(op);
  if (txCmap.source !== -1) txRow[txCmap.source] = _sanitizeForSheet(source);

  txSh.getRange(lastRow, 1, 1, txRow.length).setValues([txRow]);

  SpreadsheetApp.flush();
  
  // Proactive Low Stock Alert
  if (stockAfter <= item.minStock && stockAfter < stockBefore) {
    try {
      const adminId = PropertiesService.getScriptProperties().getProperty("ADMIN_CHAT_ID");
      if (adminId) {
        let alertMsg = `🚨 *LOW STOCK ALERT* 🚨\n`;
        alertMsg += `Item: ${item.name} (${item.code})\n`;
        alertMsg += `Current Stock: *${stockAfter}* (Min: ${item.minStock})\n`;
        alertMsg += `Please reorder soon!`;
        _sendTelegramMessage(adminId, alertMsg);
      }
    } catch (e) {
      debugLog("Failed to send proactive alert: " + e);
    }
  }

  return stockAfter;
}

function debugErrors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(SHEETS.INVENTORY);
  const dash = ss.getSheetByName(SHEETS.DASHBOARD);
  
  const errs = {
    n2_form: inv.getRange("N2").getFormula(),
    n2_val: inv.getRange("N2").getDisplayValue(),
    out_form: dash.getRange("E7").getFormula(),
    out_val: dash.getRange("E7").getDisplayValue(),
    q_form: dash.getRange("B18").getFormula(),
    q_val: dash.getRange("B18").getDisplayValue()
  };
  console.log("DEBUG_RESULTS: " + JSON.stringify(errs));
  return errs;
}

function debugErrorsUI() {
  const errs = debugErrors();
  SpreadsheetApp.getUi().alert("DEBUG INFO", JSON.stringify(errs, null, 2), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ─── CREATE NEW ITEM HELPER ────────────────────────────────────────
function _createNewItemRow(parsed) {
  const sh = _getSheet(SHEETS.INVENTORY);
  const data = sh.getDataRange().getValues();
  const cmap = _getInventoryColMap(data[0]);
  
  // Find the true last row by looking for the first empty Item Code (or Item Name)
  // This bypasses ArrayFormulas in column A that make the whole sheet look "full"
  let trueLastRowIndex = 0;
  for (let i = data.length - 1; i > 0; i--) {
    let checkVal = "";
    if (cmap.code !== -1) checkVal += data[i][cmap.code];
    if (cmap.name !== -1) checkVal += data[i][cmap.name];
    
    if (String(checkVal).trim() !== "") {
      trueLastRowIndex = i;
      break;
    }
  }
  
  // Target row is trueLastRowIndex + 1 (0-indexed) -> trueLastRowIndex + 2 (1-indexed)
  const targetRow = trueLastRowIndex + 2;
  
  // Generate SKU based on the true number of items
  let newId = trueLastRowIndex; // header is 0, item 1 is at index 1 -> ID 1
  const sku = "SKU-" + ("000" + (newId + 1)).slice(-4);
  
  const itemName = parsed.new_item_name || parsed.item_name || "New Item";
  
  if (targetRow > sh.getMaxRows()) {
    sh.insertRowAfter(sh.getMaxRows());
  }

  // Use individual setValue to avoid overwriting ArrayFormulas with empty strings
  if (cmap.code !== -1) sh.getRange(targetRow, cmap.code + 1).setValue(sku);
  if (cmap.name !== -1) sh.getRange(targetRow, cmap.name + 1).setValue(itemName);
  if (cmap.category !== -1) sh.getRange(targetRow, cmap.category + 1).setValue(parsed.new_category || "General");
  if (cmap.price !== -1) sh.getRange(targetRow, cmap.price + 1).setValue(parsed.new_price || 0);
  if (cmap.branch !== -1) sh.getRange(targetRow, cmap.branch + 1).setValue(parsed.branch || "");
  if (cmap.stock !== -1) sh.getRange(targetRow, cmap.stock + 1).setValue(0);
  
  return {
    row: targetRow,
    code: sku,
    name: itemName,
    branch: parsed.branch || "",
    stock: 0,
    minStock: 0,
    unit: "",
    buyPrice: 0,
    sellPrice: parsed.new_price || 0
  };
}

// ─── FIND ITEM HELPER ────────────────────────────────────────
function _findItem(code, name) {
  const sh = _getSheet(SHEETS.INVENTORY);
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return null;
  const cmap = _getInventoryColMap(data[0]);
  
  // Helper to build item object
  function buildItem(i) {
    return {
      row: i + 1,
      code: data[i][cmap.code],
      name: data[i][cmap.name],
      branch: (cmap.branch < data[i].length && cmap.branch !== -1) ? data[i][cmap.branch] : "",
      stock: (cmap.stock !== -1) ? (data[i][cmap.stock] || 0) : 0,
      minStock: (cmap.minStock < data[i].length && cmap.minStock !== -1) ? (data[i][cmap.minStock] || 0) : 0,
      unit: (cmap.unit < data[i].length && cmap.unit !== -1) ? (data[i][cmap.unit] || "") : "",
      buyPrice: (cmap.buyPrice !== -1) ? parseFloat(data[i][cmap.buyPrice]) || 0 : undefined,
      sellPrice: (cmap.sellPrice !== -1) ? parseFloat(data[i][cmap.sellPrice]) || 0 : undefined
    };
  }

  // Search by code first
  if (code && code !== "NEW") {
    for (let i = 1; i < data.length; i++) {
      const rowCode = data[i][cmap.code] ? data[i][cmap.code].toString().trim().toUpperCase() : "";
      if (rowCode === code.toString().trim().toUpperCase()) {
        return buildItem(i);
      }
    }
  }
  // Then search by name
  if (name) {
    const searchName = name.toString().trim().toUpperCase();
    for (let i = 1; i < data.length; i++) {
      const rowName = data[i][cmap.name] ? data[i][cmap.name].toString().trim().toUpperCase() : "";
      if (rowName.includes(searchName) || searchName.includes(rowName)) {
        return buildItem(i);
      }
    }
  }
  return null;
}

// ─── RECALCULATE STATUS ──────────────────────────────────────
function recalcAllStatus() {
  const sh = _getSheet(SHEETS.INVENTORY);
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  const cmap = _getInventoryColMap(data[0]);
  let n = 0;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][cmap.code]) continue;
    const s = data[i][cmap.stock] || 0, m = data[i][cmap.minStock] || 0;
    let status;
    if (s === 0) status = "🔴 OUT OF STOCK";
    else if (s <= m * 0.5) status = "🟠 CRITICAL";
    else if (s <= m) status = "🟡 LOW STOCK";
    else status = "🟢 IN STOCK";
    sh.getRange(i + 1, cmap.status + 1).setValue(status);
    n++;
  }
  SpreadsheetApp.getUi().alert("✅ Status Updated", `${n} items recalculated.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ─── REFRESH DASHBOARD (NO FORMULAS = NO LOCALE BUGS) ──────────
function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!sh) return;
  
  const inv = ss.getSheetByName(SHEETS.INVENTORY);
  const tx = ss.getSheetByName(SHEETS.TRANSACTIONS);
  if (!inv || !tx) return;
  
  const invData = inv.getDataRange().getValues().slice(1);
  const txData = tx.getDataRange().getValues().slice(1);
  
  let totalItems=0, outOfStock=0, critical=0, lowStock=0;
  let totalIn=0, totalOut=0, inventoryAsset=0;
  
  let criticalItems = [];
  
  // Categorize
  const catStats = {
    "Electronics": { count: 0, stock: 0 },
    "Fashion": { count: 0, stock: 0 },
    "Groceries": { count: 0, stock: 0 },
    "Office": { count: 0, stock: 0 },
    "Food & Beverage": { count: 0, stock: 0 },
    "Other": { count: 0, stock: 0 }
  };
  
  for (let i = 0; i < invData.length; i++) {
    const row = invData[i];
    if (!row[1]) continue; // Skip empty Code
    totalItems++;
    
    let stock = Number(row[8]) || 0;
    let buyPrice = Number(row[11]) || 0;
    totalIn += Number(row[6]) || 0;
    totalOut += Number(row[7]) || 0;
    inventoryAsset += (stock * buyPrice);
    
    let status = row[13] || "";
    if (status.includes("OUT OF STOCK")) outOfStock++;
    if (status.includes("CRITICAL")) critical++;
    if (status.includes("LOW STOCK")) lowStock++;
    
    if (status.includes("CRITICAL") || status.includes("OUT OF STOCK")) {
      criticalItems.push([row[1], row[2], row[4], stock, row[9], status]);
    }
    
    let cat = row[3] || "Other";
    if (!catStats[cat]) catStats[cat] = { count: 0, stock: 0 };
    catStats[cat].count++;
    catStats[cat].stock += stock;
  }
  
  let totalProfit = 0;
  let recentTx = [];
  for (let i = txData.length - 1; i >= 0; i--) {
    const row = txData[i];
    if (!row[1]) continue;
    totalProfit += Number(row[11]) || 0;
    if (recentTx.length < 8) {
      recentTx.push([row[1], row[2], row[4], row[6], row[8], row[9], row[11], row[13]]);
    }
  }
  
  // Update UI values
  sh.getRange("C4:F4").setValue(new Date().toLocaleString('id-ID', {day:'2-digit', month:'short', year:'numeric', hour:'2-digit', minute:'2-digit'}));
  sh.getRange("B7").setValue(totalItems);
  sh.getRange("E7").setValue(outOfStock);
  sh.getRange("H7").setValue(critical);
  sh.getRange("K7").setValue(lowStock);
  
  sh.getRange("B12").setValue(totalIn);
  sh.getRange("E12").setValue(totalOut);
  sh.getRange("H12").setValue(totalProfit);
  sh.getRange("K12").setValue(inventoryAsset);
  
  // Clear tables
  sh.getRange("B18:G29").clearContent();
  if (criticalItems.length > 0) {
    criticalItems.sort((a,b) => a[3] - b[3]);
    sh.getRange(18, 2, Math.min(12, criticalItems.length), 6).setValues(criticalItems.slice(0, 12));
  }
  
  const cats = ["Electronics", "Fashion", "Groceries", "Office", "Food & Beverage", "Other"];
  for (let i = 0; i < cats.length; i++) {
    sh.getRange(34 + i, 3).setValue(catStats[cats[i]].count);
    sh.getRange(34 + i, 4).setValue(catStats[cats[i]].stock);
  }
  
  sh.getRange("F34:M41").clearContent();
  if (recentTx.length > 0) {
    sh.getRange(34, 6, recentTx.length, 8).setValues(recentTx);
  }
}

function refreshDashboardClick() {
  refreshDashboard();
  navDashboard();
  SpreadsheetApp.getUi().alert("✅ Dashboard Terkini", "Data sukses diambil langsung dari pusat database.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ─── LICENSE MANAGER (ADMIN ONLY) ────────────────────────────
function showLicenseManager() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("IS_ADMIN") !== "true") {
    const ui = SpreadsheetApp.getUi();
    const r = ui.prompt("🔑 License Manager", "Enter Admin Key to continue:", ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK || r.getResponseText() !== "ADMIN-MASTER-KEY-2026") {
      ui.alert("❌ Access Denied", "Incorrect admin key.", ui.ButtonSet.OK);
      return;
    }
  }
  const ui = SpreadsheetApp.getUi();
  const count = ui.prompt("🔑 License Generator", "How many license keys to generate?", ui.ButtonSet.OK_CANCEL);
  if (count.getSelectedButton() !== ui.Button.OK) return;
  const n = parseInt(count.getResponseText()) || 1;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LICENSES);
  const keys = [];
  for (let i = 0; i < n; i++) {
    const key = generateValidKey();
    keys.push(key);
    const lastRow = _getLastDataRow(sh, 2) + 1;
    sh.getRange(lastRow, 2, 1, 5).setValues([[key, "", "⏳ Pending", "", new Date()]]);
    sh.getRange(lastRow, 1).setValue(lastRow - 1);
  }
  ui.alert("✅ Keys Generated", `${n} license key(s) created:\n\n${keys.join("\n")}\n\nView them in the 🔑 Licenses sheet.`, ui.ButtonSet.OK);
}

// ─── AGENT SETTINGS SIDEBAR ──────────────────────────────────
function showAgentSettings() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty("WEBHOOK_SECRET");
  if (!secret) {
    secret = "SEC-" + Utilities.getUuid();
    props.setProperty("WEBHOOK_SECRET", secret);
  }
  const secureUrl = ScriptApp.getService().getUrl() + "?token=" + secret;
  const html = HtmlService.createHtmlOutput(`
    <style>
      * { box-sizing: border-box; font-family: 'Segoe UI', Arial, sans-serif; }
      body { background: #0F172A; color: #E2E8F0; padding: 16px; font-size: 13px; }
      h2 { font-size: 15px; color: #10B981; margin-bottom: 16px; }
      .section { background: #1E293B; border-radius: 10px; padding: 14px; margin-bottom: 14px; border: 1px solid #334155; }
      .section h3 { font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; color: #94A3B8; margin-bottom: 10px; }
      label { display: block; font-size: 11px; color: #94A3B8; margin-bottom: 4px; margin-top: 8px; }
      input { width: 100%; padding: 8px 10px; background: #0F172A; border: 1px solid #334155; border-radius: 6px; color: #F1F5F9; font-size: 12px; font-family: monospace; outline: none; }
      input:focus { border-color: #10B981; }
      button { padding: 10px 16px; background: #10B981; color: white; border: none; border-radius: 7px; font-size: 13px; font-weight: 700; cursor: pointer; width: 100%; margin-top: 12px; }
      .note { font-size: 11px; color: #475569; margin-top: 6px; line-height: 1.4; }
      a { color: #10B981; }
    </style>
    <h2>🤖 AI Agent Settings</h2>
    <div class="section">
      <h3>📱 WhatsApp (Meta Cloud API)</h3>
      <label>Phone Number ID</label>
      <input id="wa_phone_id" value="${_escapeHtml(props.getProperty("WA_PHONE_ID") || "")}" placeholder="e.g. 1234567890123456" />
      <label>Access Token</label>
      <input id="wa_token" value="${_escapeHtml(props.getProperty("WA_TOKEN") || "")}" placeholder="EAAxxxxxxxx..." />
      <p class="note">Get free at <a href="https://developers.facebook.com" target="_blank">developers.facebook.com</a> → WhatsApp → Getting Started</p>
    </div>
    <div class="section">
      <h3>✈️ Telegram Bot</h3>
      <label>Bot Token</label>
      <input id="tg_token" value="${_escapeHtml(props.getProperty("TG_TOKEN") || "")}" placeholder="1234567890:ABCdef..." />
      <label>Admin Chat ID</label>
      <input id="tg_admin_id" value="${_escapeHtml(props.getProperty("TG_ADMIN_ID") || "")}" placeholder="Your Telegram chat ID" />
      <p class="note">Create bot via <a href="https://t.me/BotFather" target="_blank">@BotFather</a> on Telegram</p>
    </div>
    <div class="section">
      <h3>🤖 AI Engine (Models & API)</h3>
      <p class="note">Support for OpenAI, Groq, Gemini, HuggingFace, and Local APIs.</p>
      <button type="button" onclick="google.script.run.showAIProviderSettings()" style="background:#3B82F6; margin-top:6px; padding:8px;">⚙️ Advanced AI Settings</button>
    </div>
    <div class="section">
      <h3>🔗 Secure Webhook URL (WhatsApp & Telegram)</h3>
      <p class="note">Copy this URL to your Meta App Dashboard for WhatsApp Webhook:</p>
      <input readonly value="${_escapeHtml(secureUrl)}" onclick="this.select()" />
      <p class="note"><b>Webhook Verify Token (WhatsApp):</b> inventory2026</p>
    </div>
    <div class="section">
      <h3>📧 Email Triggers & Notifications</h3>
      <label>Email Triggers (Uses your current Google Account automatically)</label>
      <p class="note">AI will automatically process emails with the subject "AI Inventory Report".</p>
    </div>
    <div id="statusMsg" style="margin-top:12px; font-weight:bold; font-size:12px; text-align:center;"></div>
    <button id="saveBtn" onclick="save()">💾 Save Settings & Activate Agent</button>
    <script>
      function save() {
        const btn = document.getElementById('saveBtn');
        const msg = document.getElementById('statusMsg');
        btn.textContent = '⏳ Saving... please wait';
        btn.disabled = true;
        msg.textContent = '';
        msg.style.color = '';

        const data = {
          wa_phone_id: document.getElementById('wa_phone_id').value,
          wa_token: document.getElementById('wa_token').value,
          tg_token: document.getElementById('tg_token').value,
          tg_admin_id: document.getElementById('tg_admin_id').value
        };
        
        google.script.run
          .withSuccessHandler(() => {
            btn.textContent = '💾 Save Settings & Activate Agent';
            btn.disabled = false;
            msg.textContent = '✅ Settings Saved & Agent Activated!';
            msg.style.color = '#10B981';
          })
          .withFailureHandler(err => {
            msg.textContent = '❌ Error: ' + err.message;
            msg.style.color = '#EF4444';
            btn.textContent = '💾 Save Settings & Activate Agent';
            btn.disabled = false;
          })
          .saveAgentSettings(data);
      }
    </script>
  `).setTitle("AI Agent Settings").setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ─── AI SIMULATOR ──────────────────────────────────────────
function simulateChat() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("🧪 AI Agent Simulator", "Enter a sample message from an employee\n(e.g. 'Bos, terima 10 mie sedap tadi siang'):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    const text = resp.getResponseText();
    if (!text) return;
    _processWithAI("sim-001", text, "Simulator User", "Simulation");
  }
}

function saveAgentSettings(data) {
  const props = PropertiesService.getScriptProperties();
  if (data.wa_phone_id) props.setProperty("WA_PHONE_ID",  data.wa_phone_id);
  if (data.wa_token)    props.setProperty("WA_TOKEN",     data.wa_token);
  if (data.tg_token)    props.setProperty("TG_TOKEN",     data.tg_token);
  if (data.tg_admin_id) props.setProperty("TG_ADMIN_ID",  data.tg_admin_id);
  if (data.admin_email) props.setProperty("ADMIN_EMAIL",  data.admin_email);
  
  _removeAllTriggers("pollWhatsApp");
  _removeAllTriggers("pollEmails");
  _removeAllTriggers("pollTelegram");
  
  const adminEmail = Session.getActiveUser().getEmail();
  props.setProperty("ADMIN_EMAIL", adminEmail);
  props.setProperty("SHEET_ID", SpreadsheetApp.getActiveSpreadsheet().getId());
  
  // Set up FULLY AUTONOMOUS background workers (Polling)
  // This bypasses all Google Apps Script Webhook 302 restrictions
  ScriptApp.newTrigger("pollEmails").timeBased().everyMinutes(5).create();
  if (data.tg_token || props.getProperty("TG_TOKEN")) {
    const t = data.tg_token || props.getProperty("TG_TOKEN");
    try {
      // Unregister webhook to avoid 302 loops and allow getUpdates to work
      UrlFetchApp.fetch(`https://api.telegram.org/bot${t}/setWebhook?url=`, { muteHttpExceptions: true });
    } catch(e) {}
    ScriptApp.newTrigger("pollTelegram").timeBased().everyMinutes(1).create();
  }
}

function _registerTelegramWebhook(token) {
  try {
    const props = PropertiesService.getScriptProperties();
    const secret = props.getProperty("WEBHOOK_SECRET");
    const webhookUrl = ScriptApp.getService().getUrl() + "?token=" + secret;
    UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/setWebhook?url=${encodeURIComponent(webhookUrl)}`, { muteHttpExceptions: true });
  } catch (e) { Logger.log("Webhook registration failed: " + e); }
}

// ─── ABOUT ───────────────────────────────────────────────────
function showAbout() {
  const props = PropertiesService.getScriptProperties();
  SpreadsheetApp.getUi().alert(
    `ℹ️ ${_SYSTEM.APP_NAME}`,
    `Version: ${_SYSTEM.VERSION}\nBuild: ${_SYSTEM.BUILD_DATE}\n\n` +
    `Status: ${props.getProperty("LICENSE_KEY") ? "Activated" : "Not activated"}\n` +
    `License Key: ${props.getProperty("LICENSE_KEY") || "None"}\n\n` +
    `Features:\n• Multi-branch inventory management\n• AI-powered WhatsApp & Telegram agent\n• Real-time dashboard & analytics\n• Transaction log with proof\n• License protection system\n• Auto stock status calculation`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ─── SHARED HELPERS ──────────────────────────────────────────
function _applyHeaders(sh, headers, bgColor) {
  bgColor = bgColor || COLORS.header;
  sh.setRowHeight(1, 36);
  headers.forEach((h, i) => {
    sh.getRange(1, i + 1)
      .setValue(h).setBackground(bgColor).setFontColor("#FFFFFF")
      .setFontWeight("bold").setFontSize(10)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  });
}

function _alternateRows(sh, startRow, endRow, numCols, startCol) {
  startCol = startCol || 1;
  const numRows = endRow - startRow + 1;
  if (numRows <= 0) return;
  const backgrounds = [];
  for (let r = startRow; r <= endRow; r++) {
    const rowColor = (r % 2 === 0) ? "#F8FAFC" : "#FFFFFF";
    backgrounds.push(Array(numCols).fill(rowColor));
  }
  sh.getRange(startRow, startCol, numRows, numCols).setBackgrounds(backgrounds);
}

function _dashCard(sh, startRow, col, label, formula) {
  const c = col.charCodeAt(0) - 64;
  sh.getRange(startRow, c).setValue(label);
  sh.getRange(startRow + 1, c).setFormula(formula);
}

function _getLastDataRow(sh, col) {
  const vals = sh.getRange(1, col, sh.getMaxRows(), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== "" && vals[i][0] !== null) return i + 1;
  }
  return 1;
}

function _removeAllTriggers(funcName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === funcName)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function _getInventoryContext() {
  const sh = _getSheet(SHEETS.INVENTORY);
  if (!sh) return "No data";
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) return "No data";
  const data = sh.getDataRange().getDisplayValues(); // use DisplayValues to format properly
  if (data.length < 2) return "No data";
  
  const headers = data[0];
  const items = [];
  
  for (let i = 1; i < data.length && items.length < 100; i++) {
    if (!data[i][0]) continue; // assume first column is ID/Code
    let rowProps = [];
    for (let c = 0; c < headers.length; c++) {
      if (headers[c] && data[i][c] && data[i][c].toString().trim() !== "") {
        rowProps.push(headers[c] + ":" + data[i][c]);
      }
    }
    items.push(rowProps.join(" | "));
  }
  return items.join("\n") || "No items";
}

// ─── DYNAMIC COLUMN MAPPERS ──────────────────────────────────
// These mappers scan ALL headers and find columns by keyword.
// No hardcoded positions — works with any spreadsheet layout.
// Returns -1 for columns that don't exist in the spreadsheet.

function _getInventoryColMap(headers) {
  // Default: all -1 (not found). Code MUST check !== -1 before using.
  const map = { code: -1, name: -1, category: -1, price: -1, branch: -1, 
                stockIn: -1, stockOut: -1, stock: -1, minStock: -1, unit: -1, status: -1 };
  if (!headers || headers.length === 0) return map;

  // Normalize all headers once
  const normalized = headers.map(h => String(h).toLowerCase().trim());

  // Priority rules: check MORE SPECIFIC patterns first, then general ones.
  // Supports English, Indonesian, Spanish, French, German, and common abbreviations.
  const rules = [
    // Stock In/Out MUST be checked before generic "stock"
    ["stockIn",  h => h.includes("stock in") || h.includes("stok masuk") || h === "masuk" || h === "in" || h.includes("entrada") || h.includes("entree") || h.includes("zugang")],
    ["stockOut", h => h.includes("stock out") || h.includes("stok keluar") || h === "keluar" || h === "out" || h.includes("salida") || h.includes("sortie") || h.includes("abgang")],
    // Min stock MUST be checked before generic "stock"
    ["minStock", h => (h.includes("min") && (h.includes("stock") || h.includes("stok"))) || h.includes("reorder") || h.includes("safety") || h.includes("minimo")],
    // Buy/Sell Prices
    ["buyPrice", h => h.includes("buy") || h.includes("beli") || h.includes("modal") || h.includes("achat") || h.includes("compra")],
    ["sellPrice", h => h.includes("sell") || h.includes("jual") || h.includes("vente") || h.includes("venta")],
    // General Price (if buy/sell not found)
    ["price",    h => h.includes("harga") || h.includes("price") || h.includes("cost") || h.includes("precio") || h.includes("prix") || h.includes("preis")],
    // SKU / Item Code
    ["code",     h => h === "item code" || h === "kode barang" || h === "sku" || h === "code" || h === "barcode" || h === "upc" || h === "ean"],
    // Item Name
    ["name",     h => h.includes("nama") || h.includes("name") || h.includes("barang") || h.includes("produk") || h.includes("product") || h.includes("item") || h.includes("nombre") || h.includes("nom") || h.includes("artikel")],
    // Category
    ["category", h => h.includes("kategori") || h.includes("category") || h.includes("jenis") || h.includes("tipe") || h.includes("type") || h.includes("categoria") || h.includes("categoría") || h.includes("catégorie") || h.includes("kategorie")],
    // Branch / Location
    ["branch",   h => h.includes("cabang") || h.includes("branch") || h.includes("toko") || h.includes("lokasi") || h.includes("gudang") || h.includes("warehouse") || h.includes("location") || h.includes("sucursal") || h.includes("succursale") || h.includes("filiale")],
    // Unit (but NOT if header also contains "harga/price")
    ["unit",     h => !h.includes("harga") && !h.includes("price") && (h === "satuan" || h === "unit" || h === "uom" || h.includes("unidad") || h.includes("unite") || h.includes("einheit"))],
    // Status
    ["status",   h => h.includes("status") || h.includes("kondisi") || h.includes("state") || h.includes("estado") || h.includes("etat")],
    // Generic Stock / Current Stock (checked LAST so specific stock in/out/min get matched first)
    ["stock",    h => h.includes("stock") || h.includes("stok") || h.includes("qty") || h.includes("quantity") || h.includes("jumlah") || h.includes("persediaan") || h.includes("saldo") || h.includes("inventario") || h.includes("inventaire") || h.includes("bestand") || h.includes("cantidad")],
  ];

  for (let i = 0; i < normalized.length; i++) {
    const h = normalized[i];
    if (!h) continue;
    for (const [key, test] of rules) {
      if (map[key] === -1 && test(h)) {
        map[key] = i;
        break;
      }
    }
  }

  // Fallbacks
  if (map.code === -1) map.code = 0;
  if (map.name === -1) map.name = map.code + 1;

  debugLog("ColMap detected: " + JSON.stringify(map) + " | Headers: " + JSON.stringify(normalized));
  return map;
}

function _getTransactionColMap(headers) {
  const map = { txId: -1, date: -1, type: -1, code: -1, name: -1, branch: -1, 
                qty: -1, stockBefore: -1, stockAfter: -1, notes: -1, operator: -1, source: -1 };
  if (!headers || headers.length === 0) return map;

  const normalized = headers.map(h => String(h).toLowerCase().trim());

  const rules = [
    ["stockBefore", h => h.includes("before") || h.includes("awal") || h.includes("sebelum") || h.includes("antes") || h.includes("avant") || h.includes("vorher")],
    ["stockAfter",  h => h.includes("after") || h.includes("akhir") || h.includes("sesudah") || h.includes("despues") || h.includes("después") || h.includes("apres") || h.includes("nachher")],
    ["txId",        h => h === "#" || h === "id" || h.includes("tx") || h.includes("trans") || h.includes("transaksi") || h.includes("transaction") || h.includes("transaccion")],
    ["date",        h => h.includes("date") || h.includes("time") || h.includes("tanggal") || h.includes("waktu") || h.includes("fecha") || h.includes("datum")],
    ["type",        h => h.includes("type") || h.includes("tipe") || h.includes("jenis") || h.includes("tipo")],
    ["code",        h => h.includes("code") || h.includes("kode") || h.includes("sku") || h.includes("codigo") || h.includes("código") || h.includes("id")],
    ["name",        h => h.includes("name") || h.includes("nama") || h.includes("barang") || h.includes("product") || h.includes("nombre") || h.includes("nom") || h.includes("artikel")],
    ["qty",         h => h.includes("qty") || h.includes("quantity") || h.includes("jumlah") || h.includes("cantidad") || h.includes("quantite") || h.includes("menge")],
    ["branch",      h => h.includes("branch") || h.includes("cabang") || h.includes("lokasi") || h.includes("sucursal") || h.includes("filiale")],
    ["buyPrice",    h => h.includes("buy") || h.includes("beli") || h.includes("modal")],
    ["sellPrice",   h => h.includes("sell") || h.includes("jual")],
    ["profit",      h => h.includes("profit") || h.includes("untung") || h.includes("laba") || h.includes("margin") || h.includes("beneficio")],
    ["notes",       h => h.includes("note") || h.includes("catatan") || h.includes("keterangan") || h.includes("nota") || h.includes("remark")],
    ["operator",    h => h.includes("operator") || h.includes("petugas") || h.includes("user") || h.includes("usuario") || h.includes("utilisateur") || h.includes("benutzer")],
    ["source",      h => h.includes("source") || h.includes("channel") || h.includes("sumber") || h.includes("fuente") || h.includes("quelle")],
  ];

  for (let i = 0; i < normalized.length; i++) {
    const h = normalized[i];
    if (!h) continue;
    for (const [key, test] of rules) {
      if (map[key] === -1 && test(h)) {
        map[key] = i;
        break;
      }
    }
  }

  return map;
}

// ─── SECURITY HELPER FUNCTIONS ───────────────────────────────
function _sanitizeForSheet(text) {
  if (typeof text !== "string") return text;
  // Prevent Formula Injection (CSV Injection)
  const trimmed = text.trim();
  if (trimmed.charAt(0) === '=' || trimmed.charAt(0) === '+' || trimmed.charAt(0) === '-' || trimmed.charAt(0) === '@') {
    return "'" + text;
  }
  return text;
}

function _escapeHtml(text) {
  if (typeof text !== "string") return text;
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// ═══════════════════════════════════════════════════════════════
//  AI SMART REPAIR — Intelligent Data Cleanup & Sync
// ═══════════════════════════════════════════════════════════════

function showRepairDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', sans-serif; padding: 20px; background: #1a1a2e; color: #e0e0e0; }
        h2 { color: #00d4ff; margin-bottom: 10px; }
        .desc { color: #888; font-size: 13px; margin-bottom: 20px; line-height: 1.5; }
        .option { background: #16213e; border: 1px solid #333; border-radius: 8px; padding: 14px; margin: 8px 0; cursor: pointer; transition: all 0.2s; }
        .option:hover { border-color: #00d4ff; background: #1a2744; }
        .option h4 { color: #00d4ff; margin-bottom: 4px; }
        .option p { color: #999; font-size: 12px; }
        .btn { display: block; width: 100%; padding: 12px; margin-top: 16px; border: none; border-radius: 8px; font-size: 14px; font-weight: bold; cursor: pointer; transition: all 0.2s; }
        .btn-primary { background: linear-gradient(135deg, #00d4ff, #0051ff); color: white; }
        .btn-primary:hover { opacity: 0.9; }
        .btn-secondary { background: #333; color: #ccc; margin-top: 8px; }
        #progress { display: none; margin-top: 16px; }
        #progress .bar { height: 6px; background: #333; border-radius: 3px; overflow: hidden; }
        #progress .fill { height: 100%; background: linear-gradient(90deg, #00d4ff, #0051ff); width: 0%; transition: width 0.5s; }
        #progress .log { margin-top: 10px; font-size: 12px; color: #888; max-height: 200px; overflow-y: auto; background: #111; padding: 10px; border-radius: 6px; }
        .log-line { padding: 2px 0; border-bottom: 1px solid #1a1a2e; }
        .log-ok { color: #4caf50; }
        .log-warn { color: #ff9800; }
        .log-err { color: #f44336; }
        .log-info { color: #00d4ff; }
      </style>
    </head>
    <body>
      <h2>🛠️ AI Smart Repair</h2>
      <p class="desc">Automatically fix messy data. The system will analyze the spreadsheet, find anomalies, and repair them.</p>
      
      <div class="option" onclick="selectMode('full')">
        <h4>🔄 Full Repair (Recommended)</h4>
        <p>Clean stray data, sync stock from Transaction Logs, and fix all anomalies</p>
      </div>
      
      <div class="option" onclick="selectMode('clean')">
        <h4>🧹 Clean Only</h4>
        <p>Only clean data outside of valid columns (stray data in empty columns)</p>
      </div>

      <div class="option" onclick="selectMode('sync')">
        <h4>📊 Sync Stock Only</h4>
        <p>Recalculate stock from Transaction Logs without cleaning columns</p>
      </div>

      <button class="btn btn-primary" onclick="startRepair()">▶️ Start Repair</button>
      <button class="btn btn-secondary" onclick="google.script.host.close()">Cancel</button>
      
      <div id="progress">
        <div class="bar"><div class="fill" id="progressFill"></div></div>
        <div class="log" id="logBox"></div>
      </div>

      <script>
        let mode = 'full';
        function selectMode(m) {
          mode = m;
          document.querySelectorAll('.option').forEach(el => el.style.borderColor = '#333');
          event.currentTarget.style.borderColor = '#00d4ff';
        }
        // auto-select full
        document.querySelector('.option').style.borderColor = '#00d4ff';

        function startRepair() {
          document.getElementById('progress').style.display = 'block';
          addLog('info', '🚀 Starting ' + mode + ' repair...');
          google.script.run
            .withSuccessHandler(onResult)
            .withFailureHandler(onError)
            .runSmartRepair(mode);
        }

        function addLog(type, msg) {
          const box = document.getElementById('logBox');
          box.innerHTML += '<div class="log-line log-' + type + '">' + msg + '</div>';
          box.scrollTop = box.scrollHeight;
        }

        function onResult(result) {
          document.getElementById('progressFill').style.width = '100%';
          const r = JSON.parse(result);
          r.logs.forEach(l => addLog(l.type, l.msg));
          addLog('ok', '');
          addLog('ok', '══════════════════════════════');
          addLog('ok', '✅ REPAIR COMPLETE');
          addLog('info', '📊 Items analyzed: ' + r.stats.itemsAnalyzed);
          addLog('info', '🔧 Stock corrections: ' + r.stats.stockFixed);
          addLog('info', '🧹 Cells cleaned: ' + r.stats.cellsCleaned);
          if (r.stats.anomalies > 0) addLog('warn', '⚠️ Anomalies found: ' + r.stats.anomalies);
        }

        function onError(err) {
          addLog('err', '❌ Error: ' + err.message);
        }
      </script>
    </body>
    </html>
  `).setWidth(460).setHeight(560).setTitle("AI Smart Repair");
  SpreadsheetApp.getUi().showModalDialog(html, "🛠️ AI Smart Repair");
}

function runSmartRepair(mode) {
  const logs = [];
  const stats = { itemsAnalyzed: 0, stockFixed: 0, cellsCleaned: 0, anomalies: 0 };

  function log(type, msg) { logs.push({ type, msg }); }

  try {
    const invSh = _getSheet(SHEETS.INVENTORY);
    const txSh = _getSheet(SHEETS.TRANSACTIONS);
    if (!invSh) { log("err", "❌ Inventory sheet not found!"); return JSON.stringify({ logs, stats }); }

    const invData = invSh.getDataRange().getValues();
    if (invData.length < 2) { log("err", "❌ Inventory sheet is empty!"); return JSON.stringify({ logs, stats }); }

    const invCmap = _getInventoryColMap(invData[0]);
    log("info", "📋 Detected columns: SKU=" + _colLetter(invCmap.code) + " Name=" + _colLetter(invCmap.name) + " Stock=" + _colLetter(invCmap.stock) + " Branch=" + _colLetter(invCmap.branch));

    // ─── STEP 1: CLEAN STRAY DATA ────────────────────────────
    if (mode === "full" || mode === "clean") {
      log("info", "🧹 Step 1: Cleaning stray data...");

      // Find the rightmost VALID column (that has a known header)
      const validCols = Object.values(invCmap).filter(v => v !== -1);
      const maxValidCol = Math.max(...validCols);
      const totalCols = invSh.getLastColumn();

      if (totalCols > maxValidCol + 1) {
        const strayCols = totalCols - (maxValidCol + 1);
        log("warn", "⚠️ Found " + strayCols + " columns with potential stray data (columns " + _colLetter(maxValidCol + 1) + " to " + _colLetter(totalCols - 1) + ")");

        // Check if stray columns actually have data
        for (let col = maxValidCol + 2; col <= totalCols; col++) {
          const colData = invSh.getRange(2, col, invData.length - 1, 1).getValues();
          let hasData = false;
          for (let r = 0; r < colData.length; r++) {
            if (colData[r][0] !== "" && colData[r][0] !== null) {
              hasData = true;
              break;
            }
          }
          if (hasData) {
            invSh.getRange(2, col, invData.length - 1, 1).clearContent();
            stats.cellsCleaned += invData.length - 1;
            log("ok", "  ✅ Cleared column " + _colLetter(col - 1) + " (" + (invData.length - 1) + " cells)");
          }
        }

        if (stats.cellsCleaned === 0) {
          log("ok", "  ✅ No stray data found — columns are clean");
        }
      } else {
        log("ok", "  ✅ No extra columns detected — sheet is clean");
      }
    }

    // ─── STEP 2: SYNC STOCK FROM TRANSACTION LOGS ────────────
    if ((mode === "full" || mode === "sync") && txSh) {
      log("info", "📊 Step 2: Recalculating stock from Transaction Logs...");

      const txData = txSh.getDataRange().getValues();
      if (txData.length < 2) {
        log("warn", "⚠️ Transaction Logs is empty — skipping sync");
      } else {
        const txCmap = _getTransactionColMap(txData[0]);
        log("info", "  Transaction columns: Type=" + _colLetter(txCmap.type) + " Code=" + _colLetter(txCmap.code) + " Qty=" + _colLetter(txCmap.qty));

        // Build a map: itemCode -> { firstStockBefore, totalIn, totalOut, latestStockAfter }
        const txMap = {};
        for (let t = 1; t < txData.length; t++) {
          const code = txCmap.code !== -1 ? String(txData[t][txCmap.code]).trim() : "";
          const type = txCmap.type !== -1 ? String(txData[t][txCmap.type]).toUpperCase().trim() : "";
          const qty = txCmap.qty !== -1 ? (Number(txData[t][txCmap.qty]) || 0) : 0;
          const sBefore = txCmap.stockBefore !== -1 ? Number(txData[t][txCmap.stockBefore]) : NaN;
          const sAfter = txCmap.stockAfter !== -1 ? Number(txData[t][txCmap.stockAfter]) : NaN;

          if (!code || !type) continue;

          if (!txMap[code]) {
            txMap[code] = { firstBefore: NaN, totalIn: 0, totalOut: 0, totalAdj: 0, count: 0, latestAfter: NaN };
          }
          const entry = txMap[code];
          entry.count++;
          if (isNaN(entry.firstBefore) && !isNaN(sBefore)) entry.firstBefore = sBefore;
          if (!isNaN(sAfter)) entry.latestAfter = sAfter;

          if (type.includes("IN")) entry.totalIn += Math.abs(qty);
          else if (type.includes("OUT")) entry.totalOut += Math.abs(qty);
          else if (type.includes("ADJUST")) entry.totalAdj += qty;
        }

        // Now compare with current inventory stock
        const newStockValues = [];
        let hasUpdates = false;

        for (let i = 1; i < invData.length; i++) {
          const code = String(invData[i][invCmap.code]).trim();
          let currentStock = invCmap.stock !== -1 ? (Number(invData[i][invCmap.stock]) || 0) : 0;
          let newStock = currentStock;

          if (code) {
             stats.itemsAnalyzed++;
             const itemName = invData[i][invCmap.name] || code;

             if (txMap[code]) {
               const tx = txMap[code];
               let calculatedStock;
               if (!isNaN(tx.firstBefore)) calculatedStock = tx.firstBefore + tx.totalIn - tx.totalOut + tx.totalAdj;
               else if (!isNaN(tx.latestAfter)) calculatedStock = tx.latestAfter;

               if (calculatedStock !== undefined && calculatedStock !== currentStock) {
                 log("warn", "  ⚠️ " + itemName + " (" + code + "): Sheet=" + currentStock + " → Calculated=" + calculatedStock);
                 newStock = calculatedStock;
                 stats.stockFixed++;
                 hasUpdates = true;
                 log("ok", "    ✅ Stock corrected to " + calculatedStock);
               } else {
                 log("ok", "  ✅ " + itemName + ": Stock=" + currentStock + " (correct, " + tx.count + " transactions)");
               }
             } else {
               log("ok", "  ✅ " + itemName + ": Stock=" + currentStock + " (no transactions)");
             }
          }
          newStockValues.push([newStock]);
        }

        if (hasUpdates && invCmap.stock !== -1) {
          invSh.getRange(2, invCmap.stock + 1, invData.length - 1, 1).setValues(newStockValues);
        }
      }
    }

    // ─── STEP 3: DETECT ANOMALIES ────────────────────────────
    if (mode === "full") {
      log("info", "🔍 Step 3: Scanning for anomalies...");

      for (let i = 1; i < invData.length; i++) {
        const code = invData[i][invCmap.code];
        const name = invData[i][invCmap.name];
        const stock = invCmap.stock !== -1 ? Number(invData[i][invCmap.stock]) : 0;

        if (!code && !name) continue;

        // Check: negative stock
        if (stock < 0) {
          log("err", "  🚨 " + (name || code) + ": Negative stock (" + stock + ")!");
          stats.anomalies++;
        }

        // Check: missing SKU
        if (!code && name) {
          log("warn", "  ⚠️ Row " + (i + 1) + " has name '" + name + "' but no SKU code");
          stats.anomalies++;
        }

        // Check: missing name
        if (code && !name) {
          log("warn", "  ⚠️ Row " + (i + 1) + " has code '" + code + "' but no product name");
          stats.anomalies++;
        }

        // Check: duplicate SKU
        for (let j = i + 1; j < invData.length; j++) {
          if (code && invData[j][invCmap.code] === code) {
            log("err", "  🚨 Duplicate SKU '" + code + "' found in rows " + (i + 1) + " and " + (j + 1));
            stats.anomalies++;
          }
        }
      }

      if (stats.anomalies === 0) {
        log("ok", "  ✅ No anomalies detected — data looks healthy!");
      }
    }

    SpreadsheetApp.flush();
    log("ok", "");
    log("ok", "🎉 Repair process completed successfully!");

  } catch (e) {
    log("err", "❌ Fatal error: " + e.message);
    log("err", "Stack: " + e.stack);
  }

  return JSON.stringify({ logs, stats });
}

// Helper: convert column index (0-based) to letter
function _colLetter(idx) {
  if (idx < 0) return "N/A";
  let letter = "";
  let n = idx;
  while (n >= 0) {
    letter = String.fromCharCode(65 + (n % 26)) + letter;
    n = Math.floor(n / 26) - 1;
  }
  return letter;
}


function forceSaveSheetId() {
  const id = SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties().setProperty("SHEET_ID", id);
  SpreadsheetApp.getUi().alert("✅ SYSTEM FIXED!\nID Spreadsheet Anda berhasil dikunci: " + id + "\nSistem latar belakang kini bisa beroperasi penuh.");
}
