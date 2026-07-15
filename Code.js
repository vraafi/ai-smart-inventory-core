let _clientProps = null;

function setClientProperties(propsService) {
  _clientProps = propsService;
}

function _getScriptProps() {
  if (_clientProps) return _clientProps.getScriptProperties();
  return PropertiesService.getScriptProperties();
}
/**
 * Automated E-commerce Inventory System (BASIC PACKAGE)
 * =======================================================
 * File: Code.js
 * Description: Handles UI triggers, custom menus,// Core Library Entry Point - Force Updates.
 */

// Runs automatically when the spreadsheet is opened
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 Inventory System')
    // ── Navigation ──
    .addItem('🏠 Go to Dashboard',              'navDashboard')
    .addItem('📦 View Inventory',               'navInventory')
    .addItem('🔄 View Transactions',            'navTransactions')
    .addItem('🏢 Manage Branches',              'navBranches')
    .addSeparator()
    // ── Smart Action Center ──
    .addItem('💬 Smart Action Center',           'openSmartChatUI')
    .addSeparator()
    // ── Stock Alerts ──
    .addItem('⚠️ Check Low Stock (Popup)',       'checkLowStockAlerts')
    .addItem('📲 Send Low Stock to Telegram',    'sendLowStockToTelegram')
    .addSeparator()
    // ── Email & Reports ──
    .addItem('📧 Process Inbox Emails (AI)',     'runInboxWatcher')
    .addItem('📨 Manual Poll Emails',            'pollEmails')
    .addItem('🤖 Aktifkan Email Otomatis',       'installEmailTrigger')
    .addItem('🧪 Test Email Integration',         'testEmailIntegration')
    .addItem('📊 Generate Monthly Report (AI)',  'runMonthlyAnalysis')
    .addSeparator()
    // ── Dashboard & Repair ──
    .addItem('🔄 Refresh Dashboard',             'refreshDashboardClick')
    .addItem('📊 Recalculate All Status',        'recalcAllStatus')
    .addItem('🛠️ AI Smart Repair',               'showRepairDialog')
    .addSeparator()
    // ── Configuration ──
    .addItem('🤖 Konfigurasi AI (Universal)',    'openAiConfigUI')
    .addItem('📱 AI Config / Agent Settings',    'showAgentSettings')
    .addItem('⚙️ Set Telegram Credentials',      'promptTelegramCredentials')
    .addItem('⚙️ Set WhatsApp Config (Meta)',     'openWaConfigUI')
    .addSeparator()
    // ── Security & Admin ──
    .addItem('🔑 License Manager (Admin)',       'showLicenseManager')
    .addSeparator()
    // ── System Tools ──
    .addItem('🔧 Fix Background Error',          'forceSaveSheetId')
    .addItem('🧪 Generate Test Data (Enterprise)','generateTestData')
    .addItem('🐛 Debug Error Message',           'debugErrorsUI')
    .addItem('⚙️ Run Automated Tests',            'runAllTests')
    .addItem('Debug System',                      'debugInventory')
    .addSeparator()
    .addItem('Aktivasi Lisensi',                  'showLicenseDialog')
    .addItem('ℹ️ About',                          'showAbout')
    .addToUi();
}

function openSmartChatUI() {
  LicenseClient.require();
  const html = HtmlService.createHtmlOutputFromFile('SmartChatUI')
      .setWidth(450)
      .setHeight(650)
      .setTitle('💬 Smart Action Center');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Prompt user to add a new item
function promptNewItem() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Add New Item',
    'Enter Item Name:',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const itemName = result.getResponseText();
    if (itemName.trim() !== '') {
      const newId = InventoryCore.generateItemId();
      InventoryCore.insertNewItem(newId, itemName);
      ui.alert('Success', `Item "${itemName}" added with ID: ${newId}`, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Item name cannot be empty.', ui.ButtonSet.OK);
    }
  }
}

// Wrapper for recalculating stock
function recalculateStock() {
  const ui = SpreadsheetApp.getUi();
  try {
    InventoryCore.updateAllCurrentStock();
    ui.alert('Success', 'Stock recalculated successfully based on In/Out entries.', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'Failed to recalculate stock: ' + error.message, ui.ButtonSet.OK);
  }
}

// Wrapper for checking low stock alerts
function checkLowStockAlerts() {
  const ui = SpreadsheetApp.getUi();
  const lowItems = InventoryCore.getLowStockItems();

  if (lowItems.length > 0) {
    let msg = 'The following items are low on stock:\n\n';
    lowItems.forEach(item => {
      msg += `- [${item.id}] ${item.name} (Remaining: ${item.stock})\n`;
    });
    ui.alert('⚠️ Low Stock Alert', msg, ui.ButtonSet.OK);
  } else {
    ui.alert('✅ Stock Status', 'All items are adequately stocked.', ui.ButtonSet.OK);
  }
}

// DEBUG FUNCTION
function debugInventory() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();

    // Test Header
    const searchRange = sheet.getRange(1, 1, 10, 2).getValues();
    let foundHeader = -1;
    let headerStr = "";
    for(let i=0; i<searchRange.length; i++){
      let a = String(searchRange[i][0]).toLowerCase().replace(/\s+/g, '');
      if(a.includes('itemid')) { foundHeader = i+1; headerStr = String(searchRange[i][0]); break; }
    }

    const headerRow = InventoryCore.findHeaderRow(sheet);
    const startRow = headerRow + 1;

    let msg = `Last Row: ${lastRow}\nMy HeaderRow: ${headerRow} (New test found: ${foundHeader} val: '${headerStr}')\nStart Row: ${startRow}\n\n`;

    if (lastRow >= 2) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      msg += `Scanning Rows 2 to ${lastRow}:\n`;

      dataRange.forEach((row, idx) => {
        const actualRow = idx + 2;
        const id = String(row[0]);
        const stock = row[6];
        if (id.includes('ITM') || stock !== "") {
          msg += `R${actualRow} -> ID: '${id}', Stock: '${stock}' (${typeof stock})\n`;
          let nStock = Number(stock);
          let cond = (id !== "") && !isNaN(nStock) && (nStock > 0) && (nStock <= 5);
          msg += `   Cond: ${cond} (id!=="", !isNaN, >0, <=5)\n`;
        }
      });
    } else {
      msg += "No data rows found.";
    }

    ui.alert('Debug Info V2', msg, ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('Debug Error', e.message, ui.ButtonSet.OK);
  }
}

// PREMIUM FEATURE: Open Secure Transaction UI
function openTransactionUI() {
  const html = HtmlService.createHtmlOutputFromFile('TransactionUI')
      .setWidth(400)
      .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, '🔒 Secure Transaction Input');
}

// PREMIUM FEATURE: Process Transaction from Frontend
function processTransaction(data) {
  try {
    const userEmail = Session.getActiveUser().getEmail() || "Unknown User";
    const branch = data.branch || 'Pusat';
    InventoryCore.logAndProcessTransaction(branch, data.itemId, data.type, data.quantity, data.notes, userEmail);

    // Check stock and send telegram if needed
    checkAndSendTelegramSilent();

    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ULTRA-PREMIUM: Open Universal AI Configuration (delegates to UniversalAI.js)
function openAiConfigUI() {
  showAIProviderSettings();
}

// WhatsApp Config UI
function openWaConfigUI() {
  const html = HtmlService.createHtmlOutputFromFile('WaConfigUI')
      .setWidth(450)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '⚙️ WhatsApp Cloud API Config');
}

// Save WhatsApp Config
function saveWaConfig(accessToken, phoneId, verifyToken) {
  try {
    WhatsAppService.saveCredentials(accessToken, phoneId, verifyToken);
    return { success: true };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// Load saved WhatsApp Config
function loadWaConfig() {
  try {
    var creds = WhatsAppService.getCredentials();
    return { success: true, accessToken: creds.accessToken, phoneId: creds.phoneId, verifyToken: creds.verifyToken };
  } catch(e) {
    return { success: false };
  }
}

// Load saved Telegram Config
function loadTelegramConfig() {
  try {
    var props = _getScriptProps();
    return {
      success: true,
      botToken: props.getProperty('TELEGRAM_BOT_TOKEN') || '',
      chatId: props.getProperty('TELEGRAM_CHAT_ID') || ''
    };
  } catch(e) {
    return { success: false };
  }
}

// ULTRA-PREMIUM: Open AI Input UI
function openAiInputUI() {
  const html = HtmlService.createHtmlOutputFromFile('AiInputUI')
      .setWidth(450)
      .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '🧠 AI Autonomous Input');
}

// ULTRA-PREMIUM: Process Autonomous Input via Gemini
function processAutonomousInput(rawText) {
  try {
    LicenseClient.require();
    const parsedData = AIAgent.parseTransaction(rawText);

    // Clarification fallback
    if (parsedData && !Array.isArray(parsedData) && parsedData.cmd === "ASK_USER") {
       return { success: false, isQuestion: true, message: parsedData.question };
    }

    const userEmail = Session.getActiveUser().getEmail() || "AI Autonomous Agent";

    // Handle both single object and array of transactions
    const transactions = Array.isArray(parsedData) ? parsedData : [parsedData];

    // 1. Process all transactions and build report
    let reportMsg = `🤖 <b>Laporan AI (Groq/Gemini)</b>\n`;
    reportMsg += `Memproses ${transactions.length} transaksi:\n\n`;
    let successCount = 0;

    transactions.forEach((tx, index) => {
      try {
        const branch = tx.branch || 'Pusat';
        InventoryCore.logAndProcessTransaction(
          branch,
          tx.itemId,
          tx.type,
          tx.quantity,
          tx.notes,
          userEmail
        );
        successCount++;
        let icon = tx.type === 'IN' ? '🟢' : '🔴';
        reportMsg += `${index + 1}. [${branch}] ${icon} <b>${tx.type}</b>: ${tx.quantity} unit [<code>${tx.itemId}</code>]\n`;
        reportMsg += `   <i>Catatan: ${tx.notes}</i>\n`;

        // PREMIUM FEATURE: Auto-Invoicing
        if (tx.type === 'OUT' && tx.customerEmail) {
          InvoiceService.sendInvoice(tx.customerEmail, tx);
          reportMsg += `   📧 <i>Invoice sent to: ${tx.customerEmail}</i>\n`;
        }
      } catch (err) {
        reportMsg += `${index + 1}. ❌ <b>GAGAL</b>: ${tx.quantity} unit [<code>${tx.itemId}</code>]\n`;
        reportMsg += `   <i>Alasan: ${err.message}</i>\n`;
      }
    });

    reportMsg += `\n<i>Diproses otomatis oleh sistem (${successCount}/${transactions.length} berhasil).</i>`;



    try {
      TelegramService.sendMessage(reportMsg);
    } catch(err) {
      console.error("Gagal mengirim laporan Telegram: " + err.message);
    }

    // 3. Check low stock alerts
    checkAndSendTelegramSilent();

    return { success: true, parsed: parsedData, count: transactions.length };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Send low stock report to Telegram (Triggered manually via Menu)
function sendLowStockToTelegram() {
  const ui = SpreadsheetApp.getUi();
  const lowItems = InventoryCore.getLowStockItems();

  if (lowItems.length === 0) {
    ui.alert('✅ All Good', 'No low stock items to report.', ui.ButtonSet.OK);
    return;
  }

  let message = '🚨 <b>LOW STOCK ALERT</b> 🚨\n';
  message += 'The following items require your attention:\n\n';

  lowItems.forEach(item => {
    message += `🏢 <b>Cabang: ${item.branch}</b>\n📦 <b>${item.name}</b>\nID: <code>${item.id}</code>\nRemaining: <b>${item.stock}</b>\n\n`;
  });

  message += '<i>Sent from Automated Inventory System</i>';

  try {
    TelegramService.sendMessage(message);
    ui.alert('✅ Success', 'Low stock alert sent to Telegram.', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error', 'Failed to send to Telegram: ' + e.message, ui.ButtonSet.OK);
  }
}

// Background Silent Trigger for AI and Transaction UI
function checkAndSendTelegramSilent() {
  const lowItems = InventoryCore.getLowStockItems();
  if (lowItems.length > 0) {
    let message = '🚨 <b>AUTO SYSTEM ALERT</b> 🚨\n';
    message += 'Stock dropped below threshold after recent transaction:\n\n';
    lowItems.forEach(item => {
      message += `🏢 <b>Cabang: ${item.branch}</b>\n📦 <b>${item.name}</b>\nID: <code>${item.id}</code>\nRemaining: <b>${item.stock}</b>\n\n`;
    });
    try {
      TelegramService.sendMessage(message);
    } catch(e) {
      console.error("Silent Telegram Error: " + e.message);
    }
  }
}

// LEVEL 3: Inbox Watcher (Process emails automatically)
function processIncomingEmails() {
  try {
    // Search for unread emails with a specific label or subject.
    // For general use, we search for a label "inventory-ai"
    const labelName = 'inventory-ai';
    let label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      label = GmailApp.createLabel(labelName);
    }

    const threads = label.getThreads(0, 10);
    let processedCount = 0;

    threads.forEach(thread => {
      if (thread.isUnread()) {
        const messages = thread.getMessages();
        const lastMessage = messages[messages.length - 1];
        let body = lastMessage.getPlainBody();
        const sender = lastMessage.getFrom();
        const attachments = lastMessage.getAttachments();

        attachments.forEach(att => {
          const mimeType = att.getContentType();
          const name = att.getName();

          if (mimeType === 'text/csv' || name.toLowerCase().endsWith('.csv')) {
            body += `\n\n[Attachment: ${name}]\n`;
            body += att.getDataAsString();
          }
          else if (mimeType === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                   mimeType === 'application/vnd.ms-excel' ||
                   name.toLowerCase().endsWith('.xlsx') ||
                   name.toLowerCase().endsWith('.xls')) {
            try {
              const blob = att.copyBlob();
              const fileConfig = { title: "Temp_AI_Excel", mimeType: MimeType.GOOGLE_SHEETS };
              const newFile = Drive.Files.insert(fileConfig, blob);

              const tempSs = SpreadsheetApp.openById(newFile.id);
              const tempSheet = tempSs.getSheets()[0];
              const data = tempSheet.getDataRange().getDisplayValues(); // use DisplayValues to keep formatting

              body += `\n\n[Attachment Excel: ${name}]\n`;
              data.forEach(row => {
                body += row.join(" | ") + "\n";
              });

              DriveApp.getFileById(newFile.id).setTrashed(true);
            } catch (err) {
              console.error("Failed to parse Excel attachment: " + err.message);
              body += `\n\n[Failed to read Excel Attachment: ${name} - ${err.message}]`;
            }
          }
        });

        // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
                    // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
                    // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
                    // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
                    // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
                    // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
                    // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
                    // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
                    // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
                    // Extract sender email specifically
        const senderEmail = (sender.match(/<(.+)>/) || [, sender])[1];
        const senderName = sender.replace(/<.*>/g, "").trim() || sender;
        const subject = lastMessage.getSubject();
        const fullContext = "Subject: " + subject + "\n\nBody: " + body;

        // Route to the central AI Agent (Universal Routing)
        try {
          if (typeof _processWithAI !== "undefined") {
            _processWithAI(senderEmail, fullContext, senderName, "Email", null);
            thread.markRead();
            processedCount++;
          } else {
            console.error("AI Agent (_processWithAI) not found. Using fallback.");
            const result = processAutonomousInput("From Email (" + sender + "):\n" + body);
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
            if (result.success) {
              thread.markRead();
              processedCount++;
            }
          }
        } catch (err) {
          console.error("Error routing email to AI: " + err.message);
        }
      }
    });

    if (processedCount > 0) {
      console.log(`Processed ${processedCount} email threads.`);
    }
  } catch (e) {
    console.error("Error in processIncomingEmails: " + e.message);
  }
}

// LEVEL 3: Inbox Watcher Menu Wrapper
function runInboxWatcher() {
  LicenseClient.require();
  const ui = SpreadsheetApp.getUi();
  try {
    processIncomingEmails();
    ui.alert('✅ Inbox Watcher', 'Email processing complete. Check Telegram for results.\n\nTip: You can set up a Time-Driven Trigger to run this every hour automatically!\n(Extensions > Apps Script > Triggers)', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error', 'Failed to process emails: ' + e.message, ui.ButtonSet.OK);
  }
}

// LEVEL 3: Monthly Analysis Menu Wrapper
function runMonthlyAnalysis() {
  LicenseClient.require();
  const ui = SpreadsheetApp.getUi();
  try {
    const result = AnalysisService.generateMonthlyReport();
    ui.alert('✅ Analysis Complete', 'Monthly report has been sent to Telegram!\n\nCheck your Telegram for the full AI-powered analysis.', ui.ButtonSet.OK);
  } catch(e) {
    ui.alert('❌ Error', 'Failed to generate report: ' + e.message, ui.ButtonSet.OK);
  }
}

// =======================================================
// WEBHOOK ENDPOINTS (WhatsApp Official API)
// =======================================================

/**
 * Handle HTTP GET - Used by Meta for Webhook Verification
 */
function doGet(e) {
  const creds = WhatsAppService.getCredentials();

  if (e.parameter['hub.mode'] === 'subscribe' && e.parameter['hub.verify_token'] === creds.verifyToken) {
    return ContentService.createTextOutput(e.parameter['hub.challenge']);
  }

  return ContentService.createTextOutput("Forbidden: Invalid Verify Token").setStatusCode(403);
}



// LEVEL 3: Auto-Install Trigger for Email
function installEmailTrigger() {
  const ui = SpreadsheetApp.getUi();
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let exists = false;
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === "processIncomingEmails") {
        exists = true;
        break;
      }
    }

    if (exists) {
      ui.alert("✅ Sudah Aktif", "Sistem Auto-Pilot Email sudah aktif sebelumnya!", ui.ButtonSet.OK);
      return;
    }

    ScriptApp.newTrigger("processIncomingEmails")
      .timeBased()
      .everyHours(1)
      .create();

    ui.alert("🚀 Berhasil!", "Sistem Auto-Pilot Email berhasil diaktifkan.\nAI akan mengecek kotak masuk email Anda setiap 1 jam sekali secara otomatis tanpa Anda perlu menekan tombol apa pun.", ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("� � Error", "Gagal mengaktifkan trigger: " + e.message, ui.ButtonSet.OK);
  }
}



function showOnboardingDialog() {
  const html = HtmlService.createHtmlOutputFromFile('OnboardingUI')
    .setWidth(800)
    .setHeight(600)
    .setTitle('Registrasi Barang Baru (Smart Onboarding)');
  SpreadsheetApp.getUi().showModalDialog(html, 'Registrasi Barang Baru');
}

function processOnboarding(formObj) {
  if (!formObj.new_item_name) throw new Error("Nama Barang wajib diisi");

  const parsed = {
    item_new: true,
    new_item_name: formObj.new_item_name,
    new_category: formObj.new_category || "General",
    new_price: formObj.new_price || 0,
    quantity: formObj.quantity || 0
  };

  const rowObj = _createNewItemRow(parsed);
  if (rowObj && rowObj.row) {
    const sh = _getSheet(SHEETS.INVENTORY);
    if (sh) {
       sh.getRange(rowObj.row, 6).setValue(formObj.quantity || 0); // stock
       sh.getRange(rowObj.row, 12).setValue(formObj.buyPrice || 0); // buy price
       sh.getRange(rowObj.row, 13).setValue(formObj.new_price || 0); // sell price
    }
    return { success: true, itemName: formObj.new_item_name };
  }
  return { success: false, message: "Gagal membuat item." };
}

// Force Clasp Sync

// Force Clasp Sync 2

// --- PENGUJIAN EMAIL KE AI ---
function testEmailIntegration() {
  const ui = SpreadsheetApp.getUi();
  try {
    LicenseClient.require();
    if (typeof _processWithAI === "undefined") {
       throw new Error("Fungsi _processWithAI tidak ditemukan di sistem. Pastikan file Agent tersedia.");
    }

    // Simulasi payload email
    const dummyEmail = Session.getActiveUser().getEmail() || "test@example.com";
    const dummySubject = "Laporan Stok Dummy";
    const dummyBody = "Ada barang masuk 10 unit laptop asus";
    const fullContext = "Subject: " + dummySubject + "\n\nBody: " + dummyBody;

    // Meneruskan ke AI Agent
    _processWithAI(dummyEmail, fullContext, "Penguji Sistem", "Email", null);

    ui.alert("✅ Pengujian Berhasil", "Sistem berhasil meneruskan pesan simulasi ke AI.\n\nSilakan periksa kotak masuk Email Anda atau chat Telegram untuk melihat balasan dari AI.", ui.ButtonSet.OK);
  } catch (err) {
    ui.alert("❌ Pengujian Gagal", "Pesan Error: " + err.message, ui.ButtonSet.OK);
  }
}
