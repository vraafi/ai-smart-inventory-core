// ============================================================
//  AI SMART INVENTORY — AGENT MODULE
//  WhatsApp (Meta Cloud API) + Telegram + Email + Gemini AI
//  Paste this as a SECOND file in Apps Script
// ============================================================

function debugLog(msg) {
  try {
    const ss = _getSpreadsheet();
    if (!ss) return;
    let sheet = ss.getSheetByName("DebugLogs");
    if (!sheet) {
      sheet = ss.insertSheet("DebugLogs");
      sheet.appendRow(["Timestamp", "Message"]);
    }
    sheet.insertRowBefore(2);
    sheet.getRange("A2:B2").setValues([[new Date(), typeof msg === 'object' ? JSON.stringify(msg) : msg]]);
  } catch(e) {
    Logger.log("debugLog error: " + e);
  }
}

// ─── WEBHOOK ENTRY POINT ─────────────────────────────────────
function doPost(e) {
  debugLog("doPost Triggered! Event: " + JSON.stringify(e));
  try {
    LicenseClient.require();
    const props = _getScriptProps();
    const secret = props.getProperty("WEBHOOK_SECRET");
    if (!secret) {
      return ContentService.createTextOutput("500 Internal Server Error: Webhook secret is not configured on the server.");
    }
    if (e.parameter.token !== secret) {
      return ContentService.createTextOutput("403 Forbidden: Invalid Webhook Token");
    }

    const body = JSON.parse(e.postData.contents);

    // ── WhatsApp Cloud API webhook ──
    if (body.object === "whatsapp_business_account") {
      const entry = body.entry?.[0];
      const change = entry?.changes?.[0];
      const msg = change?.value?.messages?.[0];
      if (msg) {
        const phone = msg.from;
        const text  = msg.type === "text" ? msg.text.body : (msg.type === "button" ? msg.button.text : "");
        const name  = change?.value?.contacts?.[0]?.profile?.name || "Employee";
        if (text) _processWhatsAppMessage(phone, text, name);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Telegram webhook ──
    if (body.update_id) {
      const updateIdStr = body.update_id.toString();
      const cache = CacheService.getScriptCache();
      if (cache.get("TG_UPDATE_" + updateIdStr)) {
        debugLog("Duplicate TG Update ID ignored: " + updateIdStr);
        return HtmlService.createHtmlOutput("OK");
      }
      cache.put("TG_UPDATE_" + updateIdStr, "1", 3600); // Store for 1 hour

      const msg = body.message;
      if (msg) {
        const chatId = msg.chat.id.toString();
        const text   = msg.text || "";
        const from   = msg.from;
        const name   = [from.first_name, from.last_name].filter(Boolean).join(" ") || from.username || "Employee";
        
        if (text || msg.document) {
          debugLog("Processing TG Message from " + name + (msg.document ? " [Document attached]" : "") + ": " + text);
          _processTelegramMessage(chatId, text, name, msg.document);
        }
      }
    }
  } catch (err) {
    Logger.log("doPost error: " + err);
    debugLog("doPost error: " + err + " | " + err.stack);
    return HtmlService.createHtmlOutput("ERROR: " + err);
  }
  return HtmlService.createHtmlOutput("OK");
}

// ─── ACCESS CONTROL & WHITELIST ─────────────────────────────────
function _checkAccess(idStr, text, platform) {
  const props = _getScriptProps();
  const adminId = props.getProperty("ADMIN_CHAT_ID");

  if (!adminId) {
    props.setProperty("ADMIN_CHAT_ID", idStr);
    props.setProperty("WHITELIST_IDS", idStr);
    debugLog("Saved new ADMIN_CHAT_ID: " + idStr);
    return true;
  }

  const isPublic = props.getProperty("PUBLIC_MODE") === "ON";
  const whitelistStr = props.getProperty("WHITELIST_IDS") || adminId;
  const whitelist = whitelistStr.split(",").map(s => s.trim());
  const isAllowed = isPublic || whitelist.includes(idStr);

  if (idStr === adminId) {
    if (text.startsWith("/allow ")) {
      const newId = text.replace("/allow ", "").trim();
      if (!whitelist.includes(newId)) {
        whitelist.push(newId);
        props.setProperty("WHITELIST_IDS", whitelist.join(","));
        _sendAgentMsg(platform, idStr, `✅ ID ${newId} berhasil didaftarkan ke sistem.`);
      } else {
        _sendAgentMsg(platform, idStr, `ℹ️ ID ${newId} sudah ada di dalam whitelist.`);
      }
      return "ADMIN_CMD";
    }
    if (text.startsWith("/block ")) {
      const targetId = text.replace("/block ", "").trim();
      if (targetId === adminId) {
        _sendAgentMsg(platform, idStr, `❌ Tidak dapat memblokir diri sendiri.`);
        return "ADMIN_CMD";
      }
      const newWhitelist = whitelist.filter(w => w !== targetId);
      props.setProperty("WHITELIST_IDS", newWhitelist.join(","));
      _sendAgentMsg(platform, idStr, `🚫 ID ${targetId} telah dicabut aksesnya.`);
      return "ADMIN_CMD";
    }
    if (text === "/users") {
      _sendAgentMsg(platform, idStr, `📋 *Daftar Akses (Whitelist):*\n- ` + whitelist.join("\n- ") + `\n\nMode Publik: ${isPublic ? "ON 🔓" : "OFF 🔒"}`);
      return "ADMIN_CMD";
    }
    if (text === "/public_mode on") {
      props.setProperty("PUBLIC_MODE", "ON");
      _sendAgentMsg(platform, idStr, `🔓 Mode Publik DIAKTIFKAN.`);
      return "ADMIN_CMD";
    }
    if (text === "/public_mode off") {
      props.setProperty("PUBLIC_MODE", "OFF");
      _sendAgentMsg(platform, idStr, `🔒 Mode Publik DIMATIKAN.`);
      return "ADMIN_CMD";
    }
  }

  if (!isAllowed) {
    _sendAgentMsg(platform, idStr, `⛔ *Akses Ditolak*\nAnda tidak terdaftar di sistem ini.\nID Anda: \`${idStr}\`\n\nBerikan ID ini kepada Admin untuk didaftarkan.`);
    return false;
  }
  return true;
}

// ─── WHATSAPP MESSAGE HANDLER ──────────────────────────────
function _processWhatsAppMessage(phone, text, senderName) {
  const access = _checkAccess(phone.toString(), text || "", "whatsapp");
  if (access === "ADMIN_CMD" || access === false) return;

  if (text === "/start" || text === "/help") {
    _sendWhatsApp(phone, _buildWelcomeMessage(senderName));
    return;
  }
  if (text === "/status" || text === "/dashboard") {
    _sendAgentMsg("whatsapp", phone, _buildDashboardSummary());
    return;
  }
  if (text.startsWith("/check") || text.startsWith("/stock")) {
    const q = text.replace(/\/(check|stock)\s*/i, "").trim();
    _sendAgentMsg("whatsapp", phone, _buildStockCheckResult(q));
    return;
  }
  // Natural language — process with AI
  _sendWhatsApp(phone, "⏳ AI is processing your report...");
  _processWithAI(phone, text, senderName, "WhatsApp");
}

// ─── TELEGRAM MESSAGE HANDLER ────────────────────────────────
function _processTelegramMessage(chatId, text, senderName, document) {
  let finalContext = text || "";
  const access = _checkAccess(chatId.toString(), finalContext, "telegram");
  if (access === "ADMIN_CMD" || access === false) return;
  
  let bulkDataArray = null;
  if (document) {
    _sendTelegram(chatId, "⏳ Membaca dokumen Excel/CSV...");
    try {
      const props = _getScriptProps();
      const token = props.getProperty("TG_TOKEN");
      const fileResp = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/getFile?file_id=${document.file_id}`);
      const fileData = JSON.parse(fileResp.getContentText());
      if (fileData.ok) {
        const filePath = fileData.result.file_path;
        const fileUrl = `https://api.telegram.org/file/bot${token}/${filePath}`;
        const blobResp = UrlFetchApp.fetch(fileUrl);
        let blob = blobResp.getBlob();
        blob.setName(document.file_name || "document");
        blob.setContentType(document.mime_type || "");
        
        const mime = blob.getContentType();
        const name = blob.getName().toLowerCase();

        if (name.endsWith(".csv") || name.endsWith(".xlsx") || mime.includes("csv") || mime.includes("spreadsheetml") || mime.includes("excel")) {
            bulkDataArray = _getRawDataArrayFromBlob(blob);
            if (bulkDataArray && bulkDataArray.length > 1) {
                finalContext += `\n\n[FILE ATTACHMENT RECEIVED: ${document.file_name} - Processing as Bulk Import]`;
            } else {
                bulkDataArray = null;
                const extractedText = _extractTextFromBlob(blob);
                if (extractedText) finalContext += `\n\n[FILE ATTACHMENT CONTENT: ${document.file_name}]\n` + extractedText;
            }
        } else {
            const extractedText = _extractTextFromBlob(blob);
            if (extractedText) finalContext += `\n\n[FILE ATTACHMENT CONTENT: ${document.file_name}]\n` + extractedText;
        }
      }
    } catch(e) {
      _sendTelegram(chatId, "❌ Gagal mengekstrak dokumen: " + e.message);
      return;
    }
  }

  if (finalContext === "/start" || finalContext === "/help") {
    _sendTelegram(chatId, _buildWelcomeMessage(senderName));
    return;
  }
  if (finalContext === "/status" || finalContext === "/dashboard") {
    _sendAgentMsg("telegram", chatId, _buildDashboardSummary());
    return;
  }
  if (finalContext.startsWith("/check") || finalContext.startsWith("/stock")) {
    const q = finalContext.replace(/\/(check|stock)\s*/i, "").trim();
    _sendAgentMsg("telegram", chatId, _buildStockCheckResult(q));
    return;
  }

  if (!bulkDataArray && !finalContext.match(/https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/)) {
     _sendTelegram(chatId, "⏳ AI is processing your report...");
  }
  _processWithAI(chatId, finalContext, senderName, "Telegram", bulkDataArray);
}

// ─── WHATSAPP POLLING (backup if webhook fails) ──────────────
function pollWhatsApp() {
  // WhatsApp Cloud API doesn't support polling — use webhook only
  // This trigger is intentionally a no-op
}

// ─── TELEGRAM POLLING (backup for webhook) ───────────────────
function pollTelegram() {
  const props = _getScriptProps();
  const token = props.getProperty("TG_TOKEN");
  if (!token) return;
  let offset = parseInt(props.getProperty("TG_OFFSET") || "0");
  try {
    const resp = UrlFetchApp.fetch(
      `https://api.telegram.org/bot${token}/getUpdates?offset=${offset}&limit=10&timeout=0`,
      { muteHttpExceptions: true }
    );
    const data = JSON.parse(resp.getContentText());
    if (!data.ok || !data.result.length) return;
    data.result.forEach(update => {
      if (update.message) {
        const msg = update.message;
        const chatId = msg.chat.id.toString();
        const text   = msg.text || "";
        const from   = msg.from;
        const name   = [from.first_name, from.last_name].filter(Boolean).join(" ") || "Employee";
        if (text || msg.document) {
          debugLog("Processing TG Message from " + name + (msg.document ? " [Document attached]" : "") + ": " + text);
          try {
            _processTelegramMessage(chatId, text, name, msg.document);
          } catch(e) {
            debugLog("Error processing msg: " + e);
          }
        }
      }
      offset = update.update_id + 1;
    });
    props.setProperty("TG_OFFSET", offset.toString());
  } catch (err) { debugLog("pollTelegram error: " + err); }
}

// ─── EMAIL POLLING ────────────────────────────────────────────
function pollEmails() {
  LicenseClient.require();
  debugLog("pollEmails trigger fired!");
  try {
    const threads = GmailApp.search(`is:unread in:anywhere (subject:"Nexus AI Report" OR subject:"AI Inventory Notification" OR subject:"Inventory" OR subject:"Update" OR subject:"Laporan" OR subject:"Stok" OR subject:"Stock")`, 0, 5);
    const props = PropertiesService.getScriptProperties();
    let processed = props.getProperty("PROCESSED_EMAILS") || "";
    
    debugLog(`Found ${threads.length} unread threads matching query.`);
    
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        if (!msg.isUnread()) {
          return;
        }
        const msgId = msg.getId();
        
        // Skip if already processed in this property string
        if (processed.includes(msgId)) continue;

        if (threadProcessedCount >= 1) {
           // We only process the 1 newest unread message per thread to prevent infinite reply loops.
           // Mark older unread messages in the same thread as read so they don't clog future polls.
           msg.markRead();
           continue;
        }
        
        debugLog(`Processing NEW UNREAD message: ${msgId}`);
        
        const subject   = msg.getSubject();

        // Anti Infinite-Loop: Skip emails sent by the AI itself, but allow user replies (which start with Re: or Fwd:)
        const lowerSubject = subject.toLowerCase();
        if ((subject.includes("AI Inventory Notification") || subject.includes("Nexus AI Report"))
            && !lowerSubject.startsWith("re:")
            && !lowerSubject.startsWith("fwd:")) {
           msg.markRead();
           continue;
        }

        const rawBody   = msg.getPlainBody().substring(0, 1500);
        const from      = msg.getFrom();
        const senderName = from.replace(/<.*>/g, "").trim() || from;
        const senderEmail = (from.match(/<(.+)>/) || [, from])[1];
        
        // Combine Subject and Body for AI Context
        let fullContext = `Subject: ${subject}\n\nBody: ${rawBody}`;

        // Extract Attachments
        const attachments = msg.getAttachments();
        let bulkDataArray = null;
        if (attachments.length > 0) {
          try {
             for (let j = 0; j < attachments.length; j++) {
                const mime = attachments[j].getContentType();
                const name = attachments[j].getName().toLowerCase();
                if (name.endsWith(".csv") || name.endsWith(".cvs") || name.endsWith(".xlsx") || mime.includes("csv") || mime.includes("spreadsheetml") || mime.includes("excel")) {
                    if (!bulkDataArray) {
                        bulkDataArray = _getRawDataArrayFromBlob(attachments[j]);
                    }
                    if (bulkDataArray && bulkDataArray.length > 1) {
                        fullContext += `\n\n[FILE ATTACHMENT RECEIVED: ${attachments[j].getName()} - Processing as Bulk Import]`;
                    } else {
                        bulkDataArray = null;
                        let extractedText = _extractTextFromBlob(attachments[j]);
                        if (extractedText) fullContext += `\n\n[FILE ATTACHMENT CONTENT: ${attachments[j].getName()}]\n` + extractedText;
                    }
                } else {
                    let extractedText = _extractTextFromBlob(attachments[j]);
                    if (extractedText) fullContext += `\n\n[FILE ATTACHMENT CONTENT: ${attachments[j].getName()}]\n` + extractedText;
                }
             }
          } catch(e) {
             debugLog("Failed to extract email attachments: " + e.message);
          }
        }

        // Process with AI and pass the senderEmail so it can reply
        try {
          _processWithAI(senderEmail, fullContext, senderName, "Email");
        } catch(aiErr) {
          debugLog(`AI Processing failed for msg ${msgId}: ${aiErr.message}`);
        }
        
        // Mark as processed using PropertiesService and also mark Read in Gmail for hygiene
        processed += "," + msgId;
        props.setProperty("PROCESSED_EMAILS", processed.slice(-8000));
        msg.markRead();

        threadProcessedCount++;
        globalProcessedCount++;
      }
    }
  } catch (err) { debugLog("pollEmails error: " + err); }
}

// ─── CORE AI PROCESSING ──────────────────────────────────────

function _logSecurityAudit(source, senderName, senderId, messageText) {
  try {
    const ss = _getSpreadsheet();
    let sheet = ss.getSheetByName("Security Audit");
    if (sheet) {
      const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
      sheet.appendRow([timestamp, source, senderName || "Unknown", senderId || "-", messageText || "-"]);
    }
  } catch(e) {
    // Abaikan jika sheet dikunci atau error lainnya
  }
}

function _processWithAI(chatId, rawText, senderName, source, bulkDataArray = null) {
  _logSecurityAudit(source, senderName, chatId, rawText);
  const lowerText = rawText.toLowerCase();

  if (lowerText.includes("/wipe") || lowerText.includes("/format") || lowerText.includes("/onboarding")) {
    _sendAgentMsg(source, chatId, "❌ FITUR DIBATASI (RESTRICTED FEATURE)\\n\\nFitur registrasi barang baru, hapus data, dan format tabel telah dipindahkan secara fisik ke UI Google Sheets demi keamanan Enterprise.\\n\\nSilakan minta Admin/Bos untuk membuka Google Sheets dan menggunakan menu '📦 Smart Inventory'.");
    return;
  }

  // Handle Smart Bulk Import (Google Sheets URL or Bulk Data Array from CSV/Excel)
  const urlMatch = rawText.match(/https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (bulkDataArray || urlMatch) {
    _sendAgentMsg(source, chatId, "⏳ Membaca struktur data dokumen/tabel...");
    try {
      const rawData = bulkDataArray || _getRawDataArrayFromUrl(urlMatch[0]);
      if (rawData && rawData.length > 1) {
         const previewData = rawData.slice(0, 10);

         const mappingPrompt = `You are an expert data mapper. Analyze the following first 10 rows of a Google Sheet and determine the sheet type (INVENTORY, TRANSACTION, or UNSTRUCTURED).
Source Data Preview: ${JSON.stringify(previewData)}

Rules:
1. Return ONLY a valid JSON object.
2. The "sheet_type" must be "INVENTORY", "TRANSACTION", or "UNSTRUCTURED".
3. If the data is a complex report, pivot table, matrix, or lacks a clear single header row, set "sheet_type": "UNSTRUCTURED".
4. If it's INVENTORY, map to these keys: item_code, item_name, category, branch, initial_stock, unit, buy_price, sell_price, status
5. If it's TRANSACTION, map to these keys: date, type, item_code, item_name, branch, qty, buy_price, sell_price, notes, operator
6. For standard tables, the values in your mapping must be the EXACT matching source header name from the first row.

Example output for INVENTORY:
{
  "sheet_type": "INVENTORY",
  "mapping": {
    "item_code": "SKU",
    "item_name": "Nama Produk"
  }
}

Example output for UNSTRUCTURED:
{
  "sheet_type": "UNSTRUCTURED",
  "mapping": {}
}
`;
         _sendAgentMsg(source, chatId, "🧠 Menghubungi AI untuk mengklasifikasi dan memetakan kolom otomatis...");
         const aiRaw = callAI(mappingPrompt, "You are a JSON generator.");
         const jsonMatch = aiRaw.match(/\{[\s\S]*\}/);
         if (!jsonMatch) throw new Error("AI gagal memetakan kolom.");

         const parsed = JSON.parse(jsonMatch[0]);
         const sheetType = parsed.sheet_type || "INVENTORY";
         const mapping = parsed.mapping || {};

         if (sheetType === "UNSTRUCTURED") {
             _sendAgentMsg(source, chatId, "🧩 Tabel terdeteksi sebagai Laporan/Matriks kompleks. Memulai Mode Pemahaman Mendalam AI...");
             if (!bulkDataArray && urlMatch) {
                 const extractedText = _extractTextFromUrl(urlMatch[0]);
                 if (extractedText) rawText += `\n\n[GOOGLE SHEETS CONTENT]\n` + extractedText;
             }
             // Do NOT return here. Fall through to the main AI processing block!
         } else {
             if (sheetType === "INVENTORY" && !mapping.item_name) {
                 throw new Error("AI mengklasifikasi ini sebagai Inventory, tapi tidak bisa menemukan kolom Nama Barang.");
             }

             const sourceHeaders = rawData[0]; // Assume first row is header for tabular data
             // Process mapping indices
             const headerIndices = {};
             for (const targetField in mapping) {
                 const sourceField = mapping[targetField];
                 headerIndices[targetField] = sourceField ? sourceHeaders.indexOf(sourceField) : -1;
             }

             const numRowsToCopy = rawData.length - 1;
             _sendAgentMsg(source, chatId, `⚡ Mengimpor ${numRowsToCopy} baris data ke tab ${sheetType === "INVENTORY" ? "Inventory" : "Transaction"} secara instan...`);

             const ss = _getSpreadsheet();
             let targetSheetName = sheetType === "INVENTORY" ? "Inventory" : "Transaction";
             let targetSheet = ss.getSheetByName(targetSheetName);
             if (!targetSheet) throw new Error(`Tab ${targetSheetName} tidak ditemukan.`);

             const newRows = [];
             for (let r = 1; r < rawData.length; r++) {
                 const row = rawData[r];
                 if (row.join("").trim() === "") continue; // Skip empty rows

                 if (sheetType === "INVENTORY") {
                     const itemCode = headerIndices.item_code > -1 ? row[headerIndices.item_code] : `MIG-${Date.now()}-${r}`;
                     const itemName = headerIndices.item_name > -1 ? row[headerIndices.item_name] : `Item ${r}`;
                     if (!itemName || itemName.trim() === "") continue;

                     const category = headerIndices.category > -1 ? row[headerIndices.category] : "Migrated";
                     const branch = headerIndices.branch > -1 ? row[headerIndices.branch] : "";
                     const initialStock = headerIndices.initial_stock > -1 ? parseFloat(row[headerIndices.initial_stock] || 0) : 0;
                     const unit = headerIndices.unit > -1 ? row[headerIndices.unit] : "Pcs";
                     const buyPrice = headerIndices.buy_price > -1 && row[headerIndices.buy_price] ? parseFloat(String(row[headerIndices.buy_price]).replace(/[^0-9.-]+/g,"") || 0) : 0;
                     const sellPrice = headerIndices.sell_price > -1 && row[headerIndices.sell_price] ? parseFloat(String(row[headerIndices.sell_price]).replace(/[^0-9.-]+/g,"") || 0) : 0;
                     const status = headerIndices.status > -1 ? row[headerIndices.status] : "Active";

                     const uniqueId = Utilities.getUuid ? Utilities.getUuid().substring(0,8).toUpperCase() : `ID-${Date.now()}-${r}`;

                     newRows.push([
                         uniqueId, itemCode, itemName, category, branch, initialStock, 0, 0, "", 0, unit, buyPrice, sellPrice, status, "", new Date()
                     ]);
                 } else {
                     // TRANSACTION schema: ["#", "Date & Time", "Type", "Item Code", "Item Name", "Branch", "Qty", "Stock Before", "Stock After", "Buy Price", "Sell Price", "Profit", "Notes", "Operator", "Source"]
                     const date = headerIndices.date > -1 && row[headerIndices.date] ? row[headerIndices.date] : new Date();
                     const type = headerIndices.type > -1 ? row[headerIndices.type] : "MIGRATION";
                     const itemCode = headerIndices.item_code > -1 ? row[headerIndices.item_code] : "";
                     const itemName = headerIndices.item_name > -1 ? row[headerIndices.item_name] : "";
                     const branch = headerIndices.branch > -1 ? row[headerIndices.branch] : "";
                     const qty = headerIndices.qty > -1 && row[headerIndices.qty] ? parseFloat(String(row[headerIndices.qty]).replace(/[^0-9.-]+/g,"") || 0) : 1;
                     const buyPrice = headerIndices.buy_price > -1 && row[headerIndices.buy_price] ? parseFloat(String(row[headerIndices.buy_price]).replace(/[^0-9.-]+/g,"") || 0) : 0;
                     const sellPrice = headerIndices.sell_price > -1 && row[headerIndices.sell_price] ? parseFloat(String(row[headerIndices.sell_price]).replace(/[^0-9.-]+/g,"") || 0) : 0;
                     const profit = sellPrice - buyPrice;
                     const notes = headerIndices.notes > -1 ? row[headerIndices.notes] : "";
                     const operator = headerIndices.operator > -1 ? row[headerIndices.operator] : "Auto-Import";

                     newRows.push([
                         "", date, type, itemCode, itemName, branch, qty, 0, 0, buyPrice, sellPrice, profit, notes, operator, "System"
                     ]);
                 }
             }

             if (newRows.length > 0) {
                const lastRow = Math.max(targetSheet.getLastRow(), 1);
                targetSheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);

                if (sheetType === "INVENTORY") {
                    const formulas = [];
                    for(let nr = lastRow + 1; nr <= lastRow + newRows.length; nr++) {
                        formulas.push([`=F${nr}+G${nr}-H${nr}`]);
                    }
                    targetSheet.getRange(lastRow + 1, 9, newRows.length, 1).setFormulas(formulas);
                } else if (sheetType === "TRANSACTION") {
                    const formulas = [];
                    for(let nr = lastRow + 1; nr <= lastRow + newRows.length; nr++) {
                        formulas.push([`=IF(C${nr}="","",ROW()-1)`]);
                    }
                    targetSheet.getRange(lastRow + 1, 1, newRows.length, 1).setFormulas(formulas);
                }

                _sendAgentMsg(source, chatId, `✅ Sukses! ${newRows.length} baris berhasil dimigrasikan ke tab ${targetSheetName}.`);
             } else {
                _sendAgentMsg(source, chatId, "⚠️ Tidak ada baris data yang valid untuk diimpor.");
             }

             return; // Stop further parsing so it doesn't try to parse transactions as chat
         }

      } else {
         _sendAgentMsg(source, chatId, "⚠️ Google Sheets tersebut kosong.");
         return;
      }
    } catch(e) {
      _sendAgentMsg(source, chatId, "❌ " + e.message);
      return;
    }
  }

  const itemContext = _getInventoryContext();
  
  // Anti Prompt-Injection: Escape HTML/XML chars to prevent breaking out of tags
  const sanitizedText = rawText.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  const prompt = `**Task:** You are a strict inventory data parser. Extract transactions from the employee report below.

**CRITICAL CONSTRAINTS - VIOLATING THESE WILL CAUSE A SYSTEM CRASH:**
- DO NOT print "User request:", "Interpretation:", "Action:", "Sheet:", or "Range:".
- DO NOT use bullet points or asterisks (*).
- DO NOT think step-by-step.
- YOU MUST START YOUR ENTIRE RESPONSE WITH THE CHARACTER '{' and END WITH '}'.

<employee_report>
${sanitizedText}
</employee_report>

AVAILABLE ITEMS (SKU | Name | Category | CurrentStock):
${itemContext}

**Rules:**
1. UNREGISTERED ITEMS: If the user mentions an item that is NOT in the AVAILABLE ITEMS list, set "type": "UNKNOWN" and "item_code": null, and write a strict warning in 'notes'. Do NOT hallucinate it as a NEW item!
2. FUZZY / PARTIAL MATCHING: If the user omits minor details but the main attributes match, use that item.
3. IDENTICAL DUPLICATES: If multiple items perfectly match, pick the FIRST matching item_code.
4. Transaction types: "IN", "OUT", "ADJUSTMENT", "CHECK", "UNKNOWN".
5. quantity must always be a positive number.
6. Each separate item in the report = separate transaction object.
7. If AMBIGUOUS, set item_code="AMBIGUOUS" and write a short question in 'notes'.
8. ANTI-FRAUD VALIDATION: If reported price differs from system price without a valid reason, set "type": "UNKNOWN". If valid reason, prefix notes with "[PRICE OVERRIDE: reason]".

CRITICAL LOCALIZATION RULE: Detect the language of the user's input. You MUST provide the translations for the system messages below in the exact same language.

**Output Format Example:**
{
  "transactions": [
    {
      "type": "OUT",
      "item_code": "SKU-123",
      "item_name": "Item Name",
      "quantity": 1,
      "branch": null,
      "notes": "sold",
      "confidence": 90,
      "ai_reasoning": "matched"
    }
  ],
  "sys_err_intent": "Translated: 🤔 I couldn't determine the intent...",
  "sys_err_qty": "Translated: ❌ Could not detect a valid quantity for",
  "sys_confirm": "Translated: 🤔 AI is {conf}% confident about {name}..."
}`;

  let parsedList;
  let rootParsed = {};

  debugLog("Sending to AI...");

  try {
    // Delegate entirely to AIAgent's unified Smart Loop QA
    let parsed = AIAgent.parseTransaction(rawText);

    if (parsed) {
        rootParsed = parsed || {};
        if (!Array.isArray(parsed) && typeof parsed === 'object') {
          // If the AI returns a single object that wraps an array (e.g., {"transactions": [...]})
          if (parsed.transactions && Array.isArray(parsed.transactions)) {
            parsed = parsed.transactions;
          } else {
            const keys = Object.keys(parsed);
            for (const key of keys) {
              if (Array.isArray(parsed[key])) {
                parsed = parsed[key];
                break;
              }
            }
          }
        }
        parsedList = Array.isArray(parsed) ? parsed : [parsed];
    } else {
        throw new Error("Failed to parse transactions after multiple attempts.");
    }
  } catch (err) {
    Logger.log("AI error: " + err);
    debugLog("AI Error: " + err.message);
    if (chatId) {
      _sendAgentMsg(source, chatId, `❌ AI Provider Error: ${err.message}`);
    }
    return;
  }

  debugLog("Parsed " + parsedList.length + " transactions from AI");

  if (parsedList.length === 0) {
    const errMsg = rootParsed.sys_err_intent || "🤔 I couldn't determine the intent of this part of your report.\n\nPlease clarify: is this a Stock IN, Stock OUT, or Stock Adjustment?";
    if (chatId) _sendAgentMsg(source, chatId, errMsg);
    return;
  }

  for (let idx = 0; idx < parsedList.length; idx++) {
    const parsed = parsedList[idx];
    debugLog("Processing tx #" + (idx+1) + ": type=" + parsed.type + " code=" + parsed.item_code + " name=" + parsed.item_name + " qty=" + parsed.quantity + " conf=" + parsed.confidence);
    if (parsed.type === "UNKNOWN") {
      const errMsg = rootParsed.sys_err_intent || "🤔 I couldn't determine the intent of this part of your report.\n\nPlease clarify: is this a Stock IN, Stock OUT, or Stock Adjustment?";

      // Fraud Detection Hook
      if (parsed.notes && parsed.notes.toUpperCase().includes("FRAUD WARNING")) {
          _logSecurityThreat(senderName, "Price Manipulation", rawText, parsed.notes);
          if (chatId) _sendAgentMsg(source, chatId, `🚨 ${parsed.notes}`);
          continue;
      }

      if (chatId) _sendAgentMsg(source, chatId,
        `${errMsg}\n\nAI Note: ${parsed.ai_reasoning || "Not specified"}`
      );
      continue;
    }

    if (parsed.type === "CHECK") {
      if (chatId) _sendAgentMsg(source, chatId, _buildStockCheckResult(parsed.item_name || parsed.item_code));
      continue;
    }

    if (parsed.quantity <= 0 && parsed.type !== "ADJUSTMENT" && !(parsed.type === "IN" && (parsed.item_new === true || parsed.item_new === "true"))) {
      const errQty = rootParsed.sys_err_qty || "❌ Could not detect a valid quantity for";
      if (chatId) _sendAgentMsg(source, chatId, `${errQty} ${parsed.item_name || parsed.new_item_name || "item"}.`);
      continue;
    }

    if (parsed.confidence < 75) {
      if (chatId) {
        const confTemplate = rootParsed.sys_confirm || "🤔 AI is {conf}% confident about {name}. Please confirm:\nType: {type}\nQty: {qty}\n\nReply YES to confirm.";
        const msg = confTemplate
          .replace("{conf}", parsed.confidence)
          .replace("{name}", parsed.item_name || parsed.new_item_name)
          .replace("{type}", parsed.type)
          .replace("{qty}", parsed.quantity);

        CacheService.getUserCache().put("pending_" + chatId, JSON.stringify({ parsed, rawText, senderName, source }), 300);
        _sendAgentMsg(source, chatId, msg);
      }
      continue;
    }

    debugLog("Executing transaction for: " + parsed.item_name);
    try {
      _executeAgentTransaction(chatId, parsed, senderName, source, rawText);
      debugLog("Transaction SUCCESS for: " + parsed.item_name);
    } catch(txErr) {
      debugLog("Transaction FAILED for " + parsed.item_name + ": " + txErr + " | Stack: " + txErr.stack);
      if (chatId) _sendAgentMsg(source, chatId, "❌ Error processing " + parsed.item_name + ": " + txErr.message);
    }
  }
}

// ─── EXECUTE TRANSACTION FROM AI ─────────────────────────────
function _executeAgentTransaction(chatId, parsed, senderName, source, rawText) {
  const lock = LockService.getScriptLock();
  try {
    // Wait up to 30 seconds for other processes to finish writing
    lock.waitLock(30000);
    
    debugLog("_executeAgentTransaction START: code=" + parsed.item_code + " name=" + parsed.item_name);
    if (parsed.item_code === "AMBIGUOUS") { if (chatId) _sendAgentMsg(source, chatId, "⚠️ **Ambiguitas Ditemukan:** " + parsed.notes); return; }
    let item = _findItem(parsed.item_code, parsed.item_name);
    debugLog("_findItem result: " + (item ? "FOUND row=" + item.row + " name=" + item.name + " stock=" + item.stock : "NULL"));

  if (!item) {
    if (chatId) _sendAgentMsg(source, chatId, `❌ **Transaction Denied:** Item "${parsed.item_name || 'Tidak diketahui'}" is not registered in the database.\\n\\nPlease ask the Admin to add it via the Google Sheets menu (Smart Inventory ➔ Registrasi Barang Baru).`);
    return;
  }

  const stockBefore = item.stock;
  const qty = parseInt(parsed.quantity) || 0;

  // Validate stock for OUT transactions
  if (parsed.type === "OUT" && qty > stockBefore) {
    if (chatId) _sendAgentMsg(source, chatId,
      `❌ <b>Insufficient Stock</b>\n\nItem: ${item.name}\nRequested: ${qty}\nAvailable: ${stockBefore}\n\nTransaction cancelled.`
    );
    return;
  }

  const txType = parsed.type === "IN" ? "STOCK IN ➕" : parsed.type === "OUT" ? "STOCK OUT ➖" : "ADJUSTMENT 🔧";
  const notes  = (parsed.notes || "") + ` | Report by: ${senderName} via ${source}`;
  const overrideNewStock = parsed.type === "ADJUSTMENT" ? qty : undefined;

  const stockAfter = _recordTransaction(item, txType, qty, notes, source, overrideNewStock);

  // Calculate new status
  const minStk = item.minStock;
  let newStatus;
  if (stockAfter === 0) newStatus = "🔴 OUT OF STOCK";
  else if (stockAfter <= minStk * 0.5) newStatus = "🟠 CRITICAL";
  else if (stockAfter <= minStk) newStatus = "🟡 LOW STOCK";
  else newStatus = "🟢 IN STOCK";

  // Build proof message
  const proof = _buildTransactionProof({
    txId:         "TRX-" + Date.now(),
    timestamp:    _formatTimestamp(new Date()),
    type:         txType,
    itemCode:     item.code,
    itemName:     item.name,
    branch:       parsed.branch || item.branch,
    qty:          qty,
    stockBefore:  stockBefore,
    stockAfter:   stockAfter,
    newStatus:    newStatus,
    operator:     senderName,
    source:       source,
    notes:        parsed.notes || "",
    rawReport:    rawText,
    aiConfidence: parsed.confidence,
    aiReasoning:  parsed.ai_reasoning,
  });

  // Send proof to employee
  if (chatId) {
    if (source === "Email") {
      _sendEmailNotification(chatId, `✅ Transaction Confirmed: ${txType}`, proof.html);
    } else {
      _sendAgentMsg(source, chatId, proof.message);
    }
  }

  // Notify admin
  const adminTgId = _getScriptProps().getProperty("TG_ADMIN_ID");
  if (adminTgId && adminTgId !== chatId) {
    _sendTelegram(adminTgId, `📢 <b>New Report from ${senderName} via ${source}</b>\n\n` + proof.message);
  }

  // Email admin
  const adminEmail = _getScriptProps().getProperty("ADMIN_EMAIL");
  if (adminEmail) {
    _sendEmailNotification(adminEmail,
      `🚨 New Transaction via ${source}`,
      proof.html
    );
  }

  // Alert if critical/out
  if (newStatus.includes("CRITICAL") || newStatus.includes("OUT OF STOCK")) {
    const alert = `\n\n⚠️ <b>STOCK ALERT</b>\n${item.name} is now <b>${newStatus}</b>\nRemaining: ${stockAfter} ${item.unit}`;
    if (chatId) _sendAgentMsg(source, chatId, alert);
    if (adminTgId) _sendTelegram(adminTgId, alert);
  }

  } catch (e) {
    Logger.log("Lock or Execution Error in _executeAgentTransaction: " + e.message);
    if (e.message.toLowerCase().includes("lock")) {
      if (chatId) _sendAgentMsg(source, chatId, "❌ System is busy (Concurrency limit). Please repeat your report in a few seconds.");
    } else {
      throw e;
    }
  } finally {
    lock.releaseLock();
  }
}

// ─── PROOF MESSAGE BUILDER ────────────────────────────────────
function _buildTransactionProof(d) {
  const line = "━━━━━━━━━━━━━━━━━━━━━━━━━";
  const delta = d.type.includes("IN") ? `+${d.qty}` : d.type.includes("OUT") ? `-${d.qty}` : `→${d.stockAfter}`;

  const message =
    `✅ <b>TRANSACTION CONFIRMED</b>\n${line}\n` +
    `🆔 ID: <code>${d.txId}</code>\n` +
    `📅 Time: ${d.timestamp}\n` +
    `🏷️ Type: <b>${d.type}</b>\n${line}\n` +
    `📦 <b>Item Details:</b>\n` +
    `• Code: <code>${d.itemCode}</code>\n` +
    `• Name: ${d.itemName}\n` +
    `• Branch: ${d.branch || "—"}\n` +
    `• Quantity: <b>${d.qty.toLocaleString()}</b>\n${line}\n` +
    `📊 <b>Stock Calculation Proof:</b>\n` +
    `  Before : ${d.stockBefore.toLocaleString()}\n` +
    `  Change : ${delta}\n` +
    `  After  : <b>${d.stockAfter.toLocaleString()}</b>\n` +
    `  Status : ${d.newStatus}\n${line}\n` +
    `👤 Operator: ${d.operator}\n` +
    `📡 Channel: ${d.source}\n` +
    `🤖 AI Confidence: ${d.aiConfidence}%\n` +
    `📝 Notes: ${d.notes || "—"}\n${line}\n` +
    `<i>Original report: "${d.rawReport.substring(0, 80)}"</i>`;

  const html = `
    <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;background:#F8FAFC;padding:20px;">
      <div style="background:#064E3B;color:white;padding:20px;border-radius:12px;text-align:center;margin-bottom:20px;">
        <h2 style="margin:0;font-size:20px;">✅ Transaction Confirmed</h2>
        <p style="margin:4px 0 0;opacity:0.8;font-size:13px;">AI Smart Inventory System v3.0</p>
      </div>
      <div style="background:white;border-radius:10px;padding:18px;margin-bottom:14px;border-left:4px solid #10B981;">
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          <tr><td style="padding:5px 0;color:#64748B;width:38%;">Transaction ID</td><td style="font-family:monospace;font-weight:bold;">${d.txId}</td></tr>
          <tr><td style="padding:5px 0;color:#64748B;">Timestamp</td><td>${d.timestamp}</td></tr>
          <tr><td style="padding:5px 0;color:#64748B;">Type</td><td><b>${d.type}</b></td></tr>
          <tr><td style="padding:5px 0;color:#64748B;">Channel</td><td>${d.source}</td></tr>
        </table>
      </div>
      <div style="background:white;border-radius:10px;padding:18px;margin-bottom:14px;">
        <h3 style="margin:0 0 12px;font-size:14px;color:#0F172A;">📦 Item Details</h3>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          <tr><td style="padding:5px 0;color:#64748B;width:38%;">Item Code</td><td style="font-family:monospace;">${d.itemCode}</td></tr>
          <tr><td style="padding:5px 0;color:#64748B;">Item Name</td><td><b>${d.itemName}</b></td></tr>
          <tr><td style="padding:5px 0;color:#64748B;">Branch</td><td>${d.branch || "—"}</td></tr>
          <tr><td style="padding:5px 0;color:#64748B;">Quantity</td><td><b style="font-size:16px;">${d.qty.toLocaleString()}</b></td></tr>
        </table>
      </div>
      <div style="background:#F0FDF4;border:1px solid #A7F3D0;border-radius:10px;padding:18px;margin-bottom:14px;">
        <h3 style="margin:0 0 14px;font-size:14px;color:#064E3B;">📊 Stock Calculation Proof</h3>
        <table style="width:100%;font-size:14px;border-collapse:collapse;">
          <tr style="border-bottom:1px solid #D1FAE5;"><td style="padding:8px 0;color:#374151;">Stock Before</td><td style="text-align:right;font-family:monospace;">${d.stockBefore.toLocaleString()}</td></tr>
          <tr style="border-bottom:1px solid #D1FAE5;">
            <td style="padding:8px 0;color:${d.type.includes("IN") ? "#065F46" : "#991B1B"};">
              ${d.type.includes("IN") ? "➕ Added" : d.type.includes("OUT") ? "➖ Removed" : "🔧 Adjusted to"}
            </td>
            <td style="text-align:right;font-family:monospace;color:${d.type.includes("IN") ? "#059669" : "#DC2626"};">
              ${d.type.includes("IN") ? "+" : d.type.includes("OUT") ? "-" : "→"}${d.qty.toLocaleString()}
            </td>
          </tr>
          <tr><td style="padding:8px 0;font-weight:bold;font-size:15px;">Stock After</td><td style="text-align:right;font-family:monospace;font-weight:bold;font-size:20px;">${d.stockAfter.toLocaleString()}</td></tr>
        </table>
        <div style="margin-top:10px;text-align:center;padding:8px;background:white;border-radius:6px;font-size:15px;">Status: <b>${d.newStatus}</b></div>
      </div>
      <div style="background:#F1F5F9;border-radius:10px;padding:14px;font-size:12px;color:#64748B;">
        <p style="margin:0 0 4px;"><b>Operator:</b> ${d.operator}</p>
        <p style="margin:0 0 4px;"><b>AI Confidence:</b> ${d.aiConfidence}%</p>
        <p style="margin:0 0 4px;"><b>AI Reasoning:</b> ${d.aiReasoning}</p>
        <p style="margin:0;font-style:italic;"><b>Original report:</b> "${d.rawReport}"</p>
      </div>
    </div>`;

  return { message, html };
}

// ─── DASHBOARD SUMMARY ────────────────────────────────────────
function _buildDashboardSummary() {
  const sh = _getSheet(SHEETS.INVENTORY);
  if (!sh) return "❌ Inventory sheet not found.";
  const data = sh.getDataRange().getValues();
  let total = 0, out = 0, crit = 0, low = 0, safe = 0;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    total++;
    const s = (data[i][12] || "").toString();
    if (s.includes("OUT")) out++;
    else if (s.includes("CRITICAL")) crit++;
    else if (s.includes("LOW")) low++;
    else safe++;
    Utilities.sleep(100);
  }
  return (
    `📊 <b>Inventory Dashboard</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━\n\n` +
    `📦 Total Items: <b>${total}</b>\n` +
    `🟢 In Stock: <b>${safe}</b>\n` +
    `🟡 Low Stock: <b>${low}</b>\n` +
    `🟠 Critical: <b>${crit}</b>\n` +
    `🔴 Out of Stock: <b>${out}</b>\n\n` +
    `<i>Updated: ${_formatTimestamp(new Date())}</i>`
  );
}

// ─── STOCK CHECK ──────────────────────────────────────────────
function _buildStockCheckResult(query) {
  if (!query) return "Please specify an item name or code to check.";
  const sh = _getSheet(SHEETS.INVENTORY);
  if (!sh) return "❌ Inventory sheet not found.";
  const data = sh.getDataRange().getValues();
  const q = query.toLowerCase();
  const results = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue;
    if ((data[i][1] + data[i][2] + data[i][3]).toLowerCase().includes(q)) {
      results.push({ code: data[i][1], name: data[i][2], branch: data[i][4], stock: data[i][8] || 0, unit: data[i][10] || "", status: data[i][12] || "🟢 IN STOCK" });
    }
  }
  if (!results.length) return `❌ No items found matching "<b>${query}</b>".`;
  let msg = `📦 <b>Stock Check: "${query}"</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━\n`;
  results.slice(0, 6).forEach((r, i) => {
    msg += `\n<b>${i + 1}. ${r.name}</b>\n`;
    msg += `   Code: <code>${r.code}</code> | Branch: ${r.branch}\n`;
    msg += `   Stock: <b>${r.stock.toLocaleString()} ${r.unit}</b> — ${r.status}\n`;
  });
  if (results.length > 6) msg += `\n<i>…and ${results.length - 6} more items</i>`;
  return msg;
}


function _logSecurityThreat(operatorName, threatType, originalMessage, aiNotes) {
  try {
    const ss = _getSpreadsheet();
    // Use string "Security Audit" matching Setup.js SHEETS.SECURITY
    let sh = ss.getSheetByName("Security Audit");
    if (sh) {
       sh.appendRow([
         new Date(),
         operatorName,
         threatType,
         originalMessage,
         aiNotes
       ]);
       debugLog(`SECURITY THREAT LOGGED: ${operatorName} - ${threatType}`);
    }
  } catch (e) {
    debugLog("Failed to log security threat: " + e.message);
  }
}

// ─── WELCOME MESSAGE ─────────────────────────────────────────
function _buildWelcomeMessage(name) {
  return (
    `👋 Hello <b>${name}</b>! I'm the AI Inventory Agent.\n\n` +
    `You can report stock in any natural language:\n\n` +
    `📥 <b>Stock In (Receiving):</b>\n` +
    `• "received 10 iphone 15 from supplier"\n` +
    `• "20 boxes of water arrived at branch"\n` +
    `• "ELE-001 stock in: 50 units"\n\n` +
    `📤 <b>Stock Out (Dispatch):</b>\n` +
    `• "sent 5 samsung s24 to Jakarta"\n` +
    `• "customer bought 3 MacBook Pro"\n` +
    `• "OUT: 100 boxes Aqua to Bandung branch"\n\n` +
    `🔧 <b>Stock Adjustment (Physical Count):</b>\n` +
    `• "physical count: paper A4 = 8 reams"\n` +
    `• "actual stock check: iPhone 15 has 42 units"\n\n` +
    `📊 <b>Quick Commands:</b>\n` +
    `• /check [item] — check stock level\n` +
    `• /stock [item] — same as /check\n` +
    `• /dashboard — view summary\n` +
    `• /help — show this message\n\n` +
    `🛡️ <b>Admin Commands:</b>\n` +
    `• /users — show whitelist\n` +
    `• /allow [ID] — grant access\n` +
    `• /block [ID] — revoke access\n` +
    `• /public_mode [on/off] — toggle security\n\n` +
    `Just type naturally — AI will understand! 🚀`
  );
}

// ─── MESSAGING HELPERS ────────────────────────────────────────
function _sendAgentMsg(source, recipient, text) {
  if (source === "Simulation") {
    SpreadsheetApp.getUi().alert("🧪 Simulation AI Response", text, SpreadsheetApp.getUi().ButtonSet.OK);
  } else if (source === "Telegram") {
    _sendTelegram(recipient, text);
  } else if (source === "WhatsApp") {
    _sendWhatsApp(recipient, text);
  } else if (source === "Email") {
    if (text.startsWith("⏳") || text.startsWith("🧠") || text.startsWith("🧩") || text.startsWith("⚡")) return;
    const htmlBody = text.replace(/\n/g, "<br>");
    _sendEmailNotification(recipient, "AI Inventory Notification", htmlBody);
  }
}

function _sendTelegram(chatId, text) {
  const token = _getScriptProps().getProperty("TG_TOKEN");
  if (!token || !chatId) return;
  try {
    const payload = {
      chat_id: chatId,
      text: text,
      parse_mode: "HTML"
    };
    UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
  } catch (e) { Logger.log("Telegram send error: " + e); }
}

function _sendWhatsApp(phone, text) {
  const props = _getScriptProps();
  const phoneId = props.getProperty("WA_PHONE_ID");
  const token   = props.getProperty("WA_TOKEN");
  
  if (!phoneId || !token) {
    Logger.log("WhatsApp credentials missing in Agent Settings.");
    return;
  }
  
  try {
    UrlFetchApp.fetch(`https://graph.facebook.com/v20.0/${phoneId}/messages`, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({
        messaging_product: "whatsapp",
        to: phone,
        type: "text",
        text: { body: text.replace(/\*/g, "").replace(/_/g, "").replace(/`/g, "") },
      }),
      muteHttpExceptions: true,
    });
  } catch (e) { Logger.log("WhatsApp send error: " + e); }
}

  // --- NOTIFICATION HANDLERS --------------------------------------------------
  function _sendEmailNotification(recipient, subject, body) {
    try {
      // The body might be plain text from the AI but htmlBody requires <br> for newlines.
      const isHtml = body.includes('<br>') || body.includes('<div>') || body.includes('<p>');
      const htmlBody = isHtml ? body : body.replace(/\n/g, '<br>');
      GmailApp.sendEmail(recipient, subject, body, { htmlBody: htmlBody });
    } catch (e) {
      debugLog("EMAIL FAILED TO SEND: " + e.message);
      const adminId = _getScriptProps().getProperty("ADMIN_CHAT_ID");
      if (adminId) {
        _sendTelegram(adminId, `⚠️ <b>SYSTEM ALERT</b> ⚠️\n\nGagal mengirim email balasan ke ${recipient} karena limit kuota Gmail.\n\n<b>Status Sistem:</b> Perintah pengguna telah diproses oleh AI, namun balasan tertahan.`);
      }
    }
  }

// ─── GEMINI AI ────────────────────────────────────────────────
function _callGemini(prompt, apiKey) {
  const url  = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 8192 },
    }),
    muteHttpExceptions: true,
  });
  const json = JSON.parse(resp.getContentText());
  if (json.error) throw new Error(json.error.message);
  return json.candidates[0].content.parts[0].text;
}

// ─── UTILITIES ────────────────────────────────────────────────
function _formatTimestamp(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm:ss");
}

// ─── WHATSAPP WEBHOOK VERIFICATION ───────────────────────────
function doGet(e) {
  if (e.parameter["action"] === "test_debuglog") {
    debugLog("HELLO FROM TEST_DEBUGLOG");
    return ContentService.createTextOutput("Test DebugLog executed");
  }
  if (e.parameter["action"] === "get_logs") {
    const ss = _getSpreadsheet();
    if (!ss) return ContentService.createTextOutput("No Spreadsheet");
    const sheet = ss.getSheetByName("DebugLogs");
    if (!sheet) return ContentService.createTextOutput("No Logs sheet");
    const data = sheet.getDataRange().getValues();
    const lastLogs = data.slice(-15);
    return ContentService.createTextOutput(JSON.stringify(lastLogs));
  }
  if (e.parameter["action"] === "diagnose") {
    let result = "DIAGNOSTICS:\n";
    try {
      const ss = _getSpreadsheet();
      result += "1. _getSpreadsheet: SUCCESS. ID=" + ss.getId() + "\n";
      const sheet = ss.getSheetByName("DebugLogs");
      if (sheet) result += "2. DebugLogs Sheet: SUCCESS.\n";
      const id = _getScriptProps().getProperty("SHEET_ID");
      result += "3. SHEET_ID in Properties: " + id + "\n";
      const data = sheet.getDataRange().getValues();
      const lastLogs = data.slice(0, 15).map(r => r.join(" | ")).join("\n");
      return ContentService.createTextOutput(result + "\nLOGS:\n" + lastLogs);
    } catch(err) {
      return ContentService.createTextOutput(result + "ERROR: " + err.message + "\n" + err.stack);
    }
  }
  if (e.parameter["action"] === "force_poll") {
    pollEmails();
    return ContentService.createTextOutput("Forced poll executed successfully. Check DebugLogs.");
  }
  if (e.parameter["action"] === "set_sheet_id") {
    const id = e.parameter["id"];
    if (id) {
      _getScriptProps().setProperty("SHEET_ID", id);
      return ContentService.createTextOutput("SHEET_ID successfully set to: " + id);
    }
    return ContentService.createTextOutput("Missing id parameter");
  }

  if (e.parameter["action"] === "register_webhook") {
    try {
      const props = _getScriptProps();
      const token = props.getProperty("TG_TOKEN");
      if (!token) return ContentService.createTextOutput("No TG_TOKEN");
      const secret = props.getProperty("WEBHOOK_SECRET");
      const webhookUrl = ScriptApp.getService().getUrl() + "?token=" + secret;
      const resp = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/setWebhook?url=${encodeURIComponent(webhookUrl)}&drop_pending_updates=true`, { muteHttpExceptions: true });
      return ContentService.createTextOutput("Webhook Registration Response: " + resp.getContentText());
    } catch (err) {
      return ContentService.createTextOutput("Webhook Registration Error: " + err);
    }
  }

  // WhatsApp Cloud API webhook verification challenge
  const mode  = e.parameter["hub.mode"];
  const token = e.parameter["hub.verify_token"];
  const challenge = e.parameter["hub.challenge"];
  const myToken = "inventory2026";
  if (mode === "subscribe" && token === myToken) {
    return ContentService.createTextOutput(challenge);
  }
  return ContentService.createTextOutput("Forbidden");
}

// ─── WEEKLY AI SUMMARY (CRON) ───────────────────────────────
function generateWeeklySummary() {
  try {
    const adminId = _getScriptProps().getProperty("ADMIN_CHAT_ID");
    if (!adminId) return;

    const txSh = _getSheet(SHEETS.TRANSACTIONS);
    if (!txSh) return;
  
  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  
  const data = txSh.getDataRange().getValues();
  if (data.length < 2) return;
  
  const cmap = _getTransactionColMap(data[0]);
  let recentTx = [];
  
  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][cmap.date]);
    if (rowDate >= sevenDaysAgo) {
      let tType = (cmap.type !== -1) ? data[i][cmap.type] : "TX";
      let tName = (cmap.name !== -1) ? data[i][cmap.name] : "Item";
      let tQty = (cmap.qty !== -1) ? data[i][cmap.qty] : 0;
      let tProf = (cmap.profit !== -1) ? data[i][cmap.profit] : 0;
      recentTx.push(`- ${tType}: ${tName}, Qty: ${tQty}, Profit: $${tProf}`);
    }
  }
  
  if (recentTx.length === 0) {
    _sendTelegramMessage(adminId, "📊 *Weekly AI Summary*\n\nThere were no transactions in the last 7 days.");
    return;
  }
  
    const prompt = `You are a smart business assistant. Here are the transactions from the last 7 days:\n${recentTx.join("\n")}\n\nWrite a short, professional weekly summary (max 2 paragraphs) in English for the boss. Highlight total sales, top items, and total profit. Keep it enthusiastic!`;
  const systemPrompt = "You are an expert AI inventory analyst.";
  
  const summary = _callAIAuto(prompt, systemPrompt);
  _sendTelegramMessage(adminId, "📊 *Weekly AI Summary*\n\n" + summary);
  
  } catch (e) {
    Logger.log("Weekly Summary Error: " + e.message);
  }
}

// ─── WIPE EXECUTION ──────────────────────────────────────────
function _executeWipeAction(actionArr, source, chatId) {
  try {
    const ss = _getSpreadsheet();
    let executedCount = 0;
    
    for (let act of actionArr) {
      const sh = ss.getSheetByName(act.sheet);
      if (!sh) continue;
      
      if (act.cmd === "CLEAR_CONTENTS" && act.range) {
        let rStr = act.range;
        if (rStr.endsWith("Z") && !rStr.match(/\d+$/)) rStr += sh.getMaxRows();
        try { sh.getRange(rStr).clearContent(); executedCount++; } catch(e){}
      }
      else if (act.cmd === "CLEAR_FORMATS" && act.range) {
        let rStr = act.range;
        if (rStr.endsWith("Z") && !rStr.match(/\d+$/)) rStr += sh.getMaxRows();
        try { sh.getRange(rStr).clearFormat(); executedCount++; } catch(e){}
      }
      else if (act.cmd === "DELETE_COLUMNS" && act.start && act.count) {
        let count = act.count;
        if (act.start <= sh.getMaxColumns()) {
          if (act.start + count - 1 >= sh.getMaxColumns()) count = sh.getMaxColumns() - act.start + 1;
          // Prevent deleting all columns
          if (act.start === 1 && count === sh.getMaxColumns()) count = count - 1;
          if (count > 0) { sh.deleteColumns(act.start, count); executedCount++; }
        }
      }
      else if (act.cmd === "DELETE_ROWS" && act.start && act.count) {
        let count = act.count;
        if (act.start <= sh.getMaxRows()) {
          if (act.start + count - 1 >= sh.getMaxRows()) count = sh.getMaxRows() - act.start + 1;
          // Prevent deleting all rows
          if (act.start === 1 && count === sh.getMaxRows()) count = count - 1;
          if (count > 0) { sh.deleteRows(act.start, count); executedCount++; }
        }
      }
    }
    
    if (executedCount > 0) {
      _sendAgentMsg(source, chatId, `✅ **EXECUTION SUCCESSFUL**: ${executedCount} structural actions have been applied.`);
    } else {
      _sendAgentMsg(source, chatId, "⚠️ Tidak ada tindakan yang dieksekusi. Mungkin format sheet tidak valid.");
    }
  } catch (e) {
    _sendAgentMsg(source, chatId, "❌ **FAILED**: " + e.message);
  }
}

// ─── FORMAT EXECUTION ──────────────────────────────────────────
function _executeFormatAction(actionArr, source, chatId) {
  try {
    const ss = _getSpreadsheet();
    let executedCount = 0;

    let actionsToRun = [];
    for (let act of actionArr) {
      let sheetName = String(act.sheet).toUpperCase();
      if (sheetName === "ALL" || sheetName === "SEMUA") {
        const allSheets = ss.getSheets();
        for (let s of allSheets) {
          actionsToRun.push(Object.assign({}, act, { sheet: s.getName() }));
        }
      } else {
        actionsToRun.push(act);
      }
    }

    for (let act of actionsToRun) {
      const sh = ss.getSheetByName(act.sheet);
      if (!sh) continue;

      try {
        if (act.cmd === "RESET_FORMAT") {
          let fullRange = sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns());

          // Fase 1: Format Visual pada SELURUH sheet
          fullRange.breakApart();
          fullRange.clearFormat();
          fullRange.clearDataValidations();
          fullRange.setBackground(null);
          fullRange.setFontColor("#000000");
          sh.clearConditionalFormatRules();

          let bandings = sh.getBandings();
          for (let i = 0; i < bandings.length; i++) {
            bandings[i].remove();
          }

          // Fase 2: Menghapus Sel Error Tanpa Merusak Rumus Sehat (Menggunakan RangeList)
          // Gunakan getDataRange() untuk mencari error agar lebih cepat (tidak memproses 1000+ baris kosong)
          let dataRange = sh.getDataRange();
          if (dataRange) {
            let displayValues = dataRange.getDisplayValues();
            let errorCells = [];

            let getColLetter = (c) => {
               let temp, letter = '';
               while (c > 0) {
                 temp = (c - 1) % 26;
                 letter = String.fromCharCode(temp + 65) + letter;
                 c = (c - temp - 1) / 26;
               }
               return letter;
            };

            for (let r = 0; r < displayValues.length; r++) {
              for (let c = 0; c < displayValues[r].length; c++) {
                if (String(displayValues[r][c]).indexOf("#") === 0) {
                  errorCells.push(`${getColLetter(c + 1)}${r + 1}`);
                }
              }
            }

            if (errorCells.length > 0) {
              const chunkSize = 400;
              for (let i = 0; i < errorCells.length; i += chunkSize) {
                sh.getRangeList(errorCells.slice(i, i + chunkSize)).clearContent();
              }
            }
          }
          let maxCols = sh.getMaxColumns();
          if (maxCols > 0) sh.setColumnWidths(1, maxCols, 100);

          let maxRows = sh.getMaxRows();
          if (maxRows > 0) sh.setRowHeights(1, maxRows, 21);
          executedCount++;
        }
        else if (act.cmd === "REPAIR_FORMULA") {
          let maxRows = sh.getMaxRows();
          let maxCols = sh.getMaxColumns();
          if (maxRows >= 2 && maxCols > 0) {
             let headers = sh.getRange(1, 1, 1, maxCols).getValues()[0];
             let colCurrentStock = headers.findIndex(h => String(h).toLowerCase().includes("current stock")) + 1;
             let colMinStock = headers.findIndex(h => String(h).toLowerCase().includes("min stock")) + 1;
             let colStatus = headers.findIndex(h => String(h).toLowerCase() === "status") + 1;
             let colQR = headers.findIndex(h => String(h).toLowerCase().includes("qr")) + 1;
             let colUpdate = headers.findIndex(h => String(h).toLowerCase().includes("update")) + 1;

             let colInit = headers.findIndex(h => String(h).toLowerCase().includes("initial")) + 1 || 6;
             let colIn = headers.findIndex(h => String(h).toLowerCase().includes("in")) + 1 || 7;
             let colOut = headers.findIndex(h => String(h).toLowerCase().includes("out")) + 1 || 8;

             let getColLetter = (col) => {
               let temp, letter = '';
               while (col > 0) {
                 temp = (col - 1) % 26;
                 letter = String.fromCharCode(temp + 65) + letter;
                 col = (col - temp - 1) / 26;
               }
               return letter;
             };

             let initL = getColLetter(colInit);
             let inL = getColLetter(colIn);
             let outL = getColLetter(colOut);
             let currL = colCurrentStock > 0 ? getColLetter(colCurrentStock) : "I";
             let minL = colMinStock > 0 ? getColLetter(colMinStock) : "J";

             let numRows = maxRows - 1;

             let currentStockFormulas = [];
             let statusFormulas = [];
             let qrFormulas = [];
             let updateFormulas = [];

             for (let r = 2; r <= maxRows; r++) {
                if (colCurrentStock > 0) currentStockFormulas.push([`=IF(B${r}="","",${initL}${r}+${inL}${r}-${outL}${r})`]);
                if (colStatus > 0) statusFormulas.push([`=IF(B${r}="","",IF(${currL}${r}=0,"🔴 OUT OF STOCK",IF(${currL}${r}<=${minL}${r}/2,"🟠 CRITICAL",IF(${currL}${r}<=${minL}${r},"🟡 LOW STOCK","🟢 IN STOCK"))))`]);
                if (colQR > 0) qrFormulas.push([`=IF(B${r}="","",IMAGE("https://quickchart.io/qr?text=" & B${r} & "&size=100", 4, 100, 100))`]);
                if (colUpdate > 0) updateFormulas.push([`=IF(B${r}="","",TEXT(NOW(),"dd/MM/yyyy HH:mm"))`]);
             }

             // Batch update formulas (konsep yang diambil dari REST API v4 batchUpdate tapi native GAS)
             if (colCurrentStock > 0) sh.getRange(2, colCurrentStock, numRows, 1).setFormulas(currentStockFormulas);
             if (colStatus > 0) sh.getRange(2, colStatus, numRows, 1).setFormulas(statusFormulas);
             if (colQR > 0) sh.getRange(2, colQR, numRows, 1).setFormulas(qrFormulas);
             if (colUpdate > 0) sh.getRange(2, colUpdate, numRows, 1).setFormulas(updateFormulas);
          }
          executedCount++;
        }
        else if (act.cmd === "TIDY_UP") {
          // Sweep all columns
          let numCols = sh.getMaxColumns();
          if (numCols > 0) {
            sh.autoResizeColumns(1, numCols);
          }
          // Wrap text & vertical align on data range
          let dataRange = sh.getDataRange();
          if (dataRange) {
            dataRange.setWrap(true);
            dataRange.setVerticalAlignment("middle");
          }
          executedCount++;
        }
        else if (act.cmd === "SET_COLUMN_WIDTH") {
          let startCol = parseInt(act.startCol) || 1;
          let width = parseInt(String(act.width).replace(/\D/g, '')) || 100;
          let numCols = act.numCols;
          if (!numCols || typeof numCols === 'string') numCols = sh.getMaxColumns() - startCol + 1;
          if (numCols > 0) {
            sh.setColumnWidths(startCol, numCols, width);
            executedCount++;
          }
        }
        else if (act.cmd === "AUTO_RESIZE_COLUMNS") {
          let startCol = parseInt(act.startCol) || 1;
          let numCols = act.numCols;
          if (!numCols || typeof numCols === 'string') numCols = sh.getMaxColumns() - startCol + 1;
          if (numCols > 0) {
            // Because autoResizeColumns with multiple columns might be slower, try catch it
            sh.autoResizeColumns(startCol, numCols);
            executedCount++;
          }
        }
        else if (act.cmd === "SET_ROW_HEIGHT") {
          let startRow = parseInt(act.startRow) || 1;
          let height = parseInt(String(act.height).replace(/\D/g, '')) || 21;
          let numRows = act.numRows;
          if (!numRows || typeof numRows === 'string') numRows = sh.getMaxRows() - startRow + 1;
          if (numRows > 0) {
            sh.setRowHeights(startRow, numRows, height);
            executedCount++;
          }
        }
        else if ((act.cmd === "SET_BACKGROUND_COLOR" || act.cmd === "SET_BACKGROUND") && act.range && act.color) {
          sh.getRange(act.range).setBackground(act.color);
          executedCount++;
        }
        else if (act.cmd === "SET_FONT_WEIGHT" && act.range && act.weight) {
          sh.getRange(act.range).setFontWeight(act.weight);
          executedCount++;
        }
        else if (act.cmd === "SET_FONT_COLOR" && act.range && act.color) {
          sh.getRange(act.range).setFontColor(act.color);
          executedCount++;
        }
        else if (act.cmd === "SET_HORIZONTAL_ALIGNMENT" && act.range && act.alignment) {
          sh.getRange(act.range).setHorizontalAlignment(act.alignment);
          executedCount++;
        }
        else if (act.cmd === "CLEAR_CONTENT" && act.range) {
          sh.getRange(act.range).clearContent();
          executedCount++;
        }
      } catch (e) {
        // Skip invalid formats silently
      }
    }

    let result = { success: false, message: "", count: executedCount };
    if (executedCount > 0) {
      result.success = true;
      result.message = `✅ **FORMAT SUCCESSFUL**: ${executedCount} format actions have been applied.`;
      if (source) _sendAgentMsg(source, chatId, result.message);
    } else {
      result.message = "⚠️ Tidak ada format yang dieksekusi. Mungkin format data tidak valid.";
      if (source) _sendAgentMsg(source, chatId, result.message);
    }
    return result;
  } catch (e) {
    let errMsg = "❌ **FORMAT FAILED**: " + e.message;
    if (source) _sendAgentMsg(source, chatId, errMsg);
    return { success: false, message: errMsg, count: 0 };
  }
}
