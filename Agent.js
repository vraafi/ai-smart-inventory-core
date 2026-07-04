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
    sheet.appendRow([new Date(), typeof msg === 'object' ? JSON.stringify(msg) : msg]);
  } catch(e) {
    Logger.log("debugLog error: " + e);
  }
}

// ─── WEBHOOK ENTRY POINT ─────────────────────────────────────
function doPost(e) {
  debugLog("doPost Triggered! Event: " + JSON.stringify(e));
  try {
    const props = PropertiesService.getScriptProperties();
    const secret = props.getProperty("WEBHOOK_SECRET");
    if (secret && e.parameter.token !== secret) {
      return ContentService.createTextOutput("Forbidden: Invalid Webhook Token");
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
      const msg = body.message;
      if (msg) {
        const chatId = msg.chat.id.toString();
        const text   = msg.text || "";
        const from   = msg.from;
        const name   = [from.first_name, from.last_name].filter(Boolean).join(" ") || from.username || "Employee";
        
        // Save first person who chats as ADMIN automatically for push notifications
        if (text) {
          const existingAdmin = props.getProperty("ADMIN_CHAT_ID");
          if (!existingAdmin) {
            props.setProperty("ADMIN_CHAT_ID", chatId);
            debugLog("Saved new ADMIN_CHAT_ID: " + chatId);
          }
          debugLog("Processing TG Message from " + name + ": " + text);
          _processTelegramMessage(chatId, text, name);
        }
      }
    }
  } catch (err) {
    Logger.log("doPost error: " + err);
    return ContentService.createTextOutput("ERROR: " + err + " | " + err.stack);
  }
  return ContentService.createTextOutput("OK");
}

// ─── WHATSAPP MESSAGE HANDLER ──────────────────────────────
function _processWhatsAppMessage(phone, text, senderName) {
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
function _processTelegramMessage(chatId, text, senderName) {
  if (text === "/start" || text === "/help") {
    _sendTelegram(chatId, _buildWelcomeMessage(senderName));
    return;
  }
  if (text === "/status" || text === "/dashboard") {
    _sendAgentMsg("telegram", chatId, _buildDashboardSummary());
    return;
  }
  if (text.startsWith("/check") || text.startsWith("/stock")) {
    const q = text.replace(/\/(check|stock)\s*/i, "").trim();
    _sendAgentMsg("telegram", chatId, _buildStockCheckResult(q));
    return;
  }
  // Natural language — process with AI
  _sendTelegram(chatId, "⏳ AI is processing your report...");
  _processWithAI(chatId, text, senderName, "Telegram");
}

// ─── WHATSAPP POLLING (backup if webhook fails) ──────────────
function pollWhatsApp() {
  // WhatsApp Cloud API doesn't support polling — use webhook only
  // This trigger is intentionally a no-op
}

// ─── TELEGRAM POLLING (backup for webhook) ───────────────────
function pollTelegram() {
  const props = PropertiesService.getScriptProperties();
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
        if (text) {
          debugLog("Processing TG Message from " + name + ": " + text);
          try {
            _processTelegramMessage(chatId, text, name);
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
  try {
    const code = "AI Inventory Report";
    // Search for unread emails that have the subject code
    const threads = GmailApp.search(`is:unread subject:"${code}"`, 0, 10);
    
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        if (!msg.isUnread()) return;
        const subject   = msg.getSubject();
        const rawBody   = msg.getPlainBody().substring(0, 1500);
        const from      = msg.getFrom();
        const senderName = from.replace(/<.*>/g, "").trim() || from;
        const senderEmail = (from.match(/<(.+)>/) || [, from])[1];
        
        // Combine Subject and Body for AI Context
        const fullContext = `Subject: ${subject}\n\nBody: ${rawBody}`;

        // Process with AI and pass the senderEmail so it can reply
        _processWithAI(senderEmail, fullContext, senderName, "Email");
        msg.markRead();
      });
    });
  } catch (err) { Logger.log("pollEmails error: " + err); }
}

// ─── CORE AI PROCESSING ──────────────────────────────────────
function _processWithAI(chatId, rawText, senderName, source) {
  const cache = CacheService.getScriptCache();
  const wipeStateKey = "WIPE_STATE_" + chatId;

  // 1. If user sends /wipe command
  if (rawText.trim().toLowerCase().startsWith("/wipe")) {
    _sendAgentMsg(source, chatId, "⏳ AI is analyzing your destructive request...");
    
    const wipePrompt = `You are a destructive AI Action Engine for a Google Sheets Inventory System.
The user wants to wipe/delete data or structure: "${rawText}".
We have 4 main sheets: Dashboard, Inventory, Transactions, Branches.

Generate exactly 4 options (a, b, c, d) for the user to confirm. 
Option 'a' MUST be the exact destructive action they requested.
Option 'b' MUST be a safer alternative (e.g. clear contents only).
Option 'c' is always "Add your own custom instruction".
Option 'd' is always "Cancel".

For options 'a' and 'b', you must generate a JSON array of specific Apps Script execution commands.
Supported commands:
- {"cmd": "CLEAR_CONTENTS", "sheet": "SheetName", "range": "A2:Z1000"} (clears text only)
- {"cmd": "CLEAR_FORMATS", "sheet": "SheetName", "range": "A2:Z1000"} (clears colors/designs)
- {"cmd": "DELETE_COLUMNS", "sheet": "SheetName", "start": 1, "count": 5} (deletes physical columns)
- {"cmd": "DELETE_ROWS", "sheet": "SheetName", "start": 2, "count": 100} (deletes physical rows)

Return ONLY a valid JSON object exactly like this:
{
  "message": "⚠️ **Destructive Risk Warning!** ⚠️\\n\\na: [Describe action A and its risk]\\nb: [Describe action B as a safer alternative]\\nc: Add your own custom instruction\\nd: Cancel\\n\\nReply with a/b/c/d.",
  "actions": {
    "a": [{"cmd": "...", "sheet": "...", ...}],
    "b": [{"cmd": "...", "sheet": "...", ...}]
  }
}`;
    
    let aiRaw = callAI(wipePrompt, "You are a Google Apps Script JSON generator.");
    const jsonMatch = aiRaw.match(/\{[\s\S]*\}/);
    if (jsonMatch) {
      try {
        let parsed = JSON.parse(jsonMatch[0]);
        cache.put(wipeStateKey, JSON.stringify(parsed.actions), 300); // Wait for 5 minutes
        _sendAgentMsg(source, chatId, parsed.message);
      } catch (e) {
        _sendAgentMsg(source, chatId, "❌ AI failed to formulate an action. Please try another instruction.");
      }
      return;
    }
  }

  // 2. If user is in WIPE_PENDING state
  const pendingWipeStr = cache.get(wipeStateKey);
  if (pendingWipeStr) {
    const ans = rawText.trim().toLowerCase();
    
    if (ans === "d") {
      cache.remove(wipeStateKey);
      _sendAgentMsg(source, chatId, "✅ Wipe process **cancelled**.");
      return;
    }
    if (ans === "c") {
      cache.remove(wipeStateKey);
      _sendAgentMsg(source, chatId, "🔄 State reset. Please send a new /wipe command with your specific instructions.");
      return;
    }
    
    let actionsObj;
    try { actionsObj = JSON.parse(pendingWipeStr); } catch(e) {}
    
    if (actionsObj && (ans === "a" || ans === "b")) {
      cache.remove(wipeStateKey);
      const actionArr = actionsObj[ans];
      if (!actionArr || !Array.isArray(actionArr)) {
         _sendAgentMsg(source, chatId, "❌ Opsi tidak memiliki aksi yang valid.");
         return;
      }
      _executeWipeAction(actionArr, source, chatId);
      return;
    }
    
    // If they reply anything else while pending
    _sendAgentMsg(source, chatId, "❌ Invalid choice. Please reply with a, b, c, or d. Or type 'd' to cancel.");
    return;
  }

  const itemContext = _getInventoryContext();
  
  // Anti Prompt-Injection: Escape HTML/XML chars to prevent breaking out of tags
  const sanitizedText = rawText.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  const prompt = `You are a strict inventory data parser. Extract transactions from the employee report below.

<employee_report>
${sanitizedText}
</employee_report>

AVAILABLE ITEMS (SKU | Name | Category | CurrentStock):
${itemContext}

RULES:
1. Match items by EXACT NAME. "Dell Latitude 5430" must match "Dell Latitude 5430", NOT "ThinkPad" or any other item.
2. "MacBook Pro 14 M2" must match "MacBook Pro 14 M2", NOT "ThinkPad T14" or similar.  
3. "Mouse Logitech K380" or "Logitech K380" must match "Logitech K380", NOT any laptop.
4. If the message contains "/onboarding" and no exact match exists, set item_new=true, item_code="NEW", and extract new_item_name, new_category, new_price. If the message does NOT contain "/onboarding", NEVER set item_new=true; instead, assume it is a typo and match it to the closest existing item.
5. Use the SKU code from the AVAILABLE ITEMS list above as item_code (e.g. "LAP-003", "ACC-002").
6. Transaction types:
   - "IN" = receiving/restocking (kedatangan, masuk gudang, terima dari supplier)
   - "OUT" = selling/dispatching (laku, terjual, kirim, dibawa kurir)
   - "OUT" = returns/retur where customer returns defective item AND user says reduce stock (dikurangi)
   - "ADJUSTMENT" = physical count correction ONLY (e.g. "stok fisik = 50 unit")
   - "CHECK" = stock inquiry
   - "UNKNOWN" = unrelated to inventory
7. quantity must always be a positive number.
8. Each separate item in the report = separate transaction object.

Respond ONLY with a JSON array. No other text:
[
  {
    "type": "IN|OUT|ADJUSTMENT|CHECK|UNKNOWN",
    "item_code": "exact SKU from list above or NEW",
    "item_name": "exact item name from list above",
    "quantity": 0,
    "branch": null,
    "notes": "brief notes",
    "confidence": 90,
    "item_new": false,
    "new_item_name": null,
    "new_category": "General",
    "new_price": 0,
    "ai_reasoning": "why this match"
  }
]`;

  let parsedList;
  let aiRaw = "";
  debugLog("Sending to Groq AI...");
  try {
    aiRaw = callAI(prompt, null);
    debugLog("AI Raw Output: " + aiRaw);
    const jsonMatch = aiRaw.match(/\[[\s\S]*\]|\{[\s\S]*\}/);
    if (jsonMatch) {
      let parsed = JSON.parse(jsonMatch[0]);
      // Unwrap if AI returned {transactions: [...]} or {data: [...]} etc.
      if (!Array.isArray(parsed) && typeof parsed === 'object') {
        const keys = Object.keys(parsed);
        for (const key of keys) {
          if (Array.isArray(parsed[key])) {
            parsed = parsed[key];
            break;
          }
        }
      }
      parsedList = Array.isArray(parsed) ? parsed : [parsed];
    } else {
      throw new Error("No JSON block found");
    }
  } catch (err) {
    Logger.log("AI error: " + err);
    const errStr = err.toString().toLowerCase();
    
    if (errStr.includes("failed") || errStr.includes("unauthorized") || errStr.includes("api key") || errStr.includes("key not valid") || errStr.includes("provider")) {
      if (chatId) _sendAgentMsg(source, chatId,
        "⚠️ AI System is not configured properly or API Key is invalid. Please contact the Admin to set the API Key in the Agent Settings menu."
      );
      return;
    }
    
    if (chatId) _sendAgentMsg(source, chatId,
      `❌ Could not parse AI response. \n\nRAW AI OUTPUT:\n${aiRaw}\n\nPlease try a clearer format:\n• "received 10 [item name]"`
    );
    return;
  }

  debugLog("Parsed " + parsedList.length + " transactions from AI");
  for (let idx = 0; idx < parsedList.length; idx++) {
    const parsed = parsedList[idx];
    debugLog("Processing tx #" + (idx+1) + ": type=" + parsed.type + " code=" + parsed.item_code + " name=" + parsed.item_name + " qty=" + parsed.quantity + " conf=" + parsed.confidence);
    if (parsed.type === "UNKNOWN") {
      if (chatId) _sendAgentMsg(source, chatId,
        `🤔 I couldn't determine the intent of this part of your report.\n\nAI Note: ${parsed.ai_reasoning || "Not specified"}\n\nPlease clarify: is this a Stock IN, Stock OUT, or Stock Adjustment?`
      );
      continue;
    }

    if (parsed.type === "CHECK") {
      if (chatId) _sendAgentMsg(source, chatId, _buildStockCheckResult(parsed.item_name || parsed.item_code));
      continue;
    }

    if (parsed.quantity <= 0 && parsed.type !== "ADJUSTMENT") {
      if (chatId) _sendAgentMsg(source, chatId, `❌ Could not detect a valid quantity for ${parsed.item_name}.`);
      continue;
    }

    if (parsed.confidence < 75) {
      if (chatId) {
        CacheService.getUserCache().put("pending_" + chatId, JSON.stringify({ parsed, rawText, senderName, source }), 300);
        _sendAgentMsg(source, chatId,
          `🤔 AI is ${parsed.confidence}% confident about ${parsed.item_name}. Please confirm:\n\n` +
          `Type: <b>${parsed.type}</b>\n` +
          `Qty: <b>${parsed.quantity}</b>\n\n` +
          `Reply <b>YES</b> to confirm.`
        );
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
    const item = _findItem(parsed.item_code, parsed.item_name);
    debugLog("_findItem result: " + (item ? "FOUND row=" + item.row + " name=" + item.name + " stock=" + item.stock : "NULL"));

  if (!item) {
    if (parsed.item_new === true) {
      if (rawText.toLowerCase().includes("/onboarding")) {
        item = _createNewItemRow(parsed);
        if (chatId) _sendAgentMsg(source, chatId, `✅ New item added: *${item.name}* (${item.code}). Category: ${parsed.new_category}, Price: Rp${parsed.new_price}. Continuing transaction...`);
      } else {
        if (chatId) _sendAgentMsg(source, chatId, `❌ **Transaction Denied:** You mentioned item "${parsed.new_item_name}" which is not registered. If this is a new item, you must use the /onboarding keyword. If this is an existing item, please correct your spelling (typo).`);
        return;
      }
    } else if (parsed.item_name) {
      if (chatId) _sendAgentMsg(source, chatId, `❌ Item "${parsed.item_name}" is not found in the database.`);
      return;
    } else {
      return;
    }
  } else {
    // If item exists, but AI thought it was new, or user used /onboarding but it matched an existing item.
    if (rawText.toLowerCase().includes("/onboarding") && parsed.item_new === true) {
      if (chatId) _sendAgentMsg(source, chatId, `⚠️ You used /onboarding for "${parsed.item_name}", but this item ALREADY EXISTS in the database with code ${item.code}. Item creation cancelled. Continuing stock update...`);
    }
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
  const adminTgId = PropertiesService.getScriptProperties().getProperty("TG_ADMIN_ID");
  if (adminTgId && adminTgId !== chatId) {
    _sendTelegram(adminTgId, `📢 <b>New Report from ${senderName} via ${source}</b>\n\n` + proof.message);
  }

  // Email admin
  const adminEmail = PropertiesService.getScriptProperties().getProperty("ADMIN_EMAIL");
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
    if (chatId) _sendAgentMsg(source, chatId, "❌ System is busy (Concurrency limit). Please repeat your report in a few seconds.");
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
    const htmlBody = text.replace(/\n/g, "<br>");
    _sendEmailNotification(recipient, "🤖 AI Inventory Notification", htmlBody);
  }
}

function _sendTelegram(chatId, text) {
  const token = PropertiesService.getScriptProperties().getProperty("TG_TOKEN");
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
  const props = PropertiesService.getScriptProperties();
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

function _sendEmailNotification(to, subject, htmlBody) {
  try { GmailApp.sendEmail(to, subject, "", { htmlBody }); }
  catch (e) { Logger.log("Email error: " + e); }
}

// ─── GEMINI AI ────────────────────────────────────────────────
function _callGemini(prompt, apiKey) {
  const url  = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 1024 },
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
  if (e.parameter["action"] === "get_logs") {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
    if (!sheet) return ContentService.createTextOutput("No Logs sheet");
    const data = sheet.getDataRange().getValues();
    const lastLogs = data.slice(-5);
    return ContentService.createTextOutput(JSON.stringify(lastLogs));
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
    const adminId = PropertiesService.getScriptProperties().getProperty("ADMIN_CHAT_ID");
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
    const ss = SpreadsheetApp.getActiveSpreadsheet();
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
