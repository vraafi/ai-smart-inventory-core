/**
 * Smart Action Center Backend Controller
 * Handles Chat UI requests (Onboarding, Wipe, Format, File uploads)
 */

function processSmartChat(payload) {
  try {
    let text = (payload.text || "").trim();
    
    // --- STATE MANAGEMENT UNTUK WEB UI ---
    const webUiChatId = "WEB_UI_USER";
    const pendingState = _getPendingState(webUiChatId);
    if (pendingState && pendingState.originalText) {
      // Combine previous text with user's new answer
      text = pendingState.originalText + "\n[AI Asked: " + pendingState.aiQuestion + "]\n[User Answer: " + text + "]";
      _clearPendingState(webUiChatId);
    }
    
    let fullContext = text;
    
    // 1. Handle File Attachment
    if (payload.fileData && payload.fileName) {
      const decodedData = Utilities.base64Decode(payload.fileData);
      const blob = Utilities.newBlob(decodedData, payload.fileType, payload.fileName);
      
      const extractedText = _extractTextFromBlob(blob);
      if (extractedText) {
        fullContext += `\n\n[FILE CONTENT: ${payload.fileName}]\n${extractedText}`;
      } else {
         return "❌ Failed extracting text from file. Ensure file is valid.";
      }
    }

    // 2. Handle Google Sheets Link
    const urlMatch = text.match(/https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (urlMatch) {
       const extractedText = _extractTextFromUrl(urlMatch[0]);
       if (extractedText) {
          fullContext += `\n\n[GOOGLE SHEETS CONTENT]\n${extractedText}`;
       } else {
          return "❌ Failed extracting text from Google Sheets link.";
       }
    }

    // 3. Handle Commands
    const lowerText = text.toLowerCase();
    
    if (lowerText.startsWith("/wipe")) {
       const wipeCmd = text.replace(/^\/wipe/i, "").trim().toLowerCase();
       if (wipeCmd !== "" && wipeCmd !== "confirm") {
           // AI-powered wipe: use a dedicated wipe prompt with destructive commands
           const wipePrompt = `**Task:** Translate the user's Google Sheets DATA DELETION request into a JSON array of wipe commands.
**CRITICAL CONSTRAINTS - VIOLATING THESE WILL CAUSE A SYSTEM CRASH:**
- DO NOT print "User request:", "Interpretation:", "Action:", "Sheet:", or "Range:".
- DO NOT use bullet points or asterisks (*).
- DO NOT think step-by-step.
- YOU MUST ENCLOSE YOUR ENTIRE JSON OUTPUT WITHIN A MARKDOWN BLOCK (\`\`\`json ... \`\`\`).

**Available Commands (cmd):**
1. CLEAR_CONTENTS (needs: sheet, range) -> clear data content without clearing format. Example range: "A2:Z" to clear all rows except header.
2. CLEAR_FORMATS (needs: sheet, range) -> clear format only, data stays.
3. DELETE_COLUMNS (needs: sheet, start, count) -> physically delete columns. 'start' = column number (1=A, 2=B, ..., 10=J). 'count' = number of deleted columns.
4. DELETE_ROWS (needs: sheet, start, count) -> physically delete rows. 'start' = row number. 'count' = number of rows.
5. ASK_USER (needs: question) -> if ambiguous, fill 'question' with a specific English question.

**Rules:**
- "column number 10" = 10th column (column J). Use DELETE_COLUMNS with start:10, count:1.
- "delete all columns from 10 to 13" = DELETE_COLUMNS with start:10, count:4.
- "delete rows 5 to 20" = DELETE_ROWS with start:5, count:16.
- "delete all data in inventory" = CLEAR_CONTENTS with range "A2:Z".
- "delete content of column J" = CLEAR_CONTENTS with range "J2:J".
- If user mentions a sheet name, use that name. Default: "Inventory".
- NEVER delete row 1 (header).

**User Request:** "${wipeCmd}"`;
           
           const aiRaw = callAI(wipePrompt);
           try {
             let actionArr = AIAgent._extractJson(aiRaw);
             if (!Array.isArray(actionArr)) {
               if (actionArr && actionArr.cmd === "ASK_USER") {
                  const webUiChatId = "WEB_UI_USER";
                  const qMsg = `🤔 AI Needs Clarification:\n${actionArr.question}`;
                  _savePendingState(webUiChatId, fullContext, qMsg);
                  return qMsg;
               }
               actionArr = [actionArr];
             }
             if (actionArr.length === 0 || !actionArr[0] || !actionArr[0].cmd) {
               return "❌ AI Failed to translate your wipe request into valid actions.";
             }
             
             // Execute using the WIPE handler (supports DELETE_COLUMNS, DELETE_ROWS, CLEAR_CONTENTS)
             const ss = SpreadsheetApp.getActiveSpreadsheet();
             let executedCount = 0;
             for (let act of actionArr) {
               const sh = ss.getSheetByName(act.sheet || "Inventory");
               if (!sh) continue;
               try {
                 if (act.cmd === "CLEAR_CONTENTS" && act.range) {
                   let rStr = act.range;
                   if (rStr.endsWith("Z") && !rStr.match(/\d+$/)) rStr += sh.getMaxRows();
                   sh.getRange(rStr).clearContent(); executedCount++;
                 }
                 else if (act.cmd === "CLEAR_FORMATS" && act.range) {
                   let rStr = act.range;
                   if (rStr.endsWith("Z") && !rStr.match(/\d+$/)) rStr += sh.getMaxRows();
                   sh.getRange(rStr).clearFormat(); executedCount++;
                 }
                 else if (act.cmd === "DELETE_COLUMNS" && act.start && act.count) {
                   let count = act.count;
                   if (act.start <= sh.getMaxColumns()) {
                     if (act.start + count - 1 >= sh.getMaxColumns()) count = sh.getMaxColumns() - act.start + 1;
                     if (act.start === 1 && count === sh.getMaxColumns()) count = count - 1;
                     if (count > 0) { sh.deleteColumns(act.start, count); executedCount++; }
                   }
                 }
                 else if (act.cmd === "DELETE_ROWS" && act.start && act.count) {
                   let count = act.count;
                   if (act.start <= sh.getMaxRows()) {
                     if (act.start + count - 1 >= sh.getMaxRows()) count = sh.getMaxRows() - act.start + 1;
                     if (act.start === 1 && count === sh.getMaxRows()) count = count - 1;
                     if (count > 0) { sh.deleteRows(act.start, count); executedCount++; }
                   }
                 }
               } catch(e) { /* skip invalid action */ }
             }
             
             if (executedCount > 0) {
               return `✅ Success! AI-based wipe successfully applied (${executedCount} actions).`;
             } else {
               return "⚠️ No wipe actions successfully executed. Check sheet name and parameters.";
             }
           } catch(e) {
             return "❌ Error parsing wipe AI data: " + e.message;
           }
       }
       
       // Default /wipe tanpa argumen: bersihkan semua data Inventory & Transactions
       const ss = SpreadsheetApp.getActiveSpreadsheet();
       let wiped = 0;
       try {
         const invSh = ss.getSheetByName(SHEETS.INVENTORY);
         if (invSh) {
           invSh.getRange("A2:H" + invSh.getMaxRows()).clearContent();
           invSh.getRange("J2:M" + invSh.getMaxRows()).clearContent();
           wiped++;
         }
         const trxSh = ss.getSheetByName(SHEETS.TRANSACTIONS);
         if (trxSh) {
           trxSh.getRange("A2:N" + trxSh.getMaxRows()).clearContent();
           wiped++;
         }
         return `✅ Success! /wipe feature successfully cleared data in ${wiped} sheets without deleting formulas.`;
       } catch (e) {
         return "❌ Failed melakukan Wipe: " + e.message;
       }
    }

    
    if (lowerText.startsWith("/format")) {
       try {
         const formatCmdText = lowerText.replace("/format", "").trim();
         
         if (formatCmdText === "") {
           return "Please explain which part you want to format. Example: /format color the table header red";
         } 
         
         const actualFormatCmdText = formatCmdText;
         
         if (actualFormatCmdText !== "") {
           // AI Formatting
           let actionArr = AIAgent.parseFormattingAI(actualFormatCmdText);
           if (!Array.isArray(actionArr)) {
             if (actionArr && actionArr.cmd === "ASK_USER") {
                const webUiChatId = "WEB_UI_USER";
                const qMsg = `🤔 AI Needs Clarification:\n${actionArr.question}`;
                _savePendingState(webUiChatId, fullContext, qMsg);
                return qMsg;
             }
             actionArr = [actionArr];
           }
           if (actionArr.length === 0 || !actionArr[0] || !actionArr[0].cmd) {
             return "❌ AI Failed to translate your formatting request into valid actions.";
           }
           
           // Run the format actions with null source and chatId
           const result = _executeFormatAction(actionArr, null, null);
           
           if (result.success) {
             return `✅ Success! AI-based custom formatting successfully applied (${result.count} actions).`;
           } else {
             return `⚠️ Beberapa format mungkin Failed: ${result.message}`;
           }
         }
       } catch (e) {
          return "❌ Failed melakukan Format: " + e.message;
       }
    }
    
    if (lowerText.startsWith("/onboarding")) {
       if (lowerText === "/onboarding" && !payload.fileData) {
          return "Please enter New item details, attach a receipt file, or paste a Google Sheets link. Example: '/onboarding 10pcs Macbook Pro M4 price 2500'";
       }
               const prompt = `**Task:** Extract New item registration data from the following text:\n\n"${fullContext}"\n\n**CRITICAL CONSTRAINTS - VIOLATING THESE WILL CAUSE A SYSTEM CRASH:**\n- DO NOT print conversational filler.\n- DO NOT think step-by-step.\n- YOU MUST ENCLOSE YOUR ENTIRE JSON OUTPUT WITHIN A MARKDOWN BLOCK (\`\`\`json ... \`\`\`).\n\n**Output Format Example:**\n\`\`\`json\n[{"item_code":"kode_barang", "new_item_name":"item_full_name", "quantity":stock_amount, "min_stock":minimum_limit, "unit":"Unit", "new_price":sell_price, "buy_price":buy_price, "branch":"branch_name", "new_category":"item_category"}]\n\`\`\`\n\nIf the data is SO confusing that you CANNOT guess the Item Name at all, output JSON: \`\`\`json\n{ "cmd": "ASK_USER", "question": "Write your specific English question here, DO NOT use ellipses (...)" }\n\`\`\`\n\n**Rules:**\n- Extract as much data as possible (like code, name, Stock, Category).\n- Ignore repetitive text or typos from the user, construct the Item Name logically.\n- If any data is not mentioned, ignore it (do not include in JSON) or set to 0.\n- ALWAYS prioritize extracting into a JSON array over asking for clarification, as long as there is an Item Name.`;
       const aiRaw = callAI(prompt);
       
       try {
          let parsedData = AIAgent._extractJson(aiRaw);
          
          if (parsedData && !Array.isArray(parsedData) && parsedData.cmd === "ASK_USER") {
             const webUiChatId = "WEB_UI_USER";
             const qMsg = `🤔 AI Membutuhkan Klarifikasi:\n${parsedData.question}`;
             _savePendingState(webUiChatId, fullContext, qMsg);
             return qMsg;
          }
          
          let items = Array.isArray(parsedData) ? parsedData : [parsedData];
          
          if (!Array.isArray(items) || items.length === 0) {
             return "❌ AI Failed mengekstrak data barang dari teks/file yang diberikan.";
          }
          
          let successCount = 0;
          let errors = [];
          let debugInfo = "";
          for (let item of items) {
             if (!item.new_item_name) {
                errors.push("Data tidak memiliki 'new_item_name'");
                continue;
             }
             const result = processOnboarding(item);
             if (result.success) {
               successCount++;
               if (!debugInfo && result._debug) debugInfo = result._debug;
             }
             else errors.push(`${item.new_item_name}: ${result.message}`);
          }
          
          let msg = `✅ Success! ${successCount} New item berhasil diregistrasi.`;
          if (errors.length > 0) msg += `\n⚠️ Failed meregistrasi ${errors.length} barang:\n- ` + errors.join("\n- ");
          if (debugInfo) msg += `\n🔍 DEBUG: ${debugInfo}`;
          return msg;
       } catch (e) {
          return "❌ Error parsing onboarding AI data: " + e.message;
       }
    }

    // 4. Default: Send to AI Autonomous Input
    if (!fullContext) return "Silakan ketik sesuatu atau lampirkan file.";
    
    const result = processAutonomousInput(fullContext);
    
    if (result.success) {
       return `✅ Success memproses ${result.count} data transaksi dari pesan Anda.`;
    } else if (result.isQuestion) {
       const webUiChatId = "WEB_UI_USER";
       const qMsg = `🤔 AI Membutuhkan Klarifikasi:\n${result.message}`;
       _savePendingState(webUiChatId, fullContext, qMsg);
       return qMsg;
    } else {
       return `❌ Failed memproses transaksi: ${result.message}`;
    }

  } catch (err) {
    return "❌ System Error: " + err.message;
  }
}
