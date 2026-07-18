/**
 * Smart Action Center Backend Controller
 * Handles Chat UI requests (Onboarding, Wipe, Format, File uploads)
 */

function processSmartChat(payload) {
  try {
    const text = (payload.text || "").trim();
    let fullContext = text;
    
    // 1. Handle File Attachment
    if (payload.fileData && payload.fileName) {
      const decodedData = Utilities.base64Decode(payload.fileData);
      const blob = Utilities.newBlob(decodedData, payload.fileType, payload.fileName);
      
      const extractedText = _extractTextFromBlob(blob);
      if (extractedText) {
        fullContext += `\n\n[FILE CONTENT: ${payload.fileName}]\n${extractedText}`;
      } else {
         return "❌ Gagal mengekstrak teks dari file. Pastikan file valid.";
      }
    }

    // 2. Handle Google Sheets Link
    const urlMatch = text.match(/https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (urlMatch) {
       const extractedText = _extractTextFromUrl(urlMatch[0]);
       if (extractedText) {
          fullContext += `\n\n[GOOGLE SHEETS CONTENT]\n${extractedText}`;
       } else {
          return "❌ Gagal mengekstrak teks dari tautan Google Sheets.";
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
1. CLEAR_CONTENTS (butuh: sheet, range) -> hapus isi data tanpa hapus format. Contoh range: "A2:Z" untuk hapus semua baris data kecuali header.
2. CLEAR_FORMATS (butuh: sheet, range) -> hapus format saja, data tetap.
3. DELETE_COLUMNS (butuh: sheet, start, count) -> hapus kolom secara fisik. 'start' = nomor kolom (1=A, 2=B, ..., 10=J). 'count' = jumlah kolom yang dihapus.
4. DELETE_ROWS (butuh: sheet, start, count) -> hapus baris secara fisik. 'start' = nomor baris. 'count' = jumlah baris.
5. ASK_USER (butuh: question) -> jika ambigu, isi 'question' dengan pertanyaan spesifik dalam bahasa Indonesia.

**Rules:**
- "kolom nomor 10" = kolom ke-10 (kolom J). Gunakan DELETE_COLUMNS dengan start:10, count:1.
- "hapus semua kolom dari 10 sampai 13" = DELETE_COLUMNS dengan start:10, count:4.
- "hapus baris 5 sampai 20" = DELETE_ROWS dengan start:5, count:16.
- "hapus semua data di inventory" = CLEAR_CONTENTS dengan range "A2:Z".
- "hapus isi kolom J" = CLEAR_CONTENTS dengan range "J2:J".
- Jika user menyebut nama sheet, gunakan nama itu. Default: "Inventory".
- JANGAN pernah hapus baris 1 (header).

**User Request:** "${wipeCmd}"`;
           
           const aiRaw = callAI(wipePrompt);
           try {
             let actionArr = AIAgent._extractJson(aiRaw);
             if (!Array.isArray(actionArr)) {
               if (actionArr && actionArr.cmd === "ASK_USER") {
                  return `🤔 AI Membutuhkan Klarifikasi:\n${actionArr.question}`;
               }
               actionArr = [actionArr];
             }
             if (actionArr.length === 0 || !actionArr[0] || !actionArr[0].cmd) {
               return "❌ AI gagal menerjemahkan permintaan wipe Anda menjadi tindakan valid.";
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
               return `✅ Sukses! Wipe berbasis AI berhasil diterapkan (${executedCount} tindakan).`;
             } else {
               return "⚠️ Tidak ada tindakan wipe yang berhasil dieksekusi. Periksa nama sheet dan parameter.";
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
         return `✅ Sukses! Fitur /wipe berhasil membersihkan data pada ${wiped} lembar tabel tanpa menghapus rumus.`;
       } catch (e) {
         return "❌ Gagal melakukan Wipe: " + e.message;
       }
    }

    
    if (lowerText.startsWith("/format")) {
       try {
         const formatCmdText = lowerText.replace("/format", "").trim();
         
         if (formatCmdText === "") {
           return "Silakan jelaskan bagian mana yang ingin Anda format. Contoh: /format warnai header tabel menjadi merah";
         } 
         
         const actualFormatCmdText = formatCmdText;
         
         if (actualFormatCmdText !== "") {
           // AI Formatting
           let actionArr = AIAgent.parseFormattingAI(actualFormatCmdText);
           if (!Array.isArray(actionArr)) {
             if (actionArr && actionArr.cmd === "ASK_USER") {
                return `🤔 AI Membutuhkan Klarifikasi:\n${actionArr.question}`;
             }
             actionArr = [actionArr];
           }
           if (actionArr.length === 0 || !actionArr[0] || !actionArr[0].cmd) {
             return "❌ AI gagal menerjemahkan permintaan format Anda menjadi tindakan valid.";
           }
           
           // Run the format actions with null source and chatId
           const result = _executeFormatAction(actionArr, null, null);
           
           if (result.success) {
             return `✅ Sukses! Format khusus berbasis AI berhasil diterapkan (${result.count} tindakan).`;
           } else {
             return `⚠️ Beberapa format mungkin gagal: ${result.message}`;
           }
         }
       } catch (e) {
          return "❌ Gagal melakukan Format: " + e.message;
       }
    }
    
    if (lowerText.startsWith("/onboarding")) {
       if (lowerText === "/onboarding" && !payload.fileData) {
          return "Silakan masukkan detail barang baru, lampirkan file nota, atau paste link Google Sheets. Contoh: '/onboarding 10pcs Macbook Pro M4 harga 25000000'";
       }
               const prompt = `**Task:** Ekstrak data registrasi barang baru dari teks berikut:\n\n"${fullContext}"\n\n**CRITICAL CONSTRAINTS - VIOLATING THESE WILL CAUSE A SYSTEM CRASH:**\n- DO NOT print conversational filler.\n- DO NOT think step-by-step.\n- YOU MUST ENCLOSE YOUR ENTIRE JSON OUTPUT WITHIN A MARKDOWN BLOCK (\`\`\`json ... \`\`\`).\n\n**Output Format Example:**\n\`\`\`json\n[{"item_code":"kode_barang", "new_item_name":"nama_barang_lengkap", "quantity":jumlah_stok, "min_stock":batas_minimum, "unit":"satuan", "new_price":harga_jual, "buy_price":harga_beli, "branch":"nama_cabang", "new_category":"kategori_barang"}]\n\`\`\`\n\nJika data SANGAT membingungkan sehingga Anda TIDAK BISA menebak nama barangnya sama sekali, keluarkan JSON: \`\`\`json\n{ "cmd": "ASK_USER", "question": "Tuliskan pertanyaan spesifik Anda di sini, JANGAN gunakan titik-titik (...)" }\n\`\`\`\n\n**Rules:**\n- Ekstrak sebanyak mungkin data yang ada (seperti kode, nama, stok, kategori).\n- Abaikan teks yang berulang atau typo dari user, rangkai nama barang dengan logis.\n- Jika ada data yang tidak disebutkan, abaikan (jangan dimasukkan ke JSON) atau setel ke 0.\n- SELALU utamakan mengekstrak menjadi array of JSON daripada meminta klarifikasi, asalkan ada nama barang.`;
       const aiRaw = callAI(prompt);
       
       try {
          let parsedData = AIAgent._extractJson(aiRaw);
          
          if (parsedData && !Array.isArray(parsedData) && parsedData.cmd === "ASK_USER") {
             return `🤔 AI Membutuhkan Klarifikasi:\n${parsedData.question}`;
          }
          
          let items = Array.isArray(parsedData) ? parsedData : [parsedData];
          
          if (!Array.isArray(items) || items.length === 0) {
             return "❌ AI gagal mengekstrak data barang dari teks/file yang diberikan.";
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
          
          let msg = `✅ Sukses! ${successCount} barang baru berhasil diregistrasi.`;
          if (errors.length > 0) msg += `\n⚠️ Gagal meregistrasi ${errors.length} barang:\n- ` + errors.join("\n- ");
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
       return `✅ Sukses memproses ${result.count} data transaksi dari pesan Anda.`;
    } else if (result.isQuestion) {
       return `🤔 AI Membutuhkan Klarifikasi:\n${result.message}`;
    } else {
       return `❌ Gagal memproses transaksi: ${result.message}`;
    }

  } catch (err) {
    return "❌ System Error: " + err.message;
  }
}
