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
       const wipeCmd = lowerText.replace("/wipe", "").trim();
       if (wipeCmd !== "") {
           let actionArr = AIAgent.parseFormattingAI(wipeCmd);
           if (!Array.isArray(actionArr)) {
             if (actionArr && actionArr.cmd === "ASK_USER") {
                return `🤔 AI Membutuhkan Klarifikasi:\n${actionArr.question}`;
             }
             actionArr = [actionArr];
           }
           if (actionArr.length === 0 || !actionArr[0] || !actionArr[0].cmd) {
             return "❌ AI gagal menerjemahkan permintaan wipe Anda menjadi tindakan valid.";
           }
           const result = _executeFormatAction(actionArr, null, null);
           if (result.success) {
             return `✅ Sukses! Wipe/Format berbasis AI berhasil diterapkan (${result.count} tindakan).`;
           } else {
             return `⚠️ Wipe gagal: ${result.message}`;
           }
       }

       const ss = SpreadsheetApp.getActiveSpreadsheet();
       let wiped = 0;
       try {
         const invSh = ss.getSheetByName(SHEETS.INVENTORY);
         if (invSh) {
           invSh.getRange("A2:H").clearContent();
           invSh.getRange("J2:M").clearContent();
           wiped++;
         }
         const trxSh = ss.getSheetByName(SHEETS.TRANSACTIONS);
         if (trxSh) {
           trxSh.getRange("A2:N").clearContent();
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

       const prompt = `**Task:** Ekstrak data registrasi barang baru dari konteks teks/file berikut:\n\n"${fullContext}"\n\n**CRITICAL CONSTRAINTS - VIOLATING THESE WILL CAUSE A SYSTEM CRASH:**\n- DO NOT print "User request:", "Interpretation:", "Action:", "Sheet:", or "Range:".\n- DO NOT use bullet points or asterisks (*).\n- DO NOT think step-by-step.\n- YOU MUST START YOUR ENTIRE RESPONSE WITH THE CHARACTER '[' OR '{' and END WITH ']' OR '}'.\n\n**Output Format Example:**\n[{"new_item_name":"nama_barang_lengkap", "quantity":jumlah_stok_awal_angka, "new_price":harga_jual_angka}]\n\nJika data tidak lengkap atau membingungkan, keluarkan JSON object tunggal: { "cmd": "ASK_USER", "question": "Tuliskan pertanyaan Anda secara spesifik dalam bahasa Indonesia di sini (misal: 'Berapa stok awalnya?')" }\n\n**Rules:** Hanya keluarkan array/object JSON yang valid.`;
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
          for (let item of items) {
             if (!item.new_item_name) {
                errors.push("Data tidak memiliki 'new_item_name'");
                continue;
             }
             const result = processOnboarding(item);
             if (result.success) successCount++;
             else errors.push(`${item.new_item_name}: ${result.message}`);
          }

          let msg = `✅ Sukses! ${successCount} barang baru berhasil diregistrasi.`;
          if (errors.length > 0) msg += `\n⚠️ Gagal meregistrasi ${errors.length} barang:\n- ` + errors.join("\n- ");
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
