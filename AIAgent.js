/**
 * ULTRA-PREMIUM AI AGENT (Universal API Edition)
 * =======================================================
 * File: AIAgent.js
 * 
 * Supports:
 * - Gemini Native (Google AI Studio)
 * - Universal APIs (OpenRouter, Groq, OpenAI)
 */

const AIAgent = {

  // ── Configuration Helpers ──
  getConfig: function() {
    var props = _getScriptProps();
    return {
      provider: props.getProperty('AI_PROVIDER') || 'gemini_native',
      endpoint: props.getProperty('AI_ENDPOINT') || 'https://openrouter.ai/api/v1/chat/completions',
      model: props.getProperty('AI_MODEL') || 'gemini-2.5-flash',
      apiKey: props.getProperty('GEMINI_API_KEY') || ''
    };
  },

  /**
   * Fetches current inventory dictionary for AI context.
   * Includes RAG Lite filtering based on raw user text to prevent payload limits.
   */
  getInventoryContext: function(rawText) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(InventoryCore.SHEET_INVENTORY) || ss.getActiveSheet();
    const lastRow = sheet.getLastRow();
    const headerRow = InventoryCore.findHeaderRow(sheet);
    const startRow = headerRow + 1;
    
    if (lastRow < startRow) return "[]";
    
    const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 3).getValues();
    
    // RAG Lite: Extract meaningful keywords (>2 chars) from user's message
    const rawLower = rawText ? String(rawText).toLowerCase() : "";
    const words = rawLower.match(/[a-z0-9]+/g) || [];
    const searchTerms = words.filter(w => w.length > 2);

    let items = [];
    dataRange.forEach(row => {
      if (row[1] !== "") {
        const branch = String(row[0]);
        const id = String(row[1]);
        const name = String(row[2]);
        
        if (!rawText || searchTerms.length === 0) {
          // Fallback if no text: load max 150 items to ensure safety
          if (items.length < 150) items.push({ branch, id, name });
          return;
        }

        const searchable = (branch + " " + id + " " + name).toLowerCase();
        let match = false;
        
        for (let w of searchTerms) {
          if (searchable.includes(w)) {
            match = true;
            break;
          }
        }
        
        if (match) {
          items.push({ branch, id, name });
        }
      }
    });
    
    // Return max 150 matched items to prevent Context Window overflow
    return JSON.stringify(items.slice(0, 150));
  },

  /**
   * Builds the system instruction prompt.
   */
  _buildSystemPrompt: function(rawText) {
    const inventoryContext = this.getInventoryContext(rawText);
    return `**Task:** You are an automated inventory assistant managing multiple branches. Your job is to read unstructured messages and extract inventory transaction data into JSON.

**CRITICAL CONSTRAINTS - VIOLATING THESE WILL CAUSE A SYSTEM CRASH:**
- DO NOT print conversational filler like "Here is the JSON", "User request:", or "Action:".
- DO NOT think step-by-step.
- YOU MUST ENCLOSE YOUR ENTIRE JSON OUTPUT WITHIN A MARKDOWN BLOCK (\`\`\`json ... \`\`\`).

**JSON Schema / Available Commands (cmd):**
1. UPDATE_CELL (memperbarui nilai sel yang ada)
2. APPEND_ROW (menambah baris data transaksi/inventory baru)
3. DELETE_ROW (menghapus baris data)
4. ASK_USER (jika ambigu)
5. UNKNOWN (jika barang tidak ditemukan)

**Rules:**
1. "branch" MUST be the name of the store/branch mentioned. If not mentioned, set it to "Pusat".
2. "type" MUST be exactly "IN" (restock, bought, received) or "OUT" (sold, sent, lost).
3. "quantity" MUST be a positive integer.
4. "itemId": Jika item tidak ditemukan dalam database, JANGAN asumsikan itu item baru kecuali terdapat kata kunci registrasi eksplisit (seperti /onboarding atau [NEW]). Jika tidak ada, kembalikan cmd "UNKNOWN" atau "NOT_FOUND". Jika ADA kata kunci /onboarding, gunakan cmd "APPEND_ROW" dan buat itemId sementara seperti "NEW_ITEM".
5. If the message mentions MULTIPLE items, output a JSON ARRAY of objects.
6. If the message mentions only ONE item, output a single JSON object.
7. "notes" should be a short summary of what happened.
8. DILARANG KERAS MENGUBAH FORMAT TAMPILAN SHEET (seperti row height, font, warna). HANYA GUNAKAN COMMAND UNTUK MENGUBAH DATA INVENTARIS/TRANSAKSI.
9. Jika permintaan sangat membingungkan atau ambigu, keluarkan: { "cmd": "ASK_USER", "question": "Tuliskan pertanyaan spesifik Anda dalam bahasa Indonesia di sini" }
10. KONTEKS PERCAKAPAN: Teks input mungkin berisi percakapan gabungan (contoh: teks asli + pertanyaan AI + jawaban singkat pengguna). Pahami seluruh riwayat tersebut untuk mengekstrak data utuh. Jika pengguna menjawab "Goreng", terapkan pada barang "Indomie Goreng" dari laporan awal.

**Inventory Context:**
${inventoryContext}

**Output Format Example (single item):**
{ "cmd": "APPEND_ROW", "branch": "string", "itemId": "string", "type": "IN" | "OUT", "quantity": number, "notes": "string", "customerEmail": "string|null" }

**Output Format Example (multiple items):**
[ { "cmd": "APPEND_ROW", "branch": "string", "itemId": "string", "type": "IN" | "OUT", "quantity": number, "notes": "string", "customerEmail": "string|null" } ]`;
  },

  /**
   * Main entry point: parse raw text into structured transaction(s).
   */
  parseTransaction: function(rawText) {
    const systemPrompt = this._buildSystemPrompt(rawText);
    
    // --- METODE BERFIKIR 2X (SMART LOOP QA) ---
    let maxAttempts = 3;
    let attempt = 1;
    let currentPrompt = rawText;
    let finalParsed = null;
    let success = false;
    
    while (attempt <= maxAttempts && !success) {
      const aiOutput = callAI(currentPrompt, systemPrompt);
      let parsed = this._extractJson(aiOutput);
      
      if (parsed) {
         // SELF VERIFICATION
         let verifySys = "Anda adalah auditor QA internal. Verifikasi apakah JSON ini sudah memenuhi instruksi pengguna dengan tepat. Jawab HANYA dengan kata 'SUDAH' jika benar. Jika ada kesalahan logika, salah qty, salah nama barang, atau klasifikasi IN/OUT yang salah, sebutkan detail kesalahannya agar AI utama bisa memperbaiki.";
         let verifyPrompt = "\nInstruksi Pengguna: \"" + rawText + "\"\n\nJSON yang dihasilkan:\n" + JSON.stringify(parsed) + "\n\nApakah JSON ini sudah tepat sasaran?";
         let verifyResult = callAI(verifyPrompt, verifySys);
         
         if (verifyResult.trim().toUpperCase().startsWith("SUDAH") || attempt === maxAttempts) {
             finalParsed = parsed;
             success = true;
             break;
         } else {
             // Berikan feedback ke currentPrompt
             currentPrompt = rawText + "\n\n[PERINGATAN DARI AUDITOR QA (Percobaan " + attempt + ")]:\n" + verifyResult + "\n\nTolong evaluasi kesalahan Anda dan perbaiki JSON Anda berdasarkan feedback di atas!";
             attempt++;
         }
      } else {
         currentPrompt = rawText + "\n\n[ERROR FORMATTING (Percobaan " + attempt + ")]:\nGagal membaca JSON. Pastikan output HANYA JSON block.";
         attempt++;
      }
    }
    return finalParsed;
  },

  _extractJson: function(aiOutput) {
    let text = aiOutput || "";
    let cleaned = text.trim();
    // 1. Try markdown block
    const mdRegex = /```(?:json)?\s*([\s\S]*?)\s*```/i;
    const match = cleaned.match(mdRegex);
    if (match && match[1]) {
      try { return JSON.parse(match[1]); } catch(e) {}
    }

    // 2. Try all combinations of [ ]
    let firstBracket = cleaned.indexOf('[');
    while (firstBracket !== -1) {
      let lastBracket = cleaned.lastIndexOf(']');
      while (lastBracket > firstBracket) {
        try {
          return JSON.parse(cleaned.substring(firstBracket, lastBracket + 1));
        } catch(e) {
          lastBracket = cleaned.lastIndexOf(']', lastBracket - 1);
        }
      }
      firstBracket = cleaned.indexOf('[', firstBracket + 1);
    }

    // 3. Try all combinations of { }
    let firstBrace = cleaned.indexOf('{');
    while (firstBrace !== -1) {
      let lastBrace = cleaned.lastIndexOf('}');
      while (lastBrace > firstBrace) {
        try {
          return JSON.parse(cleaned.substring(firstBrace, lastBrace + 1));
        } catch(e) {
          lastBrace = cleaned.lastIndexOf('}', lastBrace - 1);
        }
      }
      firstBrace = cleaned.indexOf('{', firstBrace + 1);
    }

    // 4. Try parsing the whole thing
    try {
      return JSON.parse(cleaned);
    } catch(e) {
      throw new Error("Gagal memproses output AI menjadi JSON. Raw output:\n" + text.substring(0, 500));
    }
  },

  /**
   * Parse natural language formatting request into JSON actions.
   */
  parseFormattingAI: function(rawText) {
    const config = this.getConfig();
    const systemPrompt = `**Task:** Translate the user's Google Sheets request into a JSON array of formatting commands.
**CRITICAL CONSTRAINTS - VIOLATING THESE WILL CAUSE A SYSTEM CRASH:**
- DO NOT print "User request:", "Interpretation:", "Action:", "Sheet:", or "Range:".
- DO NOT use bullet points or asterisks (*).
- DO NOT think step-by-step.
- YOU MUST ENCLOSE YOUR ENTIRE JSON OUTPUT WITHIN A MARKDOWN BLOCK (\`\`\`json ... \`\`\`).

**Available Commands (cmd):**
1. SET_BACKGROUND (butuh: range, color) -> color HEX format (#ffffff)
2. SET_FONT_COLOR (butuh: range, color) -> color HEX format
3. SET_FONT_FAMILY (butuh: range, fontFamily)
4. SET_FONT_SIZE (butuh: range, size) -> size angka
5. SET_TEXT_STYLE (butuh: range, isBold, isItalic, isStrikethrough, isUnderline) -> nilai boolean
6. SET_HORIZONTAL_ALIGNMENT (butuh: range, alignment) -> "left", "center", "right"
7. SET_ROW_HEIGHT (butuh: startRow, height, numRows)
8. AUTO_RESIZE_COLUMNS (butuh: startCol, numCols)
9. SET_FONT_WEIGHT (butuh: range, weight) -> "bold" atau "normal"
10. REPAIR_FORMULA (tidak butuh parameter) -> perbaiki #ERROR!
11. ASK_USER (butuh: question) -> jika ambigu, isi 'question' dengan pertanyaan spesifik dalam bahasa Indonesia
12. CLEAR_CONTENT (butuh: range) -> hapus teks tanpa hapus format
13. SET_BORDER (butuh: range, top, left, bottom, right, vertical, horizontal, color, style) -> boolean untuk sisi (true/false), color HEX, style = "SOLID", "SOLID_MEDIUM", "SOLID_THICK", "DASHED", "DOTTED", "DOUBLE"

**Rules:**
- Jika sebut "semua kolom ke 1", asumsikan "range": "A:A" atau "startCol": 1.
- Jika sebut huruf kolom ("kolom d"), format "range": "D:D".
- Warna basic = "#ffffff" (bg) atau "#000000" (teks).
- Jika user meminta "garis batas" atau "border tabel", wajib gunakan SET_BORDER.
- Jika error/perbaiki rumus = WAJIB REPAIR_FORMULA.
- Jika basic/kembali ke basic = WAJIB RESET_FORMAT.
- (Jika tidak disebut sheet, default "Inventory")

**Output Format Example:**
[
  { "sheet": "Inventory", "cmd": "CLEAR_CONTENT", "range": "H1:H20" }
]`;

    const aiOutput = callAI(rawText, systemPrompt);
    return this._extractJson(aiOutput);
  }
};
