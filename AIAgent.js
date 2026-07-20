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
1. UPDATE_CELL (update existing cell value)
2. APPEND_ROW (add new transaction/inventory row)
3. DELETE_ROW (delete a row)
4. ASK_USER (if ambiguous)
5. UNKNOWN (if item Not found)

**Rules:**
1. "branch" MUST be the name of the store/branch mentioned. If not mentioned, set it to "Head Office".
2. "type" MUST be exactly "IN" (restock, bought, received) or "OUT" (sold, sent, lost).
3. "quantity" MUST be a positive integer.
4. "itemId": If the item is Not found in the database, DO NOT assume it's a new item unless there is an explicit registration keyword (like /onboarding or [NEW]). If none, return cmd "UNKNOWN" or "NOT_FOUND". If there IS an /onboarding keyword, use cmd "APPEND_ROW" and make a temporary itemId like "NEW_ITEM".
5. If the message mentions MULTIPLE items, output a JSON ARRAY of objects.
6. If the message mentions only ONE item, output a single JSON object.
7. "notes" should be a short summary of what happened.
8. STRICTLY PROHIBITED TO CHANGE SHEET DISPLAY FORMATTING (like row height, font, color). ONLY USE COMMANDS TO CHANGE INVENTORY/TRANSACTION DATA.
9. If the request is highly confusing or ambiguous, output: { "cmd": "ASK_USER", "question": "Write your specific question in English here" }
10. CONVERSATION CONTEXT: The input text might contain a combined conversation (Example: original text + AI question + user short answer). Understand the entire history to extract the full data. If the user answers "Fried", apply it to the item "Indomie Fried" from the initial report.

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
         let verifySys = "You are an internal QA auditor. Verify if this JSON precisely fulfills the user's instructions. Answer ONLY with the word 'DONE' if correct. If there are logic errors, wrong qty, wrong Item Name, or wrong IN/OUT classification, state the error details so the main AI can fix it.";
         let verifyPrompt = "\nUser Instructions: \"" + rawText + "\"\n\nGenerated JSON:\n" + JSON.stringify(parsed) + "\n\nIs this JSON accurate?";
         let verifyResult = callAI(verifyPrompt, verifySys);
         
         if (verifyResult.trim().toUpperCase().startsWith("DONE") || attempt === maxAttempts) {
             finalParsed = parsed;
             success = true;
             break;
         } else {
             // Give feedback to currentPrompt
             currentPrompt = rawText + "\n\n[WARNING FROM QA AUDITOR (Attempt " + attempt + ")]:\n" + verifyResult + "\n\nPlease evaluate your mistake and fix your JSON based on the feedback above!";
             attempt++;
         }
      } else {
         currentPrompt = rawText + "\n\n[FORMATTING ERROR (Attempt " + attempt + ")]:\nFailed reading JSON. Ensure output is ONLY a JSON block.";
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
      throw new Error("Failed memproses output AI menjadi JSON. Raw output:\n" + text.substring(0, 500));
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
1. SET_BACKGROUND (needs: range, color) -> color HEX format (#ffffff)
2. SET_FONT_COLOR (needs: range, color) -> color HEX format
3. SET_FONT_FAMILY (needs: range, fontFamily)
4. SET_FONT_SIZE (needs: range, size) -> size is number
5. SET_TEXT_STYLE (needs: range, isBold, isItalic, isStrikethrough, isUnderline) -> boolean values
6. SET_HORIZONTAL_ALIGNMENT (needs: range, alignment) -> "left", "center", "right"
7. SET_ROW_HEIGHT (needs: startRow, height, numRows)
8. AUTO_RESIZE_COLUMNS (needs: startCol, numCols)
9. SET_FONT_WEIGHT (needs: range, weight) -> "bold" or "normal"
10. REPAIR_FORMULA (no params) -> fix #ERROR!
11. ASK_USER (needs: question) -> if ambiguous, fill 'question' with a specific English question
12. CLEAR_CONTENT (needs: range) -> clear text without clearing format
13. SET_BORDER (needs: range, top, left, bottom, right, vertical, horizontal, color, style) -> boolean for sides (true/false), color HEX, style = "SOLID", "SOLID_MEDIUM", "SOLID_THICK", "DASHED", "DOTTED", "DOUBLE"

**Rules:**
- If mention "all column 1", assume "range": "A:A" or "startCol": 1.
- If mention column letter ("column d"), format "range": "D:D".
- Basic colors = "#ffffff" (bg) or "#000000" (text).
- If user asks for "table border" or "boundaries", MUST use SET_BORDER.
- If fix formula error = MUST REPAIR_FORMULA.
- If basic/return to basic = MUST RESET_FORMAT.
- (If no sheet is mentioned, default "Inventory")

**Output Format Example:**
[
  { "sheet": "Inventory", "cmd": "CLEAR_CONTENT", "range": "H1:H20" }
]`;

    const aiOutput = callAI(rawText, systemPrompt);
    return this._extractJson(aiOutput);
  }
};
