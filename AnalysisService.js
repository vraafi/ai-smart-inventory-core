/**
 * Automated E-commerce Inventory System (GOD-TIER / LEVEL 3)
 * =======================================================
 * File: AnalysisService.js
 * Description: AI-powered predictive analysis and monthly reporting.
 */

const AnalysisService = {

  /**
   * Generates a monthly performance report using AI and sends it via Telegram.
   */
  generateMonthlyReport: function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Transaction Logs');
    if (!logSheet) {
      throw new Error("Sheet 'Transaction Logs' not found. No data to analyze.");
    }

    const lastRow = logSheet.getLastRow();
    if (lastRow <= 1) {
      throw new Error("No transaction data available for analysis.");
    }

    // 1. Gather this month's transactions
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    const allData = logSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    let monthlyData = [];
    let totalIn = 0;
    let totalOut = 0;
    let branchSummary = {};

    allData.forEach(row => {
      const timestamp = new Date(row[0]);
      if (timestamp >= monthStart) {
        const branch = String(row[2]) || 'Pusat';
        const itemId = String(row[3]);
        const type = String(row[4]);
        const qty = Number(row[5]) || 0;

        monthlyData.push({ date: timestamp, branch: branch, itemId: itemId, type: type, qty: qty });

        if (type === 'IN') totalIn += qty;
        if (type === 'OUT') totalOut += qty;

        if (!branchSummary[branch]) {
          branchSummary[branch] = { in: 0, out: 0, items: {} };
        }
        if (!branchSummary[branch].items[itemId]) {
          branchSummary[branch].items[itemId] = { in: 0, out: 0 };
        }
        
        if (type === 'IN') {
          branchSummary[branch].in += qty;
          branchSummary[branch].items[itemId].in += qty;
        }
        if (type === 'OUT') {
          branchSummary[branch].out += qty;
          branchSummary[branch].items[itemId].out += qty;
        }
      }
    });

    if (monthlyData.length === 0) {
      throw new Error("No transactions found for this month.");
    }

    // 2. Build a summary string for the AI
    let dataSummary = `Monthly Transaction Summary (${now.toLocaleString('default', { month: 'long', year: 'numeric' })}):\n`;
    dataSummary += `Total Transactions: ${monthlyData.length}\n`;
    dataSummary += `Total Stock IN: ${totalIn} units\n`;
    dataSummary += `Total Stock OUT: ${totalOut} units\n\n`;
    dataSummary += `Per-Branch Breakdown:\n`;

    for (const branch in branchSummary) {
      dataSummary += `\n[Branch: ${branch}]\n`;
      dataSummary += `- IN: ${branchSummary[branch].in}, OUT: ${branchSummary[branch].out}\n`;
      dataSummary += `- Items:\n`;
      for (const id in branchSummary[branch].items) {
        dataSummary += `  * ${id}: IN=${branchSummary[branch].items[id].in}, OUT=${branchSummary[branch].items[id].out}\n`;
      }
    }

    // 3. Get current stock levels
    const invSheet = ss.getSheetByName(InventoryCore.SHEET_INVENTORY) || ss.getActiveSheet();
    const invLastRow = invSheet.getLastRow();
    const headerRow = InventoryCore.findHeaderRow(invSheet);
    const startRow = headerRow + 1;

    if (invLastRow >= startRow) {
      const invData = invSheet.getRange(startRow, 1, invLastRow - startRow + 1, 9).getValues();
      dataSummary += `\nCurrent Stock Levels (Across Branches):\n`;
      invData.forEach(row => {
        if (row[1] !== "") {
          dataSummary += `- [${row[0]}] ${row[1]} (${row[2]}): ${row[7]} units [Status: ${row[8]}]\n`;
        }
      });
    }

    // 4. Send to AI for analysis
    const config = AIAgent.getConfig();
    if (!config.apiKey) {
      throw new Error("API Key belum dikonfigurasi.");
    }

    const analysisPrompt = `You are a business analyst AI managing a multi-branch inventory system. Analyze the following inventory data and provide:
1. Top selling items and best performing branches.
2. Items that need restocking soon (predict based on sales velocity) per branch.
3. A brief performance summary (good/bad trends) across all branches.
4. Specific actionable recommendations for the business owner.

Keep your response concise (under 200 words), in Bahasa Indonesia, and use emoji for readability.

Data:
${dataSummary}`;

    let aiResponse;
    if (config.provider === 'openrouter') {
      aiResponse = this._callAI_OpenRouter(config, analysisPrompt);
    } else {
      aiResponse = this._callAI_Gemini(config, analysisPrompt);
    }

    // 5. Send to Telegram
    let telegramMsg = `📊 <b>LAPORAN ANALISIS BULANAN</b> 📊\n`;
    telegramMsg += `<i>${now.toLocaleString('default', { month: 'long', year: 'numeric' })}</i>\n\n`;
    telegramMsg += `📈 Total Transaksi: <b>${monthlyData.length}</b>\n`;
    telegramMsg += `🟢 Total Masuk: <b>${totalIn}</b> unit\n`;
    telegramMsg += `🔴 Total Keluar: <b>${totalOut}</b> unit\n\n`;
    telegramMsg += `<b>🤖 Analisis AI:</b>\n${aiResponse}\n\n`;
    telegramMsg += `<i>Generated by AI Inventory System</i>`;

    TelegramService.sendMessage(telegramMsg);
    return { success: true, report: aiResponse };
  },

  // Helper: Call OpenRouter/Groq for analysis
  _callAI_OpenRouter: function(config, prompt) {
    const payload = {
      model: config.model,
      messages: [
        { role: "user", content: prompt }
      ],
      temperature: 0.7
    };
    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "Authorization": "Bearer " + config.apiKey,
        "HTTP-Referer": "https://script.google.com",
        "X-Title": "GAS Inventory Analysis"
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(config.endpoint, options);
    const json = JSON.parse(response.getContentText());
    return json.choices[0].message.content;
  },

  // Helper: Call Gemini Native for analysis
  _callAI_Gemini: function(config, prompt) {
    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7 }
    };
    const modelName = config.model || "gemini-2.5-flash";
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${config.apiKey}`;
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(apiUrl, options);
    const json = JSON.parse(response.getContentText());
    return json.candidates[0].content.parts[0].text;
  }
};
