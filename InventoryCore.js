/**
 * Automated E-commerce Inventory System (PREMIUM / AI-RESISTANT)
 * =======================================================
 * File: InventoryCore.js
 * Description: Core business logic for inventory management.
 */

const InventoryCore = {
  
  // Sheet Names configuration
  SHEET_INVENTORY: 'Inventory', // Atau biarkan mencari nama sheet pertama
  LOW_STOCK_THRESHOLD: 5,
  
  /**
   * Helper: Find which row contains the headers dynamically.
   * Scans rows 1 to 10 for the word "Item ID" or "Item Name".
   */
  findHeaderRow: function(sheet) {
    const searchRange = sheet.getRange(1, 1, 10, 2).getValues(); // Check first 10 rows, columns A and B
    for (let i = 0; i < searchRange.length; i++) {
      const colA = String(searchRange[i][0]).toLowerCase().replace(/\s+/g, '');
      const colB = String(searchRange[i][1]).toLowerCase().replace(/\s+/g, '');
      if (colA.includes('itemid') || colB.includes('itemname')) {
        return i + 1; // Row numbers are 1-indexed
      }
    }
    return 1; // Default to row 1 if not found
  },

  /**
   * Generates a unique Item ID based on timestamp
   */
  generateItemId: function() {
    const date = new Date();
    const prefix = "ITM";
    const timestamp = date.getTime().toString().slice(-6);
    return `${prefix}-${timestamp}`;
  },

  /**
   * Inserts a new row into the Inventory sheet
   */
  insertNewItem: function(branchName, itemId, itemName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Use active sheet to be more flexible
    
    // Structure: [Branch, Item ID, Item Name, Category, Initial Stock, Stock In, Stock Out, Current Stock, Status]
    const initialRow = [
      branchName || 'Pusat',
      itemId,      
      itemName,    
      'General',   
      0,           
      0,           
      0,           
      '', // We will set formula manually
      'OK'         
    ];
    
    sheet.appendRow(initialRow);
    
    const lastRow = sheet.getLastRow();
    sheet.getRange(`H${lastRow}`).setFormula(`=E${lastRow}+F${lastRow}-G${lastRow}`);
    this.updateStatusColumn(sheet, lastRow);
  },

  /**
   * Recalculates the status column for a specific row
   */
  updateStatusColumn: function(sheet, rowIndex) {
    const currentStock = sheet.getRange(`H${rowIndex}`).getValue();
    const statusCell = sheet.getRange(`I${rowIndex}`);
    
    // Validasi apakah ini angka (bukan teks header)
    if (isNaN(currentStock) || currentStock === "") return;

    if (currentStock <= 0) {
      statusCell.setValue('OUT OF STOCK').setFontColor('red');
    } else if (currentStock <= this.LOW_STOCK_THRESHOLD) {
      statusCell.setValue('LOW STOCK').setFontColor('orange');
    } else {
      statusCell.setValue('OK').setFontColor('green');
    }
  },

  /**
   * Iterates through all items to update calculations and statuses
   */
  updateAllCurrentStock: function() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    const headerRow = this.findHeaderRow(sheet);
    const startRow = headerRow + 1;
    
    if (lastRow < startRow) return;
    
    const numRows = lastRow - startRow + 1;
    
    // Batch set formulas for column H
    const formulas = [];
    for (let i = startRow; i <= lastRow; i++) {
      formulas.push([`=E${i}+F${i}-G${i}`]);
    }
    sheet.getRange(startRow, 8, numRows, 1).setFormulas(formulas);
    
    SpreadsheetApp.flush(); // Force calculation before reading back values
    
    const currentStocks = sheet.getRange(startRow, 8, numRows, 1).getValues();
    const statusValues = [];
    const statusColors = [];
    
    for (let i = 0; i < numRows; i++) {
      const stock = currentStocks[i][0];
      if (isNaN(stock) || stock === "") {
        statusValues.push([sheet.getRange(startRow + i, 9).getValue()]); // Keep existing
        statusColors.push(["#000000"]);
      } else if (stock <= 0) {
        statusValues.push(["OUT OF STOCK"]);
        statusColors.push(["red"]);
      } else if (stock <= this.LOW_STOCK_THRESHOLD) {
        statusValues.push(["LOW STOCK"]);
        statusColors.push(["orange"]);
      } else {
        statusValues.push(["OK"]);
        statusColors.push(["green"]);
      }
    }
    
    const statusRange = sheet.getRange(startRow, 9, numRows, 1);
    statusRange.setValues(statusValues);
    statusRange.setFontColors(statusColors);
  },

  /**
   * Returns a list of items that are low on stock
   */
  getLowStockItems: function() {
    // Kita gunakan ActiveSheet agar bekerja meski nama sheet diganti oleh AI (seperti "Sheet1" -> "Inventory")
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const lastRow = sheet.getLastRow();
    const headerRow = this.findHeaderRow(sheet);
    const startRow = headerRow + 1;
    
    if (lastRow < startRow) return [];
    
    const numRowsToRead = lastRow - startRow + 1;
    const dataRange = sheet.getRange(startRow, 1, numRowsToRead, 9).getValues();
    const lowItems = [];
    
    dataRange.forEach(row => {
      const branch = row[0]; // Column A
      const id = row[1]; // Column B
      const name = row[2]; // Column C
      const stock = row[7]; // Column H
      
      // Pastikan bukan baris kosong dan merupakan angka
      if (id !== "" && !isNaN(stock) && stock > 0 && stock <= this.LOW_STOCK_THRESHOLD) {
        lowItems.push({ branch: branch, id: id, name: name, stock: stock });
      }
    });
    
    return lowItems;
  },

  /**
   * PREMIUM FEATURE: Logs a transaction and updates stock dynamically
   */
  logAndProcessTransaction: function(branchName, itemId, type, qty, note, userEmail) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const branch = branchName || 'Pusat';
    
    // 1. Ensure Log Sheet exists
    let logSheet = ss.getSheetByName('Transaction Logs');
    if (!logSheet) {
      logSheet = ss.insertSheet('Transaction Logs');
      logSheet.appendRow(['Timestamp', 'User Email', 'Branch', 'Item ID', 'Type', 'Quantity', 'Notes']);
      logSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#f3f3f3');
      logSheet.setFrozenRows(1);
    }
    
    // 2. Append Log
    const timestamp = new Date();
    logSheet.appendRow([timestamp, userEmail, branch, itemId, type, qty, note]);
    
    // 3. Update Main Inventory
    const invSheet = ss.getSheetByName(this.SHEET_INVENTORY) || ss.getActiveSheet();
    const lastRow = invSheet.getLastRow();
    const headerRow = this.findHeaderRow(invSheet);
    const startRow = headerRow + 1;
    
    if (lastRow < startRow) throw new Error("Inventory is empty.");
    
    const headers = invSheet.getRange(headerRow, 1, 1, invSheet.getLastColumn()).getValues()[0];
    let cmap;
    try {
      cmap = typeof _getInventoryColMap === 'function' ? _getInventoryColMap(headers) : { branch: 0, code: 1, initialStock: 4, stockIn: 5, stockOut: 6, stock: 7 };
    } catch(e) {
      cmap = { branch: 0, code: 1, initialStock: 4, stockIn: 5, stockOut: 6, stock: 7 };
    }

    const dataRange = invSheet.getRange(startRow, 1, lastRow - startRow + 1, invSheet.getLastColumn()).getValues();
    let targetRowIndex = -1;
    let currentIn = 0;
    let currentOut = 0;
    let currentStock = 0;
    
    for (let i = 0; i < dataRange.length; i++) {
      let bVal = cmap.branch !== -1 ? String(dataRange[i][cmap.branch]) : "";
      let cVal = cmap.code !== -1 ? String(dataRange[i][cmap.code]) : "";
      
      if (bVal.toLowerCase() === branch.toLowerCase() && cVal === String(itemId)) {
        targetRowIndex = startRow + i;
        currentIn = cmap.stockIn !== -1 ? (Number(dataRange[i][cmap.stockIn]) || 0) : 0;
        currentOut = cmap.stockOut !== -1 ? (Number(dataRange[i][cmap.stockOut]) || 0) : 0;
        currentStock = cmap.stock !== -1 ? (Number(dataRange[i][cmap.stock]) || 0) : 0;
        break;
      }
    }
    
    if (targetRowIndex === -1) {
      throw new Error(`Item ID "${itemId}" not found in branch "${branch}".`);
    }
    
    // ANOMALY VALIDATION: Prevent negative stock
    if (type === 'OUT' && qty > currentStock) {
      throw new Error(`InsufficientStockError: Cannot sell ${qty}. Only ${currentStock} available for "${itemId}" at "${branch}".`);
    }
    
    // 4. Increment the correct column
    if (type === 'IN' && cmap.stockIn !== -1) {
      invSheet.getRange(targetRowIndex, cmap.stockIn + 1).setValue(currentIn + qty);
    } else if (type === 'OUT' && cmap.stockOut !== -1) {
      invSheet.getRange(targetRowIndex, cmap.stockOut + 1).setValue(currentOut + qty);
    }
    
    // 5. Update Status
    if (cmap.stock !== -1) {
        const letterInitial = cmap.initialStock !== -1 ? String.fromCharCode(65 + cmap.initialStock) : "E";
        const letterIn = cmap.stockIn !== -1 ? String.fromCharCode(65 + cmap.stockIn) : "F";
        const letterOut = cmap.stockOut !== -1 ? String.fromCharCode(65 + cmap.stockOut) : "G";
        invSheet.getRange(targetRowIndex, cmap.stock + 1).setFormula(`=${letterInitial}${targetRowIndex}+${letterIn}${targetRowIndex}-${letterOut}${targetRowIndex}`);
    }
    this.updateStatusColumn(invSheet, targetRowIndex);
  }
};
