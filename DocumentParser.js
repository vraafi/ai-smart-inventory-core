// ============================================================
//  AI SMART INVENTORY — DOCUMENT PARSER MODULE
//  Extracts text data from CSV and XLSX files sent via Telegram/Email
// ============================================================

/**
 * Extracts raw text from a CSV or XLSX Blob
 * @param {GoogleAppsScript.Base.Blob} blob
 * @returns {string} The extracted text
 */
function _extractTextFromBlob(blob) {
  const mime = blob.getContentType() || "";
  const name = blob.getName() || "";
  let text = "";

  try {
    if (name.toLowerCase().endsWith(".csv") || mime.includes("csv") || mime === "text/comma-separated-values") {
      const csvData = Utilities.parseCsv(blob.getDataAsString());
      text = _convert2DArrayToText(csvData);
    } else if (name.toLowerCase().endsWith(".xlsx") || mime.includes("spreadsheetml") || mime.includes("excel")) {
      text = _readExcelViaDrive(blob);
    } else {
      throw new Error("Unsupported file type: " + name + " (" + mime + "). Only CSV and XLSX are supported.");
    }
  } catch (e) {
    Logger.log("DocumentParser error: " + e.message);
    throw new Error("Gagal membaca dokumen: " + e.message);
  }
  
  return text;
}

function _readExcelViaDrive(blob) {
  // Convert the Excel Blob to a Google Sheet
  const resource = {
    title: "TEMP_EXCEL_PARSE_" + Date.now(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  
  // Insert the file, converting it
  const file = Drive.Files.insert(resource, blob, {convert: true});
  
  let text = "";
  try {
    const ss = SpreadsheetApp.openById(file.id);
    const sheets = ss.getSheets();
    
    // Read up to first 3 sheets to avoid massive payload
    for (let i = 0; i < Math.min(sheets.length, 3); i++) {
      const sheet = sheets[i];
      const data = sheet.getDataRange().getValues();
      if (data.length > 0) {
        text += `\n--- Lembar: ${sheet.getName()} ---\n`;
        text += _convert2DArrayToText(data);
      }
    }
  } finally {
    // Always delete the temporary file!
    try {
      Drive.Files.trash(file.id);
    } catch(e) {}
  }
  
  return text.trim();
}

function _convert2DArrayToText(data) {
  let output = "";
  for (let r = 0; r < data.length; r++) {
    // skip completely empty rows
    const rowStr = data[r].join("").trim();
    if (rowStr === "") continue;
    
    output += data[r].map(cell => {
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return String(cell).replace(/\n/g, " ");
    }).join(" | ") + "\n";
  }
  // limit output to max 8000 chars to avoid hitting AI token limits
  if (output.length > 8000) {
    output = output.substring(0, 8000) + "\n...[TRUNCATED DUE TO SIZE]";
  }
  return output;
}
