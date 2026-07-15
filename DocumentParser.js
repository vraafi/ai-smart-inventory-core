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
    if (name.toLowerCase().endsWith(".csv") || name.toLowerCase().endsWith(".cvs") || mime.includes("csv") || mime === "text/comma-separated-values") {
      const csvData = Utilities.parseCsv(blob.getDataAsString());
      text = _convert2DArrayToText(csvData);
    } else if (name.toLowerCase().endsWith(".xlsx") || mime.includes("spreadsheetml") || mime.includes("excel")) {
      text = _readExcelViaDrive(blob);
    } else if (name.toLowerCase().endsWith(".pdf") || mime.includes("image") || mime.includes("pdf") || name.toLowerCase().endsWith(".png") || name.toLowerCase().endsWith(".jpg") || name.toLowerCase().endsWith(".jpeg")) {
      text = _readImageOrPdfViaDriveOcr(blob);
    } else {
      throw new Error("Unsupported file type: " + name + " (" + mime + "). Only CSV, XLSX, PDF, and Images are supported.");
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

function _getRawDataArrayFromBlob(blob) {
  const mime = blob.getContentType();
  const name = blob.getName().toLowerCase();
  
  if (name.endsWith(".csv") || name.endsWith(".cvs") || mime.includes("csv") || mime === "text/comma-separated-values") {
    try {
      return Utilities.parseCsv(blob.getDataAsString());
    } catch(e) {
      return null;
    }
  } else if (name.endsWith(".xlsx") || mime.includes("spreadsheetml") || mime.includes("excel")) {
    const resource = {
      title: "TEMP_EXCEL_PARSE_" + Date.now(),
      mimeType: MimeType.GOOGLE_SHEETS
    };
    let file;
    try {
      file = Drive.Files.insert(resource, blob, {convert: true});
      const ss = SpreadsheetApp.openById(file.id);
      const sheet = ss.getSheets()[0];
      return sheet.getDataRange().getDisplayValues();
    } catch(e) {
      return null;
    } finally {
      if (file) {
        try { Drive.Files.trash(file.id); } catch(e) {}
      }
    }
  }
  return null;
}

function _getRawDataArrayFromUrl(url) {
  try {
    const match = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) return [];
    const sheetId = match[1];
    
    // Attempt to open. Will fail if private/no permission
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0]; // Just read the first sheet for bulk import
    const data = sheet.getDataRange().getDisplayValues();
    return data;
  } catch (e) {
    if (e.message.toLowerCase().includes("permission")) {
       throw new Error("Akses ditolak (Permission Denied). Mohon ubah akses link Google Sheets Anda menjadi 'Siapa saja yang memiliki link dapat melihat' (Anyone with the link can view).");
    }
    throw new Error("Gagal mengekstrak link: " + e.message);
  }
}

function _extractTextFromUrl(url) {
  try {
    const match = url.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) return "";
    const sheetId = match[1];
    
    // Attempt to open. Will fail if private/no permission
    const ss = SpreadsheetApp.openById(sheetId);
    const sheets = ss.getSheets();
    
    let text = "";
    // Read up to first 3 sheets to prevent massive payload timeouts
    for (let i = 0; i < Math.min(sheets.length, 3); i++) {
      const sheet = sheets[i];
      const data = sheet.getDataRange().getDisplayValues();
      if (data.length > 0) {
        text += `\n--- Lembar: ${sheet.getName()} ---\n`;
        text += _convert2DArrayToText(data);
      }
    }
    return text.trim();
  } catch (e) {
    if (e.message.toLowerCase().includes("permission")) {
       throw new Error("Akses ditolak (Permission Denied). Mohon ubah akses link Google Sheets Anda menjadi 'Siapa saja yang memiliki link dapat melihat' (Anyone with the link can view).");
    }
    throw new Error("Gagal mengekstrak link: " + e.message);
  }
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

function _readImageOrPdfViaDriveOcr(blob) {
  let mime = blob.getContentType() || "";
  const name = (blob.getName() || "").toLowerCase();
  
  if (mime === "" || mime === "application/octet-stream") {
      if (name.endsWith(".pdf")) mime = "application/pdf";
      else if (name.endsWith(".png")) mime = "image/png";
      else if (name.endsWith(".jpg") || name.endsWith(".jpeg")) mime = "image/jpeg";
      
      blob.setContentType(mime); // Force the blob to have the correct source mime type
  }

  const resource = {
    title: "TEMP_OCR_PARSE_" + Date.now(),
    mimeType: mime // Use the source mime type. Drive will convert it to Docs because of {ocr: true}
  };
  
  // Insert the file into Drive with OCR enabled
  const file = Drive.Files.insert(resource, blob, { ocr: true });
  
  let text = "";
  try {
    const doc = DocumentApp.openById(file.id);
    text = doc.getBody().getText();
  } catch (e) {
    Logger.log("OCR Error: " + e.message);
    throw new Error("Gagal membaca gambar/PDF (OCR Error): " + e.message);
  } finally {
    // Always delete the temporary file!
    try {
      Drive.Files.trash(file.id);
    } catch(e) {}
  }
  
  // Limit output if it's too massive
  if (text.length > 8000) {
    text = text.substring(0, 8000) + "\n...[TRUNCATED DUE TO SIZE]";
  }
  return text.trim();
}
