/**
 * Dunia Simulasi: AI Smart Inventory System demo
 * Digunakan untuk menguji pustaka (Library) secara lokal tanpa harus terikat ke Klien.
 */

function runSimulationDemo() {
  Logger.log("=== MEMASUKI DUNIA SIMULASI ===");
  
  // 1. Injeksi Kredensial
  // Karena kita menjalankan dari dalam Library, _getScriptProps() secara otomatis
  // akan mengambil dari Script Properties milik proyek SmartInventoryCore ini.
  Logger.log("[Simulasi] Kredensial API akan diambil dari Script Properties lokal...");
  
  // 2. Monkey-Patching SpreadsheetApp
  // Menimpa fungsi getActiveSpreadsheet agar selalu mengembalikan dokumen klien Anda
  const TARGET_SHEET_ID = "1XasNdOu0HkmXR1_0_PcFEDC1TMW41R9L5jm2Eq0YZhQ";
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  
  try {
    SpreadsheetApp.getActiveSpreadsheet = function() {
      return SpreadsheetApp.openById(TARGET_SHEET_ID);
    };
    Logger.log(`[Simulasi] Spreadsheet dialihkan ke Klien Asli (ID: ${TARGET_SHEET_ID})`);
  } catch (e) {
    Logger.log("[Simulasi] Peringatan: Failed memodifikasi SpreadsheetApp (Native protection). Menggunakan fallback...");
  }
  
  // 3. Menjalankan Tes
  try {
    runAllTests();
  } catch(e) {
    Logger.log("ERROR DALAM SIMULASI: " + e.message);
  } finally {
    // Kembalikan ke keadaan semula
    try {
      SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    } catch(e) {}
    Logger.log("=== KELUAR DARI DUNIA SIMULASI ===");
  }
}
