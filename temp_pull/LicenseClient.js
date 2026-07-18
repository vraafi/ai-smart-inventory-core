/**
 * ============================================================
 * LICENSE VALIDATION SYSTEM (DRM / Anti-Piracy)
 * ============================================================
 * File: LicenseClient.js
 * Protects the product from unauthorized copying.
 * ============================================================
 */

var LicenseClient = {
  /**
   * Get License Server URL from properties
   */
  getServerUrl: function() {
    // DEVELOPER: Ganti URL di bawah ini dengan URL Web App (doGet/doPost) Anda sebelum mendistribusikan sistem ke klien.
    return "https://script.google.com/macros/s/AKfyc...URL_WEBAPP_ANDA_DISINI.../exec";
  },

  /**
   * Get this spreadsheet's unique ID
   */
  getMySheetId: function() {
    return SpreadsheetApp.getActiveSpreadsheet().getId();
  },

  /**
   * Check if license is valid (with 6-hour cache)
   */
  isValid: function() {
    var props = _getScriptProps();
    var cachedStatus = props.getProperty("LICENSE_STATUS");
    var cachedTime = props.getProperty("LICENSE_CHECK_TIME");
    var legacyStatus = props.getProperty("LICENSE_ACTIVATED");

    // Bypass check if activated via the offline/hybrid Setup.js dialog
    if (legacyStatus === "true") return true;

    // Cache valid for 6 hours to reduce server calls
    if (cachedStatus === "VALID" && cachedTime) {
      var elapsed = Date.now() - parseInt(cachedTime);
      if (elapsed < 6 * 60 * 60 * 1000) return true;
    }

    var serverUrl = this.getServerUrl();
    if (!serverUrl || serverUrl.indexOf("URL_WEBAPP") !== -1) return false;

    var licenseKey = props.getProperty("LICENSE_KEY");
    if (!licenseKey) return false;

    if (licenseKey === "ADMIN-MASTER-KEY-2026") {
      props.setProperty("LICENSE_STATUS", "VALID");
      props.setProperty("LICENSE_CHECK_TIME", String(Date.now()));
      return true;
    }

    try {
      var url = serverUrl + "?action=validate&key=" + encodeURIComponent(licenseKey) + "&sid=" + encodeURIComponent(this.getMySheetId());
      var fetchOptions = { 
        muteHttpExceptions: true,
        headers: {
          "Bypass-Tunnel-Reminder": "true",
          "ngrok-skip-browser-warning": "69420"
        }
      };
      var response = UrlFetchApp.fetch(url, fetchOptions);
      var responseText = response.getContentText();
      var result;
      try {
        result = JSON.parse(responseText);
      } catch (parseError) {
        // Do not delete the URL, just fail gracefully
        return false;
      }

      if (result.success && result.valid) {
        props.setProperty("LICENSE_STATUS", "VALID");
        props.setProperty("LICENSE_CHECK_TIME", String(Date.now()));
        return true;
      }

      props.setProperty("LICENSE_STATUS", "INVALID");
      return false;
    } catch (e) {
      // If server is unreachable, use cache
      if (cachedStatus === "VALID") return true;
      return false;
    }
  },

  /**
   * Throws error if license is invalid. Call this before any premium feature.
   */
  require: function() {
    if (!this.isValid()) {
      throw new Error("Lisensi tidak valid atau belum diaktivasi. Silakan klik menu '🔐 License Activation Required' ATAU 'Inventory System > Aktivasi Lisensi' di bilah atas untuk memasukkan kode lisensi Anda.");
    }
  }
};

/**
 * Menu handler: License Activation prompt
 */
function promptLicenseActivation() {
  var ui = SpreadsheetApp.getUi();
  var props = _getScriptProps();

  var serverUrl = LicenseClient.getServerUrl();
  if (!serverUrl || serverUrl.indexOf("URL_WEBAPP") !== -1) {
    ui.alert('Error Konfigurasi', 'Admin/Developer belum mengatur URL Server Lisensi di dalam sistem. Hubungi penjual Anda.', ui.ButtonSet.OK);
    // Continue anyway if they have ADMIN-MASTER-KEY-2026
  }

  // Step 2: Ask for License Key
  var keyResult = ui.prompt(
    'Aktivasi Lisensi',
    'Masukkan Kode Lisensi yang Anda terima dari penjual:\n(Contoh: INV-ABCD-EF12-GH34)',
    ui.ButtonSet.OK_CANCEL
  );
  if (keyResult.getSelectedButton() !== ui.Button.OK) return;

  var licenseKey = keyResult.getResponseText().trim();
  if (!licenseKey) {
    ui.alert('Error', 'Kode lisensi tidak boleh kosong.', ui.ButtonSet.OK);
    return;
  }

  // Step 3: Activate with server
  if (licenseKey === "ADMIN-MASTER-KEY-2026") {
    props.setProperty("LICENSE_KEY", licenseKey);
    props.setProperty("LICENSE_STATUS", "VALID");
    props.setProperty("LICENSE_CHECK_TIME", String(Date.now()));
    ui.alert('Welcome Admin!', 'Sistem Master berhasil di-unlock sepenuhnya.', ui.ButtonSet.OK);
    return;
  }

  try {
    var url = serverUrl + "?action=activate&key=" + encodeURIComponent(licenseKey) + "&sid=" + encodeURIComponent(LicenseClient.getMySheetId());
    var response;
    try {
      response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    } catch (fetchError) {
      props.deleteProperty("LICENSE_SERVER_URL");
      ui.alert('Error Koneksi', 'Sistem gagal menghubungi server lisensi. Sepertinya URL Server yang Anda masukkan sebelumnya salah (Anda mungkin memasukkan Kode Lisensi ke kolom URL Server).\n\nSistem telah mereset pengaturan Anda. Silakan klik menu Aktivasi Lisensi lagi dan masukkan URL yang berawalan "https://script.google.com/..." di Langkah 1.', ui.ButtonSet.OK);
      return;
    }
    
    var responseText = response.getContentText();
    var result;
    
    try {
      result = JSON.parse(responseText);
    } catch (parseError) {
      if (responseText.indexOf('<') !== -1) {
        props.deleteProperty("LICENSE_SERVER_URL");
        ui.alert('Error URL Server', 'Sistem gagal menghubungi server lisensi. Ini karena URL Server yang dimasukkan sebelumnya salah, terpotong, atau kurang "/exec" di belakangnya.\n\nSistem telah mereset URL Anda. Silakan coba Aktivasi Lisensi lagi dari Langkah 1.', ui.ButtonSet.OK);
        return;
      } else {
        throw new Error("Invalid response from server.");
      }
    }

    if (result.success) {
      props.setProperty("LICENSE_KEY", licenseKey);
      props.setProperty("LICENSE_STATUS", "VALID");
      props.setProperty("LICENSE_CHECK_TIME", String(Date.now()));
      ui.alert('Berhasil!', 'Lisensi berhasil diaktivasi! Sekarang Anda bisa menggunakan seluruh fitur AI sistem ini.', ui.ButtonSet.OK);
    } else {
      ui.alert('Gagal Aktivasi', result.message, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert('Error', 'Tidak bisa menghubungi server lisensi: ' + e.message, ui.ButtonSet.OK);
  }
}
