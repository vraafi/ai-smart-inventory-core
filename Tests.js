/**
 * Automated Test Suite for AI Smart Inventory
 * Run 'runAllTests' from the Google Apps Script Editor to verify all features.
 */

function runAllTests() {
  let uiOutput = "=== STARTING AUTOMATED TESTS ===\n\n";
  Logger.log("=== STARTING AUTOMATED TESTS ===");
  const testChatId = "TEST_USER_001"; // Or use your own TG ID to receive real messages
  const testName = "AutoTester";
  
  const scenarios = [
    {
      name: "1. AMBIGUOUS MATCH TEST",
      text: "Laku 1 unit Galaxy S24 Ultra Titanium Black",
      expectedBehavior: "Should ask for clarification (256GB or 512GB)"
    },
    {
      name: "2. EXACT MATCH (OUT) TEST",
      text: "Laku 1 unit Galaxy S24 Ultra Titanium Black 12GB/256GB",
      expectedBehavior: "Should deduct 1 stock from row 2"
    },
    {
      name: "3. EXACT MATCH (IN) TEST",
      text: "Barang datang 5 unit Galaxy S24 Ultra Titanium Black 12GB/256GB",
      expectedBehavior: "Should add 5 stock to row 2"
    },
    {
      name: "4. ONBOARDING (NEW ITEM) TEST",
      text: "/onboarding Tolong tambahkan produk baru iPhone 15 Pro Max 256GB warna Natural Titanium harga 25000000",
      expectedBehavior: "Should create a new row in Inventory with stock 0"
    },
    {
      name: "5. WIPE COMMAND TEST",
      text: "/wipe tolong hapus semua isi tabel cabang",
      expectedBehavior: "Should generate 4 options (a,b,c,d) for wiping data"
    }
  ];

  for (const s of scenarios) {
    uiOutput += `--- Running: ${s.name} ---\nInput: ${s.text}\nExpected: ${s.expectedBehavior}\n`;
    Logger.log("\\n--- Running: " + s.name + " ---");
    Logger.log("Input: " + s.text);
    Logger.log("Expected: " + s.expectedBehavior);
    try {
      _processWithAI(testChatId, s.text, testName, "Telegram");
      uiOutput += "Status: Sent to AI successfully\n\n";
    } catch (e) {
      uiOutput += `TEST FAILED WITH ERROR: ${e.message}\n\n`;
      Logger.log("TEST FAILED WITH ERROR: " + e.message);
    }
    // Small delay to prevent rate limits
    Utilities.sleep(2000); 
  }
  
  uiOutput += "=== TESTS FINISHED ===\nPlease check your Telegram/Email/Logs to verify the Agent's responses!";
  Logger.log("\\n=== TESTS FINISHED ===");
  Logger.log("Please check your Telegram/Email to verify the Agent's responses!");
  
  try {
    SpreadsheetApp.getUi().alert("Automated Tests Result", uiOutput, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}
