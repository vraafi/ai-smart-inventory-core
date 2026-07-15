/**
 * Automated E-commerce Inventory System (STANDARD PACKAGE)
 * =======================================================
 * File: Telegram.js
 * Description: Handles Telegram API integration and credentials UI.
 */

const TelegramService = {

  /**
   * Send a message via Telegram Bot API
   * @param {string} messageText
   */
  sendMessage: function(messageText) {
    const props = _getScriptProps();
    const botToken = props.getProperty('TELEGRAM_BOT_TOKEN');
    const chatId = props.getProperty('TELEGRAM_CHAT_ID');

    if (!botToken || !chatId) {
      throw new Error("Telegram credentials are not set. Please set them via the menu first.");
    }

    const url = `https://api.telegram.org/bot${botToken}/sendMessage`;
    const payload = {
      chat_id: chatId,
      text: messageText,
      parse_mode: 'HTML'
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (!result.ok) {
      throw new Error(`Telegram API Error: ${result.description}`);
    }

    return result;
  },

  /**
   * UI Dialog to set credentials
   */
  promptCredentials: function() {
    const ui = SpreadsheetApp.getUi();
    const props = _getScriptProps();

    const currentToken = props.getProperty('TELEGRAM_BOT_TOKEN') || '';
    const currentChatId = props.getProperty('TELEGRAM_CHAT_ID') || '';

    const tokenResponse = ui.prompt(
      'Telegram Settings (1/2)',
      `Enter Bot Token (Current: ${currentToken ? '***' + currentToken.slice(-5) : 'None'}):`,
      ui.ButtonSet.OK_CANCEL
    );

    if (tokenResponse.getSelectedButton() !== ui.Button.OK) return;

    const chatResponse = ui.prompt(
      'Telegram Settings (2/2)',
      `Enter Chat ID (Current: ${currentChatId || 'None'}):`,
      ui.ButtonSet.OK_CANCEL
    );

    if (chatResponse.getSelectedButton() !== ui.Button.OK) return;

    const tokenInput = tokenResponse.getResponseText().trim();
    const chatInput = chatResponse.getResponseText().trim();

    if (tokenInput) props.setProperty('TELEGRAM_BOT_TOKEN', tokenInput);
    if (chatInput) props.setProperty('TELEGRAM_CHAT_ID', chatInput);

    ui.alert('Success', 'Telegram credentials saved successfully!', ui.ButtonSet.OK);
  }
};

// Wrapper function for menu call
function promptTelegramCredentials() {
  TelegramService.promptCredentials();
}
