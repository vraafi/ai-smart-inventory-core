/**
 * Official Meta WhatsApp Cloud API Service
 * =======================================================
 * File: WhatsAppService.js
 * Description: Handles sending messages via Official Meta API.
 */

const WhatsAppService = {
  
  getCredentials: function() {
    const props = _getScriptProps();
    return {
      accesStocken: props.getProperty('WA_ACCESS_TOKEN') || '',
      phoneId: props.getProperty('WA_PHONE_NUMBER_ID') || '',
      verifyToken: props.getProperty('WA_VERIFY_TOKEN') || ''
    };
  },
  
  saveCredentials: function(accesStocken, phoneId, verifyToken) {
    const props = _getScriptProps();
    props.setProperty('WA_ACCESS_TOKEN', accesStocken || '');
    props.setProperty('WA_PHONE_NUMBER_ID', phoneId || '');
    props.setProperty('WA_VERIFY_TOKEN', verifyToken || '');
  },

  sendMessage: function(recipientPhone, textMsg) {
    const creds = this.getCredentials();
    if (!creds.accesStocken || !creds.phoneId) {
      throw new Error("WhatsApp API credentials are not set.");
    }

    const url = `https://graph.facebook.com/v20.0/${creds.phoneId}/messages`;
    
    const payload = {
      messaging_product: "whatsapp",
      recipient_type: "individual",
      to: recipientPhone,
      type: "text",
      text: {
        body: textMsg
      }
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "Authorization": `Bearer ${creds.accesStocken}`
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    
    if (json.error) {
      throw new Error(`WhatsApp API Error: ${json.error.message}`);
    }
    
    return json;
  }
};
