/**
 * Automated E-commerce Inventory System (GOD-TIER / LEVEL 3)
 * =======================================================
 * File: InvoiceService.js
 * Description: Generates PDF invoices and sends them via Email.
 */

const InvoiceService = {

  /**
   * Generates a PDF invoice and sends it to the customer email.
   * @param {string} customerEmail
   * @param {object} txData - The transaction data containing itemId, quantity, etc.
   */
  sendInvoice: function(customerEmail, txData) {
    if (!customerEmail || customerEmail === 'null' || customerEmail === '') return;

    try {
      const date = new Date().toLocaleDateString();
      const invoiceNumber = "INV-" + new Date().getTime().toString().slice(-6);

      const htmlBody = `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #eee;">
          <h2 style="color: #333; text-align: center;">INVOICE / RECEIPT</h2>
          <hr style="border: 0; border-top: 1px solid #eee;" />
          <p><strong>Invoice #:</strong> ${invoiceNumber}</p>
          <p><strong>Date:</strong> ${date}</p>
          <p><strong>Bill To:</strong> ${customerEmail}</p>
          <br>
          <table style="width: 100%; border-collapse: collapse;">
            <thead>
              <tr style="background-color: #f8f9fa;">
                <th style="padding: 10px; border: 1px solid #ddd; text-align: left;">Item ID</th>
                <th style="padding: 10px; border: 1px solid #ddd; text-align: center;">Quantity</th>
                <th style="padding: 10px; border: 1px solid #ddd; text-align: right;">Total</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td style="padding: 10px; border: 1px solid #ddd;">${txData.itemId}</td>
                <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">${txData.quantity}</td>
                <td style="padding: 10px; border: 1px solid #ddd; text-align: right;">Paid</td>
              </tr>
            </tbody>
          </table>
          <br>
          <p style="text-align: center; color: #777; font-size: 12px;">Thank you for your business!</p>
        </div>
      `;

      // Create PDF blob
      const blob = Utilities.newBlob(htmlBody, MimeType.HTML)
          .setName("Invoice_" + invoiceNumber + ".pdf")
          .getAs(MimeType.PDF);

      // Send Email
      MailApp.sendEmail({
        to: customerEmail,
        subject: "Your Invoice: " + invoiceNumber,
        htmlBody: "Please find your invoice attached.",
        attachments: [blob]
      });

    } catch(e) {
      console.error("Failed to send invoice: " + e.message);
    }
  }
};
