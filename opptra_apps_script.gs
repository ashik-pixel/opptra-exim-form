/**
 * Opptra EXIM — Google Apps Script
 *
 * HOW TO USE:
 * 1. Open your Google Sheet → Extensions → Apps Script
 * 2. Delete all existing code and paste this entire file
 * 3. Change NOTIFY_EMAIL below to your actual EXIM team email
 * 4. Click Deploy → New Deployment → Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the URL and paste it in the setup banner on the form
 *
 * Every form submission will:
 * - Add a new row to "Sheet1" with all shipment data
 * - Auto-format headers (bold, navy background, frozen row)
 * - Send an email notification to the EXIM team
 */

const SHEET_NAME = "Sheet1";
const NOTIFY_EMAIL = "your-email@opptra.com"; // ← CHANGE THIS to EXIM team email

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Build header row on first submission
    if (sheet.getLastRow() === 0) {
      const headers = Object.keys(data);
      sheet.appendRow(headers);
      // Format header: bold, navy bg, white text, freeze
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#131A48")
        .setFontColor("#FFFFFF")
        .setFontFamily("Arial");
      sheet.setFrozenRows(1);
    }

    // Append data in correct column order
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = headers.map(h => data[h] || "");
    sheet.appendRow(row);

    // Auto-resize columns for readability
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }

    // Alternate row coloring for readability
    const lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, headers.length).setBackground("#F7F7F5");
    }

    // Email notification
    if (NOTIFY_EMAIL && NOTIFY_EMAIL !== "your-email@opptra.com") {
      const subject = "🚢 New EXIM Request: " + (data.ref_number || "N/A") + " — " + (data.brand || "");

      let body = "New shipment request submitted via Opptra EXIM form.\n";
      body += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n";
      body += "Reference: " + (data.ref_number || "N/A") + "\n";
      body += "Submitted: " + (data.submitted_at || "N/A") + "\n\n";

      body += "SHIPMENT DETAILS\n";
      body += "Brand: " + (data.brand || "") + "\n";
      body += "Readiness Date: " + (data.readiness_date || "") + "\n";
      body += "Mode: " + (data.mode || "") + "\n";
      body += "Incoterm: " + (data.incoterm || "") + "\n\n";

      body += "ROUTING\n";
      body += "Origin: " + (data.source_country || "") + " → " + (data.port_of_loading || "") + "\n";
      body += "Destination: " + (data.destination_country || "") + " → " + (data.port_of_discharge || "") + "\n";
      body += "Delivery ZIP: " + (data.final_delivery_zip || "") + "\n\n";

      body += "CARGO\n";
      body += "Gross: " + (data.gross_weight_kg || "") + " KG | Net: " + (data.net_weight_kg || "") + " KG | CBM: " + (data.cbm_m3 || "") + " M³\n";
      body += "Category: " + (data.product_category || "") + "\n";
      body += "DG: " + (data.dangerous_goods || "") + "\n";
      body += "MRP Labeling: " + (data.mrp_labeling || "") + "\n\n";

      body += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
      body += "View full details in Google Sheet.\n";

      MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", row: sheet.getLastRow() }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Health check endpoint — used by the form's "Save & Test" button
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ready", message: "Opptra EXIM endpoint active" }))
    .setMimeType(ContentService.MimeType.JSON);
}
