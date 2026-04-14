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
 * 5. Copy the URL — it is already hardcoded in the form, nothing else needed.
 *
 * Every form submission will:
 * - Add a new row to "Sheet1" with all shipment data
 * - Auto-format headers (bold, navy background, frozen row)
 * - Create a Google Drive folder named after the submission ref number
 * - Save all uploaded files (CI, PL, MSDS, UN38.3, IIP) into that folder
 * - Add a clickable Drive folder link to the sheet row
 * - Send an email notification to the EXIM team
 */

const SHEET_NAME   = "Sheet1";
const NOTIFY_EMAIL = "your-email@opptra.com"; // ← CHANGE THIS to your EXIM team email
const DRIVE_FOLDER_NAME = "Opptra EXIM Submissions"; // Root folder in My Drive

// ─── doPost: receives form submission ──────────────────────────────────────
function doPost(e) {
  try {
    const raw  = e.postData.contents;
    const data = JSON.parse(raw);

    Logger.log('Received payload keys: ' + Object.keys(data).join(', '));
    Logger.log('Payload size (chars): ' + raw.length);

    // ── 1. Save files to Google Drive ─────────────────────────────────────
    let folderUrl = '';
    const filesPayload = data.files || {};
    const hasFiles = Object.values(filesPayload).some(arr => arr && arr.length > 0);

    Logger.log('hasFiles: ' + hasFiles);
    Logger.log('files keys: ' + Object.keys(filesPayload).join(', '));

    if (hasFiles) {
      // Get or create the root submissions folder
      let rootFolder;
      const existing = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
      rootFolder = existing.hasNext() ? existing.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);

      // Create a sub-folder for this specific submission
      const subFolder = rootFolder.createFolder(data.ref_number || ('SUB-' + Date.now()));
      folderUrl = subFolder.getUrl();

      const typeLabels = { ci: 'CI', pl: 'PL', msds: 'MSDS', un38: 'UN38.3', iip: 'IIP' };

      Object.entries(filesPayload).forEach(([type, fileList]) => {
        if (!fileList || fileList.length === 0) return;
        fileList.forEach(function(f) {
          try {
            const decoded = Utilities.base64Decode(f.data);
            const blob    = Utilities.newBlob(decoded, f.mimeType || 'application/octet-stream', f.name);
            subFolder.createFile(blob);
          } catch (fileErr) {
            Logger.log('File save error (' + type + '): ' + fileErr.toString());
          }
        });
      });
    }

    // Remove raw file data before saving to sheet (keep only the folder link)
    delete data.files;
    data.files_folder = folderUrl || 'No files uploaded';

    // ── 2. Save row to Google Sheet ───────────────────────────────────────
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      // First submission — write headers
      const headers = Object.keys(data);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#131A48")
        .setFontColor("#FFFFFF")
        .setFontFamily("Arial");
      sheet.setFrozenRows(1);
    }

    // Append data in the same column order as headers
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row     = headers.map(h => data[h] !== undefined ? data[h] : '');
    sheet.appendRow(row);

    // Make the Drive folder cell a clickable hyperlink
    if (folderUrl) {
      const lastRow   = sheet.getLastRow();
      const folderCol = headers.indexOf('files_folder') + 1;
      if (folderCol > 0) {
        sheet.getRange(lastRow, folderCol)
          .setFormula('=HYPERLINK("' + folderUrl + '","📁 Open Files")');
      }
    }

    // Auto-resize columns and alternate row color
    const lastRow = sheet.getLastRow();
    for (let i = 1; i <= headers.length; i++) sheet.autoResizeColumn(i);
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, headers.length).setBackground("#F7F7F5");
    }

    // ── 3. Email notification ─────────────────────────────────────────────
    if (NOTIFY_EMAIL && NOTIFY_EMAIL !== "your-email@opptra.com") {
      const subject = "🚢 New EXIM Request: " + (data.ref_number || "N/A") + " — " + (data.brand || "");

      let body = "New shipment request submitted via Opptra EXIM form.\n";
      body += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n";
      body += "Reference:  " + (data.ref_number || "N/A") + "\n";
      body += "Submitted:  " + (data.submitted_at || "N/A") + "\n\n";

      body += "SHIPMENT DETAILS\n";
      body += "Brand:           " + (data.brand || "") + "\n";
      body += "Readiness Date:  " + (data.readiness_date || "") + "\n";
      body += "Mode:            " + (data.mode || "") + "\n";
      body += "Incoterm:        " + (data.incoterm || "") + "\n\n";

      body += "ROUTING\n";
      body += "Origin:      " + (data.source_country || "") + " → " + (data.port_of_loading || "") + "\n";
      body += "Destination: " + (data.destination_country || "") + " → " + (data.port_of_discharge || "") + "\n";
      body += "Delivery ZIP: " + (data.final_delivery_zip || "") + "\n\n";

      body += "CARGO\n";
      body += "Gross: " + (data.gross_weight_kg || "") + " KG  |  ";
      body += "Net: "   + (data.net_weight_kg  || "") + " KG  |  ";
      body += "CBM: "   + (data.cbm_m3         || "") + " M³\n";
      body += "Category: " + (data.product_category || "") + "\n";
      body += "DG: "        + (data.dangerous_goods  || "") + "\n";
      body += "MRP Labeling: " + (data.mrp_labeling  || "") + "\n\n";

      if (folderUrl) {
        body += "FILES\n";
        body += "Drive folder: " + folderUrl + "\n\n";
      }

      body += "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n";
      body += "View full details in Google Sheet.\n";

      MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", row: sheet.getLastRow(), folder: folderUrl }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── doGet: health check ───────────────────────────────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ready", message: "Opptra EXIM endpoint active" }))
    .setMimeType(ContentService.MimeType.JSON);
}
