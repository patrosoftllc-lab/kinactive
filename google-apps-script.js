// ============================================================
//  TOY SHOP — Google Apps Script Backend  (Code.gs)
//  Handles: toys list, orders list, place order,
//           mark fulfilled, toggle stock, email notification
// ============================================================

const ADMIN_EMAIL   = "your-sons-gmail@gmail.com"; // ← change this
const SHEET_TOYS    = "Toys";
const SHEET_ORDERS  = "Orders";

// ── Column positions in Orders sheet (1-indexed) ──────────
const O_COL = {
  TIMESTAMP:  1,
  BUYER_NAME: 2,
  BUYER_EMAIL:3,
  BUYER_PHONE:4,
  TOY_ID:     5,
  TOY_NAME:   6,
  PRICE:      7,
  NOTES:      8,
  STATUS:     9,
};

// ── Column positions in Toys sheet (1-indexed) ────────────
const T_COL = {
  NAME:       1,
  DESCRIPTION:2,
  PRICE:      3,
  CATEGORY:   4,
  IMAGE_URL:  5,
  IN_STOCK:   6,
};

// ============================================================
//  GET  →  ?action=toys  or  ?action=orders
// ============================================================
function doGet(e) {
  const action = (e.parameter.action || "toys").toLowerCase();
  if (action === "toys")   return jsonResp(getToys());
  if (action === "orders") return jsonResp(getOrders());
  return jsonResp({ error: "Unknown action" });
}

// ============================================================
//  POST  →  body JSON: { action, ...data }
//  Actions: order | fulfill | toggleStock
// ============================================================
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = (data.action || "order").toLowerCase();

    if (action === "order")       return jsonResp(placeOrder(data));
    if (action === "fulfill")     return jsonResp(markFulfilled(data.rowIndex));
    if (action === "togglestock") return jsonResp(toggleStock(data.toyId, data.instock));

    return jsonResp({ success: false, message: "Unknown action" });
  } catch (err) {
    return jsonResp({ success: false, message: err.toString() });
  }
}

// ============================================================
//  Read all toys from "Toys" sheet
// ============================================================
function getToys() {
  const sheet = getSheet(SHEET_TOYS);
  if (!sheet) return [];
  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0].map(h => h.toString().trim().toLowerCase().replace(/\s+/g,"_"));
  return rows.slice(1)
    .map((row, i) => {
      if (!row[0]) return null;
      const toy = { id: i + 2 }; // row number (2-based) as stable ID
      headers.forEach((h, idx) => toy[h] = row[idx]);
      return toy;
    })
    .filter(Boolean);
}

// ============================================================
//  Read all orders from "Orders" sheet
// ============================================================
function getOrders() {
  const sheet = getSheet(SHEET_ORDERS);
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];
  return rows.slice(1).map((row, i) => ({
    rowIndex:   i + 2,
    timestamp:  row[O_COL.TIMESTAMP - 1],
    buyerName:  row[O_COL.BUYER_NAME - 1],
    buyerEmail: row[O_COL.BUYER_EMAIL - 1],
    buyerPhone: row[O_COL.BUYER_PHONE - 1],
    toyId:      row[O_COL.TOY_ID - 1],
    toyName:    row[O_COL.TOY_NAME - 1],
    price:      row[O_COL.PRICE - 1],
    notes:      row[O_COL.NOTES - 1],
    status:     row[O_COL.STATUS - 1] || "Ordered",
  })).filter(o => o.buyerEmail); // skip blank rows
}

// ============================================================
//  Save new order
// ============================================================
function placeOrder(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName(SHEET_ORDERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ORDERS);
    sheet.appendRow(["Timestamp","Buyer Name","Buyer Email","Buyer Phone","Toy ID","Toy Name","Price","Notes","Status"]);
  }
  const row = [
    new Date().toISOString(),
    data.buyerName  || "",
    data.buyerEmail || "",
    data.buyerPhone || "",
    data.toyId      || "",
    data.toyName    || "",
    data.price      || "",
    data.notes      || "",
    "Ordered",
  ];
  sheet.appendRow(row);
  const rowIndex = sheet.getLastRow();

  // Email notification to admin
  try {
    MailApp.sendEmail(
      ADMIN_EMAIL,
      `🛒 New Order: ${data.toyName}`,
      `New order received!\n\nToy     : ${data.toyName}\nPrice   : ₹${data.price}\nBuyer   : ${data.buyerName}\nEmail   : ${data.buyerEmail}\nPhone   : ${data.buyerPhone || "—"}\nNotes   : ${data.notes || "—"}\nTime    : ${new Date().toLocaleString()}\n\nCheck the Orders tab in your Google Sheet.`
    );
  } catch(e) { /* email optional */ }

  return { success: true, rowIndex };
}

// ============================================================
//  Mark order as Fulfilled
// ============================================================
function markFulfilled(rowIndex) {
  const sheet = getSheet(SHEET_ORDERS);
  if (!sheet) return { success: false, message: "Orders sheet not found" };
  sheet.getRange(rowIndex, O_COL.STATUS).setValue("Fulfilled");

  // Get row data to send confirmation email to buyer
  try {
    const row       = sheet.getRange(rowIndex, 1, 1, 9).getValues()[0];
    const buyerEmail = row[O_COL.BUYER_EMAIL - 1];
    const toyName    = row[O_COL.TOY_NAME - 1];
    if (buyerEmail) {
      MailApp.sendEmail(
        buyerEmail,
        `✅ Your order for "${toyName}" is fulfilled!`,
        `Hi ${row[O_COL.BUYER_NAME - 1]},\n\nGreat news! Your order for "${toyName}" has been marked as fulfilled by Aryan.\n\nThank you for shopping! 🧸`
      );
    }
  } catch(e) { /* email optional */ }

  return { success: true };
}

// ============================================================
//  Toggle toy InStock status
// ============================================================
function toggleStock(toyId, newValue) {
  const sheet = getSheet(SHEET_TOYS);
  if (!sheet) return { success: false };
  // toyId = row number (2-based)
  sheet.getRange(Number(toyId), T_COL.IN_STOCK).setValue(newValue);
  return { success: true };
}

// ============================================================
//  Helper: get sheet by name
// ============================================================
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

// ============================================================
//  Helper: JSON response with CORS
// ============================================================
function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
//  SETUP INSTRUCTIONS
// ============================================================
//
//  STEP 1 — GOOGLE SHEET STRUCTURE
//  ----------------------------------
//  Tab name: Toys
//  Row 1 headers (exact):
//    Name | Description | Price | Category | ImageURL | InStock
//
//  InStock: Yes / No
//  ImageURL: public link (Google Drive "Anyone with link" or Imgur)
//
//  Tab "Orders" is created automatically when first order arrives.
//
//
//  STEP 2 — PASTE THIS SCRIPT
//  ----------------------------------
//  Google Sheet → Extensions → Apps Script
//  → Paste this entire file into Code.gs
//  → Set ADMIN_EMAIL to your son's Gmail
//  → Save (Ctrl+S)
//
//
//  STEP 3 — DEPLOY AS WEB APP
//  ----------------------------------
//  → Deploy → New deployment
//  → Type: Web app
//  → Execute as: Me
//  → Who has access: Anyone
//  → Deploy → Authorize → Copy the URL
//
//
//  STEP 4 — GOOGLE OAUTH (for login)
//  ----------------------------------
//  1. Go to https://console.cloud.google.com
//  2. New project → name it "Toy Shop"
//  3. APIs & Services → OAuth consent screen
//     → External → Fill app name, email → Save
//  4. Credentials → Create Credentials → OAuth 2.0 Client ID
//     → Application type: Web application
//     → Authorised JavaScript origins:
//        Add your deployed URL (e.g. https://yourname.netlify.app)
//        Also add http://localhost for local testing
//     → Create → Copy the Client ID
//
//
//  STEP 5 — UPDATE index.html
//  ----------------------------------
//  Find the CFG block in index.html:
//
//    const CFG = {
//      GOOGLE_CLIENT_ID: "YOUR_GOOGLE_CLIENT_ID...",  ← paste OAuth Client ID
//      APPS_SCRIPT_URL:  "YOUR_APPS_SCRIPT_URL_HERE", ← paste Apps Script URL
//      ADMIN_EMAIL:      "your-sons-gmail@gmail.com",  ← same as above
//    };
//
//
//  STEP 6 — DEPLOY index.html
//  ----------------------------------
//  Netlify Drop (easiest):
//    → Go to https://app.netlify.com/drop
//    → Drag index.html onto the page
//    → Done! Share the URL.
//
// ============================================================
