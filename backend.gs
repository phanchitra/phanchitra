/************************************************************
 * Intersys Stock Control — Google Apps Script Backend (Web App)
 * Works with your App.jsx (actions: selfTest, bootstrap, listEntries,
 * appendStructured, rebuildTotals, uploadImage, DO/RQ save+generate).
 *
 * Tabs ensured:
 * README
 * DO_Header, DO_Items, RQ_Header, RQ_Items, Settings
 * Products, Total Stock, Stock (All), Stock Out, Receipt Log
 ************************************************************/

/* ===== Global: Stock Image Request Form folder ===== */
const IMAGE_FOLDER_ID = "1hZlYsP8Zky7NEcdgqfeelLuVZedkh_w0"; // <— verify this ID

/* ===== DO CONFIG ===== */
const DO_CFG = {
  SHEET_ID: "1e3PryQ0UjyU0Qx0pDYxN0pCUnFVNEM2_n8HxqgtqA8o",
  DOC_TEMPLATE_ID: "1VhHDMgdrgtOgiiHk8iVKy5fm0mEPKSvrJyYKRUUuhhc",
  DOC_FOLDER_ID: "1TWKGVpka7Cq2h1P2ohD86cVAPnYjV7oY",
  PDF_FOLDER_ID: "1QtDFstt1t61658c2Yu2wgA3hbZ6clSs9",
  PREFIX: "DO",
  COUNTER_PAD: 4,
  TZ: "Asia/Phnom_Penh",
  DATE_FMT: "dd/MM/yyyy",
};

const DO_SH = { HEADER: "DO_Header", ITEMS: "DO_Items" };
const DO_H_HEADERS = [
  "DO_No","DO_Date","COMPANY_NAME","COMPANY_ADDR","COMPANY_REFE","PROJECT","PO_REF","NOTE",
  "TOTAL_QTY","Created_TS","Status","Doc_Link","PDF_Link","Output_Folder"
];
const DO_I_HEADERS = ["DO_No","Line_No","ITEM_NAME","ITEM_DESC","ITEM_CODE","UNIT","QTY","REMARK","ITEM_ID"];

/* ===== RQ CONFIG ===== */
const RQ_CFG = {
  SHEET_ID: "1e3PryQ0UjyU0Qx0pDYxN0pCUnFVNEM2_n8HxqgtqA8o",
  DOC_TEMPLATE_ID: "1WA9SCtkRrvfXUUo3HGwXvRywqSWqS5daNt-5NtOw-QY",
  DOC_FOLDER_ID: "1ZSfdXvJ8atanHXdBwaR33bF1AB5G0J7_",
  PDF_FOLDER_ID: "1D9Ku46EnbPt17ndVbPCwG-9qiPgAnna3",
  PREFIX: "RQ",
  FY: "", // '' = auto (yy)
  COUNTER_PAD: 4,
  TZ: "Asia/Phnom_Penh",
  DATE_FMT: "dd/MM/yyyy",
};
const RQ_SH = { HEADER: "RQ_Header", ITEMS: "RQ_Items", SETTINGS: "Settings" };

/* ===== Stock sheet names & headers ===== */
const STOCK_SHEETS = {
  IN: {
    name: "Stock (All)",
    header: [
      "Date LPO","Type","LPO No","SO","Project","Category","Supplier",
      "Product Name","Description","Part No","Unit","Qty","Unit Price","Amount","Currency"
    ],
  },

  // Type-specific inputs (NO "Type" column)
  IN_PO: {
    name: "PO",
    header: [
      "Date LPO","LPO No","SO","Project","Category","Supplier",
      "Product Name","Description","Part No","Unit","Qty","Unit Price","Amount","Currency"
    ],
  },
  IN_LOCAL: {
    name: "Local",
    header: [
      "Date LPO","LPO No","SO","Project","Category","Supplier",
      "Product Name","Description","Part No","Unit","Qty","Unit Price","Amount","Currency"
    ],
  },
  IN_NONPO: {
    name: "NON PO",
    header: [
      "Date LPO","LPO No","SO","Project","Category","Supplier",
      "Product Name","Description","Part No","Unit","Qty","Unit Price","Amount","Currency"
    ],
  },
  OUT: {
    name: "Stock Out",
    header: [
      "No. Ref (DO No)",
      "Item",
      "SKU / Barcode",
      "Description",
      "Quantity Out (pcs)",
      "Stock Out Date (Date)",      // <— restored to keep qty from being interpreted as a date
      "Remarks",
      "Company",
      "Address",
      "Project",
      "Ref",
      "Note",
      "P.O Reference"
    ],
  },
  RECEIPT: {
    name: "Receipt Log",
    header: [
      "Timestamp","Item Code","Transaction Type","Date","SO","Project","Ref.Quot/PO",
      "Request Type","Description","Qty","Return Date (Date)","Reason","Remark",
      "Company Name","Image Upload"
    ], // 15 columns total
  },
};

const SHEET_PRODUCTS = "Products";
const SHEET_TOTAL = "Total Stock";
const SHEET_README = "README";
// ---- Sizing helpers for Doc images ----
const CM_TO_PT = 28.3464567;                 // 1 cm in points
const RQ_IMAGE_BOX_CM = 1.85;                // desired image width (cm)
const RQ_IMAGE_BOX_PT = Math.round(RQ_IMAGE_BOX_CM * CM_TO_PT); // ≈ 52 pt

const MAX_BYTES = 8 * 1024 * 1024; // 8MB image upload

// =====================================================================
// =========================== HELPERS =================================
function _normalizeFlow_(s){
  s = String(s || "").toUpperCase().trim();
  if (/^IN[-_]?ADV$/.test(s)) return "IN";   // IN_ADV / IN-ADV → IN
  if (s === "IN_PO" || s === "IN_LOCAL" || s === "IN_NONPO") return s;
  return s;
}
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _fmt_(d, tz, fmt) { return Utilities.formatDate(d, tz, fmt); }
function _openSS_(id) { return SpreadsheetApp.openById(id); }
function _normCurrency_(s){
  const t = String(s || "").trim().toUpperCase();
  if (t === "KHR" || t === "RIEL" || t === "៛") return "KHR";
  if (t === "USD" || t === "$") return "USD";
  return t || "USD";
}

function _ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (headers && sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}
function _replaceAll_(body, map) {
  Object.keys(map).forEach(k => body.replaceText("{{" + k + "}}", map[k] ?? ""));
}
/** Build/refresh "Total Request" sheet from Receipt Log (Transaction Type = "Request")
 * Columns: [Part No / SKU, Description, Total Request Qty]
 */
function buildRequestSummary_(ss, opts) {
  opts = opts || {};
  const receiptSheet = opts.receiptSheet || STOCK_SHEETS.RECEIPT.name;
  const outName = opts.outName || "Total Request";

  const recSh = ss.getSheetByName(receiptSheet);
  const data = recSh ? recSh.getDataRange().getValues() : [];
  if (!data || data.length <= 1) {
    const sh = _ensureSheet_(ss, outName, ["Part No / SKU","Description","Total Request Qty"]);
    // clear old rows (keep header)
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow()-1, sh.getMaxColumns()).clearContent();
    return { ok: true, updated: 0, sheet: outName };
  }

  const header = data[0] || [];
  const rows = data.slice(1);

  const idxKey   = findHeaderIndex(header, [/^item\s*code$/i, /sku|barcode/i, /^part\s*no$/i]);
  const idxDesc  = findHeaderIndex(header, [/^description$/i]);
  const idxQty   = findHeaderIndex(header, [/^qty$/i, /^quantity$/i]);
  const idxType  = findHeaderIndex(header, [/^transaction\s*type$/i, /^type$/i]);

  const agg = {};    // key => { qty: number, desc: string }
  rows.forEach(r => {
    const t = (idxType >= 0 ? String(r[idxType] || "") : "").trim().toLowerCase();
    if (t !== "request") return;

    const key  = (idxKey  >= 0 ? String(r[idxKey]  || "") : "").trim();
    if (!key) return;

    const qty  = Number(idxQty >= 0 ? r[idxQty] : 0) || 0;
    const desc = (idxDesc >= 0 ? String(r[idxDesc] || "") : "").trim();

    if (!agg[key]) agg[key] = { qty: 0, desc: "" };
    agg[key].qty += qty;
    // prefer a non-empty description; keep the first seen non-empty
    if (desc && !agg[key].desc) agg[key].desc = desc;
  });

  const outHeader = ["Part No / SKU","Description","Total Request Qty"];
  const outSh = _ensureSheet_(ss, outName, outHeader);
  if (outSh.getLastRow() > 1) {
    outSh.getRange(2, 1, outSh.getLastRow()-1, outSh.getMaxColumns()).clearContent();
  }

  const keys = Object.keys(agg).sort();
  const outRows = keys.map(k => [k, agg[k].desc || "", agg[k].qty]);

  if (outRows.length) {
    outSh.getRange(2, 1, outRows.length, outHeader.length).setValues(outRows);
  }

  return { ok: true, updated: outRows.length, sheet: outName };
}

/** Open by full URL or raw ID (accepts /d/<ID>/ patterns) */
function openSheetByIdOrUrl(idOrUrl) {
  if (!idOrUrl) throw new Error("Missing sheetId (empty).");
  const s = String(idOrUrl).trim();
  try {
    if (/^https?:\/\//i.test(s)) return SpreadsheetApp.openByUrl(s);
    const m = s.match(/\/d\/([a-zA-Z0-9-_]{20,})/);
    if (m && m[1]) return SpreadsheetApp.openById(m[1]);
    const id = (s.match(/[a-zA-Z0-9-_]{20,}/) || [s])[0];
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error("openSheetByIdOrUrl failed: " + e.message + " | input=" + s);
  }
}

/** Dates */
function toISOOrBlank(value, opts) {
  opts = opts || {};
  const defaultNow = !!opts.defaultNow;
  if (!value) return defaultNow ? new Date().toISOString() : "";
  try {
    const s = String(value).trim();
    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s; // already ISO
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return new Date(s + "T00:00:00").toISOString();
    const d = (Object.prototype.toString.call(value) === "[object Date]") ? value : new Date(s);
    if (isNaN(d.getTime())) return defaultNow ? new Date().toISOString() : "";
    return d.toISOString();
  } catch (e) {
    return defaultNow ? new Date().toISOString() : "";
  }
}

/** Parse POST body safely */
function _parseBody_(e) {
  try {
    if (e && e.postData && e.postData.contents) return JSON.parse(e.postData.contents);
  } catch (_) {}
  return {};
}

/** Unified action detector */
function _getAction_(e, body) {
  return (e && e.parameter && e.parameter.action)
    ? e.parameter.action
    : (body && body.action) ? body.action : "";
}

function findHeaderIndex(header, regexList) {
  for (let i = 0; i < header.length; i++) {
    const h = String(header[i] || "").trim();
    for (let r = 0; r < regexList.length; r++) {
      if (regexList[r].test(h)) return i;
    }
  }
  return -1;
}

/** ===== Diagnostics: check IDs & write access ===== */
function diagCheck_() {
  const out = { ok: true, checks: [] };
  function add(name, fn) {
    try { const v = fn(); out.checks.push({ name, ok: true, detail: v }); }
    catch (e) { out.ok = false; out.checks.push({ name, ok: false, error: e.message || String(e) }); }
  }

  add("IMAGE_FOLDER_ID", () => DriveApp.getFolderById(IMAGE_FOLDER_ID).getName());
  add("RQ DOC_TEMPLATE_ID", () => DriveApp.getFileById(RQ_CFG.DOC_TEMPLATE_ID).getName());
  add("RQ DOC_FOLDER_ID", () => DriveApp.getFolderById(RQ_CFG.DOC_FOLDER_ID).getName());
  add("RQ PDF_FOLDER_ID", () => DriveApp.getFolderById(RQ_CFG.PDF_FOLDER_ID).getName());
  add("Sheet open", () => SpreadsheetApp.openById(RQ_CFG.SHEET_ID).getName());

  // Quick write test in DOC folder (then trash)
  add("Write test to DOC_FOLDER_ID", () => {
    const f = DriveApp.getFolderById(RQ_CFG.DOC_FOLDER_ID)
      .createFile("diag.txt", "ok:" + new Date().toISOString());
    const id = f.getId();
    f.setTrashed(true);
    return "created+trashed: " + id;
  });

  return out;
}

/** Return {header, rows, objects[]} for a sheet */
function readSheetAsObjects_(sh) {
  if (!sh) return { header: [], rows: [], objects: [] };
  var values = sh.getDataRange().getValues();
  if (!values || values.length === 0) return { header: [], rows: [], objects: [] };

  var header = values[0].map(function (h) { return String(h || ""); });
  var rows = values.slice(1);
  var objects = rows.map(function (r) {
    var o = {};
    for (var i = 0; i < header.length; i++) o[header[i]] = r[i];
    return o;
  });
  return { header: header, rows: rows, objects: objects };
}

/** Case-insensitive object key getter */
function getCI_(obj, key) {
  var k = String(key);
  var low = k.toLowerCase();
  for (var p in obj) if (Object.prototype.hasOwnProperty.call(obj, p) && String(p).toLowerCase() === low) return obj[p];
  return undefined;
}

/** Filter list of objects by substring query (all fields joined) */
function filterByQuery_(objs, q) {
  if (!q) return objs;
  var t = String(q).toLowerCase().trim();
  if (!t) return objs;
  return objs.filter(function (o) {
    return Object.keys(o).some(function (k) {
      var v = o[k];
      return v != null && String(v).toLowerCase().indexOf(t) !== -1;
    });
  });
}

/** Limit + reverse (newest first) like your listEntries */
function takeLimitNewest_(objs, limit) {
  var lim = Math.min(Number(limit || 50), 5000);
  return objs.slice().reverse().slice(0, lim);
}

// =====================================================================
// ========================= BOOTSTRAP =================================
// =====================================================================

/** Ensure ALL required tabs exist with headers (including README) */
function handleBootstrap(data) {
  try {
    data = data || {};
    if (!data.sheetId) data.sheetId = DO_CFG.SHEET_ID;

    const ss = openSheetByIdOrUrl(data.sheetId);

    _ensureSheet_(ss, SHEET_README, ["Welcome", "Notes"]);
    const readme = ss.getSheetByName(SHEET_README);
    if (readme.getLastRow() < 2) {
      readme.appendRow([
        "Intersys Simple Backend",
        "Use the web app endpoints to append Stock (All)/Out/Receipt, and use rebuildTotals to refresh Total Stock."
      ]);
    }

    // Stock tabs
    _ensureSheet_(ss, STOCK_SHEETS.IN.name,       STOCK_SHEETS.IN.header);
    _ensureSheet_(ss, STOCK_SHEETS.IN_PO.name,    STOCK_SHEETS.IN_PO.header);
    _ensureSheet_(ss, STOCK_SHEETS.IN_LOCAL.name, STOCK_SHEETS.IN_LOCAL.header);
    _ensureSheet_(ss, STOCK_SHEETS.IN_NONPO.name, STOCK_SHEETS.IN_NONPO.header);
    _ensureSheet_(ss, STOCK_SHEETS.OUT.name,      STOCK_SHEETS.OUT.header);
    _ensureSheet_(ss, STOCK_SHEETS.RECEIPT.name,  STOCK_SHEETS.RECEIPT.header);

    // Apply number/date formats
    ensureStockOutFormatting_(ss);
    ensureStockAllFormatting_(ss);
    ensureStockTypedFormatting_(ss);
    // Products & Totals (aligned header)
    _ensureSheet_(ss, SHEET_PRODUCTS, ["SKU / Barcode","Product Name","Barcode","Category"]);
    _ensureSheet_(ss, SHEET_TOTAL, ["SKU / Barcode","Product Name","Total In","Total Out","Balance"]);

    // DO / RQ sheets
    _ensureSheet_(ss, DO_SH.HEADER, DO_H_HEADERS);
    _ensureSheet_(ss, DO_SH.ITEMS,  DO_I_HEADERS);
    _ensureSheet_(ss, RQ_SH.HEADER, [
      "RQ_No","RQ_Date","Company_Name","SO","Project","Ref_QPO","Types","Note",
      "TOTAL_QTY","Created_TS","Status","Doc_Link","PDF_Link","Output_Folder"
    ]);
    _ensureSheet_(ss, RQ_SH.ITEMS, [
      "No. Ref (RQ No)","Line_No","Description","Item_Code","Unit","Qty","Return_Date","Reason","Image","Remark"
    ]);
    _ensureSheet_(ss, RQ_SH.SETTINGS, ["Counter","FY"]);

    return jsonOut({ ok: true, message: "Bootstrap done" });
  } catch (err) {
    return jsonOut({ ok: false, error: "bootstrap failed: " + (err && err.message ? err.message : err) });
  }
}
/** Ensure the Stock Out sheet has expected header & number/date formats */
function ensureStockOutFormatting_(ss) {
  const expectedHeader = STOCK_SHEETS.OUT.header;
  const shOut = _ensureSheet_(ss, STOCK_SHEETS.OUT.name, expectedHeader);

  // Column 5 = Quantity Out (pcs) -> integer
  formatColumnAsNumber_(shOut, 5, "0");

  // Column 6 = Stock Out Date -> yyyy-mm-dd hh:mm
  const lastRow = Math.max(shOut.getLastRow(), 2);
  if (lastRow > 1) {
    shOut.getRange(2, 6, lastRow - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm");
  }
}



// Formats the consolidated "Stock (All)" sheet
function ensureStockAllFormatting_(ss) {
  const shAll = ss.getSheetByName(STOCK_SHEETS.IN.name);
  if (!shAll) return;
  const lastRow = Math.max(shAll.getLastRow(), 2);
  if (lastRow > 1) {
    shAll.getRange(2, 1, lastRow - 1, 1).setNumberFormat("yyyy-mm-dd");
    shAll.getRange(2, 12, lastRow - 1, 1).setNumberFormat("0");
    shAll.getRange(2, 13, lastRow - 1, 1).setNumberFormat("#,##0.00");
    shAll.getRange(2, 14, lastRow - 1, 1).setNumberFormat("#,##0.00");
  }
}

// Formats the typed Stock-In sheets: PO, Local, NON PO
function ensureStockTypedFormatting_(ss) {
  // Shared layout:
  //  1:Date LPO, 11:Qty, 12:Unit Price, 13:Amount, 14:Currency (text)
  [STOCK_SHEETS.IN_PO.name, STOCK_SHEETS.IN_LOCAL.name, STOCK_SHEETS.IN_NONPO.name].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const lastRow = Math.max(sh.getLastRow(), 2);
    if (lastRow > 1) {
      sh.getRange(2, 1,  lastRow - 1, 1).setNumberFormat("yyyy-mm-dd");
      sh.getRange(2, 11, lastRow - 1, 1).setNumberFormat("0");          // Qty
      sh.getRange(2, 12, lastRow - 1, 1).setNumberFormat("#,##0.00");   // Unit Price
      sh.getRange(2, 13, lastRow - 1, 1).setNumberFormat("#,##0.00");   // Amount
      // Currency (col 14) stays as text
    }
  });
}


/** Format a whole used column (from row 2 down) as Number/plain text */
function formatColumnAsNumber_(sh, colIndex, pattern) {
  const lastRow = Math.max(sh.getLastRow(), 2);
  if (lastRow < 2) return;
  sh.getRange(2, colIndex, lastRow - 1, 1).setNumberFormat(pattern || "0");
}

// =====================================================================
// ========================= STOCK SECTION ==============================
// =====================================================================

/** Primary append used by React UI (IN / OUT / RECEIPT) */
function handleAppendStructured(data) {
  data = data || {};
  if (!data.sheetId) return jsonOut({ ok: false, error: "Missing sheetId" });

  const ss = openSheetByIdOrUrl(data.sheetId);
  const flow = _normalizeFlow_(data.flow || "");
  const d = data.data || {};

  if (flow === "IN" || flow === "IN_PO" || flow === "IN_LOCAL" || flow === "IN_NONPO") {
    // Normalize type from UI: "PO" | "Local" | "NON PO"
    const typeRaw = (d.type || d.Type || "").toString().trim();
    const normType = (/^non\s*po$/i.test(typeRaw)) ? "NON PO"
                  : (/^local$/i.test(typeRaw))     ? "Local"
                  : "PO"; // default

    // If flow explicitly picked a variant, force it
    const type = (flow === "IN_PO")     ? "PO"
               : (flow === "IN_LOCAL")  ? "Local"
               : (flow === "IN_NONPO")  ? "NON PO"
               : normType;

    // Dates: accept dateLPO, datePO, or stockInDate; fall back to now
    const dateVal = d.dateLPO || d.datePO || d.stockInDate;
    const dateObj = dateVal ? new Date(dateVal) : new Date();
    // Numbers and money (safe coercions)
    const _qty       = Math.max(0, Number(d.qty ?? d.quantity ?? 0) || 0);
    const _unitPrice = Number(d.unitPrice ?? d.price ?? 0) || 0;
    const _amount    = Number(d.amount ?? (_qty * _unitPrice)) || 0;
    const _currency  = _normCurrency_(d.currency ?? d.curr ?? "");

    // Row for typed sheet (no "Type")
    const rowTyped = [
      dateObj,                         // Date LPO
      d.lpoNo || d.poNo || "",         // LPO No
      d.so || d.soNo || "",            // SO
      d.project || d.projectName || "",// Project
      d.category || "",                // Category
      d.supplier || "",                // Supplier
      d.productName || "",             // Product Name
      d.description || "",             // Description
      d.partNo || d.itemCode || "",    // Part No
      d.unit || "",                    // Unit
      _qty,                            // Qty
      _unitPrice,                      // Unit Price
      _amount,                         // Amount
      _currency                        // Currency
    ];
    // Consolidated row (insert "Type" after Date)
    const rowAll = rowTyped.slice();
    rowAll.splice(1, 0, type);

    const shAll   = _ensureSheet_(ss, STOCK_SHEETS.IN.name, STOCK_SHEETS.IN.header);
    shAll.appendRow(rowAll);

    let target = STOCK_SHEETS.IN_PO;
    if (type === "Local")   target = STOCK_SHEETS.IN_LOCAL;
    else if (type === "NON PO") target = STOCK_SHEETS.IN_NONPO;

    const shTyped = _ensureSheet_(ss, target.name, target.header);
    shTyped.appendRow(rowTyped);

    return jsonOut({ ok: true, to: [STOCK_SHEETS.IN.name, target.name], typeUsed: type });
  }

  if (flow === "OUT") {
    const shOut = _ensureSheet_(ss, STOCK_SHEETS.OUT.name, STOCK_SHEETS.OUT.header);

    // Quantity: force number >= 0 (integer)
    const qtyNum = Math.max(0, Number(d.qty ?? d.quantity ?? 0) || 0);

    shOut.appendRow([
      d.doNo || d.DO_NO || "",     // No. Ref (DO No)
      d.lineNo || d.itemNo || d.line || "", // Item
      d.sku || d.barcode || "",    // SKU / Barcode
      d.description || "",         // Description
      qtyNum,                      // Quantity Out (pcs) -> numeric
      new Date(),                  // Stock Out Date (Date)
      d.remarks || "",             // Remarks
      d.company || "",             // Company
      d.address || "",             // Address
      d.project || "",             // Project
      d.ref || "",                 // Ref
      d.note || "",                // Note
      d.poRef || d.po || ""        // P.O Reference
    ]);

    ensureStockOutFormatting_(ss);
    rebuildTotalsSafe_(data.sheetId);
    return jsonOut({ ok: true, to: STOCK_SHEETS.OUT.name });
  }
  if (flow === "RECEIPT") {
    const shRec = _ensureSheet_(ss, STOCK_SHEETS.RECEIPT.name, STOCK_SHEETS.RECEIPT.header);
    shRec.appendRow([
      toISOOrBlank(d.timestamp, { defaultNow: true }),    // Timestamp
      d.itemCode || d.sku || "",                          // Item Code
      d.transactionType || "Request",                     // Transaction Type
      toISOOrBlank(d.date, { defaultNow: true }),         // Date
      d.so || "",                                         // SO
      d.project || "",                                    // Project
      d.refQuotPo || d.ref || "",                         // Ref.Quot/PO
      d.requestType || "",                                // Request Type
      d.description || "",                                // Description
      Number(d.qty || 0) || 0,                            // Qty
      toISOOrBlank(d.returnDate, { defaultNow: false }),  // Return Date (Date)
      d.reason || "",                                     // Reason
      d.remark || "",                                     // Remark
      d.companyName || "",                                // Company Name
      ensureDriveImageUrl_(d.imageUrl || d.image || "", d.rqNo || d.requestNo || "", (d.itemCode || d.sku || "image") + ".png")
    ]);
    return jsonOut({ ok: true, to: STOCK_SHEETS.RECEIPT.name });
  }
  rebuildTotalsSafe_(data.sheetId);
  return jsonOut({ ok: true, to: STOCK_SHEETS.RECEIPT.name });
}


/** Reader for any sheet tab (with query + limit) */
function handleListEntries(data) {
  data = data || {};
  if (!data.sheetId) return jsonOut({ ok: false, error: "Missing sheetId" });
  if (!data.sheetName) return jsonOut({ ok: false, error: "Missing sheetName" });

  const ss = openSheetByIdOrUrl(data.sheetId);
  const sh = ss.getSheetByName(data.sheetName);
  if (!sh) return jsonOut({ ok: false, error: "Sheet not found: " + data.sheetName });

  let values = sh.getDataRange().getValues();
  if (!values || values.length === 0) return jsonOut({ ok: true, header: [], rows: [] });

  const header = values[0];
  let rows = values.slice(1).reverse();

  // Convert Date objects to ISO strings
  rows = rows.map(r => r.map(cell => (
    Object.prototype.toString.call(cell) === "[object Date]" ? new Date(cell).toISOString() : cell
  )));

  const q = (data.query || "").toString().trim().toLowerCase();
  if (q) rows = rows.filter(r => r.join(" ").toLowerCase().indexOf(q) !== -1);

  const limit = Math.min(Number(data.limit || 50), 5000);
  rows = rows.slice(0, limit);

  return jsonOut({ ok: true, header, rows, count: rows.length });
}

/** Build/refresh "Total Stock" from "Stock (All)" + "Stock Out" */
function rebuildTotalsSafe_(sheetId) {
  try { handleRebuildTotals({ sheetId: sheetId || DO_CFG.SHEET_ID }); }
  catch (e) { console.warn("rebuildTotalsSafe_:", e && e.message ? e.message : e); }
}

function handleRebuildTotals(data) {
  data = data || {};
  if (!data.sheetId) data.sheetId = DO_CFG.SHEET_ID;

  const ss = openSheetByIdOrUrl(data.sheetId);
  const productsSheet = data.productsSheet || SHEET_PRODUCTS;
  const totalSheet    = data.totalSheet    || SHEET_TOTAL;
  const inSheet       = data.inSheet       || STOCK_SHEETS.IN.name;       // Stock (All)
  const outSheet      = data.outSheet      || STOCK_SHEETS.OUT.name;      // Stock Out
  const receiptSheet  = data.receiptSheet  || STOCK_SHEETS.RECEIPT.name;  // Receipt Log

  // ---- IN (Stock All)
  const inSh = ss.getSheetByName(inSheet);
  const ins = inSh ? inSh.getDataRange().getValues() : [];
  const inHeader = ins[0] || [];
  const inRows = ins.length > 1 ? ins.slice(1) : [];
  const inKeyIdx  = findHeaderIndex(inHeader, [/^part\s*no$/i, /^item\s*code$/i, /sku|barcode/i]);
  const inQtyIdx  = findHeaderIndex(inHeader, [/^qty$/i, /^quantity$/i, /quantity\s*in/i]);
  const inDescIdx = findHeaderIndex(inHeader, [/^description$/i, /desc/i, /product\s*name/i]);

  const inMap = {};
  const nameMap = {};
  inRows.forEach(r => {
    const key = (inKeyIdx >= 0 ? String(r[inKeyIdx] || "") : "").trim();
    if (!key) return;
    const qty = Number(inQtyIdx >= 0 ? r[inQtyIdx] : 0) || 0;
    if (qty) inMap[key] = (inMap[key] || 0) + qty;
    const nm = inDescIdx >= 0 ? String(r[inDescIdx] || "").trim() : "";
    if (nm && !nameMap[key]) nameMap[key] = nm;
  });

  // ---- OUT (Stock Out)
  const outSh = ss.getSheetByName(outSheet);
  const outs = outSh ? outSh.getDataRange().getValues() : [];
  const outHeader = outs[0] || [];
  const outRows = outs.length > 1 ? outs.slice(1) : [];
  const outKeyIdx = findHeaderIndex(outHeader, [/^part\s*no$/i, /^item\s*code$/i, /sku|barcode/i]);
  const outQtyIdx = findHeaderIndex(outHeader, [/quantity\s*out/i, /^qty$/i, /^quantity$/i]);

  const outMap = {};
  outRows.forEach(r => {
    const key = (outKeyIdx >= 0 ? String(r[outKeyIdx] || "") : "").trim();
    if (!key) return;
    const qty = Number(outQtyIdx >= 0 ? r[outQtyIdx] : 0) || 0;
    if (qty) outMap[key] = (outMap[key] || 0) + qty;
  });

  // ---- REQUEST (Receipt Log → Transaction Type == "Request")
  const recSh = ss.getSheetByName(receiptSheet);
  const recs = recSh ? recSh.getDataRange().getValues() : [];
  const reqMap = {};
  if (recs.length > 1) {
    const rh = recs[0] || [];
    const rRows = recs.slice(1);
    const rKeyIdx  = findHeaderIndex(rh, [/^item\s*code$/i, /^item$/i, /sku|barcode/i, /^part\s*no$/i, /^item\s*id$/i]);
    const rQtyIdx  = findHeaderIndex(rh, [/^qty$/i, /^quantity$/i]);
    const rTypeIdx = findHeaderIndex(rh, [/^transaction\s*type$/i, /^type$/i]);
    const rDescIdx = findHeaderIndex(rh, [/^description$/i, /desc/i]);

    rRows.forEach(rr => {
      const t = (rTypeIdx >= 0 ? String(rr[rTypeIdx] || "") : "").trim().toLowerCase();
      if (t !== "request") return;
      const key = (rKeyIdx >= 0 ? String(rr[rKeyIdx] || "") : "").trim();
      if (!key) return;
      const qty = Number(rQtyIdx >= 0 ? rr[rQtyIdx] : 0) || 0;
      if (qty) reqMap[key] = (reqMap[key] || 0) + qty;
      const nm = rDescIdx >= 0 ? String(rr[rDescIdx] || "").trim() : "";
      if (nm && !nameMap[key]) nameMap[key] = nm;
    });
  }

  // ---- Names from Products (highest priority)
  const prodSh = ss.getSheetByName(productsSheet);
  if (prodSh) {
    const pv = prodSh.getDataRange().getValues();
    if (pv && pv.length > 1) {
      const ph = pv[0] || [];
      const keyI = findHeaderIndex(ph, [/^part\s*no$/i, /^item\s*code$/i, /sku|barcode/i]);
      const nmI  = findHeaderIndex(ph, [/^product\s*name$/i, /^name$/i, /item/i]);
      for (let i = 1; i < pv.length; i++) {
        const p = (keyI >= 0 ? String(pv[i][keyI] || "") : "").trim();
        if (!p) continue;
        const n = (nmI  >= 0 ? String(pv[i][nmI]  || "") : "").trim();
        if (n) nameMap[p] = n;
      }
    }
  }

  // ---- Write Total Stock
  const header = ["Part No / SKU","Product Name","Total In","Total Out","Total Request","Balance"];
  const shTotal = _ensureSheet_(ss, totalSheet, header);
  if (shTotal.getLastRow() > 1) {
    shTotal.getRange(2, 1, shTotal.getLastRow() - 1, header.length).clearContent();
  }

  const keys = {};
  Object.keys(inMap).forEach(k => keys[k] = true);
  Object.keys(outMap).forEach(k => keys[k] = true);
  Object.keys(reqMap).forEach(k => keys[k] = true);

  const rows = Object.keys(keys).sort().map(k => {
    const ti = inMap[k] || 0;
    const to = outMap[k] || 0;
    const tr = reqMap[k] || 0;
    const bal = ti - to - tr;
    return [k, nameMap[k] || "", ti, to, tr, bal];
  });

  if (rows.length) {
    shTotal.getRange(2, 1, rows.length, header.length).setValues(rows);
  }

  return jsonOut({ ok: true, updated: rows.length, sheet: shTotal.getName(), columns: header });
}



// =====================================================================
// ========================= DO SECTION =================================
// =====================================================================
function _headerWidth_(sh) {
  const max = Math.max(26, sh.getMaxColumns());
  const headerRow = sh.getRange(1, 1, 1, max).getValues()[0];
  let w = 0;
  for (let i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i] || "").trim() !== "") w = i + 1;
  }
  return w || 0;
}

function _alignToWidth_(rows, width) {
  return rows.map(r => {
    const copy = r.slice(0, width);
    while (copy.length < width) copy.push("");
    return copy;
  });
}
function DO_getCounter_(increment) {
  const fy = Utilities.formatDate(new Date(), DO_CFG.TZ, "yy");
  const key = `DO_COUNTER_${fy}`;
  const props = PropertiesService.getScriptProperties();
  let n = parseInt(props.getProperty(key) || "0", 10);
  if (increment) { n += 1; props.setProperty(key, String(n)); }
  return { fy, n };
}
function DO_nextDoNo_() {
  const { fy, n } = DO_getCounter_(true);
  const pad = String(n).padStart(DO_CFG.COUNTER_PAD, "0");
  return `${DO_CFG.PREFIX}-${fy}-${pad}`;
}
function DO_normalizeDoNoInput_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  if (/^\d+$/.test(s)) {
    const fy = Utilities.formatDate(new Date(), DO_CFG.TZ, "yy");
    const num = String(parseInt(s, 10)).padStart(DO_CFG.COUNTER_PAD, "0");
    return `${DO_CFG.PREFIX}-${fy}-${num}`;
  }
  return s;
}
function DO_isDoNoUsed_(ss, doNo) {
  const sh = _ensureSheet_(ss, DO_SH.HEADER, DO_H_HEADERS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  return vals.some(r => String(r[0]) === String(doNo));
}
function DO_nextFreeLike_(base) {
  const m = /^([A-Za-z]+)-(\d{2})-(\d+)$/.exec(base);
  if (!m) {
    let c = DO_nextDoNo_();
    const ss = _openSS_(DO_CFG.SHEET_ID);
    let guard = 0;
    while (DO_isDoNoUsed_(ss, c) && guard++ < 10000) c = DO_nextDoNo_();
    return c;
  }
  const prefix = m[1], yy = m[2], start = parseInt(m[3], 10);
  let n = start;
  const ss = _openSS_(DO_CFG.SHEET_ID);
  for (let guard = 0; guard < 10000; guard++) {
    const cand = `${prefix}-${yy}-${String(n).padStart(DO_CFG.COUNTER_PAD, "0")}`;
    if (!DO_isDoNoUsed_(ss, cand)) return cand;
    n++;
  }
  throw new Error("Could not find a free DO number");
}
function DO_findHeaderRow_(sh, doNo) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return null;
  const data = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) if (String(data[i][0]) === String(doNo)) return i + 2;
  return null;
}
function DO_fillItemsTable_(body, items) {
  const ElementType = DocumentApp.ElementType;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const el = body.getChild(i);
    if (el.getType() !== ElementType.TABLE) continue;

    const table = el.asTable();

    for (let r = 0; r < table.getNumRows(); r++) {
      const row = table.getRow(r);

      if (row.getText().indexOf("{{ITEM_NAME}}") !== -1) {
        items.forEach((it, idx) => {
          const newRow = table.insertTableRow(r + 1 + idx, row.copy());
          const set = (ci, val) =>
            newRow.getCell(ci).clear().editAsText().setText(val == null ? "" : String(val));

          set(0, String(idx + 1));
          set(1, it.ITEM_DESC || it.description || "");
          set(2, it.ITEM_CODE || it.itemCode || "");
          set(3, it.UNIT || it.unit || "Pcs");
          set(4, Number(it.QTY ?? it.qty ?? 0) || 0);
          if (newRow.getNumCells() > 5) set(5, it.REMARK || it.remark || "");
        });

        table.removeRow(r);
        return true;
      }
    }
  }
  return false;
}

/** 1) DO save — write ONLY to "Stock Out" */
function DO_saveToSheet(payload) {
  try {
    if (!payload || !payload.header || !Array.isArray(payload.items) || payload.items.length === 0) {
      return { ok: false, error: "Missing header/items" };
    }

    const ss = _openSS_(DO_CFG.SHEET_ID);
    const shOut = _ensureSheet_(ss, STOCK_SHEETS.OUT.name, STOCK_SHEETS.OUT.header);

    let doNo = DO_normalizeDoNoInput_(String(payload.header.DO_NO || "")) || DO_nextDoNo_();
    try { if (DO_isDoNoUsed_(ss, doNo)) doNo = DO_nextFreeLike_(doNo); } catch (_) {}

    const H = payload.header || {};
    const companyName  = H.COMPANY_NAME || "";
    const address      = H.COMPANY_ADDR || "";
    const project      = H.PROJECT || "";
    const ref          = H.COMPANY_REFE || "";
    const note         = H.NOTE || "";
    const poRef        = H.PO_REF || "";

    const now = new Date();
    const rows = payload.items.map((it, idx) => {
      const qty = Math.max(0, Number(it.QTY ?? it.qty ?? 0) || 0);
      const sku  = it.ITEM_CODE ?? it.ITEM_ID ?? it.BARCODE ?? it.SKU ?? "";
      const desc = it.ITEM_DESC ?? it.description ?? it.DESC ?? "";
      return [
        doNo, idx + 1, String(sku || ""), String(desc || ""), qty, now,
        it.REMARK || it.remark || "", companyName, address, project, ref, note, poRef
      ];
    });
    if (rows.length) {
      shOut.getRange(shOut.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      ensureStockOutFormatting_(ss);
    }

    const totalQty = payload.items.reduce((s, it) => s + (Number(it.QTY ?? it.qty ?? 0) || 0), 0);

    const headerRowIndex = DO_headerUpsert_(ss, doNo, H, totalQty);

    return {
      ok: true,
      doNo,
      totalQty,
      wroteTo: STOCK_SHEETS.OUT.name,
      rowsAppended: rows.length,
      headerRowIndex
    };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}


/** 2) DO generate doc/pdf (only update DO_Header if a matching row exists) */
function DO_generateDocPdf(args) {
  try {
    if (!args || !args.header || !args.items || !args.items.length) {
      return { ok: false, error: "Missing args.header/items" };
    }
    let { doNo, header, items, headerRowIndex } = args;

    if (!doNo) {
      const wantedDoNo = (header && header.DO_NO) ? String(header.DO_NO).trim() : "";
      if (wantedDoNo) doNo = wantedDoNo;
    }
    if (!doNo) {
      const saved = DO_saveToSheet({ header, items });
      if (!saved.ok) return saved;
      doNo = saved.doNo;
      headerRowIndex = saved.headerRowIndex;
    }

    const docFolder = DriveApp.getFolderById(DO_CFG.DOC_FOLDER_ID);
    const pdfFolder = DriveApp.getFolderById(DO_CFG.PDF_FOLDER_ID);
    const copyName = `${doNo} — ${header?.COMPANY_NAME || "Customer"}`;
    const fileCopy = DriveApp.getFileById(DO_CFG.DOC_TEMPLATE_ID).makeCopy(copyName, docFolder);
    const docId = fileCopy.getId();
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    const totalQty = (items || []).reduce((s, it) => s + (Number(it.QTY ?? it.qty) || 0), 0);
    _replaceAll_(body, {
      DO_NO: doNo,
      DO_DATE: header?.DO_DATE || "",
      COMPANY_NAME: header?.COMPANY_NAME || "",
      COMPANY_ADDR: header?.COMPANY_ADDR || "",
      COMPANY_REFE: header?.COMPANY_REFE || "",
      PROJECT: header?.PROJECT || "",
      PO_REF: header?.PO_REF || "",
      NOTE: header?.NOTE || "",
      TOTAL_QTY: String(totalQty),
    });
    DO_fillItemsTable_(body, items || []);
    doc.saveAndClose();

    const pdfBlob = DriveApp.getFileById(docId).getAs(MimeType.PDF).setName(copyName + ".pdf");
    const pdfFile = pdfFolder.createFile(pdfBlob);
    const docUrl = fileCopy.getUrl();
    const pdfUrl = pdfFile.getUrl();

    const ss = _openSS_(DO_CFG.SHEET_ID);
    const shH = _ensureSheet_(ss, DO_SH.HEADER, DO_H_HEADERS);

    if (!headerRowIndex) headerRowIndex = DO_headerUpsert_(ss, doNo, header, totalQty);

    const col = (name) => DO_H_HEADERS.indexOf(name) + 1;
    if (headerRowIndex) {
      shH.getRange(headerRowIndex, col("Doc_Link")).setValue(docUrl);
      shH.getRange(headerRowIndex, col("PDF_Link")).setValue(pdfUrl);
      shH.getRange(headerRowIndex, col("Status")).setValue("Completed");
    }

    return { ok: true, doNo, docUrl, pdfUrl, headerRowIndex };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}
function DO_headerUpsert_(ss, doNo, H, totalQty) {
  const shH = _ensureSheet_(ss, DO_SH.HEADER, DO_H_HEADERS);
  const row = DO_findHeaderRow_(shH, doNo);
  const col = (name) => DO_H_HEADERS.indexOf(name) + 1;

  if (row) {
    if (!shH.getRange(row, col("DO_Date")).getValue() && H?.DO_DATE) shH.getRange(row, col("DO_Date")).setValue(H.DO_DATE);
    if (!shH.getRange(row, col("COMPANY_NAME")).getValue() && H?.COMPANY_NAME) shH.getRange(row, col("COMPANY_NAME")).setValue(H.COMPANY_NAME);
    if (!shH.getRange(row, col("COMPANY_ADDR")).getValue() && H?.COMPANY_ADDR) shH.getRange(row, col("COMPANY_ADDR")).setValue(H.COMPANY_ADDR);
    if (!shH.getRange(row, col("PROJECT")).getValue() && H?.PROJECT) shH.getRange(row, col("PROJECT")).setValue(H.PROJECT);
    if (!shH.getRange(row, col("PO_REF")).getValue() && H?.PO_REF) shH.getRange(row, col("PO_REF")).setValue(H.PO_REF);
    if (!shH.getRange(row, col("NOTE")).getValue() && H?.NOTE) shH.getRange(row, col("NOTE")).setValue(H.NOTE);
    if (!shH.getRange(row, col("TOTAL_QTY")).getValue() && (totalQty != null)) shH.getRange(row, col("TOTAL_QTY")).setValue(totalQty);
    if (!shH.getRange(row, col("Created_TS")).getValue()) shH.getRange(row, col("Created_TS")).setValue(new Date());
    if (!shH.getRange(row, col("Status")).getValue()) shH.getRange(row, col("Status")).setValue("Pending");
    return row;
  }

  const w = DO_H_HEADERS.length;
  const values = [];
  values[DO_H_HEADERS.indexOf("DO_No")]      = doNo;
  values[DO_H_HEADERS.indexOf("DO_Date")]    = H?.DO_DATE || "";
  values[DO_H_HEADERS.indexOf("COMPANY_NAME")] = H?.COMPANY_NAME || "";
  values[DO_H_HEADERS.indexOf("COMPANY_ADDR")] = H?.COMPANY_ADDR || "";
  values[DO_H_HEADERS.indexOf("COMPANY_REFE")] = H?.COMPANY_REFE || "";
  values[DO_H_HEADERS.indexOf("PROJECT")]    = H?.PROJECT || "";
  values[DO_H_HEADERS.indexOf("PO_REF")]     = H?.PO_REF || "";
  values[DO_H_HEADERS.indexOf("NOTE")]       = H?.NOTE || "";
  values[DO_H_HEADERS.indexOf("TOTAL_QTY")]  = totalQty != null ? totalQty : "";
  values[DO_H_HEADERS.indexOf("Created_TS")] = new Date();
  values[DO_H_HEADERS.indexOf("Status")]     = "Pending";
  const rowIdx = shH.getLastRow() + 1;
  shH.getRange(rowIdx, 1, 1, w).setValues([_alignToWidth_([values], w)[0]]);
  return rowIdx;
}

/** do.list — list DO headers (optionally filtered) */
function handleListDO(data) {
  data = data || {};
  var ss = _openSS_(DO_CFG.SHEET_ID);
  var sh = _ensureSheet_(ss, DO_SH.HEADER, DO_H_HEADERS);
  var pack = readSheetAsObjects_(sh);
  var filtered = filterByQuery_(pack.objects, data.query || "");
  var limited = takeLimitNewest_(filtered, data.limit || 200);
  return jsonOut({ ok: true, header: pack.header, rows: limited.map(function(o){return pack.header.map(function(h){return o[h];});}), objects: limited, count: limited.length });
}

/** do.get — return one DO (header + items) by DO_No */
function handleGetDO(data) {
  data = data || {};
  var doNo = String(data.doNo || data.DO_No || "").trim();
  if (!doNo) return jsonOut({ ok: false, error: "Missing doNo" });

  var ss = _openSS_(DO_CFG.SHEET_ID);
  var shH = _ensureSheet_(ss, DO_SH.HEADER, DO_H_HEADERS);
  var shI = _ensureSheet_(ss, DO_SH.ITEMS, DO_I_HEADERS);

  var H = readSheetAsObjects_(shH).objects;
  var I = readSheetAsObjects_(shI).objects;

  var header = H.find(function (h) { return String(getCI_(h, "DO_No")) === doNo; }) || null;
  var items = I.filter(function (it) { return String(getCI_(it, "DO_No")) === doNo; });

  if (!header) return jsonOut({ ok: false, error: "DO not found: " + doNo });
  return jsonOut({ ok: true, doNo: doNo, header: header, items: items, itemsCount: items.length });
}



// =====================================================================
// ========================= REQUEST SECTION ============================
// =====================================================================

function RQ_fmtDate_(d) { return _fmt_(d, RQ_CFG.TZ, RQ_CFG.DATE_FMT); }
function RQ_FY_() {
  return (RQ_CFG.FY && RQ_CFG.FY.trim())
    ? RQ_CFG.FY.trim()
    : Utilities.formatDate(new Date(), RQ_CFG.TZ, "yy");
}
function RQ_nextCounter_() {
  const ss = _openSS_(RQ_CFG.SHEET_ID);
  let sh = ss.getSheetByName(RQ_SH.SETTINGS);
  if (!sh) sh = ss.insertSheet(RQ_SH.SETTINGS);
  const counterCell = sh.getRange("A1");
  const fyCell = sh.getRange("B1");

  const curFy = RQ_FY_();
  let lastFy = String(fyCell.getValue() || "");
  let v = Number(counterCell.getValue());
  if (!v || isNaN(v)) v = 0;

  if (!lastFy || lastFy !== curFy) {
    v = 0; fyCell.setValue(curFy);
  }
  v += 1;
  counterCell.setValue(v);
  return v;
}
function RQ_pad_(num, size) { let s = String(num); while (s.length < size) s = "0" + s; return s; }
function RQ_makeNo_(counter) { return `${RQ_CFG.PREFIX}-${RQ_FY_()}-${RQ_pad_(counter, RQ_CFG.COUNTER_PAD)}`; }

/** Write ONLY to "Receipt Log" and "RQ_Items" */
/** If val is a data: URL, upload to Drive and return a view URL; else return original */
function ensureDriveImageUrl_(val, rqNo, fileName) {
  const s = String(val || "").trim();
  if (!s) return "";
  if (/^data:/i.test(s)) {
    try {
      const up = uploadImage(s, fileName || "image.png", rqNo || "RQ");
      if (up && up.ok && (up.openUrl || up.url)) return up.openUrl || up.url;
      return s;
    } catch (_) {
      return s;
    }
  }
  return s;
}

function RQ_saveToSheet(payload) {
  try {
    if (!payload || !payload.header) return { ok: false, error: "Missing payload.header" };

    const ss = _openSS_(RQ_CFG.SHEET_ID);
    const recHeader = STOCK_SHEETS.RECEIPT.header; // 15 columns
    const recSh = _ensureSheet_(ss, STOCK_SHEETS.RECEIPT.name, recHeader);

    // RQ number (manual or auto)
    let rqNo = (payload.header.rqNo || "").trim();
    if (!rqNo || rqNo.toLowerCase() === "auto") {
      const c = RQ_nextCounter_();
      rqNo = RQ_makeNo_(c);
    }

    const H = payload.header;
    const rqDate   = H.rqDate || RQ_fmtDate_(new Date());
    const typesStr = (H.types || []).join(", ");
    const tz       = RQ_CFG.TZ || "Asia/Phnom_Penh";
    const nowStr   = Utilities.formatDate(new Date(), tz, "M/d/yyyy HH:mm:ss");

    // normalize items + auto-upload data:URL image to Drive, keep link
    let items = Array.isArray(payload.items) ? payload.items : [];
    items = items.map((it, idx) => {
      const imgIn  = it.image || it.imageUrl || "";
      const imgOut = ensureDriveImageUrl_(imgIn, rqNo, (it.itemCode || ("item" + (idx + 1))) + ".png");
      return Object.assign({}, it, { image: imgOut });
    });

    // ---- A) Receipt Log rows (15 columns — must match header order) ----
    const rowsReceipt = (items.length ? items.map(it => ([
      nowStr,                           // Timestamp (formatted)
      it.itemCode || it.sku || "",      // Item Code
      "Request",                        // Transaction Type
      rqDate,                           // Date
      H.so || "",                       // SO
      H.project || "",                  // Project
      H.refQPO || "",                   // Ref.Quot/PO
      typesStr,                         // Request Type
      it.description || "",             // Description
      Number(it.qty) || 0,              // Qty
      it.returnDate || "",              // Return Date (Date)
      it.reason || "",                  // Reason
      it.remark || "",                  // Remark
      H.companyName || "",              // Company Name
      it.image || ""                    // Image Upload (Drive link)
    ])) : [[
      nowStr, "", "Request", rqDate, H.so || "", H.project || "", H.refQPO || "",
      typesStr, "", 0, "", "", "", H.companyName || "", ""
    ]]);

    recSh.getRange(recSh.getLastRow() + 1, 1, rowsReceipt.length, recHeader.length)
         .setValues(rowsReceipt);

    // ---- B) RQ_Items rows (your visible form with No. Ref first) ----
    const itemsSh = _ensureSheet_(ss, RQ_SH.ITEMS, [
      "No. Ref (RQ No)","Line_No","Description","Item_Code","Unit","Qty","Return_Date","Reason","Image","Remark"
    ]);

    const rowsItems = (items.length ? items.map((it, idx) => ([
      rqNo,                              // No. Ref (RQ No)
      idx + 1,                           // Line_No
      it.description || "",
      it.itemCode || "",
      it.unit || "Pcs",
      Number(it.qty) || 0,
      it.returnDate || "",
      it.reason || "",
      it.image || "",                    // same Drive link
      it.remark || ""
    ])) : [[
      rqNo, 1, "", "", "Pcs", 0, "", "", "", ""
    ]]);

    itemsSh.getRange(itemsSh.getLastRow() + 1, 1, rowsItems.length, 10).setValues(rowsItems);

    const totalQty = items.reduce((s, it) => s + (Number(it.qty) || 0), 0);
    return {
      ok: true,
      rqNo,
      totalQty,
      wroteTo: [STOCK_SHEETS.RECEIPT.name, RQ_SH.ITEMS],
      rowsAppended: { receipt: rowsReceipt.length, items: rowsItems.length }
    };
  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

/** Checkbox helper used by doc template */
function RQ_makeTypeChecklist_(arr) {
  const types = ["Warranty","Borrow","T&C","Sample","Others"];
  return types.map(t => (arr.indexOf(t) > -1 ? "☑ " : "☐ ") + t).join("    ");
}
function _withRetries_(fn, tries, baseMs) {
  tries = tries || 3; baseMs = baseMs || 250;
  for (var i = 0; i < tries; i++) {
    try { return fn(); }
    catch (e) {
      if (i === tries - 1) throw e;
      Utilities.sleep(baseMs * Math.pow(2, i));
    }
  }
}

function RQ_generateDocPdf(payload) {
  const STAGE = {
    start:"start", save:"save", copy:"copyTemplate", fill:"fillDoc",
    table:"fillTable", images:"embedImages", pdf:"makePdf", header:"updateHeader", done:"done"
  };
  let stage = STAGE.start;
  try {
    if (!payload || !payload.header) throw new Error("Missing payload.header");

    stage = STAGE.save;
    const saved = RQ_saveToSheet(payload);
    if (!saved.ok) throw new Error("RQ_saveToSheet: " + saved.error);
    const rqNo = saved.rqNo;

    stage = STAGE.copy;
    const tpl = DriveApp.getFileById(RQ_CFG.DOC_TEMPLATE_ID);
    const ofolder = DriveApp.getFolderById(RQ_CFG.DOC_FOLDER_ID);
    const copy = tpl.makeCopy(rqNo, ofolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();

    stage = STAGE.fill;
    const H = payload.header || {};
    const rep = {
      "{{REQUEST_NO}}": rqNo,
      "{{DATE}}": H.rqDate || RQ_fmtDate_(new Date()),
      "{{COMPANY_NAME}}": H.companyName || "",
      "{{SO}}": H.so || "",
      "{{PROJECT}}": H.project || "",
      "{{REF_QPO}}": H.refQPO || "",
      "{{TYPES}}": RQ_makeTypeChecklist_(H.types || []),
      "{{TOTAL_QTY}}": String((payload.items || []).reduce((s, it) => s + (Number(it.qty) || 0), 0)),
      "{{NOTE}}": H.note || "",
    };
    Object.keys(rep).forEach(k => body.replaceText(k, rep[k]));

    stage = STAGE.table;
    const tables = body.getTables();
    if (!tables || !tables.length) throw new Error("Template has no tables. Expect items table at index 0.");
    const t = tables[0];

    // Ensure there is at least 2 rows (header row + one template row)
    while (t.getNumRows() > 2) t.removeRow(2);

    let items = payload.items || [];
    // ensure images are Drive links before embedding into the Doc
    items = items.map((it, idx) => {
      const img = it.image || it.imageUrl || "";
      const fixed = ensureDriveImageUrl_(img, rqNo, (it.itemCode || ("item"+(idx+1))) + ".png");
      return Object.assign({}, it, { image: fixed });
    });

    if (items.length) {
      for (let i = 1; i < items.length; i++) t.appendTableRow(t.getRow(1).copy());
      stage = STAGE.images;

      for (let i = 0; i < items.length; i++) {
        const r = t.getRow(1 + i);
        const it = items[i];

        // Columns: [No., Description, Item Code, Unit, Qty, Return Date, Reason, Image, Remark]
        const values = [
          String(i + 1),
          it.description || "",
          it.itemCode || "",
          it.unit || "Pcs",
          String(Number(it.qty) || 0),
          it.returnDate || "",
          it.reason || "",
          "", // image handled below
          it.remark || "",
        ];

        // Fill non-image cells
        for (let c = 0; c < values.length; c++) {
          if (c === 7) continue; // skip image
          if (c >= r.getNumCells()) throw new Error("Template row has fewer columns than expected. Need 9.");
          r.getCell(c).clear().setText(values[c]);
        }

        // Embed image (col index 7)
        const imgCell = r.getCell(7);
        imgCell.clear();

        const val = String(it.image || "").trim();
        if (!val) {
          continue;
        }

        // Drive ID extraction
        let fileId = "";
        let m = val.match(/\/d\/([a-zA-Z0-9_-]{20,})/); if (m) fileId = m[1];
        if (!fileId) { m = val.match(/[?&]id=([a-zA-Z0-9_-]{20,})/); if (m) fileId = m[1]; }
        if (!fileId && /^[a-zA-Z0-9_-]{20,}$/.test(val)) fileId = val;

        try {
          if (!fileId) {
            // Not a Drive file — just print the URL text
            imgCell.clear().setText(val);
          } else {
            const blob = DriveApp.getFileById(fileId).getBlob();

            // Clear cell and center the image
            imgCell.clear();
            const p = imgCell.appendParagraph("");
            p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

            // Insert and constrain width to ~1.85 cm (height auto, aspect preserved)
            const inline = p.appendInlineImage(blob);
            try { inline.setWidth(RQ_IMAGE_BOX_PT); } catch (_) {}
          }
        } catch (imgErr) {
          imgCell.clear().setText(
            val + " (image embed failed: " + (imgErr && imgErr.message ? imgErr.message : imgErr) + ")"
          );
        }
      }
    } else {
      const r = t.getRow(1);
      for (let c = 0; c < r.getNumCells(); c++) r.getCell(c).clear().setText("");
    }

    doc.saveAndClose();

    stage = STAGE.pdf;
    const pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF).setName(rqNo + ".pdf");
    const pfolder = DriveApp.getFolderById(RQ_CFG.PDF_FOLDER_ID);
    const pdf = pfolder.createFile(pdfBlob);

    stage = STAGE.header; // legacy header update if present
    const ss = _openSS_(RQ_CFG.SHEET_ID);
    const hsh = _ensureSheet_(ss, RQ_SH.HEADER, [
      "RQ_No","RQ_Date","Company_Name","SO","Project","Ref_QPO","Types","Note",
      "TOTAL_QTY","Created_TS","Status","Doc_Link","PDF_Link","Output_Folder"
    ]);
    const data = hsh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === rqNo) {
        hsh.getRange(i + 1, 12).setValue("https://docs.google.com/document/d/" + doc.getId() + "/edit");
        hsh.getRange(i + 1, 13).setValue("https://drive.google.com/file/d/" + pdf.getId() + "/view");
        try {
          const parent = DriveApp.getFileById(doc.getId()).getParents();
          if (parent.hasNext()) hsh.getRange(i + 1, 14).setValue(parent.next().getName());
        } catch (_) {}
        hsh.getRange(i + 1, 11).setValue("Completed");
        break;
      }
    }

    stage = STAGE.done;
    return { ok: true, rqNo, docId: doc.getId(), pdfId: pdf.getId() };
  } catch (e) {
    return { ok: false, error: "RQ_generateDocPdf failed at stage=" + stage + ": " + (e && e.message ? e.message : String(e)) };
  }
}


/** rq.list — list RQ headers (optionally filtered) */
function handleListRQ(data) {
  data = data || {};
  var ss = _openSS_(RQ_CFG.SHEET_ID);
  var sh = _ensureSheet_(ss, RQ_SH.HEADER, [
    "RQ_No","RQ_Date","Company_Name","SO","Project","Ref_QPO","Types","Note",
    "TOTAL_QTY","Created_TS","Status","Doc_Link","PDF_Link","Output_Folder"
  ]);
  var pack = readSheetAsObjects_(sh);
  var filtered = filterByQuery_(pack.objects, data.query || "");
  var limited = takeLimitNewest_(filtered, data.limit || 200);
  return jsonOut({ ok: true, header: pack.header, rows: limited.map(function(o){return pack.header.map(function(h){return o[h];});}), objects: limited, count: limited.length });
}

/** rq.get — return one RQ (header + items) by RQ_No */
function handleGetRQ(data) {
  data = data || {};
  var rqNo = String(data.rqNo || data.RQ_No || "").trim();
  if (!rqNo) return jsonOut({ ok: false, error: "Missing rqNo" });

  var ss = _openSS_(RQ_CFG.SHEET_ID);
  var shH = _ensureSheet_(ss, RQ_SH.HEADER, [
    "RQ_No","RQ_Date","Company_Name","SO","Project","Ref_QPO","Types","Note",
    "TOTAL_QTY","Created_TS","Status","Doc_Link","PDF_Link","Output_Folder"
  ]);

  var shI = _ensureSheet_(ss, RQ_SH.ITEMS, [
    "No. Ref (RQ No)","Line_No","Description","Item_Code","Unit","Qty","Return_Date","Reason","Image","Remark"
  ]);

  var H = readSheetAsObjects_(shH).objects;
  var I = readSheetAsObjects_(shI).objects;

  var header = H.find(function (h) { return String(getCI_(h, "RQ_No")) === rqNo; }) || null;

  // Accept either "RQ_No" or "No. Ref (RQ No)" as the key in items
  var items = I.filter(function (it) {
    var a = getCI_(it, "RQ_No");
    var b = getCI_(it, "No. Ref (RQ No)");
    return String(a || b) === rqNo;
  });

  if (!header) return jsonOut({ ok: false, error: "RQ not found: " + rqNo });
  return jsonOut({ ok: true, rqNo: rqNo, header: header, items: items, itemsCount: items.length });
}



// =====================================================================
// ====================== IMAGE UPLOAD SECTION ==========================
// =====================================================================

function uploadImage(base64, fileName, rqNo) {
  const STAGE = { start: "start", parse: "parse", decode: "decode", createFile: "createFile", share: "share", done: "done" };
  let stage = STAGE.start;
  try {
    if (!IMAGE_FOLDER_ID) throw new Error("IMAGE_FOLDER_ID not configured.");

    stage = STAGE.parse;
    if (typeof base64 !== "string") throw new Error("Image payload must be a data URL string.");
    const m = base64.match(/^data:([^;]+);base64,(.+)$/);
    if (!m) throw new Error("Malformed data URL. Expect data:<mime>;base64,<data>");

    let contentType = m[1];
    const b64 = m[2];

    stage = STAGE.decode;
    const bytes = Utilities.base64Decode(b64);
    if (bytes.length > MAX_BYTES) {
      throw new Error("Image too large (" + (bytes.length/1024/1024).toFixed(2) + " MB). Limit " + (MAX_BYTES/1024/1024) + " MB.");
    }

    const prefix = (rqNo && String(rqNo).toLowerCase() !== "auto") ? String(rqNo).trim() : "RQ";
    const safeStem = String(fileName || "image").replace(/[\\/:*?\"<>|]+/g, "_").slice(0, 60);
    const basename = prefix + "_" + Date.now() + "_" + safeStem;

    let blob = Utilities.newBlob(bytes, contentType, basename);
    const lowerCT = String(contentType).toLowerCase();
    if (!/png|jpeg|jpg/.test(lowerCT)) {
      blob = blob.getAs("image/png").setName(basename + ".png");
      contentType = "image/png";
    }

    stage = STAGE.createFile;
    const folder = DriveApp.getFolderById(IMAGE_FOLDER_ID);
    const file = folder.createFile(blob);
    const fileId = file.getId();

    stage = STAGE.share;
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (_) {}

    const viewUrl = "https://drive.google.com/file/d/" + fileId + "/view";

    stage = STAGE.done;
    return {
      ok: true,
      fileId,
      url: viewUrl,
      openUrl: viewUrl,
      name: file.getName(),
      mimeType: file.getMimeType(),
      sizeBytes: file.getSize()
    };
  } catch (err) {
    return { ok: false, error: "uploadImage failed at stage=" + stage + ": " + (err && err.message ? err.message : String(err)) };
  }
}

// =====================================================================
// ============================ HISTORY =================================
// =====================================================================

/** history.out — rows from "Stock Out" tab (newest first, with query & limit) */
function handleHistoryOut(data) {
  data = data || {};
  var ss = openSheetByIdOrUrl(data.sheetId || DO_CFG.SHEET_ID);
  var sh = _ensureSheet_(ss, STOCK_SHEETS.OUT.name, STOCK_SHEETS.OUT.header);
  var pack = readSheetAsObjects_(sh);
  var filtered = filterByQuery_(pack.objects, data.query || "");
  var limited = takeLimitNewest_(filtered, data.limit || 1000);
  return jsonOut({ ok: true, sheet: STOCK_SHEETS.OUT.name, header: pack.header, rows: limited.map(function(o){return pack.header.map(function(h){return o[h];});}), count: limited.length });
}

/** history.receipt — rows from "Receipt Log" tab (newest first, with query & limit) */
function handleHistoryReceipt(data) {
  data = data || {};
  var ss = openSheetByIdOrUrl(data.sheetId || DO_CFG.SHEET_ID);
  var sh = _ensureSheet_(ss, STOCK_SHEETS.RECEIPT.name, STOCK_SHEETS.RECEIPT.header);
  var pack = readSheetAsObjects_(sh);
  var filtered = filterByQuery_(pack.objects, data.query || "");
  var limited = takeLimitNewest_(filtered, data.limit || 1000);
  return jsonOut({ ok: true, sheet: STOCK_SHEETS.RECEIPT.name, header: pack.header, rows: limited.map(function(o){return pack.header.map(function(h){return o[h];});}), count: limited.length });
}

// =====================================================================
// ============================ MENU ====================================
// =====================================================================

/** Utility: clear all data below the header row but keep headers */
function clearSheetKeepHeader_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh) return;
  const lastRow = sh.getLastRow();
  const lastCol = Math.max(sh.getLastColumn(), 1);
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
}

function menuClearStockAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetKeepHeader_(ss, STOCK_SHEETS.IN.name);
  SpreadsheetApp.getUi().alert('Cleared: "' + STOCK_SHEETS.IN.name + '"');
}
function menuClearStockPO_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetKeepHeader_(ss, STOCK_SHEETS.IN_PO.name);
  SpreadsheetApp.getUi().alert('Cleared: "' + STOCK_SHEETS.IN_PO.name + '"');
}
function menuClearStockLocal_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetKeepHeader_(ss, STOCK_SHEETS.IN_LOCAL.name);
  SpreadsheetApp.getUi().alert('Cleared: "' + STOCK_SHEETS.IN_LOCAL.name + '"');
}
function menuClearStockNonPO_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  clearSheetKeepHeader_(ss, STOCK_SHEETS.IN_NONPO.name);
  SpreadsheetApp.getUi().alert('Cleared: "' + STOCK_SHEETS.IN_NONPO.name + '"');
}

function onOpen() {
  try {
    _withRetries_(function () {
      const ui = SpreadsheetApp.getUi();
      ui.createMenu("Intersys")
        .addItem("Bootstrap (create tabs)", "menuBootstrap_")
        .addItem("Rebuild Totals", "menuRebuildTotals_")
        .addItem("Open DO Form", "showDOForm")
        .addItem("Open Request Form", "showRQForm")
        .addSeparator()
        .addSubMenu(
          ui.createMenu("Danger Zone")
            .addItem('Clear "Stock (All)"', "menuClearStockAll_")
            .addItem('Clear "Stock In (PO)"', "menuClearStockPO_")
            .addItem('Clear "Stock In (Local)"', "menuClearStockLocal_")
            .addItem('Clear "Stock In (NON PO)"', "menuClearStockNonPO_")
        )
        .addToUi();
    }, 3, 250);
  } catch (err) {
    console.warn("onOpen failed: " + (err && err.message ? err.message : err));
    try {
      SpreadsheetApp.getActive().toast(
        "Menu init failed; use Intersys ▶ Bootstrap later or reload."
      );
    } catch (_) {}
  }
}


function menuBootstrap_(){ handleBootstrap({ sheetId: DO_CFG.SHEET_ID }); }
function menuRebuildTotals_(){ handleRebuildTotals({ sheetId: DO_CFG.SHEET_ID }); }

function showDOForm() {
  const html = HtmlService.createHtmlOutput("<div style='padding:12px'>Use the Web App UI.</div>")
    .setTitle("Create Delivery Order").setWidth(520);
  SpreadsheetApp.getUi().showSidebar(html);
}
function showRQForm() {
  const html = HtmlService.createHtmlOutput("<div style='padding:12px'>Use the Web App UI.</div>")
    .setTitle("Request (A4)").setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// =====================================================================
// ============================ ROUTING =================================
// =====================================================================

/** Auto route: saveToSheet (detect DO vs RQ by header shape) */
function saveToSheet(payload) {
  try {
    const h = payload && payload.header;
    if (h && ("DO_DATE" in h || "COMPANY_NAME" in h)) return DO_saveToSheet(payload);
    if (h && ("rqDate" in h || "rqNo" in h)) return RQ_saveToSheet(payload);
    return { ok: false, error: "Cannot detect form type" };
  } catch (e) { return { ok: false, error: e.message }; }
}

/** Auto route: generateDocPdf (detect DO vs RQ) */
function generateDocPdf(payload) {
  try {
    const h = payload && payload.header;
    if (h && ("DO_DATE" in h || "COMPANY_NAME" in h))
      return DO_generateDocPdf({ doNo: payload.doNo, headerRowIndex: payload.headerRowIndex, header: h, items: payload.items || [] });
    if (h && ("rqDate" in h || "rqNo" in h)) return RQ_generateDocPdf(payload);
    return { ok: false, error: "Cannot detect form type" };
  } catch (e) { return { ok: false, error: e.message }; }
}

/** Web App entrypoints */
function doGet(e) {
  const body = _parseBody_(e);
  const action = _getAction_(e, body);

  // Health & diagnostics
  if (action === "ping" || action === "selfTest") return jsonOut({ ok: true, message: "pong", time: new Date().toISOString() });
  if (action === "diag") return jsonOut(diagCheck_());

  // Setup & utilities
  if (action === "bootstrap")             return handleBootstrap({ sheetId: e.parameter.sheetId || body.sheetId || DO_CFG.SHEET_ID });
  if (action === "uploadImage")           return jsonOut(uploadImage(body.base64, body.fileName, body.rqNo));

  // Auto routes
  if (action === "saveToSheet")           return jsonOut(saveToSheet(body));
  if (action === "generateDocPdf")        return jsonOut(generateDocPdf(body));

  // Stock API
  if (action === "appendStructured")      return handleAppendStructured(body);
  if (action === "listEntries")           return handleListEntries({ sheetId: e.parameter.sheetId || body.sheetId || DO_CFG.SHEET_ID, sheetName: e.parameter.sheetName || body.sheetName, query: e.parameter.query || body.query, limit: e.parameter.limit || body.limit });
  if (action === "rebuildTotals")         return handleRebuildTotals({ sheetId: e.parameter.sheetId || body.sheetId || DO_CFG.SHEET_ID });

  // Direct DO/RQ create/generate
  if (action === "do.saveToSheet")        return jsonOut(DO_saveToSheet(body));
  if (action === "do.generateDocPdf")     return jsonOut(DO_generateDocPdf(body));
  if (action === "rq.saveToSheet")        return jsonOut(RQ_saveToSheet(body));
  if (action === "rq.generateDocPdf")     return jsonOut(RQ_generateDocPdf(body));

  // DO / RQ read endpoints
  if (action === "do.list")               return handleListDO(e ? e.parameter : body);
  if (action === "do.get")                return handleGetDO(e ? e.parameter : body);
  if (action === "rq.list")               return handleListRQ(e ? e.parameter : body);
  if (action === "rq.get")                return handleGetRQ(e ? e.parameter : body);

  // History feeds
  if (action === "history.out")           return handleHistoryOut(e ? e.parameter : body);
  if (action === "history.receipt")       return handleHistoryReceipt(e ? e.parameter : body);

  return jsonOut({ ok: false, error: "Unknown action", saw: { params: e && e.parameter, body } });
}

function doPost(e) {
  const body = _parseBody_(e);
  const action = _getAction_(e, body);

  // Health & diagnostics
  if (action === "ping" || action === "selfTest") return jsonOut({ ok: true, message: "pong", time: new Date().toISOString() });
  if (action === "diag") return jsonOut(diagCheck_());

  // Setup & utilities
  if (action === "bootstrap")               return handleBootstrap({ sheetId: body.sheetId || DO_CFG.SHEET_ID });
  if (action === "uploadImage")             return jsonOut(uploadImage(body.base64, body.fileName, body.rqNo));

  // Auto routes
  if (action === "saveToSheet")             return jsonOut(saveToSheet(body));
  if (action === "generateDocPdf")          return jsonOut(generateDocPdf(body));

  // Stock API
  if (action === "appendStructured")        return handleAppendStructured(body);
  if (action === "listEntries")             return handleListEntries(body);
  if (action === "rebuildTotals")           return handleRebuildTotals({ sheetId: body.sheetId || DO_CFG.SHEET_ID });

  // Direct DO/RQ create/generate
  if (action === "do.saveToSheet")          return jsonOut(DO_saveToSheet(body));
  if (action === "do.generateDocPdf")       return jsonOut(DO_generateDocPdf(body));
  if (action === "rq.saveToSheet")          return jsonOut(RQ_saveToSheet(body));
  if (action === "rq.generateDocPdf")       return jsonOut(RQ_generateDocPdf(body));

  // DO / RQ read endpoints
  if (action === "do.list")                 return handleListDO(body);
  if (action === "do.get")                  return handleGetDO(body);
  if (action === "rq.list")                 return handleListRQ(body);
  if (action === "rq.get")                  return handleGetRQ(body);

  // History feeds
  if (action === "history.out")             return handleHistoryOut(body);
  if (action === "history.receipt")         return handleHistoryReceipt(body);

  return jsonOut({ ok: false, error: "Unknown action", saw: { params: e && e.parameter, body } });
}
