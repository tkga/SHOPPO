// Google Sheets + Drive sync.
//
// How it works, in one paragraph: the shop owner pastes their own Google
// OAuth "Client ID" (created for free in Google Cloud Console — instructions
// in README.md) into Settings, signs in with their own Google account, and
// the app creates one Spreadsheet ("<shop name> - ข้อมูลร้าน") plus one
// Drive folder ("<shop name> - สลิป") the first time it connects. Every sync:
// text/number data (customers, orders, accounts, finance) is written as rows
// in the Spreadsheet; any slip/activity photo attached to an order is
// uploaded once to the Drive folder (never re-uploaded on later syncs — the
// resulting Drive file id is cached on the order itself) and referenced from
// the sheet with an =IMAGE("...") formula so it renders inline like a normal
// photo column.
//
// Everything here runs client-side with the shop owner's own OAuth token —
// there is no backend server and no Anthropic involvement in the sync itself.

const GIS_SRC = "https://accounts.google.com/gsi/client";
const SCOPES = "https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file";

let gisLoadPromise = null;
function loadGis() {
  if (window.google?.accounts?.oauth2) return Promise.resolve();
  if (gisLoadPromise) return gisLoadPromise;
  gisLoadPromise = new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = GIS_SRC;
    s.async = true;
    s.defer = true;
    s.onload = () => resolve();
    s.onerror = () => reject(new Error("โหลดสคริปต์ Google ไม่สำเร็จ (เช็คอินเทอร์เน็ต)"));
    document.head.appendChild(s);
  });
  return gisLoadPromise;
}

let tokenClient = null;
let currentToken = null; // { access_token, expires_at }

function tokenValid() {
  return currentToken && currentToken.access_token && Date.now() < currentToken.expires_at - 30000;
}

// Ask for an access token. interactive=true pops the Google consent screen
// (needed the very first time, or once permission is revoked); interactive=false
// tries a silent/no-prompt renewal first (works while the user's Google
// session is still active) and only prompts if that fails.
export async function requestAccessToken(clientId, { interactive = false } = {}) {
  if (!clientId) throw new Error("ยังไม่ได้ใส่ Google Client ID");
  if (tokenValid() && !interactive) return currentToken.access_token;
  await loadGis();
  return new Promise((resolve, reject) => {
    tokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: SCOPES,
      prompt: interactive ? "consent" : "",
      callback: (resp) => {
        if (resp.error) { reject(new Error(resp.error)); return; }
        currentToken = { access_token: resp.access_token, expires_at: Date.now() + (Number(resp.expires_in) || 3500) * 1000 };
        resolve(currentToken.access_token);
      },
      error_callback: (err) => reject(new Error(err?.message || "เข้าสู่ระบบ Google ไม่สำเร็จ")),
    });
    tokenClient.requestAccessToken({ prompt: interactive ? "consent" : "" });
  });
}

export function disconnectGoogle() {
  if (currentToken?.access_token && window.google?.accounts?.oauth2) {
    try { window.google.accounts.oauth2.revoke(currentToken.access_token, () => {}); } catch { /* ignore */ }
  }
  currentToken = null;
}

async function gfetch(url, token, opts = {}) {
  const res = await fetch(url, {
    ...opts,
    headers: { ...(opts.headers || {}), Authorization: `Bearer ${token}` },
  });
  if (!res.ok) {
    let msg = `${res.status}`;
    try { const j = await res.json(); msg = j?.error?.message || msg; } catch { /* ignore */ }
    throw new Error(`Google API error: ${msg}`);
  }
  return res.status === 204 ? null : res.json();
}

export async function fetchGoogleProfile(token) {
  const res = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", { headers: { Authorization: `Bearer ${token}` } });
  if (!res.ok) return null;
  return res.json();
}

const SHEETS = {
  orders: { title: "Orders", headers: ["วันที่สร้าง", "เลขออเดอร์", "ประเภท", "ลูกค้า", "ไอดีเกม", "รายละเอียด", "ราคา", "ชำระแล้ว", "สถานะชำระ", "สถานะเทรด/งาน", "ยกเลิก?", "รูปสลิป/รูปงาน"] },
  customers: { title: "Customers", headers: ["ชื่อในเกม", "ไอดีในเกมทั้งหมด", "Facebook", "หมายเหตุ", "ยอดใช้จ่ายสะสม"] },
  // Compact — one row per account, so this stays short even when accounts have hundreds of products.
  accounts: { title: "Accounts", headers: ["ชื่อไอดี", "จำนวนชนิดสินค้า (SKU)", "จำนวนคงเหลือรวม", "ลงทุนสะสม", "รายรับสะสม"] },
  // One row per stock item (product), its own tab so it can grow to hundreds of rows per account
  // without burying the account totals. A filter is auto-applied on sync so "ชื่อไอดี" can be
  // filtered down to one account at a time (Data ▸ Filter views, or the dropdown arrow on row 1).
  stock: { title: "Stock", headers: ["ชื่อไอดี", "ชื่อสินค้า (Pokémon)", "ชนิด", "จำนวนคงเหลือ", "แจ้งเตือนเมื่อเหลือ ≤", "รูปสินค้า"] },
  finance: { title: "Finance", headers: ["วันที่", "ประเภท", "จำนวนเงิน", "หมายเหตุ"] },
};

// Kept in sync with POKEMON_VARIANTS in PokemonGoShop.jsx (duplicated here as
// a plain copy so googleSync.js doesn't need to import from the component).
const VARIANT_LABELS = {
  normal: "ปกติ", shiny: "Shiny", shadow: "Shadow", purified: "Purified",
  lucky: "Lucky", alolan: "Alolan", galarian: "Galarian", hisuian: "Hisuian",
  mega: "Mega", xl_perfect: "XL Perfect(100%)",
};

// Create the spreadsheet with all four tabs + header rows. Returns the new spreadsheetId.
async function createSpreadsheet(token, title) {
  const body = {
    properties: { title },
    sheets: Object.values(SHEETS).map((s) => ({ properties: { title: s.title } })),
  };
  const created = await gfetch("https://sheets.googleapis.com/v4/spreadsheets", token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  const spreadsheetId = created.spreadsheetId;
  const data = Object.values(SHEETS).map((s) => ({
    range: `${s.title}!A1`,
    majorDimension: "ROWS",
    values: [s.headers],
  }));
  await gfetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`, token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ valueInputOption: "USER_ENTERED", data }),
  });
  return spreadsheetId;
}

export async function ensureSpreadsheet(token, existingId, shopName) {
  if (existingId) {
    try {
      await gfetch(`https://sheets.googleapis.com/v4/spreadsheets/${existingId}?fields=spreadsheetId`, token);
      return existingId;
    } catch {
      // old id no longer accessible (deleted/unshared) — make a new one below
    }
  }
  return createSpreadsheet(token, `${shopName || "Pokémon GO Shop"} - ข้อมูลร้าน`);
}

async function findOrCreateFolder(token, name) {
  const q = encodeURIComponent(`name = '${name.replace(/'/g, "\\'")}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false`);
  const found = await gfetch(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`, token);
  if (found.files && found.files.length) return found.files[0].id;
  const created = await gfetch("https://www.googleapis.com/drive/v3/files", token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, mimeType: "application/vnd.google-apps.folder" }),
  });
  return created.id;
}

export async function ensureDriveFolder(token, existingId, shopName) {
  if (existingId) {
    try {
      await gfetch(`https://www.googleapis.com/drive/v3/files/${existingId}?fields=id,trashed`, token);
      return existingId;
    } catch {
      // fall through to create/find
    }
  }
  return findOrCreateFolder(token, `${shopName || "Pokémon GO Shop"} - สลิป`);
}

function dataUrlToBlob(dataUrl) {
  const [meta, b64] = dataUrl.split(",");
  const mime = /data:(.*?);base64/.exec(meta)?.[1] || "image/jpeg";
  const bin = atob(b64);
  const arr = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  return new Blob([arr], { type: mime });
}

// Upload one image, make it link-viewable, return its Drive file id.
async function uploadImage(token, folderId, dataUrl, filename) {
  const blob = dataUrlToBlob(dataUrl);
  const metadata = { name: filename, parents: [folderId] };
  const form = new FormData();
  form.append("metadata", new Blob([JSON.stringify(metadata)], { type: "application/json" }));
  form.append("file", blob);
  const res = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id", {
    method: "POST",
    headers: { Authorization: `Bearer ${token}` },
    body: form,
  });
  if (!res.ok) throw new Error("อัปโหลดรูปขึ้น Drive ไม่สำเร็จ");
  const { id } = await res.json();
  // share as "anyone with the link can view" so =IMAGE() can render it and
  // so the shop owner can open the link from the sheet on any device
  await gfetch(`https://www.googleapis.com/drive/v3/files/${id}/permissions`, token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ role: "reader", type: "anyone" }),
  });
  return id;
}

function driveImageFormula(fileId) {
  // mode 4 = fixed pixel size, paired with the bigger row height/column width
  // set in applySheetFormatting() below — otherwise Sheets' default mode (fit
  // to cell) squishes the photo into the default ~21px row, basically invisible.
  // Wrapped in HYPERLINK() so clicking the thumbnail opens Drive's own full-size
  // preview in a new tab (Drive's preview page has its own zoom/download/share
  // controls) — Sheets itself has no built-in click-to-enlarge for =IMAGE().
  const viewUrl = `https://drive.google.com/file/d/${fileId}/view`;
  const imgUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
  return `=HYPERLINK("${viewUrl}", IMAGE("${imgUrl}", 4, 76, 76))`;
}

// Upload any not-yet-uploaded order images. Returns a NEW orders array with
// `driveFileId` cached on each order that had a photo, so re-syncs never
// re-upload the same photo. Call setData with the result to persist the cache.
export async function uploadPendingOrderImages(token, folderId, orders, onProgress) {
  const out = [];
  let uploaded = 0;
  const toUpload = orders.filter((o) => o.proofImageDataUrl && !o.driveFileId);
  for (const o of orders) {
    if (o.proofImageDataUrl && !o.driveFileId) {
      try {
        const fileId = await uploadImage(token, folderId, o.proofImageDataUrl, `order-${o.id}.jpg`);
        out.push({ ...o, driveFileId: fileId });
      } catch (e) {
        console.error("upload image failed for order", o.id, e);
        out.push(o);
      }
      uploaded++;
      onProgress?.(uploaded, toUpload.length);
    } else {
      out.push(o);
    }
  }
  return out;
}

// Upload any not-yet-uploaded product photos on stock items. Returns a NEW gameAccounts array
// with `photoDriveFileId` cached on each stock item that had a photo, so re-syncs never re-upload
// the same photo. Call setData with the result to persist the cache.
export async function uploadPendingStockImages(token, folderId, gameAccounts, onProgress) {
  const toUpload = [];
  gameAccounts.forEach((a) => (a.stock || []).forEach((s) => {
    if (s.photoDataUrl && !s.photoDriveFileId) toUpload.push(s.id);
  }));
  let uploaded = 0;
  const out = [];
  for (const a of gameAccounts) {
    const stock = a.stock || [];
    const newStock = [];
    for (const s of stock) {
      if (s.photoDataUrl && !s.photoDriveFileId) {
        try {
          const fileId = await uploadImage(token, folderId, s.photoDataUrl, `stock-${s.id}.jpg`);
          newStock.push({ ...s, photoDriveFileId: fileId });
        } catch (e) {
          console.error("upload image failed for stock item", s.id, e);
          newStock.push(s);
        }
        uploaded++;
        onProgress?.(uploaded, toUpload.length);
      } else {
        newStock.push(s);
      }
    }
    out.push({ ...a, stock: newStock });
  }
  return out;
}

function orderRow(o, custName, accName) {
  const desc = o.type === "sell_pokemon"
    ? `${o.pokemonName || ""} x${o.quantity || 1}`
    : `${o.hireUsed || 0}/${o.hireTotal || 0} รอบ`;
  return [
    (o.createdAt || "").slice(0, 10),
    o.code || o.id,
    o.type,
    custName(o.customerId),
    o.customerGameId || "",
    desc,
    Number(o.price) || 0,
    Number(o.paidAmount) || (o.paymentStatus === "paid" ? Number(o.price) || 0 : 0),
    o.paymentStatus,
    o.tradeStatus || o.hireStatus || "",
    o.cancelled ? "ยกเลิก" : "",
    o.driveFileId ? driveImageFormula(o.driveFileId) : "",
  ];
}

function customerRow(c, spentOf) {
  return [c.name, (c.gameIds || []).map((g) => g.value).filter(Boolean).join(", "), c.facebook || "", c.note || "", spentOf(c.id)];
}

// One account -> one summary row (stays short no matter how many products the account has).
function accountSummaryRow(a, data) {
  const invested = data.investmentHistory.filter((h) => h.accountId === a.id).reduce((s, h) => s + (Number(h.amount) || 0), 0);
  const income = data.orders.filter((o) => o.sourceAccountId === a.id && !o.cancelled && o.paymentStatus === "paid").reduce((s, o) => s + (Number(o.price) || 0), 0);
  const stock = a.stock || [];
  const totalQty = stock.reduce((s, x) => s + (Number(x.quantity) || 0), 0);
  return [a.name, stock.length, totalQty, invested, income];
}

// One account -> one row per stock item (product) for the dedicated Stock tab, e.g.:
//   MEGUMINLUCK   Rayquaza   Shiny   3   2
//   MEGUMINLUCK   Mewtwo     ปกติ    5   2
// Kept separate from Accounts so an account with hundreds of products doesn't bury the
// account-level totals; filter/sort by "ชื่อไอดี" in this tab to find one account's products.
function stockRows(a) {
  const stock = a.stock || [];
  return stock.map((s) => [
    a.name,
    s.name,
    (s.variants || []).map((v) => VARIANT_LABELS[v] || v).filter(Boolean).join(", ") || "-",
    Number(s.quantity) || 0,
    s.lowStockThreshold ?? "",
    s.photoDriveFileId ? driveImageFormula(s.photoDriveFileId) : "",
  ]);
}

async function getSheetIdMap(token, spreadsheetId) {
  const res = await gfetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=sheets(properties(sheetId,title))`, token);
  const map = {};
  (res.sheets || []).forEach((s) => { map[s.properties.title] = s.properties.sheetId; });
  return map;
}

// Self-healing migration for spreadsheets created by an older version of this app (e.g. before
// the "Stock" tab existed, or when "Accounts" had different columns): adds any missing tabs and
// rewrites every header row to match the current SHEETS definition. Safe/cheap to run every sync.
async function ensureSheetsAndHeaders(token, spreadsheetId) {
  let sheetIdMap = await getSheetIdMap(token, spreadsheetId);
  const missing = Object.values(SHEETS).filter((s) => sheetIdMap[s.title] == null);
  if (missing.length) {
    await gfetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, token, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ requests: missing.map((s) => ({ addSheet: { properties: { title: s.title } } })) }),
    });
    sheetIdMap = await getSheetIdMap(token, spreadsheetId);
  }
  const headerData = Object.values(SHEETS).map((s) => ({ range: `${s.title}!A1`, majorDimension: "ROWS", values: [s.headers] }));
  await gfetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`, token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ valueInputOption: "USER_ENTERED", data: headerData }),
  });
  return sheetIdMap;
}

// Cosmetic pass, run after the data write: (1) auto-filter on Stock (so "ชื่อไอดี" can be
// filtered down to one account even with hundreds of products) and on Orders, and (2) a taller
// row height + wider photo column on Orders so slip photos are actually visible instead of
// squished into the default ~21px row. Failures here are non-fatal — the data write already succeeded.
async function applySheetFormatting(token, spreadsheetId, sheetIdMap, { stockRowCount, orderRowCount }) {
  const requests = [];

  const stockSheetId = sheetIdMap[SHEETS.stock.title];
  if (stockSheetId != null && stockRowCount > 0) {
    requests.push({
      setBasicFilter: {
        filter: { range: { sheetId: stockSheetId, startRowIndex: 0, endRowIndex: stockRowCount + 1, startColumnIndex: 0, endColumnIndex: SHEETS.stock.headers.length } },
      },
    });
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: stockSheetId, dimension: "ROWS", startIndex: 1, endIndex: stockRowCount + 1 },
        properties: { pixelSize: 80 },
        fields: "pixelSize",
      },
    });
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: stockSheetId, dimension: "COLUMNS", startIndex: 5, endIndex: 6 },
        properties: { pixelSize: 90 },
        fields: "pixelSize",
      },
    });
  }

  const ordersSheetId = sheetIdMap[SHEETS.orders.title];
  if (ordersSheetId != null && orderRowCount > 0) {
    requests.push({
      setBasicFilter: {
        filter: { range: { sheetId: ordersSheetId, startRowIndex: 0, endRowIndex: orderRowCount + 1, startColumnIndex: 0, endColumnIndex: SHEETS.orders.headers.length } },
      },
    });
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: ordersSheetId, dimension: "ROWS", startIndex: 1, endIndex: orderRowCount + 1 },
        properties: { pixelSize: 80 },
        fields: "pixelSize",
      },
    });
    requests.push({
      updateDimensionProperties: {
        range: { sheetId: ordersSheetId, dimension: "COLUMNS", startIndex: 11, endIndex: 12 },
        properties: { pixelSize: 90 },
        fields: "pixelSize",
      },
    });
  }

  if (!requests.length) return;
  try {
    await gfetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}:batchUpdate`, token, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ requests }),
    });
  } catch (e) {
    console.error("applySheetFormatting failed (non-fatal)", e);
  }
}

// Full sync: uploads any pending order/stock photos, then overwrites all five sheets with the
// current data. Returns { orders, gameAccounts } — updated copies with cached driveFileId /
// photoDriveFileId values; caller should setData(d => ({ ...d, orders, gameAccounts })).
export async function syncAll({ token, spreadsheetId, folderId, data, onStatus }) {
  onStatus?.("กำลังอัปโหลดรูปที่ยังไม่ได้ส่ง...");
  const orders = await uploadPendingOrderImages(token, folderId, data.orders, (done, total) => {
    if (total) onStatus?.(`กำลังอัปโหลดรูปสลิป ${done}/${total}...`);
  });
  const gameAccounts = await uploadPendingStockImages(token, folderId, data.gameAccounts, (done, total) => {
    if (total) onStatus?.(`กำลังอัปโหลดรูปสินค้า ${done}/${total}...`);
  });

  const custName = (id) => data.customers.find((c) => c.id === id)?.name || "-";
  const spentOf = (id) => orders.filter((o) => o.customerId === id && !o.cancelled).reduce((s, o) => {
    const price = Number(o.price) || 0;
    return s + (o.paymentStatus === "paid" ? price : o.paymentStatus === "partial" ? Number(o.paidAmount) || 0 : 0);
  }, 0);

  onStatus?.("กำลังตรวจสอบโครงสร้างตาราง...");
  const sheetIdMap = await ensureSheetsAndHeaders(token, spreadsheetId);

  onStatus?.("กำลังเขียนข้อมูลลง Sheets...");
  const orderRowsData = orders.map((o) => orderRow(o, custName));
  const customerRows = data.customers.map((c) => customerRow(c, spentOf));
  const accountSummaryRows = gameAccounts.map((a) => accountSummaryRow(a, { ...data, orders }));
  // sorted by Pokémon name (then account name) so every account holding the same product is
  // grouped together — answers "which account(s) have Rayquaza?" at a glance, no filter needed.
  const stockSheetRows = gameAccounts
    .flatMap((a) => stockRows(a))
    .sort((x, y) => (x[1] || "").localeCompare(y[1] || "", "th") || (x[0] || "").localeCompare(y[0] || "", "th"));
  const financeRows = [
    ...data.investmentHistory.map((h) => [h.date, "ลงทุน", -(Number(h.amount) || 0), h.note || ""]),
    ...data.manualTx.map((t) => [t.date, t.type === "income" ? "รายรับ" : "รายจ่าย", (t.type === "income" ? 1 : -1) * (Number(t.amount) || 0), t.note || ""]),
  ].sort((a, b) => (a[0] || "").localeCompare(b[0] || ""));

  // clear old rows below the header first (in case the new data set is shorter), then write fresh
  await gfetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchClear`, token, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ ranges: Object.values(SHEETS).map((s) => `${s.title}!A2:Z100000`) }),
  });

  const writes = [
    { range: `${SHEETS.orders.title}!A2`, values: orderRowsData },
    { range: `${SHEETS.customers.title}!A2`, values: customerRows },
    { range: `${SHEETS.accounts.title}!A2`, values: accountSummaryRows },
    { range: `${SHEETS.stock.title}!A2`, values: stockSheetRows },
    { range: `${SHEETS.finance.title}!A2`, values: financeRows },
  ].filter((w) => w.values.length > 0);

  if (writes.length) {
    await gfetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`, token, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ valueInputOption: "USER_ENTERED", data: writes }),
    });
  }

  onStatus?.("กำลังจัดรูปแบบตาราง...");
  await applySheetFormatting(token, spreadsheetId, sheetIdMap, {
    stockRowCount: stockSheetRows.length,
    orderRowCount: orderRowsData.length,
  });

  onStatus?.("เสร็จสิ้น");
  return { orders, gameAccounts };
}

export function spreadsheetUrl(id) {
  return `https://docs.google.com/spreadsheets/d/${id}/edit`;
}
