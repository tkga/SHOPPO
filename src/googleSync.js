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
  accounts: { title: "Accounts", headers: ["ชื่อไอดี", "จำนวนสต๊อกคงเหลือ", "ลงทุนสะสม", "รายรับสะสม"] },
  finance: { title: "Finance", headers: ["วันที่", "ประเภท", "จำนวนเงิน", "หมายเหตุ"] },
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
  return `=IMAGE("https://drive.google.com/uc?export=view&id=${fileId}")`;
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

function orderRow(o, custName, accName) {
  const desc = o.type === "sell_pokemon"
    ? `${o.pokemonName || ""} x${o.quantity || 1}`
    : `${o.hireUsed || 0}/${o.hireTotal || 0} รอบ`;
  return [
    (o.createdAt || "").slice(0, 10),
    o.id,
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

function accountRow(a, data) {
  const stockCount = (a.stock || []).reduce((s, x) => s + (Number(x.quantity) || 0), 0);
  const invested = data.investmentHistory.filter((h) => h.accountId === a.id).reduce((s, h) => s + (Number(h.amount) || 0), 0);
  const income = data.orders.filter((o) => o.sourceAccountId === a.id && !o.cancelled && o.paymentStatus === "paid").reduce((s, o) => s + (Number(o.price) || 0), 0);
  return [a.name, stockCount, invested, income];
}

// Full sync: uploads any pending order photos, then overwrites all four
// sheets with the current data. Returns { orders } — the updated orders
// array with cached driveFileId values; caller should setData(d => ({...d, orders})).
export async function syncAll({ token, spreadsheetId, folderId, data, onStatus }) {
  onStatus?.("กำลังอัปโหลดรูปที่ยังไม่ได้ส่ง...");
  const orders = await uploadPendingOrderImages(token, folderId, data.orders, (done, total) => {
    if (total) onStatus?.(`กำลังอัปโหลดรูป ${done}/${total}...`);
  });

  const custName = (id) => data.customers.find((c) => c.id === id)?.name || "-";
  const spentOf = (id) => orders.filter((o) => o.customerId === id && !o.cancelled).reduce((s, o) => {
    const price = Number(o.price) || 0;
    return s + (o.paymentStatus === "paid" ? price : o.paymentStatus === "partial" ? Number(o.paidAmount) || 0 : 0);
  }, 0);

  onStatus?.("กำลังเขียนข้อมูลลง Sheets...");
  const orderRows = orders.map((o) => orderRow(o, custName));
  const customerRows = data.customers.map((c) => customerRow(c, spentOf));
  const accountRows = data.gameAccounts.map((a) => accountRow(a, { ...data, orders }));
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
    { range: `${SHEETS.orders.title}!A2`, values: orderRows },
    { range: `${SHEETS.customers.title}!A2`, values: customerRows },
    { range: `${SHEETS.accounts.title}!A2`, values: accountRows },
    { range: `${SHEETS.finance.title}!A2`, values: financeRows },
  ].filter((w) => w.values.length > 0);

  if (writes.length) {
    await gfetch(`https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`, token, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ valueInputOption: "USER_ENTERED", data: writes }),
    });
  }

  onStatus?.("เสร็จสิ้น");
  return { orders };
}

export function spreadsheetUrl(id) {
  return `https://docs.google.com/spreadsheets/d/${id}/edit`;
}
