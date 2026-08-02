import { useState, useEffect, useRef, useMemo, useCallback } from "react";
import {
  Home, Package, Users, Repeat, MoreHorizontal, Plus, X, ChevronRight,
  Wallet, TrendingUp, TrendingDown, Settings as SettingsIcon, Download,
  Upload, Gamepad2, Heart, Clock, CheckCircle2, Circle, ArrowLeft,
  Trash2, Edit2, FileDown, Printer, Search, BarChart3, Coins, ChevronDown,
  Ban, RotateCcw, AlertTriangle, Copy, Calendar, Boxes, ListFilter, Receipt, Minus,
  Target, Move
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, PieChart,
  Pie, Cell, CartesianGrid
} from "recharts";
import * as XLSX from "xlsx";
import { idbStorage, migrateFromLocalStorage } from "./idb.js";
import {
  requestAccessToken, disconnectGoogle, fetchGoogleProfile,
  ensureSpreadsheet, ensureDriveFolder, syncAll, spreadsheetUrl,
} from "./googleSync.js";

const STORAGE_KEY = "pgs-shop-data-v1";

// IndexedDB-backed storage — same { get(key)/set(key,value) } shape the app
// always used, but with a much higher quota than localStorage so a year's
// worth of orders with attached slip photos doesn't silently stop saving.
// See src/idb.js for why this replaced the old localStorage wrapper.
const storage = idbStorage;

const ORDER_TYPES = {
  sell_pokemon: { label: "ขาย Pokémon", short: "ขาย", emoji: "🐉", color: "#FFCB05" },
  hire_boss: { label: "จ้างตีบอส", short: "ตีบอส", emoji: "🎯", color: "#3B5DC9" },
  hire_invite: { label: "จ้างเชิญตี", short: "เชิญตี", emoji: "📨", color: "#33C481" },
};

const PAYMENT_STATUS = {
  pending: { label: "รอชำระ", color: "#FF5470" },
  partial: { label: "ชำระบางส่วน", color: "#FFCB05" },
  paid: { label: "ชำระแล้ว", color: "#33C481" },
};

const TRADE_STATUS = {
  waiting: { label: "รอเทรด", color: "#8B8DA3" },
  traded: { label: "เทรดแล้ว", color: "#33C481" },
  three_hearts: { label: "ทำ 3 ใจ", color: "#FFCB05" },
};

const HIRE_STATUS = {
  ongoing: { label: "ค้างอยู่", color: "#FF5470" },
  done: { label: "เสร็จสิ้น", color: "#33C481" },
};

const INVEST_TYPES = {
  topup: { label: "เติม Coin" },
  buy_pokemon: { label: "ซื้อ Pokémon" },
};

const POKEMON_VARIANTS = {
  normal: { label: "ปกติ", emoji: "⭐" },
  shiny: { label: "Shiny", emoji: "✨" },
  shadow: { label: "Shadow", emoji: "🌑" },
  purified: { label: "Purified", emoji: "💠" },
  lucky: { label: "Lucky", emoji: "🍀" },
  alolan: { label: "Alolan", emoji: "🌴" },
  galarian: { label: "Galarian", emoji: "⚔️" },
  hisuian: { label: "Hisuian", emoji: "🏔️" },
  mega: { label: "Mega", emoji: "💥" },
  xl_perfect: { label: "XL Perfect(100%)", emoji: "💯" },
};

const HIRE_MODES = {
  scheduled: { label: "ตั้งรอบ" },
  anytime: { label: "ไม่ระบุรอบ (ตีเมื่อสะดวก)" },
};

function emptyData() {
  return {
    settings: {
      shopName: "Pokémon GO Shop", createdAt: Date.now(), lastBackupAt: null, pin: "", pinQuestion: "", pinAnswer: "", logoDataUrl: "", receiptBgDataUrl: "",
      google: { clientId: "", email: "", spreadsheetId: "", folderId: "", autoSync: true, lastSyncAt: null },
    },
    customers: [],
    gameAccounts: [],
    orders: [],
    investmentHistory: [],
    manualTx: [],
    counters: { order: 0 },
  };
}

function genId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function todayStr() {
  return new Date().toISOString().slice(0, 10);
}
function fmtMoney(n) {
  const v = Number(n) || 0;
  return v.toLocaleString("th-TH", { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}
function fmtDate(d) {
  if (!d) return "-";
  try {
    return new Date(d).toLocaleDateString("th-TH", { day: "numeric", month: "short", year: "2-digit" });
  } catch { return d; }
}
function daysBetween(a, b) {
  const A = new Date(a); const B = new Date(b);
  return Math.round((B - A) / 86400000);
}
function clamp0(n) { return Math.max(0, Number(n) || 0); }

// resize an uploaded image file down to a small square data-URL (keeps storage light)
// PNG stays here because logo/receipt-bg need transparency and there's only ever
// one of each, so their size doesn't grow with the shop's order volume.
function fileToLogoDataUrl(file, maxDim = 512, quality = 0.85) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("read failed"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("image decode failed"));
      img.onload = () => {
        const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
        const w = Math.round(img.width * scale);
        const h = Math.round(img.height * scale);
        const canvas = document.createElement("canvas");
        canvas.width = w; canvas.height = h;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/png", quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

// resize + JPEG-compress an uploaded slip/activity photo. Unlike the logo/bg
// helper above, PNG is the wrong format here: PNG is lossless so it barely
// shrinks a real photo, and this runs once PER ORDER — at ~1000 orders/year
// that difference is what fills up storage. JPEG at quality 0.72 keeps slip
// text readable while landing around 40-100KB instead of 300-800KB.
function fileToJpegDataUrl(file, maxDim = 900, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("read failed"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("image decode failed"));
      img.onload = () => {
        const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
        const w = Math.round(img.width * scale);
        const h = Math.round(img.height * scale);
        const canvas = document.createElement("canvas");
        canvas.width = w; canvas.height = h;
        const ctx = canvas.getContext("2d");
        // JPEG has no alpha channel — fill white first so transparent PNG/HEIC
        // screenshots don't turn black.
        ctx.fillStyle = "#ffffff";
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

// update favicon / apple-touch-icon / manifest so the uploaded logo also shows as the app icon
function applyAppIcon(logoDataUrl, shopName) {
  try {
    const setLink = (rel, href, extra = {}) => {
      let link = document.querySelector(`link[rel="${rel}"]`);
      if (!link) {
        link = document.createElement("link");
        link.rel = rel;
        document.head.appendChild(link);
      }
      link.href = href;
      Object.entries(extra).forEach(([k, v]) => link.setAttribute(k, v));
    };
    if (logoDataUrl) {
      setLink("icon", logoDataUrl);
      setLink("apple-touch-icon", logoDataUrl);
      const manifest = {
        name: shopName || "Pokémon GO Shop",
        short_name: (shopName || "PGS Shop").slice(0, 12),
        start_url: "./",
        display: "standalone",
        background_color: "#12131c",
        theme_color: "#12131c",
        icons: [
          { src: logoDataUrl, sizes: "192x192", type: "image/png" },
          { src: logoDataUrl, sizes: "512x512", type: "image/png" },
        ],
      };
      const blob = new Blob([JSON.stringify(manifest)], { type: "application/json" });
      setLink("manifest", URL.createObjectURL(blob));
    } else {
      setLink("manifest", "./manifest.json");
    }
  } catch (e) {
    console.error("applyAppIcon failed", e);
  }
}

// balance still owed on an order, considering partial payments
function orderBalance(o) {
  const price = Number(o.price) || 0;
  if (o.paymentStatus === "paid") return 0;
  if (o.paymentStatus === "partial") return clamp0(price - (Number(o.paidAmount) || 0));
  return price;
}

// migrate older saved data shapes to the current schema (safe to re-run)
function migrateData(parsed) {
  const d = { ...emptyData(), ...parsed };
  d.settings = { ...emptyData().settings, logoDataUrl: "", receiptBgDataUrl: "", ...(parsed.settings || {}) };
  d.settings.google = { ...emptyData().settings.google, ...(parsed.settings?.google || {}) };
  if (!Array.isArray(d.customers)) d.customers = [];
  d.customers = (d.customers || []).map(c => ({
    ...c,
    gameIds: c.gameIds && c.gameIds.length ? c.gameIds : [{ id: genId(), value: c.name || "" }],
  }));
  d.gameAccounts = (d.gameAccounts || []).map(a => ({
    ...a,
    stock: (a.stock || []).map(s => ({ lowStockThreshold: 2, variants: ["normal"], ...s })),
  }));
  d.manualTx = (d.manualTx || []).map(t => ({
    ...t,
    accountId: t.accountId || "",
  }));
  d.orders = (d.orders || []).map(o => {
    const price = Number(o.price) || 0;
    return {
      ...o,
      paymentStatus: o.paymentStatus === "partial" || o.paymentStatus === "paid" ? o.paymentStatus : "pending",
      paidAmount: o.paymentStatus === "paid" ? price : clamp0(o.paidAmount),
      cancelled: !!o.cancelled,
      cancelledAt: o.cancelledAt || null,
      pokemonVariants: o.pokemonVariants || (o.type === "sell_pokemon" ? ["normal"] : []),
      stockItemId: o.stockItemId || null,
      customerGameId: o.customerGameId || "",
      proofImageDataUrl: o.proofImageDataUrl || "",
      driveFileId: o.driveFileId || null,
      hireMode: o.hireMode || "anytime",
      rounds: o.rounds || [],
      hireTotal: o.hireTotal != null ? clamp0(o.hireTotal) : (o.type !== "sell_pokemon" ? (clamp0(o.quantity) || 1) : 0),
      hireUsed: clamp0(o.hireUsed) || 0,
      hireStatus: o.hireStatus === "done" ? "done" : "ongoing",
      cancelHistory: Array.isArray(o.cancelHistory) ? o.cancelHistory : [],
    };
  });
  return d;
}

// adjust a stock item's quantity by delta (+restore / -deduct); no-op if ids missing
function adjustStock(gameAccounts, accountId, stockItemId, delta) {
  if (!accountId || !stockItemId || !delta) return gameAccounts;
  return gameAccounts.map(a => {
    if (a.id !== accountId) return a;
    return {
      ...a,
      stock: (a.stock || []).map(s => s.id === stockItemId ? { ...s, quantity: clamp0((Number(s.quantity) || 0) + delta) } : s),
    };
  });
}

const GlobalStyle = () => (
  <style>{`
    @import url('https://fonts.googleapis.com/css2?family=Baloo+2:wght@500;600;700;800&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600;700&display=swap');
    .pgs-root {
      --bg: #0c0d15;
      --surface: #161826;
      --surface2: #1e2033;
      --border: #2c2f46;
      --yellow: #ffcb05;
      --blue: #4d68e0;
      --green: #33c481;
      --red: #ff5470;
      --text: #f2f3f8;
      --muted: #8b8da6;
      --radius: 16px;
      font-family: 'Inter', system-ui, sans-serif;
      background:
        radial-gradient(circle at 15% 0%, rgba(77,104,224,0.16) 0%, transparent 45%),
        radial-gradient(circle at 100% 20%, rgba(255,203,5,0.09) 0%, transparent 40%),
        var(--bg);
      color: var(--text);
      width: 100%;
      max-width: 430px;
      margin: 0 auto;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      position: relative;
      overflow-x: hidden;
    }
    .pgs-root * { box-sizing: border-box; }
    .pgs-root button { color: inherit; font-family: inherit; }
    .pgs-display { font-family: 'Baloo 2', 'Inter', sans-serif; }
    .pgs-mono { font-family: 'JetBrains Mono', monospace; }
    .pgs-scroll {
      flex: 1;
      overflow-y: auto;
      padding: 16px 16px 96px 16px;
    }
    .pgs-scroll::-webkit-scrollbar { width: 0; }
    .pgs-header {
      position: sticky; top: 0; z-index: 20;
      background: linear-gradient(180deg, var(--bg) 80%, transparent);
      padding: 18px 16px 8px 16px;
      display: flex; align-items: center; justify-content: space-between;
    }
    .pgs-ball {
      width: 30px; height: 30px; border-radius: 50%;
      background: linear-gradient(180deg, #ee1515 0%, #ee1515 46%, #14151f 46%, #14151f 54%, #fff 54%, #fff 100%);
      border: 2px solid #14151f; position: relative; flex-shrink: 0;
      box-shadow: 0 0 0 1px rgba(255,203,5,0.25), 0 4px 14px rgba(0,0,0,0.5);
    }
    .pgs-ball::after {
      content: ''; position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%);
      width: 9px; height: 9px; border-radius: 50%; background: #fff; border: 2px solid #14151f;
    }
    .pgs-card {
      background: linear-gradient(165deg, var(--surface) 0%, rgba(30,32,51,0.7) 100%);
      border: 1px solid var(--border);
      border-radius: var(--radius);
      padding: 14px;
      box-shadow: 0 1px 0 rgba(255,255,255,0.03) inset, 0 8px 20px -12px rgba(0,0,0,0.6);
      transition: border-color .15s ease, transform .15s ease, box-shadow .15s ease;
    }
    button.pgs-card { text-decoration: none; }
    button.pgs-card:active { transform: scale(0.985); }
    .pgs-statcard {
      background: linear-gradient(165deg, var(--surface) 0%, rgba(30,32,51,0.7) 100%);
      border: 1px solid var(--border);
      border-radius: 14px;
      padding: 12px 14px;
      box-shadow: 0 1px 0 rgba(255,255,255,0.03) inset, 0 8px 20px -12px rgba(0,0,0,0.6);
    }
    .pgs-btn {
      display: inline-flex; align-items: center; justify-content: center; gap: 6px;
      border-radius: 12px; padding: 10px 16px; font-weight: 600; font-size: 14px;
      border: none; cursor: pointer; transition: transform .1s ease, box-shadow .15s ease, filter .15s ease;
    }
    .pgs-btn:active { transform: scale(0.96); }
    .pgs-btn-primary {
      background: linear-gradient(135deg, #ffe066 0%, var(--yellow) 55%, #ffb700 100%);
      color: #14151f;
      box-shadow: 0 6px 18px -6px rgba(255,203,5,0.55);
    }
    .pgs-btn-primary:hover { filter: brightness(1.04); }
    .pgs-btn-outline { background: rgba(255,255,255,0.02); color: var(--text); border: 1px solid var(--border); }
    .pgs-btn-outline:hover { border-color: rgba(255,203,5,0.4); background: rgba(255,203,5,0.05); }
    .pgs-btn-danger { background: rgba(255,84,112,0.15); color: var(--red); }
    .pgs-input, .pgs-select, .pgs-textarea {
      width: 100%; background: var(--surface2); border: 1px solid var(--border);
      color: var(--text); border-radius: 10px; padding: 10px 12px; font-size: 14px;
      outline: none; font-family: inherit; transition: border-color .15s ease, box-shadow .15s ease;
    }
    .pgs-input:focus, .pgs-select:focus, .pgs-textarea:focus { border-color: var(--yellow); box-shadow: 0 0 0 3px rgba(255,203,5,0.14); }
    .pgs-label { font-size: 12px; color: var(--muted); margin-bottom: 6px; display: block; font-weight: 500; }
    .pgs-field { margin-bottom: 14px; }
    .pgs-badge {
      display: inline-flex; align-items: center; gap: 4px; font-size: 11px; font-weight: 600;
      padding: 3px 9px; border-radius: 999px;
    }
    .pgs-bottomnav {
      position: sticky; bottom: 0; z-index: 30;
      background: rgba(18,19,30,0.85); backdrop-filter: blur(16px) saturate(160%);
      border-top: 1px solid var(--border);
      box-shadow: 0 -8px 30px -12px rgba(0,0,0,0.7);
      display: flex; justify-content: space-around; padding: 8px 4px calc(8px + env(safe-area-inset-bottom));
      max-width: 430px; margin: 0 auto; width: 100%;
    }
    .pgs-navitem {
      display: flex; flex-direction: column; align-items: center; gap: 3px;
      color: var(--muted); font-size: 10px; font-weight: 600; background: none; border: none;
      padding: 6px 10px; border-radius: 10px; cursor: pointer; transition: color .15s ease;
    }
    .pgs-navitem.active { color: var(--yellow); text-shadow: 0 0 14px rgba(255,203,5,0.5); }
    .pgs-overlay {
      position: fixed; inset: 0; background: rgba(6,7,11,0.72); backdrop-filter: blur(2px); z-index: 50;
      display: flex; align-items: flex-end; justify-content: center;
    }
    .pgs-sheet {
      background: linear-gradient(180deg, #14151f 0%, var(--bg) 100%); width: 100%; max-width: 430px; border-radius: 22px 22px 0 0;
      max-height: 88vh; overflow-y: auto; padding: 18px 16px calc(18px + env(safe-area-inset-bottom));
      border: 1px solid var(--border); border-bottom: none;
      box-shadow: 0 -20px 60px -20px rgba(0,0,0,0.8);
      animation: pgs-up .22s ease;
    }
    @keyframes pgs-up { from { transform: translateY(24px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
    .pgs-sectiontitle {
      font-size: 12px; text-transform: uppercase; letter-spacing: 0.08em; color: var(--muted);
      font-weight: 700; margin: 18px 0 8px 2px;
    }
    .pgs-empty {
      text-align: center; padding: 36px 12px; color: var(--muted); font-size: 13px;
    }
    .pgs-row { display: flex; align-items: center; justify-content: space-between; }
    .pgs-chip {
      background: var(--surface2); border: 1px solid var(--border); border-radius: 999px;
      padding: 6px 12px; font-size: 12px; font-weight: 600; color: var(--muted); cursor: pointer;
      transition: all .15s ease;
    }
    .pgs-chip:hover { border-color: rgba(255,203,5,0.35); }
    .pgs-chip.active {
      background: linear-gradient(135deg, #ffe066, var(--yellow));
      color: #14151f; border-color: var(--yellow);
      box-shadow: 0 4px 14px -4px rgba(255,203,5,0.6);
    }
    .pgs-toast {
      position: fixed; bottom: 90px; left: 50%; transform: translateX(-50%);
      background: rgba(30,32,51,0.95); backdrop-filter: blur(8px); border: 1px solid var(--border); color: var(--text);
      padding: 10px 18px; border-radius: 999px; font-size: 13px; z-index: 80; font-weight: 600;
      box-shadow: 0 8px 24px rgba(0,0,0,0.5);
    }
    .pgs-strike { text-decoration: line-through; opacity: 0.55; }
    .pgs-iconbtn {
      background: var(--surface2); border: 1px solid var(--border); border-radius: 10px;
      padding: 7px; cursor: pointer; display: inline-flex; align-items: center; justify-content: center;
      transition: border-color .15s ease;
    }
    .pgs-iconbtn:hover { border-color: rgba(255,203,5,0.4); }
    .pgs-roundrow {
      display: flex; align-items: center; gap: 6px; background: var(--surface2);
      border: 1px solid var(--border); border-radius: 10px; padding: 8px; margin-bottom: 6px;
    }
    .pgs-cancelbanner {
      background: rgba(255,84,112,0.12); color: var(--red); border: 1px solid rgba(255,84,112,0.35);
      border-radius: 10px; padding: 8px 10px; font-size: 12px; font-weight: 600; margin-bottom: 12px;
      display: flex; align-items: center; gap: 6px;
    }
    .pgs-receiptbox {
      background: var(--surface2); border: 1px solid var(--border); border-radius: 12px;
      padding: 12px; white-space: pre-wrap; font-size: 12px; line-height: 1.6; margin-bottom: 12px;
    }
  `}</style>
);

function StatusDot({ payment, trade, cancelled }) {
  if (cancelled) {
    return (
      <div style={{ display: "flex", gap: 4 }}>
        <span className="pgs-badge" style={{ background: "rgba(255,84,112,0.15)", color: "var(--red)" }}>
          <Ban size={9} /> ยกเลิกแล้ว
        </span>
      </div>
    );
  }
  return (
    <div style={{ display: "flex", gap: 4, flexWrap: "wrap", justifyContent: "flex-end" }}>
      {payment && (
        <span className="pgs-badge" style={{ background: PAYMENT_STATUS[payment].color + "22", color: PAYMENT_STATUS[payment].color }}>
          <Circle size={7} fill={PAYMENT_STATUS[payment].color} stroke="none" /> {PAYMENT_STATUS[payment].label}
        </span>
      )}
      {trade && (
        <span className="pgs-badge" style={{ background: TRADE_STATUS[trade].color + "22", color: TRADE_STATUS[trade].color }}>
          {trade === "three_hearts" ? <Heart size={9} fill={TRADE_STATUS[trade].color} stroke="none" /> : <Circle size={7} fill={TRADE_STATUS[trade].color} stroke="none" />} {TRADE_STATUS[trade].label}
        </span>
      )}
    </div>
  );
}

function StatCard({ icon: Icon, label, value, color, sub }) {
  return (
    <div className="pgs-statcard">
      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 6 }}>
        <Icon size={14} color={color || "var(--muted)"} />
        <span style={{ fontSize: 11, color: "var(--muted)", fontWeight: 600 }}>{label}</span>
      </div>
      <div className="pgs-mono pgs-display" style={{ fontSize: 20, fontWeight: 700, color: color || "var(--text)" }}>{value}</div>
      {sub && <div style={{ fontSize: 10, color: "var(--muted)", marginTop: 2 }}>{sub}</div>}
    </div>
  );
}

function Modal({ title, onClose, children, footer }) {
  return (
    <div className="pgs-overlay" onClick={onClose}>
      <div className="pgs-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="pgs-row" style={{ marginBottom: 14 }}>
          <h3 className="pgs-display" style={{ fontSize: 18, fontWeight: 700, margin: 0 }}>{title}</h3>
          <button className="pgs-btn pgs-btn-outline" style={{ padding: 8 }} onClick={onClose}><X size={16} /></button>
        </div>
        {children}
        {footer && <div style={{ marginTop: 16 }}>{footer}</div>}
      </div>
    </div>
  );
}

function EmptyState({ text }) {
  return (
    <div className="pgs-empty">
      <div style={{ fontSize: 30, marginBottom: 6 }}>🎾</div>
      {text}
    </div>
  );
}

function LockScreen({ pin, pinQuestion, pinAnswer, shopName, logoDataUrl, onUnlock, onRecover, onResetPin }) {
  const [value, setValue] = useState("");
  const [err, setErr] = useState(false);
  // recover: false (normal) | "question" (answer the recovery question) | "resetpin" (choose new PIN after a correct answer) |
  // "intro" (explain + file picker) | "setpin" (choose new PIN after a valid backup was read)
  const [recover, setRecover] = useState(false);
  const [answerValue, setAnswerValue] = useState("");
  const [answerErr, setAnswerErr] = useState("");
  const [recoveredData, setRecoveredData] = useState(null);
  const [recoverErr, setRecoverErr] = useState("");
  const [newPin, setNewPin] = useState("");
  const [confirmPin, setConfirmPin] = useState("");
  const fileRef = useRef(null);

  function submit(e) {
    e.preventDefault();
    if (value === pin) { onUnlock(); }
    else { setErr(true); setValue(""); setTimeout(() => setErr(false), 900); }
  }

  function submitAnswer(e) {
    e.preventDefault();
    if (answerValue.trim().toLowerCase() === (pinAnswer || "").trim().toLowerCase() && pinAnswer) {
      setAnswerErr("");
      setRecover("resetpin");
    } else {
      setAnswerErr("คำตอบไม่ถูกต้อง");
    }
  }

  function submitResetPin(e) {
    e.preventDefault();
    if (!/^\d{4,8}$/.test(newPin)) { setAnswerErr("PIN ต้องเป็นตัวเลข 4-8 หลัก"); return; }
    if (newPin !== confirmPin) { setAnswerErr("ยืนยัน PIN ไม่ตรงกัน"); return; }
    onResetPin(newPin);
  }

  function pickFile() {
    setRecoverErr("");
    fileRef.current?.click();
  }

  function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!parsed || typeof parsed !== "object" || !parsed.settings || !Array.isArray(parsed.customers)) {
          throw new Error("bad shape");
        }
        setRecoveredData(parsed);
        setRecover("setpin");
        setRecoverErr("");
      } catch {
        setRecoverErr("ไฟล์นี้ไม่ใช่ไฟล์ Backup ของแอปนี้ (.json) กรุณาเลือกไฟล์ที่ถูกต้อง");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  function finishRecover(skipPin) {
    if (!skipPin) {
      if (!/^\d{4,8}$/.test(newPin)) { setRecoverErr("PIN ต้องเป็นตัวเลข 4-8 หลัก"); return; }
      if (newPin !== confirmPin) { setRecoverErr("ยืนยัน PIN ไม่ตรงกัน"); return; }
    }
    onRecover(recoveredData, skipPin ? "" : newPin);
  }

  if (recover === "question") {
    return (
      <div className="pgs-root" style={{ alignItems: "center", justifyContent: "center", padding: 24 }}>
        <GlobalStyle />
        <div style={{ width: "100%", maxWidth: 300 }}>
          <div className="pgs-display" style={{ fontSize: 17, fontWeight: 700, marginBottom: 8, textAlign: "center" }}>คำถามกู้คืน PIN</div>
          <div style={{ fontSize: 12, color: "var(--muted)", marginBottom: 14, textAlign: "center" }}>{pinQuestion}</div>
          <form onSubmit={submitAnswer}>
            <label className="pgs-label">คำตอบ</label>
            <input className="pgs-input" type="password" style={{ marginBottom: 10 }} autoFocus value={answerValue} onChange={e => setAnswerValue(e.target.value)} />
            {answerErr && <div style={{ fontSize: 12, color: "var(--red)", textAlign: "center", marginBottom: 10 }}>{answerErr}</div>}
            <button type="submit" className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: 8 }}>ยืนยันคำตอบ</button>
          </form>
          <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8 }} onClick={() => { setAnswerErr(""); setRecover("intro"); }}>ตอบไม่ได้ — กู้คืนด้วยไฟล์ Backup แทน</button>
          <button className="pgs-btn pgs-btn-outline" style={{ width: "100%" }} onClick={() => { setAnswerErr(""); setAnswerValue(""); setRecover(false); }}>ย้อนกลับ</button>
        </div>
      </div>
    );
  }

  if (recover === "resetpin") {
    return (
      <div className="pgs-root" style={{ alignItems: "center", justifyContent: "center", padding: 24 }}>
        <GlobalStyle />
        <div style={{ width: "100%", maxWidth: 300 }}>
          <div className="pgs-display" style={{ fontSize: 17, fontWeight: 700, marginBottom: 4, textAlign: "center" }}>ตอบถูกแล้ว!</div>
          <div style={{ fontSize: 12, color: "var(--muted)", marginBottom: 14, textAlign: "center" }}>ตั้งรหัส PIN ใหม่ — ข้อมูลร้านทั้งหมดยังอยู่ครบ ไม่ต้องอัปโหลดไฟล์ใดๆ</div>
          <form onSubmit={submitResetPin}>
            <label className="pgs-label">PIN ใหม่ (ตัวเลข 4-8 หลัก)</label>
            <input className="pgs-input pgs-mono" style={{ marginBottom: 10 }} type="password" inputMode="numeric" autoFocus value={newPin} onChange={e => setNewPin(e.target.value)} placeholder="••••" />
            <label className="pgs-label">ยืนยัน PIN ใหม่</label>
            <input className="pgs-input pgs-mono" style={{ marginBottom: 10 }} type="password" inputMode="numeric" value={confirmPin} onChange={e => setConfirmPin(e.target.value)} placeholder="••••" />
            {answerErr && <div style={{ fontSize: 12, color: "var(--red)", textAlign: "center", marginBottom: 10 }}>{answerErr}</div>}
            <button type="submit" className="pgs-btn pgs-btn-primary" style={{ width: "100%" }}>ตั้ง PIN ใหม่</button>
          </form>
        </div>
      </div>
    );
  }

  if (recover === "intro") {
    return (
      <div className="pgs-root" style={{ alignItems: "center", justifyContent: "center", padding: 24 }}>
        <GlobalStyle />
        <div style={{ width: "100%", maxWidth: 300 }}>
          <div className="pgs-display" style={{ fontSize: 17, fontWeight: 700, marginBottom: 8, textAlign: "center" }}>กู้คืน PIN</div>
          <div style={{ fontSize: 12, color: "var(--muted)", marginBottom: 14, textAlign: "center" }}>
            เลือกไฟล์ Backup (.json) ที่เคยดาวน์โหลดไว้ ระบบจะกู้คืนข้อมูลทั้งหมดและให้ตั้ง PIN ใหม่ โดยไม่ล้างข้อมูลในเครื่อง
          </div>
          <input ref={fileRef} type="file" accept="application/json,.json" style={{ display: "none" }} onChange={handleFile} />
          {recoverErr && <div style={{ fontSize: 12, color: "var(--red)", textAlign: "center", marginBottom: 10 }}>{recoverErr}</div>}
          <button className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: 8 }} onClick={pickFile}>
            <Upload size={15} /> เลือกไฟล์ Backup
          </button>
          {pinQuestion && (
            <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8 }} onClick={() => { setRecoverErr(""); setRecover("question"); }}>ตอบคำถามกู้คืนแทน</button>
          )}
          <button className="pgs-btn pgs-btn-outline" style={{ width: "100%" }} onClick={() => { setRecover(false); setRecoverErr(""); }}>ย้อนกลับ</button>
        </div>
      </div>
    );
  }

  if (recover === "setpin") {
    return (
      <div className="pgs-root" style={{ alignItems: "center", justifyContent: "center", padding: 24 }}>
        <GlobalStyle />
        <div style={{ width: "100%", maxWidth: 300 }}>
          <div className="pgs-display" style={{ fontSize: 17, fontWeight: 700, marginBottom: 4, textAlign: "center" }}>อ่านไฟล์ Backup สำเร็จ</div>
          <div style={{ fontSize: 12, color: "var(--muted)", marginBottom: 14, textAlign: "center" }}>ตั้งรหัส PIN ใหม่เพื่อเข้าใช้งานต่อ (หรือปิดการล็อกไปเลยก็ได้)</div>
          <label className="pgs-label">PIN ใหม่ (ตัวเลข 4-8 หลัก)</label>
          <input className="pgs-input pgs-mono" style={{ marginBottom: 10 }} type="password" inputMode="numeric" value={newPin} onChange={e => setNewPin(e.target.value)} placeholder="••••" />
          <label className="pgs-label">ยืนยัน PIN ใหม่</label>
          <input className="pgs-input pgs-mono" style={{ marginBottom: 10 }} type="password" inputMode="numeric" value={confirmPin} onChange={e => setConfirmPin(e.target.value)} placeholder="••••" />
          {recoverErr && <div style={{ fontSize: 12, color: "var(--red)", textAlign: "center", marginBottom: 10 }}>{recoverErr}</div>}
          <button className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: 8 }} onClick={() => finishRecover(false)}>ตั้ง PIN ใหม่ &amp; กู้คืนข้อมูล</button>
          <button className="pgs-btn pgs-btn-outline" style={{ width: "100%" }} onClick={() => finishRecover(true)}>ไม่ตั้ง PIN (ปิดการล็อก)</button>
        </div>
      </div>
    );
  }

  return (
    <div className="pgs-root" style={{ alignItems: "center", justifyContent: "center", padding: 24 }}>
      <GlobalStyle />
      <div style={{ textAlign: "center", marginBottom: 20 }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 10 }}>
          <ShopLogo logoDataUrl={logoDataUrl} size={54} />
        </div>
        <div className="pgs-display" style={{ fontSize: 18, fontWeight: 700 }}>{shopName}</div>
        <div style={{ fontSize: 12, color: "var(--muted)", marginTop: 4 }}>ใส่รหัส PIN เพื่อเข้าใช้งาน</div>
      </div>
      <form onSubmit={submit} style={{ width: "100%", maxWidth: 280 }}>
        <input
          className="pgs-input pgs-mono" style={{ textAlign: "center", fontSize: 20, letterSpacing: 6, marginBottom: 10, borderColor: err ? "var(--red)" : undefined }}
          type="password" inputMode="numeric" autoFocus value={value}
          onChange={e => setValue(e.target.value)}
          placeholder="••••"
        />
        {err && <div style={{ fontSize: 12, color: "var(--red)", textAlign: "center", marginBottom: 10 }}>รหัส PIN ไม่ถูกต้อง</div>}
        <button type="submit" className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: 10 }}>ปลดล็อก</button>
        <button type="button" className="pgs-btn pgs-btn-outline" style={{ width: "100%" }} onClick={() => setRecover(pinQuestion ? "question" : "intro")}>ลืม PIN?</button>
      </form>
    </div>
  );
}

export default function App() {
  const [data, setData] = useState(emptyData());
  const [loaded, setLoaded] = useState(false);
  const [tab, setTab] = useState("dashboard");
  const [moreOpen, setMoreOpen] = useState(false);
  const [modal, setModal] = useState(null);
  const [toast, setToast] = useState(null);
  const [detail, setDetail] = useState(null);
  const [unlocked, setUnlocked] = useState(false);
  const [showBackupPrompt, setShowBackupPrompt] = useState(false);
  const [googleSyncing, setGoogleSyncing] = useState(false);
  const [googleStatus, setGoogleStatus] = useState("");
  const saveTimer = useRef(null);
  const backupPromptedRef = useRef(false);
  const notifAskedRef = useRef(false);
  const autoSyncTimer = useRef(null);

  // ---- load ----
  useEffect(() => {
    (async () => {
      try {
        await migrateFromLocalStorage(STORAGE_KEY); // one-time: old localStorage data -> IndexedDB
        const res = await storage.get(STORAGE_KEY);
        if (res && res.value) {
          const parsed = JSON.parse(res.value);
          setData(migrateData(parsed));
        }
      } catch (e) {
        // no existing data yet
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  // ---- save (debounced) ----
  useEffect(() => {
    if (!loaded) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        await storage.set(STORAGE_KEY, JSON.stringify(data));
      } catch (e) {
        console.error("save failed", e);
      }
    }, 400);
    return () => clearTimeout(saveTimer.current);
  }, [data, loaded]);

  // ---- auto-sync to Google Sheets/Drive (debounced) — this is the "automatic cloud backup" ----
  // Fires a while after the user stops editing so a burst of changes (e.g. filling
  // out one order) becomes a single sync instead of one per keystroke.
  useEffect(() => {
    if (!loaded) return;
    if (data.settings.pin && !unlocked) return;
    const g = data.settings.google;
    if (!g.spreadsheetId || !g.folderId || !g.autoSync) return;
    if (autoSyncTimer.current) clearTimeout(autoSyncTimer.current);
    autoSyncTimer.current = setTimeout(() => {
      if (!googleSyncing) runSync();
    }, 45000);
    return () => clearTimeout(autoSyncTimer.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [data, loaded, unlocked]);

  // ---- keep favicon / home-screen icon in sync with the uploaded shop logo ----
  useEffect(() => {
    if (!loaded) return;
    applyAppIcon(data.settings.logoDataUrl, data.settings.shopName);
  }, [loaded, data.settings.logoDataUrl, data.settings.shopName]);

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 1800);
  }, []);

  // ---------- derived ----------
  const stats = useMemo(() => {
    const orders = data.orders.filter(o => !o.cancelled);
    const today = todayStr();
    const month = today.slice(0, 7);
    const year = today.slice(0, 4);

    const incomeEntries = [
      ...orders.filter(o => o.paymentStatus === "paid").map(o => ({ date: (o.paidDate || o.createdAt).slice(0, 10), amount: Number(o.price) || 0 })),
      ...orders.filter(o => o.paymentStatus === "partial").map(o => ({ date: (o.paidDate || o.createdAt).slice(0, 10), amount: Number(o.paidAmount) || 0 })),
      ...data.manualTx.filter(t => t.type === "income").map(t => ({ date: t.date, amount: Number(t.amount) || 0 })),
    ];
    const expenseEntries = [
      ...data.investmentHistory.map(h => ({ date: h.date, amount: Number(h.amount) || 0 })),
      ...data.manualTx.filter(t => t.type === "expense").map(t => ({ date: t.date, amount: Number(t.amount) || 0 })),
    ];
    const sumBy = (arr, prefix) => arr.filter(e => e.date && e.date.startsWith(prefix)).reduce((s, e) => s + e.amount, 0);

    const incomeToday = sumBy(incomeEntries, today);
    const incomeMonth = sumBy(incomeEntries, month);
    const incomeYear = sumBy(incomeEntries, year);
    const expenseToday = sumBy(expenseEntries, today);
    const expenseMonth = sumBy(expenseEntries, month);
    const expenseYear = sumBy(expenseEntries, year);

    const totalInvestment = data.investmentHistory.reduce((s, h) => s + (Number(h.amount) || 0), 0);
    const investByAccount = {};
    data.investmentHistory.forEach(h => { investByAccount[h.accountId] = (investByAccount[h.accountId] || 0) + (Number(h.amount) || 0); });

    const pendingPayment = orders.filter(o => o.paymentStatus === "pending" || o.paymentStatus === "partial").length;
    const totalDue = orders.reduce((s, o) => s + orderBalance(o), 0);
    const pendingTrade = orders.filter(o => o.type === "sell_pokemon" && o.tradeStatus === "waiting").length;
    const threeHearts = orders.filter(o => o.type === "sell_pokemon" && o.tradeStatus === "three_hearts").length;
    const cancelledCount = data.orders.filter(o => o.cancelled).length;

    let lowStockCount = 0;
    const lowStockItems = [];
    data.gameAccounts.forEach(a => {
      (a.stock || []).forEach(s => {
        const th = s.lowStockThreshold ?? 2;
        if (clamp0(s.quantity) <= th) { lowStockCount++; lowStockItems.push({ accountId: a.id, accountName: a.name, ...s }); }
      });
    });

    // appointments / hit-rounds coming due (or overdue) within the next 7 days
    const dueSoonItems = [];
    orders.forEach(o => {
      if (o.appointmentDate) {
        const remain = daysBetween(today, o.appointmentDate);
        if (remain <= 7) dueSoonItems.push({ orderId: o.id, customerId: o.customerId, date: o.appointmentDate, remain, kind: "appointment" });
      }
      (o.rounds || []).forEach(r => {
        if (!r.done && r.date) {
          const remain = daysBetween(today, r.date);
          if (remain <= 7) dueSoonItems.push({ orderId: o.id, customerId: o.customerId, date: r.date, remain, kind: "round" });
        }
      });
    });
    dueSoonItems.sort((a, b) => a.remain - b.remain);

    return {
      incomeToday, incomeMonth, incomeYear, expenseToday, expenseMonth, expenseYear,
      profitToday: incomeToday - expenseToday, profitMonth: incomeMonth - expenseMonth, profitYear: incomeYear - expenseYear,
      totalInvestment, investByAccount, pendingPayment, totalDue, pendingTrade, threeHearts, cancelledCount,
      totalOrders: orders.length, incomeEntries, expenseEntries, lowStockCount, lowStockItems,
      dueSoonCount: dueSoonItems.length, dueSoonItems,
    };
  }, [data]);

  const custName = (id) => data.customers.find(c => c.id === id)?.name || "-";
  const accName = (id) => data.gameAccounts.find(a => a.id === id)?.name || "-";

  // ---- prominent backup reminder: pop up right when the app opens (not just a quiet dashboard card) ----
  useEffect(() => {
    if (!loaded) return;
    if (data.settings.pin && !unlocked) return; // wait until the user is actually in the app
    if (backupPromptedRef.current) return;
    backupPromptedRef.current = true;
    const days = data.settings.lastBackupAt ? daysBetween(data.settings.lastBackupAt, new Date().toISOString()) : null;
    const needs = days === null ? (data.orders.length + data.customers.length > 0) : days >= 1;
    if (needs) setShowBackupPrompt(true);
  }, [loaded, unlocked, data.settings.pin]);

  // ---- ask for notification permission once the user is in the app ----
  useEffect(() => {
    if (!loaded) return;
    if (data.settings.pin && !unlocked) return;
    if (notifAskedRef.current) return;
    if (typeof Notification === "undefined") return;
    notifAskedRef.current = true;
    if (Notification.permission === "default") Notification.requestPermission();
  }, [loaded, unlocked, data.settings.pin]);

  // ---- notify 1 day (or same-day) before a boss/invite round, appointment, or 3-hearts trade comes due ----
  useEffect(() => {
    if (!loaded) return;
    if (data.settings.pin && !unlocked) return;
    if (typeof Notification === "undefined" || Notification.permission !== "granted") return;
    const dueNow = stats.dueSoonItems.filter(it => it.remain <= 1 && it.remain >= 0);
    if (dueNow.length === 0) return;
    const notifiedKey = "pgs-notified-" + todayStr();
    let notifiedIds = [];
    try { notifiedIds = JSON.parse(window.localStorage.getItem(notifiedKey) || "[]"); } catch { notifiedIds = []; }
    const toNotify = dueNow.filter(it => !notifiedIds.includes(`${it.orderId}-${it.kind}-${it.date}`));
    if (toNotify.length === 0) return;
    toNotify.forEach(it => {
      const label = it.kind === "round" ? "รอบตี" : "นัดหมาย/ทำ 3 ใจ";
      const name = custName(it.customerId);
      const body = it.remain === 0 ? `${label} ของ ${name} ถึงกำหนดวันนี้` : `${label} ของ ${name} ถึงกำหนดพรุ่งนี้`;
      try { new Notification(data.settings.shopName || "Pokémon GO Shop", { body, tag: `${it.orderId}-${it.kind}` }); } catch { /* ignore */ }
    });
    try { window.localStorage.setItem(notifiedKey, JSON.stringify([...notifiedIds, ...toNotify.map(it => `${it.orderId}-${it.kind}-${it.date}`)])); } catch { /* ignore */ }
  }, [loaded, unlocked, data.settings.pin, stats.dueSoonItems]);

  // ---- register the service worker so the app keeps working offline ----
  useEffect(() => {
    if ("serviceWorker" in navigator) {
      navigator.serviceWorker.register("./sw.js").catch(() => { /* ignore — offline support just won't be available */ });
    }
  }, []);

  // ---------- CRUD ----------
  function saveCustomer(item) {
    setData(d => {
      const exists = d.customers.some(c => c.id === item.id);
      return { ...d, customers: exists ? d.customers.map(c => c.id === item.id ? item : c) : [...d.customers, item] };
    });
    showToast(item._isNew ? "เพิ่มลูกค้าแล้ว" : "บันทึกแล้ว");
  }
  function deleteCustomer(id) {
    setData(d => ({ ...d, customers: d.customers.filter(c => c.id !== id) }));
    showToast("ลบลูกค้าแล้ว");
  }
  function saveAccount(item) {
    setData(d => {
      const exists = d.gameAccounts.some(a => a.id === item.id);
      return { ...d, gameAccounts: exists ? d.gameAccounts.map(a => a.id === item.id ? item : a) : [...d.gameAccounts, item] };
    });
    showToast("บันทึกไอดีแล้ว");
  }
  function deleteAccount(id) {
    setData(d => ({ ...d, gameAccounts: d.gameAccounts.filter(a => a.id !== id) }));
    showToast("ลบไอดีแล้ว");
  }
  function saveOrder(item, isNew) {
    setData(d => {
      let counters = d.counters;
      let orders;
      let gameAccounts = d.gameAccounts;
      if (isNew) {
        const n = (d.counters.order || 0) + 1;
        counters = { ...d.counters, order: n };
        item.code = `ORD-${String(n).padStart(4, "0")}`;
        orders = [item, ...d.orders];
        if (!item.cancelled && item.type === "sell_pokemon" && item.stockItemId) {
          gameAccounts = adjustStock(gameAccounts, item.sourceAccountId, item.stockItemId, -clamp0(item.quantity));
        }
      } else {
        const old = d.orders.find(o => o.id === item.id);
        orders = d.orders.map(o => o.id === item.id ? item : o);
        // only reconcile stock when the order stays "active" through the edit; cancel/restore handle stock separately
        if (old && !old.cancelled && !item.cancelled) {
          if (old.type === "sell_pokemon" && old.stockItemId) {
            gameAccounts = adjustStock(gameAccounts, old.sourceAccountId, old.stockItemId, clamp0(old.quantity));
          }
          if (item.type === "sell_pokemon" && item.stockItemId) {
            gameAccounts = adjustStock(gameAccounts, item.sourceAccountId, item.stockItemId, -clamp0(item.quantity));
          }
        }
      }
      return { ...d, orders, counters, gameAccounts };
    });
    showToast(isNew ? "สร้างออเดอร์แล้ว" : "บันทึกออเดอร์แล้ว");
  }
  function quickComplete(order) {
    const price = Number(order.price) || 0;
    saveOrder({ ...order, paymentStatus: "paid", paidAmount: price, paidDate: order.paidDate || new Date().toISOString() }, false);
  }
  function quickUseHire(order, delta) {
    const total = clamp0(order.hireTotal);
    const nextUsed = clamp0((clamp0(order.hireUsed) || 0) + delta);
    const capped = total > 0 ? Math.min(nextUsed, total) : nextUsed;
    saveOrder({ ...order, hireUsed: capped }, false);
  }
  function quickSetTradeStatus(id, status) {
    setData(d => ({ ...d, orders: d.orders.map(o => o.id === id ? { ...o, tradeStatus: status } : o) }));
    showToast(TRADE_STATUS[status] ? `อัปเดตเป็น "${TRADE_STATUS[status].label}"` : "อัปเดตสถานะเทรดแล้ว");
  }
  function quickSetHireStatus(id, status) {
    setData(d => ({ ...d, orders: d.orders.map(o => o.id === id ? { ...o, hireStatus: status } : o) }));
    showToast(HIRE_STATUS[status] ? `อัปเดตเป็น "${HIRE_STATUS[status].label}"` : "อัปเดตสถานะแล้ว");
  }
  function cancelOrder(id, reason) {
    setData(d => {
      const order = d.orders.find(o => o.id === id);
      if (!order || order.cancelled) return d;
      let gameAccounts = d.gameAccounts;
      if (order.type === "sell_pokemon" && order.stockItemId) {
        gameAccounts = adjustStock(gameAccounts, order.sourceAccountId, order.stockItemId, clamp0(order.quantity));
      }
      const historyEntry = { id: genId(), date: new Date().toISOString(), reason: (reason || "").trim() || "ไม่ระบุเหตุผล" };
      const orders = d.orders.map(o => o.id === id ? {
        ...o, cancelled: true, cancelledAt: new Date().toISOString(),
        cancelHistory: [historyEntry, ...(o.cancelHistory || [])],
      } : o);
      return { ...d, orders, gameAccounts };
    });
    showToast("ยกเลิกออเดอร์แล้ว");
  }
  // used by quick "ยกเลิก" buttons outside the order form — asks the reason first via a native prompt
  function promptCancelOrder(id) {
    const reason = window.prompt("ยกเลิกออเดอร์นี้เพราะอะไร?");
    if (reason === null) return; // user pressed cancel on the prompt itself — abort, don't cancel the order
    cancelOrder(id, reason);
  }
  function restoreOrder(id) {
    setData(d => {
      const order = d.orders.find(o => o.id === id);
      if (!order || !order.cancelled) return d;
      let gameAccounts = d.gameAccounts;
      if (order.type === "sell_pokemon" && order.stockItemId) {
        gameAccounts = adjustStock(gameAccounts, order.sourceAccountId, order.stockItemId, -clamp0(order.quantity));
      }
      const orders = d.orders.map(o => o.id === id ? { ...o, cancelled: false, cancelledAt: null } : o);
      return { ...d, orders, gameAccounts };
    });
    showToast("กู้คืนออเดอร์แล้ว");
  }
  function deleteOrder(id) {
    setData(d => {
      const order = d.orders.find(o => o.id === id);
      let gameAccounts = d.gameAccounts;
      if (order && !order.cancelled && order.type === "sell_pokemon" && order.stockItemId) {
        gameAccounts = adjustStock(gameAccounts, order.sourceAccountId, order.stockItemId, clamp0(order.quantity));
      }
      return { ...d, orders: d.orders.filter(o => o.id !== id), gameAccounts };
    });
    showToast("ลบออเดอร์ถาวรแล้ว");
  }
  function saveStock(accountId, item) {
    setData(d => ({
      ...d,
      gameAccounts: d.gameAccounts.map(a => {
        if (a.id !== accountId) return a;
        const stock = a.stock || [];
        const exists = stock.some(s => s.id === item.id);
        return { ...a, stock: exists ? stock.map(s => s.id === item.id ? item : s) : [item, ...stock] };
      }),
    }));
    showToast("บันทึกสต๊อกแล้ว");
  }
  function deleteStock(accountId, stockId) {
    setData(d => ({
      ...d,
      gameAccounts: d.gameAccounts.map(a => a.id === accountId ? { ...a, stock: (a.stock || []).filter(s => s.id !== stockId) } : a),
    }));
    showToast("ลบสต๊อกแล้ว");
  }
  function saveInvestment(item) {
    setData(d => ({ ...d, investmentHistory: [item, ...d.investmentHistory] }));
    showToast("บันทึกรายการลงทุนแล้ว");
  }
  function deleteInvestment(id) {
    setData(d => ({ ...d, investmentHistory: d.investmentHistory.filter(h => h.id !== id) }));
  }
  function saveManualTx(item) {
    setData(d => ({ ...d, manualTx: [item, ...d.manualTx] }));
    showToast("บันทึกรายการแล้ว");
  }
  function deleteManualTx(id) {
    setData(d => ({ ...d, manualTx: d.manualTx.filter(t => t.id !== id) }));
  }

  // ---------- export ----------
  function exportBackup() {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `pgs-backup-${todayStr()}.json`; a.click();
    URL.revokeObjectURL(url);
    setData(d => ({ ...d, settings: { ...d.settings, lastBackupAt: new Date().toISOString() } }));
    showToast("ดาวน์โหลด Backup แล้ว");
  }
  function restoreBackup(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        setData(migrateData(parsed));
        showToast("กู้คืนข้อมูลสำเร็จ");
      } catch { showToast("ไฟล์ไม่ถูกต้อง"); }
    };
    reader.readAsText(file);
    e.target.value = "";
  }

  // ---------- Google Sheets + Drive sync ----------
  async function connectGoogle(clientId) {
    setGoogleSyncing(true);
    setGoogleStatus("กำลังเชื่อมต่อ Google...");
    try {
      const token = await requestAccessToken(clientId, { interactive: true });
      const profile = await fetchGoogleProfile(token);
      setGoogleStatus("กำลังสร้าง Sheet และโฟลเดอร์...");
      const spreadsheetId = await ensureSpreadsheet(token, data.settings.google.spreadsheetId, data.settings.shopName);
      const folderId = await ensureDriveFolder(token, data.settings.google.folderId, data.settings.shopName);
      setData(d => ({ ...d, settings: { ...d.settings, google: { ...d.settings.google, clientId, email: profile?.email || "", spreadsheetId, folderId } } }));
      showToast("เชื่อมต่อ Google สำเร็จ");
      await runSync(token, spreadsheetId, folderId);
    } catch (e) {
      console.error("connectGoogle failed", e);
      setGoogleStatus("");
      showToast(e?.message || "เชื่อมต่อ Google ไม่สำเร็จ");
    } finally {
      setGoogleSyncing(false);
    }
  }
  async function runSync(tokenArg, spreadsheetIdArg, folderIdArg) {
    const g = data.settings.google;
    const spreadsheetId = spreadsheetIdArg || g.spreadsheetId;
    const folderId = folderIdArg || g.folderId;
    if (!spreadsheetId || !folderId) return;
    setGoogleSyncing(true);
    try {
      const token = tokenArg || await requestAccessToken(g.clientId, { interactive: false });
      const result = await syncAll({ token, spreadsheetId, folderId, data, onStatus: setGoogleStatus });
      setData(d => ({ ...d, orders: result.orders, settings: { ...d.settings, google: { ...d.settings.google, lastSyncAt: new Date().toISOString() } } }));
      setGoogleStatus("ซิงค์ล่าสุดสำเร็จ");
    } catch (e) {
      console.error("sync failed", e);
      setGoogleStatus(e?.message || "ซิงค์ไม่สำเร็จ ลองใหม่อีกครั้ง");
    } finally {
      setGoogleSyncing(false);
    }
  }
  function disconnectGoogleAccount() {
    disconnectGoogle();
    setData(d => ({ ...d, settings: { ...d.settings, google: { clientId: d.settings.google.clientId, email: "", spreadsheetId: "", folderId: "", autoSync: true, lastSyncAt: null } } }));
    setGoogleStatus("");
    showToast("ยกเลิกการเชื่อมต่อ Google แล้ว (ข้อมูลใน Sheet เดิมยังอยู่)");
  }

  // used by LockScreen's "ลืม PIN?" -> security question flow — the data already lives in this
  // browser's localStorage, so a correct answer just needs to swap the PIN, nothing else changes.
  function resetPinByAnswer(newPin) {
    setData(d => ({ ...d, settings: { ...d.settings, pin: newPin } }));
    setUnlocked(true);
    showToast("ตั้ง PIN ใหม่แล้ว");
  }
  // used by LockScreen's "ลืม PIN?" flow — restores data from an uploaded backup
  // file and applies a freshly-chosen PIN (or removes the lock), without ever
  // touching the rest of localStorage.
  function recoverFromBackup(parsed, chosenPin) {
    const migrated = migrateData(parsed);
    migrated.settings.pin = chosenPin || "";
    setData(migrated);
    setUnlocked(true);
    showToast(chosenPin ? "กู้คืนข้อมูลและตั้ง PIN ใหม่แล้ว" : "กู้คืนข้อมูลแล้ว ปิดการล็อก PIN แล้ว");
  }
  function exportExcel() {
    const wb = XLSX.utils.book_new();
    const ordersSheet = data.orders.map(o => ({
      รหัสออเดอร์: o.code, ลูกค้า: custName(o.customerId), ไอดีที่ใช้: o.customerGameId || "", ประเภท: ORDER_TYPES[o.type]?.label,
      Pokemon: o.pokemonName || "", ประเภทพิเศษ: (o.pokemonVariants || []).map(v => POKEMON_VARIANTS[v]?.label).filter(Boolean).join(", "),
      จำนวน: o.quantity || "", ราคา: o.price, ชำระแล้ว: o.paymentStatus === "paid" ? o.price : (o.paidAmount || 0),
      คงค้าง: orderBalance(o),
      ไอดีต้นทาง: o.sourceAccountId ? accName(o.sourceAccountId) : "", สถานะชำระ: PAYMENT_STATUS[o.paymentStatus]?.label,
      สถานะเทรด: o.tradeStatus ? TRADE_STATUS[o.tradeStatus]?.label : "",
      โหมดตี: (o.type !== "sell_pokemon") ? HIRE_MODES[o.hireMode]?.label : "",
      จำนวนรอบ: (o.rounds || []).length || "",
      สถานะออเดอร์: o.cancelled ? "ยกเลิกแล้ว" : "ปกติ",
      วันนัด: o.appointmentDate || "",
      หมายเหตุ: o.note || "", วันที่สร้าง: (o.createdAt || "").slice(0, 10),
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ordersSheet), "Orders");
    const custSheet = data.customers.map(c => ({ ชื่อในเกม: c.name, ไอดีในเกมทั้งหมด: (c.gameIds || []).map(g => g.value).filter(Boolean).join(", "), Facebook: c.facebook || "", หมายเหตุ: c.note || "" }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(custSheet), "Customers");
    const accSheet = data.gameAccounts.map(a => ({ ชื่อไอดี: a.name, จำนวนรายการสต๊อก: (a.stock || []).length, หมายเหตุ: a.note || "" }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(accSheet), "GameAccounts");
    const stockRows = [];
    data.gameAccounts.forEach(a => (a.stock || []).forEach(s => stockRows.push({
      ไอดี: a.name, Pokemon: s.name, ประเภทพิเศษ: (s.variants || []).map(v => POKEMON_VARIANTS[v]?.label).filter(Boolean).join(", "),
      คงเหลือ: s.quantity, แจ้งเตือนเมื่อเหลือ: s.lowStockThreshold ?? 2,
    })));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stockRows), "Stock");
    const invSheet = data.investmentHistory.map(h => ({ ไอดี: accName(h.accountId), ประเภท: INVEST_TYPES[h.type]?.label, จำนวนเงิน: h.amount, วันที่: h.date, หมายเหตุ: h.note || "" }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(invSheet), "InvestmentHistory");
    const manualSheet = data.manualTx.map(t => ({
      ประเภท: t.type === "income" ? "รายรับ" : "รายจ่าย", รายการ: t.category || "อื่นๆ", จำนวนเงิน: t.amount,
      วันที่: t.date, ไอดีที่เกี่ยวข้อง: t.accountId ? accName(t.accountId) : "", หมายเหตุ: t.note || "",
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(manualSheet), "ManualTransactions");
    XLSX.writeFile(wb, `pgs-export-${todayStr()}.xlsx`);
    showToast("Export Excel แล้ว");
  }
  function exportPDF() { window.print(); }

  if (!loaded) {
    return (
      <div className="pgs-root" style={{ alignItems: "center", justifyContent: "center" }}>
        <GlobalStyle />
        <div className="pgs-ball" style={{ animation: "pgs-up 1s infinite alternate" }} />
      </div>
    );
  }

  if (data.settings.pin && !unlocked) {
    return (
      <LockScreen
        pin={data.settings.pin}
        pinQuestion={data.settings.pinQuestion}
        pinAnswer={data.settings.pinAnswer}
        shopName={data.settings.shopName}
        logoDataUrl={data.settings.logoDataUrl}
        onUnlock={() => setUnlocked(true)}
        onRecover={recoverFromBackup}
        onResetPin={resetPinByAnswer}
      />
    );
  }

  return (
    <div className="pgs-root">
      <GlobalStyle />
      <Header data={data} onMore={() => setMoreOpen(true)} />
      <div className="pgs-scroll">
        {tab === "dashboard" && <Dashboard data={data} stats={stats} custName={custName} accName={accName} goTab={setTab} openDetail={(d) => setDetail(d)} />}
        {tab === "orders" && <OrdersTab data={data} custName={custName} accName={accName} openNew={() => setModal({ type: "order", mode: "add" })} openEdit={(o) => setModal({ type: "order", mode: "edit", item: o })} openReceipt={(o) => setModal({ type: "receipt", item: o })} onQuickComplete={quickComplete} onQuickCancel={promptCancelOrder} />}
        {tab === "trade" && <TradeTab data={data} custName={custName} accName={accName} openEdit={(o) => setModal({ type: "order", mode: "edit", item: o })} onQuickTrade={quickSetTradeStatus} />}
        {tab === "hire" && <HireTab data={data} custName={custName} accName={accName} openEdit={(o) => setModal({ type: "order", mode: "edit", item: o })} onQuickUse={quickUseHire} onQuickHireStatus={quickSetHireStatus} />}
      </div>
      {moreOpen && (
        <MoreSheet
          onClose={() => setMoreOpen(false)}
          go={(t) => { setTab(t); setMoreOpen(false); }}
        />
      )}
      {tab === "customers" && (
        <div className="pgs-scroll" style={{ position: "fixed", inset: "60px 0 74px 0", maxWidth: 430, margin: "0 auto", background: "var(--bg)", zIndex: 40 }}>
          <CustomersTab
            data={data}
            openNew={() => setModal({ type: "customer", mode: "add" })}
            openEdit={(c) => setModal({ type: "customer", mode: "edit", item: c })}
            openDetail={(c) => setDetail({ type: "customer", item: c })}
            back={() => setTab("dashboard")}
          />
        </div>
      )}
      {tab === "accounts" && (
        <div className="pgs-scroll" style={{ position: "fixed", inset: "60px 0 74px 0", maxWidth: 430, margin: "0 auto", background: "var(--bg)", zIndex: 40 }}>
          <AccountsTab
            data={data} stats={stats}
            openNew={() => setModal({ type: "account", mode: "add" })}
            openDetail={(a) => setDetail({ type: "account", item: a })}
            back={() => setTab("dashboard")}
          />
        </div>
      )}
      {tab === "finance" && (
        <div className="pgs-scroll" style={{ position: "fixed", inset: "60px 0 74px 0", maxWidth: 430, margin: "0 auto", background: "var(--bg)", zIndex: 40 }}>
          <FinanceTab data={data} stats={stats} custName={custName} accName={accName} openNew={() => setModal({ type: "tx", mode: "add" })} back={() => setTab("dashboard")} onDeleteManual={deleteManualTx} openDetail={(d) => setDetail(d)} />
        </div>
      )}
      {tab === "reports" && (
        <div className="pgs-scroll" style={{ position: "fixed", inset: "60px 0 74px 0", maxWidth: 430, margin: "0 auto", background: "var(--bg)", zIndex: 40 }}>
          <ReportsTab data={data} custName={custName} accName={accName} back={() => setTab("dashboard")} />
        </div>
      )}
      {tab === "settings" && (
        <div className="pgs-scroll" style={{ position: "fixed", inset: "60px 0 74px 0", maxWidth: 430, margin: "0 auto", background: "var(--bg)", zIndex: 40 }}>
          <SettingsTab
            data={data} setData={setData} onBackup={exportBackup} onRestore={restoreBackup} onExportExcel={exportExcel} onExportPDF={exportPDF}
            showToast={showToast} back={() => setTab("dashboard")}
            googleSyncing={googleSyncing} googleStatus={googleStatus}
            onConnectGoogle={connectGoogle} onSyncNow={() => runSync()} onDisconnectGoogle={disconnectGoogleAccount}
          />
        </div>
      )}

      <BottomNav tab={tab} setTab={setTab} onMore={() => setMoreOpen(true)} />

      {modal?.type === "order" && (
        <OrderModal
          data={data} mode={modal.mode} item={modal.item}
          onClose={() => setModal(null)}
          onSave={(item) => { saveOrder(item, modal.mode === "add"); setModal(null); }}
          onCancel={(id, reason) => { cancelOrder(id, reason); setModal(null); }}
          onRestore={(id) => { restoreOrder(id); setModal(null); }}
          onDelete={(id) => { deleteOrder(id); setModal(null); }}
          onReceipt={(o) => setModal({ type: "receipt", item: o })}
        />
      )}
      {modal?.type === "receipt" && (
        <ReceiptModal
          order={modal.item} data={data} custName={custName} accName={accName}
          onClose={() => setModal(null)}
          onToast={showToast}
        />
      )}
      {modal?.type === "stock" && (
        <StockModal
          mode={modal.mode} item={modal.item}
          onClose={() => setModal(null)}
          onSave={(item) => { saveStock(modal.accountId, item); setModal(null); }}
          onDelete={modal.mode === "edit" ? () => { deleteStock(modal.accountId, modal.item.id); setModal(null); } : null}
        />
      )}
      {modal?.type === "customer" && (
        <CustomerModal
          mode={modal.mode} item={modal.item}
          onClose={() => setModal(null)}
          onSave={(item) => { saveCustomer(item); setModal(null); }}
        />
      )}
      {modal?.type === "account" && (
        <AccountModal
          mode={modal.mode} item={modal.item}
          onClose={() => setModal(null)}
          onSave={(item) => { saveAccount(item); setModal(null); }}
        />
      )}
      {modal?.type === "tx" && (
        <TxModal
          data={data}
          onClose={() => setModal(null)}
          onSaveInvestment={(item) => { saveInvestment(item); setModal(null); }}
          onSaveManual={(item) => { saveManualTx(item); setModal(null); }}
        />
      )}

      {detail?.type === "debt" && (
        <DebtModal
          data={data} custName={custName}
          onClose={() => setDetail(null)}
          onOpenCustomer={(c) => setDetail({ type: "customer", item: c })}
        />
      )}
      {showBackupPrompt && (
        <Modal title="ยังไม่ได้ Backup ข้อมูล" onClose={() => setShowBackupPrompt(false)}>
          <div style={{ display: "flex", gap: 10, alignItems: "flex-start", marginBottom: 16 }}>
            <AlertTriangle size={22} color="var(--yellow)" style={{ flexShrink: 0, marginTop: 2 }} />
            <div style={{ fontSize: 13, color: "var(--text)", lineHeight: 1.5 }}>
              {data.settings.lastBackupAt
                ? `ไม่ได้ Backup ข้อมูลมา ${daysBetween(data.settings.lastBackupAt, new Date().toISOString())} วันแล้ว ข้อมูลทั้งหมดเก็บอยู่ในเครื่องนี้เครื่องเดียว หากล้าง cache หรือเปลี่ยนเครื่องโดยไม่ Backup ไว้ ข้อมูลจะหายถาวร`
                : "ยังไม่เคย Backup ข้อมูลเลย ข้อมูลทั้งหมดเก็บอยู่ในเครื่องนี้เครื่องเดียว แนะนำให้ Backup ไว้กันข้อมูลหาย"}
            </div>
          </div>
          <button className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: 8 }} onClick={() => { exportBackup(); setShowBackupPrompt(false); }}>
            <Download size={15} /> Backup ตอนนี้
          </button>
          <button className="pgs-btn pgs-btn-outline" style={{ width: "100%" }} onClick={() => setShowBackupPrompt(false)}>เตือนทีหลัง</button>
        </Modal>
      )}
      {detail?.type === "duesoon" && (
        <DueSoonModal
          items={stats.dueSoonItems} data={data} custName={custName}
          onClose={() => setDetail(null)}
          onGoTo={(kind) => { setTab(kind === "round" ? "hire" : "trade"); setDetail(null); }}
        />
      )}
      {detail?.type === "customer" && (
        <CustomerDetail
          item={detail.item} data={data}
          onClose={() => setDetail(null)}
          onEdit={() => { setModal({ type: "customer", mode: "edit", item: detail.item }); setDetail(null); }}
          onDelete={() => { deleteCustomer(detail.item.id); setDetail(null); }}
        />
      )}
      {detail?.type === "account" && (
        <AccountDetail
          item={data.gameAccounts.find(a => a.id === detail.item.id) || detail.item} data={data} stats={stats}
          onClose={() => setDetail(null)}
          onEdit={() => { setModal({ type: "account", mode: "edit", item: detail.item }); setDetail(null); }}
          onDelete={() => { deleteAccount(detail.item.id); setDetail(null); }}
          onAddInvestment={() => { setModal({ type: "tx", mode: "add", presetAccount: detail.item.id }); setDetail(null); }}
          onDeleteInvestment={deleteInvestment}
          onAddStock={() => setModal({ type: "stock", mode: "add", accountId: detail.item.id })}
          onEditStock={(s) => setModal({ type: "stock", mode: "edit", item: s, accountId: detail.item.id })}
        />
      )}

      {toast && <div className="pgs-toast">{toast}</div>}
    </div>
  );
}

// =================== IMAGE CROPPER (drag to reposition, zoom to scale, like a profile-photo picker) ===================
// aspect = width / height of the output frame. shape "circle" shows a round mask (for the shop logo),
// shape "rect" shows a rounded-rect mask (for the receipt background). Output is always baked to a
// plain rectangle at outputW x outputH, so it can be used as a normal <img>/dataURL anywhere.
function ImageCropModal({ src, aspect = 1, shape = "rect", outputW = 640, title = "ปรับตำแหน่งรูปภาพ", onCancel, onConfirm }) {
  const frameW = Math.min(220, Math.round(280 * aspect));
  const frameH = Math.round(frameW / aspect);
  const outputH = Math.round(outputW / aspect);

  const [natural, setNatural] = useState(null); // { w, h }
  const [zoom, setZoom] = useState(1); // 1 = just covers the frame
  const [pos, setPos] = useState({ x: 0, y: 0 }); // top-left of image relative to frame, css px
  const dragState = useRef(null);

  useEffect(() => {
    let alive = true;
    const img = new Image();
    img.onload = () => {
      if (!alive) return;
      setNatural({ w: img.naturalWidth, h: img.naturalHeight });
      setZoom(1);
      setPos({ x: 0, y: 0 }); // centered once we know size, via effect below
    };
    img.src = src;
    return () => { alive = false; };
  }, [src]);

  const minScale = natural ? Math.max(frameW / natural.w, frameH / natural.h) : 1;
  const scale = minScale * zoom;
  const dw = natural ? natural.w * scale : 0;
  const dh = natural ? natural.h * scale : 0;

  // center the image the first time we learn its size
  useEffect(() => {
    if (!natural) return;
    setPos({ x: (frameW - natural.w * minScale) / 2, y: (frameH - natural.h * minScale) / 2 });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [natural]);

  function clampPos(x, y, curDw = dw, curDh = dh) {
    const minX = Math.min(0, frameW - curDw);
    const minY = Math.min(0, frameH - curDh);
    return { x: Math.min(0, Math.max(minX, x)), y: Math.min(0, Math.max(minY, y)) };
  }

  function onZoomChange(e) {
    const nextZoom = Number(e.target.value);
    const nextScale = minScale * nextZoom;
    const nextDw = natural ? natural.w * nextScale : 0;
    const nextDh = natural ? natural.h * nextScale : 0;
    // keep the frame's current center point anchored while zooming
    const cx = pos.x - frameW / 2, cy = pos.y - frameH / 2;
    const ratio = nextScale / scale;
    setPos(clampPos(frameW / 2 + cx * ratio, frameH / 2 + cy * ratio, nextDw, nextDh));
    setZoom(nextZoom);
  }

  function onPointerDown(e) {
    e.currentTarget.setPointerCapture(e.pointerId);
    dragState.current = { startX: e.clientX, startY: e.clientY, origX: pos.x, origY: pos.y };
  }
  function onPointerMove(e) {
    if (!dragState.current) return;
    const dx = e.clientX - dragState.current.startX;
    const dy = e.clientY - dragState.current.startY;
    setPos(clampPos(dragState.current.origX + dx, dragState.current.origY + dy));
  }
  function onPointerUp() { dragState.current = null; }

  function confirm() {
    if (!natural) return;
    const outScale = outputW / frameW;
    const canvas = document.createElement("canvas");
    canvas.width = outputW; canvas.height = outputH;
    const ctx = canvas.getContext("2d");
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = "high";
    const img = new Image();
    img.onload = () => {
      ctx.drawImage(img, pos.x * outScale, pos.y * outScale, dw * outScale, dh * outScale);
      onConfirm(canvas.toDataURL("image/png", 0.9));
    };
    img.src = src;
  }

  function stepZoom(delta) {
    const next = Math.min(3, Math.max(1, Math.round((zoom + delta) * 100) / 100));
    onZoomChange({ target: { value: String(next) } });
  }

  return (
    <Modal title={title} onClose={onCancel}>
      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 14 }}>
        <div
          style={{
            width: frameW, height: frameH, position: "relative", overflow: "hidden",
            borderRadius: shape === "circle" ? "50%" : 16,
            border: "2px solid var(--yellow)", background: "#0c0d15",
            touchAction: "none", cursor: dragState.current ? "grabbing" : "grab",
          }}
          onPointerDown={onPointerDown}
          onPointerMove={onPointerMove}
          onPointerUp={onPointerUp}
          onPointerLeave={onPointerUp}
        >
          {natural && (
            <img
              src={src}
              alt=""
              draggable={false}
              style={{ position: "absolute", left: pos.x, top: pos.y, width: dw, height: dh, maxWidth: "none", userSelect: "none", pointerEvents: "none" }}
            />
          )}
          {natural && (
            <div
              style={{
                position: "absolute", top: 10, left: "50%", transform: "translateX(-50%)",
                display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap",
                background: "rgba(12,13,21,0.82)", color: "#fff", fontSize: 11, fontWeight: 600,
                padding: "6px 12px", borderRadius: 999, pointerEvents: "none",
                boxShadow: "0 4px 14px rgba(0,0,0,0.35)",
              }}
            >
              <Move size={13} /> ลากรูปเพื่อปรับตำแหน่ง
            </div>
          )}
        </div>

        <div style={{ width: "100%", display: "flex", alignItems: "center", gap: 10 }}>
          <button type="button" className="pgs-iconbtn" style={{ flexShrink: 0 }} onClick={() => stepZoom(-0.1)} disabled={zoom <= 1}>
            <Minus size={14} />
          </button>
          <input type="range" min="1" max="3" step="0.01" value={zoom} onChange={onZoomChange} style={{ flex: 1 }} />
          <button type="button" className="pgs-iconbtn" style={{ flexShrink: 0 }} onClick={() => stepZoom(0.1)} disabled={zoom >= 3}>
            <Plus size={14} />
          </button>
        </div>

        <div
          style={{
            display: "flex", gap: 8, width: "100%",
            position: "sticky", bottom: 0, left: 0,
            background: "linear-gradient(180deg, #14151f 0%, var(--bg) 100%)",
            borderTop: "1px solid var(--border)",
            padding: "12px 0 calc(10px + env(safe-area-inset-bottom))", marginTop: 4,
            zIndex: 1,
          }}
        >
          <button type="button" className="pgs-btn pgs-btn-outline" style={{ flex: 1 }} onClick={onCancel}>ยกเลิก</button>
          <button type="button" className="pgs-btn pgs-btn-primary" style={{ flex: 1 }} disabled={!natural} onClick={confirm}>ใช้รูปนี้</button>
        </div>
      </div>
    </Modal>
  );
}

// =================== HEADER / NAV ===================
function ShopLogo({ logoDataUrl, size = 30 }) {
  if (logoDataUrl) {
    return <img src={logoDataUrl} alt="โลโก้ร้าน" style={{ width: size, height: size, borderRadius: "50%", objectFit: "cover", flexShrink: 0, border: "2px solid #14151f" }} />;
  }
  return <div className="pgs-ball" style={{ width: size, height: size }} />;
}

function Header({ data, onMore }) {
  return (
    <div className="pgs-header">
      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <ShopLogo logoDataUrl={data.settings.logoDataUrl} />
        <div>
          <div className="pgs-display" style={{ fontSize: 15, fontWeight: 700, lineHeight: 1.1 }}>{data.settings.shopName}</div>
          <div style={{ fontSize: 10, color: "var(--muted)" }}>ระบบจัดการร้าน</div>
        </div>
      </div>
      <button className="pgs-btn pgs-btn-outline" style={{ padding: 8 }} onClick={onMore}><MoreHorizontal size={16} /></button>
    </div>
  );
}

function BottomNav({ tab, setTab, onMore }) {
  const items = [
    { id: "dashboard", label: "หน้าแรก", icon: Home },
    { id: "orders", label: "ออเดอร์", icon: Package },
    { id: "trade", label: "เทรด", icon: Repeat },
    { id: "hire", label: "ตีบอส/เชิญตี", icon: Target },
  ];
  return (
    <div className="pgs-bottomnav">
      {items.map(it => (
        <button key={it.id} className={"pgs-navitem" + (tab === it.id ? " active" : "")} onClick={() => setTab(it.id)}>
          <it.icon size={19} />
          {it.label}
        </button>
      ))}
      <button className={"pgs-navitem" + (["accounts", "finance", "reports", "settings", "customers"].includes(tab) ? " active" : "")} onClick={onMore}>
        <MoreHorizontal size={19} />
        เพิ่มเติม
      </button>
    </div>
  );
}

function MoreSheet({ onClose, go }) {
  const items = [
    { id: "customers", label: "ลูกค้า", icon: Users, desc: "รายชื่อ & ยอดใช้จ่ายสะสม" },
    { id: "accounts", label: "ไอดีเกม", icon: Gamepad2, desc: "จัดการไอดี & เงินลงทุน" },
    { id: "finance", label: "การเงิน", icon: Wallet, desc: "รายรับ-รายจ่ายทั้งหมด" },
    { id: "reports", label: "รายงาน", icon: BarChart3, desc: "สรุปผลประกอบการ" },
    { id: "settings", label: "ตั้งค่า", icon: SettingsIcon, desc: "ร้าน, Backup, Export" },
  ];
  return (
    <div className="pgs-overlay" onClick={onClose}>
      <div className="pgs-sheet" onClick={(e) => e.stopPropagation()}>
        <div className="pgs-row" style={{ marginBottom: 14 }}>
          <h3 className="pgs-display" style={{ fontSize: 18, fontWeight: 700, margin: 0 }}>เมนูเพิ่มเติม</h3>
          <button className="pgs-btn pgs-btn-outline" style={{ padding: 8 }} onClick={onClose}><X size={16} /></button>
        </div>
        {items.map(it => (
          <button key={it.id} onClick={() => go(it.id)} className="pgs-card" style={{ display: "flex", alignItems: "center", gap: 12, width: "100%", marginBottom: 10, cursor: "pointer", textAlign: "left" }}>
            <div style={{ background: "var(--surface2)", borderRadius: 12, padding: 10 }}><it.icon size={18} color="var(--yellow)" /></div>
            <div style={{ flex: 1 }}>
              <div style={{ fontWeight: 700, fontSize: 14 }}>{it.label}</div>
              <div style={{ fontSize: 11, color: "#9ca3af" }}>{it.desc}</div>
            </div>
            <ChevronRight size={16} color="var(--muted)" />
          </button>
        ))}
      </div>
    </div>
  );
}

function SubHeader({ title, back }) {
  return (
    <div className="pgs-row" style={{ marginBottom: 14 }}>
      <button className="pgs-btn pgs-btn-outline" style={{ padding: 8 }} onClick={back}><ArrowLeft size={16} /></button>
      <h2 className="pgs-display" style={{ fontSize: 18, fontWeight: 700, margin: 0 }}>{title}</h2>
      <div style={{ width: 34 }} />
    </div>
  );
}

// =================== DASHBOARD ===================
const PERIODS = {
  today: { label: "วันนี้" },
  month: { label: "เดือนนี้" },
  year: { label: "ปีนี้" },
};

function Dashboard({ data, stats, custName, accName, goTab, openDetail }) {
  const [period, setPeriod] = useState("today");
  const recentOrders = data.orders.filter(o => !o.cancelled).slice(0, 4);

  const income = period === "today" ? stats.incomeToday : period === "month" ? stats.incomeMonth : stats.incomeYear;
  const expense = period === "today" ? stats.expenseToday : period === "month" ? stats.expenseMonth : stats.expenseYear;
  const profit = period === "today" ? stats.profitToday : period === "month" ? stats.profitMonth : stats.profitYear;

  const daysSinceBackup = data.settings.lastBackupAt ? daysBetween(data.settings.lastBackupAt, new Date().toISOString()) : null;
  const needsBackup = daysSinceBackup === null ? data.orders.length + data.customers.length > 0 : daysSinceBackup >= 7;

  return (
    <div>
      {needsBackup && (
        <button onClick={() => goTab("settings")} className="pgs-card" style={{ marginBottom: 12, width: "100%", textAlign: "left", cursor: "pointer", display: "flex", alignItems: "center", gap: 10, borderColor: "rgba(255,203,5,0.4)" }}>
          <Download size={18} color="var(--yellow)" />
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 700, fontSize: 13, color: "var(--yellow)" }}>{daysSinceBackup === null ? "ยังไม่เคย Backup ข้อมูล" : `ไม่ได้ Backup มา ${daysSinceBackup} วันแล้ว`}</div>
            <div style={{ fontSize: 10, color: "var(--muted)" }}>ข้อมูลอยู่ในเครื่องนี้เครื่องเดียว แตะเพื่อไปหน้าตั้งค่า</div>
          </div>
          <ChevronRight size={16} color="var(--muted)" />
        </button>
      )}
      <div className="pgs-row" style={{ marginBottom: 10 }}>
        <div className="pgs-sectiontitle" style={{ margin: 0 }}>ภาพรวม</div>
        <div style={{ display: "flex", gap: 4 }}>
          {Object.entries(PERIODS).map(([k, v]) => (
            <button key={k} className={"pgs-chip" + (period === k ? " active" : "")} onClick={() => setPeriod(k)}>{v.label}</button>
          ))}
        </div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <StatCard icon={TrendingUp} label={"รายรับ" + PERIODS[period].label} value={"฿" + fmtMoney(income)} color="var(--green)" />
        <StatCard icon={TrendingDown} label={"รายจ่าย" + PERIODS[period].label} value={"฿" + fmtMoney(expense)} color="var(--red)" />
      </div>

      <div className="pgs-sectiontitle">สรุปกำไร</div>
      <div className="pgs-card" style={{ marginBottom: 4 }}>
        <div className="pgs-row">
          <span style={{ fontSize: 12, color: "var(--muted)" }}>กำไรสุทธิ ({PERIODS[period].label})</span>
          <span className="pgs-mono pgs-display" style={{ fontSize: 22, fontWeight: 700, color: profit >= 0 ? "var(--green)" : "var(--red)" }}>฿{fmtMoney(profit)}</span>
        </div>
      </div>

      <div className="pgs-sectiontitle">เงินลงทุน</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <StatCard icon={Coins} label="ลงทุนทั้งหมด" value={"฿" + fmtMoney(stats.totalInvestment)} color="var(--yellow)" />
        <StatCard icon={Package} label="ออเดอร์ทั้งหมด" value={stats.totalOrders} />
      </div>

      <div className="pgs-sectiontitle">รอดำเนินการ</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8 }}>
        <button onClick={() => goTab("orders")} className="pgs-statcard" style={{ cursor: "pointer", textAlign: "left" }}>
          <div style={{ fontSize: 10, color: "var(--muted)", marginBottom: 4 }}>รอชำระ/ค้าง</div>
          <div className="pgs-mono" style={{ fontSize: 18, fontWeight: 700, color: "var(--red)" }}>{stats.pendingPayment}</div>
        </button>
        <button onClick={() => goTab("trade")} className="pgs-statcard" style={{ cursor: "pointer", textAlign: "left" }}>
          <div style={{ fontSize: 10, color: "var(--muted)", marginBottom: 4 }}>รอเทรด</div>
          <div className="pgs-mono" style={{ fontSize: 18, fontWeight: 700, color: "var(--muted)" }}>{stats.pendingTrade}</div>
        </button>
        <button onClick={() => goTab("trade")} className="pgs-statcard" style={{ cursor: "pointer", textAlign: "left" }}>
          <div style={{ fontSize: 10, color: "var(--muted)", marginBottom: 4 }}>ทำ 3 ใจ</div>
          <div className="pgs-mono" style={{ fontSize: 18, fontWeight: 700, color: "var(--yellow)" }}>{stats.threeHearts}</div>
        </button>
      </div>

      {stats.totalDue > 0 && (
        <button onClick={() => openDetail({ type: "debt" })} className="pgs-card" style={{ marginTop: 10, borderColor: "rgba(255,84,112,0.4)", width: "100%", textAlign: "left", cursor: "pointer" }}>
          <div className="pgs-row">
            <span style={{ fontSize: 12, color: "var(--muted)" }}>ยอดค้างชำระรวม · แตะเพื่อดูรายลูกค้า</span>
            <span className="pgs-mono" style={{ fontWeight: 700, fontSize: 16, color: "var(--red)" }}>฿{fmtMoney(stats.totalDue)}</span>
          </div>
        </button>
      )}

      {stats.dueSoonCount > 0 && (
        <button onClick={() => openDetail({ type: "duesoon" })} className="pgs-card" style={{ marginTop: 10, width: "100%", textAlign: "left", cursor: "pointer", display: "flex", alignItems: "center", gap: 10, borderColor: "rgba(255,203,5,0.4)" }}>
          <Clock size={18} color="var(--yellow)" />
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 700, fontSize: 13, color: "var(--yellow)" }}>นัดหมาย/รอบตีใกล้ถึงกำหนด {stats.dueSoonCount} รายการ</div>
            <div style={{ fontSize: 10, color: "var(--muted)" }}>ภายใน 7 วัน · แตะเพื่อดูรายการ แล้วไปหน้าเทรด/ตีบอสได้เลย</div>
          </div>
          <ChevronRight size={16} color="var(--muted)" />
        </button>
      )}

      {stats.lowStockCount > 0 && (
        <button onClick={() => goTab("accounts")} className="pgs-card" style={{ marginTop: 10, width: "100%", textAlign: "left", cursor: "pointer", display: "flex", alignItems: "center", gap: 10, borderColor: "rgba(255,84,112,0.4)" }}>
          <AlertTriangle size={18} color="var(--red)" />
          <div style={{ flex: 1 }}>
            <div style={{ fontWeight: 700, fontSize: 13, color: "var(--red)" }}>สต๊อกใกล้หมด {stats.lowStockCount} รายการ</div>
            <div style={{ fontSize: 10, color: "var(--muted)" }}>แตะเพื่อดูไอดีเกม</div>
          </div>
          <ChevronRight size={16} color="var(--muted)" />
        </button>
      )}

      <div className="pgs-row" style={{ marginTop: 18, marginBottom: 8 }}>
        <div className="pgs-sectiontitle" style={{ margin: 0 }}>ออเดอร์ล่าสุด</div>
        <button onClick={() => goTab("orders")} style={{ background: "none", border: "none", color: "var(--yellow)", fontSize: 12, fontWeight: 600, cursor: "pointer" }}>ดูทั้งหมด</button>
      </div>
      {recentOrders.length === 0 ? <EmptyState text="ยังไม่มีออเดอร์" /> : recentOrders.map(o => (
        <div key={o.id} className="pgs-card" style={{ marginBottom: 8 }}>
          <div className="pgs-row">
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ fontSize: 18 }}>{ORDER_TYPES[o.type].emoji}</span>
              <div>
                <div style={{ fontWeight: 700, fontSize: 13 }}>{custName(o.customerId)}</div>
                <div className="pgs-mono" style={{ fontSize: 10, color: "var(--muted)" }}>{o.code}</div>
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div className="pgs-mono" style={{ fontWeight: 700, fontSize: 13 }}>฿{fmtMoney(o.price)}</div>
              <StatusDot payment={o.paymentStatus} />
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

// =================== ORDERS ===================
function OrdersTab({ data, custName, accName, openNew, openEdit, openReceipt, onQuickComplete, onQuickCancel }) {
  const [filter, setFilter] = useState("all");
  const [q, setQ] = useState("");
  const [showAdv, setShowAdv] = useState(false);
  const [paymentFilter, setPaymentFilter] = useState("all");
  const [tradeFilter, setTradeFilter] = useState("all");
  const [cancelFilter, setCancelFilter] = useState("active");
  const [sortBy, setSortBy] = useState("date_desc");

  const filtered = data.orders.filter(o => {
    if (cancelFilter === "active" && o.cancelled) return false;
    if (cancelFilter === "cancelled" && !o.cancelled) return false;
    if (filter !== "all" && o.type !== filter) return false;
    if (paymentFilter !== "all" && o.paymentStatus !== paymentFilter) return false;
    if (tradeFilter !== "all" && o.tradeStatus !== tradeFilter) return false;
    if (q && !(custName(o.customerId).toLowerCase().includes(q.toLowerCase()) || (o.pokemonName || "").toLowerCase().includes(q.toLowerCase()) || o.code.toLowerCase().includes(q.toLowerCase()))) return false;
    return true;
  }).sort((a, b) => {
    if (sortBy === "date_asc") return (a.createdAt || "").localeCompare(b.createdAt || "");
    if (sortBy === "amount_desc") return (Number(b.price) || 0) - (Number(a.price) || 0);
    if (sortBy === "amount_asc") return (Number(a.price) || 0) - (Number(b.price) || 0);
    return (b.createdAt || "").localeCompare(a.createdAt || ""); // date_desc (default)
  });

  return (
    <div>
      <div className="pgs-row" style={{ marginBottom: 12 }}>
        <h2 className="pgs-display" style={{ fontSize: 20, fontWeight: 700, margin: 0 }}>ออเดอร์</h2>
        <button className="pgs-btn pgs-btn-primary" onClick={openNew}><Plus size={15} /> เพิ่ม</button>
      </div>
      <div style={{ position: "relative", marginBottom: 10 }}>
        <Search size={14} color="var(--muted)" style={{ position: "absolute", left: 12, top: 12 }} />
        <input className="pgs-input" style={{ paddingLeft: 32 }} placeholder="ค้นหาลูกค้า, Pokémon, รหัส..." value={q} onChange={e => setQ(e.target.value)} />
      </div>
      <select className="pgs-select" style={{ marginBottom: 8, fontSize: 12 }} value={sortBy} onChange={e => setSortBy(e.target.value)}>
        <option value="date_desc">เรียง: ใหม่สุดก่อน</option>
        <option value="date_asc">เรียง: เก่าสุดก่อน</option>
        <option value="amount_desc">เรียง: ราคามาก-น้อย</option>
        <option value="amount_asc">เรียง: ราคาน้อย-มาก</option>
      </select>
      <div style={{ display: "flex", gap: 6, overflowX: "auto", marginBottom: 8, paddingBottom: 2 }}>
        <button className={"pgs-chip" + (filter === "all" ? " active" : "")} onClick={() => setFilter("all")}>ทั้งหมด</button>
        {Object.entries(ORDER_TYPES).map(([k, v]) => (
          <button key={k} className={"pgs-chip" + (filter === k ? " active" : "")} onClick={() => setFilter(k)}>{v.emoji} {v.short}</button>
        ))}
        <button className={"pgs-chip" + (showAdv ? " active" : "")} onClick={() => setShowAdv(s => !s)}><ListFilter size={12} style={{ verticalAlign: -2 }} /> ตัวกรอง</button>
      </div>
      {showAdv && (
        <div className="pgs-card" style={{ marginBottom: 12 }}>
          <div className="pgs-field" style={{ marginBottom: 10 }}>
            <label className="pgs-label">สถานะชำระ</label>
            <select className="pgs-select" value={paymentFilter} onChange={e => setPaymentFilter(e.target.value)}>
              <option value="all">ทั้งหมด</option>
              {Object.entries(PAYMENT_STATUS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
          </div>
          <div className="pgs-field" style={{ marginBottom: 10 }}>
            <label className="pgs-label">สถานะเทรด</label>
            <select className="pgs-select" value={tradeFilter} onChange={e => setTradeFilter(e.target.value)}>
              <option value="all">ทั้งหมด</option>
              {Object.entries(TRADE_STATUS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
            </select>
          </div>
          <div className="pgs-field" style={{ marginBottom: 0 }}>
            <label className="pgs-label">สถานะออเดอร์</label>
            <div style={{ display: "flex", gap: 6 }}>
              <button className={"pgs-chip" + (cancelFilter === "active" ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => setCancelFilter("active")}>ปกติ</button>
              <button className={"pgs-chip" + (cancelFilter === "cancelled" ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => setCancelFilter("cancelled")}>ยกเลิกแล้ว</button>
              <button className={"pgs-chip" + (cancelFilter === "all" ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => setCancelFilter("all")}>ทั้งหมด</button>
            </div>
          </div>
        </div>
      )}
      {filtered.length === 0 ? <EmptyState text="ไม่พบออเดอร์" /> : filtered.map(o => {
        const balance = orderBalance(o);
        return (
          <div key={o.id} className="pgs-card" style={{ marginBottom: 8, opacity: o.cancelled ? 0.7 : 1 }} onClick={() => openEdit(o)}>
            <div className="pgs-row" style={{ alignItems: "flex-start" }}>
              <div style={{ display: "flex", gap: 8 }}>
                <span style={{ fontSize: 18 }}>{ORDER_TYPES[o.type].emoji}</span>
                <div>
                  <div className={"pgs-row" + (o.cancelled ? " pgs-strike" : "")} style={{ gap: 6, justifyContent: "flex-start" }}>
                    <span style={{ fontWeight: 700, fontSize: 13 }}>{custName(o.customerId)}</span>
                  </div>
                  <div style={{ fontSize: 11, color: "var(--muted)" }}>
                    {o.type === "sell_pokemon"
                      ? `${o.pokemonName || ""}${(o.pokemonVariants || []).filter(v => v !== "normal").length ? " (" + o.pokemonVariants.filter(v => v !== "normal").map(v => POKEMON_VARIANTS[v]?.label).join(", ") + ")" : ""} x${o.quantity || 1}`
                      : `${ORDER_TYPES[o.type].label} · ${HIRE_MODES[o.hireMode]?.label || ""}`}
                    {o.sourceAccountId ? ` · ${accName(o.sourceAccountId)}` : ""}
                  </div>
                  <div className="pgs-mono" style={{ fontSize: 10, color: "var(--muted)", marginTop: 2 }}>{o.code} · {fmtDate(o.createdAt)}</div>
                </div>
              </div>
              <div style={{ textAlign: "right" }}>
                <div className="pgs-mono" style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>฿{fmtMoney(o.price)}</div>
                <StatusDot payment={o.paymentStatus} trade={o.type === "sell_pokemon" ? o.tradeStatus : null} cancelled={o.cancelled} />
              </div>
            </div>
            {!o.cancelled && o.paymentStatus === "partial" && (
              <div style={{ fontSize: 11, color: "var(--red)", marginTop: 6 }}>ค้างชำระ ฿{fmtMoney(balance)}</div>
            )}
            {!o.cancelled && o.type !== "sell_pokemon" && (
              <div className="pgs-row" style={{ marginTop: 8, background: "var(--surface2)", borderRadius: 10, padding: "6px 8px" }}>
                <span style={{ fontSize: 11, color: "var(--muted)" }}>ใช้ไปแล้ว <span className="pgs-mono" style={{ color: "var(--text)", fontWeight: 700 }}>{o.hireUsed || 0}</span> / {o.hireTotal || 0} ตัว</span>
                <span className="pgs-badge" style={{ background: HIRE_STATUS[o.hireStatus === "done" ? "done" : "ongoing"].color + "22", color: HIRE_STATUS[o.hireStatus === "done" ? "done" : "ongoing"].color }}>
                  {HIRE_STATUS[o.hireStatus === "done" ? "done" : "ongoing"].label}
                </span>
              </div>
            )}
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 8, gap: 6 }}>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {!o.cancelled && o.type === "sell_pokemon" && o.paymentStatus !== "paid" && (
                  <button
                    className="pgs-btn pgs-btn-outline" style={{ padding: "6px 10px", fontSize: 11 }}
                    onClick={(e) => { e.stopPropagation(); onQuickComplete(o); }}
                  ><CheckCircle2 size={12} /> เสร็จสิ้น</button>
                )}
                {!o.cancelled && o.type === "sell_pokemon" && (
                  <button
                    className="pgs-btn pgs-btn-danger" style={{ padding: "6px 10px", fontSize: 11 }}
                    onClick={(e) => { e.stopPropagation(); onQuickCancel(o.id); }}
                  ><Ban size={12} /> ยกเลิก</button>
                )}
              </div>
              <button className="pgs-iconbtn" onClick={(e) => { e.stopPropagation(); openReceipt(o); }} title="ใบเสร็จ"><Receipt size={13} /></button>
            </div>
          </div>
        );
      })}
    </div>
  );
}

function VariantChips({ value, onChange, disabled, multi = true }) {
  const toggle = (k) => {
    if (disabled) return;
    if (!multi) { onChange([k]); return; }
    const has = value.includes(k);
    if (k === "normal") { onChange(["normal"]); return; }
    let next = has ? value.filter(v => v !== k) : [...value.filter(v => v !== "normal"), k];
    if (next.length === 0) next = ["normal"];
    onChange(next);
  };
  return (
    <div style={{ display: "flex", flexWrap: "wrap", gap: 6, opacity: disabled ? 0.55 : 1 }}>
      {Object.entries(POKEMON_VARIANTS).map(([k, v]) => (
        <button
          key={k}
          type="button"
          className={"pgs-chip" + (value.includes(k) ? " active" : "")}
          style={disabled ? { cursor: "not-allowed" } : undefined}
          disabled={disabled}
          onClick={() => toggle(k)}
        >
          {v.emoji} {v.label}
        </button>
      ))}
    </div>
  );
}

function RoundsEditor({ mode, rounds, onChange }) {
  const addRound = () => onChange([...rounds, { id: genId(), date: mode === "scheduled" ? todayStr() : "", count: 1, done: false }]);
  const updateRound = (id, patch) => onChange(rounds.map(r => r.id === id ? { ...r, ...patch } : r));
  const removeRound = (id) => onChange(rounds.filter(r => r.id !== id));

  if (mode === "anytime") {
    const r = rounds[0] || { id: genId(), date: "", count: 1, done: false };
    return (
      <div className="pgs-field">
        <label className="pgs-label">จำนวนรอบที่ต้องตี (ไม่ระบุวัน)</label>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <input className="pgs-input pgs-mono" type="number" min="1" style={{ flex: 1 }} value={r.count}
            onChange={e => onChange([{ ...r, count: e.target.value }])} />
          <button type="button" className={"pgs-chip" + (r.done ? " active" : "")} onClick={() => onChange([{ ...r, done: !r.done }])}>
            {r.done ? "เสร็จแล้ว ✓" : "ยังไม่เสร็จ"}
          </button>
        </div>
      </div>
    );
  }
  return (
    <div className="pgs-field">
      <label className="pgs-label">รอบที่ตั้งไว้ ({rounds.filter(r => r.done).length}/{rounds.length} เสร็จ)</label>
      {rounds.map((r, i) => (
        <div key={r.id} className="pgs-roundrow">
          <span style={{ fontSize: 11, color: "var(--muted)", width: 16 }}>{i + 1}</span>
          <input className="pgs-input" type="date" style={{ flex: 1.4 }} value={r.date} onChange={e => updateRound(r.id, { date: e.target.value })} />
          <input className="pgs-input pgs-mono" type="number" min="1" style={{ flex: 0.7 }} value={r.count} onChange={e => updateRound(r.id, { count: e.target.value })} />
          <button type="button" className="pgs-iconbtn" style={{ color: r.done ? "var(--green)" : "var(--muted)" }} onClick={() => updateRound(r.id, { done: !r.done })}><CheckCircle2 size={15} /></button>
          <button type="button" className="pgs-iconbtn" onClick={() => removeRound(r.id)}><X size={14} /></button>
        </div>
      ))}
      <button type="button" className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginTop: 4 }} onClick={addRound}><Plus size={14} /> เพิ่มรอบ</button>
    </div>
  );
}

function OrderModal({ data, mode, item, onClose, onSave, onCancel, onRestore, onDelete, onReceipt }) {
  const [form, setForm] = useState(item || {
    id: genId(), customerId: data.customers[0]?.id || "", customerGameId: data.customers[0]?.gameIds?.[0]?.value || "",
    type: "sell_pokemon", pokemonName: "", pokemonVariants: ["normal"], quantity: 1, stockItemId: null,
    price: "", sourceAccountId: data.gameAccounts[0]?.id || "",
    paymentStatus: "pending", paidAmount: 0, tradeStatus: "waiting",
    hireMode: "anytime", rounds: [], hireTotal: 1, hireUsed: 0, hireStatus: "ongoing",
    appointmentDate: "", note: "", proofImageDataUrl: "",
    createdAt: new Date().toISOString(), paidDate: "", cancelled: false, cancelledAt: null, cancelHistory: [],
  });
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [showCancelReason, setShowCancelReason] = useState(false);
  const [cancelReason, setCancelReason] = useState("");
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const isSell = form.type === "sell_pokemon";
  const isHire = !isSell;

  const selectedCustomer = data.customers.find(c => c.id === form.customerId);
  const selectedAccount = data.gameAccounts.find(a => a.id === form.sourceAccountId);
  const stockOptions = selectedAccount?.stock || [];

  function pickCustomer(id) {
    const c = data.customers.find(x => x.id === id);
    setForm(f => ({ ...f, customerId: id, customerGameId: c?.gameIds?.[0]?.value || "" }));
  }
  function pickStock(stockId) {
    if (!stockId) { set("stockItemId", null); return; }
    const s = stockOptions.find(x => x.id === stockId);
    if (!s) return;
    setForm(f => ({ ...f, stockItemId: stockId, pokemonName: s.name, pokemonVariants: s.variants && s.variants.length ? s.variants : ["normal"] }));
  }

  function submit() {
    if (!form.customerId) return;
    const price = Number(form.price) || 0;
    let paidAmount = 0;
    if (form.paymentStatus === "paid") paidAmount = price;
    else if (form.paymentStatus === "partial") paidAmount = clamp0(Math.min(Number(form.paidAmount) || 0, price));
    const payload = { ...form, price, paidAmount, quantity: Number(form.quantity) || 1 };
    if ((form.paymentStatus === "paid" || form.paymentStatus === "partial") && !payload.paidDate) payload.paidDate = new Date().toISOString();
    if (isHire) {
      payload.rounds = form.rounds;
      payload.appointmentDate = "";
      payload.hireTotal = clamp0(form.hireTotal) || 1;
      payload.hireUsed = Math.min(clamp0(form.hireUsed), payload.hireTotal);
    }
    onSave(payload);
  }

  if (data.customers.length === 0) {
    return (
      <Modal title="เพิ่มออเดอร์" onClose={onClose}>
        <EmptyState text="กรุณาเพิ่มลูกค้าก่อนสร้างออเดอร์" />
      </Modal>
    );
  }

  return (
    <Modal title={mode === "add" ? "เพิ่มออเดอร์" : `แก้ไขออเดอร์ ${form.code || ""}`} onClose={onClose}>
      {form.cancelled && (
        <div className="pgs-cancelbanner"><Ban size={13} /> ออเดอร์นี้ถูกยกเลิกแล้ว {form.cancelledAt ? `(${fmtDate(form.cancelledAt)})` : ""}</div>
      )}
      {(form.cancelHistory || []).length > 0 && (
        <div className="pgs-card" style={{ marginBottom: 12 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: "var(--muted)", marginBottom: 6 }}>ประวัติการยกเลิก</div>
          {form.cancelHistory.map(h => (
            <div key={h.id || h.date} style={{ fontSize: 11, marginBottom: 4 }}>
              <span className="pgs-mono" style={{ color: "var(--muted)" }}>{fmtDate(h.date)}</span> — {h.reason}
            </div>
          ))}
        </div>
      )}
      <div className="pgs-field">
        <label className="pgs-label">ประเภทบริการ</label>
        <div style={{ display: "flex", gap: 6 }}>
          {Object.entries(ORDER_TYPES).map(([k, v]) => (
            <button key={k} className={"pgs-chip" + (form.type === k ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => set("type", k)}>{v.emoji} {v.short}</button>
          ))}
        </div>
      </div>
      <div className="pgs-field">
        <label className="pgs-label">ลูกค้า</label>
        <select className="pgs-select" value={form.customerId} onChange={e => pickCustomer(e.target.value)}>
          {data.customers.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
        </select>
      </div>
      {selectedCustomer && (selectedCustomer.gameIds || []).length > 0 && (
        <div className="pgs-field">
          <label className="pgs-label">ไอดีของลูกค้าที่ใช้ในออเดอร์นี้</label>
          <select className="pgs-select" value={form.customerGameId} onChange={e => set("customerGameId", e.target.value)}>
            {selectedCustomer.gameIds.map(g => <option key={g.id} value={g.value}>{g.value || "(ไม่มีชื่อ)"}</option>)}
            <option value="">- ไม่ระบุ -</option>
          </select>
        </div>
      )}
      {isSell && (
        <>
          <div className="pgs-field">
            <label className="pgs-label">ไอดีต้นทาง</label>
            <select className="pgs-select" value={form.sourceAccountId} onChange={e => { set("sourceAccountId", e.target.value); set("stockItemId", null); }}>
              <option value="">- ไม่ระบุ -</option>
              {data.gameAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </div>
          {stockOptions.length > 0 && (
            <div className="pgs-field">
              <label className="pgs-label">เลือกจากสต๊อก (ตัดสต๊อกอัตโนมัติ)</label>
              <select className="pgs-select" value={form.stockItemId || ""} onChange={e => pickStock(e.target.value)}>
                <option value="">- กรอกเอง (ไม่ตัดสต๊อก) -</option>
                {stockOptions.map(s => <option key={s.id} value={s.id}>{s.name} · เหลือ {s.quantity}</option>)}
              </select>
            </div>
          )}
          <div style={{ display: "flex", gap: 10 }}>
            <div className="pgs-field" style={{ flex: 2 }}>
              <label className="pgs-label">ชื่อ Pokémon</label>
              <input className="pgs-input" value={form.pokemonName} onChange={e => set("pokemonName", e.target.value)} placeholder="เช่น Rayquaza" />
            </div>
            <div className="pgs-field" style={{ flex: 1 }}>
              <label className="pgs-label">จำนวน</label>
              <input className="pgs-input" type="number" min="1" value={form.quantity} onChange={e => set("quantity", e.target.value)} />
            </div>
          </div>
          <div className="pgs-field">
            <label className="pgs-label">ประเภท Pokémon</label>
            <VariantChips value={form.pokemonVariants} onChange={(v) => set("pokemonVariants", v)} disabled={!!form.stockItemId} />
            {form.stockItemId && (
              <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 6 }}>
                ล็อกตามรายการสต๊อกที่เลือก — ถ้าต้องการชนิดอื่น ให้เปลี่ยนตัวเลือก "เลือกจากสต๊อก" ด้านบน หรือเลือก "กรอกเอง (ไม่ตัดสต๊อก)" ก่อน
              </div>
            )}
          </div>
        </>
      )}
      {isHire && (
        <>
          <div className="pgs-field">
            <label className="pgs-label">โหมดนัดตี</label>
            <div style={{ display: "flex", gap: 6 }}>
              {Object.entries(HIRE_MODES).map(([k, v]) => (
                <button key={k} className={"pgs-chip" + (form.hireMode === k ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => set("hireMode", k)}>{v.label}</button>
              ))}
            </div>
          </div>
          <RoundsEditor mode={form.hireMode} rounds={form.rounds} onChange={(r) => set("rounds", r)} />
          <div style={{ display: "flex", gap: 10 }}>
            <div className="pgs-field" style={{ flex: 1 }}>
              <label className="pgs-label">จำนวนที่ซื้อทั้งหมด (ตัว/รอบ)</label>
              <input className="pgs-input pgs-mono" type="number" min="1" value={form.hireTotal} onChange={e => set("hireTotal", e.target.value)} />
            </div>
            <div className="pgs-field" style={{ flex: 1 }}>
              <label className="pgs-label">ใช้ไปแล้ว</label>
              <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                <button type="button" className="pgs-iconbtn" onClick={() => set("hireUsed", clamp0((Number(form.hireUsed) || 0) - 1))}><Minus size={13} /></button>
                <input className="pgs-input pgs-mono" style={{ textAlign: "center" }} type="number" min="0" value={form.hireUsed} onChange={e => set("hireUsed", e.target.value)} />
                <button type="button" className="pgs-iconbtn" onClick={() => set("hireUsed", clamp0(Math.min((Number(form.hireUsed) || 0) + 1, Number(form.hireTotal) || Infinity)))}><Plus size={13} /></button>
              </div>
            </div>
          </div>
          <div style={{ fontSize: 11, color: "var(--muted)", marginTop: -6, marginBottom: 14 }}>
            เหลืออีก {clamp0((Number(form.hireTotal) || 0) - (Number(form.hireUsed) || 0))} ตัว/รอบ — จำนวนที่ซื้อทั้งหมดจะถูกจำไว้เสมอ แม้จะปรับยอดที่ใช้ไปในภายหลัง
          </div>
        </>
      )}
      <div className="pgs-field">
        <label className="pgs-label">ราคารวม (บาท)</label>
        <input className="pgs-input pgs-mono" type="number" value={form.price} onChange={e => set("price", e.target.value)} placeholder="0" />
      </div>
      <div className="pgs-field">
        <label className="pgs-label">สถานะชำระ</label>
        <select className="pgs-select" value={form.paymentStatus} onChange={e => set("paymentStatus", e.target.value)}>
          {Object.entries(PAYMENT_STATUS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
        </select>
      </div>
      {form.paymentStatus === "partial" && (
        <div className="pgs-field">
          <label className="pgs-label">จำนวนที่ชำระแล้ว (บาท)</label>
          <input className="pgs-input pgs-mono" type="number" value={form.paidAmount} onChange={e => set("paidAmount", e.target.value)} />
          <div style={{ fontSize: 11, color: "var(--red)", marginTop: 6 }}>คงค้าง ฿{fmtMoney(clamp0((Number(form.price) || 0) - (Number(form.paidAmount) || 0)))}</div>
        </div>
      )}
      {isSell && (
        <div className="pgs-field">
          <label className="pgs-label">สถานะเทรด</label>
          <select
            className="pgs-select" value={form.tradeStatus}
            onChange={e => {
              const v = e.target.value;
              set("tradeStatus", v);
              if (v === "three_hearts") {
                const d = new Date();
                d.setDate(d.getDate() + 30);
                set("appointmentDate", d.toISOString().slice(0, 10));
              }
            }}
          >
            {Object.entries(TRADE_STATUS).map(([k, v]) => <option key={k} value={k}>{v.label}</option>)}
          </select>
        </div>
      )}
      {isSell && (
        <div className="pgs-field">
          <label className="pgs-label">วันนัด (สำหรับนัดเทรด)</label>
          <input className="pgs-input" type="date" value={form.appointmentDate} onChange={e => set("appointmentDate", e.target.value)} />
        </div>
      )}
      <div className="pgs-field">
        <label className="pgs-label">รูปภาพกิจกรรม / หลักฐาน (ถ้ามี — จะแสดงในใบเสร็จ)</label>
        <ProofImagePicker value={form.proofImageDataUrl} onChange={(v) => setForm(f => ({ ...f, proofImageDataUrl: v, driveFileId: null }))} />
      </div>
      <div className="pgs-field">
        <label className="pgs-label">หมายเหตุ</label>
        <textarea className="pgs-textarea" rows={2} value={form.note} onChange={e => set("note", e.target.value)} />
      </div>
      <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
        <button className="pgs-btn pgs-btn-primary" style={{ flex: 1 }} onClick={submit}>บันทึก</button>
        {mode === "edit" && (
          <button className="pgs-btn pgs-btn-outline" onClick={() => onReceipt(form)}><Receipt size={14} /></button>
        )}
      </div>
      {mode === "edit" && (
        <div>
          {!form.cancelled && showCancelReason && (
            <div className="pgs-field">
              <label className="pgs-label">ยกเลิกเพราะอะไร?</label>
              <textarea
                className="pgs-textarea" rows={2} autoFocus
                value={cancelReason}
                onChange={e => setCancelReason(e.target.value)}
                placeholder="ระบุเหตุผลที่ยกเลิกออเดอร์นี้..."
              />
            </div>
          )}
          <div style={{ display: "flex", gap: 8 }}>
            {form.cancelled ? (
              <button className="pgs-btn pgs-btn-outline" style={{ flex: 1 }} onClick={() => onRestore(form.id)}><RotateCcw size={14} /> กู้คืนออเดอร์</button>
            ) : !showCancelReason ? (
              <button className="pgs-btn pgs-btn-outline" style={{ flex: 1 }} onClick={() => setShowCancelReason(true)}><Ban size={14} /> ยกเลิกออเดอร์</button>
            ) : (
              <button
                className="pgs-btn pgs-btn-danger" style={{ flex: 1 }}
                disabled={!cancelReason.trim()}
                onClick={() => onCancel(form.id, cancelReason)}
              >ยืนยันยกเลิก (ระบุเหตุผลก่อน)</button>
            )}
            {!confirmDelete ? (
              <button className="pgs-btn pgs-btn-danger" style={{ flex: 1 }} onClick={() => setConfirmDelete(true)}><Trash2 size={14} /> ลบถาวร</button>
            ) : (
              <button className="pgs-btn pgs-btn-danger" style={{ flex: 1 }} onClick={() => onDelete(form.id)}>ยืนยันลบถาวร?</button>
            )}
          </div>
        </div>
      )}
    </Modal>
  );
}

// =================== RECEIPT ===================

function ProofImagePicker({ value, onChange }) {
  const ref = useRef(null);
  const [busy, setBusy] = useState(false);
  async function handleFile(e) {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    if (!file.type.startsWith("image/")) return;
    setBusy(true);
    try {
      const dataUrl = await fileToJpegDataUrl(file, 900, 0.72);
      onChange(dataUrl);
    } catch {
      // ignore — leave value unchanged on failure
    } finally {
      setBusy(false);
    }
  }
  if (value) {
    return (
      <div>
        <img src={value} alt="รูปภาพกิจกรรม" style={{ width: "100%", maxHeight: 180, objectFit: "cover", borderRadius: 12, border: "1px solid var(--border)", marginBottom: 8 }} />
        <div style={{ display: "flex", gap: 8 }}>
          <button type="button" className="pgs-btn pgs-btn-outline" style={{ flex: 1 }} onClick={() => ref.current?.click()}><Upload size={14} /> เปลี่ยนรูป</button>
          <button type="button" className="pgs-btn pgs-btn-outline" style={{ flex: 1 }} onClick={() => onChange("")}><Trash2 size={14} /> ลบรูป</button>
        </div>
        <input ref={ref} type="file" accept="image/*" style={{ display: "none" }} onChange={handleFile} />
      </div>
    );
  }
  return (
    <div>
      <button type="button" className="pgs-btn pgs-btn-outline" style={{ width: "100%" }} disabled={busy} onClick={() => ref.current?.click()}>
        <Upload size={14} /> {busy ? "กำลังอัปโหลด..." : "แนบรูปภาพ"}
      </button>
      <input ref={ref} type="file" accept="image/*" style={{ display: "none" }} onChange={handleFile} />
    </div>
  );
}

function buildReceiptLines(order, data, custName, accName) {
  const lines = [];
  lines.push(`🧾 ${data.settings.shopName}`);
  lines.push(`เลขที่: ${order.code || "-"}`);
  lines.push(`วันที่: ${fmtDate(order.createdAt)}`);
  lines.push(`--------------------------------`);
  lines.push(`ลูกค้า: ${custName(order.customerId)}`);
  if (order.customerGameId) lines.push(`ไอดีเกม: ${order.customerGameId}`);
  lines.push(`บริการ: ${ORDER_TYPES[order.type]?.label || "-"}`);
  if (order.type === "sell_pokemon") {
    const variants = (order.pokemonVariants || []).filter(v => v !== "normal").map(v => POKEMON_VARIANTS[v]?.label).filter(Boolean).join(", ");
    lines.push(`Pokémon: ${order.pokemonName || "-"}${variants ? " (" + variants + ")" : ""} x${order.quantity || 1}`);
    if (order.sourceAccountId) lines.push(`ไอดีต้นทาง: ${accName(order.sourceAccountId)}`);
  } else {
    lines.push(`โหมด: ${HIRE_MODES[order.hireMode]?.label || "-"}`);
    lines.push(`จำนวนที่ซื้อทั้งหมด: ${order.hireTotal || 0} ตัว/รอบ (ใช้ไปแล้ว ${order.hireUsed || 0})`);
    (order.rounds || []).forEach((r, i) => {
      lines.push(`  รอบ ${i + 1}: ${r.date ? fmtDate(r.date) : "ไม่ระบุวัน"} x${r.count} ${r.done ? "(เสร็จแล้ว)" : ""}`);
    });
  }
  lines.push(`--------------------------------`);
  lines.push(`ราคารวม: ฿${fmtMoney(order.price)}`);
  if (order.paymentStatus === "partial") {
    lines.push(`ชำระแล้ว: ฿${fmtMoney(order.paidAmount)}`);
    lines.push(`คงเหลือ: ฿${fmtMoney(orderBalance(order))}`);
  }
  lines.push(`สถานะชำระ: ${PAYMENT_STATUS[order.paymentStatus]?.label || "-"}`);
  if (order.type === "sell_pokemon") lines.push(`สถานะเทรด: ${TRADE_STATUS[order.tradeStatus]?.label || "-"}`);
  if (order.note) lines.push(`หมายเหตุ: ${order.note}`);
  if (order.proofImageDataUrl) lines.push(`📷 แนบรูปภาพกิจกรรม (ดูในแอป)`);
  if (order.cancelled) lines.push(`⚠️ ออเดอร์นี้ถูกยกเลิก`);
  return lines;
}

// structured summary used by the visual receipt card (both the on-screen preview and the canvas export)
function buildReceiptData(order, data, custName, accName) {
  const orderNoForCustomer = data.orders
    .filter(o => o.customerId === order.customerId)
    .slice()
    .sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""))
    .findIndex(o => o.id === order.id) + 1;

  const items = [];
  if (order.type === "sell_pokemon") {
    const variants = (order.pokemonVariants || []).filter(v => v !== "normal").map(v => POKEMON_VARIANTS[v]?.label).filter(Boolean).join(", ");
    items.push({
      label: `${order.pokemonName || "Pokémon"}${variants ? " (" + variants + ")" : ""}`,
      sub: `x${order.quantity || 1}${order.sourceAccountId ? " · " + accName(order.sourceAccountId) : ""}`,
      price: order.price,
    });
  } else {
    items.push({
      label: ORDER_TYPES[order.type]?.label || "-",
      sub: `${HIRE_MODES[order.hireMode]?.label || ""} · ใช้ไป ${order.hireUsed || 0}/${order.hireTotal || 0}`,
      price: order.price,
    });
  }

  return {
    shopName: data.settings.shopName,
    logoDataUrl: data.settings.logoDataUrl || "",
    receiptBgDataUrl: data.settings.receiptBgDataUrl || "",
    code: order.code || "-",
    dateStr: fmtDate(order.createdAt),
    customerName: custName(order.customerId),
    customerGameId: order.customerGameId || "",
    orderNoForCustomer,
    serviceEmoji: ORDER_TYPES[order.type]?.emoji || "🧾",
    serviceLabel: ORDER_TYPES[order.type]?.label || "-",
    periodStr: order.type !== "sell_pokemon" && order.rounds && order.rounds.length
      ? `${fmtDate(order.rounds[0]?.date || order.createdAt)} — ${fmtDate(order.rounds[order.rounds.length - 1]?.date || order.createdAt)}`
      : (order.appointmentDate ? `นัด ${fmtDate(order.appointmentDate)}` : null),
    items,
    total: Number(order.price) || 0,
    paidAmount: order.paymentStatus === "paid" ? Number(order.price) || 0 : Number(order.paidAmount) || 0,
    balance: orderBalance(order),
    paymentStatus: PAYMENT_STATUS[order.paymentStatus]?.label || "-",
    paymentColor: PAYMENT_STATUS[order.paymentStatus]?.color || "#8b8da6",
    tradeStatus: order.type === "sell_pokemon" ? (TRADE_STATUS[order.tradeStatus]?.label || null) : null,
    note: order.note || "",
    proofImageDataUrl: order.proofImageDataUrl || "",
    cancelled: !!order.cancelled,
  };
}

function loadImageAsync(src) {
  return new Promise((resolve) => {
    if (!src) { resolve(null); return; }
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => resolve(null);
    img.src = src;
  });
}

// rounded-rect helper shared by the canvas receipt renderer
function roundRectPath(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.arcTo(x + w, y, x + w, y + h, r);
  ctx.arcTo(x + w, y + h, x, y + h, r);
  ctx.arcTo(x, y + h, x, y, r);
  ctx.arcTo(x, y, x + w, y, r);
  ctx.closePath();
}

async function downloadReceiptImage(order, data, custName, accName) {
  const r = buildReceiptData(order, data, custName, accName);
  const [logoImg, proofImg, bgImg] = await Promise.all([loadImageAsync(r.logoDataUrl), loadImageAsync(r.proofImageDataUrl), loadImageAsync(r.receiptBgDataUrl)]);

  const width = 640;
  const pad = 28;
  let y = 0; // running cursor, computed as we lay things out top-to-bottom

  // pre-measure proof image height (contain within card width, capped)
  let proofH = 0;
  if (proofImg) {
    const maxW = width - pad * 2;
    const maxH = 320;
    const scale = Math.min(maxW / proofImg.width, maxH / proofImg.height);
    proofH = Math.round(proofImg.height * scale);
  }
  const itemsH = 46 + r.items.length * 40 + 46; // header + rows + total row
  const noteH = r.note ? 40 : 0;
  const height = pad + 96 + 70 + (r.periodStr ? 26 : 0) + 16 + itemsH + 16 + (proofImg ? 30 + proofH + 16 : 0) + 40 + noteH + (r.cancelled ? 40 : 0) + 60 + pad;

  const canvas = document.createElement("canvas");
  canvas.width = width; canvas.height = height;
  const ctx = canvas.getContext("2d");

  // background — either the user's uploaded photo (cover-fit, with a dark overlay so text stays
  // readable), or the plain gradient fallback
  roundRectPath(ctx, 0, 0, width, height, 22);
  if (bgImg) {
    ctx.save();
    ctx.clip();
    const s = Math.max(width / bgImg.width, height / bgImg.height);
    const dw = bgImg.width * s, dh = bgImg.height * s;
    ctx.drawImage(bgImg, (width - dw) / 2, (height - dh) / 2, dw, dh);
    const overlay = ctx.createLinearGradient(0, 0, width, height);
    overlay.addColorStop(0, "rgba(12,13,21,0.55)");
    overlay.addColorStop(1, "rgba(12,13,21,0.82)");
    ctx.fillStyle = overlay;
    ctx.fillRect(0, 0, width, height);
    ctx.restore();
  } else {
    const bgGrad = ctx.createLinearGradient(0, 0, width, height);
    bgGrad.addColorStop(0, "#14151f");
    bgGrad.addColorStop(1, "#0c0d15");
    ctx.fillStyle = bgGrad;
    ctx.fill();
  }
  ctx.strokeStyle = "#2c2f46";
  ctx.lineWidth = 1;
  roundRectPath(ctx, 0.5, 0.5, width - 1, height - 1, 22);
  ctx.stroke();

  y = pad;

  // top pill badge
  ctx.font = "700 12px Inter, sans-serif";
  const pillText = `${r.shopName.toUpperCase()} · RECEIPT`;
  const pillW = ctx.measureText(pillText).width + 28;
  ctx.fillStyle = "rgba(255,203,5,0.14)";
  roundRectPath(ctx, pad, y, pillW, 26, 13);
  ctx.fill();
  ctx.fillStyle = "#ffcb05";
  ctx.textBaseline = "middle";
  ctx.fillText(pillText, pad + 14, y + 14);

  // logo circle top-right
  if (logoImg) {
    const s = 40;
    ctx.save();
    ctx.beginPath();
    ctx.arc(width - pad - s / 2, y + 13, s / 2, 0, Math.PI * 2);
    ctx.closePath();
    ctx.clip();
    ctx.drawImage(logoImg, width - pad - s, y - 7, s, s);
    ctx.restore();
  }

  y += 46;
  ctx.textBaseline = "alphabetic";
  ctx.fillStyle = "#f2f3f8";
  ctx.font = "700 26px 'Baloo 2', Inter, sans-serif";
  ctx.fillText(r.cancelled ? "ใบเสร็จ (ยกเลิกแล้ว)" : "ใบเสร็จ / สรุปออเดอร์", pad, y);

  y += 26;
  ctx.font = "500 13px 'JetBrains Mono', monospace";
  ctx.fillStyle = "#8b8da6";
  ctx.fillText(`เลขที่ ${r.code}  ·  ${r.dateStr}`, pad, y);

  y += 30;
  ctx.strokeStyle = "#2c2f46";
  ctx.beginPath(); ctx.moveTo(pad, y); ctx.lineTo(width - pad, y); ctx.stroke();
  y += 26;

  // customer row
  ctx.font = "700 15px Inter, sans-serif";
  ctx.fillStyle = "#f2f3f8";
  ctx.fillText(`${r.customerName}`, pad, y);
  ctx.font = "600 11px Inter, sans-serif";
  ctx.fillStyle = "#ffcb05";
  ctx.textAlign = "right";
  ctx.fillText(`ครั้งที่ ${r.orderNoForCustomer}`, width - pad, y);
  ctx.textAlign = "left";

  y += 22;
  ctx.font = "500 12px Inter, sans-serif";
  ctx.fillStyle = "#8b8da6";
  ctx.fillText(`${r.serviceEmoji} ${r.serviceLabel}${r.customerGameId ? "  ·  ไอดี " + r.customerGameId : ""}`, pad, y);

  if (r.periodStr) {
    y += 22;
    ctx.fillText(`🗓️ ${r.periodStr}`, pad, y);
  }

  y += 24;

  // items card
  const cardTop = y;
  ctx.fillStyle = "rgba(255,255,255,0.03)";
  roundRectPath(ctx, pad, cardTop, width - pad * 2, itemsH, 14);
  ctx.fill();
  ctx.strokeStyle = "#2c2f46";
  roundRectPath(ctx, pad, cardTop, width - pad * 2, itemsH, 14);
  ctx.stroke();

  let iy = cardTop + 26;
  r.items.forEach(it => {
    ctx.font = "700 13px Inter, sans-serif";
    ctx.fillStyle = "#f2f3f8";
    ctx.fillText(it.label, pad + 16, iy);
    ctx.textAlign = "right";
    ctx.font = "700 13px 'JetBrains Mono', monospace";
    ctx.fillText(`฿${fmtMoney(it.price)}`, width - pad - 16, iy);
    ctx.textAlign = "left";
    iy += 16;
    ctx.font = "500 11px Inter, sans-serif";
    ctx.fillStyle = "#8b8da6";
    ctx.fillText(it.sub, pad + 16, iy);
    iy += 24;
  });
  ctx.strokeStyle = "#2c2f46";
  ctx.beginPath(); ctx.moveTo(pad + 16, iy - 4); ctx.lineTo(width - pad - 16, iy - 4); ctx.stroke();
  iy += 20;
  ctx.font = "600 13px Inter, sans-serif";
  ctx.fillStyle = "#8b8da6";
  ctx.fillText("รวมเงิน", pad + 16, iy);
  ctx.textAlign = "right";
  ctx.font = "700 20px 'JetBrains Mono', monospace";
  ctx.fillStyle = "#ffcb05";
  ctx.fillText(`฿${fmtMoney(r.total)}`, width - pad - 16, iy);
  ctx.textAlign = "left";

  y = cardTop + itemsH + 16;

  // activity image
  if (proofImg) {
    ctx.font = "700 11px Inter, sans-serif";
    ctx.fillStyle = "#8b8da6";
    ctx.fillText("รูปภาพกิจกรรม", pad, y);
    y += 16;
    const dw = width - pad * 2;
    roundRectPath(ctx, pad, y, dw, proofH, 12);
    ctx.save();
    ctx.clip();
    ctx.drawImage(proofImg, pad, y, dw, proofH);
    ctx.restore();
    ctx.strokeStyle = "#2c2f46";
    roundRectPath(ctx, pad, y, dw, proofH, 12);
    ctx.stroke();
    y += proofH + 16;
  }

  // status badges
  ctx.font = "700 11px Inter, sans-serif";
  const badge = (text, color, x) => {
    const w = ctx.measureText(text).width + 22;
    ctx.fillStyle = color + "22";
    roundRectPath(ctx, x, y, w, 24, 12);
    ctx.fill();
    ctx.fillStyle = color;
    ctx.textBaseline = "middle";
    ctx.fillText(text, x + 11, y + 12);
    ctx.textBaseline = "alphabetic";
    return x + w + 8;
  };
  let bx = pad;
  bx = badge(r.paymentStatus, r.paymentColor, bx);
  if (r.tradeStatus) badge(r.tradeStatus, "#ffcb05", bx);
  y += 40;

  if (r.note) {
    ctx.font = "500 12px Inter, sans-serif";
    ctx.fillStyle = "#8b8da6";
    ctx.fillText(`หมายเหตุ: ${r.note}`, pad, y);
    y += 26;
  }

  if (r.cancelled) {
    ctx.fillStyle = "rgba(255,84,112,0.15)";
    roundRectPath(ctx, pad, y - 18, width - pad * 2, 32, 10);
    ctx.fill();
    ctx.fillStyle = "#ff5470";
    ctx.font = "700 12px Inter, sans-serif";
    ctx.fillText("⚠️ ออเดอร์นี้ถูกยกเลิกแล้ว", pad + 12, y + 3);
    y += 40;
  }

  // footer
  ctx.strokeStyle = "#2c2f46";
  ctx.beginPath(); ctx.moveTo(pad, y); ctx.lineTo(width - pad, y); ctx.stroke();
  y += 26;
  ctx.font = "600 12px Inter, sans-serif";
  ctx.fillStyle = "#8b8da6";
  ctx.fillText(`ขอบคุณที่ใช้บริการ ${r.shopName} 🐾`, pad, y);

  return new Promise((resolve) => {
    canvas.toBlob(blob => {
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url; a.download = `receipt-${order.code || "order"}.png`; a.click();
      URL.revokeObjectURL(url);
      resolve();
    });
  });
}

function ReceiptModal({ order, data, custName, accName, onClose, onToast }) {
  const lines = buildReceiptLines(order, data, custName, accName);
  const text = lines.join("\n");
  const r = useMemo(() => buildReceiptData(order, data, custName, accName), [order, data]);
  const [downloading, setDownloading] = useState(false);
  async function copyText() {
    try {
      await navigator.clipboard.writeText(text);
      onToast("คัดลอกใบเสร็จแล้ว");
    } catch {
      onToast("คัดลอกไม่สำเร็จ");
    }
  }
  async function download() {
    setDownloading(true);
    try {
      await downloadReceiptImage(order, data, custName, accName);
    } finally {
      setDownloading(false);
    }
  }
  return (
    <Modal title="ใบเสร็จ / สรุปออเดอร์" onClose={onClose}>
      <div style={{
        background: r.receiptBgDataUrl
          ? `linear-gradient(165deg, rgba(12,13,21,0.55) 0%, rgba(12,13,21,0.82) 100%), url(${r.receiptBgDataUrl})`
          : "linear-gradient(165deg, #1b1d2a 0%, rgba(20,21,31,0.9) 100%)",
        backgroundSize: r.receiptBgDataUrl ? "cover" : undefined,
        backgroundPosition: r.receiptBgDataUrl ? "center" : undefined,
        border: "1px solid var(--border)", borderRadius: 18, padding: 18, marginBottom: 14,
      }}>
        <div className="pgs-row" style={{ marginBottom: 14 }}>
          <span className="pgs-badge" style={{ background: "rgba(255,203,5,0.14)", color: "var(--yellow)", fontSize: 10, letterSpacing: 0.5 }}>
            {r.shopName.toUpperCase()} · RECEIPT
          </span>
          <ShopLogo logoDataUrl={r.logoDataUrl} size={34} />
        </div>
        <div className="pgs-display" style={{ fontSize: 20, fontWeight: 700, marginBottom: 4 }}>
          {r.cancelled ? "ใบเสร็จ (ยกเลิกแล้ว)" : "ใบเสร็จ / สรุปออเดอร์"}
        </div>
        <div className="pgs-mono" style={{ fontSize: 12, color: "var(--muted)", marginBottom: 16 }}>เลขที่ {r.code} · {r.dateStr}</div>

        <div style={{ borderTop: "1px solid var(--border)", paddingTop: 14, marginBottom: 14 }}>
          <div className="pgs-row" style={{ marginBottom: 4 }}>
            <span style={{ fontWeight: 700, fontSize: 15 }}>{r.customerName}</span>
            <span style={{ fontSize: 11, fontWeight: 700, color: "var(--yellow)" }}>ครั้งที่ {r.orderNoForCustomer}</span>
          </div>
          <div style={{ fontSize: 12, color: "var(--muted)" }}>
            {r.serviceEmoji} {r.serviceLabel}{r.customerGameId ? ` · ไอดี ${r.customerGameId}` : ""}
          </div>
          {r.periodStr && <div style={{ fontSize: 12, color: "var(--muted)", marginTop: 2 }}>🗓️ {r.periodStr}</div>}
        </div>

        <div className="pgs-card" style={{ marginBottom: 14 }}>
          {r.items.map((it, i) => (
            <div key={i} className="pgs-row" style={{ alignItems: "flex-start", marginBottom: 8 }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 13 }}>{it.label}</div>
                <div style={{ fontSize: 11, color: "var(--muted)" }}>{it.sub}</div>
              </div>
              <span className="pgs-mono" style={{ fontWeight: 700, fontSize: 13 }}>฿{fmtMoney(it.price)}</span>
            </div>
          ))}
          <div className="pgs-row" style={{ borderTop: "1px solid var(--border)", paddingTop: 10, marginTop: 4 }}>
            <span style={{ fontSize: 13, color: "var(--muted)", fontWeight: 600 }}>รวมเงิน</span>
            <span className="pgs-mono pgs-display" style={{ fontWeight: 700, fontSize: 20, color: "var(--yellow)" }}>฿{fmtMoney(r.total)}</span>
          </div>
        </div>

        {r.proofImageDataUrl && (
          <div style={{ marginBottom: 14 }}>
            <div className="pgs-label" style={{ marginBottom: 8 }}>รูปภาพกิจกรรม</div>
            <img src={r.proofImageDataUrl} alt="รูปภาพกิจกรรม" style={{ width: "100%", maxHeight: 220, objectFit: "cover", borderRadius: 12, border: "1px solid var(--border)" }} />
          </div>
        )}

        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: r.note ? 10 : 0 }}>
          <span className="pgs-badge" style={{ background: r.paymentColor + "22", color: r.paymentColor }}>{r.paymentStatus}</span>
          {r.tradeStatus && <span className="pgs-badge" style={{ background: "rgba(255,203,5,0.15)", color: "var(--yellow)" }}>{r.tradeStatus}</span>}
        </div>
        {r.note && <div style={{ fontSize: 12, color: "var(--muted)" }}>หมายเหตุ: {r.note}</div>}
        {r.cancelled && <div className="pgs-cancelbanner" style={{ marginTop: 10, marginBottom: 0 }}><Ban size={13} /> ออเดอร์นี้ถูกยกเลิกแล้ว</div>}

        <div style={{ borderTop: "1px solid var(--border)", marginTop: 14, paddingTop: 12, fontSize: 11, color: "var(--muted)", textAlign: "center" }}>
          ขอบคุณที่ใช้บริการ {r.shopName} 🐾
        </div>
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <button className="pgs-btn pgs-btn-outline" style={{ flex: 1 }} onClick={copyText}><Copy size={14} /> คัดลอกข้อความ</button>
        <button className="pgs-btn pgs-btn-primary" style={{ flex: 1 }} disabled={downloading} onClick={download}><Download size={14} /> {downloading ? "กำลังสร้างรูป..." : "ดาวน์โหลดรูป"}</button>
      </div>
    </Modal>
  );
}

// =================== CUSTOMERS ===================
function CustomersTab({ data, openNew, openEdit, openDetail, back }) {
  const [q, setQ] = useState("");
  const spentOf = (id) => data.orders.filter(o => o.customerId === id && !o.cancelled).reduce((s, o) => {
    if (o.paymentStatus === "paid") return s + (Number(o.price) || 0);
    if (o.paymentStatus === "partial") return s + (Number(o.paidAmount) || 0);
    return s;
  }, 0);
  const list = data.customers
    .filter(c => !q || c.name.toLowerCase().includes(q.toLowerCase()))
    .map(c => ({ ...c, _spent: spentOf(c.id) }))
    .sort((a, b) => b._spent - a._spent);
  const medal = (i) => i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : null;
  return (
    <div>
      <SubHeader title="ลูกค้า" back={back} />
      <button className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: 12 }} onClick={openNew}><Plus size={15} /> เพิ่มลูกค้าใหม่</button>
      <div style={{ position: "relative", marginBottom: 12 }}>
        <Search size={14} color="var(--muted)" style={{ position: "absolute", left: 12, top: 12 }} />
        <input className="pgs-input" style={{ paddingLeft: 32 }} placeholder="ค้นหาลูกค้า..." value={q} onChange={e => setQ(e.target.value)} />
      </div>
      {list.length === 0 ? <EmptyState text="ยังไม่มีลูกค้า" /> : list.map((c, i) => (
        <div key={c.id} className="pgs-card" style={{ marginBottom: 8, cursor: "pointer" }} onClick={() => openDetail(c)}>
          <div className="pgs-row">
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              {medal(i) ? (
                <span style={{ fontSize: 18 }}>{medal(i)}</span>
              ) : (
                <span className="pgs-mono" style={{ fontSize: 11, color: "var(--muted)", width: 18, textAlign: "center" }}>#{i + 1}</span>
              )}
              <div>
                <div style={{ fontWeight: 700, fontSize: 14 }}>{c.name}</div>
                <div style={{ fontSize: 11, color: "var(--muted)" }}>{c.facebook || "ไม่มี Facebook"}</div>
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <div className="pgs-mono" style={{ fontWeight: 700, fontSize: 13, color: "var(--green)" }}>฿{fmtMoney(c._spent)}</div>
              <div style={{ fontSize: 10, color: "var(--muted)" }}>ยอดซื้อสะสม</div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

function CustomerModal({ mode, item, onClose, onSave }) {
  const [form, setForm] = useState(item || {
    id: genId(), name: "", facebook: "", note: "",
    gameIds: [{ id: genId(), value: "" }],
    createdAt: new Date().toISOString(), _isNew: true,
  });
  const [touched, setTouched] = useState(false);
  const updateGameId = (id, value) => setForm(f => ({ ...f, gameIds: f.gameIds.map(g => g.id === id ? { ...g, value } : g) }));
  const addGameId = () => setForm(f => ({ ...f, gameIds: [...f.gameIds, { id: genId(), value: "" }] }));
  const removeGameId = (id) => setForm(f => ({ ...f, gameIds: f.gameIds.length > 1 ? f.gameIds.filter(g => g.id !== id) : f.gameIds }));
  const hasGameId = form.gameIds.some(g => (g.value || "").trim());
  const nameOk = !!(form.name || "").trim();
  const fbOk = !!(form.facebook || "").trim();
  const canSave = nameOk && hasGameId && fbOk;
  const err = (bad) => (touched && bad ? { borderColor: "var(--red)" } : undefined);
  return (
    <Modal title={mode === "add" ? "เพิ่มลูกค้า" : "แก้ไขลูกค้า"} onClose={onClose}>
      <div className="pgs-field">
        <label className="pgs-label">ชื่อลูกค้า *</label>
        <input className="pgs-input" style={err(!nameOk)} value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} placeholder="เช่น Ash_Ketchum99" />
        {touched && !nameOk && <div style={{ fontSize: 11, color: "var(--red)", marginTop: 4 }}>กรุณากรอกชื่อลูกค้า</div>}
      </div>
      <div className="pgs-field">
        <label className="pgs-label">ไอดีในเกม (เพิ่มได้หลายไอดี) *</label>
        {form.gameIds.map((g, i) => (
          <div key={g.id} style={{ display: "flex", gap: 6, marginBottom: 6 }}>
            <input className="pgs-input" style={err(!hasGameId)} value={g.value} onChange={e => updateGameId(g.id, e.target.value)} placeholder={`ไอดี #${i + 1}`} />
            {form.gameIds.length > 1 && (
              <button type="button" className="pgs-iconbtn" onClick={() => removeGameId(g.id)}><X size={14} /></button>
            )}
          </div>
        ))}
        <button type="button" className="pgs-btn pgs-btn-outline" style={{ width: "100%" }} onClick={addGameId}><Plus size={14} /> เพิ่มไอดี</button>
        {touched && !hasGameId && <div style={{ fontSize: 11, color: "var(--red)", marginTop: 4 }}>กรุณากรอกไอดีในเกมอย่างน้อย 1 ไอดี</div>}
      </div>
      <div className="pgs-field">
        <label className="pgs-label">Facebook *</label>
        <input className="pgs-input" style={err(!fbOk)} value={form.facebook} onChange={e => setForm(f => ({ ...f, facebook: e.target.value }))} placeholder="ชื่อ / ลิงก์โปรไฟล์" />
        {touched && !fbOk && <div style={{ fontSize: 11, color: "var(--red)", marginTop: 4 }}>กรุณากรอกชื่อ Facebook</div>}
      </div>
      <div className="pgs-field">
        <label className="pgs-label">หมายเหตุ</label>
        <textarea className="pgs-textarea" rows={2} value={form.note} onChange={e => setForm(f => ({ ...f, note: e.target.value }))} />
      </div>
      <button
        className="pgs-btn pgs-btn-primary" style={{ width: "100%", opacity: canSave ? 1 : 0.6 }}
        onClick={() => { if (canSave) onSave(form); else setTouched(true); }}
      >บันทึก</button>
    </Modal>
  );
}

// dedicated list for the dashboard's "นัดหมาย/รอบตีใกล้ถึงกำหนด" widget —
// tapping an item jumps straight to the matching Trade tab (นัดเทรด) or Hire tab (รอบตี)
function DueSoonModal({ items, data, custName, onClose, onGoTo }) {
  return (
    <Modal title="นัดหมาย/รอบตีใกล้ถึงกำหนด" onClose={onClose}>
      {items.length === 0 ? <EmptyState text="ไม่มีรายการ" /> : items.map((it, i) => {
        const order = data.orders.find(o => o.id === it.orderId);
        if (!order) return null;
        const isRound = it.kind === "round";
        const overdue = it.remain < 0;
        const dueColor = overdue ? "var(--red)" : (it.remain === 0 ? "var(--yellow)" : "var(--muted)");
        const dueText = overdue ? `เลยกำหนด ${Math.abs(it.remain)} วัน` : it.remain === 0 ? "ถึงกำหนดวันนี้" : `อีก ${it.remain} วัน`;
        return (
          <div
            key={i} className="pgs-card" style={{ marginBottom: 8, cursor: "pointer" }}
            onClick={() => onGoTo(it.kind)}
          >
            <div className="pgs-row" style={{ marginBottom: 4 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                {isRound ? <Target size={15} color="var(--yellow)" /> : <Repeat size={15} color="var(--blue)" />}
                <span style={{ fontWeight: 700, fontSize: 13 }}>{custName(order.customerId)}</span>
              </div>
              <ChevronRight size={15} color="var(--muted)" />
            </div>
            <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 4 }}>
              {isRound ? `รอบตี · ${ORDER_TYPES[order.type]?.label || ""}` : `นัดเทรด · ${order.pokemonName || ""}`}
            </div>
            <div style={{ fontSize: 11, color: dueColor, display: "flex", alignItems: "center", gap: 4 }}>
              <Clock size={11} /> {fmtDate(it.date)} · {dueText}
            </div>
          </div>
        );
      })}
    </Modal>
  );
}

// aggregates unpaid/partial order balances per customer for the "ยอดค้างชำระ" overview
function DebtModal({ data, custName, onClose, onOpenCustomer }) {
  const byCustomer = useMemo(() => {
    const map = {};
    data.orders.filter(o => !o.cancelled).forEach(o => {
      const bal = orderBalance(o);
      if (bal <= 0) return;
      if (!map[o.customerId]) map[o.customerId] = { customerId: o.customerId, total: 0, count: 0 };
      map[o.customerId].total += bal;
      map[o.customerId].count += 1;
    });
    return Object.values(map).sort((a, b) => b.total - a.total);
  }, [data]);
  const grandTotal = byCustomer.reduce((s, c) => s + c.total, 0);
  return (
    <Modal title="ยอดค้างชำระตามลูกค้า" onClose={onClose}>
      <div className="pgs-card" style={{ marginBottom: 12 }}>
        <div className="pgs-row">
          <span style={{ fontSize: 12, color: "var(--muted)" }}>ค้างชำระรวมทั้งร้าน</span>
          <span className="pgs-mono" style={{ fontWeight: 700, fontSize: 18, color: "var(--red)" }}>฿{fmtMoney(grandTotal)}</span>
        </div>
      </div>
      {byCustomer.length === 0 ? <EmptyState text="ไม่มีลูกค้าติดค้างชำระ" /> : byCustomer.map((c, i) => {
        const customer = data.customers.find(x => x.id === c.customerId);
        return (
          <div
            key={c.customerId} className="pgs-row" style={{ padding: "10px 0", borderBottom: "1px solid var(--border)", cursor: customer ? "pointer" : "default" }}
            onClick={() => customer && onOpenCustomer(customer)}
          >
            <div>
              <div style={{ fontWeight: 700, fontSize: 13 }}>#{i + 1} {custName(c.customerId)}</div>
              <div style={{ fontSize: 10, color: "var(--muted)" }}>{c.count} ออเดอร์ค้างชำระ</div>
            </div>
            <span className="pgs-mono" style={{ fontWeight: 700, fontSize: 13, color: "var(--red)" }}>฿{fmtMoney(c.total)}</span>
          </div>
        );
      })}
    </Modal>
  );
}

function CustomerDetail({ item, data, onClose, onEdit, onDelete }) {
  const [period, setPeriod] = useState("all");
  const [confirmDelete, setConfirmDelete] = useState(false);
  const orders = data.orders.filter(o => o.customerId === item.id && !o.cancelled);
  const allOrderCount = data.orders.filter(o => o.customerId === item.id).length;
  const today = todayStr();
  const inPeriod = (o) => {
    if (period === "all") return true;
    const d = (o.createdAt || "").slice(0, period === "month" ? 7 : 4);
    const t = today.slice(0, period === "month" ? 7 : 4);
    return d === t;
  };
  const relevant = orders.filter(inPeriod);
  const paidAmountOf = (o) => o.paymentStatus === "paid" ? Number(o.price || 0) : (o.paymentStatus === "partial" ? Number(o.paidAmount || 0) : 0);
  const byType = (t) => relevant.filter(o => o.type === t);
  const sumType = (t) => byType(t).reduce((s, o) => s + paidAmountOf(o), 0);
  const total = ["sell_pokemon", "hire_boss", "hire_invite"].reduce((s, t) => s + sumType(t), 0);
  return (
    <Modal title={item.name} onClose={onClose}>
      <div style={{ fontSize: 12, color: "var(--muted)", marginBottom: 6 }}>{item.facebook || "ไม่มี Facebook"}{item.note ? " · " + item.note : ""}</div>
      {(item.gameIds || []).some(g => g.value) && (
        <div style={{ fontSize: 12, color: "var(--muted)", marginBottom: 12 }}>
          ไอดีในเกม: {item.gameIds.filter(g => g.value).map(g => g.value).join(", ")}
        </div>
      )}
      <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
        <button className={"pgs-chip" + (period === "month" ? " active" : "")} onClick={() => setPeriod("month")}>รายเดือน</button>
        <button className={"pgs-chip" + (period === "year" ? " active" : "")} onClick={() => setPeriod("year")}>รายปี</button>
        <button className={"pgs-chip" + (period === "all" ? " active" : "")} onClick={() => setPeriod("all")}>ทั้งหมด</button>
      </div>
      <div className="pgs-card" style={{ marginBottom: 10 }}>
        <div className="pgs-row" style={{ marginBottom: 8 }}>
          <span style={{ fontSize: 12, color: "var(--muted)" }}>ยอดรวม</span>
          <span className="pgs-mono" style={{ fontWeight: 700, fontSize: 18, color: "var(--green)" }}>฿{fmtMoney(total)}</span>
        </div>
        {Object.entries(ORDER_TYPES).map(([k, v]) => (
          <div key={k} className="pgs-row" style={{ fontSize: 12, marginBottom: 4 }}>
            <span style={{ color: "var(--muted)" }}>{v.emoji} {v.label} ({byType(k).length})</span>
            <span className="pgs-mono">฿{fmtMoney(sumType(k))}</span>
          </div>
        ))}
      </div>
      <div className="pgs-sectiontitle">ประวัติออเดอร์</div>
      {relevant.length === 0 ? <EmptyState text="ไม่มีข้อมูล" /> : relevant.slice(0, 8).map(o => (
        <div key={o.id} className="pgs-row" style={{ padding: "8px 0", borderBottom: "1px solid var(--border)", fontSize: 12 }}>
          <span>{ORDER_TYPES[o.type].emoji} {o.type === "sell_pokemon" ? o.pokemonName : ORDER_TYPES[o.type].label}</span>
          <span className="pgs-mono">฿{fmtMoney(o.price)}</span>
        </div>
      ))}
      {confirmDelete && allOrderCount > 0 && (
        <div className="pgs-cancelbanner">
          <AlertTriangle size={14} /> ลูกค้านี้มีออเดอร์อยู่ {allOrderCount} รายการ — ลบแล้วออเดอร์เหล่านั้นจะยังอยู่แต่จะไม่แสดงชื่อลูกค้า
        </div>
      )}
      <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
        <button className="pgs-btn pgs-btn-outline" style={{ flex: 1 }} onClick={onEdit}><Edit2 size={14} /> แก้ไข</button>
        {!confirmDelete ? (
          <button className="pgs-btn pgs-btn-danger" style={{ flex: 1 }} onClick={() => setConfirmDelete(true)}><Trash2 size={14} /> ลบ</button>
        ) : (
          <button className="pgs-btn pgs-btn-danger" style={{ flex: 1 }} onClick={onDelete}>ยืนยันลบถาวร?</button>
        )}
      </div>
    </Modal>
  );
}

// =================== ACCOUNTS ===================
function AccountsTab({ data, stats, openNew, openDetail, back }) {
  const [q, setQ] = useState("");
  const accounts = data.gameAccounts.filter(a => !q || a.name.toLowerCase().includes(q.toLowerCase()));
  return (
    <div>
      <SubHeader title="ไอดีเกม" back={back} />
      <button className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: 12 }} onClick={openNew}><Plus size={15} /> เพิ่มไอดีใหม่</button>
      {data.gameAccounts.length > 0 && (
        <div style={{ position: "relative", marginBottom: 12 }}>
          <Search size={14} color="var(--muted)" style={{ position: "absolute", left: 12, top: 12 }} />
          <input className="pgs-input" style={{ paddingLeft: 32 }} placeholder="ค้นหาไอดีเกม..." value={q} onChange={e => setQ(e.target.value)} />
        </div>
      )}
      {accounts.length === 0 ? <EmptyState text={data.gameAccounts.length === 0 ? "ยังไม่มีไอดีเกม" : "ไม่พบไอดีที่ค้นหา"} /> : accounts.map(a => {
        const invested = stats.investByAccount[a.id] || 0;
        const income = data.orders.filter(o => o.sourceAccountId === a.id && !o.cancelled && o.paymentStatus === "paid").reduce((s, o) => s + Number(o.price || 0), 0);
        const profit = income - invested;
        const pokemonCount = data.orders.filter(o => o.sourceAccountId === a.id && o.type === "sell_pokemon" && !o.cancelled).length;
        const lowStock = (a.stock || []).filter(s => clamp0(s.quantity) <= (s.lowStockThreshold ?? 2));
        return (
          <div key={a.id} className="pgs-card" style={{ marginBottom: 8, cursor: "pointer" }} onClick={() => openDetail(a)}>
            <div className="pgs-row" style={{ marginBottom: 8 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <Gamepad2 size={16} color="var(--yellow)" />
                <span style={{ fontWeight: 700, fontSize: 14 }}>{a.name}</span>
              </div>
              <span className="pgs-mono" style={{ fontWeight: 700, fontSize: 13, color: profit >= 0 ? "var(--green)" : "var(--red)" }}>฿{fmtMoney(profit)}</span>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6, fontSize: 10, color: "var(--muted)" }}>
              <div>ลงทุน<br /><span className="pgs-mono" style={{ color: "var(--text)", fontSize: 12 }}>฿{fmtMoney(invested)}</span></div>
              <div>รายรับ<br /><span className="pgs-mono" style={{ color: "var(--text)", fontSize: 12 }}>฿{fmtMoney(income)}</span></div>
              <div>Pokémon<br /><span className="pgs-mono" style={{ color: "var(--text)", fontSize: 12 }}>{pokemonCount}</span></div>
            </div>
            {lowStock.length > 0 && (
              <div className="pgs-badge" style={{ marginTop: 8, background: "rgba(255,84,112,0.15)", color: "var(--red)" }}>
                <AlertTriangle size={10} /> สต๊อกใกล้หมด {lowStock.length} รายการ
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function AccountModal({ mode, item, onClose, onSave }) {
  const [form, setForm] = useState(item || { id: genId(), name: "", note: "", createdAt: new Date().toISOString() });
  return (
    <Modal title={mode === "add" ? "เพิ่มไอดีเกม" : "แก้ไขไอดีเกม"} onClose={onClose}>
      <div className="pgs-field">
        <label className="pgs-label">ชื่อไอดี</label>
        <input className="pgs-input" value={form.name} onChange={e => setForm(f => ({ ...f, name: e.target.value }))} placeholder="เช่น ID-Trainer01" />
      </div>
      <div className="pgs-field">
        <label className="pgs-label">หมายเหตุ</label>
        <textarea className="pgs-textarea" rows={2} value={form.note} onChange={e => setForm(f => ({ ...f, note: e.target.value }))} />
      </div>
      <button className="pgs-btn pgs-btn-primary" style={{ width: "100%" }} disabled={!form.name} onClick={() => form.name && onSave(form)}>บันทึก</button>
    </Modal>
  );
}

function AccountDetail({ item, data, stats, onClose, onEdit, onDelete, onAddInvestment, onDeleteInvestment, onAddStock, onEditStock }) {
  const invested = stats.investByAccount[item.id] || 0;
  const income = data.orders.filter(o => o.sourceAccountId === item.id && !o.cancelled && o.paymentStatus === "paid").reduce((s, o) => s + Number(o.price || 0), 0);
  const history = data.investmentHistory.filter(h => h.accountId === item.id).sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  const waitingTrade = data.orders.filter(o => o.sourceAccountId === item.id && !o.cancelled && o.tradeStatus === "waiting").length;
  const threeHearts = data.orders.filter(o => o.sourceAccountId === item.id && !o.cancelled && o.tradeStatus === "three_hearts").length;
  const stock = item.stock || [];
  const allOrderCount = data.orders.filter(o => o.sourceAccountId === item.id).length;
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [confirmInvId, setConfirmInvId] = useState(null);
  return (
    <Modal title={item.name} onClose={onClose}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 12 }}>
        <StatCard icon={Coins} label="ลงทุนสะสม" value={"฿" + fmtMoney(invested)} color="var(--yellow)" />
        <StatCard icon={TrendingUp} label="กำไร" value={"฿" + fmtMoney(income - invested)} color={income - invested >= 0 ? "var(--green)" : "var(--red)"} />
        <StatCard icon={Clock} label="ลูกค้ารอเทรด" value={waitingTrade} />
        <StatCard icon={Heart} label="ทำ 3 ใจ" value={threeHearts} color="var(--yellow)" />
      </div>
      <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 12 }} onClick={onAddInvestment}><Plus size={14} /> บันทึกเติม Coin / ซื้อ Pokémon</button>

      <div className="pgs-row" style={{ marginBottom: 8 }}>
        <div className="pgs-sectiontitle" style={{ margin: 0 }}>สต๊อก Pokémon</div>
        <button className="pgs-iconbtn" onClick={onAddStock}><Plus size={14} /></button>
      </div>
      {stock.length === 0 ? <EmptyState text="ยังไม่มีสต๊อก" /> : stock.map(s => {
        const low = clamp0(s.quantity) <= (s.lowStockThreshold ?? 2);
        return (
          <div key={s.id} className="pgs-row" style={{ padding: "8px 0", borderBottom: "1px solid var(--border)", fontSize: 12, cursor: "pointer" }} onClick={() => onEditStock(s)}>
            <div>
              <div style={{ fontWeight: 600 }}>{s.name} {(s.variants || []).filter(v => v !== "normal").map(v => POKEMON_VARIANTS[v]?.emoji).join("")}</div>
              <div style={{ color: "var(--muted)", fontSize: 10 }}>{(s.variants || []).map(v => POKEMON_VARIANTS[v]?.label).join(", ")}</div>
            </div>
            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
              {low && <AlertTriangle size={13} color="var(--red)" />}
              <span className="pgs-mono" style={{ fontWeight: 700, color: low ? "var(--red)" : "var(--text)" }}>{s.quantity}</span>
            </div>
          </div>
        );
      })}

      <div className="pgs-sectiontitle">ประวัติการลงทุน</div>
      {history.length === 0 ? <EmptyState text="ยังไม่มีประวัติ" /> : history.map(h => (
        <div key={h.id} className="pgs-row" style={{ padding: "8px 0", borderBottom: "1px solid var(--border)", fontSize: 12 }}>
          <div>
            <div style={{ fontWeight: 600 }}>{INVEST_TYPES[h.type].label}</div>
            <div style={{ color: "var(--muted)", fontSize: 10 }}>{fmtDate(h.date)}{h.note ? " · " + h.note : ""}</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span className="pgs-mono" style={{ color: "var(--red)" }}>-฿{fmtMoney(h.amount)}</span>
            {confirmInvId === h.id ? (
              <button onClick={() => { onDeleteInvestment(h.id); setConfirmInvId(null); }} className="pgs-btn pgs-btn-danger" style={{ padding: "4px 8px", fontSize: 10 }}>ยืนยัน?</button>
            ) : (
              <button onClick={() => setConfirmInvId(h.id)} style={{ background: "none", border: "none", cursor: "pointer" }}><Trash2 size={13} color="var(--muted)" /></button>
            )}
          </div>
        </div>
      ))}
      {confirmDelete && (allOrderCount > 0 || stock.length > 0) && (
        <div className="pgs-cancelbanner">
          <AlertTriangle size={14} /> ไอดีนี้มีออเดอร์ {allOrderCount} รายการ และสต๊อก {stock.length} รายการผูกอยู่ — ลบแล้วข้อมูลเหล่านั้นจะไม่แสดงชื่อไอดีนี้อีก
        </div>
      )}
      <div style={{ display: "flex", gap: 8, marginTop: 16 }}>
        <button className="pgs-btn pgs-btn-outline" style={{ flex: 1 }} onClick={onEdit}><Edit2 size={14} /> แก้ไข</button>
        {!confirmDelete ? (
          <button className="pgs-btn pgs-btn-danger" style={{ flex: 1 }} onClick={() => setConfirmDelete(true)}><Trash2 size={14} /> ลบไอดี</button>
        ) : (
          <button className="pgs-btn pgs-btn-danger" style={{ flex: 1 }} onClick={onDelete}>ยืนยันลบถาวร?</button>
        )}
      </div>
    </Modal>
  );
}

function StockModal({ mode, item, onClose, onSave, onDelete }) {
  const [form, setForm] = useState(item || { id: genId(), name: "", variants: ["normal"], quantity: 1, lowStockThreshold: 2 });
  const [confirmDelete, setConfirmDelete] = useState(false);
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  return (
    <Modal title={mode === "add" ? "เพิ่มสต๊อก Pokémon" : "แก้ไขสต๊อก"} onClose={onClose}>
      <div className="pgs-field">
        <label className="pgs-label">ชื่อ Pokémon</label>
        <input className="pgs-input" value={form.name} onChange={e => set("name", e.target.value)} placeholder="เช่น Rayquaza" />
      </div>
      <div className="pgs-field">
        <label className="pgs-label">ประเภท (เลือกได้ 1 ชนิดต่อรายการ)</label>
        <VariantChips value={form.variants} onChange={(v) => set("variants", v)} multi={false} />
        <div style={{ fontSize: 11, color: "var(--muted)", marginTop: 6 }}>
          Pokémon ตัวเดียวกันแต่คนละชนิด (เช่น ปกติ กับ Shiny) จำนวนคงเหลือมักไม่เท่ากัน — ให้เพิ่มเป็นรายการสต๊อกแยกกันสำหรับแต่ละชนิด
        </div>
      </div>
      <div style={{ display: "flex", gap: 10 }}>
        <div className="pgs-field" style={{ flex: 1 }}>
          <label className="pgs-label">จำนวนคงเหลือ</label>
          <input className="pgs-input pgs-mono" type="number" min="0" value={form.quantity} onChange={e => set("quantity", e.target.value)} />
        </div>
        <div className="pgs-field" style={{ flex: 1 }}>
          <label className="pgs-label">แจ้งเตือนเมื่อเหลือ ≤</label>
          <input className="pgs-input pgs-mono" type="number" min="0" value={form.lowStockThreshold} onChange={e => set("lowStockThreshold", e.target.value)} />
        </div>
      </div>
      <button
        className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: onDelete ? 8 : 0 }}
        disabled={!form.name}
        onClick={() => form.name && onSave({ ...form, quantity: clamp0(form.quantity), lowStockThreshold: clamp0(form.lowStockThreshold) })}
      >บันทึก</button>
      {onDelete && (
        !confirmDelete ? (
          <button className="pgs-btn pgs-btn-danger" style={{ width: "100%" }} onClick={() => setConfirmDelete(true)}><Trash2 size={14} /> ลบสต๊อกนี้</button>
        ) : (
          <button className="pgs-btn pgs-btn-danger" style={{ width: "100%" }} onClick={onDelete}>ยืนยันลบ?</button>
        )
      )}
    </Modal>
  );
}

// =================== TRADE ===================
function TradeTab({ data, custName, accName, openEdit, onQuickTrade }) {
  const [filter, setFilter] = useState("waiting");
  const [sortDir, setSortDir] = useState("asc"); // asc = เก่าสุดอยู่บน (มาก่อนได้ก่อน)
  const orders = data.orders
    .filter(o => o.type === "sell_pokemon" && !o.cancelled && o.tradeStatus === filter)
    .slice()
    .sort((a, b) => {
      const diff = (a.createdAt || "").localeCompare(b.createdAt || "");
      return sortDir === "asc" ? diff : -diff;
    });
  return (
    <div>
      <h2 className="pgs-display" style={{ fontSize: 20, fontWeight: 700, margin: "0 0 12px 0" }}>ระบบเทรด</h2>
      <div className="pgs-row" style={{ marginBottom: 12, gap: 6 }}>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          <button className={"pgs-chip" + (filter === "waiting" ? " active" : "")} onClick={() => setFilter("waiting")}>รอเทรด</button>
          <button className={"pgs-chip" + (filter === "three_hearts" ? " active" : "")} onClick={() => setFilter("three_hearts")}>ทำ 3 ใจ</button>
          <button className={"pgs-chip" + (filter === "traded" ? " active" : "")} onClick={() => setFilter("traded")}>เทรดแล้ว</button>
        </div>
      </div>
      {(filter === "waiting" || filter === "three_hearts") && orders.length > 0 && (
        <button
          className="pgs-btn pgs-btn-outline"
          style={{ marginBottom: 10, fontSize: 12, padding: "6px 12px" }}
          onClick={() => setSortDir(d => d === "asc" ? "desc" : "asc")}
        >
          <ListFilter size={13} /> {sortDir === "asc" ? "มาก่อน อยู่บนสุด" : "มาใหม่ อยู่บนสุด"}
        </button>
      )}
      {orders.length === 0 ? <EmptyState text="ไม่มีรายการในสถานะนี้" /> : orders.map((o, i) => {
        const remain = o.appointmentDate ? daysBetween(todayStr(), o.appointmentDate) : null;
        return (
          <div key={o.id} className="pgs-card" style={{ marginBottom: 8, cursor: "pointer" }} onClick={() => openEdit(o)}>
            <div className="pgs-row" style={{ marginBottom: 6 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                {(filter === "waiting" || filter === "three_hearts") && (
                  <span className="pgs-mono" style={{ fontSize: 11, color: "var(--muted)", width: 18, textAlign: "center" }}>#{i + 1}</span>
                )}
                <div>
                  <div style={{ fontWeight: 700, fontSize: 13 }}>{custName(o.customerId)}</div>
                  <div style={{ fontSize: 11, color: "var(--muted)" }}>{o.pokemonName} x{o.quantity} · {accName(o.sourceAccountId)}</div>
                </div>
              </div>
              <StatusDot trade={o.tradeStatus} />
            </div>
            {o.appointmentDate && (
              <div style={{ fontSize: 11, color: "var(--muted)", display: "flex", alignItems: "center", gap: 4 }}>
                <Clock size={11} /> นัด {fmtDate(o.appointmentDate)}
                {remain !== null && remain >= 0 && ` · เหลืออีก ${remain} วัน`}
                {remain !== null && remain < 0 && ` · เลยกำหนด ${Math.abs(remain)} วัน`}
              </div>
            )}
            {filter === "waiting" && (
              <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
                <button
                  className="pgs-btn pgs-btn-outline" style={{ padding: "6px 10px", fontSize: 11 }}
                  onClick={(e) => { e.stopPropagation(); onQuickTrade(o.id, "traded"); }}
                ><CheckCircle2 size={12} /> เทรดแล้ว</button>
              </div>
            )}
            {filter === "three_hearts" && (
              <div style={{ display: "flex", gap: 6, marginTop: 8, flexWrap: "wrap" }}>
                <button
                  className="pgs-btn pgs-btn-outline" style={{ padding: "6px 10px", fontSize: 11 }}
                  onClick={(e) => { e.stopPropagation(); onQuickTrade(o.id, "traded"); }}
                ><CheckCircle2 size={12} /> เทรดแล้ว</button>
                <button
                  className="pgs-btn pgs-btn-outline" style={{ padding: "6px 10px", fontSize: 11, borderColor: "rgba(255,203,5,0.4)" }}
                  onClick={(e) => { e.stopPropagation(); onQuickTrade(o.id, "waiting"); }}
                ><Heart size={12} /> ทำ 3 ใจครบ</button>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

// =================== HIRE (ตีบอส / เชิญตี) ===================
function HireTab({ data, custName, accName, openEdit, onQuickUse, onQuickHireStatus }) {
  const [filter, setFilter] = useState("ongoing");
  const orders = data.orders
    .filter(o => (o.type === "hire_boss" || o.type === "hire_invite") && !o.cancelled)
    .filter(o => filter === "ongoing" ? o.hireStatus !== "done" : o.hireStatus === "done")
    .slice()
    .sort((a, b) => (a.createdAt || "").localeCompare(b.createdAt || ""));

  return (
    <div>
      <h2 className="pgs-display" style={{ fontSize: 20, fontWeight: 700, margin: "0 0 12px 0" }}>ตีบอส / เชิญตี</h2>
      <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap" }}>
        <button className={"pgs-chip" + (filter === "ongoing" ? " active" : "")} onClick={() => setFilter("ongoing")}>ค้างอยู่</button>
        <button className={"pgs-chip" + (filter === "done" ? " active" : "")} onClick={() => setFilter("done")}>เสร็จสิ้น</button>
      </div>
      {orders.length === 0 ? <EmptyState text="ไม่มีรายการในสถานะนี้" /> : orders.map(o => {
        const total = clamp0(o.hireTotal);
        const used = clamp0(o.hireUsed);
        const isFull = total > 0 && used >= total;
        return (
          <div key={o.id} className="pgs-card" style={{ marginBottom: 8, cursor: "pointer" }} onClick={() => openEdit(o)}>
            <div className="pgs-row" style={{ marginBottom: 6 }}>
              <div>
                <div style={{ fontWeight: 700, fontSize: 13 }}>{custName(o.customerId)}</div>
                <div style={{ fontSize: 11, color: "var(--muted)" }}>
                  {ORDER_TYPES[o.type].emoji} {ORDER_TYPES[o.type].label} · {HIRE_MODES[o.hireMode]?.label || ""}
                  {o.sourceAccountId ? ` · ${accName(o.sourceAccountId)}` : ""}
                </div>
              </div>
              <span className="pgs-badge" style={{ background: HIRE_STATUS[o.hireStatus === "done" ? "done" : "ongoing"].color + "22", color: HIRE_STATUS[o.hireStatus === "done" ? "done" : "ongoing"].color }}>
                {HIRE_STATUS[o.hireStatus === "done" ? "done" : "ongoing"].label}
              </span>
            </div>
            <div className="pgs-row" style={{ marginTop: 4, background: "var(--surface2)", borderRadius: 10, padding: "6px 8px" }}>
              <span style={{ fontSize: 11, color: "var(--muted)" }}>ใช้ไปแล้ว <span className="pgs-mono" style={{ color: "var(--text)", fontWeight: 700 }}>{used}</span> / {total} ตัว</span>
              {filter === "ongoing" && (
                <div style={{ display: "flex", gap: 6 }}>
                  <button
                    className="pgs-iconbtn" style={{ padding: 5 }}
                    disabled={used <= 0}
                    onClick={(e) => { e.stopPropagation(); onQuickUse(o, -1); }}
                    title="ลดจำนวนที่ใช้ (แก้ไข)"
                  ><Minus size={12} /></button>
                  <button
                    className="pgs-iconbtn" style={{ padding: 5, borderColor: "rgba(255,203,5,0.4)" }}
                    disabled={total > 0 && used >= total}
                    onClick={(e) => { e.stopPropagation(); onQuickUse(o, 1); }}
                    title="ใช้ไปวันนี้ +1"
                  ><Plus size={12} /></button>
                </div>
              )}
            </div>
            {filter === "ongoing" && isFull && (
              <button
                className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginTop: 8, fontSize: 12, padding: "6px 10px" }}
                onClick={(e) => { e.stopPropagation(); onQuickHireStatus(o.id, "done"); }}
              ><CheckCircle2 size={13} /> เสร็จสิ้น (ครบจำนวนแล้ว)</button>
            )}
          </div>
        );
      })}
    </div>
  );
}

// =================== FINANCE ===================
function FinanceTab({ data, stats, custName, accName, openNew, back, onDeleteManual, openDetail }) {
  const [confirmTxId, setConfirmTxId] = useState(null);
  const [filter, setFilter] = useState("all");
  const ledger = useMemo(() => {
    const rows = [
      ...data.orders.filter(o => !o.cancelled && o.paymentStatus === "paid").map(o => ({
        id: "o_" + o.id, type: "income", label: `${ORDER_TYPES[o.type].label} - ${custName(o.customerId)}`,
        amount: Number(o.price) || 0, date: (o.paidDate || o.createdAt).slice(0, 10), source: "order",
      })),
      ...data.orders.filter(o => !o.cancelled && o.paymentStatus === "partial").map(o => ({
        id: "op_" + o.id, type: "income", label: `${ORDER_TYPES[o.type].label} - ${custName(o.customerId)} (ชำระบางส่วน)`,
        amount: Number(o.paidAmount) || 0, date: (o.paidDate || o.createdAt).slice(0, 10), source: "order",
      })),
      ...data.investmentHistory.map(h => ({
        id: "i_" + h.id, type: "expense", label: INVEST_TYPES[h.type].label, amount: Number(h.amount) || 0, date: h.date, source: "investment", accountId: h.accountId,
      })),
      ...data.manualTx.map(t => ({
        id: "m_" + t.id, type: t.type, label: t.category || "อื่นๆ", amount: Number(t.amount) || 0, date: t.date, source: "manual", rawId: t.id, accountId: t.accountId,
      })),
    ];
    return rows.filter(r => filter === "all" || r.type === filter).sort((a, b) => (b.date || "").localeCompare(a.date || ""));
  }, [data, filter, custName]);

  return (
    <div>
      <SubHeader title="การเงิน" back={back} />
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 12 }}>
        <StatCard icon={TrendingUp} label="รายรับเดือนนี้" value={"฿" + fmtMoney(stats.incomeMonth)} color="var(--green)" />
        <StatCard icon={TrendingDown} label="รายจ่ายเดือนนี้" value={"฿" + fmtMoney(stats.expenseMonth)} color="var(--red)" />
      </div>
      {stats.totalDue > 0 && (
        <button onClick={() => openDetail?.({ type: "debt" })} className="pgs-card" style={{ marginBottom: 12, borderColor: "rgba(255,84,112,0.4)", width: "100%", textAlign: "left", cursor: "pointer" }}>
          <div className="pgs-row">
            <span style={{ fontSize: 12, color: "var(--muted)" }}>ยอดค้างชำระรวมทั้งร้าน · แตะดูรายลูกค้า</span>
            <span className="pgs-mono" style={{ fontWeight: 700, fontSize: 16, color: "var(--red)" }}>฿{fmtMoney(stats.totalDue)}</span>
          </div>
        </button>
      )}
      <button className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: 12 }} onClick={openNew}><Plus size={15} /> เพิ่มรายการ (เติม Coin / ซื้อ Pokémon / อื่นๆ)</button>
      <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>
        <button className={"pgs-chip" + (filter === "all" ? " active" : "")} onClick={() => setFilter("all")}>ทั้งหมด</button>
        <button className={"pgs-chip" + (filter === "income" ? " active" : "")} onClick={() => setFilter("income")}>รายรับ</button>
        <button className={"pgs-chip" + (filter === "expense" ? " active" : "")} onClick={() => setFilter("expense")}>รายจ่าย</button>
      </div>
      {ledger.length === 0 ? <EmptyState text="ยังไม่มีรายการ" /> : ledger.map(r => (
        <div key={r.id} className="pgs-row" style={{ padding: "10px 0", borderBottom: "1px solid var(--border)" }}>
          <div>
            <div style={{ fontWeight: 600, fontSize: 13 }}>{r.label}</div>
            <div style={{ fontSize: 10, color: "var(--muted)" }}>{fmtDate(r.date)}{r.accountId ? ` · ${accName(r.accountId)}` : ""}</div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span className="pgs-mono" style={{ fontWeight: 700, color: r.type === "income" ? "var(--green)" : "var(--red)" }}>{r.type === "income" ? "+" : "-"}฿{fmtMoney(r.amount)}</span>
            {r.source === "manual" && (
              confirmTxId === r.rawId ? (
                <button onClick={() => { onDeleteManual(r.rawId); setConfirmTxId(null); }} className="pgs-btn pgs-btn-danger" style={{ padding: "4px 8px", fontSize: 10 }}>ยืนยัน?</button>
              ) : (
                <button onClick={() => setConfirmTxId(r.rawId)} style={{ background: "none", border: "none", cursor: "pointer" }}><Trash2 size={13} color="var(--muted)" /></button>
              )
            )}
          </div>
        </div>
      ))}
    </div>
  );
}

function TxModal({ data, onClose, onSaveInvestment, onSaveManual, presetAccount }) {
  const [mode, setMode] = useState("investment");
  const [invForm, setInvForm] = useState({ id: genId(), accountId: presetAccount || data.gameAccounts[0]?.id || "", type: "topup", amount: "", date: todayStr(), note: "" });
  const [manForm, setManForm] = useState({ id: genId(), type: "expense", category: "อื่นๆ", amount: "", date: todayStr(), note: "", accountId: presetAccount || "" });

  return (
    <Modal title="เพิ่มรายการการเงิน" onClose={onClose}>
      <div style={{ display: "flex", gap: 6, marginBottom: 14 }}>
        <button className={"pgs-chip" + (mode === "investment" ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => setMode("investment")}>เติม Coin / ซื้อ Pokémon</button>
        <button className={"pgs-chip" + (mode === "manual" ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => setMode("manual")}>รายการอื่นๆ</button>
      </div>

      {mode === "investment" ? (
        <>
          {data.gameAccounts.length === 0 ? <EmptyState text="กรุณาเพิ่มไอดีเกมก่อน" /> : (
            <>
              <div className="pgs-field">
                <label className="pgs-label">ไอดีเกม</label>
                <select className="pgs-select" value={invForm.accountId} onChange={e => setInvForm(f => ({ ...f, accountId: e.target.value }))}>
                  {data.gameAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
                </select>
              </div>
              <div className="pgs-field">
                <label className="pgs-label">ประเภท</label>
                <div style={{ display: "flex", gap: 6 }}>
                  {Object.entries(INVEST_TYPES).map(([k, v]) => (
                    <button key={k} className={"pgs-chip" + (invForm.type === k ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => setInvForm(f => ({ ...f, type: k }))}>{v.label}</button>
                  ))}
                </div>
              </div>
              <div className="pgs-field">
                <label className="pgs-label">จำนวนเงิน (บาท)</label>
                <input className="pgs-input pgs-mono" type="number" value={invForm.amount} onChange={e => setInvForm(f => ({ ...f, amount: e.target.value }))} />
              </div>
              <div className="pgs-field">
                <label className="pgs-label">วันที่</label>
                <input className="pgs-input" type="date" value={invForm.date} onChange={e => setInvForm(f => ({ ...f, date: e.target.value }))} />
              </div>
              <div className="pgs-field">
                <label className="pgs-label">หมายเหตุ</label>
                <input className="pgs-input" value={invForm.note} onChange={e => setInvForm(f => ({ ...f, note: e.target.value }))} />
              </div>
              <button className="pgs-btn pgs-btn-primary" style={{ width: "100%" }} disabled={!invForm.amount || !invForm.accountId} onClick={() => onSaveInvestment({ ...invForm, amount: Number(invForm.amount) })}>บันทึก</button>
            </>
          )}
        </>
      ) : (
        <>
          <div className="pgs-field">
            <label className="pgs-label">ประเภท</label>
            <div style={{ display: "flex", gap: 6 }}>
              <button className={"pgs-chip" + (manForm.type === "income" ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => setManForm(f => ({ ...f, type: "income" }))}>รายรับ</button>
              <button className={"pgs-chip" + (manForm.type === "expense" ? " active" : "")} style={{ flex: 1, textAlign: "center" }} onClick={() => setManForm(f => ({ ...f, type: "expense" }))}>รายจ่าย</button>
            </div>
          </div>
          <div className="pgs-field">
            <label className="pgs-label">รายการ</label>
            <input className="pgs-input" value={manForm.category} onChange={e => setManForm(f => ({ ...f, category: e.target.value }))} placeholder="เช่น ค่าธรรมเนียมโอน" />
          </div>
          <div className="pgs-field">
            <label className="pgs-label">จำนวนเงิน (บาท)</label>
            <input className="pgs-input pgs-mono" type="number" value={manForm.amount} onChange={e => setManForm(f => ({ ...f, amount: e.target.value }))} />
          </div>
          <div className="pgs-field">
            <label className="pgs-label">วันที่</label>
            <input className="pgs-input" type="date" value={manForm.date} onChange={e => setManForm(f => ({ ...f, date: e.target.value }))} />
          </div>
          <div className="pgs-field">
            <label className="pgs-label">เกี่ยวข้องกับไอดีเกม (ถ้ามี)</label>
            <select className="pgs-select" value={manForm.accountId} onChange={e => setManForm(f => ({ ...f, accountId: e.target.value }))}>
              <option value="">ไม่เกี่ยวกับไอดีใด</option>
              {data.gameAccounts.map(a => <option key={a.id} value={a.id}>{a.name}</option>)}
            </select>
          </div>
          <button className="pgs-btn pgs-btn-primary" style={{ width: "100%" }} disabled={!manForm.amount} onClick={() => onSaveManual({ ...manForm, amount: Number(manForm.amount) })}>บันทึก</button>
        </>
      )}
    </Modal>
  );
}

// =================== REPORTS ===================
const PIE_COLORS = ["#ffcb05", "#4d68e0", "#33c481", "#ff5470", "#8b8da6"];

function ReportsTab({ data, custName, accName, back }) {
  const monthly = useMemo(() => {
    const map = {};
    const push = (date, key, amt) => {
      const m = (date || "").slice(0, 7);
      if (!m) return;
      map[m] = map[m] || { month: m, income: 0, expense: 0 };
      map[m][key] += amt;
    };
    data.orders.filter(o => !o.cancelled && o.paymentStatus === "paid").forEach(o => push((o.paidDate || o.createdAt), "income", Number(o.price) || 0));
    data.orders.filter(o => !o.cancelled && o.paymentStatus === "partial").forEach(o => push((o.paidDate || o.createdAt), "income", Number(o.paidAmount) || 0));
    data.manualTx.forEach(t => push(t.date, t.type, Number(t.amount) || 0));
    data.investmentHistory.forEach(h => push(h.date, "expense", Number(h.amount) || 0));
    return Object.values(map).sort((a, b) => a.month.localeCompare(b.month)).slice(-6).map(m => ({ ...m, label: m.month.slice(5) + "/" + m.month.slice(2, 4) }));
  }, [data]);

  const paidAmountOf = (o) => o.paymentStatus === "paid" ? Number(o.price || 0) : (o.paymentStatus === "partial" ? Number(o.paidAmount || 0) : 0);

  const incomeByAccount = useMemo(() => {
    const map = {};
    data.orders.filter(o => !o.cancelled && o.sourceAccountId && paidAmountOf(o) > 0).forEach(o => {
      const name = accName(o.sourceAccountId);
      map[name] = (map[name] || 0) + paidAmountOf(o);
    });
    return Object.entries(map).map(([name, value]) => ({ name, value }));
  }, [data]);

  const incomeByType = useMemo(() => {
    const map = {};
    data.orders.filter(o => !o.cancelled && paidAmountOf(o) > 0).forEach(o => {
      const label = ORDER_TYPES[o.type].short;
      map[label] = (map[label] || 0) + paidAmountOf(o);
    });
    return Object.entries(map).map(([name, value]) => ({ name, value }));
  }, [data]);

  const topCustomers = useMemo(() => {
    const map = {};
    data.orders.filter(o => !o.cancelled && paidAmountOf(o) > 0).forEach(o => {
      map[o.customerId] = (map[o.customerId] || 0) + paidAmountOf(o);
    });
    return Object.entries(map).map(([id, amount]) => ({ name: custName(id), amount })).sort((a, b) => b.amount - a.amount).slice(0, 5);
  }, [data]);

  return (
    <div>
      <SubHeader title="รายงาน" back={back} />
      <div className="pgs-sectiontitle">รายรับ-รายจ่าย 6 เดือนล่าสุด</div>
      <div className="pgs-card" style={{ marginBottom: 16, height: 190 }}>
        {monthly.length === 0 ? <EmptyState text="ยังไม่มีข้อมูล" /> : (
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={monthly}>
              <CartesianGrid strokeDasharray="3 3" stroke="#2c2f42" vertical={false} />
              <XAxis dataKey="label" tick={{ fill: "#8b8da6", fontSize: 10 }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fill: "#8b8da6", fontSize: 10 }} axisLine={false} tickLine={false} width={40} />
              <Tooltip contentStyle={{ background: "#1b1d2a", border: "1px solid #2c2f42", borderRadius: 8, fontSize: 12 }} formatter={(v) => "฿" + fmtMoney(v)} />
              <Bar dataKey="income" fill="#33c481" radius={[4, 4, 0, 0]} name="รายรับ" />
              <Bar dataKey="expense" fill="#ff5470" radius={[4, 4, 0, 0]} name="รายจ่าย" />
            </BarChart>
          </ResponsiveContainer>
        )}
      </div>

      <div className="pgs-sectiontitle">รายได้แยกตามไอดี</div>
      <div className="pgs-card" style={{ marginBottom: 16 }}>
        {incomeByAccount.length === 0 ? <EmptyState text="ยังไม่มีข้อมูล" /> : (
          <div style={{ height: 180 }}>
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie data={incomeByAccount} dataKey="value" nameKey="name" innerRadius={40} outerRadius={70} paddingAngle={2}>
                  {incomeByAccount.map((_, i) => <Cell key={i} fill={PIE_COLORS[i % PIE_COLORS.length]} />)}
                </Pie>
                <Tooltip contentStyle={{ background: "#1b1d2a", border: "1px solid #2c2f42", borderRadius: 8, fontSize: 12 }} formatter={(v) => "฿" + fmtMoney(v)} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      <div className="pgs-sectiontitle">รายได้แยกตามประเภทบริการ</div>
      <div className="pgs-card" style={{ marginBottom: 16 }}>
        {incomeByType.map((t, i) => (
          <div key={t.name} className="pgs-row" style={{ marginBottom: 8 }}>
            <span style={{ fontSize: 12, display: "flex", alignItems: "center", gap: 6 }}><span style={{ width: 8, height: 8, borderRadius: "50%", background: PIE_COLORS[i % PIE_COLORS.length], display: "inline-block" }} />{t.name}</span>
            <span className="pgs-mono" style={{ fontSize: 12, fontWeight: 700 }}>฿{fmtMoney(t.value)}</span>
          </div>
        ))}
        {incomeByType.length === 0 && <EmptyState text="ยังไม่มีข้อมูล" />}
      </div>

      <div className="pgs-sectiontitle">ลูกค้าใช้จ่ายสูงสุด</div>
      <div className="pgs-card">
        {topCustomers.length === 0 ? <EmptyState text="ยังไม่มีข้อมูล" /> : topCustomers.map((c, i) => (
          <div key={c.name} className="pgs-row" style={{ padding: "6px 0" }}>
            <span style={{ fontSize: 12 }}>#{i + 1} {c.name}</span>
            <span className="pgs-mono" style={{ fontSize: 12, fontWeight: 700, color: "var(--green)" }}>฿{fmtMoney(c.amount)}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

// =================== SETTINGS ===================
function SettingsTab({ data, setData, onBackup, onRestore, onExportExcel, onExportPDF, showToast, back, googleSyncing, googleStatus, onConnectGoogle, onSyncNow, onDisconnectGoogle }) {
  const fileRef = useRef(null);
  const logoRef = useRef(null);
  const bgRef = useRef(null);
  const [cropSrc, setCropSrc] = useState(null); // data URL of the image currently open in the cropper
  const [cropTarget, setCropTarget] = useState(null); // "logo" | "receiptBg"
  const [pinForm, setPinForm] = useState({ current: "", next: "", confirm: "", question: data.settings.pinQuestion || "", answer: "" });
  const hasPin = !!data.settings.pin;

  function savePin() {
    if (hasPin && pinForm.current !== data.settings.pin) { showToast?.("รหัส PIN ปัจจุบันไม่ถูกต้อง"); return; }
    if (!/^\d{4,8}$/.test(pinForm.next)) { showToast?.("PIN ต้องเป็นตัวเลข 4-8 หลัก"); return; }
    if (pinForm.next !== pinForm.confirm) { showToast?.("ยืนยัน PIN ไม่ตรงกัน"); return; }
    const question = pinForm.question.trim();
    const typedAnswer = pinForm.answer.trim();
    if (question && !typedAnswer && question !== data.settings.pinQuestion) { showToast?.("กรุณาระบุคำตอบของคำถามกู้คืนด้วย"); return; }
    // keep the old answer if the question is unchanged and no new answer was typed; otherwise use what was typed (or clear if question was cleared)
    const answer = !question ? "" : (typedAnswer || (question === data.settings.pinQuestion ? data.settings.pinAnswer : ""));
    setData(d => ({ ...d, settings: { ...d.settings, pin: pinForm.next, pinQuestion: question, pinAnswer: answer } }));
    setPinForm({ current: "", next: "", confirm: "", question, answer: "" });
    showToast?.(hasPin ? "เปลี่ยน PIN แล้ว" : "ตั้งรหัส PIN แล้ว");
  }
  function removePin() {
    if (pinForm.current !== data.settings.pin) { showToast?.("รหัส PIN ปัจจุบันไม่ถูกต้อง"); return; }
    setData(d => ({ ...d, settings: { ...d.settings, pin: "", pinQuestion: "", pinAnswer: "" } }));
    setPinForm({ current: "", next: "", confirm: "", question: "", answer: "" });
    showToast?.("ปิดการล็อกด้วย PIN แล้ว");
  }

  const daysSinceBackup = data.settings.lastBackupAt ? daysBetween(data.settings.lastBackupAt, new Date().toISOString()) : null;
  const g = data.settings.google;
  const [clientIdInput, setClientIdInput] = useState(g.clientId || "");
  const connected = !!g.spreadsheetId;

  function pickFile(e, target) {
    const file = e.target.files[0];
    e.target.value = "";
    if (!file) return;
    if (!file.type.startsWith("image/")) { showToast?.("กรุณาเลือกไฟล์รูปภาพ"); return; }
    const reader = new FileReader();
    reader.onload = () => {
      setCropSrc(reader.result);
      setCropTarget(target);
    };
    reader.onerror = () => showToast?.("อ่านไฟล์รูปภาพไม่สำเร็จ");
    reader.readAsDataURL(file);
  }
  function handleCropConfirm(dataUrl) {
    if (cropTarget === "logo") {
      setData(d => ({ ...d, settings: { ...d.settings, logoDataUrl: dataUrl } }));
      showToast?.("เปลี่ยนโลโก้ร้านแล้ว");
    } else if (cropTarget === "receiptBg") {
      setData(d => ({ ...d, settings: { ...d.settings, receiptBgDataUrl: dataUrl } }));
      showToast?.("เปลี่ยนพื้นหลังใบเสร็จแล้ว");
    }
    setCropSrc(null); setCropTarget(null);
  }
  function handleCropCancel() {
    setCropSrc(null); setCropTarget(null);
  }
  function removeLogo() {
    setData(d => ({ ...d, settings: { ...d.settings, logoDataUrl: "" } }));
    showToast?.("ลบโลโก้แล้ว");
  }
  function removeReceiptBg() {
    setData(d => ({ ...d, settings: { ...d.settings, receiptBgDataUrl: "" } }));
    showToast?.("ลบพื้นหลังใบเสร็จแล้ว");
  }

  return (
    <div>
      <SubHeader title="ตั้งค่า" back={back} />
      <div className="pgs-field">
        <label className="pgs-label">ชื่อร้าน</label>
        <input className="pgs-input" value={data.settings.shopName} onChange={e => setData(d => ({ ...d, settings: { ...d.settings, shopName: e.target.value } }))} />
      </div>

      <div className="pgs-sectiontitle">โลโก้ร้าน</div>
      <div className="pgs-card" style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 8 }}>
        <ShopLogo logoDataUrl={data.settings.logoDataUrl} size={56} />
        <div style={{ flex: 1, fontSize: 11, color: "var(--muted)" }}>ใช้เป็นโลโก้ที่แสดงในหน้าแรกของแอป และเป็นไอคอนตอนเปิดแอป/เพิ่มลงหน้าจอโฮม อัปโหลดแล้วสามารถลากปรับตำแหน่ง/ซูมเองได้</div>
      </div>
      <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start" }} onClick={() => logoRef.current?.click()}>
        <Upload size={15} /> {data.settings.logoDataUrl ? "เปลี่ยนโลโก้ร้าน" : "อัปโหลดโลโก้ร้าน"}
      </button>
      {data.settings.logoDataUrl && (
        <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start" }} onClick={removeLogo}>
          <Trash2 size={15} /> ลบโลโก้ (ใช้ไอคอนเริ่มต้น)
        </button>
      )}
      <input ref={logoRef} type="file" accept="image/*" style={{ display: "none" }} onChange={e => pickFile(e, "logo")} />

      <div className="pgs-sectiontitle">พื้นหลังใบเสร็จ</div>
      <div className="pgs-card" style={{ display: "flex", alignItems: "center", gap: 14, marginBottom: 8 }}>
        {data.settings.receiptBgDataUrl ? (
          <img src={data.settings.receiptBgDataUrl} alt="พื้นหลังใบเสร็จ" style={{ width: 56, height: 74, borderRadius: 10, objectFit: "cover", flexShrink: 0, border: "1px solid var(--border)" }} />
        ) : (
          <div style={{ width: 56, height: 74, borderRadius: 10, background: "rgba(255,255,255,0.04)", border: "1px dashed var(--border)", flexShrink: 0 }} />
        )}
        <div style={{ flex: 1, fontSize: 11, color: "var(--muted)" }}>รูปพื้นหลังโปร่งใสของใบเสร็จ (การ์ดจะโปร่งแสงให้เห็นภาพนี้) เปลี่ยนได้ทุกเมื่อ</div>
      </div>
      <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start" }} onClick={() => bgRef.current?.click()}>
        <Upload size={15} /> {data.settings.receiptBgDataUrl ? "เปลี่ยนพื้นหลังใบเสร็จ" : "อัปโหลดพื้นหลังใบเสร็จ"}
      </button>
      {data.settings.receiptBgDataUrl && (
        <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start" }} onClick={removeReceiptBg}>
          <Trash2 size={15} /> ลบพื้นหลัง (ใช้ดีไซน์เริ่มต้น)
        </button>
      )}
      <input ref={bgRef} type="file" accept="image/*" style={{ display: "none" }} onChange={e => pickFile(e, "receiptBg")} />

      {cropSrc && cropTarget === "logo" && (
        <ImageCropModal src={cropSrc} aspect={1} shape="circle" outputW={512} title="ปรับโลโก้ร้าน" onCancel={handleCropCancel} onConfirm={handleCropConfirm} />
      )}
      {cropSrc && cropTarget === "receiptBg" && (
        <ImageCropModal src={cropSrc} aspect={0.75} shape="rect" outputW={720} title="ปรับพื้นหลังใบเสร็จ" onCancel={handleCropCancel} onConfirm={handleCropConfirm} />
      )}

      <div className="pgs-sectiontitle">ล็อกแอปด้วย PIN</div>
      <div className="pgs-card" style={{ marginBottom: 8 }}>
        <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 10 }}>
          {hasPin ? "แอปนี้ล็อกด้วย PIN อยู่ ต้องใส่รหัสทุกครั้งที่เปิดแอป" : "ยังไม่ได้ตั้ง PIN — ใครก็เปิดแอปนี้ดูข้อมูลร้านได้ทันที"}
          {hasPin && (data.settings.pinQuestion
            ? <> ตั้งคำถามกู้คืนไว้แล้ว: "<b>{data.settings.pinQuestion}</b>"</>
            : <> ยังไม่ได้ตั้งคำถามกู้คืน — ถ้าลืม PIN จะต้องกู้คืนด้วยไฟล์ Backup แทน</>)}
        </div>
        {hasPin && (
          <div className="pgs-field" style={{ marginBottom: 8 }}>
            <label className="pgs-label">PIN ปัจจุบัน</label>
            <input className="pgs-input pgs-mono" type="password" inputMode="numeric" value={pinForm.current} onChange={e => setPinForm(f => ({ ...f, current: e.target.value }))} />
          </div>
        )}
        <div className="pgs-field" style={{ marginBottom: 8 }}>
          <label className="pgs-label">{hasPin ? "PIN ใหม่" : "ตั้ง PIN (ตัวเลข 4-8 หลัก)"}</label>
          <input className="pgs-input pgs-mono" type="password" inputMode="numeric" value={pinForm.next} onChange={e => setPinForm(f => ({ ...f, next: e.target.value }))} />
        </div>
        <div className="pgs-field" style={{ marginBottom: 10 }}>
          <label className="pgs-label">ยืนยัน PIN</label>
          <input className="pgs-input pgs-mono" type="password" inputMode="numeric" value={pinForm.confirm} onChange={e => setPinForm(f => ({ ...f, confirm: e.target.value }))} />
        </div>
        <div className="pgs-field" style={{ marginBottom: 8 }}>
          <label className="pgs-label">คำถามกู้คืน (ถ้าลืม PIN) — ไม่บังคับ</label>
          <input className="pgs-input" placeholder="เช่น ชื่อเล่นตอนเด็ก" value={pinForm.question} onChange={e => setPinForm(f => ({ ...f, question: e.target.value }))} />
        </div>
        <div className="pgs-field" style={{ marginBottom: 10 }}>
          <label className="pgs-label">คำตอบ{data.settings.pinAnswer && pinForm.question === data.settings.pinQuestion ? " (เว้นว่าง = ใช้คำตอบเดิม)" : ""}</label>
          <input className="pgs-input" type="password" value={pinForm.answer} onChange={e => setPinForm(f => ({ ...f, answer: e.target.value }))} />
        </div>
        <button className="pgs-btn pgs-btn-primary" style={{ width: "100%", marginBottom: hasPin ? 8 : 0 }} onClick={savePin}>{hasPin ? "เปลี่ยน PIN" : "ตั้ง PIN"}</button>
        {hasPin && <button className="pgs-btn pgs-btn-danger" style={{ width: "100%" }} onClick={removePin}>ปิดการล็อก PIN</button>}
      </div>

      <div className="pgs-sectiontitle">สำรองข้อมูลอัตโนมัติ (Google Sheets + Drive)</div>
      <div className="pgs-card" style={{ marginBottom: 8 }}>
        {!connected ? (
          <>
            <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 10 }}>
              ข้อมูลลูกค้า/ออเดอร์จะถูกเขียนลง Google Sheets และรูปสลิป/รูปงานจะอัปโหลดขึ้น Google Drive ของร้านเองอัตโนมัติทุกครั้งที่มีการแก้ไข ต้องใส่ "Google Client ID" ของตัวเองก่อน (สร้างฟรีครั้งเดียว — ดูขั้นตอนใน README.md หัวข้อ "ตั้งค่า Google Sync")
            </div>
            <div className="pgs-field" style={{ marginBottom: 10 }}>
              <label className="pgs-label">Google Client ID</label>
              <input className="pgs-input pgs-mono" style={{ fontSize: 11 }} placeholder="xxxxxxxx.apps.googleusercontent.com" value={clientIdInput} onChange={e => setClientIdInput(e.target.value.trim())} />
            </div>
            <button className="pgs-btn pgs-btn-primary" style={{ width: "100%" }} disabled={!clientIdInput || googleSyncing} onClick={() => onConnectGoogle(clientIdInput)}>
              {googleSyncing ? "กำลังเชื่อมต่อ..." : "เชื่อมต่อ Google"}
            </button>
          </>
        ) : (
          <>
            <div className="pgs-row" style={{ marginBottom: 6 }}>
              <span style={{ color: "var(--muted)" }}>บัญชี</span>
              <span style={{ fontSize: 12 }}>{g.email || "-"}</span>
            </div>
            <div className="pgs-row" style={{ marginBottom: 6 }}>
              <span style={{ color: "var(--muted)" }}>ซิงค์อัตโนมัติ</span>
              <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12 }}>
                <input type="checkbox" checked={g.autoSync} onChange={e => setData(d => ({ ...d, settings: { ...d.settings, google: { ...d.settings.google, autoSync: e.target.checked } } }))} />
                เปิดใช้งาน
              </label>
            </div>
            <div className="pgs-row" style={{ marginBottom: 10 }}>
              <span style={{ color: "var(--muted)" }}>ซิงค์ล่าสุด</span>
              <span className="pgs-mono" style={{ fontSize: 12 }}>{g.lastSyncAt ? fmtDate(g.lastSyncAt) : "ยังไม่เคย"}</span>
            </div>
            {googleStatus && <div style={{ fontSize: 11, color: "var(--blue)", marginBottom: 10 }}>{googleStatus}</div>}
            <a className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start", textDecoration: "none" }} href={spreadsheetUrl(g.spreadsheetId)} target="_blank" rel="noreferrer">
              <FileDown size={15} /> เปิด Google Sheet
            </a>
            <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start" }} disabled={googleSyncing} onClick={onSyncNow}>
              <Upload size={15} /> {googleSyncing ? "กำลังซิงค์..." : "ซิงค์เดี๋ยวนี้"}
            </button>
            <button className="pgs-btn pgs-btn-danger" style={{ width: "100%" }} onClick={onDisconnectGoogle}>ยกเลิกการเชื่อมต่อ</button>
          </>
        )}
      </div>

      <div className="pgs-sectiontitle">ข้อมูล & สำรองข้อมูล</div>
      {(daysSinceBackup === null || daysSinceBackup >= 7) && (
        <div className="pgs-cancelbanner" style={{ background: "rgba(255,203,5,0.12)", color: "var(--yellow)", borderColor: "rgba(255,203,5,0.35)" }}>
          <AlertTriangle size={14} />
          {daysSinceBackup === null ? "ยังไม่เคย Backup ข้อมูลเลย แนะนำให้ Backup ไว้กันข้อมูลหาย" : `ไม่ได้ Backup มา ${daysSinceBackup} วันแล้ว ข้อมูลอยู่ในเครื่องนี้เครื่องเดียว แนะนำให้ Backup`}
        </div>
      )}
      {daysSinceBackup !== null && daysSinceBackup < 7 && (
        <div style={{ fontSize: 11, color: "var(--muted)", marginBottom: 8 }}>Backup ล่าสุด: {fmtDate(data.settings.lastBackupAt)}</div>
      )}
      <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start" }} onClick={onBackup}><Download size={15} /> Backup ข้อมูล (.json)</button>
      <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start" }} onClick={() => fileRef.current?.click()}><Upload size={15} /> Restore ข้อมูลจากไฟล์</button>
      <input ref={fileRef} type="file" accept="application/json" style={{ display: "none" }} onChange={onRestore} />
      <div style={{ fontSize: 10, color: "var(--muted)", marginBottom: 8 }}>ข้อมูลทั้งหมดเก็บอยู่ในเบราว์เซอร์นี้เท่านั้น (ไม่ sync ข้ามเครื่อง/เบราว์เซอร์) หากล้าง cache หรือเปลี่ยนเครื่องโดยไม่ได้ Backup ไว้ ข้อมูลจะหายถาวร</div>

      <div className="pgs-sectiontitle">ส่งออกรายงาน</div>
      <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start" }} onClick={onExportExcel}><FileDown size={15} /> Export Excel (.xlsx)</button>
      <button className="pgs-btn pgs-btn-outline" style={{ width: "100%", marginBottom: 8, justifyContent: "flex-start" }} onClick={onExportPDF}><Printer size={15} /> Export PDF (พิมพ์ / บันทึกเป็น PDF)</button>

      <div className="pgs-sectiontitle">สรุปฐานข้อมูล</div>
      <div className="pgs-card" style={{ fontSize: 12 }}>
        <div className="pgs-row" style={{ marginBottom: 6 }}><span style={{ color: "var(--muted)" }}>ลูกค้า</span><span className="pgs-mono">{data.customers.length}</span></div>
        <div className="pgs-row" style={{ marginBottom: 6 }}><span style={{ color: "var(--muted)" }}>ออเดอร์</span><span className="pgs-mono">{data.orders.length}</span></div>
        <div className="pgs-row" style={{ marginBottom: 6 }}><span style={{ color: "var(--muted)" }}>ไอดีเกม</span><span className="pgs-mono">{data.gameAccounts.length}</span></div>
        <div className="pgs-row" style={{ marginBottom: 6 }}><span style={{ color: "var(--muted)" }}>รายการสต๊อก Pokémon</span><span className="pgs-mono">{data.gameAccounts.reduce((s, a) => s + (a.stock || []).length, 0)}</span></div>
        <div className="pgs-row" style={{ marginBottom: 6 }}><span style={{ color: "var(--muted)" }}>ออเดอร์ที่ยกเลิก</span><span className="pgs-mono">{data.orders.filter(o => o.cancelled).length}</span></div>
        <div className="pgs-row"><span style={{ color: "var(--muted)" }}>ประวัติการลงทุน</span><span className="pgs-mono">{data.investmentHistory.length}</span></div>
      </div>
      <div style={{ fontSize: 10, color: "var(--muted)", textAlign: "center", marginTop: 20 }}>ข้อมูลถูกบันทึกอัตโนมัติในเครื่องนี้ทุกครั้งที่มีการแก้ไข</div>
    </div>
  );
}
