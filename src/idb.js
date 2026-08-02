// IndexedDB-backed storage — same shape as the old localStorage wrapper
// ({ get(key) / set(key, value) } returning { key, value } or null) so the
// rest of the app doesn't need to change, but with a MUCH higher quota.
//
// Why this exists: localStorage caps out around 5-10MB per site. Once the
// shop has a few hundred orders with attached slip photos, the old
// localStorage-based storage would silently fail to save (setItem throws
// and gets swallowed) — new data just stops persisting. IndexedDB gives
// hundreds of MB to a few GB depending on the browser/device, which is what
// a shop doing ~1000 orders/year with photos actually needs.

const DB_NAME = "pgs-db";
const DB_VERSION = 1;
const STORE = "kv";

let dbPromise = null;

function openDb() {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    if (typeof indexedDB === "undefined") {
      reject(new Error("IndexedDB not available"));
      return;
    }
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
  return dbPromise;
}

export const idbStorage = {
  async get(key) {
    try {
      const db = await openDb();
      return await new Promise((resolve, reject) => {
        const tx = db.transaction(STORE, "readonly");
        const req = tx.objectStore(STORE).get(key);
        req.onsuccess = () => {
          const value = req.result;
          resolve(value === undefined ? null : { key, value });
        };
        req.onerror = () => reject(req.error);
      });
    } catch (e) {
      console.error("idbStorage.get failed", e);
      return null;
    }
  },
  async set(key, value) {
    try {
      const db = await openDb();
      return await new Promise((resolve, reject) => {
        const tx = db.transaction(STORE, "readwrite");
        tx.objectStore(STORE).put(value, key);
        tx.oncomplete = () => resolve({ key, value });
        tx.onerror = () => reject(tx.error);
      });
    } catch (e) {
      console.error("idbStorage.set failed", e);
      return null;
    }
  },
};

// One-time migration: if the old localStorage key still has data and
// IndexedDB doesn't yet, copy it over. Safe to call every startup — it's a
// no-op once migrated. Does NOT delete the localStorage copy automatically
// (kept as a safety net for a few versions), but the app stops writing to it.
export async function migrateFromLocalStorage(storageKey) {
  try {
    const existing = await idbStorage.get(storageKey);
    if (existing && existing.value) return false; // already migrated
    const old = window.localStorage.getItem(storageKey);
    if (!old) return false;
    await idbStorage.set(storageKey, old);
    return true;
  } catch (e) {
    console.error("migrateFromLocalStorage failed", e);
    return false;
  }
}
