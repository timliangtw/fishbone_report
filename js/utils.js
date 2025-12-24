
// -----------------------------
// Global error hooks (better than silent "Script error")
// -----------------------------
(function attachGlobalErrorBanner() {
    const show = (msg) => {
        const banner = document.getElementById('error-banner');
        const text = document.getElementById('error-banner-text');
        if (!banner || !text) return;
        text.textContent = String(msg || 'Unknown error');
        banner.style.display = 'block';
    };

    window.addEventListener('error', (e) => {
        // In some cross-origin cases, browsers only expose "Script error.".
        const parts = [e.message || 'Script error.'];
        if (e.filename) parts.push(`@ ${e.filename}:${e.lineno || 0}:${e.colno || 0}`);
        show(parts.join(' '));
    });

    window.addEventListener('unhandledrejection', (e) => {
        show(e?.reason?.message || e?.reason || 'Unhandled promise rejection');
    });
})();

// --- Config ---
window.LOCAL_STORAGE_KEY = 'fishbone_pro_data_v1';
window.IDB_NAME = 'fishbone_pro_idb_v1';

// -----------------------------
// IndexedDB helpers (store FileSystemFileHandle for true overwrite-save in Chrome)
// -----------------------------
function openKeyValueDb() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(window.IDB_NAME, 1);
        req.onupgradeneeded = () => {
            const db = req.result;
            if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv');
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

window.idbGet = async function (key) {
    const db = await openKeyValueDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('kv', 'readonly');
        const store = tx.objectStore('kv');
        const req = store.get(key);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
        tx.oncomplete = () => db.close();
    });
};

window.idbSet = async function (key, value) {
    const db = await openKeyValueDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('kv', 'readwrite');
        const store = tx.objectStore('kv');
        const req = store.put(value, key);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
        tx.oncomplete = () => db.close();
    });
};

window.idbDel = async function (key) {
    const db = await openKeyValueDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('kv', 'readwrite');
        const store = tx.objectStore('kv');
        const req = store.delete(key);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
        tx.oncomplete = () => db.close();
    });
};

// -----------------------------
// File System Access API availability
// -----------------------------
function isEmbeddedFrame() {
    try {
        return window.top !== window.self;
    } catch (_) {
        return true;
    }
}

window.isEmbeddedFrame = isEmbeddedFrame;

window.canUseNativeFilePicker = function () {
    return !!window.showSaveFilePicker && !!window.isSecureContext && !isEmbeddedFrame();
};

window.canUseNativeOpenPicker = function () {
    return !!window.showOpenFilePicker && !!window.isSecureContext && !isEmbeddedFrame();
};

window.ensureReadWritePermission = async function (handle) {
    if (!handle) throw new Error('Missing file handle');
    if (typeof handle.queryPermission !== 'function' || typeof handle.requestPermission !== 'function') {
        return;
    }
    const q = await handle.queryPermission({ mode: 'readwrite' });
    if (q === 'granted') return;
    const r = await handle.requestPermission({ mode: 'readwrite' });
    if (r !== 'granted') throw new Error('Permission denied for readwrite');
};

window.STATUS_OPTIONS = {
    CANT_EXECUTE: { id: 'CANT_EXECUTE', label: '⛔ Can\'t Execute', icon: '⛔', color: '#94a3b8', boneColor: '#94a3b8', bg: 'bg-gray-100 text-gray-500' },
    ON_GOING: { id: 'ON_GOING', label: '⏳ On Going', icon: '⏳', color: '#3b82f6', boneColor: '#60a5fa', bg: 'bg-blue-50 text-blue-700' },
    FINISHED: { id: 'FINISHED', label: '✅ Finished', icon: '✅', color: '#16a34a', boneColor: '#22c55e', bg: 'bg-green-50 text-green-700' },
    EXCLUDE: { id: 'EXCLUDE', label: '❌ Exclude', icon: '❌', color: '#cbd5e1', boneColor: '#94a3b8', bg: 'bg-gray-50 text-gray-400 line-through' }
};

window.DEMO_DATA = [
    { id: 1, category: "能源/電力", cause: "點火系統供電", factor: "電池沒電 (更換測試)", method: "更換全新電池", evidence: "成功點燃", status: 'FINISHED', isPriority: true },
    { id: 2, category: "能源/電力", cause: "點火系統供電", factor: "電池盒彈簧嚴重鏽蝕", method: "目視檢查", evidence: "正常", status: 'EXCLUDE', isPriority: false },
    { id: 21, category: "能源/電力", cause: "點火系統供電", factor: "電池接點接觸不良", method: "清潔接點", evidence: "無效", status: 'EXCLUDE', isPriority: false },
    { id: 3, category: "能源/電力", cause: "瓦斯供應異常", factor: "瓦斯桶沒氣", method: "搖晃確認", evidence: "有氣", status: 'EXCLUDE', isPriority: false },
    { id: 4, category: "設備硬體", cause: "點火針故障", factor: "點火針積碳", method: "清潔測試", evidence: "無效", status: 'CANT_EXECUTE', isPriority: false },
    { id: 5, category: "設備硬體", cause: "水盤異常", factor: "皮膜破裂", method: "拆機", evidence: "待確認", status: 'ON_GOING', isPriority: true },
    { id: 6, category: "人為操作", cause: "模式設定錯誤", factor: "誤切到夏天模式", method: "檢查面板", evidence: "正常", status: 'EXCLUDE', isPriority: false },
];

window.getTextWidth = (text, fontSize = 12) => {
    const canvas = window.getTextWidth.canvas || (window.getTextWidth.canvas = document.createElement("canvas"));
    const context = canvas.getContext("2d");
    context.font = `bold ${fontSize}px system-ui`;
    const metrics = context.measureText(text || "");
    return metrics.width + 10;
};
