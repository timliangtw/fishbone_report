const { useState, useMemo, useRef, useEffect, useCallback } = React;

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

// --- Icons ---
const Plus = ({ size = 20 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" /></svg>;
const Trash2 = ({ size = 20 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6" /><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" /><line x1="10" y1="11" x2="10" y2="17" /><line x1="14" y1="11" x2="14" y2="17" /></svg>;
const Save = ({ size = 20 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" /><polyline points="17 21 17 13 7 13 7 21" /><polyline points="7 3 7 8 15 8" /></svg>;
const FolderOpen = ({ size = 20 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z" /></svg>;
const Printer = ({ size = 20 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="6 9 6 2 18 2 18 9" /><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2" /><rect x="6" y="14" width="12" height="8" /></svg>;
const Maximize = ({ size = 20 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M8 3H5a2 2 0 0 0-2 2v3" /><path d="M21 8V5a2 2 0 0 0-2-2h-3" /><path d="M3 16v3a2 2 0 0 0 2 2h3" /><path d="M16 21h3a2 2 0 0 0 2-2v-3" /></svg>;
const FileText = ({ size = 20 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" /><polyline points="14 2 14 8 20 8" /><line x1="16" y1="13" x2="8" y2="13" /><line x1="16" y1="17" x2="8" y2="17" /><polyline points="10 9 9 9 8 9" /></svg>;
const Type = ({ size = 20 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="4 7 4 4 20 4 20 7" /><line x1="9" y1="20" x2="15" y2="20" /><line x1="12" y1="4" x2="12" y2="20" /></svg>;
const Star = ({ size = 16, filled = false }) => <svg width={size} height={size} viewBox="0 0 24 24" fill={filled ? "#ea580c" : "none"} stroke={filled ? "#ea580c" : "currentColor"} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2" /></svg>;
const Check = ({ size = 16 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12" /></svg>;
const GripVertical = ({ size = 20 }) => <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="9" cy="12" r="1" /><circle cx="9" cy="5" r="1" /><circle cx="9" cy="19" r="1" /><circle cx="15" cy="12" r="1" /><circle cx="15" cy="5" r="1" /><circle cx="15" cy="19" r="1" /></svg>;

// Editor Icons (kept; not used directly with Quill toolbar)
const BoldIcon = () => <span className="font-bold font-serif">B</span>;
const ItalicIcon = () => <span className="italic font-serif">I</span>;
const UnderlineIcon = () => <span className="underline font-serif">U</span>;
const StrikeIcon = () => <span className="line-through font-serif">S</span>;
const ListIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="8" y1="6" x2="21" y2="6" /><line x1="8" y1="12" x2="21" y2="12" /><line x1="8" y1="18" x2="21" y2="18" /><line x1="3" y1="6" x2="3.01" y2="6" /><line x1="3" y1="12" x2="3.01" y2="12" /><line x1="3" y1="18" x2="3.01" y2="18" /></svg>;
const ListOrderedIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="10" y1="6" x2="21" y2="6" /><line x1="10" y1="12" x2="21" y2="12" /><line x1="10" y1="18" x2="21" y2="18" /><path d="M4 6h1v4" /><path d="M4 10h2" /><path d="M6 18H4c0-1 2-2 2-3s-1-1.5-2-1" /></svg>;
const AlignLeft = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="17" y1="10" x2="3" y2="10" /><line x1="21" y1="6" x2="3" y2="6" /><line x1="21" y1="14" x2="3" y2="14" /><line x1="17" y1="18" x2="3" y2="18" /></svg>;
const AlignCenter = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="18" y1="10" x2="6" y2="10" /><line x1="21" y1="6" x2="3" y2="6" /><line x1="21" y1="14" x2="3" y2="14" /><line x1="18" y1="18" x2="6" y2="18" /></svg>;
const AlignRight = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><line x1="21" y1="10" x2="7" y2="10" /><line x1="21" y1="6" x2="3" y2="6" /><line x1="21" y1="14" x2="3" y2="14" /><line x1="21" y1="18" x2="7" y2="18" /></svg>;
const UndoIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M3 7v6h6" /><path d="M21 17a9 9 0 0 0-9-9 9 9 0 0 0-6 2.3L3 13" /></svg>;
const RedoIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><path d="M21 7v6h-6" /><path d="M3 17a9 9 0 0 1 9-9 9 9 0 0 1 6 2.3l3 2.7" /></svg>;
const ImageIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2" /><circle cx="8.5" cy="8.5" r="1.5" /><polyline points="21 15 16 10 5 21" /></svg>;
const HighlighterIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M9 11l6 6" /><path d="M14 3l7 7-4 4-7-7z" /><path d="M3 21l6-2 2-6-6-6-2 6z" /></svg>;
const LinkIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M10 13a5 5 0 0 1 0-7l1.2-1.2a5 5 0 0 1 7 7L17 13" /><path d="M14 11a5 5 0 0 1 0 7l-1.2 1.2a5 5 0 0 1-7-7L7 11" /></svg>;
const UnlinkIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 7h.01" /><path d="M7 17h.01" /><path d="M12 12l7-7" /><path d="M3 21l7-7" /><path d="M10 14a5 5 0 0 1 0-7l1.2-1.2a5 5 0 0 1 7 7L17 13" /><path d="M14 11a5 5 0 0 1 0 7l-1.2 1.2a5 5 0 0 1-7-7L7 11" /></svg>;
const EraserIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M20 20H9" /><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 0 0-4Z" /><path d="M12 14l4 4" /></svg>;
const IndentIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 6H11" /><path d="M21 12H11" /><path d="M21 18H11" /><path d="M3 8l4 4-4 4" /></svg>;
const OutdentIcon = () => <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 6H11" /><path d="M21 12H11" /><path d="M21 18H11" /><path d="M7 8l-4 4 4 4" /></svg>;

// --- Config ---
const LOCAL_STORAGE_KEY = 'fishbone_pro_data_v1';
const IDB_NAME = 'fishbone_pro_idb_v1';

// -----------------------------
// IndexedDB helpers (store FileSystemFileHandle for true overwrite-save in Chrome)
// -----------------------------
function openKeyValueDb() {
    return new Promise((resolve, reject) => {
        const req = indexedDB.open(IDB_NAME, 1);
        req.onupgradeneeded = () => {
            const db = req.result;
            if (!db.objectStoreNames.contains('kv')) db.createObjectStore('kv');
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });
}

async function idbGet(key) {
    const db = await openKeyValueDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('kv', 'readonly');
        const store = tx.objectStore('kv');
        const req = store.get(key);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
        tx.oncomplete = () => db.close();
    });
}

async function idbSet(key, value) {
    const db = await openKeyValueDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('kv', 'readwrite');
        const store = tx.objectStore('kv');
        const req = store.put(value, key);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
        tx.oncomplete = () => db.close();
    });
}

async function idbDel(key) {
    const db = await openKeyValueDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction('kv', 'readwrite');
        const store = tx.objectStore('kv');
        const req = store.delete(key);
        req.onsuccess = () => resolve(true);
        req.onerror = () => reject(req.error);
        tx.oncomplete = () => db.close();
    });
}

// -----------------------------
// File System Access API availability
// - showSaveFilePicker is blocked inside iframes (especially cross-origin) and requires secure context.
// -----------------------------
function isEmbeddedFrame() {
    try {
        return window.top !== window.self;
    } catch (_) {
        // Accessing window.top may throw in cross-origin iframes
        return true;
    }
}

function canUseNativeFilePicker() {
    return !!window.showSaveFilePicker && !!window.isSecureContext && !isEmbeddedFrame();
}

function canUseNativeOpenPicker() {
    return !!window.showOpenFilePicker && !!window.isSecureContext && !isEmbeddedFrame();
}

async function ensureReadWritePermission(handle) {
    if (!handle) throw new Error('Missing file handle');
    if (typeof handle.queryPermission !== 'function' || typeof handle.requestPermission !== 'function') {
        // Some environments may not expose permission methods; attempt write and let it fail if needed.
        return;
    }
    const q = await handle.queryPermission({ mode: 'readwrite' });
    if (q === 'granted') return;
    const r = await handle.requestPermission({ mode: 'readwrite' });
    if (r !== 'granted') throw new Error('Permission denied for readwrite');
}


const STATUS_OPTIONS = {
    CANT_EXECUTE: { id: 'CANT_EXECUTE', label: '⛔ Can\'t Execute', icon: '⛔', color: '#94a3b8', boneColor: '#94a3b8', bg: 'bg-gray-100 text-gray-500' },
    ON_GOING: { id: 'ON_GOING', label: '⏳ On Going', icon: '⏳', color: '#3b82f6', boneColor: '#60a5fa', bg: 'bg-blue-50 text-blue-700' },
    FINISHED: { id: 'FINISHED', label: '✅ Finished', icon: '✅', color: '#16a34a', boneColor: '#22c55e', bg: 'bg-green-50 text-green-700' },
    EXCLUDE: { id: 'EXCLUDE', label: '❌ Exclude', icon: '❌', color: '#cbd5e1', boneColor: '#94a3b8', bg: 'bg-gray-50 text-gray-400 line-through' }
};

const DEMO_DATA = [
    { id: 1, category: "能源/電力", cause: "點火系統供電", factor: "電池沒電 (更換測試)", method: "更換全新電池", evidence: "成功點燃", status: 'FINISHED', isPriority: true },
    { id: 2, category: "能源/電力", cause: "點火系統供電", factor: "電池盒彈簧嚴重鏽蝕", method: "目視檢查", evidence: "正常", status: 'EXCLUDE', isPriority: false },
    { id: 21, category: "能源/電力", cause: "點火系統供電", factor: "電池接點接觸不良", method: "清潔接點", evidence: "無效", status: 'EXCLUDE', isPriority: false },
    { id: 3, category: "能源/電力", cause: "瓦斯供應異常", factor: "瓦斯桶沒氣", method: "搖晃確認", evidence: "有氣", status: 'EXCLUDE', isPriority: false },
    { id: 4, category: "設備硬體", cause: "點火針故障", factor: "點火針積碳", method: "清潔測試", evidence: "無效", status: 'CANT_EXECUTE', isPriority: false },
    { id: 5, category: "設備硬體", cause: "水盤異常", factor: "皮膜破裂", method: "拆機", evidence: "待確認", status: 'ON_GOING', isPriority: true },
    { id: 6, category: "人為操作", cause: "模式設定錯誤", factor: "誤切到夏天模式", method: "檢查面板", evidence: "正常", status: 'EXCLUDE', isPriority: false },
];

const getTextWidth = (text, fontSize = 12) => {
    const canvas = getTextWidth.canvas || (getTextWidth.canvas = document.createElement("canvas"));
    const context = canvas.getContext("2d");
    context.font = `bold ${fontSize}px system-ui`;
    const metrics = context.measureText(text || "");
    return metrics.width + 10;
};

// --- HoverExpandTextarea ---
const HoverTextarea = ({ value, onChange, placeholder }) => {
    const [isExpanded, setIsExpanded] = useState(false);
    const ref = useRef(null);

    useEffect(() => {
        const el = ref.current;
        if (!el) return;
        if (isExpanded) {
            el.style.height = 'auto';
            el.style.height = `${el.scrollHeight + 2}px`;
            el.style.width = 'auto';
            el.style.minWidth = '100%';
        } else {
            el.style.height = '100%';
            el.style.width = '100%';
        }
    }, [value, isExpanded]);

    return (
        <div className="relative h-12 w-full">
            <textarea
                ref={ref}
                value={value}
                onChange={onChange}
                onMouseEnter={() => setIsExpanded(true)}
                onMouseLeave={() => { if (document.activeElement !== ref.current) setIsExpanded(false); }}
                onFocus={() => setIsExpanded(true)}
                onBlur={() => setIsExpanded(false)}
                placeholder={placeholder}
                className={`
                            w-full resize-none outline-none px-2 py-1.5 text-sm rounded transition-all duration-150
                            ${isExpanded
                        ? 'absolute top-0 left-0 z-50 bg-white shadow-xl border border-blue-300 min-h-[48px]'
                        : 'bg-transparent border-b border-transparent h-full overflow-hidden text-slate-600'
                    }
                        `}
                rows={1}
            />
        </div>
    );
};

// --- RichTextEditor (Quill-based, avoids React infinite update loop) ---
const RichTextEditor = ({ value, onChange }) => {
    const toolbarRef = useRef(null);
    const editorHostRef = useRef(null);
    const quillRef = useRef(null);
    const fileInputRef = useRef(null);

    const [selectedImg, setSelectedImg] = useState(null);
    const [imgWidth, setImgWidth] = useState(100);

    // Tracks what we believe the editor currently contains (as HTML)
    const lastHtmlRef = useRef(value || '');

    // When we programmatically paste/set HTML, Quill fires text-change.
    // This flag prevents a "ping-pong" loop: editor -> React state -> editor -> ...
    const applyingExternalRef = useRef(false);

    const applyImageSelection = (imgEl) => {
        try {
            const root = quillRef.current?.root;
            if (!root) return;
            root.querySelectorAll('img.selected').forEach(n => n.classList.remove('selected'));
            if (imgEl) {
                imgEl.classList.add('selected');
                setSelectedImg(imgEl);
                const currentWidth = imgEl.style.width || '100%';
                const parsed = parseInt(currentWidth, 10);
                setImgWidth(Number.isFinite(parsed) ? parsed : 100);
            } else {
                setSelectedImg(null);
            }
        } catch (_) {
            setSelectedImg(null);
        }
    };

    const safeEmitChange = (nextHtml) => {
        // Only emit if it actually changed (prevents maximum update depth)
        if (typeof nextHtml !== 'string') return;
        if (nextHtml === lastHtmlRef.current) return;
        lastHtmlRef.current = nextHtml;
        onChange(nextHtml);
    };

    const insertImageDataUrl = (dataUrl) => {
        const q = quillRef.current;
        if (!q) return;

        const sel = q.getSelection(true);
        const index = sel ? sel.index : q.getLength();

        q.insertEmbed(index, 'image', dataUrl, 'user');
        q.setSelection(index + 1, 0, 'silent');

        // Default styling. This direct DOM styling may not trigger Quill's delta,
        // so we emit a change only if HTML differs.
        setTimeout(() => {
            const root = q.root;
            const imgs = root.querySelectorAll('img');
            const last = imgs[imgs.length - 1];
            if (last) {
                last.style.width = '100%';
                last.style.borderRadius = '8px';
                last.style.margin = '10px 0';
                last.style.boxShadow = '0 4px 6px rgba(0,0,0,0.1)';
                safeEmitChange(q.root.innerHTML);
            }
        }, 0);
    };

    // Init Quill once
    useEffect(() => {
        if (!editorHostRef.current || !toolbarRef.current) return;
        if (!window.Quill) {
            console.error('Quill not loaded');
            return;
        }

        // Register width style attributor so image width persists across save/load.
        // Without this, Quill's clipboard conversion may drop inline width styles.
        try {
            if (!window.__fishboneQuillWidthRegistered) {
                const Parchment = window.Quill.import('parchment');
                const WidthStyle = new Parchment.Attributor.Style('width', 'width', {
                    scope: Parchment.Scope.INLINE,
                });
                window.Quill.register(WidthStyle, true);
                window.__fishboneQuillWidthRegistered = true;
            }
        } catch (e) {
            console.warn('Width attributor registration failed:', e);
        }

        const q = new window.Quill(editorHostRef.current, {
            theme: 'snow',
            placeholder: 'Write notes here…',
            modules: {
                toolbar: {
                    container: toolbarRef.current,
                    handlers: {
                        image: () => fileInputRef.current?.click(),
                    }
                }
            }
        });

        quillRef.current = q;

        // initial content
        const initial = value || '';
        if (initial) {
            applyingExternalRef.current = true;
            lastHtmlRef.current = initial;
            q.clipboard.dangerouslyPasteHTML(initial);
            // release in next tick after Quill finishes dispatching changes
            setTimeout(() => { applyingExternalRef.current = false; }, 0);
        } else {
            lastHtmlRef.current = '';
        }

        const onTextChange = () => {
            if (applyingExternalRef.current) {
                // Keep lastHtmlRef in sync without calling parent
                lastHtmlRef.current = q.root.innerHTML;
                return;
            }
            const html = q.root.innerHTML;
            safeEmitChange(html);
        };

        q.on('text-change', onTextChange);

        // Click-to-select image for resizing
        const onRootClick = (e) => {
            if (e.target?.tagName === 'IMG') {
                applyImageSelection(e.target);
            } else {
                applyImageSelection(null);
            }
        };

        // Paste image support (more deterministic than default)
        const onRootPaste = (e) => {
            const items = e.clipboardData?.items;
            if (!items) return;
            for (let i = 0; i < items.length; i++) {
                if (items[i].type && items[i].type.indexOf('image') !== -1) {
                    e.preventDefault();
                    const blob = items[i].getAsFile();
                    if (!blob) return;
                    const reader = new FileReader();
                    reader.onload = (ev) => {
                        const dataUrl = ev.target?.result;
                        if (typeof dataUrl === 'string') insertImageDataUrl(dataUrl);
                    };
                    reader.readAsDataURL(blob);
                    return;
                }
            }
        };

        q.root.addEventListener('click', onRootClick);
        q.root.addEventListener('paste', onRootPaste);

        return () => {
            try {
                q.off('text-change', onTextChange);
                q.root.removeEventListener('click', onRootClick);
                q.root.removeEventListener('paste', onRootPaste);
            } catch (_) { }
            quillRef.current = null;
            if (editorHostRef.current) editorHostRef.current.innerHTML = '';
        };
    }, []);

    // External updates (React state -> Quill)
    useEffect(() => {
        const q = quillRef.current;
        if (!q) return;

        const next = value || '';
        // If parent value equals what we already believe, do nothing
        if (next === lastHtmlRef.current) return;

        // Apply without causing a ping-pong loop
        applyingExternalRef.current = true;
        lastHtmlRef.current = next;

        const sel = q.getSelection();
        q.clipboard.dangerouslyPasteHTML(next);
        if (sel) q.setSelection(sel);

        setTimeout(() => { applyingExternalRef.current = false; }, 0);
    }, [value]);

    const handleImageUpload = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            const dataUrl = ev.target?.result;
            if (typeof dataUrl === 'string') insertImageDataUrl(dataUrl);
        };
        reader.readAsDataURL(file);
        e.target.value = '';
    };

    const updateImageSize = (newWidth) => {
        if (!selectedImg) return;
        setImgWidth(newWidth);

        const q = quillRef.current;
        const widthVal = `${newWidth}%`;

        // Apply via Quill format so it persists (Delta + clipboard).
        try {
            if (q && window.Quill && typeof window.Quill.find === 'function') {
                const blot = window.Quill.find(selectedImg);
                if (blot) {
                    const idx = q.getIndex(blot);
                    // Ensure selection is on the embed, then apply format.
                    q.setSelection(idx, 1, 'silent');
                    q.format('width', widthVal, 'user');
                }
            }
        } catch (e) {
            // Fallback to DOM style only
            console.warn('Quill width format failed, fallback to DOM style:', e);
        }

        // Always keep DOM style in sync for immediate UI.
        selectedImg.style.width = widthVal;

        // DOM style change may not trigger Quill event; emit only if changed
        if (q) safeEmitChange(q.root.innerHTML);
    };

    return (
        <div className="w-full bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden mt-6 print:border-none print:shadow-none">
            {/* Toolbar (Quill) */}
            <div className="bg-slate-50 border-b border-slate-200 px-2 py-2 print:hidden select-none">
                <div ref={toolbarRef} className="flex flex-wrap items-center gap-1">
                    {/* Header / Body */}
                    <select className="ql-header text-xs font-semibold" defaultValue="">
                        <option value="1">H1</option>
                        <option value="2">H2</option>
                        <option value="3">H3</option>
                        <option value="">Body</option>
                    </select>

                    <span className="w-px h-6 bg-slate-200 mx-1" />

                    {/* Colors */}
                    <select className="ql-color" />
                    <select className="ql-background" />

                    <span className="w-px h-6 bg-slate-200 mx-1" />

                    {/* Inline */}
                    <button className="ql-bold" />
                    <button className="ql-italic" />
                    <button className="ql-underline" />
                    <button className="ql-strike" />

                    <span className="w-px h-6 bg-slate-200 mx-1" />

                    {/* Lists */}
                    <button className="ql-list" value="bullet" />
                    <button className="ql-list" value="ordered" />
                    {/* Checklist: depending on Quill build/theme, value="check" may or may not render a checkbox style.
                               If it doesn't, we can implement a custom blot later. */}
                    <button className="ql-list" value="check" title="Checklist" />

                    <span className="w-px h-6 bg-slate-200 mx-1" />

                    {/* Link */}
                    <button className="ql-link" />

                    <span className="w-px h-6 bg-slate-200 mx-1" />

                    {/* Align / Indent */}
                    <select className="ql-align" />
                    <button className="ql-indent" value="-1" />
                    <button className="ql-indent" value="+1" />

                    <span className="w-px h-6 bg-slate-200 mx-1" />

                    {/* Super/Sub */}
                    <button className="ql-script" value="super" />
                    <button className="ql-script" value="sub" />

                    <span className="w-px h-6 bg-slate-200 mx-1" />

                    {/* Blocks */}
                    <button className="ql-blockquote" />
                    <button className="ql-code-block" />

                    <span className="w-px h-6 bg-slate-200 mx-1" />

                    {/* Image / Clean */}
                    <button className="ql-image" />
                    <button className="ql-clean" />
                </div>
                <input type="file" ref={fileInputRef} className="hidden" accept="image/*" onChange={handleImageUpload} />
            </div>

            {/* Image Context Toolbar */}
            {selectedImg && (
                <div className="bg-blue-50 border-b border-blue-100 px-3 py-1.5 flex items-center gap-4 text-sm animate-fade-in print:hidden">
                    <span className="text-blue-600 font-bold text-xs uppercase flex items-center gap-1">
                        <span className="inline-block w-2 h-2 rounded-full bg-blue-600" /> Image Size
                    </span>
                    <input
                        type="range"
                        min="10"
                        max="100"
                        value={imgWidth}
                        onChange={(e) => updateImageSize(Number(e.target.value))}
                        className="w-32 accent-blue-600 h-1.5"
                    />
                    <span className="font-mono text-blue-700 w-10">{imgWidth}%</span>
                    <div className="flex gap-1">
                        <button onClick={() => updateImageSize(25)} className="px-2 py-0.5 bg-white border border-blue-200 rounded text-xs text-blue-600 hover:bg-blue-100">S</button>
                        <button onClick={() => updateImageSize(50)} className="px-2 py-0.5 bg-white border border-blue-200 rounded text-xs text-blue-600 hover:bg-blue-100">M</button>
                        <button onClick={() => updateImageSize(75)} className="px-2 py-0.5 bg-white border border-blue-200 rounded text-xs text-blue-600 hover:bg-blue-100">L</button>
                        <button onClick={() => updateImageSize(100)} className="px-2 py-0.5 bg-white border border-blue-200 rounded text-xs text-blue-600 hover:bg-blue-100">Full</button>
                    </div>
                </div>
            )}

            {/* Editor */}
            <div className="editor-content">
                <div ref={editorHostRef} />
            </div>
        </div>
    );
};

const App = () => {
    const [problem, setProblem] = useState("熱水器不熱");
    const [problemDesc, setProblemDesc] = useState("熱水器完全沒有反應，有更換過電池");
    const [rows, setRows] = useState(DEMO_DATA);
    const [notes, setNotes] = useState("<div>Notes...</div>");
    const [scale, setScale] = useState(1);
    const [baseFontSize, setBaseFontSize] = useState(29);
    const [pan, setPan] = useState({ x: 0, y: 0 });
    const [isDragging, setIsDragging] = useState(false);
    const [dragStart, setDragStart] = useState({ x: 0, y: 0 });
    const [toast, setToast] = useState({ show: false, text: '' });
    const [saveHandle, setSaveHandle] = useState(null);
    const [saveTargetName, setSaveTargetName] = useState('');
    const [highlightRowId, setHighlightRowId] = useState(null);
    const [draggedRowIndex, setDraggedRowIndex] = useState(null);

    const embedded = useMemo(() => isEmbeddedFrame(), []);
    const nativeSaveAvailable = useMemo(() => {
        // showSaveFilePicker is not allowed in iframes; also needs secure context.
        return !!window.showSaveFilePicker && !!window.isSecureContext && !embedded;
    }, [embedded]);

    const nativeOpenAvailable = useMemo(() => {
        // showOpenFilePicker is also blocked in iframes and needs secure context.
        return !!window.showOpenFilePicker && !!window.isSecureContext && !embedded;
    }, [embedded]);
    const nativeSaveReason = useMemo(() => {
        if (nativeSaveAvailable) return '';
        if (embedded) return 'This preview is embedded (iframe). Chrome blocks showSaveFilePicker here.';
        if (!window.isSecureContext) return 'Not a secure context (needs https or localhost).';
        if (!window.showSaveFilePicker) return 'showSaveFilePicker not supported in this browser.';
        return 'Native save unavailable.';
    }, [nativeSaveAvailable, embedded]);

    const nativeOpenReason = useMemo(() => {
        if (nativeOpenAvailable) return '';
        if (embedded) return 'This preview is embedded (iframe). Chrome blocks showOpenFilePicker here.';
        if (!window.isSecureContext) return 'Not a secure context (needs https or localhost).';
        if (!window.showOpenFilePicker) return 'showOpenFilePicker not supported in this browser.';
        return 'Native open unavailable.';
    }, [nativeOpenAvailable, embedded]);

    const fileInputRef = useRef(null);
    const canvasRef = useRef(null);
    const highlightTimeoutRef = useRef(null);

    // Avoid stale closures when saving via keyboard shortcuts
    const saveHandleRef = useRef(null);
    const savingRef = useRef(false);
    const shortcutHitsRef = useRef(0);
    const [shortcutHits, setShortcutHits] = useState(0);
    const [lastShortcutAction, setLastShortcutAction] = useState('');
    const latestRef = useRef(null);
    const handleSaveCallRef = useRef(null);
    const allowDragRef = useRef(false);
    // Keep the newest snapshot for saving (prevents intermittent "saved old state" from stale closures)

    useEffect(() => {
        latestRef.current = { problem, problemDesc, rows, notes, baseFontSize };
    }, [problem, problemDesc, rows, notes, baseFontSize]);

    useEffect(() => {
        // Restore local data
        try {
            const savedData = localStorage.getItem(LOCAL_STORAGE_KEY);
            if (savedData) {
                const parsed = JSON.parse(savedData);
                if (parsed.problem) setProblem(parsed.problem);
                if (parsed.problemDesc) setProblemDesc(parsed.problemDesc);
                if (parsed.rows) setRows(parsed.rows);
                if (parsed.notes) setNotes(parsed.notes);
                if (parsed.baseFontSize) setBaseFontSize(parsed.baseFontSize);
            }
        } catch (_) {
            // ignore
        }
    }, []);

    useEffect(() => {
        // Restore last chosen save file handle (Chrome / Chromium File System Access API)
        (async () => {
            try {
                const h = await idbGet('saveHandle');
                if (h) setSaveHandle(h);
            } catch (_) { }
        })();
    }, []);

    useEffect(() => {
        // Keep ref in sync (used by Ctrl+S handler to avoid stale state)
        saveHandleRef.current = saveHandle;
    }, [saveHandle]);

    useEffect(() => {
        // Derive a friendly file name for UI (best-effort; may fail in some permission models)
        (async () => {
            try {
                if (!saveHandle) { setSaveTargetName(''); return; }
                const f = await saveHandle.getFile();
                setSaveTargetName(f?.name || '');
            } catch (_) {
                setSaveTargetName('');
            }
        })();
    }, [saveHandle]);

    useEffect(() => {
        // Keyboard shortcut: Ctrl+S / Cmd+S => Save.
        // Fixes:
        // 1) Use e.code===KeyS (layout-independent) + e.key fallback
        // 2) Ignore key repeat
        // 3) Prevent default browser "Save Page" dialog
        // 4) Use refs so we don't get a stale handleSave/saveHandle closure
        const onKeyDown = (e) => {
            const isS = (e.code === 'KeyS') || ((e.key || '').toLowerCase() === 's');
            if ((e.ctrlKey || e.metaKey) && isS) {
                if (e.repeat) return;
                try { e.preventDefault(); } catch (_) { }
                try { e.stopImmediatePropagation?.(); } catch (_) { }

                // If native save isn't available, don't trigger browser download via shortcut.
                if (!nativeSaveAvailable) {
                    shortcutHitsRef.current += 1;
                    setShortcutHits(shortcutHitsRef.current);
                    setLastShortcutAction('Blocked (native save unavailable)');
                    return;
                }

                // Guard against double-trigger while picker is opening / save is in-flight.
                if (savingRef.current) return;
                savingRef.current = true;

                shortcutHitsRef.current += 1;
                setShortcutHits(shortcutHitsRef.current);
                setLastShortcutAction('Saving…');

                Promise.resolve().then(async () => {
                    const fn = handleSaveCallRef.current || handleSave;
                    await fn();
                    setLastShortcutAction('Saved');
                }).catch((err) => {
                    console.error(err);
                    setLastShortcutAction('Save failed');
                }).finally(() => {
                    savingRef.current = false;
                });
            }
        };

        // Register on both window and document (some editors play with event handling)
        window.addEventListener('keydown', onKeyDown, true);
        document.addEventListener('keydown', onKeyDown, true);
        return () => {
            window.removeEventListener('keydown', onKeyDown, true);
            document.removeEventListener('keydown', onKeyDown, true);
        };
    }, [nativeSaveAvailable]);



    const scrollToRow = useCallback((rowId) => {
        if (!rowId) return;
        const el = document.getElementById(`verify-row-${rowId}`);
        if (!el) return;

        el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        setHighlightRowId(rowId);

        if (highlightTimeoutRef.current) {
            clearTimeout(highlightTimeoutRef.current);
        }
        highlightTimeoutRef.current = setTimeout(() => setHighlightRowId(null), 1800);
    }, []);

    const handleDragStart = (e, index) => {
        // Only allow dragging if initiated from the handle (via validation ref)
        if (!allowDragRef.current) {
            e.preventDefault();
            return;
        }
        setDraggedRowIndex(index);
        e.dataTransfer.effectAllowed = "move";
        // Required for Firefox and some other browsers to acknowledge the drag
        e.dataTransfer.setData('text/plain', index);
    };

    const handleDragEnd = () => {
        setDraggedRowIndex(null);
        allowDragRef.current = false;
    };

    const handleDragOver = (e, index) => {
        e.preventDefault(); // Necessary to allow dropping
        e.dataTransfer.dropEffect = "move";
    };

    const handleDrop = (e, targetIndex) => {
        e.preventDefault();
        if (draggedRowIndex === null || draggedRowIndex === targetIndex) return;

        const newRows = [...rows];
        const [draggedItem] = newRows.splice(draggedRowIndex, 1);
        newRows.splice(targetIndex, 0, draggedItem);
        setRows(newRows);
        setDraggedRowIndex(null);
    };

    const handleRowChange = (id, field, value) => setRows(prev => prev.map(r => r.id === id ? { ...r, [field]: value } : r));
    const insertRowAfter = (index) => {
        setRows(prev => {
            const currentRow = prev[index];
            const newId = prev.length > 0 ? Math.max(...prev.map(r => r.id)) + 1 : 1;
            const newRows = [...prev];
            newRows.splice(index + 1, 0, { ...currentRow, id: newId, factor: "New Factor", method: "", evidence: "", status: 'ON_GOING', isPriority: false });
            return newRows;
        });
    };
    const deleteRow = (id) => setRows(prev => prev.filter(r => r.id !== id));
    const addFirstRow = () => setRows([{ id: 1, category: "Primary", cause: "Secondary", factor: "Factor", method: "", evidence: "", status: 'ON_GOING', isPriority: false }]);
    const clearData = () => {
        if (confirm("Are you sure you want to clear all data?")) {
            setRows([]); setNotes(""); setProblem("Main Problem"); setProblemDesc("");
            try { localStorage.removeItem(LOCAL_STORAGE_KEY); } catch (e) { }
        }
    };
    const buildExportObject = () => {
        const s = latestRef.current || { problem, problemDesc, rows, notes, baseFontSize };
        return {
            problem: s.problem,
            problemDesc: s.problemDesc,
            rows: s.rows,
            notes: s.notes,
            baseFontSize: s.baseFontSize,
            version: "3.8-SaveConsistency",
            exportedAt: new Date().toISOString(),
        };
    };

    const saveLocalDraft = () => {
        // Manual draft save (replaces auto-save): only saved when user triggers Save / Ctrl+S / Download.
        try {
            const s = latestRef.current || { problem, problemDesc, rows, notes, baseFontSize };
            const dataToSave = { ...s, timestamp: Date.now() };
            localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(dataToSave));
        } catch (_) { }
    };

    const suggestedFilename = () => {
        const p = (latestRef.current?.problem ?? problem) || 'export';
        return `RCA_${String(p).replace(/[^a-z0-9]/gi, '_').substring(0, 20) || 'export'}.json`;
    };

    const downloadCopy = () => {
        saveLocalDraft();
        const blob = new Blob([JSON.stringify(buildExportObject(), null, 2)], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = suggestedFilename();
        link.click();
        setTimeout(() => URL.revokeObjectURL(url), 2000);
    };

    const writeToFileHandle = async (handle) => {
        // Direct overwrite, no confirmation.
        await ensureReadWritePermission(handle);
        const data = JSON.stringify(buildExportObject(), null, 2);
        const writable = await handle.createWritable();
        await writable.write(data);
        await writable.close();
    };

    const pickSaveHandle = async () => {
        const opts = {
            suggestedName: suggestedFilename(),
            types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }],
        };
        // Must be called from a user gesture (button click)
        return await window.showSaveFilePicker(opts);
    };

    const pickOpenHandle = async () => {
        const opts = {
            multiple: false,
            types: [{ description: 'JSON', accept: { 'application/json': ['.json'] } }],
        };
        const handles = await window.showOpenFilePicker(opts);
        return handles?.[0] || null;
    };

    const loadFromFileHandle = async (handle) => {
        const file = await handle.getFile();
        const text = await file.text();
        const data = JSON.parse(text);
        if (data.problem) setProblem(data.problem);
        if (data.problemDesc) setProblemDesc(data.problemDesc);
        if (data.rows) setRows(data.rows);
        if (data.notes) setNotes(data.notes);
        if (data.baseFontSize) setBaseFontSize(data.baseFontSize);
    };

    const pulseToast = (text) => {
        setToast({ show: true, text: String(text || '') });
        setTimeout(() => setToast({ show: false, text: '' }), 2500);
    };

    async function handleSave() {
        // Ctrl+S triggers this too.
        // Best-effort: blur active element so any pending UI editing commits.
        try { document.activeElement?.blur?.(); } catch (_) { }
        saveLocalDraft();

        // Direct overwrite, no confirmation.
        if (!nativeSaveAvailable) {
            downloadCopy();
            return;
        }

        // Use ref first (prevents stale state when triggered from a long-lived event handler)
        let handle = saveHandleRef.current || saveHandle;

        try {
            if (!handle) {
                handle = await pickSaveHandle();
                saveHandleRef.current = handle;
                setSaveHandle(handle);
                try { await idbSet('saveHandle', handle); } catch (_) { }
            }

            // Give React a frame to flush any queued updates
            await new Promise(requestAnimationFrame);

            await writeToFileHandle(handle);
            pulseToast('File saved');
        } catch (err) {
            if (err?.name === 'AbortError') return;
            console.error('Save failed:', err);
            downloadCopy();
        }
    }

    // Ensure keyboard shortcut always calls the newest save function
    handleSaveCallRef.current = handleSave;

    async function handleSaveAs() {
        saveLocalDraft();
        // Always choose a new file, then overwrite that chosen file.
        if (!nativeSaveAvailable) {
            downloadCopy();
            return;
        }

        try {
            const handle = await pickSaveHandle();
            saveHandleRef.current = handle;
            setSaveHandle(handle);
            try { await idbSet('saveHandle', handle); } catch (_) { }
            await writeToFileHandle(handle);
            pulseToast('File saved');
        } catch (err) {
            if (err?.name === 'AbortError') return;
            console.error('Save As failed:', err);
            downloadCopy();
        }
    }
    const handleLoad = (e) => {
        const file = e.target.files?.[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            try {
                const data = JSON.parse(ev.target?.result);
                if (data.problem) setProblem(data.problem);
                if (data.problemDesc) setProblemDesc(data.problemDesc);
                if (data.rows) setRows(data.rows);
                if (data.notes) setNotes(data.notes);
                if (data.baseFontSize) setBaseFontSize(data.baseFontSize);
            } catch (err) {
                alert('Invalid file');
            }
        };
        reader.readAsText(file);
        e.target.value = '';
    };

    const unlinkSaveTarget = async () => {
        setSaveHandle(null);
        setSaveTargetName('');
        try { await idbDel('saveHandle'); } catch (_) { }
    };

    const handleOpen = async () => {
        // Prefer native open when available (lets Save overwrite the same file later).
        if (nativeOpenAvailable) {
            try {
                const handle = await pickOpenHandle();
                if (!handle) return;
                await loadFromFileHandle(handle);
                setSaveHandle(handle); // Save will now overwrite this opened file
                try { await idbSet('saveHandle', handle); } catch (_) { }
                pulseToast('File loaded');
                return;
            } catch (err) {
                if (err?.name === 'AbortError') return;
                // Fallback to file input
            }
        }
        fileInputRef.current?.click();
    };
    const handlePrint = () => { setScale(0.85); setPan({ x: 0, y: 0 }); setTimeout(() => window.print(), 100); };

    const fishboneStructure = useMemo(() => {
        const catGroups = {};
        rows.forEach(row => {
            const cat = row.category || "Uncategorized";
            if (!catGroups[cat]) catGroups[cat] = [];
            catGroups[cat].push(row);
        });
        return Object.entries(catGroups).map(([catName, catRows]) => {
            const causeGroups = {};
            catRows.forEach(row => {
                const cause = row.cause || "Unspecified";
                if (!causeGroups[cause]) causeGroups[cause] = [];
                causeGroups[cause].push(row);
            });
            return { name: catName, causes: Object.entries(causeGroups).map(([causeName, factors]) => ({ name: causeName, factors })) };
        });
    }, [rows]);

    const layoutConfig = useMemo(() => {
        const fontGrowth = Math.max(0, baseFontSize - 14);
        const s = 1 + (fontGrowth * 0.01);
        return {
            s,
            fontSize: baseFontSize,
            smallFontSize: Math.max(10, baseFontSize - 2),
            categorySpacing: 350 + (fontGrowth * 6),
            factorSpacing: 25 + (fontGrowth * 0.5),
            minBoneLength: 220 + (fontGrowth * 2),
            verticalBranchGap: 130 + (fontGrowth * 3.5),
            problemFontSize: 24,
            headScale: 1.0
        };
    }, [baseFontSize]);

    const { elements: fishboneElements, totalWidth, totalHeight } = useMemo(() => {
        const { s, fontSize, smallFontSize, categorySpacing, verticalBranchGap, problemFontSize, headScale } = layoutConfig;

        let maxTopHeight = 0;
        let maxBottomHeight = 0;

        const categoriesWithLayout = fishboneStructure.map((cat, index) => {
            const isTop = index % 2 === 0;
            let totalSpacingNeeded = cat.causes.length * verticalBranchGap + (100 * s);
            const boneLength = Math.max(layoutConfig.minBoneLength, totalSpacingNeeded);

            if (isTop) maxTopHeight = Math.max(maxTopHeight, boneLength);
            else maxBottomHeight = Math.max(maxBottomHeight, boneLength);

            let maxCauseLen = 0;

            const causesWithLayout = cat.causes.map(cause => {
                let currentPos = 40 * s;
                const factorsWithPos = cause.factors.map((f) => {
                    const statusConfig = STATUS_OPTIONS[f.status] || STATUS_OPTIONS.ON_GOING;
                    const prefix = statusConfig.icon + " ";
                    const fullText = prefix + f.factor + (f.isPriority ? " ⭐" : "");
                    const width = getTextWidth(fullText, smallFontSize);
                    const step = width * 0.7 + (15 * s);
                    const pos = currentPos;
                    currentPos += step;
                    return { ...f, xOffset: pos, width, statusIcon: statusConfig.icon };
                });
                const totalLen = currentPos + (40 * s);
                const causeTextW = getTextWidth(cause.name, fontSize);

                // Effective length including text
                const finalLen = Math.max(totalLen, causeTextW + (60 * s));

                // Track max length for this category to prevent horizontal overlap
                if (finalLen > maxCauseLen) maxCauseLen = finalLen;

                return { ...cause, factors: factorsWithPos, totalLen: finalLen };
            });

            return { ...cat, causes: causesWithLayout, boneLength, isTop, maxCauseLen };
        });

        const verticalPadding = 150;
        const spineY = maxTopHeight + verticalPadding;
        const dynamicTotalHeight = spineY + maxBottomHeight + verticalPadding;

        // Dynamic Horizontal Layout
        // We calculate positions from Right (Head) to Left (Tail)
        // making sure each category has enough space based on its content.

        const problemTextWidth = getTextWidth(problem, problemFontSize);
        const minHeadW = 280;
        const dynamicHeadW = Math.max(minHeadW, problemTextWidth + 100);

        // Gap required for the FIRST category (closest to head)
        // It needs to clear the Head.
        // Index 0 causes extend RIGHT.
        const firstCatMaxLen = categoriesWithLayout.length > 0 ? categoriesWithLayout[0].maxCauseLen : 0;
        const headGap = Math.max(260 * s, firstCatMaxLen + (100 * s));

        // Calculate absolute X positions (relative to 0 at Head connection)
        const xPositions = [];
        let currentXCursor = 0;

        categoriesWithLayout.forEach((cat, i) => {
            // How much space does this category need to not hit the previous one?
            // If i=0, it hits the head (handled by headGap later).
            // If i>0, it hits Cat[i-1].
            // Cat[i] causes extend RIGHT.

            // We use the configured categorySpacing as a BASE minimum, 
            // but expand if content is wide.
            let gap = categorySpacing;

            // Content-aware expansion:
            // Ensure gap is at least maxCauseLen + padding
            const contentGap = cat.maxCauseLen + (80 * s);
            gap = Math.max(gap, contentGap);

            currentXCursor -= gap;
            xPositions[i] = currentXCursor;
        });

        // Determine where Head starts
        // We want the total width to fit in at least 1200 or more
        // currentXCursor is the left-most position (negative).
        const totalContentWidth = Math.abs(currentXCursor) + headGap + dynamicHeadW + 300;
        const startX = Math.max(200, 100 + Math.abs(currentXCursor) + headGap); // Shift everything right so left-most is visible

        const dynamicHeadX = startX;
        const spineEndX = startX;
        const firstBoneX = startX - headGap;

        // Map relative positions to absolute
        const finalCategoryPositions = xPositions.map(x => firstBoneX + x + categorySpacing /* adjustment for the loop logic diff */);

        // Re-calculate simpler:
        // xPositions[0] is -gap0. We want Cat0 at (HeadX - headGap).
        // Actually, let's just use the cursor logic directly:
        const finalPositions = [];
        let curX = dynamicHeadX - headGap;
        categoriesWithLayout.forEach((cat, i) => {
            finalPositions[i] = curX;

            // Calculate gap for NEXT loop
            if (i < categoriesWithLayout.length - 1) {
                const nextCat = categoriesWithLayout[i + 1];
                // Gap is determined by NEXT category's reach towards THIS one?
                // No, causes extend RIGHT. 
                // So Cat[i+1] extends towards Cat[i].
                // We need Cat[i+1] to be far enough LEFT so its causes don't hit Cat[i]'s bone.

                let gap = categorySpacing;
                const nextContentGap = nextCat.maxCauseLen + (80 * s);
                gap = Math.max(gap, nextContentGap);

                curX -= gap;
            }
        });

        const lastCatX = finalPositions[finalPositions.length - 1] || (dynamicHeadX - 200);
        const spineStartX = lastCatX - 150;

        const headW = dynamicHeadW * headScale;
        const headH = 140 * headScale;
        const requiredWidth = dynamicHeadX + headW + 250;

        const elements = [];

        elements.push(<path key="tail" d={`M ${spineStartX} ${spineY} L ${spineStartX - 60} ${spineY - 50} Q ${spineStartX - 40} ${spineY} ${spineStartX - 60} ${spineY + 50} Z`} fill="#3b82f6" opacity="0.8" />);
        elements.push(<line key="spine" x1={spineStartX} y1={spineY} x2={dynamicHeadX} y2={spineY} stroke="#475569" strokeWidth={6 * s} strokeLinecap="round" />);

        const headPath = `
                    M ${dynamicHeadX} ${spineY - headH / 2}
                    Q ${dynamicHeadX + headW / 2} ${spineY - headH / 2 - 20} ${dynamicHeadX + headW} ${spineY}
                    Q ${dynamicHeadX + headW / 2} ${spineY + headH / 2 + 20} ${dynamicHeadX} ${spineY + headH / 2}
                    Q ${dynamicHeadX - 40} ${spineY} ${dynamicHeadX} ${spineY - headH / 2}
                    Z
                `;

        elements.push(
            <g key="head">
                <path d={headPath} fill="#3b82f6" className="drop-shadow-xl" />
                <circle cx={dynamicHeadX + (headW * 0.7)} cy={spineY - 20} r={10} fill="white" />
                <circle cx={dynamicHeadX + (headW * 0.7)} cy={spineY - 20} r={4} fill="black" />
                <foreignObject x={dynamicHeadX} y={spineY - (headH / 2)} width={headW} height={headH} style={{ pointerEvents: 'none' }}>
                    <div className="flex items-center justify-center h-full px-8 text-center">
                        <span className="text-white font-bold leading-tight drop-shadow-md break-words" style={{ fontSize: `${problemFontSize}px` }}>{problem}</span>
                    </div>
                </foreignObject>
            </g>
        );

        categoriesWithLayout.forEach((cat, index) => {
            const currentX = finalPositions[index];
            const { isTop, boneLength, causes } = cat;
            const endY = isTop ? spineY - boneLength : spineY + boneLength;
            const boneEndX = currentX - (boneLength * 0.45);

            elements.push(
                <g key={`cat-${index}`}>
                    <line x1={currentX} y1={spineY} x2={boneEndX} y2={endY} stroke="#64748b" strokeWidth={4 * s} strokeLinecap="round" />
                    <g transform={`translate(${boneEndX}, ${isTop ? endY - 40 : endY + 20})`}>
                        <rect x={-70} y={-10} width={140} height={34} rx={17} fill="white" stroke="#cbd5e1" strokeWidth="1" className="drop-shadow-sm" />
                        <text x="0" y={13} textAnchor="middle" fontWeight="bold" fill="#1e293b" fontSize={fontSize}>{cat.name}</text>
                    </g>
                </g>
            );

            let distFromSpine = 70 * s;

            causes.forEach((causeGroup, cIndex) => {
                const totalDist = Math.hypot(boneEndX - currentX, endY - spineY);
                const t = distFromSpine / totalDist;
                const rootX = currentX + (boneEndX - currentX) * t;
                const rootY = spineY + (endY - spineY) * t;
                const causeLen = Math.max(160 * s, causeGroup.totalLen);
                const causeEndX = rootX + causeLen;
                const causeEndY = rootY;

                elements.push(
                    <g key={`cause-${index}-${cIndex}`} onClick={() => scrollToRow(causeGroup.factors?.[0]?.id)} style={{ cursor: 'pointer' }}>
                        <line x1={rootX} y1={rootY} x2={causeEndX} y2={causeEndY} stroke="#94a3b8" strokeWidth={3 * s} strokeLinecap="round" />
                        <text x={causeEndX + 8} y={causeEndY + 5} fontSize={fontSize} fontWeight="bold" fill="#334155" textAnchor="start">{causeGroup.name}</text>
                    </g>
                );

                causeGroup.factors.forEach((factor, fIndex) => {
                    const factorRootX = rootX + factor.xOffset;
                    const factorRootY = rootY;
                    const isUpFactor = fIndex % 2 === 0;
                    const branchLen = 30 * s;
                    const branchEndY = factorRootY + (isUpFactor ? -branchLen : branchLen);
                    const branchEndX = factorRootX + (15 * s);
                    const statusConfig = STATUS_OPTIONS[factor.status] || STATUS_OPTIONS.ON_GOING;
                    const isExclude = factor.status === 'EXCLUDE';
                    const isPriority = factor.isPriority === true;
                    const strokeColor = isPriority ? '#ea580c' : statusConfig.boneColor;
                    const strokeWidth = isPriority ? 2.5 : 1.5;
                    const textColor = isPriority ? '#c2410c' : (isExclude ? '#cbd5e1' : '#1e293b');
                    const fontWeight = isPriority ? 'bold' : 'normal';
                    const textY = isUpFactor ? branchEndY - 5 : branchEndY + 14;

                    elements.push(
                        <g key={`fact-${factor.id}`} onClick={() => scrollToRow(factor.id)} style={{ cursor: 'pointer' }}>
                            <circle cx={factorRootX} cy={factorRootY} r={2 * s} fill="#cbd5e1" />
                            <line x1={factorRootX} y1={factorRootY} x2={branchEndX} y2={branchEndY} stroke={strokeColor} strokeWidth={strokeWidth} />
                            <text x={branchEndX - 5} y={textY} fontSize={smallFontSize} fill={textColor} fontWeight={fontWeight} textDecoration={isExclude ? 'line-through' : 'none'} textAnchor="start">
                                {statusConfig.icon} {factor.factor} {isPriority ? "⭐" : ""}
                            </text>
                        </g>
                    );
                });

                distFromSpine += verticalBranchGap;
            });
        });

        return {
            elements,
            totalWidth: Math.max(1200, requiredWidth),
            totalHeight: Math.max(750, dynamicTotalHeight)
        };
    }, [fishboneStructure, layoutConfig, problem, scrollToRow]);

    useEffect(() => {
        const canvasEl = canvasRef.current;
        if (!canvasEl) return;
        const onWheel = (e) => {
            e.preventDefault();
            e.stopPropagation();
            const delta = e.deltaY > 0 ? 0.9 : 1.1;
            setScale(s => Math.min(Math.max(0.3, s * delta), 3));
        };
        canvasEl.addEventListener('wheel', onWheel, { passive: false });
        return () => canvasEl.removeEventListener('wheel', onWheel);
    }, []);

    return (
        <div className="min-h-screen bg-slate-50 flex flex-col text-slate-800 font-sans print:bg-white">
            <header className="bg-white border-b border-slate-200 px-6 py-3 flex justify-between items-center sticky top-0 z-50 print:hidden">
                <div className="flex items-center gap-3">
                    <div className="bg-slate-900 text-white p-1.5 rounded-lg"><Maximize size={20} /></div>
                    <div>
                        <h1 className="font-bold text-lg leading-tight">Fishbone Pro</h1>
                        <p className="text-xs text-slate-500">RCA Edition</p>
                    </div>
                </div>
                <div className="flex gap-4 items-center">
                    {nativeSaveAvailable && (
                        <div className="hidden xl:flex items-center gap-2 px-3 py-1.5 rounded-full bg-slate-50 text-slate-600 border border-slate-200 text-xs font-semibold" title="Debug: Ctrl+S capture + last action">
                            ⌨️ Ctrl+S: <span className="font-mono">{shortcutHits}</span>
                            <span className="text-slate-300">|</span>
                            <span className="font-mono">{lastShortcutAction || '—'}</span>
                        </div>
                    )}
                    {saveTargetName && (
                        <div className="hidden lg:flex items-center gap-2 px-3 py-1.5 rounded-full bg-slate-50 text-slate-700 border border-slate-200 text-xs font-semibold" title="Current Save Target (Save/Ctrl+S will overwrite this file)">
                            <span className="inline-block w-2 h-2 rounded-full bg-emerald-500" />
                            Target: <span className="font-mono">{saveTargetName}</span>
                            <button onClick={unlinkSaveTarget} className="ml-2 px-2 py-0.5 rounded bg-white border border-slate-200 hover:bg-slate-100 text-slate-600" title="Unlink target (next Save will ask where to save)">Unlink</button>
                        </div>
                    )}
                    {!nativeSaveAvailable && (
                        <div className="hidden md:flex items-center gap-2 px-3 py-1.5 rounded-full bg-amber-50 text-amber-800 border border-amber-200 text-xs font-semibold" title={nativeSaveReason}>
                            <span className="inline-block w-2 h-2 rounded-full bg-amber-500" />
                            Native Save disabled in this preview
                        </div>
                    )}
                    <div className={`flex items-center gap-1.5 text-xs font-bold px-3 py-1.5 rounded-full transition-all duration-300 ${toast.show ? 'opacity-100 bg-blue-100 text-blue-700 scale-105' : 'opacity-0 scale-95'}`}>
                        <Check size={14} />
                        {toast.text || '—'}
                    </div>
                    <div className="h-6 w-px bg-slate-200"></div>
                    <input type="file" ref={fileInputRef} onChange={handleLoad} accept=".json" className="hidden" />
                    <button
                        onClick={handleOpen}
                        title={nativeOpenAvailable ? 'Open (native) — enables true overwrite Save' : nativeOpenReason}
                        className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 rounded-md text-sm font-medium hover:bg-slate-50 transition"
                    ><FolderOpen size={16} /> Open</button>
                    <button
                        onClick={handleSave}
                        disabled={!nativeSaveAvailable}
                        title={!nativeSaveAvailable ? nativeSaveReason : 'Save (Ctrl+S) — overwrite the same file'}
                        className={`flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 rounded-md text-sm font-medium transition ${nativeSaveAvailable ? 'hover:bg-slate-50' : 'opacity-50 cursor-not-allowed'}`}
                    >
                        <Save size={16} /> Save (Ctrl+S)
                    </button>
                    <button
                        onClick={handleSaveAs}
                        disabled={!nativeSaveAvailable}
                        title={!nativeSaveAvailable ? nativeSaveReason : 'Save As (choose a new file)'}
                        className={`flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 rounded-md text-sm font-medium transition ${nativeSaveAvailable ? 'hover:bg-slate-50' : 'opacity-50 cursor-not-allowed'}`}
                    >
                        <Save size={16} /> Save As
                    </button>
                    <button onClick={downloadCopy} className="flex items-center gap-2 px-3 py-1.5 bg-white border border-slate-300 rounded-md text-sm font-medium hover:bg-slate-50 transition"><Save size={16} /> Download Copy</button>
                    <button onClick={handlePrint} className="flex items-center gap-2 px-4 py-1.5 bg-blue-600 text-white rounded-md text-sm font-medium hover:bg-blue-700 shadow-sm transition"><Printer size={16} /> Print / PDF</button>
                </div>
            </header>

            <div className="flex-1 flex flex-col p-6 gap-6 max-w-[98%] mx-auto w-full print:p-0">
                <div className="bg-white rounded-xl border border-slate-200 shadow-sm p-5 flex flex-col justify-center gap-3 print:hidden">
                    <div className="w-full">
                        <label className="text-xs font-bold text-slate-400 uppercase mb-1 block">Problem Title (Main Effect)</label>
                        <input type="text" value={problem} onChange={(e) => setProblem(e.target.value)} className="w-full bg-transparent font-bold text-slate-800 text-2xl outline-none border-b border-transparent focus:border-blue-500 transition-colors placeholder-slate-300" placeholder="Enter the main problem here..." />
                    </div>
                    <div className="w-full">
                        <label className="text-xs font-bold text-slate-400 uppercase mb-1 block">Description / Context</label>
                        <textarea value={problemDesc} onChange={(e) => setProblemDesc(e.target.value)} className="w-full bg-slate-50 rounded-lg p-2 text-sm text-slate-600 outline-none border border-slate-200 focus:border-blue-400 transition-colors resize-none" rows={2} placeholder="Detailed description of the problem..." />
                    </div>
                </div>

                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden relative print:border-none print:shadow-none" style={{ height: '550px' }}>
                    <div className="absolute top-4 left-4 z-10 bg-white/90 p-2 rounded-lg border border-slate-200 shadow-sm backdrop-blur flex items-center gap-3 print:hidden">
                        <Type size={16} className="text-slate-500" />
                        <div className="flex flex-col">
                            <span className="text-[10px] font-bold text-slate-400 uppercase leading-none mb-1">Font Size</span>
                            <div className="flex items-center gap-2">
                                <input type="range" min="12" max="60" step="1" value={baseFontSize} onChange={(e) => setBaseFontSize(Number(e.target.value))} className="w-24 accent-blue-600 h-1.5" />
                                <span className="text-xs font-mono w-4 text-slate-600">{baseFontSize}</span>
                            </div>
                        </div>
                    </div>
                    <div className="absolute bottom-4 right-4 z-10 flex gap-2 print:hidden">
                        <button onClick={() => { setScale(1); setPan({ x: 0, y: 0 }) }} className="p-2 bg-white border border-slate-200 rounded-lg hover:bg-slate-50 shadow-sm text-xs font-bold text-slate-600">Reset View</button>
                    </div>

                    <div
                        ref={canvasRef}
                        className={`w-full h-full bg-slate-50 cursor-grab active:cursor-grabbing print:bg-white ${isDragging ? 'cursor-grabbing' : ''}`}
                        onMouseDown={(e) => { if (e.button === 0) { setIsDragging(true); setDragStart({ x: e.clientX - pan.x, y: e.clientY - pan.y }) } }}
                        onMouseMove={(e) => { if (isDragging) setPan({ x: e.clientX - dragStart.x, y: e.clientY - dragStart.y }) }}
                        onMouseUp={() => setIsDragging(false)}
                        onMouseLeave={() => setIsDragging(false)}
                    >
                        <svg width="100%" height="100%" viewBox={`${-pan.x} ${-pan.y} ${totalWidth} ${totalHeight}`} style={{ transform: `scale(${scale})`, transformOrigin: '0 0', transition: isDragging ? 'none' : 'transform 0.1s ease-out' }}>
                            <defs>
                                <pattern id="grid" width="40" height="40" patternUnits="userSpaceOnUse">
                                    <path d="M 40 0 L 0 0 0 40" fill="none" stroke="#e2e8f0" strokeWidth="1" />
                                </pattern>
                            </defs>
                            <rect x={-5000} y={-5000} width="10000" height="10000" fill="url(#grid)" className="print:hidden" />
                            {fishboneElements}
                        </svg>
                    </div>
                </div>

                <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 print:border-none print:shadow-none print:p-0">
                    <div className="flex justify-between items-center mb-4 print:hidden">
                        <h2 className="text-lg font-bold flex items-center gap-2"><FileText size={20} className="text-blue-600" />Hypothesis Verification</h2>
                        <div className="flex gap-2">
                            {rows.length === 0 && (
                                <button onClick={addFirstRow} className="px-3 py-1.5 bg-blue-50 text-blue-600 rounded-md text-sm font-medium hover:bg-blue-100 flex items-center gap-1"><Plus size={16} /> Add First Row</button>
                            )}
                            <button onClick={clearData} className="px-3 py-1.5 text-slate-400 hover:text-red-500 hover:bg-red-50 rounded-md text-sm transition flex items-center gap-1"><Trash2 size={16} /> Clear / Reset</button>
                        </div>
                    </div>

                    <div className="overflow-x-auto">
                        <table className="w-full text-left text-sm border-collapse min-w-[1200px]">
                            <thead>
                                <tr className="border-b border-slate-200 text-slate-500 text-xs uppercase tracking-wider">
                                    <th className="py-3 px-2 w-[30px] text-center"></th>
                                    <th className="py-3 px-2 w-[4%] text-center">Priority</th>
                                    <th className="py-3 px-2 w-[10%]">Primary Bone (Category)</th>
                                    <th className="py-3 px-2 w-[12%]">Secondary Bone (Cause)</th>
                                    <th className="py-3 px-2 w-[15%] text-blue-700">Specific Factor (Hypothesis)</th>
                                    <th className="py-3 px-2 w-[22%]">Verification Method</th>
                                    <th className="py-3 px-2 w-[22%]">Evidence / Result</th>
                                    <th className="py-3 px-2 w-[10%]">Status</th>
                                    <th className="py-3 px-2 w-[5%] text-right print:hidden"></th>
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-slate-100">
                                {rows.map((row, index) => (
                                    <tr
                                        id={`verify-row-${row.id}`}
                                        key={row.id}
                                        className={`group hover:bg-slate-50 transition-colors ${highlightRowId === row.id ? 'bg-yellow-100 ring-2 ring-yellow-300' : ''}`}
                                        draggable
                                        onDragStart={(e) => handleDragStart(e, index)}
                                        onDragEnd={handleDragEnd}
                                        onDragOver={(e) => handleDragOver(e, index)}
                                        onDrop={(e) => handleDrop(e, index)}
                                    >
                                        <td
                                            className="p-2 align-middle text-center cursor-grab active:cursor-grabbing text-slate-400 hover:text-slate-600 drag-handle touch-none"
                                            onMouseDown={() => { allowDragRef.current = true; }}
                                            onMouseUp={() => { allowDragRef.current = false; }}
                                        >
                                            <GripVertical size={16} />
                                        </td>
                                        <td className="p-2 align-top text-center">
                                            <button onClick={() => handleRowChange(row.id, 'isPriority', !row.isPriority)} className="p-1.5 hover:bg-orange-50 rounded-full transition">
                                                <Star size={18} filled={row.isPriority} />
                                            </button>
                                        </td>
                                        <td className="p-2 align-top">
                                            <input type="text" value={row.category} onChange={e => handleRowChange(row.id, 'category', e.target.value)} className="w-full bg-transparent font-bold text-slate-700 outline-none border-b border-transparent focus:border-blue-400 px-1 py-1" placeholder="Category" />
                                        </td>
                                        <td className="p-2 align-top">
                                            <input type="text" value={row.cause} onChange={e => handleRowChange(row.id, 'cause', e.target.value)} className="w-full bg-transparent text-slate-700 outline-none border-b border-transparent focus:border-blue-400 px-1 py-1" placeholder="Cause" />
                                        </td>
                                        <td className="p-2 align-top bg-blue-50/30 rounded">
                                            <input type="text" value={row.factor} onChange={e => handleRowChange(row.id, 'factor', e.target.value)} className="w-full bg-transparent font-medium text-blue-900 outline-none border-b border-transparent focus:border-blue-400 px-1 py-1" placeholder="Hypothesis" />
                                        </td>
                                        <td className="p-2 align-top"><HoverTextarea value={row.method} onChange={e => handleRowChange(row.id, 'method', e.target.value)} placeholder="Method" /></td>
                                        <td className="p-2 align-top"><HoverTextarea value={row.evidence} onChange={e => handleRowChange(row.id, 'evidence', e.target.value)} placeholder="Result" /></td>
                                        <td className="p-2 align-top">
                                            <select value={row.status} onChange={e => handleRowChange(row.id, 'status', e.target.value)} className={`w-full text-xs font-bold rounded-md px-2 py-1.5 border-none outline-none cursor-pointer appearance-none ${STATUS_OPTIONS[row.status]?.bg || 'bg-gray-100'}`}>
                                                {Object.values(STATUS_OPTIONS).map(opt => <option key={opt.id} value={opt.id}>{opt.label}</option>)}
                                            </select>
                                        </td>
                                        <td className="p-2 text-right print:hidden">
                                            <div className="flex justify-end gap-1">
                                                <button onClick={() => insertRowAfter(index)} className="p-1.5 text-blue-600 hover:bg-blue-50 rounded" title="Insert Below"><Plus size={16} /></button>
                                                <button onClick={() => deleteRow(row.id)} className="p-1.5 text-red-500 hover:bg-red-50 rounded" title="Delete"><Trash2 size={16} /></button>
                                            </div>
                                        </td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>

                    <RichTextEditor value={notes} onChange={setNotes} />
                </div>
            </div>
        </div>
    );
};

// -----------------------------
// Smoke tests (quick sanity checks)
// -----------------------------
function runSmokeTests() {
    console.assert(!!React && !!ReactDOM, 'React/ReactDOM should be loaded');
    console.assert(typeof STATUS_OPTIONS === 'object', 'STATUS_OPTIONS should exist');
    console.assert(Array.isArray(DEMO_DATA), 'DEMO_DATA should be an array');
    console.assert(typeof getTextWidth === 'function', 'getTextWidth should be a function');
    console.assert(typeof document.execCommand === 'function' || typeof document.execCommand === 'undefined', 'document.execCommand may be unavailable (OK)');
    console.assert(typeof idbGet === 'function' && typeof idbSet === 'function', 'IDB helpers should exist');
    console.assert(typeof canUseNativeFilePicker === 'function', 'canUseNativeFilePicker helper should exist');
    console.assert(typeof canUseNativeFilePicker() === 'boolean', 'canUseNativeFilePicker should return a boolean');
    console.assert(typeof canUseNativeOpenPicker === 'function', 'canUseNativeOpenPicker helper should exist');
    console.assert(typeof canUseNativeOpenPicker() === 'boolean', 'canUseNativeOpenPicker should return a boolean');
    console.assert(!!window.Quill, 'Quill should be loaded');
    // Additional sanity: avoid undefined localStorage key
    console.assert(typeof LOCAL_STORAGE_KEY === 'string' && LOCAL_STORAGE_KEY.length > 0, 'LOCAL_STORAGE_KEY should be a non-empty string');
}

runSmokeTests();

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(<App />);
