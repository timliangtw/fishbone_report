// --- Setup Globals ---
const { useState, useMemo, useRef, useEffect, useCallback } = React;
const {
    Plus, Trash2, Save, FolderOpen, Printer, Maximize, FileText, Type, Star, Check, GripVertical
} = window.Icons;
const {
    STATUS_OPTIONS, DEMO_DATA, getTextWidth, idbGet, idbSet, idbDel,
    canUseNativeFilePicker, canUseNativeOpenPicker, ensureReadWritePermission,
    isEmbeddedFrame, LOCAL_STORAGE_KEY
} = window;
const { RichTextEditor, HoverTextarea } = window;


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

    // Report Metadata State
    const [projectName, setProjectName] = useState("Root Cause Analysis");
    const [owner, setOwner] = useState("");
    const [reportStage, setReportStage] = useState("Draft");


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
        latestRef.current = { problem, problemDesc, rows, notes, baseFontSize, projectName, owner, reportStage };
    }, [problem, problemDesc, rows, notes, baseFontSize, projectName, owner, reportStage]);

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
                // Load metadata with fallbacks
                if (parsed.projectName) setProjectName(parsed.projectName);
                if (parsed.owner) setOwner(parsed.owner);
                if (parsed.reportStage) setReportStage(parsed.reportStage);
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
            projectName: s.projectName,
            owner: s.owner,
            reportStage: s.reportStage,
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
        if (data.projectName) setProjectName(data.projectName);
        if (data.owner) setOwner(data.owner);
        if (data.reportStage) setReportStage(data.reportStage);
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
                // Load metadata with fallbacks
                if (data.projectName) setProjectName(data.projectName);
                if (data.owner) setOwner(data.owner);
                if (data.reportStage) setReportStage(data.reportStage);

                pulseToast('Loaded successfully');
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
        const headerHeight = 100 * s;
        const footerHeight = 60 * s;
        const spineY = maxTopHeight + verticalPadding + headerHeight;
        const dynamicTotalHeight = spineY + maxBottomHeight + verticalPadding + footerHeight;

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

        // --- Report Header (Project Name) ---
        elements.push(
            <text key="header-title" x="50" y={80 * s} fontSize={36 * s} fontWeight="bold" fill="#1e293b" fontFamily="sans-serif">
                {projectName}
            </text>
        );
        elements.push(
            <line key="header-line" x1="50" y1={100 * s} x2={requiredWidth - 50} y2={100 * s} stroke="#e2e8f0" strokeWidth={2 * s} />
        );

        // --- Fishbone Body ---

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

        // --- Report Footer ---
        const footerY = dynamicTotalHeight - (20 * s);
        const dateStr = new Date().toLocaleDateString('zh-TW');

        elements.push(
            <g key="footer">
                <line x1="50" y1={footerY - (30 * s)} x2={requiredWidth - 50} y2={footerY - (30 * s)} stroke="#e2e8f0" strokeWidth={1 * s} />
                <text x="50" y={footerY} fontSize={14 * s} fill="#64748b">
                    Generated: {dateStr}
                </text>
                <text x={requiredWidth / 2} y={footerY} textAnchor="middle" fontSize={14 * s} fill="#64748b">
                    {owner ? `Owner: ${owner}` : ''}
                </text>
                <text x={requiredWidth - 50} y={footerY} textAnchor="end" fontSize={14 * s} fontWeight="bold" fill={reportStage === 'Final' ? '#059669' : '#d97706'}>
                    {reportStage} Ver.
                </text>
            </g>
        );

        return {
            elements,
            totalWidth: Math.max(1200, requiredWidth),
            totalHeight: Math.max(750, dynamicTotalHeight)
        };
    }, [fishboneStructure, layoutConfig, problem, scrollToRow, projectName, owner, reportStage]);

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
                    <div className="bg-white p-4 rounded-xl shadow-sm border border-slate-200">
                        <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-2">Report Metadata</label>
                        <div className="space-y-3">
                            <div>
                                <label className="block text-xs text-slate-400 mb-1">Project Name</label>
                                <input
                                    type="text"
                                    value={projectName}
                                    onChange={(e) => setProjectName(e.target.value)}
                                    className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 font-bold text-slate-700"
                                    placeholder="Enter Project Name..."
                                />
                            </div>
                            <div className="grid grid-cols-2 gap-3">
                                <div>
                                    <label className="block text-xs text-slate-400 mb-1">Owner</label>
                                    <input
                                        type="text"
                                        value={owner}
                                        onChange={(e) => setOwner(e.target.value)}
                                        className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                                        placeholder="Name..."
                                    />
                                </div>
                                <div>
                                    <label className="block text-xs text-slate-400 mb-1">Status</label>
                                    <select
                                        value={reportStage}
                                        onChange={(e) => setReportStage(e.target.value)}
                                        className="w-full px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 bg-white"
                                    >
                                        <option value="Draft">Draft</option>
                                        <option value="Final">Final</option>
                                    </select>
                                </div>
                            </div>
                        </div>
                    </div>

                    <div className="w-full">
                        <label className="text-xs font-bold text-slate-400 uppercase mb-1 block">Problem Title (Main Effect)</label>
                        <input type="text" value={problem} onChange={(e) => setProblem(e.target.value)} className="w-full bg-transparent font-bold text-slate-800 text-2xl outline-none border-b border-transparent focus:border-blue-500 transition-colors placeholder-slate-300" placeholder="Enter the main problem here..." />
                    </div>
                    <div className="w-full">
                        <label className="text-xs font-bold text-slate-400 uppercase mb-1 block">Description / Context</label>
                        <RichTextEditor value={problemDesc} onChange={setProblemDesc} className="mt-1" />
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

                    <RichTextEditor value={notes} onChange={setNotes} className="mt-6" />
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
