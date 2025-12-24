
// --- HoverExpandTextarea ---
const { useState, useMemo, useRef, useEffect, useCallback } = React;

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
const RichTextEditor = ({ value, onChange, className }) => {
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
        <div className={`w-full bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden print:border-none print:shadow-none ${className || 'mt-6'}`}>
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

window.HoverTextarea = HoverTextarea;
window.RichTextEditor = RichTextEditor;
