var container;

// ── safeResize ────────────────────────────────────────────────────────────────
function safeResize(attemptsLeft) {
    if (!container) return;
    var parent = document.querySelector('.n-editor-area') || document.getElementById('container1');
    var w = parent ? parent.offsetWidth : 0;
    if (w > 0) { container.resize(); return; }
    if (attemptsLeft <= 0) return;
    requestAnimationFrame(function () { setTimeout(function () { safeResize(attemptsLeft - 1); }, 50); });
}

// ══════════════════════════════════════════════════════════════════════════════
// INDEXEDDB — Tab SFDT Storage (keeps large SFDT out of Blazor/.NET memory)
// ══════════════════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════════════════
// INDEXEDDB — Tab SFDT Storage (with full diagnostics)
// ══════════════════════════════════════════════════════════════════════════════
var _tabDb = null;
var _tabDbReady = new Promise(function (resolve) {
    var req = indexedDB.open('deditor-tabs', 1);
    req.onupgradeneeded = function (e) {
        console.log('[TabDB] Creating object store...');
        e.target.result.createObjectStore('tabs');
    };
    req.onsuccess = function (e) {
        _tabDb = e.target.result;
        console.log('[TabDB] IndexedDB ready');
        resolve();
    };
    req.onerror = function (e) {
        console.error('[TabDB] IndexedDB init FAILED:', e);
        resolve();
    };
});

// Save SFDT to IndexedDB
window.tabDbSave = async function (tabId, sfdt) {
    await _tabDbReady;
    if (!_tabDb) { console.error('[TabDB] save: DB not ready!'); return false; }
    if (!tabId) { console.error('[TabDB] save: no tabId!'); return false; }
    if (!sfdt) { console.error('[TabDB] save: no sfdt data!'); return false; }
    console.log('[TabDB] SAVE tabId=' + tabId.substring(0, 8) + '... size=' + Math.round(sfdt.length / 1024) + 'KB');
    return new Promise(function (resolve) {
        try {
            var tx = _tabDb.transaction('tabs', 'readwrite');
            tx.objectStore('tabs').put(sfdt, tabId);
            tx.oncomplete = function () {
                console.log('[TabDB] SAVE OK: ' + tabId.substring(0, 8));
                resolve(true);
            };
            tx.onerror = function (e) {
                console.error('[TabDB] SAVE FAILED:', e);
                resolve(false);
            };
        } catch (e) { console.error('[TabDB] save exception:', e); resolve(false); }
    });
};

// Load SFDT from IndexedDB
window.tabDbLoad = async function (tabId) {
    await _tabDbReady;
    if (!_tabDb) { console.error('[TabDB] load: DB not ready!'); return null; }
    if (!tabId) { console.error('[TabDB] load: no tabId!'); return null; }
    console.log('[TabDB] LOAD tabId=' + tabId.substring(0, 8) + '...');
    return new Promise(function (resolve) {
        try {
            var tx = _tabDb.transaction('tabs', 'readonly');
            var req = tx.objectStore('tabs').get(tabId);
            req.onsuccess = function () {
                var data = req.result || null;
                if (data) {
                    console.log('[TabDB] LOAD OK: ' + tabId.substring(0, 8) + ' size=' + Math.round(data.length / 1024) + 'KB');
                } else {
                    console.warn('[TabDB] LOAD EMPTY: ' + tabId.substring(0, 8) + ' — no data found!');
                }
                resolve(data);
            };
            req.onerror = function (e) {
                console.error('[TabDB] LOAD FAILED:', e);
                resolve(null);
            };
        } catch (e) { console.error('[TabDB] load exception:', e); resolve(null); }
    });
};

// Delete SFDT from IndexedDB
window.tabDbDelete = async function (tabId) {
    await _tabDbReady;
    if (!_tabDb) return;
    console.log('[TabDB] DELETE tabId=' + (tabId || '').substring(0, 8));
    try {
        var tx = _tabDb.transaction('tabs', 'readwrite');
        tx.objectStore('tabs').delete(tabId);
    } catch (e) { console.error('[TabDB] delete error:', e); }
};

// Check if tab has data in IndexedDB
window.tabDbHas = async function (tabId) {
    var data = await window.tabDbLoad(tabId);
    return data != null && data.length > 0;
};

// Save current editor content to IndexedDB (serialize + store)
window.saveCurrentEditorToDb = async function (tabId) {
    if (!container || !tabId) {
        console.warn('[TabDB] saveCurrentEditor: container=' + !!container + ' tabId=' + tabId);
        return false;
    }
    console.log('[TabDB] Serializing editor for tab ' + tabId.substring(0, 8) + '...');
    var sfdt = container.documentEditor.serialize();
    if (!sfdt || sfdt.length < 10) {
        console.warn('[TabDB] saveCurrentEditor: serialize returned empty/tiny (' + (sfdt ? sfdt.length : 0) + ' chars)');
        return false;
    }
    console.log('[TabDB] Serialized: ' + Math.round(sfdt.length / 1024) + 'KB — saving to DB...');
    return await window.tabDbSave(tabId, sfdt);
};

// Load tab from IndexedDB into editor
window.loadTabFromDb = async function (tabId) {
    console.log('[TabDB] loadTabFromDb: ' + tabId.substring(0, 8));
    showLoader();
    var sfdt = await window.tabDbLoad(tabId);
    if (sfdt && sfdt.length > 10) {
        console.log('[TabDB] Opening in editor: ' + Math.round(sfdt.length / 1024) + 'KB');
        container.documentEditor.open(sfdt);
        container.documentEditor.focusIn();
        hideLoader();
        return true;
    }
    console.warn('[TabDB] loadTabFromDb: NO DATA — opening blank');
    hideLoader();
    return false;
};

// Full tab switch: save current → load next
window.switchTab = async function (currentTabId, nextTabId, nextIsEmpty) {
    console.log('[TabDB] === SWITCH TAB ===');
    showLoader();

    // 1. Save current tab to IndexedDB
    if (currentTabId && currentTabId !== '') {
        await window.saveCurrentEditorToDb(currentTabId);
    }

    // 2. Try IndexedDB first — always (user may have typed in a "new" tab)
    var sfdt = await window.tabDbLoad(nextTabId);
    if (sfdt && sfdt.length > 0) {
        console.log('[TabDB]   Loading from DB: ' + Math.round(sfdt.length / 1024) + 'KB');
        // open() internally destroys the previous document — no openBlank() needed
        container.documentEditor.open(sfdt);
        sfdt = null;
        container.documentEditor.focusIn();
    } else {
        console.log('[TabDB]   No saved content — blank');
        container.documentEditor.openBlank();
        container.documentEditor.focusIn();
    }

    setTimeout(function () {
        try { container.resize(); } catch (e) {}
        hideLoader();
    }, 100);
    console.log('[TabDB] === SWITCH COMPLETE ===');
};

// Save SFDT string directly to IndexedDB (used after server import)
window.storeSfdtToDb = async function (tabId, sfdt) {
    console.log('[TabDB] storeSfdtToDb: tabId=' + (tabId || '').substring(0, 8) + ' size=' + (sfdt ? Math.round(sfdt.length / 1024) + 'KB' : 'null'));
    await window.tabDbSave(tabId, sfdt);
};


// ── Initialize Syncfusion Document Editor ─────────────────────────────────────
window.initializeDocumentEditor = function () {
    container = new ej.documenteditor.DocumentEditorContainer({
        height: "calc(100vh - 46px)",
        width: "100%",
        enableToolbar: true,
        toolbarMode: 'Ribbon',
        serviceUrl: '/api/documenteditor/',
    });


container.appendTo("#container1");
    container.documentEditor.enableTrackChanges = false;
    container.documentEditor.showRevisions = false;
    container.documentEditor.enableOptimizedTextMeasuring = true;
    window._deContainer = container;

    safeResize(20);

    if (typeof ResizeObserver !== 'undefined') {
        var _editorArea = document.querySelector('.n-editor-area');
        if (_editorArea) {
            var _ro = new ResizeObserver(function (entries) {
                for (var i = 0; i < entries.length; i++) {
                    if (entries[i].contentRect.width > 0 && container) safeResize(3);
                }
            });
            _ro.observe(_editorArea);
            window._deResizeObserver = _ro;
        }
    }

    var _winResizeTimer = null;
    window.addEventListener('resize', function () {
        clearTimeout(_winResizeTimer);
        _winResizeTimer = setTimeout(function () { safeResize(5); }, 80);
    });
    _interceptFileMenu();
    _disableSyncfusionSave();
};

// ── Intercept Syncfusion ribbon File > New / Open ────────────────────────────
function _interceptFileMenu() {
    document.addEventListener('click', function (e) {
        var menuItem = e.target.closest('.e-menu-item');
        if (!menuItem || !window._blazorEditorRef) return;
        var text = (menuItem.textContent || '').trim();
        if (text !== 'New' && text !== 'Open') return;

        e.preventDefault();
        e.stopImmediatePropagation();

        try {
            var fileBtn = document.querySelector('.e-ribbon-file-menu');
            if (fileBtn) fileBtn.click();
        } catch (ex) {}

        if (text === 'New') {
            window._blazorEditorRef.invokeMethodAsync('NewTabFromJS');
        } else {
            window._blazorEditorRef.invokeMethodAsync('OpenFileDirect');
        }
    }, true);
}

// ── Disable Syncfusion built-in Ctrl+S ───────────────────────────────────────
function _disableSyncfusionSave() {
    if (!window._deContainer) return;
    var de = window._deContainer.documentEditor;
    de.keyDown = function (args) {
        if (!args || !args.event) return;
        var e = args.event;
        if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
            args.isHandled = true;
        }
    };
}

function showLoader() {
    // Try to overlay just the viewer container (canvas area)
    var viewer = document.getElementById("container1_editor_viewerContainer");
    if (viewer) {
        var overlay = document.getElementById("editorLoader");
        if (!overlay) {
            overlay = document.createElement("div");
            overlay.id = "editorLoader";
            overlay.className = "editor-loader";
            overlay.innerHTML = '<div class="global-spinner"></div>';
            viewer.appendChild(overlay);
        }
        overlay.style.display = "flex";
        return;
    }
    // Fallback to global loader if editor not mounted yet
    document.getElementById("globalLoader").style.display = "flex";
}

function hideLoader() {
    var el = document.getElementById("editorLoader");
    if (el) el.style.display = "none";
    document.getElementById("globalLoader").style.display = "none";
}

// ══════════════════════════════════════════════════════════════════════════════
// FILE SYSTEM ACCESS API — Notepad-style Save / Save As
// ══════════════════════════════════════════════════════════════════════════════
// Stores FileSystemFileHandle per tab so "Save" can overwrite without a picker.
// Falls back to download-based save if the API is not supported.
// ══════════════════════════════════════════════════════════════════════════════

var _fileHandles = {};  // tabId → FileSystemFileHandle

// Check if File System Access API is available
window.hasFileSystemAccess = function () {
    return typeof window.showSaveFilePicker === 'function';
};

// Check if a tab already has a saved file handle
window.hasFileHandle = function (tabId) {
    return !!_fileHandles[tabId];
};

// Clear file handle for a tab (used when closing tab or creating new)
window.clearFileHandle = function (tabId) {
    delete _fileHandles[tabId];
};

// ── Check if saved file still exists on disk ─────────────────────────────────
// Returns: "EXISTS", "DELETED", "NO_HANDLE", or "ERROR:msg"
window.checkFileExists = async function (tabId) {
    if (!_fileHandles[tabId]) return 'NO_HANDLE';
    try {
        var handle = _fileHandles[tabId];
        // queryPermission first — if denied, treat as no handle
        var perm = await handle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted') {
            perm = await handle.requestPermission({ mode: 'readwrite' });
            if (perm !== 'granted') return 'NO_HANDLE';
        }
        // getFile() throws NotFoundError if the file was deleted from disk
        await handle.getFile();
        return 'EXISTS';
    } catch (e) {
        if (e.name === 'NotFoundError') return 'DELETED';
        if (e.name === 'NotAllowedError' || e.name === 'SecurityError') {
            delete _fileHandles[tabId];
            return 'NO_HANDLE';
        }
        return 'ERROR:' + e.message;
    }
};

// ── Get the full file name stored in the handle ──────────────────────────────
window.getFileHandleName = function (tabId) {
    if (!_fileHandles[tabId]) return '';
    return _fileHandles[tabId].name || '';
};

// ── Save to existing file handle (silent, no picker) ─────────────────────────
// Returns: "OK" on success, "NO_HANDLE" if no handle, "DELETED" if file gone, "ERROR:msg" on failure
window.saveToExistingHandle = async function (tabId, base64) {
    if (!_fileHandles[tabId]) return 'NO_HANDLE';
    try {
        var handle = _fileHandles[tabId];
        var bytes = Uint8Array.from(atob(base64), function (c) { return c.charCodeAt(0); });

        // Verify we still have write permission
        var perm = await handle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted') {
            perm = await handle.requestPermission({ mode: 'readwrite' });
            if (perm !== 'granted') return 'NO_HANDLE';
        }

        // Check file still exists before writing
        try {
            await handle.getFile();
        } catch (fileErr) {
            if (fileErr.name === 'NotFoundError') return 'DELETED';
            throw fileErr;
        }

        var writable = await handle.createWritable();
        await writable.write(bytes);
        await writable.close();
        return 'OK';
    } catch (e) {
        console.warn('[Save] write to handle failed:', e);
        if (e.name === 'NotAllowedError' || e.name === 'SecurityError') {
            delete _fileHandles[tabId];
            return 'NO_HANDLE';
        }
        if (e.name === 'NotFoundError') return 'DELETED';
        return 'ERROR:' + e.message;
    }
};

// ── Save As with native file picker ──────────────────────────────────────────
// Returns: chosen filename on success, "CANCELLED" if user cancelled, "ERROR:msg" on failure
// ── Save As — Complete flow (picker opens FIRST to preserve user gesture) ────
// Called from Blazor. Opens picker immediately, then serializes + saves.
window.saveAsComplete = async function (tabId, suggestedName, formatExt, serverSaveUrl) {
    try {
        var mimeMap = {
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.doc':  'application/msword',
            '.rtf':  'application/rtf',
            '.txt':  'text/plain',
            '.xml':  'application/xml'
        };
        var descMap = {
            '.docx': 'Word Document',
            '.doc':  'Word 97-2003 Document',
            '.rtf':  'Rich Text Format',
            '.txt':  'Plain Text',
            '.xml':  'Word XML Document'
        };

        var ext  = formatExt || '.docx';
        var mime = mimeMap[ext] || mimeMap['.docx'];
        var desc = descMap[ext] || 'Document';

        var types = [{ description: desc, accept: {} }];
        types[0].accept[mime] = [ext];
        Object.keys(mimeMap).forEach(function (e) {
            if (e !== ext) {
                var t = { description: descMap[e], accept: {} };
                t.accept[mimeMap[e]] = [e];
                types.push(t);
            }
        });

        // 1) Open picker IMMEDIATELY (user gesture is still valid)
        var handle = await window.showSaveFilePicker({
            suggestedName: suggestedName || 'Untitled.docx',
            types: types
        });

        // 2) Now serialize the document
        var sfdt = container ? container.documentEditor.serialize() : null;
        if (!sfdt) return 'ERROR:No document content';

        // Determine final filename and extension from what user picked
        var chosenName = handle.name || suggestedName;
        var chosenExt  = chosenName.lastIndexOf('.') >= 0
            ? chosenName.substring(chosenName.lastIndexOf('.'))
            : ext;

        // 3) POST to server to convert SFDT → binary
        var resp = await fetch(serverSaveUrl || '/api/documenteditor/save', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ Content: sfdt, FileName: chosenName })
        });

        if (!resp.ok) return 'ERROR:Server save failed (' + resp.status + ')';

        // 4) Write bytes to the file handle
        var blob    = await resp.blob();
        var writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();

        // 5) Store handle for future Save calls
        _fileHandles[tabId] = handle;

        return 'OK:' + chosenName;
    } catch (e) {
        if (e.name === 'AbortError') return 'CANCELLED';
        console.error('[SaveAs] failed:', e);
        return 'ERROR:' + e.message;
    }
};

// ── Open file with native picker (returns handle for write-back) ──────────────
// Returns: { fileName, base64 } on success, { error: "CANCELLED" } or { error: "msg" }
window.openAndUploadToServer = async function () {
    try {
        console.log("🚀 Open started");
        const startTime = performance.now();

        showLoader(); // ✅ START LOADING UI

       var handles = await window.showOpenFilePicker({
            multiple: false,
            types: [{
                description: 'Documents',
                accept: {
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
                    'application/msword': ['.doc'],
                    'application/rtf': ['.rtf'],
                    'text/plain': ['.txt'],
                    'application/json': ['.sfdt']
                }
            }]
        });

        if (!handles || handles.length === 0) {
            hideLoader();
            return;
        }

        console.log("📂 File selected");

      var file = await handles[0].getFile();
        var fileName = file.name || '';
        var fileExt = fileName.lastIndexOf('.') >= 0 ? fileName.substring(fileName.lastIndexOf('.')).toLowerCase() : '';

        // Validate file format
        var supportedFormats = ['.docx', '.doc', '.rtf', '.txt', '.sfdt', '.xml'];
        if (!supportedFormats.includes(fileExt)) {
            hideLoader();
            alert('Unsupported file format: "' + fileExt + '"\n\nSupported formats:\n• .docx (Word Document)\n• .doc (Word 97-2003)\n• .rtf (Rich Text Format)\n• .txt (Plain Text)\n• .sfdt (Syncfusion Document)');
            return;
        }

        var sfdt;

        // SFDT files are already in Syncfusion's native format — load directly
        if (fileExt === '.sfdt') {
            console.log("⚡ SFDT file — loading directly (no server round-trip)");
            sfdt = await file.text();
        } else {
            var formData = new FormData();
            formData.append("files", file);

            console.log("⬆ Uploading to server...");
            const uploadStart = performance.now();

            const response = await fetch('/api/documenteditor/import', {
                method: 'POST',
                body: formData
            });

            const uploadEnd = performance.now();
            console.log(`✅ Server response received in ${(uploadEnd - uploadStart).toFixed(2)} ms`);

if (!response.ok) {
                var errText = await response.text();
                console.error("❌ Server error:", errText);
                hideLoader();
                // Show user-friendly message for large files
                if (errText.indexOf('too large') >= 0 || errText.indexOf('File too large') >= 0) {
                    alert('This file is too large to open.\n\nMaximum supported file size is 100MB.\nTry splitting the document into smaller parts.');
                }
                return;
            }

            sfdt = await response.text();
        }

        console.log("🧠 Rendering document...");
        const renderStart = performance.now();

        window.loadDocument(sfdt);

        // Store SFDT to IndexedDB so tab switching works
        var btnSaveAs = document.getElementById('btnSaveAs');
        var tabId = btnSaveAs ? btnSaveAs.getAttribute('data-tab-id') : '';
        if (tabId && sfdt && sfdt.length > 10) {
            await window.storeSfdtToDb(tabId, sfdt);
            console.log('[Open] Stored to IndexedDB: tab=' + tabId.substring(0, 8) + ' size=' + Math.round(sfdt.length / 1024) + 'KB');
        }

        // Notify Blazor to update tab state (IsEmpty = false, FileName)
        if (window._blazorEditorRef && file) {
            try {
                await window._blazorEditorRef.invokeMethodAsync('OnFileOpenedFromJS', file.name || 'document.docx');
            } catch (e) { console.warn('[Open] Blazor notify failed:', e); }
        }

        setTimeout(() => {
            const endTime = performance.now();
            console.log("🎉 Document fully loaded");
            console.log(`⏱ TOTAL TIME: ${(endTime - startTime).toFixed(2)} ms`);
            hideLoader();
        }, 300);

    } catch (e) {
        hideLoader();
        if (e.name === 'AbortError') {
            console.log('[Open] User cancelled file picker');
            return;
        }
        console.error("❌ Upload failed:", e);
    }
};

// ── Fallback download (for browsers without File System Access API) ──────────
window.downloadFile = function (base64, fileName) {
    var bytes = Uint8Array.from(atob(base64), function (c) { return c.charCodeAt(0); });
    var blob  = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    var url   = URL.createObjectURL(blob);
    var a     = document.createElement("a");
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
};

window.downloadFileAs = function (base64, fileName, mimeType) {
    var bytes = Uint8Array.from(atob(base64), function (c) { return c.charCodeAt(0); });
    var blob  = new Blob([bytes], { type: mimeType || 'application/octet-stream' });
    var url   = URL.createObjectURL(blob);
    var a     = document.createElement('a');
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
};


// ══════════════════════════════════════════════════════════════════════════════
// THEME
// ══════════════════════════════════════════════════════════════════════════════
window.applyTheme = function (isDark) {
    if (isDark) document.documentElement.setAttribute('data-theme', 'dark');
    else        document.documentElement.removeAttribute('data-theme');
    try { localStorage.setItem('deditor-theme', isDark ? 'dark' : 'light'); } catch (e) {}
};

window.getThemePreference = function () {
    try { return localStorage.getItem('deditor-theme') || 'light'; } catch (e) { return 'light'; }
};

// ══════════════════════════════════════════════════════════════════════════════
// NAVIGATION PANE
// ══════════════════════════════════════════════════════════════════════════════
window.toggleNavigationPane = function (open) {
    if (!container) return;
    if (open) {
        container.documentEditor.showOptionsPane();
    } else {
        try {
            var mod = container.documentEditor.optionsPaneModule || container.documentEditor['optionsPaneModule'];
            if (mod && typeof mod.hideOptionPane === 'function') { mod.hideOptionPane(); return; }
            if (mod) {
                if (typeof mod.closePane === 'function')             { mod.closePane(); return; }
                if (typeof mod.showHideOptionPane === 'function')     { mod.showHideOptionPane(false); return; }
            }
        } catch (e) {}
        var closeBtn = document.querySelector('.e-documenteditor-optionspane .e-de-op-close-button') ||
                       document.querySelector('.e-de-op .e-de-close-icon') ||
                       document.querySelector('.e-documenteditor-optionspane [title="Close"]');
        if (closeBtn) closeBtn.click();
    }
};

// ══════════════════════════════════════════════════════════════════════════════
// TRACK CHANGES
// ══════════════════════════════════════════════════════════════════════════════
window.toggleTrackChanges = function () {
    if (!container) return false;
    var current = container.documentEditor.enableTrackChanges;
    container.documentEditor.enableTrackChanges = !current;
    return !current;
};
window.acceptAllChanges = function () { if (container) container.documentEditor.revisions.acceptAll(); };
window.rejectAllChanges = function () { if (container) container.documentEditor.revisions.rejectAll(); };

// ══════════════════════════════════════════════════════════════════════════════
// LOAD / SERIALIZE
// ══════════════════════════════════════════════════════════════════════════════
window.loadBlankDocument = function () {
    if (!container) return;
    container.documentEditor.openBlank();
    container.documentEditor.focusIn();
};
window.loadDocument = function (sfdt) {
    if (!container) return;
    console.log('[Render] Opening document (' + Math.round((sfdt || '').length / 1024) + 'KB)...');
    var start = performance.now();
    container.documentEditor.open(sfdt);
    console.log('[Render] open() took ' + (performance.now() - start).toFixed(0) + 'ms');
    container.documentEditor.focusIn();
};
window.getDocumentContent = function () {
    if (!container) return null;
    console.log('[Save] Serializing document...');
    var start = performance.now();
    var result = container.documentEditor.serialize();
    console.log('[Save] Serialized in ' + (performance.now() - start).toFixed(0) + 'ms (' + Math.round((result || '').length / 1024) + 'KB)');
    return result;
};
window.saveAndSwitch = function (newSfdt, isBlank) {
    if (!container) { console.error('saveAndSwitch: container null'); return null; }
    var oldSfdt = container.documentEditor.serialize();
    if (isBlank) container.documentEditor.openBlank();
    else         container.documentEditor.open(newSfdt);
    container.documentEditor.focusIn();
    return oldSfdt;
};

// ══════════════════════════════════════════════════════════════════════════════
// SELECTION HELPERS
// ══════════════════════════════════════════════════════════════════════════════
window.getSelectedText = function () {
    if (!container) return "";
    return container.documentEditor.selection.text;
};
window.replaceSelectedText = function (newText) {
    if (!container || !newText) return;
    var editor    = container.documentEditor.editor;
    var selection = container.documentEditor.selection;
    if (!selection || !selection.text || selection.text.length === 0) { editor.insertText(newText); return; }
    editor.delete();
    editor.insertText(newText);
};
window.hasSelection = function () {
    if (!container) return false;
    var sel = container.documentEditor.selection.text;
    return sel !== null && sel.length > 0;
};

// ══════════════════════════════════════════════════════════════════════════════
// CASE TRANSFORMATIONS
// ══════════════════════════════════════════════════════════════════════════════
window.applyTitleCase = function () {
    var text = window.getSelectedText();
    if (!text) return "NO_SELECTION";
    var minorWords = new Set(["a","an","the","and","but","or","nor","for","so","yet","as","at","by","in","of","on","to","up","via","per","vs","etc"]);
    var tokens = text.split(/(\s+)/);
    var wordIndex = 0;
    var wordCount = tokens.filter(function(t){ return t.trim().length > 0; }).length;
    var result = tokens.map(function(token) {
        if (token.trim().length === 0) return token;
        var lower = token.toLowerCase();
        var isFirst = wordIndex === 0, isLast = wordIndex === wordCount - 1;
        wordIndex++;
        return (isFirst || isLast || !minorWords.has(lower)) ? capitalizeFirst(token) : lower;
    });
    window.replaceSelectedText(result.join(""));
    return "OK";
};
window.applyInitialCaps = function () {
    var text = window.getSelectedText();
    if (!text) return "NO_SELECTION";
    window.replaceSelectedText(text.split(/(\s+)/).map(function(t){ return t.trim().length === 0 ? t : capitalizeFirst(t); }).join(""));
    return "OK";
};
window.applyEssentialCaps = function () {
    var text = window.getSelectedText();
    if (!text) return "NO_SELECTION";
    if (typeof nlp === "undefined") return "NLP_NOT_LOADED";
    var alwaysLower = new Set(["a","an","the","and","but","or","nor","for","so","yet","in","of","on","at","by","to","up","as","via","per","vs","with","from","into","onto","is","are","was","were","be","been","being","has","have","had","do","does","did"]);
    var doc = nlp(text), importantWords = new Set();
    doc.nouns().out('array').forEach(function(p){ p.split(/\s+/).forEach(function(w){ if(w) importantWords.add(w.toLowerCase()); }); });
    doc.match('#ProperNoun').out('array').forEach(function(p){ p.split(/\s+/).forEach(function(w){ if(w) importantWords.add(w.toLowerCase()); }); });
    doc.match('#Acronym').out('array').forEach(function(p){ p.split(/\s+/).forEach(function(w){ if(w) importantWords.add(w.toLowerCase()); }); });
    var result = text.split(/(\s+)/).map(function(token) {
        if (token.trim().length === 0) return token;
        var lower = token.toLowerCase();
        if (alwaysLower.has(lower)) return lower;
        if (importantWords.has(lower)) return capitalizeFirst(token);
        if (token.indexOf("-") > -1 && token.split("-").some(function(p){ return importantWords.has(p.toLowerCase()); })) return capitalizeFirst(token);
        return lower;
    });
    window.replaceSelectedText(result.join(""));
    return "OK";
};
function capitalizeFirst(word) { return word ? word.charAt(0).toUpperCase() + word.slice(1) : word; }

// ══════════════════════════════════════════════════════════════════════════════
// DOCUMENT TEXT
// ══════════════════════════════════════════════════════════════════════════════
window.getDocumentText = function () {
    if (!container) return "";
    container.documentEditor.selection.selectAll();
    var text = container.documentEditor.selection.text;
    container.documentEditor.selection.clear();
    return text || "";
};

// ══════════════════════════════════════════════════════════════════════════════
// DOCUMENT OUTLINE
// ══════════════════════════════════════════════════════════════════════════════
window.getDocumentOutline = function () {
    if (!container) return [];
    try {
        var sfdt = JSON.parse(container.documentEditor.serialize());
        var results = [];
        var sections = sfdt.sections || [];
        sections.forEach(function(section) {
            var blocks = section.blocks || [];
            blocks.forEach(function(block) {
                _extractHeadings(block, results);
            });
        });
        return results;
    } catch (e) {
        console.warn('[Outline] parse error:', e);
        return [];
    }
};

function _extractHeadings(block, results) {
    if (!block) return;
    if (block.paragraphFormat && block.paragraphFormat.styleName) {
        var style = block.paragraphFormat.styleName;
        var level = 0;
        if (/^Heading 1$/i.test(style)) level = 1;
        else if (/^Heading 2$/i.test(style)) level = 2;
        else if (/^Heading 3$/i.test(style)) level = 3;
        if (level > 0) {
            var text = '';
            (block.inlines || []).forEach(function(inline) {
                if (inline.text) text += inline.text;
            });
            text = text.trim();
            if (text) results.push({ level: level, text: text });
        }
    }
    if (block.rows) {
        block.rows.forEach(function(row) {
            (row.cells || []).forEach(function(cell) {
                (cell.blocks || []).forEach(function(b) { _extractHeadings(b, results); });
            });
        });
    }
}

// ══════════════════════════════════════════════════════════════════════════════
// DOCUMENT NAME
// ══════════════════════════════════════════════════════════════════════════════
window.setDocumentName = function (name) {
    var display = (name && name.trim()) ? name.trim() : 'Untitled.docx';
    var el = document.getElementById('left-sidebar-filename');
    if (el) el.textContent = display;
    document.title = display + ' \u2014 DEditor';
};

// ══════════════════════════════════════════════════════════════════════════════
// RESIZE
// ══════════════════════════════════════════════════════════════════════════════
window.resizeEditor = function () {
    if (!container) return;
    requestAnimationFrame(function () { requestAnimationFrame(function () { safeResize(5); }); });
};

// ══════════════════════════════════════════════════════════════════════════════
// SEARCH & REPLACE
// ══════════════════════════════════════════════════════════════════════════════
window.searchInDocument = function (query) {
    if (!container || !query) return 0;
    container.documentEditor.search.findAll(query, 'None');
    var count = container.documentEditor.search.searchResults.length;
    refocusSearch();
    return count;
};
window.searchNavigate = function (direction) {
    if (!container) return;
    var results = container.documentEditor.search.searchResults;
    if (!results || results.length === 0) return;
    if (direction === 'next') results.index = (results.index + 1) % results.length;
    else                      results.index = (results.index - 1 + results.length) % results.length;
    refocusSearch();
};
window.replaceAllInDocument = function (searchText, replaceText) {
    if (!container || !searchText) return 0;
    container.documentEditor.search.findAll(searchText, 'None');
    var count = container.documentEditor.search.searchResults.length;
    if (count > 0) container.documentEditor.search.searchResults.replaceAll(replaceText || '');
    refocusSearch();
    return count;
};
window.clearSearch = function () { if (container) container.documentEditor.search.searchResults.clear(); };
function refocusSearch() {
    setTimeout(function () {
        var el = document.getElementById('searchInput');
        if (el) { el.focus(); var len = el.value.length; el.setSelectionRange(len, len); }
    }, 30);
}

window.scrollChatToBottom = function () {
    var el = document.getElementById("chatMessages");
    if (el) el.scrollTop = el.scrollHeight;
};

// ══════════════════════════════════════════════════════════════════════════════
// AUTO-SAVE & RECOVERY
// ══════════════════════════════════════════════════════════════════════════════
var AUTOSAVE_INTERVAL = 60000;
var _autoSaveTimer    = null;
var _lastSavedSfdt    = '';

window.startAutoSave = function (tabId, fileName) {
    // AutoSave disabled
};

window.stopAutoSave = function () {
    clearInterval(_autoSaveTimer);
    _autoSaveTimer = null;
};

window.autoSaveNow = function (tabId, fileName) {
    if (!container || !tabId) return false;
    try {
        var sfdt = container.documentEditor.serialize();
        if (!sfdt || sfdt === _lastSavedSfdt) return false;
        if (sfdt.length > 4 * 1024 * 1024) {
            console.log('[AutoSave] Skipped — doc too large (' + Math.round(sfdt.length / 1024 / 1024) + 'MB)');
            return false;
        }
        var entry = JSON.stringify({ sfdt: sfdt, fileName: fileName || 'Untitled.docx', savedAt: new Date().toISOString() });
        localStorage.setItem('deditor-autosave-v1:' + tabId, entry);
        _lastSavedSfdt = sfdt;
        var ids = _getAutoSaveIds();
        if (!ids.includes(tabId)) { ids.push(tabId); localStorage.setItem('deditor-autosave-ids', JSON.stringify(ids)); }
        var ts = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        document.dispatchEvent(new CustomEvent('deditor-autosaved', { detail: ts }));
        return true;
    } catch (e) { console.warn('[AutoSave] write failed:', e); return false; }
};

window.clearAutoSaveDraft = function (tabId) {
    if (!tabId) return;
    try {
        localStorage.removeItem('deditor-autosave-v1:' + tabId);
        var ids = _getAutoSaveIds().filter(function(id){ return id !== tabId; });
        localStorage.setItem('deditor-autosave-ids', JSON.stringify(ids));
        _lastSavedSfdt = '';
    } catch (e) {}
};

window.checkAutoSaveRecovery = function () {
    try {
        return _getAutoSaveIds().map(function(tabId) {
            var raw = localStorage.getItem('deditor-autosave-v1:' + tabId);
            if (!raw) return null;
            try {
                var e = JSON.parse(raw);
                return { tabId: tabId, fileName: e.fileName, savedAt: e.savedAt, sfdt: e.sfdt };
            } catch(ex) { return null; }
        }).filter(Boolean);
    } catch (e) { return []; }
};

window.loadRecoveredDraft = function (sfdt) {
    if (!container || !sfdt) return;
    container.documentEditor.open(sfdt);
    container.documentEditor.focusIn();
};

window.dismissAllDrafts = function () {
    try {
        _getAutoSaveIds().forEach(function(id){ localStorage.removeItem('deditor-autosave-v1:' + id); });
        localStorage.removeItem('deditor-autosave-ids');
    } catch (e) {}
};

function _getAutoSaveIds() {
    try { var r = localStorage.getItem('deditor-autosave-ids'); return r ? JSON.parse(r) : []; } catch (e) { return []; }
}

window.registerAutoSaveListener = function (dotNetRef) {
    document.addEventListener('deditor-autosaved', function (e) {
        try { dotNetRef.invokeMethodAsync('OnAutoSaved', e.detail); } catch (err) {}
    });
};

// ══════════════════════════════════════════════════════════════════════════════
// RECENT FILES
// ══════════════════════════════════════════════════════════════════════════════
window.getRecentFiles = function () {
    try {
        var raw = localStorage.getItem('deditor-recent-files');
        return raw ? JSON.parse(raw) : [];
    } catch (e) { return []; }
};

window.pushRecentFile = function (name, openedAt) {
    try {
        var files = window.getRecentFiles().filter(function(f){ return f.name !== name; });
        files.unshift({ name: name, openedAt: openedAt });
        if (files.length > 10) files = files.slice(0, 10);
        localStorage.setItem('deditor-recent-files', JSON.stringify(files));
    } catch (e) {}
};

window.clearRecentFiles = function () {
    try { localStorage.removeItem('deditor-recent-files'); } catch (e) {}
};

// ══════════════════════════════════════════════════════════════════════════════
// SLASH-COMMAND SNIPPET INSERTION
// ══════════════════════════════════════════════════════════════════════════════
window.insertSnippetText = function (body) {
    if (!container || !body) return;
    var editor = container.documentEditor.editor;
    editor.delete();
    var lines = body.split('\n');
    lines.forEach(function(line, idx) {
        editor.insertText(line);
        if (idx < lines.length - 1) editor.onEnter();
    });
    container.documentEditor.focusIn();
};

// ══════════════════════════════════════════════════════════════════════════════
// KEYBOARD LISTENERS
// ══════════════════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════════════════
// SAVE AS — runs entirely in JS to preserve user gesture for file picker
// ══════════════════════════════════════════════════════════════════════════════
function _doSaveAs() {
    var btn = document.getElementById('btnSaveAs');
    var tabId     = btn ? btn.getAttribute('data-tab-id') : '';
    var fileName  = btn ? btn.getAttribute('data-filename') : 'Untitled.docx';
    var ext       = '.docx';
    if (fileName) {
        var dotIdx = fileName.lastIndexOf('.');
        if (dotIdx >= 0) ext = fileName.substring(dotIdx).toLowerCase();
    }
    if (ext !== '.docx' && ext !== '.doc' && ext !== '.rtf' && ext !== '.txt' && ext !== '.xml') ext = '.docx';

    var mimeMap = {
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.doc':  'application/msword',
        '.rtf':  'application/rtf',
        '.txt':  'text/plain',
        '.xml':  'application/xml'
    };
    var descMap = {
        '.docx': 'Word Document', '.doc': 'Word 97-2003 Document',
        '.rtf': 'Rich Text Format', '.txt': 'Plain Text', '.xml': 'Word XML Document'
    };

    var types = [{ description: descMap[ext] || 'Document', accept: {} }];
    types[0].accept[mimeMap[ext] || mimeMap['.docx']] = [ext];
    Object.keys(mimeMap).forEach(function (e) {
        if (e !== ext) { var t = { description: descMap[e], accept: {} }; t.accept[mimeMap[e]] = [e]; types.push(t); }
    });

    // 1) Open picker IMMEDIATELY (user gesture still valid)
    window.showSaveFilePicker({
        suggestedName: fileName || 'Untitled.docx',
        types: types
    }).then(function (handle) {
        // 2) Serialize
        var sfdt = container ? container.documentEditor.serialize() : null;
        if (!sfdt) { _notifyBlazorSaveAs('ERROR:No document content'); return; }
        var chosenName = handle.name || fileName;

        // 3) POST to server
        return fetch('/api/documenteditor/save', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ Content: sfdt, FileName: chosenName })
        }).then(function (resp) {
            if (!resp.ok) throw new Error('Server returned ' + resp.status);
            return resp.blob();
        }).then(function (blob) {
            // 4) Write to file
            return handle.createWritable().then(function (w) {
                return w.write(blob).then(function () { return w.close(); });
            }).then(function () {
                _fileHandles[tabId] = handle;
                _notifyBlazorSaveAs('OK:' + chosenName);
            });
        });
    }).catch(function (e) {
        if (e.name === 'AbortError') { _notifyBlazorSaveAs('CANCELLED'); return; }
        console.error('[SaveAs]', e);
        _notifyBlazorSaveAs('ERROR:' + e.message);
    });
}

function _notifyBlazorSaveAs(result) {
    if (window._blazorEditorRef) {
        try { window._blazorEditorRef.invokeMethodAsync('OnSaveAsCompleted', result); } catch (e) {}
    }
}

// Intercept Save As button clicks — must happen at DOM level to preserve gesture
document.addEventListener('click', function (e) {
    var btn = e.target.closest('#btnSaveAs');
    if (!btn) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    if (typeof window.showSaveFilePicker !== 'function') {
        // Fallback — let Blazor handle via download
        if (window._blazorEditorRef) {
            try { window._blazorEditorRef.invokeMethodAsync('OnSaveShortcut'); } catch (err) {}
        }
        return;
    }
    _doSaveAs();
}, true); // capture phase — fires before Blazor
window.registerEditorKeyListeners = function (dotNetRef) {
    window._blazorEditorRef = dotNetRef;

    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') {
            try { dotNetRef.invokeMethodAsync('OnEscapePressed'); } catch (err) {}
        }
    }, true);

    document.addEventListener('keydown', function (e) {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && !e.shiftKey) {
            e.preventDefault();
            e.stopPropagation();
            try { dotNetRef.invokeMethodAsync('OnSaveShortcut'); } catch (err) {}
            return;
        }
if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && e.shiftKey) {
            e.preventDefault();
            e.stopPropagation();
            _doSaveAs();
            return;
        }
    }, true);

    function attachSlashListener() {
        var canvas = document.getElementById('container1_editor_viewerContainer');
        if (!canvas || !window._deContainer) { setTimeout(attachSlashListener, 600); return; }

        canvas.addEventListener('keyup', function (e) {
            if (e.key !== '/') return;
            var sel = window._deContainer ? window._deContainer.documentEditor.selection : null;
            if (!sel) return;
            var rect = canvas.getBoundingClientRect();
            var x = rect.left + 80;
            var y = rect.top  + 120;
            try {
                var caretRect = sel.getPhysicalPositionOfCursor ? sel.getPhysicalPositionOfCursor() : null;
                if (caretRect) { x = caretRect.x + rect.left; y = caretRect.y + rect.top; }
            } catch (ex) {}
            try { dotNetRef.invokeMethodAsync('OnSlashTyped', x, y); } catch (err) {}
        });
    }
    setTimeout(attachSlashListener, 1500);
};