// ══════════════════════════════════════════════════════════════════════════════
// MULTI-EDITOR INSTANCE ARCHITECTURE
// ══════════════════════════════════════════════════════════════════════════════
// Instead of one editor that serializes/deserializes on every tab switch,
// we keep one DocumentEditorContainer per tab and toggle display:none.
// `container` always points to the active editor — all existing code works unchanged.
// ══════════════════════════════════════════════════════════════════════════════

var container;                // Always points to the active editor — backwards compatible
var _editors = {};            // tabId → { container, divId }
var _activeEditorTabId = '';  // Currently visible editor's tabId

// ── safeResize ────────────────────────────────────────────────────────────────
function safeResize(attemptsLeft) {
    if (!container) return;
    var parent = document.querySelector('.n-editor-area') || document.getElementById('editorArea');
    var w = parent ? parent.offsetWidth : 0;
    if (w > 0) {
        container.resize();
        // Refresh the active editor's ribbon layout
        _refreshActiveRibbon();
        return;
    }
    if (attemptsLeft <= 0) return;
    requestAnimationFrame(function () { setTimeout(function () { safeResize(attemptsLeft - 1); }, 50); });
}

// Find and refresh the ribbon component on the active editor
function _refreshActiveRibbon() {
    if (!_activeEditorTabId || !_editors[_activeEditorTabId]) return;
    var divId = _editors[_activeEditorTabId].divId;
    var editorDiv = document.getElementById(divId);
    if (!editorDiv) return;
    try {
        var ribbonEl = editorDiv.querySelector('.e-ribbon');
        if (ribbonEl && ribbonEl.ej2_instances && ribbonEl.ej2_instances.length > 0) {
            ribbonEl.ej2_instances[0].refreshLayout();
        }
    } catch (e) {}
}

// ══════════════════════════════════════════════════════════════════════════════
// INDEXEDDB — Tab SFDT Storage (kept for auto-save/recovery, NOT for tab switching)
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
            tx.oncomplete = function () { resolve(true); };
            tx.onerror = function (e) { console.error('[TabDB] SAVE FAILED:', e); resolve(false); };
        } catch (e) { console.error('[TabDB] save exception:', e); resolve(false); }
    });
};

window.tabDbLoad = async function (tabId) {
    await _tabDbReady;
    if (!_tabDb) { console.error('[TabDB] load: DB not ready!'); return null; }
    if (!tabId) { console.error('[TabDB] load: no tabId!'); return null; }
    return new Promise(function (resolve) {
        try {
            var tx = _tabDb.transaction('tabs', 'readonly');
            var req = tx.objectStore('tabs').get(tabId);
            req.onsuccess = function () { resolve(req.result || null); };
            req.onerror = function (e) { console.error('[TabDB] LOAD FAILED:', e); resolve(null); };
        } catch (e) { console.error('[TabDB] load exception:', e); resolve(null); }
    });
};

window.tabDbDelete = async function (tabId) {
    await _tabDbReady;
    if (!_tabDb) return;
    try {
        var tx = _tabDb.transaction('tabs', 'readwrite');
        tx.objectStore('tabs').delete(tabId);
    } catch (e) { console.error('[TabDB] delete error:', e); }
};

window.tabDbHas = async function (tabId) {
    var data = await window.tabDbLoad(tabId);
    return data != null && data.length > 0;
};

// Save current active editor to IndexedDB (for auto-save/recovery only)
window.saveCurrentEditorToDb = async function (tabId) {
    if (!container || !tabId) return false;
    var sfdt = container.documentEditor.serialize();
    if (!sfdt || sfdt.length < 10) return false;
    return await window.tabDbSave(tabId, sfdt);
};

// Load tab from IndexedDB into the ACTIVE editor (used for recovery/restore)
window.loadTabFromDb = async function (tabId) {
    showLoader();
    var sfdt = await window.tabDbLoad(tabId);
    if (sfdt && sfdt.length > 10) {
        container.documentEditor.open(sfdt);
        container.documentEditor.focusIn();
        hideLoader();
        return true;
    }
    hideLoader();
    return false;
};

window.storeSfdtToDb = async function (tabId, sfdt) {
    await window.tabDbSave(tabId, sfdt);
};

// ══════════════════════════════════════════════════════════════════════════════
// MULTI-EDITOR — Create / Switch / Destroy
// ══════════════════════════════════════════════════════════════════════════════

// Create a new editor instance for a tab
window.createEditorForTab = function (tabId) {
    if (_editors[tabId]) {
        console.warn('[MultiEditor] Editor already exists for tab ' + tabId.substring(0, 8));
        _setActiveEditor(tabId);
        return true;
    }

    // Hide the currently active editor FIRST so the new one gets full layout space
    if (_activeEditorTabId && _editors[_activeEditorTabId]) {
        var prevDiv = document.getElementById(_editors[_activeEditorTabId].divId);
        if (prevDiv) prevDiv.style.display = 'none';
    }

    var divId = 'de_' + tabId.replace(/-/g, '');
    var wrapper = document.createElement('div');
    wrapper.id = divId;
    // Start VISIBLE so Syncfusion can measure DOM during initialization
    wrapper.style.cssText = 'width:100%;height:100%;';
    document.getElementById('editorArea').appendChild(wrapper);

    console.log('[MultiEditor] Creating editor for tab ' + tabId.substring(0, 8) + ' → #' + divId);

    var inst = new ej.documenteditor.DocumentEditorContainer({
        height: "calc(100vh - 46px)",
        width: "100%",
        enableToolbar: true,
        toolbarMode: 'Ribbon',
        serviceUrl: '/api/documenteditor/',
    });
    inst.appendTo('#' + divId);
    inst.documentEditor.enableTrackChanges = false;
    inst.documentEditor.showRevisions = false;
    inst.documentEditor.enableOptimizedTextMeasuring = true;

    // Disable Syncfusion built-in Ctrl+S for this instance
    inst.documentEditor.keyDown = function (args) {
        if (!args || !args.event) return;
        var e = args.event;
        if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
            args.isHandled = true;
        }
    };

    _editors[tabId] = { container: inst, divId: divId };

    // Make this the active editor
    _setActiveEditor(tabId);

    // Attach slash command listener to this editor's canvas
    _attachSlashToEditor(divId);

    console.log('[MultiEditor] Editor created. Total instances: ' + Object.keys(_editors).length);
    return true;
};

// Switch active editor (just show/hide — INSTANT)
function _setActiveEditor(tabId) {
    // Hide previous
    if (_activeEditorTabId && _editors[_activeEditorTabId]) {
        var prevDiv = document.getElementById(_editors[_activeEditorTabId].divId);
        if (prevDiv) prevDiv.style.display = 'none';
    }

    _activeEditorTabId = tabId;
    var entry = _editors[tabId];
    if (!entry) { console.error('[MultiEditor] No editor for tab ' + tabId); return; }

    // Show new
    var div = document.getElementById(entry.divId);
    if (div) div.style.display = '';

    // Update global references so ALL existing code works unchanged
    container = entry.container;
    window._deContainer = container;

    // Resize after layout settles
    setTimeout(function () {
        try { container.resize(); } catch (e) {}
        try { container.documentEditor.focusIn(); } catch (e) {}
    }, 30);
    if (window.initSelectionToolbar) {
        setTimeout(window.initSelectionToolbar, 100);
    }
}

// Tab switch: now just show/hide — NO serialize, NO IndexedDB, NO open()!
window.switchTab = function (currentTabId, nextTabId, nextIsEmpty) {
    console.log('[MultiEditor] === SWITCH TAB === ' + (currentTabId || '').substring(0, 8) + ' → ' + nextTabId.substring(0, 8));
    var start = performance.now();
    _setActiveEditor(nextTabId);
    console.log('[MultiEditor] === SWITCH COMPLETE in ' + (performance.now() - start).toFixed(1) + 'ms ===');
};

// Destroy editor instance when tab is closed
window.destroyEditorForTab = function (tabId) {
    var entry = _editors[tabId];
    if (!entry) return;
    console.log('[MultiEditor] Destroying editor for tab ' + tabId.substring(0, 8));
    try { entry.container.destroy(); } catch (e) { console.warn('[MultiEditor] destroy error:', e); }
    var div = document.getElementById(entry.divId);
    if (div) div.remove();
    delete _editors[tabId];
    // Clean up IndexedDB too
    window.tabDbDelete(tabId);
    console.log('[MultiEditor] Destroyed. Remaining instances: ' + Object.keys(_editors).length);
};

// Attach slash-command listener to an editor's canvas
function _attachSlashToEditor(divId) {
    function tryAttach() {
        var div = document.getElementById(divId);
        if (!div || !window._blazorEditorRef) { setTimeout(tryAttach, 600); return; }
        var canvas = div.querySelector('[id$="_editor_viewerContainer"]');
        if (!canvas) { setTimeout(tryAttach, 600); return; }

        canvas.addEventListener('keyup', function (e) {
            if (e.key !== '/') return;
            var sel = container ? container.documentEditor.selection : null;
            if (!sel) return;
            var rect = canvas.getBoundingClientRect();
            var x = rect.left + 80;
            var y = rect.top + 120;
            try {
                var caretRect = sel.getPhysicalPositionOfCursor ? sel.getPhysicalPositionOfCursor() : null;
                if (caretRect) { x = caretRect.x + rect.left; y = caretRect.y + rect.top; }
            } catch (ex) {}
            try { window._blazorEditorRef.invokeMethodAsync('OnSlashTyped', x, y); } catch (err) {}
        });
    }
    setTimeout(tryAttach, 1500);
}


// ══════════════════════════════════════════════════════════════════════════════
// Initialize — creates the FIRST editor + global event wiring
// ══════════════════════════════════════════════════════════════════════════════
window.initializeDocumentEditor = function (firstTabId) {
    // Create the first editor instance
    window.createEditorForTab(firstTabId);

    // ResizeObserver on the editor area
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
    _winResizeTimer = setTimeout(function () {
        safeResize(5);
    }, 150);
});

    _interceptFileMenu();
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

        // Close the file menu — but ONLY on the active editor's div
        try {
            if (_activeEditorTabId && _editors[_activeEditorTabId]) {
                var activeDiv = document.getElementById(_editors[_activeEditorTabId].divId);
                if (activeDiv) {
                    var fileBtn = activeDiv.querySelector('.e-ribbon-file-menu');
                    if (fileBtn) fileBtn.click();
                }
            }
        } catch (ex) {}

        if (text === 'New') {
            window._blazorEditorRef.invokeMethodAsync('NewTabFromJS');
        } else {
            window._blazorEditorRef.invokeMethodAsync('OpenFileDirect');
        }
    }, true);
}

// ── Loader (simplified — just global overlay) ────────────────────────────────
function showLoader() {
    // Find the active editor's viewer container dynamically
    if (_activeEditorTabId && _editors[_activeEditorTabId]) {
        var divId = _editors[_activeEditorTabId].divId;
        var editorDiv = document.getElementById(divId);
        if (editorDiv) {
            var viewer = editorDiv.querySelector('[id$="_editor_viewerContainer"]');
            if (viewer) {
                var overlay = viewer.querySelector('.editor-loader');
                if (!overlay) {
                    overlay = document.createElement("div");
                    overlay.className = "editor-loader";
                    overlay.innerHTML = '<div class="global-spinner"></div>';
                    viewer.appendChild(overlay);
                }
                overlay.style.display = "flex";
                return;
            }
        }
    }
    // Fallback to global loader
    document.getElementById("globalLoader").style.display = "flex";
}

function hideLoader() {
    // Hide loader in all editors (in case tab switched while loading)
    document.querySelectorAll('.editor-loader').forEach(function (el) {
        el.style.display = "none";
    });
    document.getElementById("globalLoader").style.display = "none";
}

// ══════════════════════════════════════════════════════════════════════════════
// FILE SYSTEM ACCESS API — Notepad-style Save / Save As
// ══════════════════════════════════════════════════════════════════════════════
var _fileHandles = {};

window.hasFileSystemAccess = function () {
    return typeof window.showSaveFilePicker === 'function';
};

window.hasFileHandle = function (tabId) {
    return !!_fileHandles[tabId];
};

window.clearFileHandle = function (tabId) {
    delete _fileHandles[tabId];
};

window.checkFileExists = async function (tabId) {
    if (!_fileHandles[tabId]) return 'NO_HANDLE';
    try {
        var handle = _fileHandles[tabId];
        var perm = await handle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted') {
            perm = await handle.requestPermission({ mode: 'readwrite' });
            if (perm !== 'granted') return 'NO_HANDLE';
        }
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

window.getFileHandleName = function (tabId) {
    if (!_fileHandles[tabId]) return '';
    return _fileHandles[tabId].name || '';
};

window.saveToExistingHandle = async function (tabId, base64) {
    if (!_fileHandles[tabId]) return 'NO_HANDLE';
    try {
        var handle = _fileHandles[tabId];
        var bytes = Uint8Array.from(atob(base64), function (c) { return c.charCodeAt(0); });
        var perm = await handle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted') {
            perm = await handle.requestPermission({ mode: 'readwrite' });
            if (perm !== 'granted') return 'NO_HANDLE';
        }
        try { await handle.getFile(); } catch (fileErr) {
            if (fileErr.name === 'NotFoundError') return 'DELETED';
            throw fileErr;
        }
        var writable = await handle.createWritable();
        await writable.write(bytes);
        await writable.close();
        return 'OK';
    } catch (e) {
        console.warn('[Save] write to handle failed:', e);
        if (e.name === 'NotAllowedError' || e.name === 'SecurityError') { delete _fileHandles[tabId]; return 'NO_HANDLE'; }
        if (e.name === 'NotFoundError') return 'DELETED';
        return 'ERROR:' + e.message;
    }
};

window.saveAsComplete = async function (tabId, suggestedName, formatExt, serverSaveUrl) {
    try {
        var mimeMap = {
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.doc': 'application/msword', '.rtf': 'application/rtf',
            '.txt': 'text/plain', '.xml': 'application/xml'
        };
        var descMap = {
            '.docx': 'Word Document', '.doc': 'Word 97-2003 Document',
            '.rtf': 'Rich Text Format', '.txt': 'Plain Text', '.xml': 'Word XML Document'
        };
        var ext = formatExt || '.docx';
        var mime = mimeMap[ext] || mimeMap['.docx'];
        var desc = descMap[ext] || 'Document';
        var types = [{ description: desc, accept: {} }];
        types[0].accept[mime] = [ext];
        Object.keys(mimeMap).forEach(function (e) {
            if (e !== ext) { var t = { description: descMap[e], accept: {} }; t.accept[mimeMap[e]] = [e]; types.push(t); }
        });

        var handle = await window.showSaveFilePicker({ suggestedName: suggestedName || 'Untitled.docx', types: types });
        var sfdt = container ? container.documentEditor.serialize() : null;
        if (!sfdt) return 'ERROR:No document content';
        var chosenName = handle.name || suggestedName;

        var resp = await fetch(serverSaveUrl || '/api/documenteditor/save', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ Content: sfdt, FileName: chosenName })
        });
        if (!resp.ok) return 'ERROR:Server save failed (' + resp.status + ')';

        var blob = await resp.blob();
        var writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        _fileHandles[tabId] = handle;
        return 'OK:' + chosenName;
    } catch (e) {
        if (e.name === 'AbortError') return 'CANCELLED';
        console.error('[SaveAs] failed:', e);
        return 'ERROR:' + e.message;
    }
};

// ── Open file with native picker ─────────────────────────────────────────────
window.openAndUploadToServer = async function () {
    try {
        console.log("🚀 Open started");
        var startTime = performance.now();
        showLoader();

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

        if (!handles || handles.length === 0) { hideLoader(); return; }

        var file = await handles[0].getFile();
        var fileName = file.name || '';
        var fileExt = fileName.lastIndexOf('.') >= 0 ? fileName.substring(fileName.lastIndexOf('.')).toLowerCase() : '';

        var supportedFormats = ['.docx', '.doc', '.rtf', '.txt', '.sfdt', '.xml'];
        if (!supportedFormats.includes(fileExt)) {
            hideLoader();
            alert('Unsupported file format: "' + fileExt + '"\n\nSupported formats:\n• .docx (Word Document)\n• .doc (Word 97-2003)\n• .rtf (Rich Text Format)\n• .txt (Plain Text)\n• .sfdt (Syncfusion Document)');
            return;
        }

        var sfdt;
        if (fileExt === '.sfdt') {
            sfdt = await file.text();
        } else {
            var formData = new FormData();
            formData.append("files", file);
            var response = await fetch('/api/documenteditor/import', { method: 'POST', body: formData });
            if (!response.ok) {
                var errText = await response.text();
                hideLoader();
                if (errText.indexOf('too large') >= 0 || errText.indexOf('File too large') >= 0) {
                    alert('This file is too large to open.\n\nMaximum supported file size is 100MB.');
                }
                return;
            }
            sfdt = await response.text();
        }

        // Open in the currently active editor
        window.loadDocument(sfdt);

        // Store to IndexedDB for recovery
        var btnSaveAs = document.getElementById('btnSaveAs');
        var tabId = btnSaveAs ? btnSaveAs.getAttribute('data-tab-id') : '';
        if (tabId && sfdt && sfdt.length > 10) {
            await window.storeSfdtToDb(tabId, sfdt);
        }

        // Store file handle for Save
        if (tabId && handles[0]) {
            _fileHandles[tabId] = handles[0];
        }

        if (window._blazorEditorRef && file) {
            try { await window._blazorEditorRef.invokeMethodAsync('OnFileOpenedFromJS', file.name || 'document.docx'); } catch (e) {}
        }

        setTimeout(function () {
            console.log("🎉 Document loaded in " + (performance.now() - startTime).toFixed(0) + "ms");
            hideLoader();
        }, 300);

    } catch (e) {
        hideLoader();
        if (e.name === 'AbortError') return;
        console.error("❌ Upload failed:", e);
    }
};

// ── Download fallback ────────────────────────────────────────────────────────
window.downloadFile = function (base64, fileName) {
    var bytes = Uint8Array.from(atob(base64), function (c) { return c.charCodeAt(0); });
    var blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
};

window.downloadFileAs = function (base64, fileName, mimeType) {
    var bytes = Uint8Array.from(atob(base64), function (c) { return c.charCodeAt(0); });
    var blob = new Blob([bytes], { type: mimeType || 'application/octet-stream' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
};


// ══════════════════════════════════════════════════════════════════════════════
// THEME
// ══════════════════════════════════════════════════════════════════════════════
window.applyTheme = function (isDark) {
    if (isDark) document.documentElement.setAttribute('data-theme', 'dark');
    else document.documentElement.removeAttribute('data-theme');
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
                if (typeof mod.closePane === 'function') { mod.closePane(); return; }
                if (typeof mod.showHideOptionPane === 'function') { mod.showHideOptionPane(false); return; }
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
    if (!container) return Promise.resolve();
    console.log('[Render] Opening document (' + Math.round((sfdt || '').length / 1024) + 'KB)...');
    
    // Yield TWO frames so the browser can paint the loading spinner
    // before the synchronous open() blocks the main thread
    return new Promise(function (resolve) {
        requestAnimationFrame(function () {
            requestAnimationFrame(function () {
                var start = performance.now();
                container.documentEditor.open(sfdt);
                console.log('[Render] open() took ' + (performance.now() - start).toFixed(0) + 'ms');
                container.documentEditor.focusIn();
                resolve();
            });
        });
    });
};
window.getDocumentContent = function () {
    if (!container) return null;
    var start = performance.now();
    var result = container.documentEditor.serialize();
    console.log('[Save] Serialized in ' + (performance.now() - start).toFixed(0) + 'ms (' + Math.round((result || '').length / 1024) + 'KB)');
    return result;
};
window.saveAndSwitch = function (newSfdt, isBlank) {
    if (!container) { console.error('saveAndSwitch: container null'); return null; }
    var oldSfdt = container.documentEditor.serialize();
    if (isBlank) container.documentEditor.openBlank();
    else container.documentEditor.open(newSfdt);
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
    var editor = container.documentEditor.editor;
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
    var wordCount = tokens.filter(function (t) { return t.trim().length > 0; }).length;
    var result = tokens.map(function (token) {
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
    window.replaceSelectedText(text.split(/(\s+)/).map(function (t) { return t.trim().length === 0 ? t : capitalizeFirst(t); }).join(""));
    return "OK";
};
window.applyEssentialCaps = function () {
    var text = window.getSelectedText();
    if (!text) return "NO_SELECTION";
    if (typeof nlp === "undefined") return "NLP_NOT_LOADED";
    var alwaysLower = new Set(["a","an","the","and","but","or","nor","for","so","yet","in","of","on","at","by","to","up","as","via","per","vs","with","from","into","onto","is","are","was","were","be","been","being","has","have","had","do","does","did"]);
    var doc = nlp(text), importantWords = new Set();
    doc.nouns().out('array').forEach(function (p) { p.split(/\s+/).forEach(function (w) { if (w) importantWords.add(w.toLowerCase()); }); });
    doc.match('#ProperNoun').out('array').forEach(function (p) { p.split(/\s+/).forEach(function (w) { if (w) importantWords.add(w.toLowerCase()); }); });
    doc.match('#Acronym').out('array').forEach(function (p) { p.split(/\s+/).forEach(function (w) { if (w) importantWords.add(w.toLowerCase()); }); });
    var result = text.split(/(\s+)/).map(function (token) {
        if (token.trim().length === 0) return token;
        var lower = token.toLowerCase();
        if (alwaysLower.has(lower)) return lower;
        if (importantWords.has(lower)) return capitalizeFirst(token);
        if (token.indexOf("-") > -1 && token.split("-").some(function (p) { return importantWords.has(p.toLowerCase()); })) return capitalizeFirst(token);
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
        sections.forEach(function (section) {
            (section.blocks || []).forEach(function (block) { _extractHeadings(block, results); });
        });
        return results;
    } catch (e) { return []; }
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
            (block.inlines || []).forEach(function (inline) { if (inline.text) text += inline.text; });
            text = text.trim();
            if (text) results.push({ level: level, text: text });
        }
    }
    if (block.rows) {
        block.rows.forEach(function (row) {
            (row.cells || []).forEach(function (cell) {
                (cell.blocks || []).forEach(function (b) { _extractHeadings(b, results); });
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
    else results.index = (results.index - 1 + results.length) % results.length;
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
var _autoSaveTimer = null;
var _lastSavedSfdt = '';

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
        if (sfdt.length > 4 * 1024 * 1024) return false;
        var entry = JSON.stringify({ sfdt: sfdt, fileName: fileName || 'Untitled.docx', savedAt: new Date().toISOString() });
        localStorage.setItem('deditor-autosave-v1:' + tabId, entry);
        _lastSavedSfdt = sfdt;
        var ids = _getAutoSaveIds();
        if (!ids.includes(tabId)) { ids.push(tabId); localStorage.setItem('deditor-autosave-ids', JSON.stringify(ids)); }
        var ts = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        document.dispatchEvent(new CustomEvent('deditor-autosaved', { detail: ts }));
        return true;
    } catch (e) { return false; }
};

window.clearAutoSaveDraft = function (tabId) {
    if (!tabId) return;
    try {
        localStorage.removeItem('deditor-autosave-v1:' + tabId);
        var ids = _getAutoSaveIds().filter(function (id) { return id !== tabId; });
        localStorage.setItem('deditor-autosave-ids', JSON.stringify(ids));
        _lastSavedSfdt = '';
    } catch (e) {}
};

window.checkAutoSaveRecovery = function () {
    try {
        return _getAutoSaveIds().map(function (tabId) {
            var raw = localStorage.getItem('deditor-autosave-v1:' + tabId);
            if (!raw) return null;
            try { var e = JSON.parse(raw); return { tabId: tabId, fileName: e.fileName, savedAt: e.savedAt, sfdt: e.sfdt }; }
            catch (ex) { return null; }
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
        _getAutoSaveIds().forEach(function (id) { localStorage.removeItem('deditor-autosave-v1:' + id); });
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
    try { var raw = localStorage.getItem('deditor-recent-files'); return raw ? JSON.parse(raw) : []; } catch (e) { return []; }
};
window.pushRecentFile = function (name, openedAt) {
    try {
        var files = window.getRecentFiles().filter(function (f) { return f.name !== name; });
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
    lines.forEach(function (line, idx) {
        editor.insertText(line);
        if (idx < lines.length - 1) editor.onEnter();
    });
    container.documentEditor.focusIn();
};

// ══════════════════════════════════════════════════════════════════════════════
// SAVE AS — runs entirely in JS to preserve user gesture for file picker
// ══════════════════════════════════════════════════════════════════════════════
function _doSaveAs() {
    var btn = document.getElementById('btnSaveAs');
    var tabId = btn ? btn.getAttribute('data-tab-id') : '';
    var fileName = btn ? btn.getAttribute('data-filename') : 'Untitled.docx';
    var ext = '.docx';
    if (fileName) {
        var dotIdx = fileName.lastIndexOf('.');
        if (dotIdx >= 0) ext = fileName.substring(dotIdx).toLowerCase();
    }
    if (ext !== '.docx' && ext !== '.doc' && ext !== '.rtf' && ext !== '.txt' && ext !== '.xml') ext = '.docx';

    var mimeMap = {
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.doc': 'application/msword', '.rtf': 'application/rtf',
        '.txt': 'text/plain', '.xml': 'application/xml'
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

    window.showSaveFilePicker({
        suggestedName: fileName || 'Untitled.docx', types: types
    }).then(function (handle) {
        var sfdt = container ? container.documentEditor.serialize() : null;
        if (!sfdt) { _notifyBlazorSaveAs('ERROR:No document content'); return; }
        var chosenName = handle.name || fileName;
        return fetch('/api/documenteditor/save', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ Content: sfdt, FileName: chosenName })
        }).then(function (resp) {
            if (!resp.ok) throw new Error('Server returned ' + resp.status);
            return resp.blob();
        }).then(function (blob) {
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

document.addEventListener('click', function (e) {
    var btn = e.target.closest('#btnSaveAs');
    if (!btn) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    if (typeof window.showSaveFilePicker !== 'function') {
        if (window._blazorEditorRef) {
            try { window._blazorEditorRef.invokeMethodAsync('OnSaveShortcut'); } catch (err) {}
        }
        return;
    }
    _doSaveAs();
}, true);

// ══════════════════════════════════════════════════════════════════════════════
// KEYBOARD LISTENERS
// ══════════════════════════════════════════════════════════════════════════════
window.registerEditorKeyListeners = function (dotNetRef) {
    window._blazorEditorRef = dotNetRef;

    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') {
            try { dotNetRef.invokeMethodAsync('OnEscapePressed'); } catch (err) {}
        }
    }, true);

    document.addEventListener('keydown', function (e) {
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && !e.shiftKey) {
            e.preventDefault(); e.stopPropagation();
            try { dotNetRef.invokeMethodAsync('OnSaveShortcut'); } catch (err) {}
            return;
        }
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && e.shiftKey) {
            e.preventDefault(); e.stopPropagation();
            _doSaveAs();
            return;
        }
    }, true);

    // Note: slash listeners are now attached per-editor in createEditorForTab
};
