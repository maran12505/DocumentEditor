// ══════════════════════════════════════════════════════════════════════════════
// SERVER BASE URL — configurable for MAUI Blazor Hybrid
// In Web (WASM): empty string (same origin). In MAUI: set to server URL.
// Called from Blazor on first render via: setServerBaseUrl("http://localhost:5278")
// ══════════════════════════════════════════════════════════════════════════════
window.__serverBaseUrl = '';

window.setServerBaseUrl = function (url) {
    window.__serverBaseUrl = url || '';
};

function _apiUrl(path) {
    return (window.__serverBaseUrl || '') + path;
}

// ══════════════════════════════════════════════════════════════════════════════
// MULTI-EDITOR INSTANCE ARCHITECTURE
// ══════════════════════════════════════════════════════════════════════════════
// Instead of one editor that serializes/deserializes on every tab switch,
// we keep one DocumentEditorContainer per tab and toggle display:none.
// `container` always points to the active editor — all existing code works unchanged.
// ══════════════════════════════════════════════════════════════════════════════

// ── Patch ALL Syncfusion classes with removeKeytip ──────────────────────────
// The crash is in ect.removeKeytip (minified), NOT Ribbon.removeKeytip.
// Multiple classes have removeKeytip — we must patch ALL of them.
var _ktRetries = 0;
function _applyKeytipPatches() {
    if (typeof ej === 'undefined') {
        _ktRetries++;
        if (_ktRetries > 20) {
            console.error('[KT-FIX] ej never loaded after 10s — giving up');
            return;
        }
        console.warn('[KT-FIX] ej not loaded yet — retry ' + _ktRetries + '/20');
        setTimeout(_applyKeytipPatches, 500);
        return;
    }
    var count = 0;
    function wrapRemoveKeytip(proto, label) {
        if (proto._ktPatched) return;
        var orig = proto.removeKeytip;
        proto.removeKeytip = function () {
            try { return orig.apply(this, arguments); } catch (e) { /* hidden/destroyed editor — safe to ignore */ }
        };
        proto._ktPatched = true;
        count++;
    }

    try {
        Object.keys(ej).forEach(function (ns) {
            try {
                var sub = ej[ns];
                if (!sub || typeof sub !== 'object') return;
                Object.keys(sub).forEach(function (cls) {
                    try {
                        var C = sub[cls];
                        if (typeof C === 'function' && C.prototype && typeof C.prototype.removeKeytip === 'function') {
                            wrapRemoveKeytip(C.prototype, 'ej.' + ns + '.' + cls);
                        }
                    } catch (e) {}
                });
            } catch (e) {}
        });
    } catch (e) {}

    console.log('[KT-FIX] Total patched: ' + count + (count > 0 ? ' ✅' : ' ❌'));
    window._ktPatchCount = count;
}
_applyKeytipPatches();

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
    showLoader('Loading…');
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

    if (typeof ej === 'undefined' || !ej.documenteditor) {
        console.error('[MultiEditor] Syncfusion ej2 library not loaded! Cannot create editor.');
        return false;
    }

    var inst = new ej.documenteditor.DocumentEditorContainer({
        height: "100%",
        width: "100%",
        enableToolbar: true,
        toolbarMode: 'Ribbon',
        serviceUrl: _apiUrl('/api/documenteditor/'),
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

    var editorEntry = {
        container: inst, divId: divId, navPaneOpen: false,
        ccTagsVisible: false,
        cleanSfdt: null,    // original SFDT string (no markers)
        markedSfdt: null,   // SFDT string with tag markers inserted (cached)
        _editedInTagView: false  // true if user edited while tags were visible
    };
    _editors[tabId] = editorEntry;

    // Track edits made while in tag-view mode — invalidates stale cleanSfdt
    inst.documentEditor.contentChange = function () {
        if (editorEntry.ccTagsVisible) {
            editorEntry._editedInTagView = true;
            // Invalidate caches so HIDE will serialize fresh instead of restoring stale cleanSfdt
            editorEntry.markedSfdt = null;
        }
    };

    // Patch Syncfusion's broken removeContentControl with SFDT-based unwrap
    _patchRemoveCC(inst);

    // Make this the active editor
    _setActiveEditor(tabId);

    // Attach slash command listener to this editor's canvas
    _attachSlashToEditor(divId);

    // Inject Developer tab into this editor's ribbon
    setTimeout(function () {
        var editorDiv = document.getElementById(divId);
        if (editorDiv) _injectDeveloperTab(editorDiv);
    }, 500);

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

    // Sync global _ccTagsVisible to this tab's state
    _ccTagsVisible = entry.ccTagsVisible || false;

    // Increment toggle sequence to invalidate any in-flight toggle from the previous tab
    _toggleSeq++;

    // Resize after layout settles
    setTimeout(function () {
        try { container.resize(); } catch (e) {}
        try { container.documentEditor.focusIn(); } catch (e) {}
        // Update toggle button to reflect this tab's tag state
        var edDiv = document.getElementById(entry.divId);
        if (edDiv) {
            var btn = edDiv.querySelector('.dev-cc-toggle');
            if (btn) _updateToggleBtn(btn, _ccTagsVisible);
        }
    }, 30);
    if (window.initSelectionToolbar) {
        setTimeout(window.initSelectionToolbar, 100);
    }
}

window.switchTab = function (currentTabId, nextTabId, nextIsEmpty) {
    console.log('[MultiEditor] === SWITCH TAB === ' + (currentTabId || '').substring(0, 8) + ' → ' + nextTabId.substring(0, 8));
    var start = performance.now();
    _setActiveEditor(nextTabId);
    console.log('[MultiEditor] === SWITCH COMPLETE in ' + (performance.now() - start).toFixed(1) + 'ms ===');
};

window.isNavPaneOpen = function () {
    if (!_activeEditorTabId || !_editors[_activeEditorTabId]) return false;
    return _editors[_activeEditorTabId].navPaneOpen === true;
};

// Destroy editor instance when tab is closed
window.destroyEditorForTab = function (tabId) {
    var entry = _editors[tabId];
    if (!entry) return;
    console.log('[MultiEditor] Destroying editor for tab ' + tabId.substring(0, 8));

    delete _editors[tabId];
    if (_activeEditorTabId === tabId) {
        _activeEditorTabId = '';
        container = null;
    }

    var div = document.getElementById(entry.divId);
    if (div) div.style.display = 'none';

    setTimeout(function () {
        try { entry.container.destroy(); } catch (e) { console.warn('[MultiEditor] destroy error:', e); }
        if (div && div.parentNode) div.parentNode.removeChild(div);
        console.log('[MultiEditor] Destroyed. Remaining instances: ' + Object.keys(_editors).length);
    }, 0);

    window.tabDbDelete(tabId);
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

// ── Loader — targets the document canvas area only (not toolbar/sidebar) ────
function _findCanvasHost() {
    // Try to find the Syncfusion document viewer area inside the active editor
    if (_activeEditorTabId && _editors[_activeEditorTabId]) {
        var divId = _editors[_activeEditorTabId].divId;
        var editorDiv = document.getElementById(divId);
        if (editorDiv) {
            // Best target: the viewerContainer (the actual canvas scroll area)
            var viewer = editorDiv.querySelector('[id$="_editor_viewerContainer"]');
            if (viewer) return viewer;
            // Fallback: the e-de-ctn wrapper (document editor content area)
            var ctn = editorDiv.querySelector('.e-de-ctn');
            if (ctn) return ctn;
            // Last resort: the editor div itself (still scoped per-tab, not full screen)
            return editorDiv;
        }
    }
    // Final fallback: editorArea
    return document.getElementById('editorArea');
}

function showLoader(message) {
    var host = _findCanvasHost();
    if (!host) {
        // Absolute last resort — global overlay
        document.getElementById("globalLoader").style.display = "flex";
        return;
    }
    // Ensure host is positioned for absolute child
    if (getComputedStyle(host).position === 'static') {
        host.style.position = 'relative';
    }
    var overlay = host.querySelector('.doc-canvas-loader');
    if (!overlay) {
        overlay = document.createElement('div');
        overlay.className = 'doc-canvas-loader';
        host.appendChild(overlay);
    }
    // Update message if provided
    var msgEl = overlay.querySelector('.doc-loader-msg');
    if (msgEl) msgEl.textContent = message || '';
    else {
        overlay.innerHTML =
            '<div class="doc-loader-inner">' +
                '<div class="doc-loader-spinner"></div>' +
                '<div class="doc-loader-msg">' + (message || '') + '</div>' +
            '</div>';
    }
    overlay.style.display = 'flex';
}

function hideLoader() {
    // Hide all per-editor loaders (in case tab switched while loading)
    document.querySelectorAll('.doc-canvas-loader').forEach(function (el) {
        el.style.display = 'none';
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

        var resp = await fetch(serverSaveUrl || _apiUrl('/api/documenteditor/save'), {
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
        showLoader('Opening document…');

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
            var response = await fetch(_apiUrl('/api/documenteditor/import'), { method: 'POST', body: formData });
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
    if (_activeEditorTabId && _editors[_activeEditorTabId]) {
        _editors[_activeEditorTabId].navPaneOpen = open;
    }
    if (open) {
        container.documentEditor.showOptionsPane();
        _watchOptionsPaneClose();
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

// Watch for Syncfusion closing the options pane via its own Close button
var _optionsPaneObserver = null;

function _watchOptionsPaneClose() {
    if (_optionsPaneObserver) { try { _optionsPaneObserver.disconnect(); } catch(e){} _optionsPaneObserver = null; }

    var root = null;
    if (_activeEditorTabId && _editors[_activeEditorTabId]) {
        root = document.getElementById(_editors[_activeEditorTabId].divId);
    }
    if (!root) root = document;

    var pane = root.querySelector('.e-documenteditor-optionspane');
    if (!pane) {
        var retries = 0;
        (function retry() {
            if (retries++ > 10) return;
            pane = root.querySelector('.e-documenteditor-optionspane');
            if (pane) _attachPaneWatcher(pane);
            else setTimeout(retry, 100);
        })();
        return;
    }
    _attachPaneWatcher(pane);
}

function _attachPaneWatcher(pane) {
    var watchedTabId = _activeEditorTabId;

    _optionsPaneObserver = new MutationObserver(function (mutations) {
        for (var i = 0; i < mutations.length; i++) {
            if (mutations[i].type === 'attributes' && mutations[i].attributeName === 'style') {
                if (pane.style.display === 'none') {
                    _optionsPaneObserver.disconnect();
                    _optionsPaneObserver = null;
                    _markNavPaneClosed(watchedTabId);
                    return;
                }
            }
        }
    });
    _optionsPaneObserver.observe(pane, { attributes: true, attributeFilter: ['style'] });

    var closeBtn = pane.querySelector('.e-de-op-close-button, [title="Close"]');
    if (closeBtn) {
        closeBtn.addEventListener('click', function onClose() {
            closeBtn.removeEventListener('click', onClose);
            if (_optionsPaneObserver) { _optionsPaneObserver.disconnect(); _optionsPaneObserver = null; }
            setTimeout(function () { _markNavPaneClosed(watchedTabId); }, 50);
        });
    }
}

function _markNavPaneClosed(tabId) {
    if (_editors[tabId]) _editors[tabId].navPaneOpen = false;
    if (tabId === _activeEditorTabId && window._blazorEditorRef) {
        try { window._blazorEditorRef.invokeMethodAsync('OnNavPaneClosed'); }
        catch (e) { console.warn('[NavPane] Could not notify Blazor:', e); }
    }
}

// ══════════════════════════════════════════════════════════════════════════════
// TRACK CHANGES

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
    var ts = _getTabState();
    if (ts && ts.ccTagsVisible) {
        if (ts._editedInTagView) {
            // Edited in tag view — strip markers from current state
            console.log('[Save] Stripping markers from edited tag-view document');
            var current = container.documentEditor.serialize();
            var parsed = JSON.parse(current);
            _stripTagMarkers(parsed);
            return JSON.stringify(parsed);
        }
        if (ts.cleanSfdt) {
            console.log('[Save] Returning clean SFDT (tag markers active, ' + Math.round(ts.cleanSfdt.length / 1024) + 'KB)');
            return ts.cleanSfdt;
        }
    }
    var start = performance.now();
    var result = container.documentEditor.serialize();
    console.log('[Save] Serialized in ' + (performance.now() - start).toFixed(0) + 'ms (' + Math.round((result || '').length / 1024) + 'KB)');
    return result;
};
window.saveAndSwitch = function (newSfdt, isBlank) {
    if (!container) { console.error('saveAndSwitch: container null'); return null; }
    var ts = _getTabState();
    var oldSfdt;
    if (ts && ts.ccTagsVisible) {
        if (ts._editedInTagView) {
            // Strip markers from current state
            var current = container.documentEditor.serialize();
            var parsed = JSON.parse(current);
            _stripTagMarkers(parsed);
            oldSfdt = JSON.stringify(parsed);
            ts._editedInTagView = false;
        } else {
            oldSfdt = ts.cleanSfdt || container.documentEditor.serialize();
        }
    } else {
        oldSfdt = container.documentEditor.serialize();
    }
    // Clear cache on switch — will be rebuilt when user toggles again
    if (ts) { ts.cleanSfdt = null; ts.markedSfdt = null; }
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
        return fetch(_apiUrl('/api/documenteditor/save'), {
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

// ══════════════════════════════════════════════════════════════════════════════
// DEVELOPER TAB — Show/Hide content control tags toggle
// Injects into the EXISTING Syncfusion Developer tab (not a new tab)
//
// Architecture (v3 — per-tab isolated, race-condition-free):
//   Per-tab state lives in _editors[tabId]: { ccTagsVisible, cleanSfdt, markedSfdt }
//   All async callbacks capture references at schedule-time (container, tabId, tabState)
//   and validate the tab is still active before operating.
//   SHOW always serializes fresh — never trusts stale cleanSfdt.
//   HIDE clears ALL cache to prevent stale data on next SHOW.
//   0-CC documents skip the expensive de.open() entirely.
// ══════════════════════════════════════════════════════════════════════════════
var _ccTagsVisible = false;  // Mirrors the ACTIVE tab's state (synced on tab switch)
var _ccToggleBusy = false;
var _toggleSeq = 0;          // Monotonic sequence to detect stale callbacks

// Per-tab state helpers
function _getTabState() {
    var entry = _editors[_activeEditorTabId];
    return entry || null;
}
function _getTabCleanSfdt() {
    var s = _getTabState();
    return s ? s.cleanSfdt : null;
}
function _setTabCleanSfdt(val) {
    var s = _getTabState();
    if (s) s.cleanSfdt = val;
}
function _setTabTagsVisible(val) {
    _ccTagsVisible = val;
    var s = _getTabState();
    if (s) s.ccTagsVisible = val;
}
function _invalidateTabCache(tabState) {
    if (!tabState) return;
    tabState.cleanSfdt = null;
    tabState.markedSfdt = null;
}

// Called once per editor — listens for Developer tab clicks to inject button
function _injectDeveloperTab(editorDiv) {
    if (!editorDiv) return;

    editorDiv.addEventListener('click', function (e) {
        if (_ccToggleBusy) return;
        var tabItem = e.target.closest('.e-toolbar-item');
        if (!tabItem) return;
        var text = (tabItem.textContent || '').trim();
        if (text === 'Developer') {
            setTimeout(function () { _tryInjectToggle(editorDiv); }, 200);
        }
    }, true);

    console.log('[Developer] Listening for Developer tab clicks');
}

function _tryInjectToggle(editorDiv) {
    if (editorDiv.querySelector('.dev-cc-toggle')) return;

    var devPanel = null;
    var ribbonGroups = editorDiv.querySelectorAll('.e-ribbon-group');
    for (var i = 0; i < ribbonGroups.length; i++) {
        var text = ribbonGroups[i].textContent || '';
        if (text.indexOf('Restrict Editing') !== -1 || text.indexOf('Form Fields') !== -1 ||
            text.indexOf('Content Control') !== -1 || text.indexOf('XML Mapping') !== -1) {
            devPanel = ribbonGroups[i].closest('.e-ribbon-tab-item') || ribbonGroups[i].closest('.e-item');
            break;
        }
    }
    if (!devPanel) return;

    var group = document.createElement('div');
    group.className = 'e-ribbon-group dev-toggle-group';
    group.innerHTML =
        '<div class="e-ribbon-group-container">' +
            '<div class="e-ribbon-group-content">' +
                '<div class="e-ribbon-collection">' +
                    '<div class="e-ribbon-item">' +
                        '<button class="e-btn e-ribbon-control dev-cc-toggle" title="Show/Hide content control tags">' +
                            '<span class="e-btn-icon dev-cc-icon"></span>' +
                            '<span class="dev-btn-label"></span>' +
                        '</button>' +
                    '</div>' +
                '</div>' +
            '</div>' +
            '<div class="e-ribbon-group-header">Tags</div>' +
        '</div>';

    devPanel.appendChild(group);

    var toggleBtn = group.querySelector('.dev-cc-toggle');
    _updateToggleBtn(toggleBtn, _ccTagsVisible);

    toggleBtn.addEventListener('click', function (e) {
        e.stopPropagation();
        e.stopImmediatePropagation();
        e.preventDefault();
        if (_ccToggleBusy) return;
        _applyTagVisibility(!_ccTagsVisible);
    });

    console.log('[Developer] Toggle button injected');
}

function _updateToggleBtn(btn, visible) {
    if (!btn) return;
    var label = btn.querySelector('.dev-btn-label');
    var iconWrap = btn.querySelector('.dev-cc-icon');
    if (visible) {
        btn.classList.add('active');
        if (label) label.textContent = 'Tags On';
        if (iconWrap) iconWrap.innerHTML =
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' +
                '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>';
    } else {
        btn.classList.remove('active');
        if (label) label.textContent = 'Tags Off';
        if (iconWrap) iconWrap.innerHTML =
            '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' +
                '<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"/>' +
                '<path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/>' +
                '<path d="M14.12 14.12a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>';
    }
}

// ── Design Mode: Insert/remove visible [tag] markers in document ────────
var _CC_COLOR = '#2383E2';

// ── Pill-style tag marker formatting ──
// Markers are invisible in SFDT (white text, no highlight) — visual pills are HTML overlays
// The text runs still occupy space for offset calculations and search-based positioning
var _MARKER_CF_OPEN  = { "fsz": 7, "fc": "#4A4A4A", "hc": 1, "fn": "Consolas" };
var _MARKER_CF_CLOSE = { "fsz": 7, "fc": "#4A4A4A", "hc": 1, "fn": "Consolas" };
var _MARKER_CF = _MARKER_CF_OPEN;
// Unique zero-width marker prefix so strip logic can identify our markers reliably
var _MARKER_PREFIX = '\u200B';

// Shorten long JATS-style tag names for cleaner display
// e.g. "#book-part;id=2;type=chapter;" → "book-part"
//      "#disp-quote;id=epigraph;" → "disp-quote"
//      "title" → "title"
function _shortTagName(tag) {
    if (!tag) return '?';
    // Strip leading # if present
    var t = tag.charAt(0) === '#' ? tag.substring(1) : tag;
    // Take only the part before the first semicolon (strip attributes)
    var semi = t.indexOf(';');
    if (semi > 0) t = t.substring(0, semi);
    // Trim trailing/leading whitespace
    t = t.replace(/^\s+|\s+$/g, '');
    return t || '?';
}

// ── HTML pill overlay system ──
// Renders styled pill divs over the canvas at marker positions
// Uses Syncfusion search API to locate markers, then page bounding rect for coordinate mapping


function _applyTagVisibility(visible) {
    if (!container || _ccToggleBusy) {
        console.warn('[Toggle] Blocked: container=' + !!container + ' busy=' + _ccToggleBusy);
        return;
    }
    var tabState = _getTabState();
    if (!tabState) { console.warn('[Toggle] No tab state'); return; }

    // ── Freeze all references at invocation time ──
    // These will NOT change even if the user switches tabs during async callbacks
    var capturedContainer = container;
    var capturedTabId = _activeEditorTabId;
    var capturedDivId = tabState.divId;
    var seq = ++_toggleSeq;  // unique sequence for this toggle operation

    _ccToggleBusy = true;
    var t0 = performance.now();
    var tabId8 = capturedTabId.substring(0, 8);

    console.log('[Toggle:' + seq + '] ── ' + (visible ? 'SHOW' : 'HIDE') + ' tags for tab ' + tabId8 + ' ──');

    _showToggleOverlay(capturedDivId, true);

    // Use rAF so overlay paints before the synchronous de.open() blocks the thread
    requestAnimationFrame(function () {
        // ── Stale check: if another toggle fired after us, abort ──
        if (seq !== _toggleSeq) {
            console.warn('[Toggle:' + seq + '] Aborted — superseded by toggle:' + _toggleSeq);
            _ccToggleBusy = false;
            _showToggleOverlay(capturedDivId, false);
            return;
        }
        // ── Tab-switch check: if user switched tabs during rAF, abort ──
        if (_activeEditorTabId !== capturedTabId) {
            console.warn('[Toggle:' + seq + '] Aborted — tab switched from ' + tabId8 + ' to ' + _activeEditorTabId.substring(0, 8));
            _ccToggleBusy = false;
            _showToggleOverlay(capturedDivId, false);
            return;
        }

        try {
            var de = capturedContainer.documentEditor;

            if (visible) {
                // ── SHOW TAGS ──
                // ALWAYS serialize fresh — never trust stale cleanSfdt
                var ts1 = performance.now();
                var freshClean = de.serialize();
                console.log('[Toggle:' + seq + '] serialize() took ' + (performance.now() - ts1).toFixed(0) + 'ms (' + Math.round(freshClean.length / 1024) + 'KB)');
                tabState.cleanSfdt = freshClean;

                // Build marked SFDT
                var tp = performance.now();
                var sfdt = JSON.parse(freshClean);
                console.log('[Toggle:' + seq + '] JSON.parse() took ' + (performance.now() - tp).toFixed(0) + 'ms');

                var tw = performance.now();
                var count = _insertTagMarkers(sfdt);
                console.log('[Toggle:' + seq + '] _insertTagMarkers() patched ' + count + ' CC(s) in ' + (performance.now() - tw).toFixed(0) + 'ms');

                if (count === 0) {
                    // No CCs to mark — skip the expensive de.open(), just update state
                    console.log('[Toggle:' + seq + '] 0 CCs — skipping de.open(), no visual change');
                    tabState.markedSfdt = null;
                    _setTabTagsVisible(true);
                    _ccToggleBusy = false;
                    _showToggleOverlay(capturedDivId, false);
                    // Update button without restoring dev tab (no de.open happened)
                    var btn0 = document.getElementById(capturedDivId);
                    if (btn0) { var b = btn0.querySelector('.dev-cc-toggle'); _updateToggleBtn(b, true); }
                    console.log('[Toggle:' + seq + '] ── Total: ' + (performance.now() - t0).toFixed(0) + 'ms (no-op) ──');
                    return;
                }

                var tj = performance.now();
                tabState.markedSfdt = JSON.stringify(sfdt);
                console.log('[Toggle:' + seq + '] JSON.stringify() took ' + (performance.now() - tj).toFixed(0) + 'ms (' + Math.round(tabState.markedSfdt.length / 1024) + 'KB)');

                var to = performance.now();
                de.open(tabState.markedSfdt);
                console.log('[Toggle:' + seq + '] de.open(marked) took ' + (performance.now() - to).toFixed(0) + 'ms');


            } else {
                // ── HIDE TAGS ──
                if (tabState._editedInTagView) {
                    // User edited while tags were visible (e.g. removed a CC)
                    // Can't use stale cleanSfdt — serialize current state and strip markers
                    console.log('[Toggle:' + seq + '] Document was edited in tag-view — stripping markers from current state');
                    var ts2 = performance.now();
                    var currentSfdt = de.serialize();
                    console.log('[Toggle:' + seq + '] serialize(current) took ' + (performance.now() - ts2).toFixed(0) + 'ms');

                    var parsed = JSON.parse(currentSfdt);
                    _stripTagMarkers(parsed);
                    var stripped = JSON.stringify(parsed);

                    var to3 = performance.now();
                    de.open(stripped);
                    console.log('[Toggle:' + seq + '] de.open(stripped) took ' + (performance.now() - to3).toFixed(0) + 'ms (' + Math.round(stripped.length / 1024) + 'KB)');
                    tabState._editedInTagView = false;
                } else {
                    // No edits — safe to restore the original clean SFDT
                    var stored = tabState.cleanSfdt;
                    if (stored) {
                        var to2 = performance.now();
                        de.open(stored);
                        console.log('[Toggle:' + seq + '] de.open(clean) took ' + (performance.now() - to2).toFixed(0) + 'ms (' + Math.round(stored.length / 1024) + 'KB)');
                    } else {
                        console.warn('[Toggle:' + seq + '] No clean SFDT stored — nothing to restore');
                        _ccToggleBusy = false;
                        _showToggleOverlay(capturedDivId, false);
                        return;
                    }
                }
                // Clear ALL cache — next SHOW will serialize fresh
                _invalidateTabCache(tabState);
            }

            de.focusIn();
            _setTabTagsVisible(visible);

            console.log('[Toggle:' + seq + '] ── Total: ' + (performance.now() - t0).toFixed(0) + 'ms ──');

            // Restore Developer tab after de.open() — pass frozen references
            _restoreDevTab(capturedTabId, capturedDivId, seq);

        } catch (e) {
            _ccToggleBusy = false;
            _showToggleOverlay(capturedDivId, false);
            console.error('[Toggle:' + seq + '] Failed:', e);
        }
    });
}

// Lightweight loading overlay during toggle — targets the canvas/viewer area only
function _showToggleOverlay(divId, show) {
    var edDiv = document.getElementById(divId);
    if (!edDiv) return;

    // Find the document canvas area (not the full editor with toolbar)
    var host = edDiv.querySelector('[id$="_editor_viewerContainer"]')
            || edDiv.querySelector('.e-de-ctn')
            || edDiv;

    var overlay = host.querySelector('.dev-toggle-overlay');
    if (show && !overlay) {
        // Ensure host is positioned for absolute overlay
        if (getComputedStyle(host).position === 'static') {
            host.style.position = 'relative';
        }
        overlay = document.createElement('div');
        overlay.className = 'dev-toggle-overlay';
        host.appendChild(overlay);
    }
    if (overlay) {
        overlay.style.display = show ? 'flex' : 'none';
    }
}

function _restoreDevTab(capturedTabId, capturedDivId, seq) {
    // After de.open() the ribbon resets — click Developer tab to re-inject toggle button
    // Uses CAPTURED references (not globals) to avoid operating on the wrong tab
    setTimeout(function () {
        // Stale check
        if (seq !== _toggleSeq) {
            console.warn('[Toggle:' + seq + '] restoreDevTab aborted — superseded');
            _ccToggleBusy = false;
            _showToggleOverlay(capturedDivId, false);
            return;
        }
        var editorDiv = document.getElementById(capturedDivId);
        if (!editorDiv) {
            console.warn('[Toggle:' + seq + '] restoreDevTab — editor div gone');
            _ccToggleBusy = false;
            _showToggleOverlay(capturedDivId, false);
            return;
        }

        // Click the Developer tab header to make Syncfusion render its content
        var tabs = editorDiv.querySelectorAll('.e-tab-header .e-toolbar-item');
        var devTabFound = false;
        tabs.forEach(function (t) {
            if ((t.textContent || '').trim() === 'Developer') {
                var wrap = t.querySelector('.e-tab-wrap');
                if (wrap) { wrap.click(); devTabFound = true; }
            }
        });

        if (!devTabFound) {
            console.warn('[Toggle:' + seq + '] Developer tab header not found');
        }

        // Wait for Syncfusion to render the tab content, then update toggle button
        setTimeout(function () {
            var btn = editorDiv.querySelector('.dev-cc-toggle');
            var entry = _editors[capturedTabId];
            var vis = entry ? entry.ccTagsVisible : false;
            _updateToggleBtn(btn, vis);
            _ccToggleBusy = false;
            _showToggleOverlay(capturedDivId, false);
            console.log('[Toggle:' + seq + '] Dev tab restored for ' + capturedTabId.substring(0, 8) + ', busy lock released');
        }, 250);
    }, 100);
}

// Strip tag marker text runs from SFDT (reverse of _insertTagMarkers)
// Identifies markers by matching the _MARKER_CF signature or zero-width prefix
function _stripTagMarkers(sfdt) {
    var stripped = 0;
    var visited = new WeakSet();

    function isMarkerRun(inline) {
        if (!inline || typeof inline !== 'object') return false;
        // Check for zero-width prefix (new format)
        var text = inline.tlp;
        if (text && typeof text === 'string' && text.charAt(0) === _MARKER_PREFIX) return true;
        // Fallback: match by CF signature (supports both old and new marker formats)
        var cf = inline.cf;
        if (!cf) return false;
        // Current format (dark gray text, yellow highlight)
        if (cf.fsz === _MARKER_CF.fsz && cf.fc === _MARKER_CF.fc && cf.hc === _MARKER_CF.hc) return true;
        // Legacy: blue text, red highlight (fc:#2B79C2, hc:6, fsz:7)
        if (cf.fsz === 7 && cf.fc === '#2B79C2' && cf.hc === 6) return true;
        // Legacy: white invisible markers (fc:#FFFFFF, hc:0, fsz:6.5)
        if (cf.fsz === 6.5 && cf.fc === '#FFFFFF' && cf.hc === 0) return true;
        // Legacy format (fc:#444444, hc:14, fsz:7)
        if (cf.fsz === 7 && cf.fc === '#444444' && cf.hc === 14) return true;
        return false;
    }

    function isSpacerRun(inline) {
        // Spacer: { tlp: ' ', cf: { bi: false } } — single space, only bi:false in cf
        if (!inline || typeof inline !== 'object') return false;
        var text = inline.tlp;
        if (text !== ' ') return false;
        var cf = inline.cf;
        if (!cf) return false;
        var keys = Object.keys(cf);
        return keys.length === 1 && cf.bi === false;
    }

    function cleanInlines(arr) {
        if (!Array.isArray(arr)) return;
        for (var i = arr.length - 1; i >= 0; i--) {
            if (isMarkerRun(arr[i])) {
                arr.splice(i, 1);
                stripped++;
            } else if (isSpacerRun(arr[i])) {
                // Only remove spacer if adjacent to a marker (already removed) or at boundary
                // Check if next or prev was a marker position
                var prevIsMarker = (i > 0 && isMarkerRun(arr[i - 1]));
                var nextIsMarker = (i < arr.length - 1 && isMarkerRun(arr[i + 1]));
                var atStart = (i === 0);
                var atEnd = (i === arr.length - 1);
                if (prevIsMarker || nextIsMarker || atStart || atEnd) {
                    arr.splice(i, 1);
                    stripped++;
                }
            }
        }
    }

    function walk(obj) {
        if (!obj || typeof obj !== 'object' || visited.has(obj)) return;
        visited.add(obj);
        if (Array.isArray(obj)) {
            // Check if this array contains inline text runs (has objects with tlp)
            var hasInlines = obj.length > 0 && obj.some(function (item) {
                return item && typeof item === 'object' && ('tlp' in item);
            });
            if (hasInlines) cleanInlines(obj);
            for (var i = 0; i < obj.length; i++) walk(obj[i]);
            return;
        }
        var keys = Object.keys(obj);
        for (var k = 0; k < keys.length; k++) {
            if (obj[keys[k]] && typeof obj[keys[k]] === 'object') {
                walk(obj[keys[k]]);
            }
        }
    }

    walk(sfdt);
    console.log('[Toggle] Stripped ' + stripped + ' marker/spacer runs');
    return stripped;
}

// Walk SFDT and insert [tag] / [tag] marker text runs into each CC
function _insertTagMarkers(sfdt) {
    var count = 0;
    var visited = new WeakSet();

    function walk(obj) {
        if (!obj || typeof obj !== 'object' || visited.has(obj)) return;
        visited.add(obj);
        if (Array.isArray(obj)) {
            for (var i = 0; i < obj.length; i++) walk(obj[i]);
            return;
        }

        var ccp = obj.ccp || obj.contentControlProperties;
        if (ccp && typeof ccp === 'object' && !Array.isArray(ccp)) {
            var tag = ccp.tg || ccp.tag || ccp.tt || ccp.title || '';
            if (tag) {
                // Log first CC's properties for diagnostics
                if (count === 0) {
                    console.log('[Toggle] First CC props:', JSON.stringify(ccp).substring(0, 200));
                }
                // Set BoundingBox + color
                ccp.a = 1;
                ccp.c = _CC_COLOR;
                // Unlock CC so Remove/Edit operations work in tag-view mode
                ccp.lcc = false;  // lockContentControl = false
                ccp.lc = false;   // lockContents = false

                // Shorten tag name for display (e.g. "#book-part;id=2;type=chapter;" → "book-part")
                var shortTag = _shortTagName(tag);

                // Find the inlines array to insert markers into
                var inlines = _findFirstInlines(obj);
                var lastInlines = _findLastInlines(obj);

                if (inlines) {
                    // Prepend opening pill marker:  ❮tagname❯  (compact, monospace)
                    inlines.unshift(
                        { tlp: _MARKER_PREFIX + '\u276E' + shortTag + '\u276F', cf: _MARKER_CF_OPEN },
                        { tlp: ' ', cf: { "bi": false } }  // spacer after pill
                    );
                    count++;
                }
                if (lastInlines) {
                    // Append closing pill marker:  ❮tagname❯
                    lastInlines.push(
                        { tlp: ' ', cf: { "bi": false } },  // spacer before pill
                        { tlp: _MARKER_PREFIX + '\u276E' + shortTag + '\u276F', cf: _MARKER_CF_CLOSE }
                    );
                }
            }
        }

        // Recurse into child properties
        var keys = Object.keys(obj);
        for (var k = 0; k < keys.length; k++) {
            if (obj[keys[k]] && typeof obj[keys[k]] === 'object') {
                walk(obj[keys[k]]);
            }
        }
    }

    walk(sfdt);
    return count;
}

// Find the first inlines array (first paragraph's inlines) inside a CC
function _findFirstInlines(ccObj) {
    // Inline CC: has direct .i or .inlines
    if (Array.isArray(ccObj.i)) return ccObj.i;
    if (Array.isArray(ccObj.inlines)) return ccObj.inlines;

    // Block CC: has .b or .blocks containing paragraphs
    var blocks = ccObj.b || ccObj.blocks;
    if (Array.isArray(blocks)) {
        for (var i = 0; i < blocks.length; i++) {
            var inl = blocks[i].i || blocks[i].inlines;
            if (Array.isArray(inl)) return inl;
        }
    }
    return null;
}

// Find the last inlines array (last paragraph's inlines) inside a CC
function _findLastInlines(ccObj) {
    if (Array.isArray(ccObj.i)) return ccObj.i;
    if (Array.isArray(ccObj.inlines)) return ccObj.inlines;

    var blocks = ccObj.b || ccObj.blocks;
    if (Array.isArray(blocks)) {
        for (var i = blocks.length - 1; i >= 0; i--) {
            var inl = blocks[i].i || blocks[i].inlines;
            if (Array.isArray(inl)) return inl;
        }
    }
    return null;
}

// Expose for external use
window.toggleCCTagVisibility = function () {
    _applyTagVisibility(!_ccTagsVisible);
    return _ccTagsVisible;
};

window.setCCTagVisibility = function (visible) {
    _applyTagVisibility(visible);
    return _ccTagsVisible;
};

// Generic SFDT walker — calls callback(ccp) for every CC properties object
function _walkSfdtCCProps(sfdt, callback) {
    var count = 0, visited = new WeakSet();
    function walk(obj) {
        if (!obj || typeof obj !== 'object' || visited.has(obj)) return;
        visited.add(obj);
        if (Array.isArray(obj)) { for (var i = 0; i < obj.length; i++) walk(obj[i]); return; }
        var ccp = obj.ccp || obj.contentControlProperties;
        if (ccp && typeof ccp === 'object' && !Array.isArray(ccp)) { callback(ccp); count++; }
        var keys = Object.keys(obj);
        for (var k = 0; k < keys.length; k++) {
            if (obj[keys[k]] && typeof obj[keys[k]] === 'object') walk(obj[keys[k]]);
        }
    }
    walk(sfdt);
    return count;
}

// ══════════════════════════════════════════════════════════════════════════════
// CONTENT CONTROL — Insert CC wrapping current selection with tag & title
// ══════════════════════════════════════════════════════════════════════════════
// ══════════════════════════════════════════════════════════════════════════════
// CONTENT CONTROL — Remove CC via SFDT unwrap (Syncfusion removeContentControl is broken in v32)
// Uses selection offset to locate the exact CC at cursor, then unwraps only that one.
// ══════════════════════════════════════════════════════════════════════════════

// Flatten a blocks array (handling nested block-level CCs) into paragraphs.
// Returns array of { block, ccParentArr, ccIdx } where ccParentArr/ccIdx point
// to the innermost block-level CC containing this paragraph (or null if none).
function _flattenBlocks(blocks, ccParentArr, ccIdx) {
    var result = [];
    for (var i = 0; i < blocks.length; i++) {
        var b = blocks[i];
        var ccp = b.ccp || b.contentControlProperties;
        var innerBlocks = b.b || b.blocks;
        if (ccp && innerBlocks) {
            // Block-level CC — recurse; inner paragraphs belong to THIS CC
            var inner = _flattenBlocks(innerBlocks, blocks, i);
            for (var j = 0; j < inner.length; j++) result.push(inner[j]);
        } else {
            // Regular paragraph
            result.push({ block: b, ccParentArr: ccParentArr || null, ccIdx: ccIdx !== undefined ? ccIdx : -1 });
        }
    }
    return result;
}

// Count the total character offset consumed by an inline element in Syncfusion's offset scheme.
// Inline CC: 1 (start marker) + sum of inner text lengths + 1 (end marker)
// Text run: text length
// Other (bookmark, field, image, etc.): varies
function _inlineCharLen(inline) {
    var ccp = inline.ccp || inline.contentControlProperties;
    var innerInlines = inline.i || inline.inlines;
    if (ccp && innerInlines) {
        // Inline CC: 1 + inner text + 1
        var innerLen = 0;
        for (var j = 0; j < innerInlines.length; j++) innerLen += _inlineCharLen(innerInlines[j]);
        return 1 + innerLen + 1;
    }
    if (inline.tlp !== undefined) return inline.tlp.length;  // text run
    if (inline.text !== undefined) return inline.text.length; // full-name text
    if (inline.img || inline.image) return 1;                 // image
    if (inline.fld || inline.fieldType !== undefined) return 0; // field marker (0-width)
    if (inline.n || inline.bookmarkName) return 0;            // bookmark (0-width)
    return 0; // unknown — safe default
}

// Find the inline CC at a given character offset within a paragraph's inlines array.
// Returns { parentArr, idx } pointing to the CC node, or null.
function _findInlineCCAtOffset(inlines, charOffset) {
    if (!inlines) return null;
    var pos = 0;
    for (var i = 0; i < inlines.length; i++) {
        var il = inlines[i];
        var len = _inlineCharLen(il);
        var ccp = il.ccp || il.contentControlProperties;
        var innerInlines = il.i || il.inlines;
        if (ccp && innerInlines && charOffset >= pos && charOffset <= pos + len) {
            // Cursor is inside this CC — check for nested inline CCs
            // Skip the start marker (1 char) and recurse into inner inlines
            var nested = _findInlineCCAtOffset(innerInlines, charOffset - pos - 1);
            if (nested) return nested; // found a deeper CC
            return { parentArr: inlines, idx: i };
        }
        pos += len;
    }
    return null;
}

// Main removal function — called from context menu override and programmatically
window.removeContentControlAtCursor = function () {
    if (!container) { console.warn('[CC-Remove] No active editor'); return false; }
    var de = container.documentEditor;
    var tabState = _getTabState();

    try {
        // 1. Get cursor position BEFORE any serialization
        var offsetStr = de.selection.startOffset;
        var parts = offsetStr.split(';').map(Number);
        console.log('[CC-Remove] Cursor offset: ' + offsetStr);

        // 2. Serialize current document state
        var wasInTagView = tabState && tabState.ccTagsVisible;
        var sfdtStr;
        if (wasInTagView && tabState.cleanSfdt) {
            // Use clean SFDT (without tag markers) — offsets may differ due to markers,
            // so we re-serialize from current state and strip markers instead
            sfdtStr = de.serialize();
        } else {
            sfdtStr = de.serialize();
        }
        var sfdt = JSON.parse(sfdtStr);

        // 3. Strip tag markers if in tag-view (they add extra chars that shift offsets)
        if (wasInTagView) {
            // Tag markers change offsets, so we can't reliably use the offset approach.
            // Instead, use a simpler heuristic: find the CC closest to the cursor.
            // Strip markers, count CCs, and use offset to approximate which one.
            _stripTagMarkers(sfdt);
            console.log('[CC-Remove] Stripped tag markers — using approximate CC matching');
        }

        // 4. Parse offset and find the CC
        if (parts.length < 3) {
            console.warn('[CC-Remove] Unexpected offset format: ' + offsetStr);
            return false;
        }
        var secIdx = parts[0], blockIdx = parts[1], charOffset = parts[2];

        var section = sfdt.sec ? sfdt.sec[secIdx] : null;
        if (!section || !section.b) {
            console.warn('[CC-Remove] Section ' + secIdx + ' not found');
            return false;
        }

        // Flatten blocks to find the paragraph at blockIdx and any enclosing block-level CC
        var flat = _flattenBlocks(section.b, null, undefined);
        if (blockIdx >= flat.length) {
            console.warn('[CC-Remove] Block index ' + blockIdx + ' out of range (max ' + flat.length + ')');
            return false;
        }

        var paraInfo = flat[blockIdx];
        var targetCC = null;
        var ccTag = '(unknown)';

        // Check for inline CC at the character offset
        var inlineResult = _findInlineCCAtOffset(paraInfo.block.i || paraInfo.block.inlines, charOffset);
        if (inlineResult) {
            targetCC = inlineResult;
            var ccNode = targetCC.parentArr[targetCC.idx];
            ccTag = (ccNode.ccp || ccNode.contentControlProperties || {}).tg || ccTag;
            console.log('[CC-Remove] Found inline CC at offset (tag: ' + ccTag + ')');
        } else if (paraInfo.ccParentArr) {
            // Cursor is inside a block-level CC
            targetCC = { parentArr: paraInfo.ccParentArr, idx: paraInfo.ccIdx };
            var ccNode = targetCC.parentArr[targetCC.idx];
            ccTag = (ccNode.ccp || ccNode.contentControlProperties || {}).tg || ccTag;
            console.log('[CC-Remove] Found block-level CC (tag: ' + ccTag + ')');
        }

        if (!targetCC) {
            console.log('[CC-Remove] No content control at cursor position');
            return false;
        }

        // 5. Unwrap only that one CC
        var ccNode = targetCC.parentArr[targetCC.idx];
        var inner = ccNode.b || ccNode.blocks || ccNode.i || ccNode.inlines;
        Array.prototype.splice.apply(targetCC.parentArr, [targetCC.idx, 1].concat(inner));

        // 6. Re-open the document
        de.open(JSON.stringify(sfdt));
        de.focusIn();

        // 7. Invalidate caches since document structure changed
        if (tabState) {
            _invalidateTabCache(tabState);
            if (wasInTagView) {
                _setTabTagsVisible(false);
                var edDiv = document.getElementById(tabState.divId);
                if (edDiv) {
                    var btn = edDiv.querySelector('.dev-cc-toggle');
                    if (btn) _updateToggleBtn(btn, false);
                }
            }
        }

        console.log('[CC-Remove] Removed 1 content control (tag: ' + ccTag + ') via SFDT unwrap');
        return true;
    } catch (e) {
        console.error('[CC-Remove] Failed:', e);
        return false;
    }
};

// Monkey-patch Syncfusion's broken removeContentControl on each editor instance
function _patchRemoveCC(inst) {
    var editor = inst.documentEditor.editor;
    if (!editor || editor._ccRemovePatched) return;

    var origRemove = editor.removeContentControl;
    editor.removeContentControl = function () {
        console.log('[CC-Remove] Intercepted removeContentControl — using SFDT unwrap');
        window.removeContentControlAtCursor();
    };
    editor._ccRemovePatched = true;
    console.log('[CC-Remove] Patched removeContentControl on editor instance');
}

window.insertContentControlWithTag = function (tag, title) {
    if (!container) { console.warn('[CC] No active editor'); return false; }
    var de = container.documentEditor;
    var editor = de.editor;
    var sel = de.selection;
    if (!editor || !sel) { console.warn('[CC] Editor/selection unavailable'); return false; }

    var tabState = _getTabState();

    try {
        // If tag markers are active, first restore clean document
        if (tabState && tabState.ccTagsVisible && tabState.cleanSfdt) {
            console.log('[CC] Restoring clean doc before CC insert (was in tag-view mode)');
            de.open(tabState.cleanSfdt);
            _invalidateTabCache(tabState);
            _setTabTagsVisible(false);
            // Note: selection/cursor position is lost — user will need to re-select
        }

        // Insert a Rich Text content control around the current selection
        editor.insertContentControl('RichText');

        // Syncfusion API doesn't persist tag/title set after insert.
        // Patch them directly in the SFDT and re-open.
        var sfdt = JSON.parse(de.serialize());
        var patched = false;
        _walkSfdtCCProps(sfdt, function (ccp) {
            if (!ccp.tg && !ccp.tag) {
                ccp.tg = tag;
                ccp.tt = title || tag;
                ccp.c = _CC_COLOR;
                patched = true;
            }
        });
        if (patched) {
            de.open(JSON.stringify(sfdt));
        }
        de.focusIn();
        // Invalidate marker cache since document structure changed
        _invalidateTabCache(tabState);
        console.log('[CC] Inserted content control: tag=' + tag + ', title=' + (title || tag));
        return true;
    } catch (e) {
        console.error('[CC] Failed to insert content control:', e);
        return false;
    }
};

// ══════════════════════════════════════════════════════════════════════════════
// AI MANUSCRIPT PROCESSOR — Load comparison result with Track Changes
// ══════════════════════════════════════════════════════════════════════════════
window.loadManuscriptResult = function (sfdt) {
    if (!container || !sfdt) return;
    showLoader('Loading AI suggestions…');
    try {
        container.documentEditor.open(sfdt);
        container.documentEditor.enableTrackChanges = true;
        container.documentEditor.showRevisions = true;
        container.documentEditor.focusIn();
        console.log('[Manuscript] Loaded comparison document with track changes');

        // Open the track changes review pane after document renders
        setTimeout(function () {
            try {
                // Method 1: Syncfusion API — showComments pane or review pane
                if (typeof container.showRevisions !== 'undefined') {
                    container.showRevisions = true;
                }

                // Method 2: Click the Review tab, then find & click "Track Changes" button
                var editorDiv = _editors[_activeEditorTabId] ? document.getElementById(_editors[_activeEditorTabId].divId) : null;
                var root = editorDiv || document;

                // Click Review tab
                var tabHeaders = root.querySelectorAll('.e-ribbon .e-tab-header .e-toolbar-item');
                for (var i = 0; i < tabHeaders.length; i++) {
                    if (tabHeaders[i].textContent.trim() === 'Review') {
                        tabHeaders[i].click();
                        break;
                    }
                }

                // After Review tab activates, find and click the Changes/Track Changes button
                setTimeout(function () {
                    var allBtns = root.querySelectorAll('.e-ribbon-item .e-btn, .e-ribbon-item .e-dropdown-btn, .e-ribbon-item .e-split-btn, .e-ribbon-item button');
                    for (var j = 0; j < allBtns.length; j++) {
                        var txt = (allBtns[j].textContent || '').trim();
                        // Syncfusion's Review tab has "Changes" or "Track Changes" button
                        if (txt === 'Changes' || txt === 'Track Changes') {
                            allBtns[j].click();
                            console.log('[Manuscript] Opened Changes pane via "' + txt + '" button');
                            break;
                        }
                    }
                }, 400);
            } catch (e2) {
                console.log('[Manuscript] Could not auto-open review pane:', e2.message);
            }
        }, 500);
    } catch (e) {
        console.error('[Manuscript] Failed to load comparison:', e);
    }
    hideLoader();
};