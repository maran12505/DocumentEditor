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
        // Syncfusion's Find/Replace/Navigation pane caches its flex-child widths at
        // open time and .resize() alone won't reshrink it. When the editor area
        // changes width (e.g. AI sidebar toggled), close + reopen forces a relayout.
        _refreshOptionsPane();
        // Refresh the active editor's ribbon layout
        _refreshActiveRibbon();
        return;
    }
    if (attemptsLeft <= 0) return;
    requestAnimationFrame(function () { setTimeout(function () { safeResize(attemptsLeft - 1); }, 50); });
}

function _refreshOptionsPane() {
    try {
        var mod = container && container.documentEditor && container.documentEditor.optionsPaneModule;
        if (!mod || !mod.isOptionsPaneShow) return;

        // Capture the active sub-tab (Heading / Find / Replace) before the close/reopen
        // dance below. Syncfusion's showHideOptionsPane(true) reopens to Find by default,
        // which would silently clobber Heading/Replace whenever this runs (ribbon tab
        // click, window resize, AI-sidebar toggle all funnel through safeResize → here).
        var savedIdx = null;
        try {
            if (mod.tabInstance && typeof mod.tabInstance.selectedItem === 'number') {
                savedIdx = mod.tabInstance.selectedItem;
            }
        } catch (e) { /* tabInstance may not be ready; fall through */ }

        mod.showHideOptionsPane(false);
        mod.showHideOptionsPane(true);
        container.resize();

        // Restore the sub-tab if Syncfusion reopened to a different one.
        if (savedIdx !== null) {
            try {
                var cur = (mod.tabInstance && typeof mod.tabInstance.selectedItem === 'number')
                    ? mod.tabInstance.selectedItem : null;
                if (cur !== savedIdx && mod.tabInstance && typeof mod.tabInstance.select === 'function') {
                    mod.tabInstance.select(savedIdx);
                }
            } catch (e) { /* non-fatal */ }
        }
    } catch (e) {}
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
    var sfdt = _serializeForSave();
    if (!sfdt || sfdt.length < 10) return false;
    return await window.tabDbSave(tabId, sfdt);
};

// Load tab from IndexedDB into the ACTIVE editor (used for recovery/restore)
window.loadTabFromDb = async function (tabId) {
    showLoader('Loading…');
    var sfdt = await window.tabDbLoad(tabId);
    if (sfdt && sfdt.length > 10) {
        var entry = _editors[_activeEditorTabId] || _editors[tabId];
        if (entry) { entry.originalSfdt = sfdt; entry._loadingDocument = true; }
        var cleanSfdt = _sanitizeSfdtForOpen(sfdt);
        container.documentEditor.open(cleanSfdt);
        _postOpenFieldStrip();
        container.documentEditor.focusIn();
        if (entry) {
            entry.editedSinceOpen = false;
            setTimeout(function () { entry._loadingDocument = false; }, 100);
        }
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
        autoHeadingsEnabled: false,
        ccTagsVisible: false,
        cleanSfdt: null,    // original SFDT string (no markers)
        markedSfdt: null,   // SFDT string with tag markers inserted (cached)
        _editedInTagView: false,  // true if user edited while tags were visible
        // Zotero / Word field-instruction round-trip state (see _sanitizeSfdtForOpen)
        originalSfdt: null,         // pristine SFDT including ft:0…ft:2 instruction runs
        editedSinceOpen: false,     // flipped true by contentChange once user edits
        _loadingDocument: false     // guard: suppress contentChange during programmatic open()
    };
    _editors[tabId] = editorEntry;

    // Track edits. Ignore the synthetic contentChange fired by open(); only real
    // user edits should mark the document dirty for save-time round-trip decisions.
    inst.documentEditor.contentChange = function () {
        if (editorEntry._loadingDocument) return;
        editorEntry.editedSinceOpen = true;
        // Toggle caches: any real edit invalidates them so the next toggle rebuilds
        // fresh. Covers both edits-while-tags-visible AND edits-while-tags-hidden.
        editorEntry.cleanSfdt = null;
        editorEntry.markedSfdt = null;
        if (editorEntry.ccTagsVisible) {
            editorEntry._editedInTagView = true;
        }
    };

    // ── iPubEdit Properties wiring ──
    // 1) selection sync — populate panel when cursor moves
    var _origSelChange = inst.documentEditor.selectionChange;
    inst.documentEditor.selectionChange = function () {
        try { if (typeof _origSelChange === 'function') _origSelChange.apply(this, arguments); } catch (e) {}
        if (window.IPubProperties && typeof window.IPubProperties.onSelectionChanged === 'function') {
            window.IPubProperties.onSelectionChanged();
        }
    };
    // 2) document loaded — fire auto-open exactly once
    var _origDocChange = inst.documentEditor.documentChange;
    inst.documentEditor.documentChange = function () {
        try { if (typeof _origDocChange === 'function') _origDocChange.apply(this, arguments); } catch (e) {}
        if (window.IPubProperties && typeof window.IPubProperties.onDocumentLoaded === 'function') {
            setTimeout(function () { window.IPubProperties.onDocumentLoaded(); }, 0);
        }
    };

    // Patch Syncfusion's broken removeContentControl with SFDT-based unwrap
    _patchRemoveCC(inst);

    // Make this the active editor
    _setActiveEditor(tabId);

    // Attach slash command listener to this editor's canvas
    _attachSlashToEditor(divId);

    // Inject unified header ASAP via MutationObserver (no 500ms flash)
    _watchForRibbon(divId, inst);

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
    // Sync search button to this tab's nav pane state
    var navOpen = entry.navPaneOpen || false;
    document.querySelectorAll('.n-rbtn-search').forEach(function (b) { b.classList.toggle('active', navOpen); });

    var _switchDelay = window.__serverBaseUrl ? 150 : 30;
    setTimeout(function () {
        try { container.resize(); } catch (e) {}
        try { container.documentEditor.focusIn(); } catch (e) {}
        // Also focus the viewer container for WebView2 scroll
        var edDiv = document.getElementById(entry.divId);
        if (edDiv) {
            var vc = edDiv.querySelector('[id$="_editor_viewerContainer"]');
            if (vc) try { vc.focus(); } catch (e) {}
            var btn = edDiv.querySelector('.dev-cc-toggle');
            if (btn) _updateToggleBtn(btn, _ccTagsVisible);
        }
    }, _switchDelay);
    // MAUI/WebView2 safety net: second resize + focusIn after layout fully settles
    if (window.__serverBaseUrl) {
        setTimeout(function () {
            try { container.resize(); } catch (e) {}
            try { container.documentEditor.focusIn(); } catch (e) {}
        }, 200);
    }
    if (window.initSelectionToolbar) {
        setTimeout(window.initSelectionToolbar, 100);
    }
    // Re-setup auto-headings if nav pane is open for this tab
    if (entry.navPaneOpen) {
        // Early sync to prevent blink (before full setup at 200ms)
        setTimeout(function () { var r = _getActiveEditorRoot(); if (r) _ahSyncVisNow(r); }, 50);
        setTimeout(function () { _cleanNavigationHeadings(); _setupAutoHeadings(); }, 200);
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

    // ── MAUI/WebView2: Refocus viewer on wheel to fix scroll after focus loss ──
    if (window.__serverBaseUrl) {
        var _wheelArea = document.querySelector('#editorArea');
        if (_wheelArea) {
            var _lastWheelFocus = 0;
            _wheelArea.addEventListener('wheel', function (e) {
                var now = Date.now();
                if (now - _lastWheelFocus < 200) return;
                var ae = document.activeElement;
                if (ae && (ae.tagName === 'INPUT' || ae.tagName === 'TEXTAREA' || ae.tagName === 'SELECT' || ae.isContentEditable)) return;
                if (!container || !container.documentEditor || !_activeEditorTabId) return;
                var entry = _editors[_activeEditorTabId];
                if (!entry) return;
                var edDiv = document.getElementById(entry.divId);
                var vc = edDiv && edDiv.querySelector('[id$="_editor_viewerContainer"]');
                if (vc && vc.contains(e.target)) {
                    _lastWheelFocus = now;
                    try { container.documentEditor.focusIn(); } catch (ex) {}
                }
            }, { passive: true });
        }
    }

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

// ── Update customXml (iCore metadata) inside the bound DOCX file ─────────────
// Called from Blazor after iPubEdit Meta dialog Save. Reads the original DOCX
// bytes via the File System Access handle, POSTs to /update-customxml with the
// metadata DTO, and writes the patched DOCX bytes back through the same handle.
// Returns one of: 'OK', 'NO_HANDLE', 'PERMISSION_DENIED', 'DELETED', 'ERROR:<msg>'.
window.updateCustomXmlForActiveTab = async function (tabId, dtoJson) {
    console.log('[UpdateMeta-JS] ===== START tab=' + tabId + ' =====');
    if (!_fileHandles || !_fileHandles[tabId]) {
        console.warn('[UpdateMeta-JS] No file handle for tab — cannot persist to disk.');
        console.log('[UpdateMeta-JS] ===== END (no handle) =====');
        return 'NO_HANDLE';
    }
    var handle = _fileHandles[tabId];
    try {
        // Upgrade to readwrite (no-op if already granted).
        var perm = await handle.queryPermission({ mode: 'readwrite' });
        if (perm !== 'granted') {
            console.log('[UpdateMeta-JS] Requesting readwrite permission...');
            perm = await handle.requestPermission({ mode: 'readwrite' });
            if (perm !== 'granted') {
                console.warn('[UpdateMeta-JS] Permission denied.');
                console.log('[UpdateMeta-JS] ===== END (permission denied) =====');
                return 'PERMISSION_DENIED';
            }
        }

        var file;
        try { file = await handle.getFile(); }
        catch (fileErr) {
            if (fileErr.name === 'NotFoundError') {
                console.warn('[UpdateMeta-JS] Source file deleted on disk.');
                console.log('[UpdateMeta-JS] ===== END (deleted) =====');
                return 'DELETED';
            }
            throw fileErr;
        }
        console.log('[UpdateMeta-JS] read source file: name=' + file.name + ' size=' + file.size);

        // Build multipart payload: file + dto JSON
        var form = new FormData();
        form.append('files', file, file.name || 'document.docx');
        form.append('dto', dtoJson || '{}');

        console.log('[UpdateMeta-JS] POST /api/documenteditor/update-customxml');
        var response = await fetch(_apiUrl('/api/documenteditor/update-customxml'), { method: 'POST', body: form });
        console.log('[UpdateMeta-JS] response status = ' + response.status + ' ' + response.statusText);
        if (!response.ok) {
            var errText = await response.text();
            console.warn('[UpdateMeta-JS] non-success body: ' + errText);
            console.log('[UpdateMeta-JS] ===== END (server error) =====');
            return 'ERROR:' + response.status;
        }

        var blob = await response.blob();
        var arr = new Uint8Array(await blob.arrayBuffer());
        console.log('[UpdateMeta-JS] received patched DOCX: ' + arr.length + ' bytes. Writing back to handle...');

        var writable = await handle.createWritable();
        await writable.write(arr);
        await writable.close();
        console.log('[UpdateMeta-JS] DOCX written back to disk OK.');
        console.log('[UpdateMeta-JS] ===== END =====');
        return 'OK';
    } catch (e) {
        console.warn('[UpdateMeta-JS] EXCEPTION:', e && e.message);
        if (e.name === 'NotAllowedError' || e.name === 'SecurityError') return 'PERMISSION_DENIED';
        if (e.name === 'NotFoundError') return 'DELETED';
        console.log('[UpdateMeta-JS] ===== END (exception) =====');
        return 'ERROR:' + (e && e.message);
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
        var sfdt = _serializeForSave();
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
            mode: 'readwrite',
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

        // ── customXml -> iPubEdit Meta auto-fill (best-effort, .docx only) ──
        if (fileExt === '.docx' && window._blazorEditorRef) {
            (async function () {
                console.log('[ExtractMeta-JS] ===== START =====');
                try {
                    var metaForm = new FormData();
                    metaForm.append('files', file);
                    console.log('[ExtractMeta-JS] POST /api/documenteditor/extract-metadata, file=' + (file.name || 'document.docx') + ' size=' + file.size);
                    var metaResp = await fetch(_apiUrl('/api/documenteditor/extract-metadata'), { method: 'POST', body: metaForm });
                    console.log('[ExtractMeta-JS] response status = ' + metaResp.status + ' ' + metaResp.statusText);
                    if (metaResp.status === 204) {
                        console.log('[ExtractMeta-JS] 204 NoContent — no customXml metadata found.');
                        console.log('[ExtractMeta-JS] ===== END =====');
                        return;
                    }
                    if (!metaResp.ok) {
                        var errText = await metaResp.text();
                        console.warn('[ExtractMeta-JS] non-success body: ' + errText);
                        console.log('[ExtractMeta-JS] ===== END =====');
                        return;
                    }
                    var dto = await metaResp.json();
                    console.log('[ExtractMeta-JS] DTO received:', dto);
                    try {
                        await window._blazorEditorRef.invokeMethodAsync('OnIPubMetaExtractedFromJS', dto);
                        console.log('[ExtractMeta-JS] OnIPubMetaExtractedFromJS invoked OK.');
                    } catch (invokeErr) {
                        console.warn('[ExtractMeta-JS] invokeMethodAsync failed:', invokeErr && invokeErr.message);
                    }
                } catch (metaErr) {
                    console.warn('[ExtractMeta-JS] EXCEPTION:', metaErr && metaErr.message);
                }
                console.log('[ExtractMeta-JS] ===== END =====');
            })();
        } else {
            console.log('[ExtractMeta-JS] Skipped (ext=' + fileExt + ', blazorRef=' + (!!window._blazorEditorRef) + ')');
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
    // NOTE: Syncfusion paints the page onto a <canvas>, and its pageBackgroundColor
    // property is not reactive in this build — setting it does not repaint the
    // canvas, and a serialize/open round-trip still paints RGBA(255,255,255,255).
    // Document "paper" stays white in dark mode until Syncfusion exposes a repaint
    // hook (or we swap to their built-in theme API). Chrome around the page still
    // darkens correctly via the [data-theme="dark"] CSS rules.
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
        // Clean tag markers from navigation headings after pane renders, then setup auto-headings
        setTimeout(function () { _cleanNavigationHeadings(); _setupNavCleanObserver(); _setupAutoHeadings(); }, 300);
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
    // Deactivate search button in unified toolbar
    document.querySelectorAll('.n-rbtn-search').forEach(function (b) { b.classList.remove('active'); });
    if (tabId === _activeEditorTabId && window._blazorEditorRef) {
        try { window._blazorEditorRef.invokeMethodAsync('OnNavPaneClosed'); }
        catch (e) { console.warn('[NavPane] Could not notify Blazor:', e); }
    }
}

// Called after any user-facing de.open() to sync nav pane state
function _afterDocumentOpen() {
    document.querySelectorAll('.n-rbtn-search').forEach(function (b) { b.classList.remove('active'); });
    if (_activeEditorTabId && _editors[_activeEditorTabId]) {
        _editors[_activeEditorTabId].navPaneOpen = false;
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
    _afterDocumentOpen();
};
window.loadDocument = function (sfdt) {
    if (!container) return Promise.resolve();
    console.log('[Render] Opening document (' + Math.round((sfdt || '').length / 1024) + 'KB)...');

    // Stash pristine SFDT for round-trip save (preserves Zotero field instructions
    // that _sanitizeSfdtForOpen strips below). See _stripFieldInstructions comment.
    var entry = _editors[_activeEditorTabId];
    if (entry) { entry.originalSfdt = sfdt; entry._loadingDocument = true; }
    var cleanSfdt = _sanitizeSfdtForOpen(sfdt);

    // Yield TWO frames so the browser can paint the loading spinner
    // before the synchronous open() blocks the main thread
    return new Promise(function (resolve) {
        requestAnimationFrame(function () {
            requestAnimationFrame(function () {
                var start = performance.now();
                try { container.documentEditor.open(cleanSfdt); }
                catch (e) { console.warn('[Render] open() threw (non-fatal):', e && e.message); }
                console.log('[Render] open() took ' + (performance.now() - start).toFixed(0) + 'ms');
                try { _postOpenFieldStrip(); } catch (e) { console.warn('[Render] _postOpenFieldStrip:', e && e.message); }

                // Syncfusion 32.2.3 bug: open() returns synchronously but `processSfdt`
                // continues in an internal Promise chain. Calling focusIn() / accessing
                // selection here races against bodyWidget attachment. We:
                //   1. SKIP focusIn() entirely — the user's first click focuses anyway,
                //      and the caret-blink animation isn't worth a NullRef cascade.
                //   2. POLL for bodyWidget before resolving so callers (HydrateMetaFromCustomXml,
                //      applyMetadata) only run when the selection module is safe to touch.
                var de = container.documentEditor;
                var pollStart = performance.now();
                (function waitForBody() {
                    var ready = false;
                    try {
                        ready = !!(de && de.documentHelper && de.documentHelper.pages
                                   && de.documentHelper.pages.length > 0
                                   && de.documentHelper.pages[0].bodyWidgets
                                   && de.documentHelper.pages[0].bodyWidgets.length > 0);
                    } catch (e) { ready = false; }

                    if (ready || (performance.now() - pollStart) > 4000) {
                        try { _afterDocumentOpen(); } catch (e) { console.warn('[Render] _afterDocumentOpen:', e && e.message); }
                        if (entry) {
                            entry.editedSinceOpen = false;
                            setTimeout(function () { entry._loadingDocument = false; }, 100);
                        }
                        resolve();
                    } else {
                        setTimeout(waitForBody, 50);
                    }
                })();
            });
        });
    });
};

// ── Suppress 3 known-harmless Syncfusion 32.2.3 console errors ─────────────
// These fire from internal Promise/setTimeout callbacks during SFDT processing
// and selection-init — try/catch around open() cannot catch them. They do NOT
// affect document rendering. We match by message substring + source file so we
// don't accidentally swallow real app errors.
(function installSyncfusionNoiseFilter() {
    if (window._sfNoiseFilterInstalled) return;
    window._sfNoiseFilterInstalled = true;
    var noisySnippets = [
        "reading 'bodyWidget'",
        "Can't find local header signature",
        "reading 'length'"
    ];
    function isNoise(msg, src) {
        if (!msg) return false;
        var m = String(msg);
        if (!noisySnippets.some(function (s) { return m.indexOf(s) !== -1; })) return false;
        // Only filter when the source is Syncfusion's bundle.
        return !src || String(src).indexOf('ej2.min.js') !== -1 || String(src).indexOf('ej2.') !== -1;
    }
    window.addEventListener('error', function (e) {
        if (isNoise(e.message, e.filename)) { e.preventDefault(); e.stopImmediatePropagation(); return false; }
    }, true);
    window.addEventListener('unhandledrejection', function (e) {
        var r = e.reason;
        var msg = r && (r.message || r.toString && r.toString());
        var src = r && r.stack;
        if (isNoise(msg, src)) { e.preventDefault(); }
    });
})();
window.getDocumentContent = function () {
    if (!container) return null;
    var ts = _getTabState();

    // Round-trip fast path: if the document was opened via the field-instruction
    // sanitizer and the user hasn't edited anything since, return the pristine
    // original SFDT. This keeps Zotero / Word field instructions intact so the
    // exported DOCX still contains ADDIN ZOTERO_ITEM / PAGEREF / HYPERLINK field
    // codes and downstream plugins can refresh citations.
    // Exception: if Show Tags is active, we fall through to the tag-handling
    // branches below (tag state trumps round-trip).
    if (ts && !ts.ccTagsVisible && ts.originalSfdt && !ts.editedSinceOpen) {
        console.log('[Save] Returning pristine original SFDT (fields preserved, ' + Math.round(ts.originalSfdt.length / 1024) + 'KB)');
        return ts.originalSfdt;
    }

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
    } else if (ts && ts.originalSfdt && !ts.editedSinceOpen) {
        // Unmodified doc — return the pristine SFDT (keeps Zotero field instructions)
        oldSfdt = ts.originalSfdt;
    } else {
        oldSfdt = container.documentEditor.serialize();
    }
    // Clear cache on switch — will be rebuilt when user toggles again
    if (ts) { ts.cleanSfdt = null; ts.markedSfdt = null; }

    // Stash pristine SFDT for the INCOMING document's round-trip and sanitize for render.
    if (ts) {
        ts.originalSfdt = isBlank ? null : newSfdt;
        ts._loadingDocument = true;
    }
    if (isBlank) container.documentEditor.openBlank();
    else {
        container.documentEditor.open(_sanitizeSfdtForOpen(newSfdt));
        _postOpenFieldStrip();
    }
    container.documentEditor.focusIn();
    if (ts) {
        ts.editedSinceOpen = false;
        setTimeout(function () { ts._loadingDocument = false; }, 100);
    }
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
// AUTO-DETECTED HEADINGS (MS Word Online-style)
// ══════════════════════════════════════════════════════════════════════════════
//
// Each editor tab has its OWN nav pane DOM. _setupAutoHeadings is called
// per-tab and stores observers in _editors[tabId] so they're independent.
// The treeDiv gets class .n-heading-panel for CSS scrollbar styling.
// A MutationObserver on treeDiv's style attribute handles FIND/REPLACE hiding.
// A MutationObserver on opTab's childList keeps the footer in correct position.

var _ahRetries = 0;
var _ahToggleHTML =
    '<label class="n-toggle-switch">' +
        '<input type="checkbox" class="n-auto-headings-cb">' +
        '<span class="n-toggle-slider"></span>' +
    '</label>' +
    '<span class="n-auto-headings-label">Automatically detected headings</span>' +
    '<span class="n-auto-headings-info" title="Detect headings by formatting when the document has no heading styles">' +
        '<svg width="14" height="14" viewBox="0 0 2048 2048" fill="currentColor"><path d="M1152 640q0 53-37 91-37 37-91 37-53 0-91-37-37-37-37-91 0-53 37-91 37-37 91-37 53 0 91 37 37 37 37 91z m-256 384q0-53 37-91 37-37 91-37 53 0 91 37 37 37 37 91v384q0 53-37 91-37 37-91 37-53 0-91-37-37-37-37-91v-384z m-768 0q0-243 122-452 119-203 322-322 209-122 452-122 243 0 452 122 203 119 322 322 122 209 122 452 0 243-122 452-119 203-322 322-209 122-452 122-243 0-452-122-203-119-322-322-122-209-122-452z m896-768q-209 0-388 105-174 102-275 275-105 179-105 388 0 209 105 388 102 174 275 275 179 105 388 105 209 0 388-105 174-102 275-275 105-179 105-388 0-209-105-388-102-174-275-275-179-105-388-105z"/></svg>' +
    '</span>';

// ── Ensure correct DOM order: searchBar → treeDiv → [auto-tree] → footer
function _ahEnsureOrder(opTab) {
    var td = opTab.querySelector('[id$="_editor_treeDiv"]');
    var footer = opTab.querySelector('.n-auto-headings-footer');
    if (!td || !footer) return;
    var searchBar = opTab.querySelector('.n-heading-search-bar');
    var autoTree = opTab.querySelector('.n-auto-headings-tree');
    // searchBar before treeDiv
    if (searchBar && searchBar.nextSibling !== td) td.parentNode.insertBefore(searchBar, td);
    if (autoTree) {
        // treeDiv → autoTree → footer
        if (td.nextSibling !== autoTree) td.parentNode.insertBefore(autoTree, td.nextSibling);
        if (autoTree.nextSibling !== footer) autoTree.parentNode.insertBefore(footer, autoTree.nextSibling);
    } else {
        // treeDiv → footer
        if (td.nextSibling !== footer) td.parentNode.insertBefore(footer, td.nextSibling);
    }
}

// ── Global polling: sync visibility of footer/auto-tree in ACTIVE editor ──
// Replaces all MutationObservers. Runs every 200ms, checks the ACTIVE editor
// only — works across tab switches without re-attaching observers.
if (!window._ahVisInterval) {
    window._ahVisInterval = setInterval(function () {  // 100ms polling
        if (!_activeEditorTabId || !_editors[_activeEditorTabId]) return;
        var root = document.getElementById(_editors[_activeEditorTabId].divId);
        if (!root) return;
        var td = root.querySelector('[id$="_editor_treeDiv"]');
        if (!td) return;
        var footer = root.querySelector('.n-auto-headings-footer');
        var autoTree = root.querySelector('.n-auto-headings-tree');
        // Syncfusion sets treeDiv display:none when FIND/REPLACE is active.
        // When auto-headings toggle is ON, we also set display:none + data-auto-hidden.
        // Heading tab is active when: treeDiv is NOT hidden by Syncfusion.
        var searchContent = root.querySelector('.e-de-search-tab-content');
        var isHeadingTab = !searchContent || searchContent.style.display === 'none';
        var searchBar = root.querySelector('.n-heading-search-bar');
        if (searchBar) searchBar.style.display = isHeadingTab ? '' : 'none';
        if (footer) footer.style.display = isHeadingTab ? '' : 'none';
        if (autoTree) autoTree.style.display = isHeadingTab ? '' : 'none';
        // Syncfusion's built-in heading treeDiv isn't hidden when Replace is
        // active (only when Find is), so chapters leak into the Replace view.
        // Force-hide it whenever we're not on the Heading tab.
        if (!isHeadingTab) {
            td.style.display = 'none';
        } else if (!td.hasAttribute('data-auto-hidden')) {
            // Restore on Heading tab unless auto-headings toggle hid it explicitly
            td.style.display = '';
        } else {
            td.style.display = 'none';
        }
    }, 100);
}

// Helper to get active editor root div
function _getActiveEditorRoot() {
    if (!_activeEditorTabId || !_editors[_activeEditorTabId]) return null;
    return document.getElementById(_editors[_activeEditorTabId].divId);
}

// Immediate sync for use after setup/tab-switch (doesn't wait for interval)
function _ahSyncVisNow(root) {
    if (!root) return;
    var td = root.querySelector('[id$="_editor_treeDiv"]');
    if (!td) return;
    var footer = root.querySelector('.n-auto-headings-footer');
    var autoTree = root.querySelector('.n-auto-headings-tree');
    var searchBar = root.querySelector('.n-heading-search-bar');
    var searchContent = root.querySelector('.e-de-search-tab-content');
    var isHeadingTab = !searchContent || searchContent.style.display === 'none';
    if (searchBar) searchBar.style.display = isHeadingTab ? '' : 'none';
    if (footer) footer.style.display = isHeadingTab ? '' : 'none';
    if (autoTree) autoTree.style.display = isHeadingTab ? '' : 'none';
    if (isHeadingTab && td.hasAttribute('data-auto-hidden')) td.style.display = 'none';
}

// ── Apply flex layout (Syncfusion may reset display:block on tab switch) ──
function _ahApplyLayout(opDiv, opTab, treeDiv) {
    opDiv.style.display = 'flex';
    opDiv.style.flexDirection = 'column';
    opDiv.style.height = '100%';
    opDiv.style.overflow = 'hidden';
    opTab.style.flex = '1';
    opTab.style.display = 'flex';
    opTab.style.flexDirection = 'column';
    opTab.style.overflow = 'hidden';
    opTab.style.minHeight = '0';
    treeDiv.style.flex = '1';
    treeDiv.style.overflowY = 'auto';
    treeDiv.style.minHeight = '0';
    treeDiv.style.height = '';
    if (!treeDiv.classList.contains('n-heading-panel')) treeDiv.classList.add('n-heading-panel');
}

// ── Main setup: called per editor tab, fully idempotent ──────────────
function _setupAutoHeadings() {
    if (!_activeEditorTabId || !_editors[_activeEditorTabId]) return;
    var entry = _editors[_activeEditorTabId];
    var root = document.getElementById(entry.divId);
    if (!root) return;
    var opDiv = root.querySelector('.e-de-op');
    if (!opDiv) return;
    var opTab = opDiv.querySelector('.e-de-op-tab');
    if (!opTab) return;
    var treeDiv = opTab.querySelector('[id$="_editor_treeDiv"]');
    if (!treeDiv) {
        if (_ahRetries++ < 15) setTimeout(_setupAutoHeadings, 100);
        return;
    }
    _ahRetries = 0;

    // ── Always re-apply flex layout (Syncfusion resets on tab switch) ──
    _ahApplyLayout(opDiv, opTab, treeDiv);

    // ── Wire click on tab header for immediate sync (once per opTab) ──
    var tabHeader = opTab.querySelector('.e-tab-header');
    if (tabHeader && !tabHeader._ahClickWired) {
        tabHeader._ahClickWired = true;
        tabHeader.addEventListener('click', function () {
            requestAnimationFrame(function () {
                var r = _getActiveEditorRoot();
                if (r) _ahSyncVisNow(r);
            });
        });
    }

    // ── If footer already exists, just sync state + ensure order ──
    var existingFooter = opTab.querySelector('.n-auto-headings-footer');
    if (existingFooter) {
        _ahEnsureOrder(opTab);
        // Sync toggle checkbox
        var cb = existingFooter.querySelector('.n-auto-headings-cb');
        if (cb) {
            cb.checked = !!entry.autoHeadingsEnabled;
            if (entry.autoHeadingsEnabled) _showAutoHeadings(root);
            else _hideAutoHeadings(root);
        }
        _ahSyncVisNow(root);
        return;
    }

    // ── First-time setup: create search bar + footer + attach observers ──

    // Search bar
    var searchBar = document.createElement('div');
    searchBar.className = 'n-heading-search-bar';
    searchBar.innerHTML =
        '<span class="n-hs-icon"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg></span>' +
        '<input type="text" class="n-hs-input" placeholder="Search headings...">' +
        '<button class="n-hs-clear" title="Clear"><svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg></button>';
    treeDiv.parentNode.insertBefore(searchBar, treeDiv);

    var hsInput = searchBar.querySelector('.n-hs-input');
    var hsClear = searchBar.querySelector('.n-hs-clear');
    hsInput.addEventListener('input', function () {
        hsClear.classList.toggle('visible', hsInput.value.length > 0);
        _filterHeadingItems(opTab, hsInput.value);
    });
    hsClear.addEventListener('click', function () {
        hsInput.value = '';
        hsClear.classList.remove('visible');
        _filterHeadingItems(opTab, '');
    });
    // Prevent Syncfusion from stealing keyboard input
    hsInput.addEventListener('keydown', function (e) { e.stopPropagation(); });

    var footer = document.createElement('div');
    footer.className = 'n-auto-headings-footer';
    footer.innerHTML = _ahToggleHTML;
    treeDiv.parentNode.insertBefore(footer, treeDiv.nextSibling);
    _ahEnsureOrder(opTab);

    // Toggle event
    footer.querySelector('.n-auto-headings-cb').addEventListener('change', function () {
        var en = this.checked;
        if (_activeEditorTabId && _editors[_activeEditorTabId]) {
            _editors[_activeEditorTabId].autoHeadingsEnabled = en;
        }
        // root may be stale; re-fetch
        var r = document.getElementById(_editors[_activeEditorTabId].divId);
        if (en) _showAutoHeadings(r);
        else _hideAutoHeadings(r);
    });

    // ChildList observer: Syncfusion may re-add treeDiv → reorder
    var orderObs = new MutationObserver(function () {
        var td = opTab.querySelector('[id$="_editor_treeDiv"]');
        if (td && !td.classList.contains('n-heading-panel')) {
            _ahApplyLayout(opDiv, opTab, td);
        }
        _ahEnsureOrder(opTab);
    });
    orderObs.observe(opTab, { childList: true });

    // Sync initial state
    var cb2 = footer.querySelector('.n-auto-headings-cb');
    if (cb2) {
        cb2.checked = !!entry.autoHeadingsEnabled;
        if (entry.autoHeadingsEnabled) _showAutoHeadings(root);
    }
    _ahSyncVisNow(root);
}

// ── Sync toggle + content (convenience wrapper) ─────────────────────
function _syncAutoHeadingsState(root) {
    var cb = root.querySelector('.n-auto-headings-cb');
    if (!cb) return;
    var enabled = _activeEditorTabId && _editors[_activeEditorTabId]
        ? !!_editors[_activeEditorTabId].autoHeadingsEnabled : false;
    cb.checked = enabled;
    if (enabled) _showAutoHeadings(root);
    else _hideAutoHeadings(root);
}

// ── Show auto-detected headings ──────────────────────────────────────
function _showAutoHeadings(root) {
    if (!root) return;
    var treeDiv = root.querySelector('[id$="_editor_treeDiv"]');
    if (!treeDiv) return;
    var opTab = treeDiv.parentNode;

    // Clear search filter on rebuild
    var hsInput = opTab.querySelector('.n-heading-search-bar .n-hs-input');
    var hsClear = opTab.querySelector('.n-heading-search-bar .n-hs-clear');
    if (hsInput) hsInput.value = '';
    if (hsClear) hsClear.classList.remove('visible');

    var old = opTab.querySelector('.n-auto-headings-tree');
    if (old) old.remove();

    var headings = _detectAutoHeadings();

    // If nothing detected (e.g. doc already has Heading N styles, so _detectAutoHeadings
    // skips every block), fall back to the native Syncfusion heading list instead of
    // hiding it and showing an empty-state. The toggle stays ON (user's preference
    // preserved); it just becomes a no-op for styled docs.
    if (headings.length === 0) {
        treeDiv.removeAttribute('data-auto-hidden');
        treeDiv.style.display = '';
        return;
    }

    treeDiv.setAttribute('data-auto-hidden', 'true');
    treeDiv.style.display = 'none';

    var tree = document.createElement('div');
    tree.className = 'n-auto-headings-tree n-heading-panel';
    tree.style.flex = '1'; tree.style.overflowY = 'auto'; tree.style.minHeight = '0';
    for (var i = 0; i < headings.length; i++) {
        var h = headings[i];
        var item = document.createElement('div');
        item.className = 'n-auto-heading-item';
        item.setAttribute('data-level', h.level);
        item.textContent = h.text;
        item.title = h.text;
        item.addEventListener('click', (function (t) { return function () { _navigateToAutoHeading(t); }; })(h.text));
        tree.appendChild(item);
    }
    var footer = opTab.querySelector('.n-auto-headings-footer');
    if (footer) opTab.insertBefore(tree, footer);
    else treeDiv.parentNode.insertBefore(tree, treeDiv.nextSibling);
    _ahEnsureOrder(opTab);
}

// ── Hide auto-detected headings ──────────────────────────────────────
function _hideAutoHeadings(root) {
    if (!root) return;
    var opTab = root.querySelector('.e-de-op-tab');
    if (!opTab) return;
    // Clear search filter
    var hsInput = opTab.querySelector('.n-heading-search-bar .n-hs-input');
    var hsClear = opTab.querySelector('.n-heading-search-bar .n-hs-clear');
    if (hsInput) hsInput.value = '';
    if (hsClear) hsClear.classList.remove('visible');
    var tree = opTab.querySelector('.n-auto-headings-tree');
    if (tree) tree.remove();
    var treeDiv = root.querySelector('[id$="_editor_treeDiv"]');
    if (treeDiv) { treeDiv.removeAttribute('data-auto-hidden'); treeDiv.style.display = ''; }
    // Reset filter on manual heading items
    _filterHeadingItems(opTab, '');
}

// ── Filter heading items by search query ────────────────────────────
function _filterHeadingItems(opTab, query) {
    if (!opTab) return;
    var q = (query || '').toLowerCase().trim();
    var shown = 0;

    // Filter Syncfusion manual heading items (.e-list-item inside treeDiv)
    var treeDiv = opTab.querySelector('[id$="_editor_treeDiv"]');
    if (treeDiv && treeDiv.style.display !== 'none') {
        var items = treeDiv.querySelectorAll('.e-list-item');
        for (var i = 0; i < items.length; i++) {
            var text = (items[i].textContent || '').toLowerCase();
            var match = !q || text.indexOf(q) !== -1;
            items[i].style.display = match ? '' : 'none';
            if (match) shown++;
        }
    }

    // Filter auto-detected heading items (.n-auto-heading-item)
    var autoTree = opTab.querySelector('.n-auto-headings-tree');
    if (autoTree) {
        var autoItems = autoTree.querySelectorAll('.n-auto-heading-item');
        for (var j = 0; j < autoItems.length; j++) {
            var aText = (autoItems[j].textContent || '').toLowerCase();
            var aMatch = !q || aText.indexOf(q) !== -1;
            autoItems[j].style.display = aMatch ? '' : 'none';
            if (aMatch) shown++;
        }
        // "No matching headings" message
        var noMatch = autoTree.querySelector('.n-heading-no-match');
        if (!noMatch) {
            noMatch = document.createElement('div');
            noMatch.className = 'n-heading-no-match';
            noMatch.textContent = 'No matching headings.';
            autoTree.appendChild(noMatch);
        }
        noMatch.style.display = (q && shown === 0) ? '' : 'none';
    }
}

// ── Heading detection (unchanged logic) ──────────────────────────────
function _detectAutoHeadings() {
    if (!container) return [];
    try {
        var sfdt = JSON.parse(container.documentEditor.serialize());
        var sections = sfdt.sections || sfdt.sec || [];
        var results = [];
        var fontSizes = {};

        // First pass: find the most common font size (= body font size)
        sections.forEach(function (sec) {
            var blocks = sec.blocks || sec.b || [];
            blocks.forEach(function (block) {
                _collectFontSizes(block, fontSizes);
            });
        });
        var bodyFontSize = 0;
        var maxCount = 0;
        for (var sz in fontSizes) {
            if (fontSizes[sz] > maxCount) { maxCount = fontSizes[sz]; bodyFontSize = parseFloat(sz); }
        }

        // Second pass: detect heading-like paragraphs
        var globalBlockIdx = 0;
        for (var si = 0; si < sections.length; si++) {
            var blocks = sections[si].blocks || sections[si].b || [];
            for (var bi = 0; bi < blocks.length; bi++) {
                var block = blocks[bi];
                var inlines = block.inlines || block.i;
                if (!inlines || inlines.length === 0) { globalBlockIdx++; continue; }

                // Skip if this block already has a heading style
                var pf = block.paragraphFormat || block.pf;
                if (pf) {
                    var sn = pf.styleName || pf.sn || '';
                    if (/^Heading \d/i.test(sn)) { globalBlockIdx++; continue; }
                }

                var text = '';
                var allBold = true;
                var maxFS = 0;
                var hasText = false;
                for (var ii = 0; ii < inlines.length; ii++) {
                    var inl = inlines[ii];
                    var t = inl.text !== undefined ? inl.text : (inl.tlp !== undefined ? inl.tlp : '');
                    if (!t) continue;
                    hasText = true;
                    text += t;
                    var cf = inl.characterFormat || inl.cf || block.characterFormat || block.cf || {};
                    if (!cf.bold && !cf.b) allBold = false;
                    var fs = cf.fontSize || cf.fsz || 0;
                    if (fs > maxFS) maxFS = fs;
                }
                text = text.trim();
                if (!hasText || !text || text.length > 120) { globalBlockIdx++; continue; }

                // Skip single-character or very short non-alpha paragraphs
                if (text.length < 3) { globalBlockIdx++; continue; }

                var level = _classifyHeading(text, allBold, maxFS, bodyFontSize);
                if (level > 0) {
                    results.push({ level: level, text: text, secIdx: si, blockIdx: bi, globalBlockIdx: globalBlockIdx });
                }
                globalBlockIdx++;
            }
        }
        return results;
    } catch (e) { console.warn('[AutoHeadings] detection error:', e); return []; }
}

function _collectFontSizes(block, sizes) {
    var inlines = block.inlines || block.i;
    if (inlines) {
        for (var i = 0; i < inlines.length; i++) {
            var t = inlines[i].text !== undefined ? inlines[i].text : (inlines[i].tlp !== undefined ? inlines[i].tlp : '');
            if (!t || t.trim().length === 0) continue;
            var cf = inlines[i].characterFormat || inlines[i].cf || block.characterFormat || block.cf || {};
            var fs = cf.fontSize || cf.fsz || 0;
            if (fs > 0) sizes[fs] = (sizes[fs] || 0) + t.length;
        }
    }
    var rows = block.rows || block.r;
    if (rows) {
        rows.forEach(function (row) {
            (row.cells || row.c || []).forEach(function (cell) {
                (cell.blocks || cell.b || []).forEach(function (b) { _collectFontSizes(b, sizes); });
            });
        });
    }
}

function _classifyHeading(text, allBold, maxFS, bodyFS) {
    var isAllCaps = text === text.toUpperCase() && /[A-Z]/.test(text);

    // Pattern: "Scheme 1", "Table 1", "Figure 1" → level 1
    if (/^(Scheme|Table|Figure)\s+\d/i.test(text)) return 1;

    // Pattern: numbered sub-section "2.1", "2.1.1" → level 3
    if (/^\d+\.\s?\d+(\.\d+)?\s/.test(text)) return 3;

    // Pattern: numbered section "1.", "2." → level 2
    if (/^\d+\.\s+[A-Z]/.test(text)) return 2;

    // ALL-CAPS short text → level 1
    if (isAllCaps && text.length < 100 && text.length >= 3) return 1;

    // Bold + larger than body → level 1
    if (allBold && bodyFS > 0 && maxFS > bodyFS && text.length < 100) return 1;

    // Bold + short paragraph (under 80 chars) → level 2
    if (allBold && text.length < 80 && text.length >= 3) return 2;

    return 0;
}

// Navigate to a detected heading by searching for its text.
// Matches the native Syncfusion heading list: cursor moves to the heading,
// no selection-highlight remains (search.find leaves the match selected,
// so we collapse it to a zero-width cursor at the paragraph start).
function _navigateToAutoHeading(headingText) {
    if (!container || !headingText) return;
    try {
        // Use a unique-enough snippet (first 60 chars) to find the heading
        var query = headingText.length > 60 ? headingText.substring(0, 60) : headingText;
        var sel = container.documentEditor.selection;
        // Move to document start so find() searches from the top
        sel.moveToDocumentStart();
        // find() selects the first match and scrolls it into view
        container.documentEditor.search.find(query, 'None');
        // Collapse the resulting selection so no highlight remains — the
        // viewport has already scrolled to the match, we just drop the cursor
        // at the start of the heading paragraph (native-list parity).
        try { sel.moveToParagraphStart(); } catch (e) { /* non-fatal */ }
    } catch (e) { console.warn('[AutoHeadings] navigate error:', e); }
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
        var sfdt = _serializeForSave();
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
    var entry = _editors[_activeEditorTabId];
    if (entry) { entry.originalSfdt = sfdt; entry._loadingDocument = true; }
    container.documentEditor.open(_sanitizeSfdtForOpen(sfdt));
    _postOpenFieldStrip();
    container.documentEditor.focusIn();
    _afterDocumentOpen();
    if (entry) {
        entry.editedSinceOpen = false;
        setTimeout(function () { entry._loadingDocument = false; }, 100);
    }
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
        var sfdt = _serializeForSave();
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

// Returns SFDT suitable for persisting (Save / Save As / autosave).
// Guarantees no tag-view markers are present even if tags are currently visible.
//   - tags hidden          → de.serialize() directly (current state IS clean)
//   - tags visible, no edit → return cached cleanSfdt (instant, no work)
//   - tags visible, edited → serialize current, strip markers, return stripped
// Without this, saving while tags are visible persists the markers as real text.
function _serializeForSave() {
    if (!container || !container.documentEditor) return null;
    var de = container.documentEditor;
    var s = _getTabState();
    if (!s || !s.ccTagsVisible) {
        return de.serialize();
    }
    if (s.cleanSfdt && !s._editedInTagView) {
        return s.cleanSfdt;
    }
    try {
        var live = de.serialize();
        var parsed = JSON.parse(live);
        _stripTagMarkers(parsed);
        return JSON.stringify(parsed);
    } catch (e) {
        console.warn('[Save] strip-on-save failed; falling back to raw serialize:', e);
        return de.serialize();
    }
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

// ─────────────────────────────────────────────────────────────────────────
// iPubEdit Meta — top-level Ribbon tab via container.ribbon.addTab
// (DocumentEditorContainer's ribbon WRAPPER — NOT the underlying EJ2 Ribbon
// at .e-ribbon.ej2_instances[0]; the wrapper has the contextualTabManager
// that re-registers contextual tab handlers after addTab).
// Reference: https://help.syncfusion.com/document-processing/word/word-processor/javascript-es5/how-to/customize-ribbon
// ─────────────────────────────────────────────────────────────────────────
function _addIPubEditMetaTab(container) {
    if (!container) return;
    var divId = (container.element && container.element.id) || '?';
    var tag = '[iPubEditMeta:' + divId.slice(0, 11) + ']';

    // container.ribbon is wired AFTER appendTo + first layout. Poll up to 30s.
    if (!container.ribbon || typeof container.ribbon.addTab !== 'function') {
        var attempts = (container.__ipubMetaAttempts = (container.__ipubMetaAttempts || 0) + 1);
        if (attempts > 300) {
            console.warn(tag + ' container.ribbon never became ready after ' + attempts + ' attempts (~30s)');
            return;
        }
        setTimeout(function () { _addIPubEditMetaTab(container); }, 100);
        return;
    }
    var ribbon = container.ribbon;

    // Per-container idempotency. NEVER guard via document.* — multiple editors
    // share the document; a global hit would block every later instance.
    try {
        var existing = ribbon.tabManager && ribbon.tabManager.tabCollection;
        if (existing) {
            for (var i = 0; i < existing.length; i++) {
                var h = existing[i] && (existing[i].header || (existing[i].tabHeader && existing[i].tabHeader.header));
                if (String(h || '').trim() === 'iPubEdit Meta') {
                    console.log(tag + ' already present, skip');
                    return;
                }
            }
        }
    } catch (e) { /* fall through, addTab will throw if truly duplicate */ }

    // Unique id per container so rendered DOM IDs don't collide across editors.
    var uid = (divId.replace(/[^a-zA-Z0-9]/g, '') || ('e' + Date.now()));

    var ribbonTab = {
        header: 'iPubEdit',
        id: 'ipub_edit_meta_tab_' + uid,
        groups: [{
            header: 'Meta',
            id: 'ipub_edit_meta_group_' + uid,
            collections: [{
                items: [
                    {
                        type: 'Button',
                        buttonSettings: {
                            content: 'Meta',
                            iconCss: 'e-icons e-edit',
                            clicked: function () {
                                var bridge = document.getElementById('nBtnMeta');
                                if (bridge) bridge.click();
                            }
                        }
                    },
                    {
                        type: 'Button',
                        buttonSettings: {
                            content: 'Properties',
                            iconCss: 'e-icons e-settings',
                            clicked: function () {
                                if (window.IPubProperties && typeof window.IPubProperties.toggle === 'function')
                                    window.IPubProperties.toggle();
                                else {
                                    var bridge = document.getElementById('nBtnProperties');
                                    if (bridge) bridge.click();
                                }
                            }
                        }
                    }
                ]
            }]
        }]
    };

    try {
        ribbon.addTab(ribbonTab);
        console.log(tag + ' tab added via container.ribbon.addTab');
    } catch (e) {
        console.warn(tag + ' container.ribbon.addTab failed:', e);
    }
}

function _updateToggleBtn(btn, visible) {
    if (!btn) return;
    var label = btn.querySelector('.dev-btn-label');
    var iconWrap = btn.querySelector('.dev-cc-icon');
    if (visible) {
        btn.classList.add('active');
        if (label) label.textContent = 'Tags';
        if (iconWrap) iconWrap.innerHTML =
            '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' +
                '<path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>';
    } else {
        btn.classList.remove('active');
        if (label) label.textContent = 'Tags';
        if (iconWrap) iconWrap.innerHTML =
            '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' +
                '<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"/>' +
                '<path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/>' +
                '<path d="M14.12 14.12a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>';
    }
}

// ══════════════════════════════════════════════════════════════════════════
// UNIFIED HEADER — Inject logo (left) + custom buttons (right) into ribbon
// ══════════════════════════════════════════════════════════════════════════

// Watch for the ribbon to appear and inject immediately (no setTimeout flash)
function _watchForRibbon(divId, container) {
    var editorDiv = document.getElementById(divId);
    if (!editorDiv) return;

    // Try immediately first (ribbon may already exist)
    var tabHeader = editorDiv.querySelector('.e-tab-header');
    if (tabHeader) {
        _injectDeveloperTab(editorDiv);
        _injectUnifiedHeader(editorDiv);
        _addIPubEditMetaTab(container);
        return;
    }

    // Otherwise watch for it via MutationObserver
    var observer = new MutationObserver(function () {
        var th = editorDiv.querySelector('.e-tab-header');
        if (th) {
            observer.disconnect();
            _injectDeveloperTab(editorDiv);
            _injectUnifiedHeader(editorDiv);
            _addIPubEditMetaTab(container);
        }
    });
    observer.observe(editorDiv, { childList: true, subtree: true });

    // Safety timeout — disconnect observer after 10s even if ribbon never appeared
    setTimeout(function () { observer.disconnect(); }, 10000);
}

function _injectUnifiedHeader(editorDiv) {
    if (!editorDiv) return;
    var tabHeader = editorDiv.querySelector('.e-tab-header');
    if (!tabHeader || editorDiv.querySelector('.n-ribbon-header-row')) return;

    var isDark = document.documentElement.getAttribute('data-theme') === 'dark';

    // ── Left group: logo image only (no text — logo includes the name) ──
    var left = document.createElement('div');
    left.className = 'n-ribbon-left';
    left.innerHTML =
        '<img src="_content/Deditor.Shared.UI/images/ProjectX_logo.png" alt="Logo" class="n-ribbon-logo-img" />';

    // ── Right group: search, focus, theme, save, AI, chat, avatar ──
    var right = document.createElement('div');
    right.className = 'n-ribbon-right';
    right.innerHTML =
        // Search
        '<button class="n-rbtn n-rbtn-icon n-rbtn-search" title="Find & Replace (Ctrl+F)">' +
            '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">' +
                '<circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/>' +
            '</svg>' +
        '</button>' +
        // Focus mode
        '<button class="n-rbtn n-rbtn-icon n-rbtn-focus" title="Focus Mode">' +
            '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5">' +
                '<polyline points="4 14 4 20 10 20"/><polyline points="20 10 20 4 14 4"/>' +
                '<line x1="14" y1="10" x2="20" y2="4"/><line x1="4" y1="20" x2="10" y2="14"/>' +
            '</svg>' +
        '</button>' +
        '<span class="n-ribbon-divider"></span>' +
        // Theme toggle
        '<button class="n-rbtn n-rbtn-icon n-rbtn-theme" title="Toggle Theme">' +
            (isDark
                ? '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">' +
                    '<circle cx="12" cy="12" r="5"/>' +
                    '<line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/>' +
                    '<line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/>' +
                    '<line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/>' +
                    '<line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>' +
                  '</svg>'
                : '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">' +
                    '<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>' +
                  '</svg>') +
        '</button>' +
        // Save group
        '<div class="n-ribbon-save-group">' +
            '<button class="n-rbtn n-rbtn-icon n-rbtn-save" title="Save (Ctrl+S)">' +
                '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                    '<path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>' +
                    '<polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>' +
                '</svg>' +
            '</button>' +
            '<span class="n-ribbon-divider"></span>' +
            '<button class="n-rbtn n-rbtn-icon n-rbtn-saveas" title="Save As (Ctrl+Shift+S)">' +
                '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                    '<path d="M13 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v4"/>' +
                    '<polyline points="7 3 7 8 15 8"/>' +
                    '<path d="M17 21.5l4.5-4.5-2-2L15 19.5V21.5h2z"/>' +
                '</svg>' +
            '</button>' +
        '</div>' +
        '<span class="n-ribbon-divider"></span>' +
        // AI Chat
        '<button class="n-rbtn n-rbtn-ai" title="Toggle AI Assistant">' +
            '<span class="n-ai-star">✦</span> AI' +
        '</button>' +
        // Avatar
        '<div class="n-ribbon-avatar">U</div>';

    // Build a flex wrapper row: [Logo] [File] [tab-header(tabs + right-buttons)]
    // This avoids making .e-ribbon-tab itself flex (which breaks the content area below)
    var ribbonTab = tabHeader.parentElement; // .e-ribbon-tab
    var fileBtn = ribbonTab ? ribbonTab.querySelector('.e-ribbon-file-menu') : null;
    var row = document.createElement('div');
    row.className = 'n-ribbon-header-row';
    row.appendChild(left);                 // 1. Logo
    if (fileBtn) row.appendChild(fileBtn); // 2. File button (moved from ribbon-tab)
    // 3. Tab header (moved from ribbon-tab into the row)
    ribbonTab.insertBefore(row, tabHeader);
    row.appendChild(tabHeader);
    // Right buttons go inside the tab header (flex: 1 area)
    tabHeader.appendChild(right);

    // ── Wire click forwarding to hidden Blazor buttons ──
    var fwd = function (cls, targetId) {
        var btn = tabHeader.querySelector('.' + cls);
        if (btn) btn.addEventListener('click', function (e) {
            e.stopPropagation();
            var t = document.getElementById(targetId);
            if (t) t.click();
        });
    };
    fwd('n-rbtn-search', 'nBtnSearch');
    fwd('n-rbtn-focus', 'nBtnFocusMode');
    fwd('n-rbtn-theme', 'nBtnTheme');
    fwd('n-rbtn-save', 'nBtnSave');
    fwd('n-rbtn-saveas', 'btnSaveAs');
    fwd('n-rbtn-ai', 'nBtnAiChat');

    // ── Update padding-top on .e-de-ctn + CSS variable for sidebars ──
    function _updateRibbonHeight() {
        var ctnrRibbon = editorDiv.querySelector('.e-de-ctnr-ribbon');
        var deCtn = editorDiv.querySelector('.e-de-ctn');
        if (ctnrRibbon && deCtn) {
            var h = ctnrRibbon.offsetHeight;
            deCtn.style.paddingTop = h + 'px';
            // Set CSS variable on :root so sidebars can use it for padding-top
            document.documentElement.style.setProperty('--n-ribbon-h', h + 'px');
        }
    }
    _updateRibbonHeight();

    // ── Watch for ribbon height changes (tab switch, collapse/expand) ──
    var ribbonEl = editorDiv.querySelector('.e-de-ctnr-ribbon');
    if (ribbonEl) {
        var _lastAppliedH = ribbonEl.offsetHeight;
        var _roTimer = null;
        var ribbonRO = new ResizeObserver(function () {
            clearTimeout(_roTimer);
            _roTimer = setTimeout(function () {
                requestAnimationFrame(function () {
                    var h = ribbonEl.offsetHeight;
                    if (h !== _lastAppliedH && h > 0) {
                        _lastAppliedH = h;
                        _updateRibbonHeight();
                        safeResize(3);
                    }
                });
            }, 100);
        });
        ribbonRO.observe(ribbonEl);
    }

    console.log('[UnifiedHeader] Injected into ribbon tab header');
}

// ── State sync: called from Blazor after render ──
window.syncRibbonHeader = function (state) {
    // Update ALL instances (one per editor tab)
    document.querySelectorAll('.n-rbtn-search').forEach(function (b) { b.classList.toggle('active', !!state.navPaneOpen); });
    document.querySelectorAll('.n-rbtn-ai').forEach(function (b) { b.classList.toggle('active', !!state.chatOpen); });

    // Theme icon + logo
    var isDark = !!state.isDark;
    var sunSvg = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">' +
                    '<circle cx="12" cy="12" r="5"/>' +
                    '<line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/>' +
                    '<line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/>' +
                    '<line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/>' +
                    '<line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/>' +
                 '</svg>';
    var moonSvg = '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2">' +
                    '<path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/>' +
                  '</svg>';
    var themeHtml = isDark ? sunSvg : moonSvg;
    var themeTitle = isDark ? 'Switch to Light Mode' : 'Switch to Dark Mode';
    document.querySelectorAll('.n-rbtn-theme').forEach(function (b) {
        b.title = themeTitle;
        b.innerHTML = themeHtml;
    });

};

// ── Design Mode: Insert/remove visible [tag] markers in document ────────
var _CC_COLOR = '#2383E2';

// ── Pill-style tag marker formatting ──
// Markers are invisible in SFDT (white text, no highlight) — visual pills are HTML overlays
// The text runs still occupy space for offset calculations and search-based positioning
var _MARKER_CF_OPEN  = { "fsz": 9, "fc": "#2383E2", "hc": "#d2e4faff", "fn": "Consolas", "b": true };
var _MARKER_CF_CLOSE = { "fsz": 9, "fc": "#2383E2", "hc": "#d2e4faff", "fn": "Consolas", "b": true };
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
                // Fast path: caches still valid (no edits since last toggle).
                // Skip serialize + parse + insertMarkers + stringify; reuse markedSfdt directly.
                // contentChange handler nulls these caches on any real user edit, so reaching
                // here guarantees the cached marked SFDT matches the current document.
                if (tabState.cleanSfdt && tabState.markedSfdt) {
                    var toFast = performance.now();
                    tabState._loadingDocument = true;
                    de.open(tabState.markedSfdt);
                    setTimeout(function () { tabState._loadingDocument = false; }, 100);
                    console.log('[Toggle:' + seq + '] de.open(cached marked) took ' + (performance.now() - toFast).toFixed(0) + 'ms (cache hit, ' + Math.round(tabState.markedSfdt.length / 1024) + 'KB)');
                    de.focusIn();
                    _setTabTagsVisible(true);
                    _markNavPaneClosed(capturedTabId);
                    console.log('[Toggle:' + seq + '] ── Total: ' + (performance.now() - t0).toFixed(0) + 'ms (cache hit) ──');
                    _restoreDevTab(capturedTabId, capturedDivId, seq);
                    return;
                }

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
                tabState._loadingDocument = true;
                de.open(tabState.markedSfdt);
                setTimeout(function () { tabState._loadingDocument = false; }, 100);
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
                    tabState._loadingDocument = true;
                    de.open(stripped);
                    setTimeout(function () { tabState._loadingDocument = false; }, 100);
                    console.log('[Toggle:' + seq + '] de.open(stripped) took ' + (performance.now() - to3).toFixed(0) + 'ms (' + Math.round(stripped.length / 1024) + 'KB)');
                    tabState._editedInTagView = false;
                    // Cache stripped as new clean baseline; markedSfdt stays null
                    // (will be rebuilt on next SHOW since structure changed).
                    tabState.cleanSfdt = stripped;
                    tabState.markedSfdt = null;
                } else {
                    // No edits — safe to restore the original clean SFDT
                    var stored = tabState.cleanSfdt;
                    if (stored) {
                        var to2 = performance.now();
                        tabState._loadingDocument = true;
                        de.open(stored);
                        setTimeout(function () { tabState._loadingDocument = false; }, 100);
                        console.log('[Toggle:' + seq + '] de.open(clean) took ' + (performance.now() - to2).toFixed(0) + 'ms (' + Math.round(stored.length / 1024) + 'KB)');
                    } else {
                        console.warn('[Toggle:' + seq + '] No clean SFDT stored — nothing to restore');
                        _ccToggleBusy = false;
                        _showToggleOverlay(capturedDivId, false);
                        return;
                    }
                }
                // Caches kept; next SHOW reuses markedSfdt via the fast path above.
                // Edits invalidate via contentChange handler.
            }

            de.focusIn();
            _setTabTagsVisible(visible);

            // de.open() closes the nav pane — sync search button
            _markNavPaneClosed(capturedTabId);

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
            // Clean tag markers from Navigation pane headings
            setTimeout(_cleanNavigationHeadings, 200);
            setTimeout(function () { _setupNavCleanObserver(); }, 500);
            console.log('[Toggle:' + seq + '] Dev tab restored for ' + capturedTabId.substring(0, 8) + ', busy lock released');
        }, 250);
    }, 100);
}

// ── Clean tag markers from Navigation pane headings ──
// Markers are emitted by _insertTagMarkers as \u200B + tag-text (no whitespace)
// followed by a spacer run of ' '. Syncfusion's heading extractor concatenates
// every inline's text, so marker+spacer leaks into .e-list-text. Strip both.
var _NAV_MARKER_RE = /\u200B\S*\s?/g;
var _ZWS_RE = /\u200B/g;

function _cleanNavigationHeadings() {
    // Find ALL navigation/options panes (one per editor tab)
    var panes = document.querySelectorAll('.e-documenteditor-optionspane');
    for (var p = 0; p < panes.length; p++) {
        var items = panes[p].querySelectorAll('.e-list-text');
        for (var i = 0; i < items.length; i++) {
            var text = items[i].textContent || '';
            if (_NAV_MARKER_RE.test(text) || _ZWS_RE.test(text)) {
                _NAV_MARKER_RE.lastIndex = 0;
                _ZWS_RE.lastIndex = 0;
                var cleaned = text.replace(_NAV_MARKER_RE, '').replace(_ZWS_RE, '').trim();
                items[i].textContent = cleaned;
                // Also clean the tooltip on the parent .e-list-item
                var listItem = items[i].closest('.e-list-item');
                if (listItem && listItem.title) {
                    _NAV_MARKER_RE.lastIndex = 0;
                    _ZWS_RE.lastIndex = 0;
                    listItem.title = listItem.title.replace(_NAV_MARKER_RE, '').replace(_ZWS_RE, '').trim();
                }
            }
            _NAV_MARKER_RE.lastIndex = 0;
            _ZWS_RE.lastIndex = 0;
        }
    }
}

function _setupNavCleanObserver() {
    // Attach a per-tab observer so each editor tab's pane is watched
    if (!_activeEditorTabId || !_editors[_activeEditorTabId]) return;
    var entry = _editors[_activeEditorTabId];
    if (entry._navCleanObserver) return; // already watching this tab
    var root = document.getElementById(entry.divId);
    if (!root) return;
    var target = root.querySelector('.e-documenteditor-optionspane');
    if (!target) return;
    entry._navCleanObserver = new MutationObserver(function () {
        _cleanNavigationHeadings();
    });
    entry._navCleanObserver.observe(target, { childList: true, subtree: true, characterData: true });
    _cleanNavigationHeadings(); // initial clean
}

// ────────────────────────────────────────────────────────────────────────────
// SFDT field-instruction sanitizer — Zotero / Word field rendering workaround
// ────────────────────────────────────────────────────────────────────────────
// Syncfusion's DocumentEditor v32.1.x renders the text between field-begin
// ({ft:0}) and field-separator ({ft:2}) — i.e. the field instruction code —
// as visible text instead of hiding it. In narrow two-column layouts this
// overflows and overlaps neighbouring runs (see ao5c04145.docx). Word, Word
// Online and LibreOffice all hide instruction runs; Syncfusion does not.
//
// Fix: before handing SFDT to documentEditor.open(), walk each inline array
// and remove runs sandwiched between {ft:0} and {ft:2}. Leave the ft:0, ft:2,
// and ft:1 markers themselves untouched so the field skeleton survives.
//
// Round-trip: the caller is expected to stash the original (pre-sanitize)
// SFDT on the editor entry as `originalSfdt`; on save, getDocumentContent()
// returns that pristine copy when the document hasn't been edited, which
// keeps Zotero field instructions intact for plugin round-trip. If the user
// edited the document, we fall back to the live (sanitized) serialize —
// in-session edits forfeit Zotero round-trip fidelity, by design.
// ────────────────────────────────────────────────────────────────────────────
function _stripFieldInstructions(sfdtObj) {
    var stripped = 0;
    var visited = new WeakSet();

    function cleanInlines(arr) {
        if (!Array.isArray(arr)) return;
        var i = 0;
        while (i < arr.length) {
            var item = arr[i];
            // Field begin marker: {ft:0} (often co-present with hfe, fpd, etc.)
            if (item && typeof item === 'object' && item.ft === 0) {
                // Look ahead for the matching separator (ft:2). Stop at field-end
                // (ft:1) too — a field with no separator has no instruction to strip.
                var j = i + 1;
                while (j < arr.length) {
                    var mid = arr[j];
                    if (mid && typeof mid === 'object' && (mid.ft === 2 || mid.ft === 1)) break;
                    j++;
                }
                if (j < arr.length && arr[j].ft === 2 && j > i + 1) {
                    // Splice out the instruction runs between begin and separator.
                    // Skeleton (ft:0 ... ft:2 ... result ... ft:1) stays intact.
                    var removed = j - i - 1;
                    arr.splice(i + 1, removed);
                    stripped += removed;
                }
            }
            i++;
        }
    }

    function walk(obj) {
        if (!obj || typeof obj !== 'object' || visited.has(obj)) return;
        visited.add(obj);
        if (Array.isArray(obj)) {
            var looksLikeInlines = obj.length > 0 && obj.some(function (item) {
                return item && typeof item === 'object' && ('tlp' in item || 'ft' in item);
            });
            if (looksLikeInlines) cleanInlines(obj);
            for (var i = 0; i < obj.length; i++) walk(obj[i]);
            return;
        }
        var keys = Object.keys(obj);
        for (var k = 0; k < keys.length; k++) {
            var v = obj[keys[k]];
            if (v && typeof v === 'object') walk(v);
        }
    }

    walk(sfdtObj);
    return stripped;
}

// Parse → strip → stringify. On any failure returns the input unchanged so
// the open() flow can't regress on malformed SFDT. No-op (returns original
// reference) when zero runs are stripped, to avoid a wasted stringify.
function _sanitizeSfdtForOpen(sfdt) {
    if (!sfdt) return sfdt;
    var isString = (typeof sfdt === 'string');
    // Fast-path: raw SFDT string with no {ft:0} literally — nothing to strip.
    // NB: Syncfusion's "optimized SFDT" envelope (`{"sfdt":"UEsDB..."}` — a base64
    // ZIP) also lacks a literal `"ft":0`, so that format falls into this fast
    // path too. The post-open sanitizer below catches it: once the editor has
    // decoded the envelope, `serialize()` yields plain JSON SFDT that DOES
    // contain the markers, and we strip + re-open on that expanded form.
    if (isString && sfdt.indexOf('"ft":0') === -1 && sfdt.indexOf('"ft": 0') === -1) {
        return sfdt;
    }
    var obj;
    try {
        obj = isString ? JSON.parse(sfdt) : sfdt;
    } catch (e) {
        console.warn('[Render] SFDT parse failed in field sanitizer, passing through:', e);
        return sfdt;
    }
    var start = performance.now();
    var n = _stripFieldInstructions(obj);
    if (n > 0) {
        console.log('[Render] Stripped ' + n + ' field-instruction run(s) in ' + (performance.now() - start).toFixed(0) + 'ms');
        return isString ? JSON.stringify(obj) : obj;
    }
    return sfdt;
}

// Post-open companion to _sanitizeSfdtForOpen. Covers the case where the input
// to open() was the optimized (base64-ZIP) SFDT envelope — the pre-open
// sanitizer can't see inline `"ft":0` markers inside the base64 blob, so it's a
// no-op. Once the editor has decoded + parsed the envelope, `serialize()`
// returns the expanded JSON SFDT where the markers are literal strings. We
// strip instruction runs on that expanded form and re-open to force the
// renderer to discard the offending internal model. Idempotent: a second call
// after a successful strip+reopen finds nothing further to strip and is a
// cheap no-op (indexOf returns > -1 but the stripper removes zero runs).
//
// Callers must run this while `_loadingDocument` is still true so the
// synthetic `contentChange` from the re-open doesn't flip `editedSinceOpen`.
function _postOpenFieldStrip() {
    if (!container || !container.documentEditor) return 0;
    var live;
    try { live = container.documentEditor.serialize(); }
    catch (e) { console.warn('[Render] post-open serialize failed:', e); return 0; }
    if (!live || live.indexOf('"ft":0') === -1) return 0;
    var obj;
    try { obj = JSON.parse(live); }
    catch (e) { console.warn('[Render] post-open SFDT parse failed:', e); return 0; }
    var start = performance.now();
    var n = _stripFieldInstructions(obj);
    if (n > 0) {
        var reopen = JSON.stringify(obj);
        console.log('[Render] Post-open stripped ' + n + ' field-instruction run(s) in ' + (performance.now() - start).toFixed(0) + 'ms; re-opening (' + Math.round(reopen.length / 1024) + 'KB)');
        var reopenStart = performance.now();
        container.documentEditor.open(reopen);
        console.log('[Render] post-open re-open() took ' + (performance.now() - reopenStart).toFixed(0) + 'ms');
    }
    return n;
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
        // Current format (blue text, no highlight)
        if (cf.fsz === _MARKER_CF.fsz && cf.fc === _MARKER_CF.fc && cf.hc === _MARKER_CF.hc) return true;
        // Legacy: blue text, turquoise highlight (fc:#2383E2, hc:3, fsz:7)
        if (cf.fsz === 7 && cf.fc === '#2383E2' && cf.hc === 3) return true;
        // Legacy: dark gray text, yellow highlight (fc:#4A4A4A, hc:1, fsz:7)
        if (cf.fsz === 7 && cf.fc === '#4A4A4A' && cf.hc === 1) return true;
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

                // Show opening tag BEFORE content and closing tag AFTER content.
                // e.g. tag="#book-part;id=2;type=chapter;" →
                //   open  = "#book-part;id=2;type=chapter;"
                //   close = "#/book-part;"
                // For plain tags (no leading '#'), close is "/<tag>" (e.g. BOOK_TITLE → /BOOK_TITLE).
                var openTag  = tag;
                var closeTag;
                if (tag.charAt(0) === '#') {
                    closeTag = '#/' + _shortTagName(tag) + ';';
                } else {
                    closeTag = '/' + tag;
                }

                // Opening marker on the first inlines array (same as before)
                var firstInlines = _findFirstInlines(obj);
                var lastInlines  = _findLastInlines(obj);

                if (firstInlines) {
                    firstInlines.unshift(
                        { tlp: _MARKER_PREFIX + openTag, cf: _MARKER_CF_OPEN },
                        { tlp: ' ', cf: { "bi": false } }  // spacer after opening pill
                    );
                    count++;
                }
                // Append closing marker to the last inlines array (handles block + inline CCs)
                if (lastInlines) {
                    lastInlines.push(
                        { tlp: ' ', cf: { "bi": false } }, // spacer before closing pill
                        { tlp: _MARKER_PREFIX + closeTag, cf: _MARKER_CF_CLOSE }
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
        var entry = _editors[_activeEditorTabId];
        if (entry) { entry.originalSfdt = sfdt; entry._loadingDocument = true; }
        container.documentEditor.open(_sanitizeSfdtForOpen(sfdt));
        _postOpenFieldStrip();
        if (entry) {
            entry.editedSinceOpen = false;
            setTimeout(function () { entry._loadingDocument = false; }, 100);
        }
        container.documentEditor.enableTrackChanges = true;
        container.documentEditor.showRevisions = true;
        container.documentEditor.focusIn();
        _afterDocumentOpen();
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
// ══════════════════════════════════════════════════════════════════════
// iPubEdit Meta Information — document-aware metadata persistence + editor binding
// ══════════════════════════════════════════════════════════════════════
var _IPUB_META_PREFIX = 'ipubedit.meta.';          // keyed by documentId (tabId)
var _IPUB_META_LEGACY = 'ipubedit.meta';           // old single-key blob (migrated on read)

// Normalize a CC tag for lookup: strip "#", drop ";…attributes", lowercase,
// dashes → underscores. "#book-part;id=2" → "book_part".
function _ipubNormalizeTag(rawTag) {
    if (!rawTag) return '';
    var s = _shortTagName(rawTag);      // strips leading "#" and ";..." attributes
    return s.toLowerCase().replace(/-/g, '_');
}

// Tag → metadata field resolver. Keys are post-_ipubNormalizeTag.
// Accepts both the user's UPPER_SNAKE style (BOOK_TITLE) and JATS kebab (book-title).
var _IPUB_TAG_FIELD = {
    'book_title':           'BookTitle',
    'isbn':                 'Isbn',
    'doi':                  'Doi',
    'publisher':            'Publisher',
    'publisher_name':       'Publisher',
    'publisher_imprint':    'PublisherImprint',
    'journal_name':         'JournalName',
    'journal_title':        'JournalTitle',
    'p_issn':               'PIssn',
    'pissn':                'PIssn',
    'issn':                 'PIssn',
    'e_issn':               'EIssn',
    'eissn':                'EIssn',
    'article_id':           'ArticleId',
    'article_type':         'ArticleType',
    'document_type':        'DocumentType',
    'year':                 'Year',
    'pub_year':             'Year',
    'month':                'Month',
    'pub_month':            'Month',
    'country':              'Country',
    'copyrights':           'Copyrights',
    'copyright_statement':  'Copyrights',
    'dtd':                  'Dtd',
    'job_card':             'JobCard',
    'ce_template':          'CeTemplate',
    'book_id':              'BookOrJournalId',
    'journal_id':           'BookOrJournalId',
    'pagination_platform':  'PaginationPlatform',
    'article_title':        'ArticleTitle',
    'chapter_title':        'ChapterTitle',
    'volume':               'Volume',
    'issue':                'Issue',
    'fpage':                'FirstPage',
    'first_page':           'FirstPage',
    'lpage':                'LastPage',
    'last_page':            'LastPage',
    'elocation_id':         'ElocationId',
    'edition':              'Edition',
    'series':               'Series',
    'day':                  'Day',
    'pub_day':              'Day',
    'publisher_loc':        'PublisherLoc'
};

// Pull a value from meta by field name, tolerating PascalCase or camelCase keys.
function _ipubGetMetaValue(meta, field) {
    if (!meta || !field) return '';
    if (meta[field] != null) return String(meta[field]);
    var lc = field.charAt(0).toLowerCase() + field.substring(1);
    if (meta[lc] != null) return String(meta[lc]);
    return '';
}

// Apply metadata to the active editor's content controls.
// Walks the serialized SFDT tree, matches each CC's tag to a metadata field,
// and rewrites the first inlines array inside the CC with the new value.
// Preserves the first non-marker run's char formatting (cf) so fonts survive.
// Returns { replaced, skipped, unmapped: [..up to 3 tag names..] }.
window.applyMetadata = function (meta) {
    if (!meta) return { replaced: 0, skipped: 0, unmapped: [] };
    // Expose XmlType for the iPubEdit Properties panel (chooses Book vs Journal tree)
    try {
        window.__ipubEditMeta = meta;
        if (meta.XmlType) localStorage.setItem('ipubedit.meta.xmlType', meta.XmlType);
    } catch (e) {}
    if (!container || !container.documentEditor) {
        console.warn('[iPubEdit] No active editor — skipping applyMetadata');
        return { replaced: 0, skipped: 0, unmapped: [] };
    }
    var de = container.documentEditor;
    var sfdtStr;
    try { sfdtStr = de.serialize(); }
    catch (e) { console.warn('[iPubEdit] serialize() failed', e); return { replaced: 0, skipped: 0, unmapped: [] }; }
    var parsed;
    try { parsed = typeof sfdtStr === 'string' ? JSON.parse(sfdtStr) : sfdtStr; }
    catch (e) { console.warn('[iPubEdit] parse SFDT failed', e); return { replaced: 0, skipped: 0, unmapped: [] }; }

    // Strip any tag-view marker runs so we don't overwrite real text with marker pills
    try { _stripTagMarkers(parsed); } catch (e) {}

    var updated = 0;
    var skipped = 0;
    var unmapped = [];
    var visited = new WeakSet();

    function walk(obj) {
        if (!obj || typeof obj !== 'object' || visited.has(obj)) return;
        visited.add(obj);
        if (Array.isArray(obj)) { for (var i = 0; i < obj.length; i++) walk(obj[i]); return; }

        var ccp = obj.ccp || obj.contentControlProperties;
        if (ccp && typeof ccp === 'object' && !Array.isArray(ccp)) {
            var rawTag = ccp.tg || ccp.tag || ccp.tt || ccp.title || '';
            var key = _ipubNormalizeTag(rawTag);
            var field = key ? _IPUB_TAG_FIELD[key] : null;
            if (field) {
                var inlines = _findFirstInlines(obj);
                if (inlines) {
                    var value = _ipubGetMetaValue(meta, field);
                    // Harvest cf from the first non-marker run so font/size survive
                    var preservedCf = {};
                    for (var j = 0; j < inlines.length; j++) {
                        var r = inlines[j];
                        var t = r && (r.tlp != null ? r.tlp : r.text);
                        if (typeof t === 'string' && t.charAt(0) !== _MARKER_PREFIX && r.cf) {
                            preservedCf = r.cf;
                            break;
                        }
                    }
                    // Replace inlines with a single run carrying the new value
                    inlines.length = 0;
                    inlines.push({ tlp: value, cf: preservedCf });
                    updated++;
                } else {
                    skipped++;
                }
            } else if (rawTag) {
                skipped++;
                if (unmapped.length < 3) unmapped.push(rawTag);
            }
        }

        var keys = Object.keys(obj);
        for (var k = 0; k < keys.length; k++) {
            var v = obj[keys[k]];
            if (v && typeof v === 'object') walk(v);
        }
    }

    try { walk(parsed); }
    catch (e) { console.warn('[iPubEdit] CC walk failed', e); return { replaced: 0, skipped: 0, unmapped: [] }; }

    if (updated > 0) {
        try {
            de.open(JSON.stringify(parsed));
            // If the user had tag-view on for this tab, re-apply markers after re-open
            try {
                var ts = (typeof _getTabState === 'function') ? _getTabState() : null;
                if (ts && ts.ccTagsVisible && typeof _applyTagVisibility === 'function') {
                    _applyTagVisibility(true);
                }
            } catch (eInner) {}
        } catch (e) {
            console.warn('[iPubEdit] de.open() after CC rewrite failed', e);
        }
    }

    console.log('[iPubEdit] applied ' + updated + ' value(s) to content control(s); '
        + 'skipped ' + skipped + ' unmapped tag(s)'
        + (unmapped.length ? ' e.g. ' + unmapped.join(', ') : ''),
        meta);
    return { replaced: updated, skipped: skipped, unmapped: unmapped };
};

// Backwards-compat alias — old callers that wanted a simple count.
window.ipubReplaceTokens = function (meta) {
    var res = window.applyMetadata(meta);
    return res ? (res.replaced || 0) : 0;
};

// Save metadata keyed by document id (tabId). documentId is REQUIRED for
// per-document persistence; callers pass the active tabId from Blazor.
window.saveIPubEditMeta = function (documentId, meta) {
    if (!documentId) {
        console.warn('[iPubEdit] saveIPubEditMeta called without documentId — aborting');
        return false;
    }
    try {
        var payload = Object.assign({}, meta, {
            _documentId: documentId,
            _savedAt: new Date().toISOString()
        });
        localStorage.setItem(_IPUB_META_PREFIX + documentId, JSON.stringify(payload));
        console.log('[iPubEdit] Saved meta for doc ' + documentId, payload);
    } catch (e) {
        console.warn('[iPubEdit] localStorage save failed', e);
        return false;
    }
    try { window.applyMetadata(meta); } catch (e) {}
    return true;
};

// Load metadata for a specific document. Falls back to the legacy global blob
// (one-time migration) if no per-doc entry exists.
window.loadIPubEditMeta = function (documentId) {
    if (!documentId) return null;
    try {
        var raw = localStorage.getItem(_IPUB_META_PREFIX + documentId);
        if (raw) return JSON.parse(raw);
        // Legacy fallback — pull the old single-key blob, stash it under this doc, then clear it
        var legacy = localStorage.getItem(_IPUB_META_LEGACY);
        if (legacy) {
            try {
                var parsed = JSON.parse(legacy);
                localStorage.setItem(_IPUB_META_PREFIX + documentId, legacy);
                localStorage.removeItem(_IPUB_META_LEGACY);
                return parsed;
            } catch (e) {}
        }
        return null;
    } catch (e) {
        return null;
    }
};

// Delete metadata for a document (e.g. on tab close if desired)
window.deleteIPubEditMeta = function (documentId) {
    if (!documentId) return;
    try { localStorage.removeItem(_IPUB_META_PREFIX + documentId); } catch (e) {}
};
