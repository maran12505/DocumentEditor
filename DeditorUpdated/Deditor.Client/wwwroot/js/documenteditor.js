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
    container.documentEditor.enableTrackChanges = true;
    container.documentEditor.showRevisions = true;
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
            window._blazorEditorRef.invokeMethodAsync('OpenFileFromJS');
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
window.saveWithFilePicker = async function (tabId, base64, suggestedName, formatExt) {
    try {
        var mimeMap = {
            '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            '.doc':  'application/msword',
            '.rtf':  'application/rtf',
            '.txt':  'text/plain'
        };
        var descMap = {
            '.docx': 'Word Document',
            '.doc':  'Word 97-2003 Document',
            '.rtf':  'Rich Text Format',
            '.txt':  'Plain Text'
        };

        var ext  = formatExt || '.docx';
        var mime = mimeMap[ext] || mimeMap['.docx'];
        var desc = descMap[ext] || 'Document';

        // Build file type filters — selected format first, then others
        var types = [];
        types.push({
            description: desc,
            accept: {}
        });
        types[0].accept[mime] = [ext];

        // Add other formats as additional options
        Object.keys(mimeMap).forEach(function (e) {
            if (e !== ext) {
                var t = { description: descMap[e], accept: {} };
                t.accept[mimeMap[e]] = [e];
                types.push(t);
            }
        });

        var handle = await window.showSaveFilePicker({
            suggestedName: suggestedName || 'Untitled.docx',
            types: types
        });

        var bytes = Uint8Array.from(atob(base64), function (c) { return c.charCodeAt(0); });
        var writable = await handle.createWritable();
        await writable.write(bytes);
        await writable.close();

        // Store handle for future "Save" calls
        _fileHandles[tabId] = handle;

        // Return the actual chosen filename
        return handle.name || suggestedName;
    } catch (e) {
        if (e.name === 'AbortError') return 'CANCELLED';
        console.error('[SaveAs] picker failed:', e);
        return 'ERROR:' + e.message;
    }
};

// ── Open file with native picker (returns handle for write-back) ──────────────
// Returns: { fileName, base64 } on success, { error: "CANCELLED" } or { error: "msg" }
window.openWithFilePicker = async function (tabId) {
    try {
        var types = [
            {
                description: 'Documents',
                accept: {
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'],
                    'application/msword': ['.doc'],
                    'application/rtf': ['.rtf'],
                    'text/plain': ['.txt']
                }
            }
        ];

        var handles = await window.showOpenFilePicker({
            multiple: false,
            types: types
        });

        if (!handles || handles.length === 0) return { error: 'CANCELLED' };

        var handle = handles[0];
        var file = await handle.getFile();

        // Read file as base64
        var arrayBuffer = await file.arrayBuffer();
        var bytes = new Uint8Array(arrayBuffer);
        var binary = '';
        for (var i = 0; i < bytes.length; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        var base64 = btoa(binary);

        // Store the handle so Save can write back to this file
        _fileHandles[tabId] = handle;

        return { fileName: file.name, base64: base64 };
    } catch (e) {
        if (e.name === 'AbortError') return { error: 'CANCELLED' };
        console.error('[Open] picker failed:', e);
        return { error: e.message };
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
    container.documentEditor.open(sfdt);
    container.documentEditor.focusIn();
};
window.getDocumentContent = function () {
    if (!container) return null;
    return container.documentEditor.serialize();
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
    clearInterval(_autoSaveTimer);
    _lastSavedSfdt = '';
    _autoSaveTimer = setInterval(function () {
        window.autoSaveNow(tabId, fileName);
    }, AUTOSAVE_INTERVAL);
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
            try { dotNetRef.invokeMethodAsync('OnSaveAsShortcut'); } catch (err) {}
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