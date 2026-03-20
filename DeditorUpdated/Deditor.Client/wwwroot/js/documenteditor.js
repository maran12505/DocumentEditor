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
 
        // Close the file menu popup
        try {
            var fileBtn = document.querySelector('.e-ribbon-file-menu');
            if (fileBtn) fileBtn.click();
        } catch (ex) {}
 
        if (text === 'New') {
            window._blazorEditorRef.invokeMethodAsync('NewTabFromJS');
        } else {
            window._blazorEditorRef.invokeMethodAsync('OpenFileFromJS');
        }
    }, true); // capture phase — fires before Syncfusion's handler
}

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
// DOWNLOAD
// ══════════════════════════════════════════════════════════════════════════════
window.downloadFile = function (base64, fileName) {
    var bytes = Uint8Array.from(atob(base64), function(c){ return c.charCodeAt(0); });
    var blob  = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    var url   = URL.createObjectURL(blob);
    var a     = document.createElement("a");
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
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
// Parses the serialized SFDT for heading-level paragraphs.
// Returns [{level:1,text:"Intro"}, {level:2,text:"Background"}, ...]
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
    // Paragraph with a heading style
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
    // Recurse into rows/cells (tables)
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
// Storage keys:
//   deditor-autosave-v1:{tabId}  →  { sfdt, fileName, savedAt }
//   deditor-autosave-ids         →  JSON array of tabIds
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

// Bridge custom event → Blazor DotNetObjectReference
window.registerAutoSaveListener = function (dotNetRef) {
    document.addEventListener('deditor-autosaved', function (e) {
        try { dotNetRef.invokeMethodAsync('OnAutoSaved', e.detail); } catch (err) {}
    });
};

// ══════════════════════════════════════════════════════════════════════════════
// RECENT FILES
// Stored as JSON array (max 10) under 'deditor-recent-files'
// Each entry: { name: string, openedAt: ISO string }
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
// Deletes the "/" that triggered the picker, then inserts the snippet body.
// ══════════════════════════════════════════════════════════════════════════════
window.insertSnippetText = function (body) {
    if (!container || !body) return;
    var editor = container.documentEditor.editor;
    // Delete the "/" character that triggered the picker
    editor.delete();
    // Insert each line; use insertText for lines and onEnter for newlines
    var lines = body.split('\n');
    lines.forEach(function(line, idx) {
        editor.insertText(line);
        if (idx < lines.length - 1) editor.onEnter();
    });
    container.documentEditor.focusIn();
};

// ══════════════════════════════════════════════════════════════════════════════
// KEYBOARD LISTENERS
// Registers:
//   • Escape  → Blazor.OnEscapePressed  (exit focus mode, close slash picker)
//   • "/" key → Blazor.OnSlashTyped     (open slash snippet picker)
// ══════════════════════════════════════════════════════════════════════════════
window.registerEditorKeyListeners = function (dotNetRef) {
      window._blazorEditorRef = dotNetRef;
    // Escape — capture phase so it fires before Syncfusion
    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape') {
            try { dotNetRef.invokeMethodAsync('OnEscapePressed'); } catch (err) {}
        }
    }, true);

    // Slash — wait for Syncfusion to render the char, then check if "/" was typed
    // We listen on the viewer container (the actual editable canvas area)
    function attachSlashListener() {
        var canvas = document.getElementById('container1_editor_viewerContainer');
        if (!canvas || !window._deContainer) { setTimeout(attachSlashListener, 600); return; }

        canvas.addEventListener('keyup', function (e) {
            if (e.key !== '/') return;
            // Get cursor position on screen to position the picker
            var sel = window._deContainer ? window._deContainer.documentEditor.selection : null;
            if (!sel) return;
            // Use the caret rectangle from the viewer if available, else fall back to mouse position
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