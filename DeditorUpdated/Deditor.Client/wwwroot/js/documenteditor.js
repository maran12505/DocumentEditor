var container;

// ── safeResize — poll until .n-editor-area has a real pixel width ─────────────
// Prevents "SVG attribute width: Expected length, NaN".
// Syncfusion reads offsetWidth from #container1 which starts at 0 until the
// flex layout has painted. We measure the PARENT (.n-editor-area) to avoid a
// feedback loop (Syncfusion overwrites #container1's inline width on every resize).
function safeResize(attemptsLeft) {
    if (!container) return;
    var parent = document.querySelector('.n-editor-area') ||
                 document.getElementById('container1');
    var w = parent ? parent.offsetWidth : 0;
    if (w > 0) {
        container.resize();
        return;
    }
    if (attemptsLeft <= 0) return;
    requestAnimationFrame(function () {
        setTimeout(function () { safeResize(attemptsLeft - 1); }, 50);
    });
}

// ── Initialize the Syncfusion Document Editor Container ──────────────────────
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

    // ── Initial resize: wait for flex layout to paint ────────────────────────
    safeResize(20);

    // ── ResizeObserver on .n-editor-area (the flex parent of #container1) ────
    // Observing the PARENT avoids the feedback loop where container.resize()
    // sets #container1's inline width, which would re-fire the observer.
    // Catches: sidebar open/close, window resize, DevTools open/close.
    if (typeof ResizeObserver !== 'undefined') {
        var _editorArea = document.querySelector('.n-editor-area');
        if (_editorArea) {
            var _ro = new ResizeObserver(function (entries) {
                for (var i = 0; i < entries.length; i++) {
                    if (entries[i].contentRect.width > 0 && container) {
                        safeResize(3);
                    }
                }
            });
            _ro.observe(_editorArea);
            window._deResizeObserver = _ro;
        }
    }

    // ── window 'resize' — belt-and-suspenders for DevTools toggle ────────────
    // DevTools open/close changes window.innerWidth. Some browsers do NOT
    // propagate this to ResizeObserver on interior flex children, so we
    // handle it explicitly with an 80 ms debounce.
    var _winResizeTimer = null;
    window.addEventListener('resize', function () {
        clearTimeout(_winResizeTimer);
        _winResizeTimer = setTimeout(function () { safeResize(5); }, 80);
    });
};

// ══════════════════════════════════════════════════════════════════════════════
// THEME MANAGEMENT
// ══════════════════════════════════════════════════════════════════════════════

// Apply or remove dark mode by toggling [data-theme="dark"] on <html>
window.applyTheme = function (isDark) {
    if (isDark) {
        document.documentElement.setAttribute('data-theme', 'dark');
    } else {
        document.documentElement.removeAttribute('data-theme');
    }
    // Persist choice so it survives page refresh
    try {
        localStorage.setItem('deditor-theme', isDark ? 'dark' : 'light');
    } catch (e) { /* storage may be unavailable */ }
};

// Read persisted theme preference
window.getThemePreference = function () {
    try {
        return localStorage.getItem('deditor-theme') || 'light';
    } catch (e) {
        return 'light';
    }
};

// ── Toggle Syncfusion Navigation Pane (Find / Replace) ───────────────────────
window.toggleNavigationPane = function (open) {
    if (!container) return;
    if (open) {
        container.documentEditor.showOptionsPane();
    } else {
        try {
            var mod = container.documentEditor.optionsPaneModule ||
                      container.documentEditor['optionsPaneModule'];
            if (mod && typeof mod.hideOptionPane === 'function') { mod.hideOptionPane(); return; }
            if (mod && typeof mod.destroy === 'function') {
                if (typeof mod.closePane === 'function') { mod.closePane(); return; }
                if (typeof mod.showHideOptionPane === 'function') { mod.showHideOptionPane(false); return; }
            }
        } catch (e) { /* fall through */ }
        var closeBtn = document.querySelector('.e-documenteditor-optionspane .e-de-op-close-button') ||
                       document.querySelector('.e-de-op .e-de-close-icon') ||
                       document.querySelector('.e-documenteditor-optionspane [title="Close"]');
        if (closeBtn) closeBtn.click();
    }
};

// ── Toggle Track Changes ──────────────────────────────────────────────────────
window.toggleTrackChanges = function () {
    if (!container) return false;
    var current = container.documentEditor.enableTrackChanges;
    container.documentEditor.enableTrackChanges = !current;
    return !current;
};

window.acceptAllChanges = function () {
    if (!container) return;
    container.documentEditor.revisions.acceptAll();
};

window.rejectAllChanges = function () {
    if (!container) return;
    container.documentEditor.revisions.rejectAll();
};

// ── Load / Serialize documents ────────────────────────────────────────────────
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
    if (isBlank) { container.documentEditor.openBlank(); }
    else         { container.documentEditor.open(newSfdt); }
    container.documentEditor.focusIn();
    return oldSfdt;
};

// ── File download from base64 ────────────────────────────────────────────────
window.downloadFile = function (base64, fileName) {
    var bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0));
    var blob  = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    var url   = URL.createObjectURL(blob);
    var a     = document.createElement("a");
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
};

// ── Selection helpers ────────────────────────────────────────────────────────
window.getSelectedText = function () {
    if (!container) return "";
    return container.documentEditor.selection.text;
};

window.replaceSelectedText = function (newText) {
    if (!container || !newText) return;
    var editor    = container.documentEditor.editor;
    var selection = container.documentEditor.selection;
    if (!selection || !selection.text || selection.text.length === 0) {
        editor.insertText(newText);
        return;
    }
    editor.delete();
    editor.insertText(newText);
};

window.hasSelection = function () {
    if (!container) return false;
    var sel = container.documentEditor.selection.text;
    return sel !== null && sel.length > 0;
};

// ══════════════════════════════════════════════════════════════════════════════
// CASE TRANSFORMATION LOGIC
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
    var result = text.split(/(\s+)/).map(function(token) {
        return token.trim().length === 0 ? token : capitalizeFirst(token);
    });
    window.replaceSelectedText(result.join(""));
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
        if (token.indexOf("-") > -1) {
            var parts = token.split("-");
            if (parts.some(function(p){ return importantWords.has(p.toLowerCase()); })) return capitalizeFirst(token);
        }
        return lower;
    });
    window.replaceSelectedText(result.join(""));
    return "OK";
};

function capitalizeFirst(word) {
    if (!word) return word;
    return word.charAt(0).toUpperCase() + word.slice(1);
}

// ── Get full document plain text ─────────────────────────────────────────────
window.getDocumentText = function () {
    if (!container) return "";
    container.documentEditor.selection.selectAll();
    var text = container.documentEditor.selection.text;
    container.documentEditor.selection.clear();
    return text || "";
};

// ── Scroll chat to bottom ─────────────────────────────────────────────────────
window.scrollChatToBottom = function () {
    var el = document.getElementById("chatMessages");
    if (el) el.scrollTop = el.scrollHeight;
};

// ── Live document name — direct DOM patch, no Blazor re-render needed ─────────
// Called from C# the instant the filename is known (even before upload completes).
// Bypasses Blazor's render cycle so the sidebar name updates in <1ms.
window.setDocumentName = function (name) {
    var display = (name && name.trim()) ? name.trim() : 'Untitled.docx';
    var el = document.getElementById('left-sidebar-filename');
    console.log('[JS setDocumentName] name="' + name + '" | display="' + display + '" | el found=' + !!el);
    if (el) {
        el.textContent = display;
        console.log('[JS setDocumentName] ✅ set to "' + el.textContent + '"');
    } else {
        console.warn('[JS setDocumentName] ⚠️ #left-sidebar-filename NOT IN DOM — sidebar may be closed!');
    }
    document.title = display + ' \u2014 DEditor';
};

// ── Resize editor when sidebar opens/closes ───────────────────────────────────
// Called by Blazor (ToggleLeftPane / ToggleChat / ToggleNavigationPane).
// Double-rAF: first frame ends the current JS task; second frame runs after
// the browser has committed the new flex widths — then safeResize polls until
// offsetWidth > 0, preventing both the NaN error and the visual shake.
window.resizeEditor = function () {
    if (!container) return;
    requestAnimationFrame(function () {
        requestAnimationFrame(function () {
            safeResize(5);
        });
    });
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
    if (direction === 'next') { results.index = (results.index + 1) % results.length; }
    else                      { results.index = (results.index - 1 + results.length) % results.length; }
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

window.clearSearch = function () {
    if (!container) return;
    container.documentEditor.search.searchResults.clear();
};

function refocusSearch() {
    setTimeout(function () {
        var el = document.getElementById('searchInput');
        if (el) { el.focus(); var len = el.value.length; el.setSelectionRange(len, len); }
    }, 30);
}