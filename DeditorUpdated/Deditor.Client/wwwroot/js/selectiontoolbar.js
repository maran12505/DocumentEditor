// ══════════════════════════════════════════════════════════════════════════════
// FLOATING SELECTION TOOLBAR — Multi-Editor Compatible
// Re-attaches listeners whenever the active editor changes.
// ══════════════════════════════════════════════════════════════════════════════

(function () {

    var _toolbar     = null;
    var _hideTimer   = null;
    var _showTimer   = null;
    var _attachedCanvas = null;   // track which canvas we're currently bound to
    var _mouseupHandler = null;
    var _keydownHandler = null;
    var _docMousedownHandler = null;

    // ── Build toolbar DOM (once) ────────────────────────────────────────────
    function createToolbar() {
        var el = document.createElement('div');
        el.id        = 'sel-toolbar';
        el.className = 'sel-toolbar';
        el.innerHTML = [
            btn('Bold',          '<b>B</b>',    'sel-fmt-bold',    'Bold (Ctrl+B)'),
            btn('Italic',        '<i>I</i>',    'sel-fmt-italic',  'Italic (Ctrl+I)'),
            btn('Underline',     '<u>U</u>',    'sel-fmt-under',   'Underline (Ctrl+U)'),
            btn('Strikethrough', '<s>S</s>',    'sel-fmt-strike',  'Strikethrough'),
            divider(),
            btn('TitleCase',     'Tt',          'sel-fmt-title',   'Title Case'),
            btn('InitialCaps',   'Aa',          'sel-fmt-initial', 'Initial Caps'),
            btn('EssentialCaps', '✦&thinsp;Ec', 'sel-fmt-ess',     'Essential Caps'),
            divider(),
            btn('Font',
                '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' +
                '<text x="3" y="17" font-size="14" fill="currentColor" stroke="none" font-family="serif">A</text>' +
                '<path d="M21 17l-5-10-5 10M8 14h6"/></svg>Font',
                'sel-fmt-font', 'Open Font dialog'),
        ].join('');
        document.body.appendChild(el);

        el.addEventListener('mousedown', function (e) {
            e.preventDefault();
            e.stopPropagation();
            var b = e.target.closest('[data-action]');
            if (b) handleAction(b.dataset.action);
        });

        el.addEventListener('mouseenter', cancelHide);
        el.addEventListener('mouseleave', function () { scheduleHide(400); });

        return el;
    }

    function btn(action, label, cls, title) {
        return '<button class="sel-tb-btn ' + cls +
               '" data-action="' + action +
               '" title="' + title + '">' + label + '</button>';
    }
    function divider() { return '<span class="sel-tb-div"></span>'; }

    // ── Position toolbar above the mouse release point ──────────────────────
    function showAt(x, y) {
        cancelHide();
        if (!_toolbar) _toolbar = createToolbar();
        _toolbar.style.display = 'flex';

        requestAnimationFrame(function () {
            var tw = _toolbar.offsetWidth  || 360;
            var th = _toolbar.offsetHeight || 36;
            var left = Math.max(8, Math.min(x - tw / 2, window.innerWidth - tw - 8));
            var top  = y - th - 10;
            if (top < 8) top = y + 14;
            _toolbar.style.left = left + 'px';
            _toolbar.style.top  = top  + 'px';
        });
    }

    function hideToolbar() {
        if (_toolbar) _toolbar.style.display = 'none';
        if (_showTimer) { clearTimeout(_showTimer); _showTimer = null; }
    }

    function scheduleHide(ms) {
        cancelHide();
        _hideTimer = setTimeout(hideToolbar, ms);
    }

    function cancelHide() {
        if (_hideTimer) { clearTimeout(_hideTimer); _hideTimer = null; }
    }

    // ── Detach old listeners ────────────────────────────────────────────────
    function detachFromCanvas() {
        if (_attachedCanvas) {
            if (_mouseupHandler) _attachedCanvas.removeEventListener('mouseup', _mouseupHandler);
            if (_keydownHandler) _attachedCanvas.removeEventListener('keydown', _keydownHandler);
            _attachedCanvas = null;
        }
        if (_docMousedownHandler) {
            document.removeEventListener('mousedown', _docMousedownHandler, true);
            _docMousedownHandler = null;
        }
    }

    // ── Attach to the currently active editor's canvas ──────────────────────
    function attachToActiveEditor() {
        if (!window._deContainer) {
            setTimeout(attachToActiveEditor, 500);
            return;
        }

        // Find the active editor's viewer container
        var canvas = null;
        if (typeof _activeEditorTabId !== 'undefined' && typeof _editors !== 'undefined' && _editors[_activeEditorTabId]) {
            var divId = _editors[_activeEditorTabId].divId;
            var editorDiv = document.getElementById(divId);
            if (editorDiv) {
                canvas = editorDiv.querySelector('[id$="_editor_viewerContainer"]');
            }
        }

        // Fallback: find any visible viewer container
        if (!canvas) {
            canvas = document.querySelector('#editorArea > div:not([style*="display: none"]) [id$="_editor_viewerContainer"]')
                  || document.querySelector('[id$="_editor_viewerContainer"]');
        }

        if (!canvas) {
            setTimeout(attachToActiveEditor, 500);
            return;
        }

        // Already attached to this exact canvas — skip
        if (_attachedCanvas === canvas) return;

        // Detach old, attach new
        detachFromCanvas();
        _attachedCanvas = canvas;

        var deEditor = window._deContainer.documentEditor;

        _mouseupHandler = function (e) {
            clearTimeout(_showTimer);
            _showTimer = setTimeout(function () {
                var text = deEditor.selection ? deEditor.selection.text : '';
                if (text && text.trim().length > 0) {
                    showAt(e.clientX, e.clientY);
                } else {
                    hideToolbar();
                }
            }, 80);
        };

        _keydownHandler = function () {
            setTimeout(function () {
                var text = deEditor.selection ? deEditor.selection.text : '';
                if (!text || text.trim().length === 0) hideToolbar();
            }, 80);
        };

        _docMousedownHandler = function (e) {
            if (!_toolbar || _toolbar.style.display === 'none') return;
            if (_toolbar.contains(e.target)) return;
            if (canvas.contains(e.target)) return;
            hideToolbar();
        };

        canvas.addEventListener('mouseup', _mouseupHandler);
        canvas.addEventListener('keydown', _keydownHandler);
        document.addEventListener('mousedown', _docMousedownHandler, true);
    }

    // ── Action handler ──────────────────────────────────────────────────────
    function handleAction(action) {
        if (!window._deContainer) return;
        var ed = window._deContainer.documentEditor.editor;

        switch (action) {
            case 'Bold':          ed.toggleBold();              break;
            case 'Italic':        ed.toggleItalic();            break;
            case 'Underline':     ed.toggleUnderline('Single'); break;
            case 'Strikethrough': ed.toggleStrikethrough();     break;
            case 'TitleCase':     window.applyTitleCase   && window.applyTitleCase();   break;
            case 'InitialCaps':   window.applyInitialCaps && window.applyInitialCaps(); break;
            case 'EssentialCaps': window.applyEssentialCaps && window.applyEssentialCaps(); break;
            case 'Font':
                window._deContainer.documentEditor.showDialog('Font');
                hideToolbar();
                return;
        }

        setTimeout(function () {
            var text = window._deContainer.documentEditor.selection
                       ? window._deContainer.documentEditor.selection.text : '';
            if (!text || text.trim().length === 0) hideToolbar();
        }, 150);
    }

    // ── Boot ────────────────────────────────────────────────────────────────
    window.addEventListener('load', function () {
        setTimeout(attachToActiveEditor, 1500);
    });

    // Public API — called from documenteditor.js after tab switch
    window.initSelectionToolbar = function () {
        hideToolbar();
        attachToActiveEditor();
    };

})();