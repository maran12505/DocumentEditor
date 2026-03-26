// ══════════════════════════════════════════════════════════════════════════════
// FLOATING SELECTION TOOLBAR
// Syncfusion Document Editor renders on a <canvas> — NOT an iframe.
// We listen on the viewerContainer div for mouseup, then check Syncfusion's
// own selection.text property. Toolbar shows above selected text.
// ══════════════════════════════════════════════════════════════════════════════

(function () {

    var _toolbar     = null;
    var _hideTimer   = null;
    var _showTimer   = null;

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

        // Prevent editor losing selection; stop event reaching global hide listener
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
            if (top < 8) top = y + 14;          // flip below if near top edge
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

    // ── Attach once the editor is ready ─────────────────────────────────────
    function attachToEditor() {
        // The viewer container is the scrollable canvas area
        var canvas = document.querySelector('#editorArea > div:not([style*="display:none"]) [id$="_editor_viewerContainer"]') || document.querySelector('[id$="_editor_viewerContainer"]');
        if (!canvas || !window._deContainer) {
            setTimeout(attachToEditor, 500);
            return;
        }

        var deEditor = window._deContainer.documentEditor;

        // ── mouseup on the canvas: check selection after Syncfusion settles ──
        canvas.addEventListener('mouseup', function (e) {
            clearTimeout(_showTimer);
            _showTimer = setTimeout(function () {
                var text = deEditor.selection ? deEditor.selection.text : '';
                if (text && text.trim().length > 0) {
                    showAt(e.clientX, e.clientY);
                } else {
                    hideToolbar();
                }
            }, 80); // 80ms is enough for Syncfusion to update selection.text
        });

        // ── Global mousedown (capture): hide when clicking outside canvas ───
        document.addEventListener('mousedown', function (e) {
            if (!_toolbar || _toolbar.style.display === 'none') return;
            // Toolbar itself — stopPropagation handles it
            if (_toolbar.contains(e.target)) return;
            // Click landed on the canvas → mouseup will decide
            if (canvas.contains(e.target)) return;
            // Anything else (ribbon, topbar, sidebar…) → hide immediately
            hideToolbar();
        }, true);

        // ── Keyboard: hide when selection is cleared ─────────────────────────
        canvas.addEventListener('keydown', function () {
            setTimeout(function () {
                var text = deEditor.selection ? deEditor.selection.text : '';
                if (!text || text.trim().length === 0) hideToolbar();
            }, 80);
        });
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

        // Keep toolbar only if selection still exists after the action
        setTimeout(function () {
            var text = window._deContainer.documentEditor.selection
                       ? window._deContainer.documentEditor.selection.text : '';
            if (!text || text.trim().length === 0) hideToolbar();
        }, 150);
    }

    // ── Boot ────────────────────────────────────────────────────────────────
    window.addEventListener('load', function () {
        setTimeout(attachToEditor, 1500);
    });

    window.initSelectionToolbar = function () { attachToEditor(); };

})();