// ══════════════════════════════════════════════════════════════════════════════
// iPubEdit Properties — right-side panel
//   - Tree of elements from elements-{book|journal}.json (lazy render)
//   - Real-time search filter
//   - Selection sync: shows tag of CC under cursor
//   - Buttons: Reload (re-read CC), Activate (wrap selection), Remove (unwrap)
//   - Toggle via Ctrl+Alt+M, ribbon button, or first-doc-load auto-open
// ══════════════════════════════════════════════════════════════════════════════
(function () {
    'use strict';

    // Try the RCL-mounted path first (Web/Maui hosts), then fall back to a root path
    // (e.g. when the Shared.UI assets are served from the host's own wwwroot).
    var DATA_BASE_CANDIDATES = [
        '/_content/Deditor.Shared.UI/data/',
        '/data/'
    ];
    var STORAGE_KEY = 'ipubProps.open';  // 'open' | 'closed'
    var AUTO_OPENED_KEY = 'ipubProps.autoOpenedOnce';

    // ── State ─────────────────────────────────────────────────────────────────
    var _treeData = null;        // array of root nodes (after fetch)
    var _byId = new Map();       // id → node
    var _byName = new Map();     // lowercased name → node[]
    var _flatList = [];          // depth-first flat order, used for filter scan
    var _expanded = new Set();   // expanded node ids
    var _selectedId = null;      // currently highlighted tree node id
    var _activeCC = null;        // current CC under cursor (Syncfusion object)
    var _hookedContainerIds = new Set(); // container element ids we've already wired
    var _xmlTypeLoaded = null;   // 'Book' | 'Journal'
    var _filterTimer = null;
    var _selSyncTimer = null;

    // ── Public API (window.IPubProperties) ────────────────────────────────────
    var api = {
        init: init,                       // call after editor mounts; idempotent
        ensureLoaded: ensureLoaded,       // called by toggle/auto-open
        toggle: toggle,
        open: open,
        close: close,
        isOpen: function () { return _isOpen(); },
        onSelectionChanged: throttledSyncFromSelection,
        onDocumentLoaded: onDocumentLoaded,
        getXmlType: function () { return _xmlTypeLoaded; },
    };
    window.IPubProperties = api;

    // ── Init: wire panel DOM events once ──────────────────────────────────────
    function init() {
        var panel = document.getElementById('ipubPropertiesPanel');
        if (!panel || panel.__ipubWired) return;
        panel.__ipubWired = true;

        var closeBtn   = panel.querySelector('.ipp-close');
        var btnReload  = panel.querySelector('#ippReload');
        var btnActivate= panel.querySelector('#ippActivate');
        var btnRemove  = panel.querySelector('#ippRemove');
        var btnSearchTag = panel.querySelector('#ippSearchBtn');
        var tagInput   = panel.querySelector('#ippTagInput');

        if (closeBtn)    closeBtn.addEventListener('click', close);
        if (btnReload)   btnReload.addEventListener('click', reload);
        if (btnActivate) btnActivate.addEventListener('click', activateSelection);
        if (btnRemove)   btnRemove.addEventListener('click', removeAtCursor);
        if (btnSearchTag) btnSearchTag.addEventListener('click', function () {
            if (tagInput) applyFilter(tagInput.value || '');
        });
        // Top "tag name" input doubles as the tree search/filter (live, debounced).
        if (tagInput) {
            tagInput.addEventListener('input', function (e) {
                if (_filterTimer) clearTimeout(_filterTimer);
                var val = e.target.value;
                _filterTimer = setTimeout(function () { applyFilter(val); }, 120);
            });
        }
        // Tab buttons (only Element implemented in v1)
        panel.querySelectorAll('.ipp-tabs button').forEach(function (b) {
            b.addEventListener('click', function () {
                if (b.classList.contains('ipp-tab-disabled')) return;
                panel.querySelectorAll('.ipp-tabs button').forEach(function (x) { x.classList.remove('active'); });
                b.classList.add('active');
            });
        });

        _attachVisibilityObserver();

        // Always pre-populate the attributes section so the layout never jumps
        // when the user first selects a node — section is visible from the start.
        renderAttrs([]);

        // Restore last open/closed pref. If user closed it last session, leave closed.
        var pref = localStorage.getItem(STORAGE_KEY);
        if (pref === 'open' && !_isOpen()) ensureLoaded().then(_clickBridge);
    }

    // ── XmlType detection (read from iPubEdit Meta if available) ──────────────
    function detectXmlType() {
        try {
            var raw = localStorage.getItem('ipubedit.meta.xmlType');
            if (raw === 'Journal' || raw === 'Book') return raw;
        } catch (e) {}
        // Fallback: peek at any global meta exposed by Razor
        if (window.__ipubEditMeta && window.__ipubEditMeta.XmlType) return window.__ipubEditMeta.XmlType;
        return 'Book';
    }

    // ── Load elements JSON (book or journal) ──────────────────────────────────
    function ensureLoaded() {
        var want = detectXmlType();
        if (_treeData && _xmlTypeLoaded === want) return Promise.resolve();
        var fileName = want === 'Journal' ? 'elements-journal.json' : 'elements-book.json';

        // Try each candidate base in order until one succeeds.
        function tryNext(idx) {
            if (idx >= DATA_BASE_CANDIDATES.length) {
                console.error('[IPubProperties] All candidate URLs failed for ' + fileName);
                return Promise.resolve();
            }
            var url = DATA_BASE_CANDIDATES[idx] + fileName;
            return fetch(url, { cache: 'force-cache' }).then(function (r) {
                if (!r.ok) throw new Error('HTTP ' + r.status);
                return r.json();
            }).then(function (json) {
                _treeData = json;
                _xmlTypeLoaded = want;
                _byId.clear(); _byName.clear(); _flatList = []; _expanded.clear();
                indexTree(json, 0);
                // Auto-expand root level so the tree shows structure immediately
                // (matches desktop iPubEdit "Core Elements > Body > title" reference view).
                for (var ri = 0; ri < json.length; ri++) {
                    if (json[ri].children && json[ri].children.length) _expanded.add(json[ri].id);
                }
                renderTree();
                console.log('[IPubProperties] Loaded ' + want + ' from ' + url + ': ' + _byId.size + ' nodes');
            }).catch(function () { return tryNext(idx + 1); });
        }
        return tryNext(0);
    }

    function indexTree(list, depth) {
        for (var i = 0; i < list.length; i++) {
            var n = list[i];
            n.depth = depth;
            _byId.set(n.id, n);
            _flatList.push(n);
            var key = (n.name || '').toLowerCase();
            if (key) {
                if (!_byName.has(key)) _byName.set(key, []);
                _byName.get(key).push(n);
            }
            if (n.children && n.children.length) indexTree(n.children, depth + 1);
        }
    }

    // ── Tree rendering (lazy, only renders expanded subtree) ──────────────────
    function renderTree() {
        var root = document.getElementById('ippTreeRoot');
        if (!root || !_treeData) return;
        root.innerHTML = '';
        var ul = document.createElement('ul');
        ul.className = 'ipp-tree-ul ipp-tree-root-ul';
        for (var i = 0; i < _treeData.length; i++) ul.appendChild(buildNodeLi(_treeData[i]));
        root.appendChild(ul);
    }

    function buildNodeLi(node) {
        var li = document.createElement('li');
        li.className = 'ipp-tree-li';
        li.dataset.id = node.id;

        var hasKids = node.children && node.children.length > 0;
        var row = document.createElement('div');
        row.className = 'ipp-tree-row' + (hasKids ? ' has-kids' : '');
        if (_selectedId === node.id) row.classList.add('selected');

        var twisty = document.createElement('span');
        twisty.className = 'ipp-twisty';
        twisty.textContent = hasKids ? (_expanded.has(node.id) ? '▾' : '▸') : '';
        row.appendChild(twisty);

        var label = document.createElement('span');
        label.className = 'ipp-tree-label';
        label.textContent = node.name || '?';
        label.title = node.rawTag || node.name;
        row.appendChild(label);

        row.addEventListener('click', function (e) {
            e.stopPropagation();
            selectNode(node.id);
            if (hasKids) toggleExpand(node.id);
        });
        row.addEventListener('dblclick', function (e) {
            e.stopPropagation();
            e.preventDefault();
            applyContentControl(node);
        });
        li.appendChild(row);

        if (hasKids && _expanded.has(node.id)) {
            var ul = document.createElement('ul');
            ul.className = 'ipp-tree-ul';
            for (var i = 0; i < node.children.length; i++) ul.appendChild(buildNodeLi(node.children[i]));
            li.appendChild(ul);
        }
        return li;
    }

    function toggleExpand(id) {
        if (_expanded.has(id)) _expanded.delete(id); else _expanded.add(id);
        renderTree();
    }

    function selectNode(id) {
        _selectedId = id;
        var n = _byId.get(id);
        if (!n) return;
        var tagInput = document.getElementById('ippTagInput');
        if (tagInput) tagInput.value = (n.name || '');
        var tagLabel = document.getElementById('ippTagLabel');
        if (tagLabel) tagLabel.textContent = formatTagLabel(n);
        renderAttrs(n.attributes || []);
        renderTree();
        // ensure visible row in scroll
        var li = document.querySelector('#ippTreeRoot li[data-id="' + id + '"]');
        if (li) li.scrollIntoView({ block: 'nearest' });
    }

    function expandAncestors(node) {
        var p = node && node.parentId;
        while (p) {
            _expanded.add(p);
            var pn = _byId.get(p);
            p = pn && pn.parentId;
        }
    }

    // ── Filter: substring match on name; auto-expand all matched ancestors ───
    function applyFilter(query) {
        var q = (query || '').trim().toLowerCase();
        var root = document.getElementById('ippTreeRoot');
        if (!root) return;

        if (!q) {
            // Reset to root view
            _expanded.clear();
            renderTree();
            return;
        }
        // Find all matching nodes, expand their ancestors, hide non-matching subtrees
        var matchSet = new Set();
        var visibleSet = new Set();
        for (var i = 0; i < _flatList.length; i++) {
            var n = _flatList[i];
            if ((n.name || '').toLowerCase().indexOf(q) !== -1) {
                matchSet.add(n.id);
                visibleSet.add(n.id);
                // walk ancestors to make them visible + expanded
                var p = n.parentId;
                while (p) {
                    visibleSet.add(p);
                    _expanded.add(p);
                    var pn = _byId.get(p);
                    p = pn && pn.parentId;
                }
            }
        }
        renderFilteredTree(matchSet, visibleSet, q);
    }

    function renderFilteredTree(matchSet, visibleSet, query) {
        var root = document.getElementById('ippTreeRoot');
        if (!root) return;
        root.innerHTML = '';
        var ul = document.createElement('ul');
        ul.className = 'ipp-tree-ul ipp-tree-root-ul';
        for (var i = 0; i < _treeData.length; i++) {
            var li = buildFilteredLi(_treeData[i], matchSet, visibleSet, query);
            if (li) ul.appendChild(li);
        }
        root.appendChild(ul);
    }

    function buildFilteredLi(node, matchSet, visibleSet, query) {
        if (!visibleSet.has(node.id)) return null;
        var li = document.createElement('li');
        li.className = 'ipp-tree-li';
        li.dataset.id = node.id;

        var hasKids = node.children && node.children.length > 0;
        var row = document.createElement('div');
        row.className = 'ipp-tree-row' + (hasKids ? ' has-kids' : '');
        if (_selectedId === node.id) row.classList.add('selected');
        if (matchSet.has(node.id)) row.classList.add('matched');

        var twisty = document.createElement('span');
        twisty.className = 'ipp-twisty';
        twisty.textContent = hasKids ? '▾' : '';
        row.appendChild(twisty);

        var label = document.createElement('span');
        label.className = 'ipp-tree-label';
        label.innerHTML = highlightMatch(node.name || '?', query);
        label.title = node.rawTag || node.name;
        row.appendChild(label);

        row.addEventListener('click', function (e) {
            e.stopPropagation();
            selectNode(node.id);
        });
        row.addEventListener('dblclick', function (e) {
            e.stopPropagation();
            e.preventDefault();
            applyContentControl(node);
        });
        li.appendChild(row);

        if (hasKids) {
            var ul = document.createElement('ul');
            ul.className = 'ipp-tree-ul';
            var childAdded = false;
            for (var i = 0; i < node.children.length; i++) {
                var c = buildFilteredLi(node.children[i], matchSet, visibleSet, query);
                if (c) { ul.appendChild(c); childAdded = true; }
            }
            if (childAdded) li.appendChild(ul);
        }
        return li;
    }

    function highlightMatch(text, query) {
        if (!query) return escapeHtml(text);
        var idx = text.toLowerCase().indexOf(query);
        if (idx < 0) return escapeHtml(text);
        return escapeHtml(text.slice(0, idx)) +
            '<mark>' + escapeHtml(text.slice(idx, idx + query.length)) + '</mark>' +
            escapeHtml(text.slice(idx + query.length));
    }

    function escapeHtml(s) {
        return String(s).replace(/[&<>"]/g, function (c) {
            return c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : '&quot;';
        });
    }

    // ── Attribute table ───────────────────────────────────────────────────────
    function renderAttrs(attrs) {
        var tbl = document.getElementById('ippAttrTable');
        if (!tbl) return;
        tbl.innerHTML = '';
        // Group header row — matches desktop "▾ Attributes" collapser style
        var hdr = document.createElement('tr');
        hdr.className = 'ipp-attr-group';
        hdr.innerHTML = '<td colspan="2"><span class="ipp-attr-twisty">▾</span> Attributes</td>';
        tbl.appendChild(hdr);
        if (!attrs || attrs.length === 0) {
            var tr = document.createElement('tr');
            tr.innerHTML = '<td colspan="2" class="ipp-attr-empty">— No attributes —</td>';
            tbl.appendChild(tr);
            return;
        }
        for (var i = 0; i < attrs.length; i++) {
            var a = attrs[i];
            var tr = document.createElement('tr');
            var nameTd = document.createElement('td');
            nameTd.className = 'ipp-attr-name';
            nameTd.textContent = a.name;
            var valTd = document.createElement('td');
            valTd.className = 'ipp-attr-val';
            var input = document.createElement('input');
            input.type = 'text';
            input.value = a.value || '';
            input.placeholder = '';
            input.dataset.attr = a.name;
            valTd.appendChild(input);
            tr.appendChild(nameTd);
            tr.appendChild(valTd);
            tbl.appendChild(tr);
        }
    }

    // ── Selection sync: find innermost CC at cursor, populate panel ──────────
    function throttledSyncFromSelection() {
        if (_selSyncTimer) return;
        _selSyncTimer = setTimeout(function () {
            _selSyncTimer = null;
            try { syncFromSelection(); } catch (e) { /* swallow — non-critical */ }
        }, 60);
    }

    function syncFromSelection() {
        var c = window.container;
        if (!c || !c.documentEditor) return;
        var de = c.documentEditor;
        var cc = findInnermostCCAtCursor(de);
        _activeCC = cc;
        if (!cc) {
            updateTagDisplay('', null);
            return;
        }
        var tag = readCCTag(cc);
        var parsed = parseTag(tag);
        updateTagDisplay(tag, parsed);
        // Highlight in tree
        if (parsed.name) {
            var matches = _byName.get(parsed.name.toLowerCase());
            if (matches && matches.length) {
                var node = matches[0];
                _selectedId = node.id;
                expandAncestors(node);
                // Merge attribute names from JSON catalog with values from CC
                var attrsForRender = (node.attributes || []).map(function (a) {
                    return { name: a.name, value: parsed.attrs[a.name] || '' };
                });
                // Add any attrs found in tag but not in catalog
                Object.keys(parsed.attrs).forEach(function (k) {
                    if (!attrsForRender.some(function (x) { return x.name === k; }))
                        attrsForRender.push({ name: k, value: parsed.attrs[k] });
                });
                renderAttrs(attrsForRender);
                renderTree();
                var li = document.querySelector('#ippTreeRoot li[data-id="' + node.id + '"]');
                if (li) li.scrollIntoView({ block: 'nearest' });
            }
        }
    }

    function findInnermostCCAtCursor(de) {
        try {
            var helper = de.documentHelper;
            if (!helper || !helper.contentControlCollection) return null;
            var sel = de.selection;
            if (!sel || !sel.start) return null;
            var cursor = sel.start;
            var best = null;
            var bestDepth = -1;
            var ccs = helper.contentControlCollection;
            for (var i = 0; i < ccs.length; i++) {
                var cc = ccs[i];
                if (!cc || !cc.contentControlWidgetType) continue;
                if (containsCursor(cc, cursor)) {
                    // depth approximation: use list index as tie-breaker (innermost = last added)
                    if (i > bestDepth) { best = cc; bestDepth = i; }
                }
            }
            return best;
        } catch (e) { return null; }
    }

    function containsCursor(cc, cursor) {
        try {
            var s = cc.contentControlWidgetType === 'Block' ? cc.contentControlStart : cc.start;
            var e = cc.contentControlWidgetType === 'Block' ? cc.contentControlEnd   : cc.end;
            // Syncfusion: TextPosition.isExistAfter(other)
            if (!s || !e) return false;
            var afterStart = !cursor.isExistBefore(s);
            var beforeEnd  = cursor.isExistBefore(e);
            return afterStart && beforeEnd;
        } catch (err) { return false; }
    }

    function readCCTag(cc) {
        var ccp = cc.contentControlProperties || cc;
        return ccp.tag || ccp.tg || ccp.title || ccp.tt || '';
    }

    function parseTag(tag) {
        // "#book-part;id=2;type=chapter;" → { name: 'book-part', attrs: { id:'2', type:'chapter' } }
        // "snm" → { name: 'snm', attrs: {} }
        var out = { name: '', attrs: {} };
        if (!tag) return out;
        var t = String(tag).trim();
        if (t.charAt(0) === '#') t = t.slice(1);
        var parts = t.split(';');
        out.name = parts[0] || '';
        for (var i = 1; i < parts.length; i++) {
            var p = parts[i].trim();
            if (!p) continue;
            var eq = p.indexOf('=');
            if (eq > 0) out.attrs[p.slice(0, eq).trim()] = p.slice(eq + 1).trim();
        }
        return out;
    }

    function updateTagDisplay(rawTag, parsed) {
        var tagLabel = document.getElementById('ippTagLabel');
        var tagInput = document.getElementById('ippTagInput');
        var attrTag  = document.getElementById('ippAttrTag');
        var name = (parsed && parsed.name) || '';
        var pretty = name ? '#' + name + ';' : '— no tag —';
        if (tagLabel) tagLabel.textContent = pretty;
        if (attrTag)  attrTag.textContent  = pretty;
        if (tagInput && name) tagInput.value = name;
    }

    function formatTagLabel(node) {
        var t = node.rawTag || node.name;
        if (!t) return '';
        if (t.charAt(0) === '#') return t;
        return t;
    }

    // ── Action buttons ────────────────────────────────────────────────────────
    function reload() { syncFromSelection(); }

    function activateSelection() {
        var c = window.container;
        if (!c || !c.documentEditor) return;
        var de = c.documentEditor;
        // If cursor is already inside a CC, focus/highlight it and refresh the panel
        var existingCC = findInnermostCCAtCursor(de);
        if (existingCC) {
            if (typeof de.focusIn === 'function') de.focusIn();
            syncFromSelection();
            return;
        }
        // Otherwise apply a new CC from the tree-selected element
        var n = _selectedId ? _byId.get(_selectedId) : null;
        if (!n) { console.warn('[IPubProperties] No element selected in tree'); return; }
        applyContentControl(n);
    }

    function applyContentControl(node) {
        var c = window.container;
        if (!c || !c.documentEditor) return;
        var de = c.documentEditor;
        if (!de.selection) return;

        // Guard: require a non-empty text selection
        var selText = '';
        try { selText = de.selection.text || ''; } catch (e) {}
        if (!selText || !selText.trim()) {
            _flashTagLabel('⚠ Select text first');
            return;
        }

        // Build tag string from element name + any filled attribute inputs
        var attrParts = [];
        var tbl = document.getElementById('ippAttrTable');
        if (tbl) {
            tbl.querySelectorAll('input[data-attr]').forEach(function (inp) {
                var v = inp.value.trim();
                if (v) attrParts.push(inp.dataset.attr + '=' + v);
            });
        }
        var tagStr = '#' + node.name + ';' + (attrParts.length ? attrParts.join(';') + ';' : '');

        try {
            // Wrap the current text selection in a Rich Text content control
            de.editor.insertContentControl('RichText');

            // Stamp the element name as title + tag on the newly inserted CC
            var ccInfo = { title: node.name, tag: tagStr, lockContentControl: false, lockContents: false };
            if (typeof de.editor.setContentControlInfo === 'function') {
                de.editor.setContentControlInfo(ccInfo);
            } else {
                // Fallback: directly patch the widget's properties object
                var newCC = findInnermostCCAtCursor(de);
                if (newCC && newCC.contentControlProperties) {
                    newCC.contentControlProperties.tag = tagStr;
                    newCC.contentControlProperties.title = node.name;
                }
            }

            console.log('[IPubProperties] Applied: ' + tagStr);
            if (_selectedId !== node.id) selectNode(node.id);
            setTimeout(function () {
                syncFromSelection();
                if (typeof de.focusIn === 'function') de.focusIn();
            }, 50);
        } catch (e) {
            console.error('[IPubProperties] Apply content control failed:', e);
        }
    }

    function _flashTagLabel(msg) {
        var tagLabel = document.getElementById('ippTagLabel');
        if (!tagLabel) return;
        var prev = tagLabel.textContent;
        var prevColor = tagLabel.style.color;
        tagLabel.textContent = msg;
        tagLabel.style.color = '#e65c5c';
        setTimeout(function () { tagLabel.textContent = prev; tagLabel.style.color = prevColor; }, 2000);
    }

    function removeAtCursor() {
        var c = window.container;
        if (!c || !c.documentEditor) return;
        var de = c.documentEditor;
        var cc = findInnermostCCAtCursor(de);
        if (!cc) {
            _flashTagLabel('⚠ Cursor not in a tagged element');
            return;
        }
        try {
            if (typeof de.editor.deleteContentControl === 'function') {
                de.editor.deleteContentControl();
            } else if (typeof de.editor.setContentControlInfo === 'function') {
                // Fallback: clear tag metadata — SDT wrapper stays but becomes untagged
                de.editor.setContentControlInfo({ title: '', tag: '', lockContentControl: false, lockContents: false });
                console.warn('[IPubProperties] deleteContentControl unavailable; cleared tag metadata only');
            } else {
                console.warn('[IPubProperties] No remove API found in this Syncfusion build');
            }
            _activeCC = null;
            updateTagDisplay('', null);
            renderAttrs([]);
        } catch (e) {
            console.error('[IPubProperties] Remove failed:', e);
        }
    }

    // ── Document-load auto-open (first time only unless user closed last time) ─
    function onDocumentLoaded() {
        try {
            var pref = localStorage.getItem(STORAGE_KEY);
            if (pref === 'closed') return;     // user explicitly closed; respect it
            var auto = sessionStorage.getItem(AUTO_OPENED_KEY);
            if (auto === '1') return;          // already auto-opened this session
            sessionStorage.setItem(AUTO_OPENED_KEY, '1');
            ensureLoaded().then(open);
        } catch (e) {}
    }

    // ── Selection-change wiring (called from documenteditor.js) ──────────────
    api.attachToContainer = function (containerEl, syncfusionContainer) {
        if (!syncfusionContainer || !syncfusionContainer.documentEditor) return;
        var key = (containerEl && containerEl.id) || 'default';
        if (_hookedContainerIds.has(key)) return;
        _hookedContainerIds.add(key);
        try {
            syncfusionContainer.documentEditor.selectionChange = function () {
                throttledSyncFromSelection();
            };
        } catch (e) { console.warn('[IPubProperties] selectionChange hook failed:', e); }
    };

    // ── Open/close ────────────────────────────────────────────────────────────
    // Source of truth: Blazor's _showPropertiesPanel (controls panel `hidden` attribute).
    // All open/close paths click the hidden bridge button #nBtnProperties so Blazor
    // stays in sync. A MutationObserver below watches the `hidden` attribute and
    // performs side-effects (localStorage persistence, data load, editor resize).
    function _isOpen() {
        var p = document.getElementById('ipubPropertiesPanel');
        return !!(p && !p.hidden);
    }
    function _clickBridge() {
        var b = document.getElementById('nBtnProperties');
        if (b) b.click();
    }
    function open()  { if (!_isOpen()) _clickBridge(); }
    function close() { if (_isOpen())  _clickBridge(); }
    function toggle(){ ensureLoaded().then(_clickBridge); }

    // Watch the panel's hidden attribute → run open/close side-effects
    function _attachVisibilityObserver() {
        var p = document.getElementById('ipubPropertiesPanel');
        if (!p || p.__ipubVisObs) return;
        p.__ipubVisObs = true;
        var lastHidden = p.hidden;
        var obs = new MutationObserver(function () {
            if (p.hidden === lastHidden) return;
            lastHidden = p.hidden;
            if (!p.hidden) {
                document.body.classList.add('ipub-props-open');
                try { localStorage.setItem(STORAGE_KEY, 'open'); } catch (e) {}
                ensureLoaded().then(function () { syncFromSelection(); });
            } else {
                document.body.classList.remove('ipub-props-open');
                try { localStorage.setItem(STORAGE_KEY, 'closed'); } catch (e) {}
            }
            if (typeof window.safeResize === 'function')
                setTimeout(function () { window.safeResize(8); }, 50);
        });
        obs.observe(p, { attributes: true, attributeFilter: ['hidden'] });
        // Run initial side-effects to match current state
        if (!p.hidden) {
            document.body.classList.add('ipub-props-open');
            ensureLoaded().then(function () { syncFromSelection(); });
        }
    }

    // Wire up Ctrl+Alt+M globally (idempotent)
    if (!window.__ipubPropsKeyHook) {
        window.__ipubPropsKeyHook = true;
        document.addEventListener('keydown', function (e) {
            if ((e.ctrlKey || e.metaKey) && e.altKey && (e.key === 'm' || e.key === 'M')) {
                e.preventDefault();
                e.stopPropagation();
                toggle();
            }
        }, true);
    }

    // Init when DOM is ready (panel may be lazily mounted by Razor)
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        setTimeout(init, 0);
    }
    // Also re-init whenever Blazor re-renders the panel
    var moInit = new MutationObserver(function () {
        if (document.getElementById('ipubPropertiesPanel')) init();
    });
    moInit.observe(document.body, { childList: true, subtree: true });
})();
