// ==============================================================================
// CC TAG PICKER v2 — Injects searchable tag dropdown into Syncfusion's
// built-in Content Control Properties dialog via MutationObserver.
// 405 JATS elements from List_Elements_DB.xlsx Journal sheet.
// ==============================================================================

(function () {
    "use strict";

    // ── Element tag database ─────────────────────────────────────────────
    var _tagDB = [{"t":"access-date","d":"access-date","g":"Back Matter"},{"t":"add-material","d":"add-material","g":"Back Matter"},{"t":"annotation","d":"annotation","g":"Back Matter"},{"t":"app","d":"app","g":"Back Matter"},{"t":"app-group","d":"app-group","g":"Back Matter"},{"t":"atl","d":"article-title","g":"Back Matter"},{"t":"arxiv","d":"arxiv","g":"Back Matter"},{"t":"au;name-style=western","d":"Au_western","g":"Back Matter"},{"t":"au","d":"author","g":"Back Matter"},{"t":"aug","d":"authorgroup","g":"Back Matter"},{"t":"back","d":"back","g":"Back Matter"},{"t":"book","d":"Book","g":"Back Matter"},{"t":"broadcast","d":"broadcast","g":"Back Matter"},{"t":"btl","d":"btl","g":"Back Matter"},{"t":"chapter","d":"Chapter","g":"Back Matter"},{"t":"cty","d":"city","g":"Back Matter"},{"t":"col","d":"collab","g":"Back Matter"},{"t":"communication","d":"Communication","g":"Back Matter"},{"t":"conf-proc","d":"Conf-Proc","g":"Back Matter"},{"t":"ctl","d":"ctl","g":"Back Matter"},{"t":"data","d":"data","g":"Back Matter"},{"t":"dic","d":"date-in-citation","g":"Back Matter"},{"t":"dept","d":"department","g":"Back Matter"},{"t":"discussion","d":"Discussion","g":"Back Matter"},{"t":"doi","d":"doi","g":"Back Matter"},{"t":"ed","d":"ed","g":"Back Matter"},{"t":"edn","d":"edn","g":"Back Matter"},{"t":"eds","d":"eds","g":"Back Matter"},{"t":"element-citation","d":"element-citation","g":"Back Matter"},{"t":"etal","d":"etal","g":"Back Matter"},{"t":"fpg","d":"firstpage","g":"Back Matter"},{"t":"fn","d":"fn","g":"Back Matter"},{"t":"fn-group","d":"fn-group","g":"Back Matter"},{"t":"gov","d":"gov","g":"Back Matter"},{"t":"govt","d":"Gov","g":"Back Matter"},{"t":"instnm","d":"institution-name","g":"Back Matter"},{"t":"isni","d":"isni","g":"Back Matter"},{"t":"iso","d":"iso","g":"Back Matter"},{"t":"item","d":"item","g":"Back Matter"},{"t":"journal","d":"Journal","g":"Back Matter"},{"t":"jtl","d":"jtl","g":"Back Matter"},{"t":"lpg","d":"lastpage","g":"Back Matter"},{"t":"legal","d":"legal","g":"Back Matter"},{"t":"legal-case","d":"legal-case","g":"Back Matter"},{"t":"letter","d":"Letter","g":"Back Matter"},{"t":"level1","d":"level1","g":"Back Matter"},{"t":"level2","d":"level2","g":"Back Matter"},{"t":"level3","d":"level3","g":"Back Matter"},{"t":"level4","d":"level4","g":"Back Matter"},{"t":"milestone-end","d":"milestone-end","g":"Back Matter"},{"t":"milestone-start","d":"milestone-start","g":"Back Matter"},{"t":"misc","d":"misc","g":"Back Matter"},{"t":"mixed-citation","d":"mixed-citation","g":"Back Matter"},{"t":"mnm","d":"mnm","g":"Back Matter"},{"t":"nlm-citation","d":"nlm-citation","g":"Back Matter"},{"t":"nonjournal","d":"nonjournal","g":"Back Matter"},{"t":"note","d":"note","g":"Back Matter"},{"t":"other","d":"Other","g":"Back Matter"},{"t":"pg","d":"page-range","g":"Back Matter"},{"t":"paper","d":"Paper","g":"Back Matter"},{"t":"ptl","d":"part-title","g":"Back Matter"},{"t":"patent","d":"Patent","g":"Back Matter"},{"t":"pnt","d":"patent","g":"Back Matter"},{"t":"ptname","d":"patent-name","g":"Back Matter"},{"t":"person-group","d":"person-group","g":"Back Matter"},{"t":"preprint","d":"preprint","g":"Back Matter"},{"t":"pub-id","d":"pub-id","g":"Back Matter"},{"t":"ref","d":"ref","g":"Back Matter"},{"t":"ref-list","d":"ref-list","g":"Back Matter"},{"t":"report","d":"Report","g":"Back Matter"},{"t":"review","d":"Review","g":"Back Matter"},{"t":"review-meta","d":"review-meta","g":"Back Matter"},{"t":"ringgold","d":"ringgold","g":"Back Matter"},{"t":"ssn","d":"season","g":"Back Matter"},{"t":"series","d":"series","g":"Back Matter"},{"t":"stl","d":"SeriesTitle","g":"Back Matter"},{"t":"size","d":"size","g":"Back Matter"},{"t":"size;specific-use=runing time;units=hours","d":"SizeHours","g":"Back Matter"},{"t":"size;specific-use=runing time;units=minutes","d":"SizeMinutes","g":"Back Matter"},{"t":"size;specific-use=runing time;units=seconds","d":"SizeSecond","g":"Back Matter"},{"t":"social","d":"social","g":"Back Matter"},{"t":"software","d":"software","g":"Back Matter"},{"t":"st","d":"st","g":"Back Matter"},{"t":"std","d":"std","g":"Back Matter"},{"t":"string-conf","d":"string-conf","g":"Back Matter"},{"t":"string-date","d":"string-date","g":"Back Matter"},{"t":"string-name","d":"string-name","g":"Back Matter"},{"t":"thesis","d":"Thesis","g":"Back Matter"},{"t":"time-stamp","d":"time-stamp","g":"Back Matter"},{"t":"trans-source","d":"trans-source","g":"Back Matter"},{"t":"unpublished","d":"unpublished","g":"Back Matter"},{"t":"web","d":"Web","g":"Back Matter"},{"t":"yr","d":"year","g":"Back Matter"},{"t":"alt-text","d":"alt-text","g":"Body Matter"},{"t":"body","d":"body","g":"Body Matter"},{"t":"bold","d":"bold","g":"Body Matter"},{"t":"break","d":"break","g":"Body Matter"},{"t":"chem-struct","d":"chem-struct","g":"Body Matter"},{"t":"chem-struct-wrap","d":"chem-struct-wrap","g":"Body Matter"},{"t":"db-link","d":"db-link","g":"Body Matter"},{"t":"dbond","d":"dbond","g":"Body Matter"},{"t":"def","d":"def","g":"Body Matter"},{"t":"def-head","d":"def-head","g":"Body Matter"},{"t":"def-item","d":"def-item","g":"Body Matter"},{"t":"disp-formula","d":"disp-formula","g":"Body Matter"},{"t":"disp-formula-group","d":"disp-formula-group","g":"Body Matter"},{"t":"disp-quote","d":"disp-quote","g":"Body Matter"},{"t":"genus-species","d":"genus-species","g":"Body Matter"},{"t":"glyph-data","d":"glyph-data","g":"Body Matter"},{"t":"glyph-ref","d":"glyph-ref","g":"Body Matter"},{"t":"hidden","d":"hidden","g":"Body Matter"},{"t":"hr","d":"hr","g":"Body Matter"},{"t":"inline-graphic","d":"inline-graphic","g":"Body Matter"},{"t":"inline-supplementary-material","d":"inline-supplementary-material","g":"Body Matter"},{"t":"italic","d":"italic","g":"Body Matter"},{"t":"keep-together","d":"keep-together","g":"Body Matter"},{"t":"long-desc","d":"long-desc","g":"Body Matter"},{"t":"media","d":"media","g":"Body Matter"},{"t":"mml:math","d":"mml:math","g":"Body Matter"},{"t":"monospace","d":"monospace","g":"Body Matter"},{"t":"object-id","d":"object-id","g":"Body Matter"},{"t":"overline","d":"overline","g":"Body Matter"},{"t":"overline-end","d":"overline-end","g":"Body Matter"},{"t":"overline-start","d":"overline-start","g":"Body Matter"},{"t":"private-char","d":"private-char","g":"Body Matter"},{"t":"ptn","d":"ptn","g":"Body Matter"},{"t":"roman","d":"roman","g":"Body Matter"},{"t":"sans-serif","d":"sans-serif","g":"Body Matter"},{"t":"sbond","d":"sbond","g":"Body Matter"},{"t":"sc","d":"sc","g":"Body Matter"},{"t":"si","d":"si","g":"Body Matter"},{"t":"statement","d":"statement","g":"Body Matter"},{"t":"strike","d":"strike","g":"Body Matter"},{"t":"sub","d":"sub","g":"Body Matter"},{"t":"sup","d":"sup","g":"Body Matter"},{"t":"target","d":"target","g":"Body Matter"},{"t":"tbond","d":"tbond","g":"Body Matter"},{"t":"term","d":"term","g":"Body Matter"},{"t":"term-head","d":"term-head","g":"Body Matter"},{"t":"tex-math","d":"tex-math","g":"Body Matter"},{"t":"textual-form","d":"textual-form","g":"Body Matter"},{"t":"underline","d":"underline","g":"Body Matter"},{"t":"underline-end","d":"underline-end","g":"Body Matter"},{"t":"underline-start","d":"underline-start","g":"Body Matter"},{"t":"unstructured-kwd-group","d":"unstructured-kwd-group","g":"Body Matter"},{"t":"abbrev","d":"abbrev","g":"Common"},{"t":"ack","d":"ack","g":"Common"},{"t":"alternatives","d":"alternatives","g":"Common"},{"t":"aqg","d":"AQ_Group","g":"Common"},{"t":"array","d":"array","g":"Common"},{"t":"article","d":"article","g":"Common"},{"t":"attrib","d":"attrib","g":"Common"},{"t":"aq","d":"author-query","g":"Common"},{"t":"boxed-text","d":"boxed-text","g":"Common"},{"t":"caption","d":"caption","g":"Common"},{"t":"code","d":"code","g":"Common"},{"t":"colgroup","d":"colgroup","g":"Common"},{"t":"cmt","d":"comment","g":"Common"},{"t":"compound-kwd","d":"compound-kwd","g":"Common"},{"t":"compound-kwd-part","d":"compound-kwd-part","g":"Common"},{"t":"date","d":"date","g":"Common"},{"t":"day","d":"day","g":"Common"},{"t":"def-list","d":"def-list","g":"Common"},{"t":"eqn","d":"eqn","g":"Common"},{"t":"fig","d":"fig","g":"Common"},{"t":"fig-group","d":"fig-group","g":"Common"},{"t":"fg","d":"figgroup","g":"Common"},{"t":"floats-group","d":"floats-group","g":"Common"},{"t":"fnref","d":"fnref","g":"Common"},{"t":"front-stub","d":"front-stub","g":"Common"},{"t":"gnm","d":"given-names","g":"Common"},{"t":"glossary","d":"glossary","g":"Common"},{"t":"graphic","d":"graphic","g":"Common"},{"t":"head1","d":"head1","g":"Common"},{"t":"head2","d":"head2","g":"Common"},{"t":"head3","d":"head3","g":"Common"},{"t":"head4","d":"head4","g":"Common"},{"t":"head5","d":"head5","g":"Common"},{"t":"head6","d":"head6","g":"Common"},{"t":"ieqn","d":"ieqn","g":"Common"},{"t":"igALL","d":"igALL","g":"Common"},{"t":"igXML","d":"igXML","g":"Common"},{"t":"inline-formula","d":"inline-formula","g":"Common"},{"t":"INSaiLabel","d":"INSaiLabel","g":"Common"},{"t":"INScolorfiginfo","d":"INScolorfiginfo","g":"Common"},{"t":"INScommon","d":"INScommon","g":"Common"},{"t":"INSenlargethispage","d":"INSenlargethispage","g":"Common"},{"t":"INShskip","d":"INShskip","g":"Common"},{"t":"INShspace","d":"INShspace","g":"Common"},{"t":"INSjournalmonth","d":"INSjournalmonth","g":"Common"},{"t":"INSmonthprint","d":"INSmonthprint","g":"Common"},{"t":"INSnotename","d":"INSnotename","g":"Common"},{"t":"INSsetcounter","d":"INSsetcounter","g":"Common"},{"t":"INSspanrule","d":"INSspanrule","g":"Common"},{"t":"INStblbelowspace","d":"INStblbelowspace","g":"Common"},{"t":"INStblbrk","d":"INStblbrk","g":"Common"},{"t":"INStxtsuperscript","d":"INStxtsuperscript","g":"Common"},{"t":"INSvskip","d":"INSvskip","g":"Common"},{"t":"INSvspace","d":"INSvspace","g":"Common"},{"t":"INSxmlelement","d":"INSxmlelement","g":"Common"},{"t":"INTb","d":"INTb","g":"Common"},{"t":"INTB","d":"INTB","g":"Common"},{"t":"INTbi","d":"INTbi","g":"Common"},{"t":"INTBI","d":"INTBI","g":"Common"},{"t":"INTbkmerge","d":"INTbkmerge","g":"Common"},{"t":"INTcolspec","d":"INTcolspec","g":"Common"},{"t":"INTcontfig","d":"INTcontfig","g":"Common"},{"t":"INTdup","d":"INTdup","g":"Common"},{"t":"INTenditem","d":"INTenditem","g":"Common"},{"t":"INTendlabel","d":"INTendlabel","g":"Common"},{"t":"INTentry","d":"INTentry","g":"Common"},{"t":"INTfloatid","d":"INTfloatid","g":"Common"},{"t":"INTfootitem","d":"INTfootitem","g":"Common"},{"t":"INTfootlabel","d":"INTfootlabel","g":"Common"},{"t":"INThbox","d":"INThbox","g":"Common"},{"t":"INThspace","d":"INThspace","g":"Common"},{"t":"INTIT","d":"INTIT","g":"Common"},{"t":"INTit","d":"INTit","g":"Common"},{"t":"INTLC","d":"INTLC","g":"Common"},{"t":"INTmerge","d":"INTmerge","g":"Common"},{"t":"INTpicture","d":"INTpicture","g":"Common"},{"t":"INTqauthor","d":"INTqauthor","g":"Common"},{"t":"INTqpage","d":"INTqpage","g":"Common"},{"t":"INTqstitle","d":"INTqstitle","g":"Common"},{"t":"INTqtitle","d":"INTqtitle","g":"Common"},{"t":"INTqtoc","d":"INTqtoc","g":"Common"},{"t":"INTqtocauthor","d":"INTqtocauthor","g":"Common"},{"t":"INTqtoctitle","d":"INTqtoctitle","g":"Common"},{"t":"INTRM","d":"INTRM","g":"Common"},{"t":"INTrm","d":"INTrm","g":"Common"},{"t":"INTrotate","d":"INTrotate","g":"Common"},{"t":"INTSC","d":"INTSC","g":"Common"},{"t":"INTswapfile","d":"INTswapfile","g":"Common"},{"t":"INTswaptext","d":"INTswaptext","g":"Common"},{"t":"INTtabimage","d":"INTtabimage","g":"Common"},{"t":"INTtrack","d":"INTtrack","g":"Common"},{"t":"INTUP","d":"INTUP","g":"Common"},{"t":"INTvspace","d":"INTvspace","g":"Common"},{"t":"INTxmlentity","d":"INTxmlentity","g":"Common"},{"t":"jel","d":"jel","g":"Common"},{"t":"journaltitle","d":"journaltitle","g":"Common"},{"t":"kwd","d":"kwd","g":"Common"},{"t":"kwd-group","d":"kwd-group","g":"Common"},{"t":"lbl","d":"label","g":"Common"},{"t":"list","d":"list","g":"Common"},{"t":"list-item","d":"list-item","g":"Common"},{"t":"listtitle","d":"listtitle","g":"Common"},{"t":"lrh","d":"lrh","g":"Common"},{"t":"mth","d":"month","g":"Common"},{"t":"name","d":"name","g":"Common"},{"t":"named-content","d":"named-content","g":"Common"},{"t":"notes","d":"notes","g":"Common"},{"t":"orcid","d":"orcid","g":"Common"},{"t":"p","d":"p","g":"Common"},{"t":"particle","d":"particle","g":"Common"},{"t":"pref","d":"prefix","g":"Common"},{"t":"preformat","d":"preformat","g":"Common"},{"t":"response","d":"response","g":"Common"},{"t":"rrh","d":"rrh","g":"Common"},{"t":"sec","d":"sec","g":"Common"},{"t":"sec-meta","d":"sec-meta","g":"Common"},{"t":"shorttitle","d":"shorttitle","g":"Common"},{"t":"sig","d":"sig","g":"Common"},{"t":"sig-block","d":"sig-block","g":"Common"},{"t":"src","d":"source","g":"Common"},{"t":"speaker","d":"speaker","g":"Common"},{"t":"speech","d":"speech","g":"Common"},{"t":"srh","d":"srh","g":"Common"},{"t":"styled-content","d":"styled-content","g":"Common"},{"t":"sub-article","d":"sub-article","g":"Common"},{"t":"subtitle","d":"subtitle","g":"Common"},{"t":"suff","d":"suffix","g":"Common"},{"t":"supertitle","d":"supertitle","g":"Common"},{"t":"snm","d":"surname","g":"Common"},{"t":"table","d":"table","g":"Common"},{"t":"table-wrap","d":"table-wrap","g":"Common"},{"t":"table-wrap-foot","d":"table-wrap-foot","g":"Common"},{"t":"table-wrap-group","d":"table-wrap-group","g":"Common"},{"t":"tg","d":"tablegroup","g":"Common"},{"t":"tbody","d":"tbody","g":"Common"},{"t":"td","d":"td","g":"Common"},{"t":"tfoot","d":"tfoot","g":"Common"},{"t":"th","d":"th","g":"Common"},{"t":"thead","d":"thead","g":"Common"},{"t":"title","d":"title","g":"Common"},{"t":"tr","d":"tr","g":"Common"},{"t":"ueqn","d":"ueqn","g":"Common"},{"t":"uri","d":"uri","g":"Common"},{"t":"verse-group","d":"verse-group","g":"Common"},{"t":"verse-line","d":"verse-line","g":"Common"},{"t":"vol","d":"volume","g":"Common"},{"t":"volume-id","d":"volume-id","g":"Common"},{"t":"volume-series","d":"volume-series","g":"Common"},{"t":"xref","d":"xref","g":"Common"},{"t":"abbrev-journal-title","d":"abbrev-journal-title","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"abstract","d":"abstract","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"adrl","d":"addr-line","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"address","d":"address","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"aff","d":"aff","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"ag","d":"aff-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"alt-title","d":"alt-title","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"anon","d":"anonymous","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"article-categories","d":"article-categories","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"article-id","d":"article-id","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"article-meta","d":"article-meta","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"author-comment","d":"author-comment","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"author-notes","d":"author-notes","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"award-group","d":"award-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"award-id","d":"award-id","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"bibdataset","d":"bibdataset","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"bibissue","d":"bibissue","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"bio","d":"bio","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"Category","d":"Category","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"conf-acronym","d":"conf-acronym","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"cfdate","d":"conf-date","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"cfloc","d":"conf-loc","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"cfname","d":"conf-name","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"conf-num","d":"conf-num","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"conf-sponsor","d":"conf-sponsor","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"conf-theme","d":"conf-theme","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"cftitle","d":"conf-title","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"conference","d":"conference","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"contrib","d":"contrib","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"contrib-group","d":"contrib-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"copyright-holder","d":"copyright-holder","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"copyright-statement","d":"copyright-statement","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"copyright-year","d":"copyright-year","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"corresp","d":"corresp","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"cnt","d":"country","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"counts","d":"counts","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"custom-meta","d":"custom-meta","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"custom-meta-group","d":"custom-meta-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"datasettitle","d":"datasettitle","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"deg","d":"degrees","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"elocation-id","d":"elocation-id","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"email","d":"email","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"equation-count","d":"equation-count","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"ext-link","d":"ext-link","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"fax","d":"fax","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"fig-count","d":"fig-count","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"front","d":"front","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"funding-group","d":"funding-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"funding-source","d":"funding-source","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"funding-statement","d":"funding-statement","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"history","d":"history","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"index-div","d":"index-div","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"index-entry","d":"index-entry","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"inst","d":"institution","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"isbn","d":"isbn","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"issn","d":"issn","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"iss","d":"issue","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"issue-id","d":"issue-id","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"issue-part","d":"issue-part","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"issue-sponsor","d":"issue-sponsor","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"itl","d":"issue-title","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"journal-id","d":"journal-id","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"journal-meta","d":"journal-meta","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"journal-subtitle","d":"journal-subtitle","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"journal-title-group","d":"journal-title-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"license","d":"license","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"license-p","d":"license-p","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"meta-name","d":"meta-name","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"meta-value","d":"meta-value","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"Num-p","d":"Num-p","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"on-behalf-of","d":"on-behalf-of","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"open-access","d":"open-access","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"p-unindent","d":"p-unindent","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"page-count","d":"page-count","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"permissions","d":"permissions","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"phone","d":"phone","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"PMI","d":"PMI","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"pcode","d":"postal-code","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"pbox","d":"postbox","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"price","d":"price","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"principal-award-recipient","d":"principal-award-recipient","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"principal-investigator","d":"principal-investigator","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"product","d":"product","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"pub-date","d":"pub-date","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"pub","d":"publisher","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"loc","d":"publisher-loc","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"ref-count","d":"ref-count","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"related-article","d":"related-article","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"related-object","d":"related-object","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"role","d":"role","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"see","d":"see","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"see-also","d":"see-also","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"self-uri","d":"self-uri","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"series-text","d":"series-text","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"series-title","d":"series-title","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"subj-group","d":"subj-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"subject","d":"subject","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"supplement","d":"supplement","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"supplementary-material","d":"supplementary-material","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"table-count","d":"table-count","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"title-group","d":"title-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"toc","d":"toc","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"toc-entry","d":"toc-entry","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"trans-abstract","d":"trans-abstract","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"trans-kwd-group","d":"trans-kwd-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"trans-subtitle","d":"trans-subtitle","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"trans-title","d":"trans-title","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"trans-title-group","d":"trans-title-group","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"translation","d":"translation","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"Un_List","d":"Un_List","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"},{"t":"word-count","d":"word-count","g":"Journal Metadata PLACEHOLDER_TAGS Front Matter"}];

    // ── Group the tags ───────────────────────────────────────────────────
    var _grouped = {};
    for (var i = 0; i < _tagDB.length; i++) {
        var g = _tagDB[i].g;
        if (!_grouped[g]) _grouped[g] = [];
        _grouped[g].push(_tagDB[i]);
    }
    var _groupNames = Object.keys(_grouped).sort();

    // Track state
    var _injected = false;
    var _dropdownEl = null;

    // ── MutationObserver: watch for Syncfusion CC Properties dialog ──────
    function startWatching() {
        var observer = new MutationObserver(function (mutations) {
            for (var m = 0; m < mutations.length; m++) {
                for (var n = 0; n < mutations[m].addedNodes.length; n++) {
                    var node = mutations[m].addedNodes[n];
                    if (node.nodeType !== 1) continue;
                    checkForCCDialog(node);
                }
            }
        });
        observer.observe(document.body, { childList: true, subtree: true });
        console.log("[CCPicker] MutationObserver active, watching for CC dialog (" + _tagDB.length + " tags)");
    }

    function checkForCCDialog(node) {
        // Syncfusion CC dialog has title "Content Control Properties"
        var titleEl = node.querySelector
            ? node.querySelector("#" + CSS.escape(node.id) + "_title, .e-dlg-header")
            : null;

        if (!titleEl) {
            // node itself might be inside the dialog
            var dialogs = document.querySelectorAll(".e-de-dlg-target.e-popup-open");
            for (var d = 0; d < dialogs.length; d++) {
                var hdr = dialogs[d].querySelector(".e-dlg-header");
                if (hdr && hdr.textContent.trim() === "Content Control Properties") {
                    injectDropdown(dialogs[d]);
                    return;
                }
            }
            return;
        }

        if (titleEl.textContent.trim() === "Content Control Properties") {
            injectDropdown(node);
        }
    }

    // Also poll on attribute changes (dialog visibility)
    function startDialogPoll() {
        var pollObserver = new MutationObserver(function () {
            if (_injected) return;
            var dialogs = document.querySelectorAll(".e-de-dlg-target");
            for (var d = 0; d < dialogs.length; d++) {
                var hdr = dialogs[d].querySelector(".e-dlg-header");
                if (hdr && hdr.textContent.trim() === "Content Control Properties" &&
                    dialogs[d].classList.contains("e-popup-open")) {
                    injectDropdown(dialogs[d]);
                    return;
                }
            }
        });
        pollObserver.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ["class", "style"] });
    }

    // ── Inject dropdown into the Tag input field ─────────────────────────
    function injectDropdown(dialogEl) {
        if (_injected) return;

        // Find the Tag input — it's the second textbox (Title is first)
        var inputs = dialogEl.querySelectorAll("input.e-textbox");
        var tagInput = null;
        var titleInput = null;

        for (var i = 0; i < inputs.length; i++) {
            var label = inputs[i].closest(".e-float-input");
            if (!label) continue;
            var labelText = label.querySelector(".e-float-text");
            if (!labelText) continue;
            var lt = labelText.textContent.trim().toLowerCase();
            if (lt === "tag") tagInput = inputs[i];
            if (lt === "title") titleInput = inputs[i];
        }

        if (!tagInput) {
            console.warn("[CCPicker] Tag input not found in dialog");
            return;
        }

        _injected = true;
        console.log("[CCPicker] Injecting dropdown into CC Properties dialog");

        // Create dropdown container
        var wrapper = tagInput.closest(".e-float-input") || tagInput.parentElement;
        wrapper.style.position = "relative";

        _dropdownEl = document.createElement("div");
        _dropdownEl.className = "ccp-dropdown";
        _dropdownEl.id = "ccpDropdownInject";
        wrapper.appendChild(_dropdownEl);

        // Render on focus/input
        tagInput.addEventListener("focus", function () {
            renderDropdown(tagInput.value, tagInput, titleInput);
        });

        tagInput.addEventListener("input", function () {
            renderDropdown(tagInput.value, tagInput, titleInput);
        });

        // Close dropdown when clicking outside
        document.addEventListener("mousedown", function closeDDHandler(e) {
            if (!_dropdownEl) { document.removeEventListener("mousedown", closeDDHandler); return; }
            if (!wrapper.contains(e.target) && !_dropdownEl.contains(e.target)) {
                _dropdownEl.style.display = "none";
            }
        });

        // Watch for dialog close to reset state
        var closeObserver = new MutationObserver(function () {
            if (!dialogEl.classList.contains("e-popup-open")) {
                _injected = false;
                _dropdownEl = null;
                closeObserver.disconnect();
                console.log("[CCPicker] Dialog closed, ready for next open");
            }
        });
        closeObserver.observe(dialogEl, { attributes: true, attributeFilter: ["class"] });

        // Also handle the dialog close button and OK/Cancel
        var closeBtn = dialogEl.querySelector(".e-dlg-closeicon-btn");
        var okBtn = dialogEl.querySelector(".e-para-okay");
        var cancelBtn = dialogEl.querySelector(".e-para-cancel");
        var resetOnClose = function () {
            setTimeout(function () { _injected = false; _dropdownEl = null; }, 100);
        };
        if (closeBtn) closeBtn.addEventListener("click", resetOnClose);
        if (okBtn) okBtn.addEventListener("click", resetOnClose);
        if (cancelBtn) cancelBtn.addEventListener("click", resetOnClose);
    }

    // ── Render the dropdown list ─────────────────────────────────────────
    function renderDropdown(filter, tagInput, titleInput) {
        if (!_dropdownEl) return;

        var html = "";
        var f = (filter || "").toLowerCase();
        var count = 0;
        var maxItems = 150;

        for (var gi = 0; gi < _groupNames.length; gi++) {
            var gName = _groupNames[gi];
            var items = _grouped[gName];
            var matchedItems = [];

            for (var ii = 0; ii < items.length; ii++) {
                if (count >= maxItems) break;
                var item = items[ii];
                if (!f || item.t.toLowerCase().indexOf(f) !== -1 || item.d.toLowerCase().indexOf(f) !== -1) {
                    matchedItems.push(item);
                    count++;
                }
            }

            if (matchedItems.length > 0) {
                html += '<div class="ccp-dd-group">' + esc(gName) + '</div>';
                for (var mi = 0; mi < matchedItems.length; mi++) {
                    var m = matchedItems[mi];
                    var desc = m.d !== m.t ? ' <span class="ccp-dd-desc">' + esc(m.d) + '</span>' : '';
                    html += '<div class="ccp-dd-item" data-tag="' + esc(m.t) + '" data-display="' + esc(m.d) + '">' +
                        '<span class="ccp-dd-tag">' + esc(m.t) + '</span>' + desc + '</div>';
                }
            }
        }

        if (count === 0) {
            html = '<div class="ccp-dd-empty">No matching tags found</div>';
        }

        _dropdownEl.innerHTML = html;
        _dropdownEl.style.display = "block";

        // Click handlers on items
        var ddItems = _dropdownEl.querySelectorAll(".ccp-dd-item");
        for (var i = 0; i < ddItems.length; i++) {
            ddItems[i].addEventListener("mousedown", function (e) {
                e.preventDefault(); // prevent blur
                var tag = this.getAttribute("data-tag");

                // Set tag input value — use Syncfusion's method
                setInputValue(tagInput, tag);

                // Auto-fill title too
                if (titleInput) {
                    setInputValue(titleInput, tag);
                }

                _dropdownEl.style.display = "none";
                console.log("[CCPicker] Selected tag: " + tag);
            });
        }
    }

    // ── Set value on a Syncfusion EJ2 TextBox properly ───────────────────
    function setInputValue(input, value) {
        // Set the native value
        var nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
        nativeSetter.call(input, value);

        // Fire events so Syncfusion picks up the change
        input.dispatchEvent(new Event("input", { bubbles: true }));
        input.dispatchEvent(new Event("change", { bubbles: true }));

        // Also try the EJ2 instance directly
        var ejInst = input.ej2_instances;
        if (ejInst && ejInst.length > 0) {
            ejInst[0].value = value;
            if (ejInst[0].dataBind) ejInst[0].dataBind();
        }

        // Ensure float label moves up
        var wrapper = input.closest(".e-float-input");
        if (wrapper) {
            var label = wrapper.querySelector(".e-float-text");
            if (label) {
                label.classList.remove("e-label-bottom");
                label.classList.add("e-label-top");
            }
            wrapper.classList.add("e-valid-input");
        }
    }

    function esc(s) {
        if (!s) return "";
        return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    }

    // ── Expose DB ────────────────────────────────────────────────────────
    window.getCCTagDB = function () { return _tagDB; };

    // ── Initialize ───────────────────────────────────────────────────────
    if (document.readyState === "complete") {
        startWatching();
        startDialogPoll();
    } else {
        window.addEventListener("load", function () {
            startWatching();
            startDialogPoll();
        });
    }

})();