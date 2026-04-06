// ==============================================================================
// SCHEMA PARSER v8 — Dynamic XSD + Content Control Tag Matching
// - Reads ANY XSD schema (journal, book, etc.) to generate fields
// - 455-entry JATS tag lookup for display names and grouping
// - CC Tag matching via Syncfusion API + SFDT fallback
// - Dynamic field creation for unmatched CC tags
// ==============================================================================

(function () {
    "use strict";

    var _schemaTree = null;

    // ── JATS tag lookup: normalizedTag → { d: displayName, g: group } ────
    var _tagLookup = {"copyrightstatement":{"d":"copyright-statement","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"copyrightholder":{"d":"copyright-holder","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"copyrightyear":{"d":"copyright-year","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"license":{"d":"license","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"licensep":{"d":"license-p","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"mmlmath":{"d":"mml:math","g":"Body Matter"},"permissions":{"d":"permissions","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"atl":{"d":"article-title","g":"Back Matter"},"articletitle":{"d":"article-title","g":"Back Matter"},"aff":{"d":"aff","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"col":{"d":"collab","g":"Back Matter"},"collab":{"d":"collab","g":"Back Matter"},"cfloc":{"d":"conf-loc","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"confloc":{"d":"conf-loc","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"cfname":{"d":"conf-name","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"confname":{"d":"conf-name","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"objectid":{"d":"object-id","g":"Body Matter"},"isbn":{"d":"isbn","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"issn":{"d":"issn","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"iss":{"d":"issue","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"issue":{"d":"issue","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"issueid":{"d":"issue-id","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"issuepart":{"d":"issue-part","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"issuesponsor":{"d":"issue-sponsor","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"itl":{"d":"issue-title","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"issuetitle":{"d":"issue-title","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"journalid":{"d":"journal-id","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"role":{"d":"role","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"transtitlegroup":{"d":"trans-title-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"transsubtitle":{"d":"trans-subtitle","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"transtitle":{"d":"trans-title","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"vol":{"d":"volume","g":"Common"},"volume":{"d":"volume","g":"Common"},"volumeid":{"d":"volume-id","g":"Common"},"volumeseries":{"d":"volume-series","g":"Common"},"app":{"d":"app","g":"Back Matter"},"pub":{"d":"publisher","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"publisher":{"d":"publisher","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"publishername":{"d":"publisher-name","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"loc":{"d":"publisher-loc","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"publisherloc":{"d":"publisher-loc","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"fpg":{"d":"firstpage","g":"Back Matter"},"firstpage":{"d":"firstpage","g":"Back Matter"},"lpg":{"d":"lastpage","g":"Back Matter"},"lastpage":{"d":"lastpage","g":"Back Matter"},"pg":{"d":"page-range","g":"Back Matter"},"pagerange":{"d":"page-range","g":"Back Matter"},"size":{"d":"size","g":"Back Matter"},"elocationid":{"d":"elocation-id","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"mixedcitation":{"d":"mixed-citation","g":"Back Matter"},"elementcitation":{"d":"element-citation","g":"Back Matter"},"address":{"d":"address","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"adrl":{"d":"addr-line","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"addrline":{"d":"addr-line","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"cnt":{"d":"country","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"country":{"d":"country","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"email":{"d":"email","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"fax":{"d":"fax","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"inst":{"d":"institution","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"institution":{"d":"institution","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"phone":{"d":"phone","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"uri":{"d":"uri","g":"Common"},"date":{"d":"date","g":"Common"},"ssn":{"d":"season","g":"Back Matter"},"season":{"d":"season","g":"Back Matter"},"yr":{"d":"year","g":"Back Matter"},"year":{"d":"year","g":"Back Matter"},"stringdate":{"d":"string-date","g":"Back Matter"},"stringname":{"d":"string-name","g":"Back Matter"},"name":{"d":"name","g":"Common"},"snm":{"d":"surname","g":"Common"},"surname":{"d":"surname","g":"Common"},"gnm":{"d":"given-names","g":"Common"},"givennames":{"d":"given-names","g":"Common"},"pref":{"d":"prefix","g":"Common"},"prefix":{"d":"prefix","g":"Common"},"suff":{"d":"suffix","g":"Common"},"suffix":{"d":"suffix","g":"Common"},"extlink":{"d":"ext-link","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"attrib":{"d":"attrib","g":"Common"},"def":{"d":"def","g":"Body Matter"},"lbl":{"d":"label","g":"Common"},"label":{"d":"label","g":"Common"},"price":{"d":"price","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"title":{"d":"title","g":"Common"},"relatedarticle":{"d":"related-article","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"journal":{"d":"Journal","g":"Back Matter"},"book":{"d":"Book","g":"Back Matter"},"other":{"d":"Other","g":"Back Matter"},"web":{"d":"Web","g":"Back Matter"},"confproc":{"d":"Conf-Proc","g":"Back Matter"},"ack":{"d":"ack","g":"Common"},"bio":{"d":"bio","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"notes":{"d":"notes","g":"Common"},"alttext":{"d":"alt-text","g":"Body Matter"},"longdesc":{"d":"long-desc","g":"Body Matter"},"custommetagroup":{"d":"custom-meta-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"custommeta":{"d":"custom-meta","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"metaname":{"d":"meta-name","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"metavalue":{"d":"meta-value","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"alternatives":{"d":"alternatives","g":"Common"},"textualform":{"d":"textual-form","g":"Body Matter"},"subjgroup":{"d":"subj-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"subject":{"d":"subject","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"seriestitle":{"d":"series-title","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"btl":{"d":"btl","g":"Back Matter"},"authornotes":{"d":"author-notes","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"product":{"d":"product","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"history":{"d":"history","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"selfuri":{"d":"self-uri","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"abstract":{"d":"abstract","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"jtl":{"d":"jtl","g":"Back Matter"},"kwdgroup":{"d":"kwd-group","g":"Common"},"kwd":{"d":"kwd","g":"Common"},"compoundkwd":{"d":"compound-kwd","g":"Common"},"compoundkwdpart":{"d":"compound-kwd-part","g":"Common"},"unstructuredkwdgroup":{"d":"unstructured-kwd-group","g":"Body Matter"},"corresp":{"d":"corresp","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"pubdate":{"d":"pub-date","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"conference":{"d":"conference","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"confacronym":{"d":"conf-acronym","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"confnum":{"d":"conf-num","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"confsponsor":{"d":"conf-sponsor","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"conftheme":{"d":"conf-theme","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"stringconf":{"d":"string-conf","g":"Back Matter"},"counts":{"d":"counts","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"equationcount":{"d":"equation-count","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"figcount":{"d":"fig-count","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"ctl":{"d":"ctl","g":"Back Matter"},"pagecount":{"d":"page-count","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"wordcount":{"d":"word-count","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"titlegroup":{"d":"title-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"subtitle":{"d":"subtitle","g":"Common"},"contribgroup":{"d":"contrib-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"contrib":{"d":"contrib","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"deg":{"d":"degrees","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"degrees":{"d":"degrees","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"onbehalfof":{"d":"on-behalf-of","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"authorcomment":{"d":"author-comment","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"edn":{"d":"edn","g":"Back Matter"},"etal":{"d":"etal","g":"Back Matter"},"boxedtext":{"d":"boxed-text","g":"Common"},"chemstructwrap":{"d":"chem-struct-wrap","g":"Body Matter"},"chemstruct":{"d":"chem-struct","g":"Body Matter"},"chapter":{"d":"Chapter","g":"Back Matter"},"fig":{"d":"fig","g":"Common"},"caption":{"d":"caption","g":"Common"},"graphic":{"d":"graphic","g":"Common"},"media":{"d":"media","g":"Body Matter"},"inlinegraphic":{"d":"inline-graphic","g":"Body Matter"},"preformat":{"d":"preformat","g":"Common"},"supplementarymaterial":{"d":"supplementary-material","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"tablewrapgroup":{"d":"table-wrap-group","g":"Common"},"tablewrap":{"d":"table-wrap","g":"Common"},"tablewrapfoot":{"d":"table-wrap-foot","g":"Common"},"hr":{"d":"hr","g":"Body Matter"},"break":{"d":"break","g":"Body Matter"},"bold":{"d":"bold","g":"Body Matter"},"italic":{"d":"italic","g":"Body Matter"},"monospace":{"d":"monospace","g":"Body Matter"},"roman":{"d":"roman","g":"Body Matter"},"sansserif":{"d":"sans-serif","g":"Body Matter"},"sc":{"d":"sc","g":"Body Matter"},"overline":{"d":"overline","g":"Body Matter"},"strike":{"d":"strike","g":"Body Matter"},"sub":{"d":"sub","g":"Body Matter"},"sup":{"d":"sup","g":"Body Matter"},"underline":{"d":"underline","g":"Body Matter"},"overlinestart":{"d":"overline-start","g":"Body Matter"},"overlineend":{"d":"overline-end","g":"Body Matter"},"underlinestart":{"d":"underline-start","g":"Body Matter"},"underlineend":{"d":"underline-end","g":"Body Matter"},"fundinggroup":{"d":"funding-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"fundingstatement":{"d":"funding-statement","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"openaccess":{"d":"open-access","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"awardgroup":{"d":"award-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"fundingsource":{"d":"funding-source","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"awardid":{"d":"award-id","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"principalawardrecipient":{"d":"principal-award-recipient","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"principalinvestigator":{"d":"principal-investigator","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"journalmeta":{"d":"journal-meta","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"journaltitlegroup":{"d":"journal-title-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"journaltitle":{"d":"journal-title","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"journalsubtitle":{"d":"journal-subtitle","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"abbrevjournaltitle":{"d":"abbrev-journal-title","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"fn":{"d":"fn","g":"Back Matter"},"target":{"d":"target","g":"Body Matter"},"xref":{"d":"xref","g":"Common"},"inlinesupplementarymaterial":{"d":"inline-supplementary-material","g":"Body Matter"},"deflist":{"d":"def-list","g":"Common"},"thesis":{"d":"Thesis","g":"Back Matter"},"report":{"d":"Report","g":"Back Matter"},"paper":{"d":"Paper","g":"Back Matter"},"discussion":{"d":"Discussion","g":"Back Matter"},"patent":{"d":"Patent","g":"Back Matter"},"communication":{"d":"Communication","g":"Back Matter"},"list":{"d":"list","g":"Common"},"letter":{"d":"Letter","g":"Back Matter"},"inlineformula":{"d":"inline-formula","g":"Common"},"review":{"d":"Review","g":"Back Matter"},"dispformulagroup":{"d":"disp-formula-group","g":"Body Matter"},"nlmcitation":{"d":"nlm-citation","g":"Back Matter"},"p":{"d":"p","g":"Common"},"dispquote":{"d":"disp-quote","g":"Body Matter"},"speech":{"d":"speech","g":"Common"},"speaker":{"d":"speaker","g":"Common"},"statement":{"d":"statement","g":"Body Matter"},"versegroup":{"d":"verse-group","g":"Common"},"verseline":{"d":"verse-line","g":"Common"},"abbrev":{"d":"abbrev","g":"Common"},"milestonestart":{"d":"milestone-start","g":"Back Matter"},"milestoneend":{"d":"milestone-end","g":"Back Matter"},"namedcontent":{"d":"named-content","g":"Common"},"styledcontent":{"d":"styled-content","g":"Common"},"reflist":{"d":"ref-list","g":"Back Matter"},"ref":{"d":"ref","g":"Back Matter"},"note":{"d":"note","g":"Back Matter"},"accessdate":{"d":"access-date","g":"Back Matter"},"annotation":{"d":"annotation","g":"Back Matter"},"chaptertitle":{"d":"chapter-title","g":"Back Matter"},"cmt":{"d":"comment","g":"Common"},"comment":{"d":"comment","g":"Common"},"dic":{"d":"date-in-citation","g":"Back Matter"},"dateincitation":{"d":"date-in-citation","g":"Back Matter"},"edition":{"d":"edition","g":"Back Matter"},"gov":{"d":"gov","g":"Back Matter"},"ptl":{"d":"part-title","g":"Back Matter"},"parttitle":{"d":"part-title","g":"Back Matter"},"pnt":{"d":"patent","g":"Back Matter"},"persongroup":{"d":"person-group","g":"Back Matter"},"pubid":{"d":"pub-id","g":"Back Matter"},"series":{"d":"series","g":"Back Matter"},"std":{"d":"std","g":"Back Matter"},"src":{"d":"source","g":"Common"},"source":{"d":"source","g":"Common"},"timestamp":{"d":"time-stamp","g":"Back Matter"},"transsource":{"d":"trans-source","g":"Back Matter"},"relatedobject":{"d":"related-object","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"floatsgroup":{"d":"floats-group","g":"Common"},"sec":{"d":"sec","g":"Common"},"secmeta":{"d":"sec-meta","g":"Common"},"table":{"d":"table","g":"Common"},"thead":{"d":"thead","g":"Common"},"tfoot":{"d":"tfoot","g":"Common"},"tbody":{"d":"tbody","g":"Common"},"colgroup":{"d":"colgroup","g":"Common"},"tr":{"d":"tr","g":"Common"},"th":{"d":"th","g":"Common"},"td":{"d":"td","g":"Common"},"privatechar":{"d":"private-char","g":"Body Matter"},"glyphdata":{"d":"glyph-data","g":"Body Matter"},"glyphref":{"d":"glyph-ref","g":"Body Matter"},"article":{"d":"article","g":"Common"},"front":{"d":"front","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"body":{"d":"body","g":"Body Matter"},"back":{"d":"back","g":"Back Matter"},"subarticle":{"d":"sub-article","g":"Common"},"frontstub":{"d":"front-stub","g":"Common"},"response":{"d":"response","g":"Common"},"articlemeta":{"d":"article-meta","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"articlecategories":{"d":"article-categories","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"articleid":{"d":"article-id","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"anon":{"d":"anonymous","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"anonymous":{"d":"anonymous","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"seriestext":{"d":"series-text","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"transabstract":{"d":"trans-abstract","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"tablecount":{"d":"table-count","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"refcount":{"d":"ref-count","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"alttitle":{"d":"alt-title","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"cty":{"d":"city","g":"Back Matter"},"city":{"d":"city","g":"Back Matter"},"st":{"d":"st","g":"Back Matter"},"ed":{"d":"ed","g":"Back Matter"},"supplement":{"d":"supplement","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"appgroup":{"d":"app-group","g":"Back Matter"},"intvspace":{"d":"INTvspace","g":"Common"},"fngroup":{"d":"fn-group","g":"Back Matter"},"glossary":{"d":"glossary","g":"Common"},"array":{"d":"array","g":"Common"},"sigblock":{"d":"sig-block","g":"Common"},"au":{"d":"author","g":"Back Matter"},"author":{"d":"author","g":"Back Matter"},"aug":{"d":"authorgroup","g":"Back Matter"},"authorgroup":{"d":"authorgroup","g":"Back Matter"},"booktitle":{"d":"booktitle","g":"Back Matter"},"page":{"d":"page","g":"Back Matter"},"eds":{"d":"eds","g":"Back Matter"},"state":{"d":"state","g":"Back Matter"},"sig":{"d":"sig","g":"Common"},"termhead":{"d":"term-head","g":"Body Matter"},"defhead":{"d":"def-head","g":"Body Matter"},"defitem":{"d":"def-item","g":"Body Matter"},"term":{"d":"term","g":"Body Matter"},"listitem":{"d":"list-item","g":"Common"},"dispformula":{"d":"disp-formula","g":"Body Matter"},"texmath":{"d":"tex-math","g":"Body Matter"},"figgroup":{"d":"fig-group","g":"Common"},"level1":{"d":"level1","g":"Back Matter"},"level2":{"d":"level2","g":"Back Matter"},"level3":{"d":"level3","g":"Back Matter"},"level4":{"d":"level4","g":"Back Matter"},"item":{"d":"item","g":"Back Matter"},"editors":{"d":"editors","g":"Back Matter"},"dept":{"d":"department","g":"Back Matter"},"department":{"d":"department","g":"Back Matter"},"instnm":{"d":"institution-name","g":"Back Matter"},"institutionname":{"d":"institution-name","g":"Back Matter"},"pcode":{"d":"postal-code","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"postalcode":{"d":"postal-code","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"ag":{"d":"aff-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"affgroup":{"d":"aff-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"fg":{"d":"figgroup","g":"Common"},"tg":{"d":"tablegroup","g":"Common"},"tablegroup":{"d":"tablegroup","g":"Common"},"aq":{"d":"author-query","g":"Common"},"authorquery":{"d":"author-query","g":"Common"},"aqg":{"d":"AQ_Group","g":"Common"},"aqgroup":{"d":"AQ_Group","g":"Common"},"eqn":{"d":"eqn","g":"Common"},"ueqn":{"d":"ueqn","g":"Common"},"ieqn":{"d":"ieqn","g":"Common"},"igxml":{"d":"igXML","g":"Common"},"igall":{"d":"igALL","g":"Common"},"translation":{"d":"translation","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"misc":{"d":"misc","g":"Back Matter"},"supertitle":{"d":"supertitle","g":"Common"},"doi":{"d":"doi","g":"Back Matter"},"cfdate":{"d":"conf-date","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"confdate":{"d":"conf-date","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"ptname":{"d":"patent-name","g":"Back Matter"},"patentname":{"d":"patent-name","g":"Back Matter"},"cftitle":{"d":"conf-title","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"conftitle":{"d":"conf-title","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"intentry":{"d":"INTentry","g":"Common"},"mnm":{"d":"mnm","g":"Back Matter"},"lrh":{"d":"lrh","g":"Common"},"rrh":{"d":"rrh","g":"Common"},"day":{"d":"day","g":"Common"},"intmerge":{"d":"INTmerge","g":"Common"},"intqpage":{"d":"INTqpage","g":"Common"},"intqstitle":{"d":"INTqstitle","g":"Common"},"intqtitle":{"d":"INTqtitle","g":"Common"},"intqauthor":{"d":"INTqauthor","g":"Common"},"intqtoc":{"d":"INTqtoc","g":"Common"},"intqtoctitle":{"d":"INTqtoctitle","g":"Common"},"intqtocauthor":{"d":"INTqtocauthor","g":"Common"},"intdup":{"d":"INTdup","g":"Common"},"intfloatid":{"d":"INTfloatid","g":"Common"},"intendlabel":{"d":"INTendlabel","g":"Common"},"intenditem":{"d":"INTenditem","g":"Common"},"intfootlabel":{"d":"INTfootlabel","g":"Common"},"mth":{"d":"month","g":"Common"},"month":{"d":"month","g":"Common"},"intfootitem":{"d":"INTfootitem","g":"Common"},"intswaptext":{"d":"INTswaptext","g":"Common"},"intswapfile":{"d":"INTswapfile","g":"Common"},"intbkmerge":{"d":"INTbkmerge","g":"Common"},"inttrack":{"d":"INTtrack","g":"Common"},"intpicture":{"d":"INTpicture","g":"Common"},"intup":{"d":"INTUP","g":"Common"},"intsc":{"d":"INTSC","g":"Common"},"intlc":{"d":"INTLC","g":"Common"},"intrm":{"d":"INTRM","g":"Common"},"intb":{"d":"INTb","g":"Common"},"intbi":{"d":"INTbi","g":"Common"},"introtate":{"d":"INTrotate","g":"Common"},"inthspace":{"d":"INThspace","g":"Common"},"inthbox":{"d":"INThbox","g":"Common"},"inttabimage":{"d":"INTtabimage","g":"Common"},"intcontfig":{"d":"INTcontfig","g":"Common"},"intit":{"d":"INTIT","g":"Common"},"srh":{"d":"srh","g":"Common"},"intcolspec":{"d":"INTcolspec","g":"Common"},"insvspace":{"d":"INSvspace","g":"Common"},"insailabel":{"d":"INSaiLabel","g":"Common"},"insspanrule":{"d":"INSspanrule","g":"Common"},"insmonthprint":{"d":"INSmonthprint","g":"Common"},"insjournalmonth":{"d":"INSjournalmonth","g":"Common"},"inshspace":{"d":"INShspace","g":"Common"},"insvskip":{"d":"INSvskip","g":"Common"},"inshskip":{"d":"INShskip","g":"Common"},"insxmlelement":{"d":"INSxmlelement","g":"Common"},"intxmlentity":{"d":"INTxmlentity","g":"Common"},"inscolorfiginfo":{"d":"INScolorfiginfo","g":"Common"},"insnotename":{"d":"INSnotename","g":"Common"},"instblbrk":{"d":"INStblbrk","g":"Common"},"instxtsuperscript":{"d":"INStxtsuperscript","g":"Common"},"inssetcounter":{"d":"INSsetcounter","g":"Common"},"instblbelowspace":{"d":"INStblbelowspace","g":"Common"},"insenlargethispage":{"d":"INSenlargethispage","g":"Common"},"inscommon":{"d":"INScommon","g":"Common"},"listtitle":{"d":"listtitle","g":"Common"},"pbox":{"d":"postbox","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"postbox":{"d":"postbox","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"reviewmeta":{"d":"review-meta","g":"Back Matter"},"govt":{"d":"Gov","g":"Back Matter"},"head1":{"d":"head1","g":"Common"},"head2":{"d":"head2","g":"Common"},"head3":{"d":"head3","g":"Common"},"head4":{"d":"head4","g":"Common"},"head5":{"d":"head5","g":"Common"},"head6":{"d":"head6","g":"Common"},"jel":{"d":"jel","g":"Common"},"orcid":{"d":"orcid","g":"Common"},"nump":{"d":"Num-p","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"pmi":{"d":"PMI","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"category":{"d":"Category","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"unlist":{"d":"Un_List","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"punindent":{"d":"p-unindent","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"toc":{"d":"toc","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"tocentry":{"d":"toc-entry","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"transkwdgroup":{"d":"trans-kwd-group","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"indexdiv":{"d":"index-div","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"indexentry":{"d":"index-entry","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"see":{"d":"see","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"seealso":{"d":"see-also","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"hidden":{"d":"hidden","g":"Body Matter"},"fnref":{"d":"fnref","g":"Common"},"code":{"d":"code","g":"Common"},"legalcase":{"d":"legal-case","g":"Back Matter"},"data":{"d":"data","g":"Back Matter"},"ringgold":{"d":"ringgold","g":"Back Matter"},"isni":{"d":"isni","g":"Back Matter"},"social":{"d":"social","g":"Back Matter"},"broadcast":{"d":"broadcast","g":"Back Matter"},"legal":{"d":"legal","g":"Back Matter"},"arxiv":{"d":"arxiv","g":"Back Matter"},"preprint":{"d":"preprint","g":"Back Matter"},"software":{"d":"software","g":"Back Matter"},"nonjournal":{"d":"nonjournal","g":"Back Matter"},"unpublished":{"d":"unpublished","g":"Back Matter"},"iso":{"d":"iso","g":"Back Matter"},"addmaterial":{"d":"add-material","g":"Back Matter"},"genusspecies":{"d":"genus-species","g":"Body Matter"},"si":{"d":"si","g":"Body Matter"},"dblink":{"d":"db-link","g":"Body Matter"},"sbond":{"d":"sbond","g":"Body Matter"},"dbond":{"d":"dbond","g":"Body Matter"},"tbond":{"d":"tbond","g":"Body Matter"},"keeptogether":{"d":"keep-together","g":"Body Matter"},"ptn":{"d":"ptn","g":"Body Matter"},"bibissue":{"d":"bibissue","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"bibdataset":{"d":"bibdataset","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"datasettitle":{"d":"datasettitle","g":"Journal Metadata PLACEHOLDER_LOOKUP Front Matter"},"particle":{"d":"particle","g":"Common"},"stl":{"d":"SeriesTitle","g":"Back Matter"},"size;specificuse=runingtime;units=minutes":{"d":"SizeMinutes","g":"Back Matter"},"sizeminutes":{"d":"SizeMinutes","g":"Back Matter"},"size;specificuse=runingtime;units=hours":{"d":"SizeHours","g":"Back Matter"},"sizehours":{"d":"SizeHours","g":"Back Matter"},"size;specificuse=runingtime;units=seconds":{"d":"SizeSecond","g":"Back Matter"},"sizesecond":{"d":"SizeSecond","g":"Back Matter"},"au;namestyle=western":{"d":"Au_western","g":"Back Matter"},"auwestern":{"d":"Au_western","g":"Back Matter"},"shorttitle":{"d":"shorttitle","g":"Common"}};

    // ==========================================================================
    // DYNAMIC XSD PARSER — reads any schema structure (journal, book, etc.)
    // ==========================================================================
    window.parseXsdSchema = function (xsdText) {
        var parser = new DOMParser();
        var xsdDoc = parser.parseFromString(xsdText, "application/xml");
        var ns = "http://www.w3.org/2001/XMLSchema";

        // Collect all top-level element definitions: name → DOM element
        var elemDefs = {};
        var root = xsdDoc.documentElement;
        for (var i = 0; i < root.children.length; i++) {
            var ch = root.children[i];
            if (ch.localName === "element" && ch.getAttribute("name"))
                elemDefs[ch.getAttribute("name")] = ch;
        }

        // Collect all named complexType and simpleType definitions
        var typeDefs = {};
        for (var i = 0; i < root.children.length; i++) {
            var ch = root.children[i];
            if ((ch.localName === "complexType" || ch.localName === "simpleType") && ch.getAttribute("name"))
                typeDefs[ch.getAttribute("name")] = ch;
        }

        // Collect named groups
        var groupDefs = {};
        for (var i = 0; i < root.children.length; i++) {
            var ch = root.children[i];
            if (ch.localName === "group" && ch.getAttribute("name"))
                groupDefs[ch.getAttribute("name")] = ch;
        }

        // Detect root element (article, book, book-part, etc.)
        var rootName = detectRootElement(elemDefs);
        console.log("[XSD] Root element: " + rootName + " | Elements: " + Object.keys(elemDefs).length +
            " | Types: " + Object.keys(typeDefs).length);

        // Build fields by walking the schema
        var fields = [];
        var seen = {};

        // 1. Extract attributes from root element
        if (elemDefs[rootName]) {
            extractAttributes(elemDefs[rootName], ns, rootName, fields, seen, typeDefs);
        }

        // 2. Walk the root element's children recursively (max 3 levels deep)
        if (elemDefs[rootName]) {
            walkElement(elemDefs[rootName], ns, "", 0, fields, seen, elemDefs, typeDefs, groupDefs);
        }

        // 3. Also add any top-level elements not yet seen (orphans)
        var elemNames = Object.keys(elemDefs);
        for (var i = 0; i < elemNames.length; i++) {
            var eName = elemNames[i];
            if (seen[eName] || eName === rootName) continue;
            addField(eName, "", fields, seen);
        }

        console.log("[XSD] Generated " + fields.length + " schema fields from " + rootName);
        _schemaTree = fields;
        return JSON.stringify(_schemaTree);
    };

    // Detect which element is the root (article, book, book-part, etc.)
    function detectRootElement(elemDefs) {
        // Priority order: known root element names
        var roots = ["article", "book", "book-part", "book-part-wrapper", "collection", "series"];
        for (var i = 0; i < roots.length; i++) {
            if (elemDefs[roots[i]]) return roots[i];
        }
        // Fallback: first element in the map
        var keys = Object.keys(elemDefs);
        return keys.length > 0 ? keys[0] : "unknown";
    }

    // Extract attributes from an element definition
    function extractAttributes(elemDef, ns, parentName, fields, seen, typeDefs) {
        var attrs = [];
        // Direct attributes
        collectAttrs(elemDef, ns, attrs, typeDefs);
        // Also check named type
        var typeName = elemDef.getAttribute("type");
        if (typeName && typeDefs[typeName]) {
            collectAttrs(typeDefs[typeName], ns, attrs, typeDefs);
        }

        for (var i = 0; i < attrs.length; i++) {
            var a = attrs[i];
            var key = "@" + a.name;
            if (seen[key]) continue;
            seen[key] = true;

            var info = lookupTag(a.name) || lookupTag(parentName + "-" + a.name);
            var group = info ? info.g : capitalize(parentName);
            var label = info ? info.d : a.name;

            fields.push({
                xpath: key,
                label: label,
                type: a.enumValues.length > 0 ? "dropdown" : "text",
                group: group,
                required: a.required,
                value: "",
                enumValues: a.enumValues
            });
        }
    }

    // Collect xs:attribute from a node (recursing into complexType)
    function collectAttrs(node, ns, out, typeDefs) {
        if (!node) return;
        var children = node.children;
        for (var i = 0; i < children.length; i++) {
            var ch = children[i];
            if (ch.localName === "attribute" && ch.getAttribute("name")) {
                var attrName = ch.getAttribute("name");
                var required = ch.getAttribute("use") === "required";
                var enumVals = extractEnumValues(ch, ns, typeDefs);
                out.push({ name: attrName, required: required, enumValues: enumVals });
            }
            // Recurse into complexType, complexContent, extension, restriction
            if (ch.localName === "complexType" || ch.localName === "complexContent" ||
                ch.localName === "simpleContent" || ch.localName === "extension" ||
                ch.localName === "restriction") {
                collectAttrs(ch, ns, out, typeDefs);
            }
        }
    }

    // Extract enumeration values from an attribute or simple type
    function extractEnumValues(node, ns, typeDefs) {
        var vals = [];
        // Check for inline restriction
        var enums = node.getElementsByTagNameNS(ns, "enumeration");
        for (var i = 0; i < enums.length; i++) {
            var v = enums[i].getAttribute("value");
            if (v) vals.push(v);
        }
        // Check named type reference
        if (vals.length === 0) {
            var typeName = node.getAttribute("type");
            if (typeName && typeDefs[typeName]) {
                var enums2 = typeDefs[typeName].getElementsByTagNameNS(ns, "enumeration");
                for (var i = 0; i < enums2.length; i++) {
                    var v = enums2[i].getAttribute("value");
                    if (v) vals.push(v);
                }
            }
        }
        return vals;
    }

    // Walk an element's complex type to find child elements
    function walkElement(elemDef, ns, parentGroup, depth, fields, seen, elemDefs, typeDefs, groupDefs) {
        if (depth > 3) return; // prevent infinite recursion

        var children = getChildElements(elemDef, ns, elemDefs, typeDefs, groupDefs);

        for (var i = 0; i < children.length; i++) {
            var child = children[i];
            var eName = child.name;
            if (seen[eName]) continue;

            var childDef = elemDefs[eName] || null;
            var hasChildren = childDef ? hasChildElements(childDef, ns, elemDefs, typeDefs, groupDefs) : false;
            var maxOccurs = child.maxOccurs;
            var minOccurs = child.minOccurs;

            // Determine group from JATS lookup or parent
            var info = lookupTag(eName);
            var group = info ? info.g : (parentGroup || "Schema Fields");
            var label = info ? info.d : eName;

            // Determine type
            var fieldType = "text";
            if (hasChildren || maxOccurs === "unbounded" || parseInt(maxOccurs) > 1) {
                fieldType = "collection";
            } else if (isMixedContent(childDef, typeDefs)) {
                fieldType = "richtext";
            }

            seen[eName] = true;
            fields.push({
                xpath: eName,
                label: label,
                type: fieldType,
                group: group,
                required: minOccurs !== "0",
                value: "",
                children: fieldType === "collection" ? [] : undefined,
                enumValues: []
            });

            // Also extract attributes of this child element
            if (childDef) {
                extractAttributes(childDef, ns, eName, fields, seen, typeDefs);
            }

            // Recurse into children (for grouping sub-elements)
            if (childDef && hasChildren) {
                walkElement(childDef, ns, group, depth + 1, fields, seen, elemDefs, typeDefs, groupDefs);
            }
        }
    }

    // Get child element references from an element's complexType
    function getChildElements(elemDef, ns, elemDefs, typeDefs, groupDefs) {
        var results = [];
        var visited = {};

        function scanNode(node) {
            if (!node) return;
            var children = node.children;
            for (var i = 0; i < children.length; i++) {
                var ch = children[i];
                if (ch.localName === "element") {
                    var name = ch.getAttribute("name") || ch.getAttribute("ref");
                    if (name && !visited[name]) {
                        visited[name] = true;
                        results.push({
                            name: name,
                            minOccurs: ch.getAttribute("minOccurs") || "1",
                            maxOccurs: ch.getAttribute("maxOccurs") || "1"
                        });
                    }
                }
                // Follow group references
                if (ch.localName === "group" && ch.getAttribute("ref")) {
                    var gRef = ch.getAttribute("ref");
                    if (groupDefs[gRef]) scanNode(groupDefs[gRef]);
                }
                // Recurse into structural nodes
                if (ch.localName === "complexType" || ch.localName === "sequence" ||
                    ch.localName === "choice" || ch.localName === "all" ||
                    ch.localName === "complexContent" || ch.localName === "extension" ||
                    ch.localName === "restriction") {
                    scanNode(ch);
                }
            }
        }

        scanNode(elemDef);

        // Also check named type
        var typeName = elemDef.getAttribute("type");
        if (typeName && typeDefs[typeName]) {
            scanNode(typeDefs[typeName]);
        }

        return results;
    }

    // Check if an element has child elements (is complex)
    function hasChildElements(elemDef, ns, elemDefs, typeDefs, groupDefs) {
        return getChildElements(elemDef, ns, elemDefs, typeDefs, groupDefs).length > 0;
    }

    // Check if element has mixed content
    function isMixedContent(elemDef, typeDefs) {
        if (!elemDef) return false;
        // Check inline complexType
        var cts = elemDef.getElementsByTagNameNS("http://www.w3.org/2001/XMLSchema", "complexType");
        for (var i = 0; i < cts.length; i++) {
            if (cts[i].getAttribute("mixed") === "true") return true;
        }
        // Check named type
        var typeName = elemDef.getAttribute("type");
        if (typeName && typeDefs[typeName]) {
            if (typeDefs[typeName].getAttribute("mixed") === "true") return true;
        }
        return false;
    }

    // Add a field from element name using JATS lookup
    function addField(eName, parentGroup, fields, seen) {
        if (seen[eName]) return;
        seen[eName] = true;
        var info = lookupTag(eName);
        fields.push({
            xpath: eName,
            label: info ? info.d : eName,
            type: "text",
            group: info ? info.g : (parentGroup || "Schema Fields"),
            required: false,
            value: "",
            enumValues: []
        });
    }

    function capitalize(str) {
        if (!str) return "Schema Fields";
        return str.charAt(0).toUpperCase() + str.slice(1).replace(/-/g, " ");
    }

    // ==========================================================================
    // XML PARSER
    // ==========================================================================
    window.parseXmlForSchema = function (xmlText) {
        var parser = new DOMParser();
        var xmlDoc = parser.parseFromString(xmlText, "application/xml");
        if (!_schemaTree) return "[]";
        var article = xmlDoc.querySelector("article") || xmlDoc.documentElement;
        for (var i = 0; i < _schemaTree.length; i++) {
            var field = _schemaTree[i];
            if (field.xpath.charAt(0) === "@") {
                field.value = article.getAttribute(field.xpath.substring(1)) || "";
            } else if (field.type !== "collection" && field.type !== "sections") {
                var el = article.querySelector(field.xpath);
                if (el) field.value = el.textContent.trim();
            }
        }
        return JSON.stringify(_schemaTree);
    };

    // ==========================================================================
    // TAG NORMALIZATION & MAP
    // ==========================================================================

    function normalizeTag(str) {
        if (!str) return "";
        return str.toLowerCase().replace(/[#;]/g, "").replace(/[\s_\-.:\/]+/g, "").replace(/[^a-z0-9]/g, "");
    }

    function buildTagMap(schemaTree) {
        var map = {};
        for (var i = 0; i < schemaTree.length; i++) {
            var f = schemaTree[i];
            var normLabel = normalizeTag(f.label);
            if (normLabel) map[normLabel] = i;
            var normXpath = normalizeTag(f.xpath);
            if (normXpath && normXpath !== normLabel) map[normXpath] = i;
        }
        return map;
    }

    // Look up a raw tag in the JATS database to get display name & group
    function lookupTag(rawTag) {
        var norm = normalizeTag(rawTag);
        if (_tagLookup[norm]) return _tagLookup[norm];
        // Try with # and ; stripped (tags from picker come as "copyright-statement", DB has them normalized)
        var clean = rawTag.replace(/^#/, "").replace(/;$/, "");
        var norm2 = normalizeTag(clean);
        if (_tagLookup[norm2]) return _tagLookup[norm2];
        return null;
    }

    // ==========================================================================
    // SFDT ACCESSORS
    // ==========================================================================

    function getSections(sfdt) { return sfdt.sections || sfdt.sec || []; }
    function getBlocks(section) { return section.blocks || section.b || []; }
    function getInlines(block) { return block.inlines || block.i || []; }
    function getTableRows(block) { return block.rows || block.r || []; }
    function getTableCells(row) { return row.cells || row.c || []; }

    function getInlineText(inline) {
        if (inline.text !== undefined && inline.text !== null) return inline.text;
        if (inline.tlp !== undefined && inline.tlp !== null) return inline.tlp;
        return "";
    }

    // ==========================================================================
    // CONTENT CONTROL EXTRACTION FROM SFDT
    // ==========================================================================

    function getCCTag(ccp) { return ccp.tag || ccp.tg || ""; }
    function getCCTitle(ccp) { return ccp.title || ccp.tt || ""; }

    function extractTextFromInlines(inlines) {
        if (!inlines) return "";
        var text = "";
        for (var i = 0; i < inlines.length; i++) {
            var inl = inlines[i];
            var t = getInlineText(inl);
            if (t) text += t;
            var nested = inl.inlines || inl.i;
            if (nested) text += extractTextFromInlines(nested);
        }
        return text;
    }

    function extractTextFromBlocks(blocks) {
        if (!blocks) return "";
        var text = "";
        for (var i = 0; i < blocks.length; i++) {
            var block = blocks[i];
            var inlines = getInlines(block);
            if (inlines && inlines.length > 0) {
                var t = extractTextFromInlines(inlines);
                if (t) text += (text ? " " : "") + t;
            }
            var nested = block.blocks || block.b;
            if (nested) {
                var nt = extractTextFromBlocks(nested);
                if (nt) text += (text ? " " : "") + nt;
            }
        }
        return text.trim();
    }

    function walkBlocksForCC(blocks, results) {
        if (!blocks) return;
        for (var i = 0; i < blocks.length; i++) {
            var block = blocks[i];
            var ccp = block.contentControlProperties || block.ccp;
            if (ccp) {
                var tag = getCCTag(ccp);
                var title = getCCTitle(ccp);
                var childBlocks = block.blocks || block.b || [];
                var text = extractTextFromBlocks(childBlocks);
                if (!text) text = extractTextFromInlines(getInlines(block));
                results.push({ tag: tag, title: title, text: text, source: "sfdt-block" });
                walkBlocksForCC(childBlocks, results);
                continue;
            }
            var inlines = getInlines(block);
            if (inlines) walkInlinesForCC(inlines, results);
            var nested = block.blocks || block.b;
            if (nested) walkBlocksForCC(nested, results);
            var rows = getTableRows(block);
            if (rows && rows.length > 0) {
                for (var ri = 0; ri < rows.length; ri++) {
                    var cells = getTableCells(rows[ri]);
                    for (var ci = 0; ci < cells.length; ci++) {
                        walkBlocksForCC(cells[ci].blocks || cells[ci].b || [], results);
                    }
                }
            }
        }
    }

    function walkInlinesForCC(inlines, results) {
        if (!inlines) return;
        for (var i = 0; i < inlines.length; i++) {
            var inl = inlines[i];
            var ccp = inl.contentControlProperties || inl.ccp;
            if (ccp) {
                var tag = getCCTag(ccp);
                var title = getCCTitle(ccp);
                var nested = inl.inlines || inl.i || [];
                var text = extractTextFromInlines(nested);
                results.push({ tag: tag, title: title, text: text, source: "sfdt-inline" });
            } else {
                var n = inl.inlines || inl.i;
                if (n) walkInlinesForCC(n, results);
            }
        }
    }

    // ==========================================================================
    // CC TYPE MAPPING — maps Syncfusion CC type to display name
    // API returns: "RichText", "Text", "Picture", "ComboBox", "DropDownList", "DatePicker", "CheckBox"
    // ==========================================================================

    function mapCCType(ccType) {
        if (!ccType) return "";
        var t = ccType.toLowerCase().replace(/\s+/g, "");
        if (t === "richtext") return "Rich Text";
        if (t === "text" || t === "plaintext") return "Plain Text";
        if (t === "picture") return "Picture";
        if (t === "combobox") return "Combo Box";
        if (t === "dropdownlist") return "Drop-Down";
        if (t === "datepicker") return "Date Picker";
        if (t === "checkbox") return "Check Box";
        return ccType; // fallback: return as-is
    }

    // ==========================================================================
    // MATCH — maps CCs to schema fields, collects repeated tags as children
    // ==========================================================================

    function matchCCsToSchema(ccList, tagMap) {
        var matchedCount = 0;
        var createdCount = 0;
        var hitCount = {};   // fieldIdx → count
        var tagSummary = {}; // tag → count (for logging)

        for (var ci = 0; ci < ccList.length; ci++) {
            var cc = ccList[ci];
            var normTag = normalizeTag(cc.tag);
            var normTitle = normalizeTag(cc.title);
            var ccText = (cc.text || "").trim();
            if (!ccText) continue; // skip empty CCs

            // Try matching existing schema field
            var fieldIdx = (normTag && tagMap[normTag] !== undefined) ? tagMap[normTag] : null;
            if (fieldIdx === null) {
                fieldIdx = (normTitle && tagMap[normTitle] !== undefined) ? tagMap[normTitle] : null;
            }

            if (fieldIdx !== null) {
                // ── Matched existing field ────────────────────────────────
                var field = _schemaTree[fieldIdx];
                if (!hitCount[fieldIdx]) hitCount[fieldIdx] = 0;
                hitCount[fieldIdx]++;

                var truncText = ccText.length > 200 ? ccText.substring(0, 200) + "..." : ccText;

                if (hitCount[fieldIdx] === 1) {
                    // First hit — set value directly
                    field.value = truncText;
                    // Update field type from actual CC type
                    var mappedType = mapCCType(cc.ccType);
                    if (mappedType) field.type = mappedType;
                } else {
                    // Subsequent hit — add as child
                    if (hitCount[fieldIdx] === 2) {
                        // Convert first value to child too
                        if (!field.children) field.children = [];
                        field.children.unshift({
                            label: field.label + " 1",
                            depth: 0,
                            fields: [{ key: "text", label: "Text", value: field.value }],
                            children: []
                        });
                    }
                    if (!field.children) field.children = [];
                    field.children.push({
                        label: field.label + " " + hitCount[fieldIdx],
                        depth: 0,
                        fields: [{ key: "text", label: "Text", value: truncText }],
                        children: []
                    });
                    field.value = hitCount[fieldIdx] + " instance(s)";
                }

                matchedCount++;
                var sk = cc.tag || cc.title;
                tagSummary[sk] = (tagSummary[sk] || 0) + 1;

            } else {
                // ── No match — auto-create a new field ───────────────────
                var rawTag = cc.tag || cc.title || "";
                if (!rawTag) continue;

                var info = lookupTag(rawTag);
                var displayName = info ? info.d : rawTag.replace(/^#/, "").replace(/;$/, "");
                var groupName = info ? info.g : "Document Tags";

                var newField = {
                    xpath: rawTag,
                    label: displayName,
                    type: mapCCType(cc.ccType) || "Rich Text",
                    group: groupName,
                    required: false,
                    value: ccText.length > 200 ? ccText.substring(0, 200) + "..." : ccText,
                    children: []
                };

                _schemaTree.push(newField);
                var newIdx = _schemaTree.length - 1;
                hitCount[newIdx] = 1;
                if (normTag) tagMap[normTag] = newIdx;
                if (normTitle && normTitle !== normTag) tagMap[normTitle] = newIdx;

                createdCount++;
                tagSummary[rawTag] = 1;
            }
        }

        // Log summary per tag (not per CC)
        var tagKeys = Object.keys(tagSummary).sort(function (a, b) { return tagSummary[b] - tagSummary[a]; });
        for (var ti = 0; ti < Math.min(tagKeys.length, 30); ti++) {
            var tk = tagKeys[ti];
            console.log("[Schema]   " + tk + " \u00d7" + tagSummary[tk]);
        }
        if (tagKeys.length > 30) {
            console.log("[Schema]   ... and " + (tagKeys.length - 30) + " more tag types");
        }

        return { matched: matchedCount, created: createdCount };
    }

    // ==========================================================================
    // MAIN EXTRACTION
    // ==========================================================================

    window.extractDocContentForSchema = function () {
        if (!_schemaTree) { _schemaTree = []; console.log("[Schema] No XSD — CC-only mode"); }

        var ec = window._deContainer || window.container;
        if (!ec || !ec.documentEditor) { console.warn("[Schema] No editor"); return "[]"; }

        console.log("[Schema] v7.1 \u2014 Dynamic CC Tag Matching (with collection)");

        // Reset values but keep structure (remove only dynamic fields)
        var baseCount = 0;
        var cleaned = [];
        for (var i = 0; i < _schemaTree.length; i++) {
            var f = _schemaTree[i];
            if (f._dynamic) continue;
            f.value = "";
            f.children = [];
            cleaned.push(f);
            baseCount++;
        }
        _schemaTree = cleaned;

        var tagMap = buildTagMap(_schemaTree);
        var ccList = [];

        // ── PRIMARY: Syncfusion API ──────────────────────────────────────
        try {
            var apiData = ec.documentEditor.exportContentControlData();
            if (apiData && apiData.length > 0) {
                console.log("[Schema] API found " + apiData.length + " content control(s)");
                for (var ai = 0; ai < apiData.length; ai++) {
                    var item = apiData[ai];
                    ccList.push({
                        tag: item.tag || "",
                        title: item.title || "",
                        text: item.value || "",
                        ccType: item.type || "",
                        source: "api"
                    });
                }
            } else {
                console.log("[Schema] API returned 0 CCs, falling back to SFDT scan");
            }
        } catch (e) {
            console.warn("[Schema] API failed: " + e.message + ", falling back to SFDT scan");
        }

        // ── FALLBACK: SFDT scan ──────────────────────────────────────────
        if (ccList.length === 0) {
            var sfdtStr;
            try { sfdtStr = ec.documentEditor.serialize(); } catch (e) { console.error("[Schema] serialize:", e); return "[]"; }
            var sfdt;
            try { sfdt = JSON.parse(sfdtStr); } catch (e) { console.error("[Schema] parse:", e); return "[]"; }
            var sections = getSections(sfdt);
            for (var si = 0; si < sections.length; si++) {
                walkBlocksForCC(getBlocks(sections[si]), ccList);
            }
            console.log("[Schema] SFDT scan found " + ccList.length + " CC(s)");
        }

        // ── MATCH & AUTO-CREATE ──────────────────────────────────────────
        var result = matchCCsToSchema(ccList, tagMap);

        // Mark dynamic fields
        for (var i = baseCount; i < _schemaTree.length; i++) {
            _schemaTree[i]._dynamic = true;
        }

        // ── SUMMARY ──────────────────────────────────────────────────────
        var filledCount = 0;
        for (var i = 0; i < _schemaTree.length; i++) {
            var f = _schemaTree[i];
            if (f.value || (f.children && f.children.length > 0)) filledCount++;
        }
        console.log("[Schema] Done: " + filledCount + "/" + _schemaTree.length + " fields populated" +
            " (" + result.matched + " CCs matched, " + result.created + " new tags, " +
            (_schemaTree.length - baseCount) + " dynamic fields)");

        return JSON.stringify(_schemaTree);
    };

    // ==========================================================================
    // HELPERS
    // ==========================================================================
    window.getSchemaTree = function () { return _schemaTree ? JSON.stringify(_schemaTree) : "[]"; };
    window.clearSchemaTree = function () { _schemaTree = null; console.log("[Schema] Schema tree cleared"); };

    // ==========================================================================
    // NAVIGATION
    // ==========================================================================
    window.navigateToDocText = function (searchText) {
        var ec = window._deContainer || window.container;
        if (!ec || !ec.documentEditor || !searchText) return false;
        var query = searchText.trim();
        if (!query) return false;
        if (query.length > 50) query = query.substring(0, 50);
        try {
            ec.documentEditor.search.find(query, "None");
            return true;
        } catch (e) {
            if (query.length > 25) {
                try { ec.documentEditor.search.find(query.substring(0, 25), "None"); return true; }
                catch (e2) {}
            }
            return false;
        }
    };
})();