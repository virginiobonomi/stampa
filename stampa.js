/*
 * Atti — modelli di stampa degli originali (come RedAtto):
 *   A4 o A3 a libretto, 25 o 50 righe per pagina, interlineato (righe completate con trattini) o no.
 * Dal contenuto del documento (OOXML restituito da Word) costruisce un nuovo .docx pronto per la stampa.
 *
 * Interlineatura (regole di RedAtto, manuale par. 5.3.4):
 *   - la parte vuota dell'ultima riga di ogni paragrafo si riempie di trattini;
 *   - i paragrafi centrati hanno trattini prima e dopo, quelli allineati a destra prima;
 *   - le righe vuote si barrano per intero;
 *   - NON si interlinea una riga chiusa con a capo manuale (Maiusc+Invio) né un paragrafo
 *     che termina con uno spazio unificatore (Ctrl+Maiusc+Spazio in Windows, Opzione+Spazio su Mac).
 */
(function (root) {
  "use strict";

  var W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  var R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  var PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
  var REL = "http://schemas.openxmlformats.org/package/2006/relationships";

  // misure dei modelli RedAtto (twip): margini 2181/2188/1359/1359, 25 righe = interlinea esatta 562.
  // Word scrive il testo in fondo alla riga "esatta"; RedAtto in alto: con 25 righe si alza la gabbia di 230 twip
  // (11,5 pt) perché le righe cadano alla stessa altezza degli originali RedAtto.
  var MODELLI = {
    "a4-25": { nome: "A4 · 25 righe", a3: false, riga: 562, su: 230 },
    "a4-50": { nome: "A4 · 50 righe", a3: false, riga: 281, su: 0 },
    "a3-25": { nome: "A3 a libretto · 25 righe", a3: true, riga: 562, su: 230 },
    "a3-50": { nome: "A3 a libretto · 50 righe", a3: true, riga: 281, su: 0 }
  };
  var PAGINA = { w: 11906, h: 16838, sx: 2181, dx: 2188, su: 1359, giu: 1359, intest: 720, piede: 361 };
  var LARGHEZZA = PAGINA.w - PAGINA.sx - PAGINA.dx;     // 7537 twip di testo
  var CAR_RIGA = 60;    // oltre questa lunghezza un paragrafo centrato o a destra occupa più righe

  // ordine dei figli di w:pPr (lo schema lo impone)
  var ORDINE_PPR = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
    "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
    "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];

  function figli(n, nome) {
    var out = [];
    for (var c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1 && c.namespaceURI === W && c.localName === nome) out.push(c);
    return out;
  }
  function figlio(n, nome) { return figli(n, nome)[0] || null; }
  function crea(doc, nome, attr) {
    var e = doc.createElementNS(W, "w:" + nome);
    Object.keys(attr || {}).forEach(function (k) { e.setAttributeNS(W, "w:" + k, String(attr[k])); });
    return e;
  }
  function attr(e, nome) { return e ? (e.getAttributeNS(W, nome) || e.getAttribute("w:" + nome)) : null; }

  function pPrDi(p) {
    var ppr = figlio(p, "pPr");
    if (!ppr) { ppr = crea(p.ownerDocument, "pPr"); p.insertBefore(ppr, p.firstChild); }
    return ppr;
  }
  function imposta(ppr, nome, el) {
    figli(ppr, nome).forEach(function (x) { ppr.removeChild(x); });
    if (!el) return;
    var pos = ORDINE_PPR.indexOf(nome), dopo = null;
    for (var c = ppr.firstChild; c; c = c.nextSibling) {
      if (c.nodeType === 1 && ORDINE_PPR.indexOf(c.localName) > pos) { dopo = c; break; }
    }
    ppr.insertBefore(el, dopo);
  }

  // testo del paragrafo (\t per le tabulazioni, \n per gli a capo manuali)
  function testoParagrafo(p) {
    var s = "";
    (function giro(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.localName === "del" || c.localName === "pPr" || c.localName === "rPr" || c.localName === "instrText") continue;
        if (c.localName === "t") s += c.textContent;
        else if (c.localName === "tab" && c.parentNode.localName === "r") s += "\t";
        else if (c.localName === "br" || c.localName === "cr") s += "\n";
        else giro(c);
      }
    })(p);
    return s;
  }
  function haOggetti(p) {
    return p.getElementsByTagNameNS(W, "drawing").length > 0 || p.getElementsByTagNameNS(W, "pict").length > 0 ||
      p.getElementsByTagNameNS(W, "object").length > 0;
  }

  function runTab(doc) {
    var r = crea(doc, "r");
    r.appendChild(crea(doc, "tab"));
    return r;
  }

  function aggiungiTab(ppr, val, pos) {
    var tabs = figlio(ppr, "tabs");
    if (!tabs) { tabs = crea(ppr.ownerDocument, "tabs"); imposta(ppr, "tabs", tabs); }
    // niente doppioni nella stessa posizione
    figli(tabs, "tab").forEach(function (t) { if (+attr(t, "pos") === pos) tabs.removeChild(t); });
    tabs.appendChild(crea(ppr.ownerDocument, "tab", { val: val, leader: "hyphen", pos: pos }));   // trattini, come RedAtto
  }

  function interlinea(p, doc) {
    var testo = testoParagrafo(p);
    if (haOggetti(p)) return;
    if (/ $/.test(testo) || /\n$/.test(testo)) return;          // spazio unificatore o a capo manuale in fondo
    var ppr = pPrDi(p);
    var ind = figlio(ppr, "ind");
    var sx = ind ? (+attr(ind, "left") || +attr(ind, "start") || 0) : 0;
    var dx = ind ? (+attr(ind, "right") || +attr(ind, "end") || 0) : 0;
    var fine = LARGHEZZA - dx;
    var jc = attr(figlio(ppr, "jc"), "val") || "";
    var vuoto = !testo.replace(/\s/g, "");
    var unaRiga = testo.indexOf("\n") < 0 && testo.length <= CAR_RIGA;
    if (vuoto) {
      aggiungiTab(ppr, "right", fine);
      p.appendChild(runTab(doc));
      return;
    }
    if ((jc === "center") && unaRiga) {
      imposta(ppr, "jc", crea(doc, "jc", { val: "left" }));
      if (ind) { ind.removeAttributeNS(W, "firstLine"); ind.removeAttributeNS(W, "hanging"); }
      aggiungiTab(ppr, "center", Math.round(sx + (fine - sx) / 2));
      aggiungiTab(ppr, "right", fine);
      p.insertBefore(runTab(doc), ppr.nextSibling);
      p.appendChild(runTab(doc));
      return;
    }
    if ((jc === "right" || jc === "end") && unaRiga) {
      imposta(ppr, "jc", crea(doc, "jc", { val: "left" }));
      if (ind) { ind.removeAttributeNS(W, "firstLine"); ind.removeAttributeNS(W, "hanging"); }
      aggiungiTab(ppr, "right", fine);
      p.insertBefore(runTab(doc), ppr.nextSibling);
      return;
    }
    if (jc === "center" || jc === "right" || jc === "end") imposta(ppr, "jc", crea(doc, "jc", { val: "both" }));
    aggiungiTab(ppr, "right", fine);
    p.appendChild(runTab(doc));
  }

  function impagina(p, doc, modello) {
    var ppr = pPrDi(p);
    imposta(ppr, "keepNext", null);
    imposta(ppr, "keepLines", null);
    imposta(ppr, "widowControl", crea(doc, "widowControl", { val: "0" }));
    imposta(ppr, "spacing", crea(doc, "spacing", { before: 0, after: 0, beforeAutospacing: 0, afterAutospacing: 0, line: modello.riga, lineRule: "exact" }));
    imposta(ppr, "contextualSpacing", null);
    imposta(ppr, "sectPr", null);                       // un'unica sezione
  }

  function uniformaRun(rpr) {
    ["rFonts", "sz", "szCs"].forEach(function (n) { figli(rpr, n).forEach(function (x) { rpr.removeChild(x); }); });
  }

  // ---------------------------------------------------------------- pacchetto

  function leggiPacchetto(xml) {
    var dom = new root.DOMParser().parseFromString(xml, "application/xml");
    var parti = {};
    var lista = dom.getElementsByTagNameNS(PKG, "part");
    for (var i = 0; i < lista.length; i++) {
      var part = lista[i];
      var nome = part.getAttributeNS(PKG, "name") || part.getAttribute("pkg:name");
      var tipo = part.getAttributeNS(PKG, "contentType") || part.getAttribute("pkg:contentType");
      var xd = part.getElementsByTagNameNS(PKG, "xmlData")[0];
      var bd = part.getElementsByTagNameNS(PKG, "binaryData")[0];
      var voce = { nome: nome, tipo: tipo };
      if (xd) {
        var radice = null;
        for (var c = xd.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) { radice = c; break; }
        var d = new root.DOMParser().parseFromString(new root.XMLSerializer().serializeToString(radice), "application/xml");
        voce.dom = d;
      } else if (bd) voce.base64 = bd.textContent.replace(/\s+/g, "");
      parti[nome] = voce;
    }
    return parti;
  }

  function relsDi(parti, nomeParte) {
    var i = nomeParte.lastIndexOf("/");
    var n = nomeParte.slice(0, i) + "/_rels/" + nomeParte.slice(i + 1) + ".rels";
    if (!parti[n]) {
      parti[n] = { nome: n, tipo: "application/vnd.openxmlformats-package.relationships+xml",
        dom: new root.DOMParser().parseFromString('<Relationships xmlns="' + REL + '"/>', "application/xml") };
    }
    return parti[n];
  }

  function aggiungiRel(rels, id, tipo, target) {
    var radice = rels.dom.documentElement;
    var tutte = radice.getElementsByTagNameNS(REL, "Relationship");
    for (var i = tutte.length - 1; i >= 0; i--) {
      if (tutte[i].getAttribute("Id") === id || tutte[i].getAttribute("Type") === tipo && /settings$/.test(tipo)) radice.removeChild(tutte[i]);
    }
    var r = rels.dom.createElementNS(REL, "Relationship");
    r.setAttribute("Id", id); r.setAttribute("Type", tipo); r.setAttribute("Target", target);
    radice.appendChild(r);
  }

  var XML_PIEDE = '<w:ftr xmlns:w="' + W + '" xmlns:r="' + R + '"><w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/></w:rPr></w:pPr>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">- </w:t></w:r>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/></w:rPr><w:t>1</w:t></w:r>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r>' +
    '<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve"> -</w:t></w:r></w:p></w:ftr>';

  function xmlImpostazioni(a3) {
    return '<w:settings xmlns:w="' + W + '"><w:defaultTabStop w:val="720"/>' +
      (a3 ? '<w:bookFoldPrinting/>' : "") +
      '<w:characterSpacingControl w:val="doNotCompress"/>' +
      '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>';
  }

  // xmlPacchetto: OOXML di body.getOoxml(); opz: { modello, interlineato, uniforma }
  function prepara(xmlPacchetto, opz) {
    var modello = MODELLI[opz.modello];
    if (!modello) throw new Error("Modello di stampa sconosciuto");
    var parti = leggiPacchetto(xmlPacchetto);
    var nomeDoc = "/word/document.xml";
    if (!parti[nomeDoc]) {
      Object.keys(parti).forEach(function (n) { if (/document\.main\+xml$/.test(parti[n].tipo || "")) nomeDoc = n; });
    }
    var doc = parti[nomeDoc].dom;
    var body = doc.getElementsByTagNameNS(W, "body")[0];

    // paragrafi del corpo (non quelli nelle tabelle)
    var paragrafi = [];
    (function giro(n) {
      for (var c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType !== 1) continue;
        if (c.namespaceURI === W && c.localName === "p") paragrafi.push(c);
        else if (c.namespaceURI === W && (c.localName === "sdt" || c.localName === "sdtContent" || c.localName === "customXml" || c.localName === "smartTag")) giro(c);
      }
    })(body);
    // l'ultimo paragrafo vuoto che Word aggiunge sempre non va barrato
    var ultimo = paragrafi[paragrafi.length - 1];
    if (ultimo && !testoParagrafo(ultimo).replace(/\s/g, "") && paragrafi.length > 1) paragrafi.pop();

    paragrafi.forEach(function (p) {
      impagina(p, doc, modello);
      if (opz.interlineato) interlinea(p, doc);
    });
    // anche i paragrafi nelle tabelle prendono l'interlinea del modello
    Array.prototype.forEach.call(body.getElementsByTagNameNS(W, "tc"), function (tc) {
      figli(tc, "p").forEach(function (p) { impagina(p, doc, modello); });
    });
    if (opz.uniforma) {
      Array.prototype.slice.call(doc.getElementsByTagNameNS(W, "rPr")).forEach(uniformaRun);
    }

    // sezione unica con le misure del modello
    figli(body, "sectPr").forEach(function (s) { body.removeChild(s); });
    var sect = crea(doc, "sectPr");
    var fr = doc.createElementNS(W, "w:footerReference");
    fr.setAttributeNS(W, "w:type", "default");
    fr.setAttributeNS(R, "r:id", "rIdAttiPiede");
    sect.appendChild(fr);
    sect.appendChild(modello.a3 ? crea(doc, "pgSz", { w: 23811, h: 16838, orient: "landscape" }) : crea(doc, "pgSz", { w: PAGINA.w, h: PAGINA.h }));
    sect.appendChild(crea(doc, "pgMar", { top: PAGINA.su - modello.su, right: PAGINA.dx, bottom: PAGINA.giu + modello.su, left: PAGINA.sx, header: PAGINA.intest, footer: PAGINA.piede, gutter: 0 }));
    sect.appendChild(crea(doc, "cols", { space: 720 }));
    body.appendChild(sect);

    // stili: Arial 12 come carattere predefinito
    var stili = null;
    Object.keys(parti).forEach(function (n) { if (/styles\+xml$/.test(parti[n].tipo || "") && /styles\.xml$/.test(n)) stili = parti[n]; });
    if (stili) {
      var sd = stili.dom;
      var dd = sd.getElementsByTagNameNS(W, "docDefaults")[0];
      if (!dd) { dd = crea(sd, "docDefaults"); sd.documentElement.insertBefore(dd, sd.documentElement.firstChild); }
      var rprd = figlio(dd, "rPrDefault") || dd.insertBefore(crea(sd, "rPrDefault"), dd.firstChild);
      var rpr = figlio(rprd, "rPr") || rprd.appendChild(crea(sd, "rPr"));
      uniformaRun(rpr);
      rpr.insertBefore(crea(sd, "rFonts", { ascii: "Arial", hAnsi: "Arial", cs: "Arial", eastAsia: "Arial" }), rpr.firstChild);
      rpr.appendChild(crea(sd, "sz", { val: 24 })); rpr.appendChild(crea(sd, "szCs", { val: 24 }));
      if (opz.uniforma) Array.prototype.slice.call(sd.getElementsByTagNameNS(W, "style")).forEach(function (st) {
        var r2 = figlio(st, "rPr"); if (r2) uniformaRun(r2);
      });
    }

    // piè di pagina con il numero, impostazioni (libretto per l'A3)
    parti["/word/footerAtti.xml"] = { nome: "/word/footerAtti.xml", tipo: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
      dom: new root.DOMParser().parseFromString(XML_PIEDE, "application/xml") };
    Object.keys(parti).forEach(function (n) { if (/settings\+xml$/.test(parti[n].tipo || "")) delete parti[n]; });
    parti["/word/settingsAtti.xml"] = { nome: "/word/settingsAtti.xml", tipo: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
      dom: new root.DOMParser().parseFromString(xmlImpostazioni(modello.a3), "application/xml") };
    var rels = relsDi(parti, nomeDoc);
    aggiungiRel(rels, "rIdAttiPiede", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer", "footerAtti.xml");
    aggiungiRel(rels, "rIdAttiImp", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settingsAtti.xml");
    if (!parti["/_rels/.rels"]) {
      parti["/_rels/.rels"] = { nome: "/_rels/.rels", tipo: "application/vnd.openxmlformats-package.relationships+xml",
        dom: new root.DOMParser().parseFromString('<Relationships xmlns="' + REL + '"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="' + nomeDoc.slice(1) + '"/></Relationships>', "application/xml") };
    }
    return parti;
  }

  // parti -> file .docx (base64) con JSZip
  function docx(parti) {
    var zip = new root.JSZip();
    var tipi = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>';
    var ser = new root.XMLSerializer();
    Object.keys(parti).forEach(function (n) {
      var p = parti[n];
      if (!/\.rels$/.test(n)) tipi += '<Override PartName="' + n + '" ContentType="' + p.tipo + '"/>';
      if (p.dom && !/\.rels$/.test(n)) zip.file(n.slice(1), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ser.serializeToString(p.dom), { createFolders: false });
      else if (p.base64) zip.file(n.slice(1), p.base64, { base64: true, createFolders: false });
    });
    tipi += "</Types>";
    // relazioni verso parti che non ci sono (es. intestazioni del documento di partenza): si tolgono
    Object.keys(parti).forEach(function (n) {
      if (!/\.rels$/.test(n) || !parti[n].dom) return;
      var base = n.replace(/_rels\/[^\/]*\.rels$/, "");
      var tutte = parti[n].dom.getElementsByTagNameNS(REL, "Relationship");
      for (var i = tutte.length - 1; i >= 0; i--) {
        var r = tutte[i];
        if (r.getAttribute("TargetMode") === "External") continue;
        var t = r.getAttribute("Target");
        var assoluto = t.charAt(0) === "/" ? t : risolvi(base + t);
        if (!parti[assoluto]) r.parentNode.removeChild(r);
      }
    });
    Object.keys(parti).forEach(function (n) {
      var p = parti[n];
      if (/\.rels$/.test(n) && p.dom) zip.file(n.slice(1), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + ser.serializeToString(p.dom), { createFolders: false });
    });
    zip.file("[Content_Types].xml", tipi, { createFolders: false });
    return zip.generateAsync({ type: "base64", compression: "DEFLATE" });
  }

  function risolvi(percorso) {
    var out = [];
    percorso.split("/").forEach(function (x) { if (x === "..") out.pop(); else if (x && x !== ".") out.push(x); });
    return "/" + out.join("/");
  }

  function originale(xmlPacchetto, opz) {
    return docx(prepara(xmlPacchetto, opz));
  }

  root.Stampa = { MODELLI: MODELLI, prepara: prepara, docx: docx, originale: originale, LARGHEZZA: LARGHEZZA };
})(this);
