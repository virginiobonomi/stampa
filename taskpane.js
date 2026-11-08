/* Pannello Stampa originali: crea dal documento aperto un nuovo documento secondo i modelli di stampa RedAtto. */
(function () {
  "use strict";

  var $ = function (id) { return document.getElementById(id); };
  var inWord = false;

  function stato(id, msg, tipo) { var s = $(id); s.textContent = msg || ""; s.className = "stato" + (tipo ? " " + tipo : ""); }

  function creaOriginale() {
    if (!inWord) { stato("statoStampa", "Il pannello non è aperto dentro Word.", "err"); return; }
    var req = Office.context.requirements;
    if (req && !req.isSetSupported("WordApi", "1.3")) {
      stato("statoStampa", "Questa versione di Word non permette di creare un nuovo documento dal pannello: aggiorna Word.", "err"); return;
    }
    var opz = { modello: $("modelloStampa").value, interlineato: $("interlineato").checked, uniforma: $("uniforma").checked };
    try { window.localStorage.setItem("stampa.opzioni", JSON.stringify(opz)); } catch (e) { /* ignora */ }
    stato("statoStampa", "Preparazione dell'originale…");
    var a3 = Stampa.MODELLI[opz.modello].a3;
    // il libretto non passa dal file (Word scarta le impostazioni del documento): si imposta sulla pagina
    var librettoPossibile = a3 && req && req.isSetSupported("WordApiDesktop", "1.3") && req.isSetSupported("WordApiHiddenDocument", "1.3");
    var b64 = null;
    function apri(conLibretto) {
      return Word.run(function (ctx) {
        var nuovo = ctx.application.createDocument(b64);
        if (conLibretto) nuovo.sections.getFirst().pageSetup.bookFoldPrinting = true;
        return ctx.sync().then(function () { nuovo.open(); return ctx.sync(); });
      });
    }
    Word.run(function (ctx) {
      var ox = ctx.document.body.getOoxml();
      return ctx.sync().then(function () { return Stampa.originale(ox.value, opz); });
    }).then(function (b) {
      b64 = b;
      if (!librettoPossibile) return apri(false).then(function () { return false; });
      return apri(true).then(function () { return true; }, function () { return apri(false).then(function () { return false; }); });
    }).then(function (libretto) {
      var nome = Stampa.MODELLI[opz.modello].nome + (opz.interlineato ? ", interlineato" : "");
      if (a3) {
        stato("statoStampa", "Originale creato (" + nome + ") nella nuova finestra. Se lì le pagine non appaiono come A4 affiancati, " +
          "apri Stampa originali in quella finestra e premi \"Imposta il libretto A3 su questo documento\".", libretto ? "ok" : "err");
      } else stato("statoStampa", "Originale creato (" + nome + "): è nella nuova finestra.", "ok");
    }).catch(function (e) {
      stato("statoStampa", "Non riuscito: " + (e && e.message ? e.message : e), "err");
    });
  }

  // aggiunge lo spazio unificatore in fondo ai paragrafi selezionati (se non c'è già)
  function nonInterlineare() {
    if (!inWord) { stato("statoSel", "Il pannello non è aperto dentro Word.", "err"); return; }
    Word.run(function (ctx) {
      var pars = ctx.document.getSelection().paragraphs;
      pars.load("items/text");
      return ctx.sync().then(function () {
        var n = 0;
        pars.items.forEach(function (p) {
          if (/ $/.test(p.text)) return;
          p.insertText(" ", "End");
          n++;
        });
        return ctx.sync().then(function () { return n; });
      });
    }).then(function (n) {
      stato("statoSel", n ? (n === 1 ? "1 paragrafo non sarà interlineato." : n + " paragrafi non saranno interlineati.") : "I paragrafi selezionati erano già esclusi.", "ok");
    }).catch(function (e) { stato("statoSel", "Non riuscito: " + (e && e.message ? e.message : e), "err"); });
  }

  function supporta(set, v) {
    var req = window.Office && Office.context && Office.context.requirements;
    return !!(req && req.isSetSupported(set, v));
  }

  // libretto sul documento attivo (l'originale aperto nella nuova finestra)
  function impostaLibretto() {
    if (!inWord) { stato("statoLibretto", "Il pannello non è aperto dentro Word.", "err"); return; }
    if (!supporta("WordApiDesktop", "1.3")) {
      stato("statoLibretto", "Questa versione di Word non permette al pannello di impostare il libretto: fallo a mano come qui sotto (un minuto).", "err");
      $("manuale").hidden = false; return;
    }
    Word.run(function (ctx) {
      var ps = ctx.document.sections.getFirst().pageSetup;
      ps.load("pageWidth,pageHeight,orientation");
      return ctx.sync().then(function () {
        // il foglio è l'A3 orizzontale; il libretto lo divide in due pagine A4
        if (ps.pageWidth < ps.pageHeight) ps.orientation = "Landscape";
        ps.bookFoldPrinting = true;
        return ctx.sync();
      }).then(function () {
        ps.load("bookFoldPrinting,pageWidth,pageHeight");
        return ctx.sync().then(function () { return ps.bookFoldPrinting; });
      });
    }).then(function (ok) {
      if (ok) stato("statoLibretto", "Libretto impostato: a video ogni pagina è un A4. Stampa fronte-retro, lato corto.", "ok");
      else { stato("statoLibretto", "Word non ha accettato l'impostazione: fallo a mano come qui sotto.", "err"); $("manuale").hidden = false; }
    }).catch(function (e) {
      stato("statoLibretto", "Non riuscito (" + (e && e.message ? e.message : e) + "): fallo a mano come qui sotto.", "err");
      $("manuale").hidden = false;
    });
  }

  function collega() {
    try {
      var o = JSON.parse(window.localStorage.getItem("stampa.opzioni") || "null");
      if (o) { $("modelloStampa").value = o.modello; $("interlineato").checked = o.interlineato !== false; $("uniforma").checked = o.uniforma !== false; }
    } catch (e) { /* ignora */ }
    $("creaOriginale").addEventListener("click", creaOriginale);
    $("nonInterlineare").addEventListener("click", nonInterlineare);
    $("impostaLibretto").addEventListener("click", impostaLibretto);
    if (inWord) $("diagnosi").textContent = "Funzioni di Word disponibili: impostazione pagina " + (supporta("WordApiDesktop", "1.3") ? "sì" : "no") +
      ", documento nuovo " + (supporta("WordApi", "1.3") ? "sì" : "no") + ", documento nascosto " + (supporta("WordApiHiddenDocument", "1.3") ? "sì" : "no") + ".";
  }

  var pronto = { dom: false, office: !(window.Office && Office.onReady) };
  function avvia() { if (pronto.dom && pronto.office) collega(); }
  document.addEventListener("DOMContentLoaded", function () { pronto.dom = true; avvia(); });
  if (window.Office && Office.onReady) {
    Office.onReady(function (info) { inWord = !!(info && info.host === Office.HostType.Word); pronto.office = true; avvia(); });
  }
})();
