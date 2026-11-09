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
    Word.run(function (ctx) {
      var ox = ctx.document.body.getOoxml();
      return ctx.sync().then(function () { return Stampa.originale(ox.value, opz); }).then(function (b64) {
        ctx.application.createDocument(b64).open();
        return ctx.sync();
      });
    }).then(function () {
      var nome = Stampa.MODELLI[opz.modello].nome + (opz.interlineato ? ", interlineato" : "");
      stato("statoStampa", "Originale creato (" + nome + "): è nella nuova finestra. Stampa con la preimpostazione «" +
        (a3 ? "Originale A3 libretto" : "Originale A4") + "».", "ok");
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

  function collega() {
    try {
      var o = JSON.parse(window.localStorage.getItem("stampa.opzioni") || "null");
      if (o) { $("modelloStampa").value = o.modello; $("interlineato").checked = o.interlineato !== false; $("uniforma").checked = o.uniforma !== false; }
    } catch (e) { /* ignora */ }
    $("creaOriginale").addEventListener("click", creaOriginale);
    $("nonInterlineare").addEventListener("click", nonInterlineare);

  }

  var pronto = { dom: false, office: !(window.Office && Office.onReady) };
  function avvia() { if (pronto.dom && pronto.office) collega(); }
  document.addEventListener("DOMContentLoaded", function () { pronto.dom = true; avvia(); });
  if (window.Office && Office.onReady) {
    Office.onReady(function (info) { inWord = !!(info && info.host === Office.HostType.Word); pronto.office = true; avvia(); });
  }
})();
