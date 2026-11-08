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
      if (!a3) return apri(false);
      return apri(true).catch(function () { return apri(false); });
    }).then(function () {
      var nome = Stampa.MODELLI[opz.modello].nome + (opz.interlineato ? ", interlineato" : "");
      stato("statoStampa", "Originale creato (" + nome + "): è nella nuova finestra." + (a3 ? " Stampa fronte-retro, lato corto." : ""), "ok");
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

  // piega libro sul documento attivo (l'originale A3 nella nuova finestra)
  function impostaLibretto() {
    if (!inWord) { stato("statoLibretto", "Il pannello non è aperto dentro Word.", "err"); return; }
    Word.run(function (ctx) {
      var ps = ctx.document.sections.getFirst().pageSetup;
      ps.bookFoldPrinting = true;
      return ctx.sync().then(function () { ps.load("bookFoldPrinting"); return ctx.sync(); }).then(function () { return ps.bookFoldPrinting; });
    }).then(function (ok) {
      stato("statoLibretto", ok ? "Fatto: ogni pagina è un A4. Stampa fronte-retro, lato corto." : "Word non ha accettato l'impostazione.", ok ? "ok" : "err");
    }).catch(function (e) {
      stato("statoLibretto", "Word non l'ha accettato (" + (e && e.message ? e.message : e) + ").", "err");
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

  }

  var pronto = { dom: false, office: !(window.Office && Office.onReady) };
  function avvia() { if (pronto.dom && pronto.office) collega(); }
  document.addEventListener("DOMContentLoaded", function () { pronto.dom = true; avvia(); });
  if (window.Office && Office.onReady) {
    Office.onReady(function (info) { inWord = !!(info && info.host === Office.HostType.Word); pronto.office = true; avvia(); });
  }
})();
