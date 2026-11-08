/*
 * Impaginazione a libretto: da un PDF di pagine A4 a un PDF A3 orizzontale con due pagine per facciata,
 * nell'ordine per piegare e rilegare (come RedAtto): 1° foglio fronte 16|1, retro 2|15, ...
 * Stampa fronte-retro, lato corto.
 */
(function (root) {
  "use strict";

  var A3 = { w: 1190.55, h: 841.89 };

  // ordine delle pagine: per ogni facciata [sinistra, destra] (indici da 0; -1 = pagina bianca)
  function ordine(n) {
    var tot = Math.ceil(n / 4) * 4, facciate = [];
    var p = function (i) { return i < n ? i : -1; };
    for (var i = 0; i < tot / 4; i++) {
      facciate.push([p(tot - 1 - 2 * i), p(2 * i)]);          // fronte
      facciate.push([p(2 * i + 1), p(tot - 2 - 2 * i)]);      // retro
    }
    return facciate;
  }

  // byte del PDF A4 -> byte del PDF A3 a libretto
  function impagina(byte, PDFLib) {
    PDFLib = PDFLib || root.PDFLib;
    return PDFLib.PDFDocument.load(byte).then(function (src) {
      return PDFLib.PDFDocument.create().then(function (out) {
        var n = src.getPageCount();
        var indici = [];
        for (var i = 0; i < n; i++) indici.push(i);
        return out.embedPdf(src, indici).then(function (pagine) {
          ordine(n).forEach(function (f) {
            var foglio = out.addPage([A3.w, A3.h]);
            f.forEach(function (idx, lato) {
              if (idx < 0) return;
              var pg = pagine[idx];
              var metà = A3.w / 2;
              var s = Math.min(metà / pg.width, A3.h / pg.height);
              var w = pg.width * s, h = pg.height * s;
              foglio.drawPage(pg, { x: lato * metà + (metà - w) / 2, y: (A3.h - h) / 2, width: w, height: h });
            });
          });
          return out.save();
        }).then(function (b) { return { byte: b, pagine: n, fogli: Math.ceil(n / 4) }; });
      });
    });
  }

  var api = { ordine: ordine, impagina: impagina };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Libretto = api;
})(this);
