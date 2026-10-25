/*
 * Salvataggio locale nel browser (localStorage) e caricamento del
 * modello d'istituto pubblicato insieme all'app (modello-istituto.json).
 */
(function (root) {
  'use strict';

  var model = root.PDL.model;
  var CHIAVE_MODELLO = 'pdl.modello';
  var CHIAVE_PIANI = 'pdl.piani';
  var CHIAVE_PREFERENZE = 'pdl.preferenze';

  function leggi(chiave) {
    try {
      var s = root.localStorage.getItem(chiave);
      return s ? JSON.parse(s) : null;
    } catch (e) {
      return null;
    }
  }

  function scrivi(chiave, valore) {
    try {
      root.localStorage.setItem(chiave, JSON.stringify(valore));
      return true;
    } catch (e) {
      return false;
    }
  }

  // Alcuni contesti (riquadri protetti, certe finestre anonime) non permettono di salvare.
  function memoriaDisponibile() {
    try {
      root.localStorage.setItem('pdl.prova', '1');
      root.localStorage.removeItem('pdl.prova');
      return true;
    } catch (e) {
      return false;
    }
  }

  function caricaModello() {
    var m = leggi(CHIAVE_MODELLO);
    return m ? model.normalizzaModello(m) : null;
  }

  function salvaModello(m) {
    return scrivi(CHIAVE_MODELLO, m);
  }

  function caricaPiani() {
    var p = leggi(CHIAVE_PIANI);
    return Array.isArray(p) ? p.map(model.normalizzaPiano).filter(Boolean) : [];
  }

  function salvaPiani(piani) {
    return scrivi(CHIAVE_PIANI, piani);
  }

  function preferenze() {
    return leggi(CHIAVE_PREFERENZE) || {};
  }

  function salvaPreferenza(nome, valore) {
    var p = preferenze();
    p[nome] = valore;
    scrivi(CHIAVE_PREFERENZE, p);
  }

  // Il modello d'istituto può essere pubblicato accanto a index.html.
  // Quando l'app è aperta come file locale la richiesta fallisce: nessun problema.
  function caricaModelloIstituto() {
    if (!root.fetch || root.location.protocol === 'file:') return Promise.resolve(null);
    return root.fetch('modello-istituto.json', { cache: 'no-store' })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (json) {
        if (!json) return null;
        var letto = model.leggiPacchetto(json);
        return letto.tipo === 'modello' ? letto.modello : null;
      })
      .catch(function () { return null; });
  }

  root.PDL.store = {
    memoriaDisponibile: memoriaDisponibile,
    caricaModello: caricaModello,
    salvaModello: salvaModello,
    caricaPiani: caricaPiani,
    salvaPiani: salvaPiani,
    preferenze: preferenze,
    salvaPreferenza: salvaPreferenza,
    caricaModelloIstituto: caricaModelloIstituto
  };
})(window);
