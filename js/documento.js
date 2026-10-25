/*
 * Generazione del documento finale (anteprima, stampa/PDF, Word).
 * Il piano viene prima tradotto in una sequenza di blocchi (titoli,
 * paragrafi, elenchi, tabelle, firma); da questa derivano sia l'HTML
 * dell'anteprima sia il file Word (vedi word.js), così l'aspetto è lo stesso.
 * Tutto il testo inserito dall'utente viene sempre passato da esc().
 */
(function (root) {
  'use strict';

  var model = root.PDL && root.PDL.model;
  if (!model && typeof require === 'function') model = require('./model.js');

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function testo(s) { return String(s == null ? '' : s); }

  /* ---------- Costruzione dei blocchi ---------- */

  function par(t, classe) { return { t: 'p', runs: [{ testo: t }], classe: classe }; }

  function paragrafi(s) {
    return testo(s).split(/\n\s*\n/).map(function (b) { return b.trim(); }).filter(Boolean)
      .map(function (b) { return par(b); });
  }

  var VUOTO = par('Sezione non compilata.', 'vuoto');

  // Voci selezionate nell'ordine del modello, seguite da eventuali voci
  // non più presenti nel modello e dalle voci aggiunte a mano ("altro").
  function vociSelezionate(v, opzioni) {
    var sel = v.sel || [];
    var voci = opzioni.filter(function (o) { return sel.indexOf(o) >= 0; });
    sel.forEach(function (s) { if (voci.indexOf(s) < 0) voci.push(s); });
    return voci.concat(model.righe(v.altro));
  }

  function lista(voci, numerata) {
    return voci.length ? [{ t: 'lista', numerata: !!numerata, voci: voci }] : [];
  }

  function bloccoSituazione(piano, sez, modello) {
    var v = model.valoreSezione(piano, sez);
    var n = v.num;
    var out = [];
    var tot = Number(n.alunni) || 0;
    if (tot) {
      var runs = [{ testo: 'La classe è composta da ' }, { testo: String(tot), b: true }, { testo: ' alunni' }];
      if (n.maschi || n.femmine) runs.push({ testo: ' (' + (Number(n.maschi) || 0) + ' maschi, ' + (Number(n.femmine) || 0) + ' femmine)' });
      runs.push({ testo: '.' });
      out.push({ t: 'p', runs: runs });
    }
    var altri = model.CAMPI_SITUAZIONE.slice(3).filter(function (c) { return Number(n[c.id]) > 0; });
    if (altri.length) {
      out.push({
        t: 'tabella', classe: 'orizzontale',
        intestazione: altri.map(function (c) { return c.etichetta; }),
        righe: [altri.map(function (c) { return testo(n[c.id]); })]
      });
    }
    if (modello.livelli.some(function (l) { return Number(v.livelli[l]) > 0; })) {
      out.push(par('Fasce di livello', 'sotto'));
      out.push({
        t: 'tabella', classe: 'orizzontale',
        intestazione: modello.livelli.slice(),
        righe: [modello.livelli.map(function (l) { return String(Number(v.livelli[l]) || 0); })]
      });
    }
    var strumenti = vociSelezionate(v, sez.opzioni);
    if (strumenti.length) out.push(par('Strumenti di rilevazione', 'sotto'));
    out = out.concat(lista(strumenti));
    if (testo(v.note).trim()) out = out.concat([par('Osservazioni sulla classe', 'sotto')], paragrafi(v.note));
    return out.length ? out : [VUOTO];
  }

  function bloccoChecklist(piano, sez, modello) {
    var v = model.valoreSezione(piano, sez);
    var out = lista(vociSelezionate(v, model.opzioniSezione(sez, piano, modello)), sez.tipo === 'traguardi');
    var campi = (sez.campi || []).filter(function (c) { return testo((v.campi || {})[c.id]).trim(); });
    if (campi.length) {
      out.push({
        t: 'tabella', classe: 'campi', intestazioneColonna: true, larghezze: [60, 40],
        righe: campi.map(function (c) { return [c.etichetta, testo(v.campi[c.id])]; })
      });
    }
    out = out.concat(paragrafi(v.note));
    return out.length ? out : [VUOTO];
  }

  function bloccoUda(piano, sez) {
    var unita = model.valoreSezione(piano, sez).unita.filter(function (u) {
      return ['titolo', 'obiettivi', 'contenuti'].some(function (k) { return testo(u[k]).trim(); });
    });
    if (!unita.length) return [VUOTO];
    return [{
      t: 'tabella', classe: 'uda',
      intestazione: ['N.', 'Unità e tempi', 'Nuclei tematici', 'Obiettivi di apprendimento', 'Contenuti e attività'],
      larghezze: [5, 21, 18, 28, 28],
      righe: unita.map(function (u, i) {
        var tempi = [u.periodo, u.ore ? u.ore + ' ore' : ''].filter(Boolean).join(' · ');
        var prima = [{ testo: testo(u.titolo).trim(), b: true }];
        if (tempi) prima.push({ testo: '\n' + tempi, i: true });
        return [String(i + 1), prima, (u.nuclei || []).join('; '), testo(u.obiettivi).trim(), testo(u.contenuti).trim()];
      })
    }];
  }

  function corpoSezione(piano, sez, modello) {
    switch (sez.tipo) {
      case 'situazione': return bloccoSituazione(piano, sez, modello);
      case 'checklist':
      case 'traguardi': return bloccoChecklist(piano, sez, modello);
      case 'uda': return bloccoUda(piano, sez);
      case 'testo': {
        var p = paragrafi(model.valoreSezione(piano, sez).note);
        return p.length ? p : [VUOTO];
      }
      default: return [];
    }
  }

  function intestazione(piano, modello) {
    var sc = modello.scuola;
    var righe = [];
    if (sc.nome) righe.push({ testo: sc.nome, stile: 'scuola' });
    if (sc.sottotitolo) righe.push({ testo: sc.sottotitolo, stile: 'sottotitolo' });
    righe.push({ testo: 'Piano di lavoro annuale', stile: 'titolo' });
    righe.push({ testo: 'Anno scolastico ' + (piano.anno || sc.annoScolastico), stile: 'anno' });
    var dati = model.CAMPI_DATI.filter(function (c) {
      return c.id !== 'anno' && c.id !== 'sezione' && testo(piano[c.id]).trim();
    }).map(function (c) {
      return [c.etichetta, c.id === 'classe' ? piano.classe + (piano.sezione ? ' ' + piano.sezione : '') : testo(piano[c.id])];
    });
    return [{ t: 'testata', righe: righe }, { t: 'tabella', classe: 'dati', intestazioneColonna: true, larghezze: [30, 70], righe: dati }];
  }

  function dataLeggibile(iso) {
    var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(testo(iso));
    return m ? m[3] + '/' + m[2] + '/' + m[1] : '';
  }

  function bloccoFirma(piano, modello) {
    var f = model.firmaValida(piano.firma);
    return {
      t: 'firma',
      luogo: modello.scuola.citta || '',
      data: f ? dataLeggibile(f.data) : '',
      docente: piano.docente || '',
      immagine: f ? f.immagine : ''
    };
  }

  function strutturaPiano(piano, modello) {
    var blocchi = intestazione(piano, modello);
    var n = 0;
    model.sezioniAttive(modello).forEach(function (s) {
      if (s.tipo === 'dati') return;
      n++;
      blocchi.push({ t: 'h2', testo: n + '. ' + s.titolo });
      blocchi = blocchi.concat(corpoSezione(piano, s, modello));
    });
    (piano.extra || []).forEach(function (x) {
      n++;
      blocchi.push({ t: 'h2', testo: n + '. ' + (testo(x.titolo).trim() || 'Sezione aggiuntiva') });
      var p = paragrafi(x.testo);
      blocchi = blocchi.concat(p.length ? p : [VUOTO]);
    });
    blocchi.push(bloccoFirma(piano, modello));
    return blocchi;
  }

  /* ---------- HTML ---------- */

  function htmlRuns(runs) {
    return runs.map(function (r) {
      var h = esc(r.testo).replace(/\n/g, '<br>');
      if (r.i) h = '<em>' + h + '</em>';
      if (r.b) h = '<strong>' + h + '</strong>';
      return h;
    }).join('');
  }

  function htmlCella(c) {
    return Array.isArray(c) ? htmlRuns(c) : esc(c).replace(/\n/g, '<br>');
  }

  function htmlTabella(b) {
    var larg = function (i) { return b.larghezze ? ' style="width:' + b.larghezze[i] + '%"' : ''; };
    var h = '<table class="doc-tab doc-' + b.classe + '">';
    if (b.intestazione) {
      h += '<thead><tr>' + b.intestazione.map(function (x, i) { return '<th' + larg(i) + '>' + esc(x) + '</th>'; }).join('') + '</tr></thead>';
    }
    h += '<tbody>' + b.righe.map(function (r) {
      return '<tr>' + r.map(function (c, i) {
        var tag = b.intestazioneColonna && i === 0 ? 'th' : 'td';
        return '<' + tag + (b.intestazione ? '' : larg(i)) + '>' + htmlCella(c) + '</' + tag + '>';
      }).join('') + '</tr>';
    }).join('') + '</tbody></table>';
    return h;
  }

  function htmlBlocco(b) {
    switch (b.t) {
      case 'testata':
        return '<header class="doc-testata">' + b.righe.map(function (r) {
          return r.stile === 'titolo' ? '<h1 class="doc-titolo">' + esc(r.testo) + '</h1>' : '<div class="doc-' + r.stile + '">' + esc(r.testo) + '</div>';
        }).join('') + '</header>';
      case 'h2': return '<h2>' + esc(b.testo) + '</h2>';
      case 'p': return '<p' + (b.classe ? ' class="doc-' + b.classe + '"' : '') + '>' + htmlRuns(b.runs) + '</p>';
      case 'lista': {
        var tag = b.numerata ? 'ol' : 'ul';
        return '<' + tag + '>' + b.voci.map(function (v) { return '<li>' + esc(v) + '</li>'; }).join('') + '</' + tag + '>';
      }
      case 'tabella': return htmlTabella(b);
      case 'firma':
        return '<div class="doc-firma"><div>' + (b.luogo ? esc(b.luogo) + ', ' : 'Luogo e data: ') +
          (b.data ? esc(b.data) : '____________________') + '</div>' +
          '<div class="doc-firma-docente">Il/La docente' +
          (b.immagine ? '<img class="doc-firma-img" alt="Firma" src="' + esc(b.immagine) + '">' : '') +
          '<span' + (b.immagine ? ' class="firmato"' : '') + '>' + esc(b.docente) + '</span></div></div>';
      default: return '';
    }
  }

  function renderPiano(piano, modello) {
    return '<article class="doc">' + strutturaPiano(piano, modello).map(htmlBlocco).join('') + '</article>';
  }

  // Stili del documento per anteprima e stampa (word.js li riproduce in Word).
  var CSS_DOCUMENTO = [
    '.doc{font-family:"Times New Roman",Georgia,serif;font-size:12pt;line-height:1.4;color:#111;}',
    '.doc-testata{text-align:center;margin-bottom:14pt;}',
    '.doc-scuola{font-size:14pt;font-weight:bold;text-transform:uppercase;}',
    '.doc-sottotitolo{font-size:11pt;}',
    '.doc-titolo{font-size:18pt;margin:12pt 0 2pt;text-transform:uppercase;letter-spacing:1px;}',
    '.doc-anno{font-size:12pt;}',
    '.doc h2{font-size:13pt;margin:16pt 0 6pt;padding-bottom:2pt;border-bottom:1px solid #444;}',
    '.doc-tab{border-collapse:collapse;width:100%;margin:6pt 0;}',
    '.doc-tab th,.doc-tab td{border:1px solid #666;padding:3pt 5pt;vertical-align:top;text-align:left;font-size:11pt;}',
    '.doc-tab th{background:#eee;}',
    '.doc-campi th{font-weight:normal;}',
    '.doc-uda td,.doc-uda th{font-size:10pt;}',
    '.doc-sotto{font-weight:bold;margin:8pt 0 2pt;}',
    '.doc-vuoto{color:#777;font-style:italic;}',
    '.doc ul,.doc ol{margin:4pt 0 4pt 18pt;padding:0;}',
    '.doc p{margin:4pt 0;}',
    '.doc-firma{margin-top:28pt;display:flex;flex-wrap:wrap;gap:12pt;justify-content:space-between;align-items:flex-end;break-inside:avoid;}',
    '.doc-firma-docente{text-align:center;min-width:160px;}',
    '.doc-firma-img{display:block;margin:4pt auto 0;max-width:200px;max-height:70px;}',
    '.doc-firma-docente span{display:inline-block;margin-top:24pt;border-top:1px solid #111;min-width:160px;padding-top:2pt;}',
    '.doc-firma-docente span.firmato{margin-top:0;}'
  ].join('\n');

  var api = {
    esc: esc,
    strutturaPiano: strutturaPiano,
    renderPiano: renderPiano,
    dataLeggibile: dataLeggibile,
    CSS_DOCUMENTO: CSS_DOCUMENTO
  };

  root.PDL = root.PDL || {};
  root.PDL.documento = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
