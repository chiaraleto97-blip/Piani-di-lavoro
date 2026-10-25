/*
 * Logica dei dati: modello d'istituto, piani di lavoro, completamento.
 * Nessuna dipendenza dal DOM: il file è usato sia dal browser sia dai test.
 */
(function (root) {
  'use strict';

  var defaults = root.PDL && root.PDL.defaults;
  if (!defaults && typeof require === 'function') defaults = require('./defaults.js');

  var FORMATO_MODELLO = 'modello-piano-di-lavoro';
  var FORMATO_PIANI = 'piani-di-lavoro';

  var CAMPI_DATI = [
    { id: 'docente', etichetta: 'Docente', obbligatorio: true },
    { id: 'disciplina', etichetta: 'Disciplina', obbligatorio: true },
    { id: 'classe', etichetta: 'Classe', obbligatorio: true },
    { id: 'sezione', etichetta: 'Sezione', obbligatorio: true },
    { id: 'anno', etichetta: 'Anno scolastico', obbligatorio: true },
    { id: 'ore', etichetta: 'Ore settimanali' },
    { id: 'libro', etichetta: 'Libro di testo' },
    { id: 'plesso', etichetta: 'Plesso / sede' }
  ];

  var CAMPI_SITUAZIONE = [
    { id: 'alunni', etichetta: 'Alunni' },
    { id: 'maschi', etichetta: 'Maschi' },
    { id: 'femmine', etichetta: 'Femmine' },
    { id: 'ripetenti', etichetta: 'Ripetenti' },
    { id: 'h104', etichetta: 'Con disabilità (L. 104/92)' },
    { id: 'dsa', etichetta: 'Con DSA' },
    { id: 'bes', etichetta: 'Altri BES' },
    { id: 'nai', etichetta: 'Non italofoni / NAI' }
  ];

  var MESI = ['Settembre', 'Ottobre', 'Novembre', 'Dicembre', 'Gennaio', 'Febbraio',
    'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Intero anno'];

  function clona(x) { return JSON.parse(JSON.stringify(x)); }

  function uid() {
    return 'p' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  }

  function testo(x) { return x == null ? '' : String(x); }

  function righe(s) {
    return testo(s).split(/\r?\n/).map(function (r) { return r.trim(); }).filter(Boolean);
  }

  function isOggetto(x) { return x !== null && typeof x === 'object' && !Array.isArray(x); }

  function elencoStringhe(a) {
    return Array.isArray(a) ? a.map(testo).map(function (s) { return s.trim(); }).filter(Boolean) : [];
  }

  /* ---------- Modello ---------- */

  // Completa un modello (anche parziale o di una versione precedente)
  // con i valori predefiniti, così che l'interfaccia possa fidarsi della forma.
  function normalizzaModello(m) {
    var base = defaults.modelloPredefinito();
    if (!isOggetto(m)) return base;
    var out = {
      schema: 1,
      aggiornato: testo(m.aggiornato),
      bloccato: !!m.bloccato,
      scuola: Object.assign({}, base.scuola, isOggetto(m.scuola) ? m.scuola : {}),
      periodi: elencoStringhe(m.periodi).length ? elencoStringhe(m.periodi) : base.periodi,
      livelli: elencoStringhe(m.livelli).length ? elencoStringhe(m.livelli) : base.livelli,
      sezioni: [],
      discipline: []
    };

    var visti = {};
    (Array.isArray(m.sezioni) ? m.sezioni : base.sezioni).forEach(function (s) {
      if (!isOggetto(s) || !s.id || visti[s.id]) return;
      if (defaults.TIPI_SEZIONE.indexOf(s.tipo) < 0) return;
      visti[s.id] = true;
      out.sezioni.push({
        id: testo(s.id),
        tipo: s.tipo,
        titolo: testo(s.titolo) || 'Sezione',
        attiva: s.attiva !== false,
        facoltativa: !!s.facoltativa,
        personalizzata: !!s.personalizzata,
        guida: testo(s.guida),
        opzioni: elencoStringhe(s.opzioni),
        campi: Array.isArray(s.campi) ? s.campi.filter(function (c) { return isOggetto(c) && c.id; }).map(function (c) {
          return { id: testo(c.id), etichetta: testo(c.etichetta) || testo(c.id), tipo: c.tipo === 'numero' ? 'numero' : 'testo' };
        }) : [],
        testo: testo(s.testo)
      });
    });
    // La sezione "dati" è sempre presente, attiva e in prima posizione.
    var iDati = out.sezioni.findIndex(function (s) { return s.tipo === 'dati'; });
    var dati = iDati >= 0 ? out.sezioni.splice(iDati, 1)[0] : clona(base.sezioni[0]);
    dati.attiva = true;
    out.sezioni.unshift(dati);

    var nomi = {};
    (Array.isArray(m.discipline) ? m.discipline : base.discipline).forEach(function (d) {
      if (!isOggetto(d)) return;
      var nome = testo(d.nome).trim();
      if (!nome || nomi[nome.toLowerCase()]) return;
      nomi[nome.toLowerCase()] = true;
      out.discipline.push({
        nome: nome,
        ore: d.ore === '' || d.ore == null ? '' : Number(d.ore) || '',
        nuclei: elencoStringhe(d.nuclei),
        traguardi: elencoStringhe(d.traguardi)
      });
    });
    return out;
  }

  function trovaDisciplina(modello, nome) {
    var n = testo(nome).trim().toLowerCase();
    for (var i = 0; i < modello.discipline.length; i++) {
      if (modello.discipline[i].nome.toLowerCase() === n) return modello.discipline[i];
    }
    return null;
  }

  function sezioniAttive(modello) {
    return modello.sezioni.filter(function (s) { return s.attiva; });
  }

  // Opzioni proposte per una sezione: per i traguardi dipendono dalla disciplina.
  function opzioniSezione(sezione, piano, modello) {
    if (sezione.tipo === 'traguardi') {
      var d = trovaDisciplina(modello, piano.disciplina);
      return d ? d.traguardi : [];
    }
    return sezione.opzioni || [];
  }

  /* ---------- Piani ---------- */

  function valoreVuoto(sezione) {
    switch (sezione.tipo) {
      case 'situazione': return { num: {}, livelli: {}, sel: [], altro: '', note: sezione.testo || '' };
      case 'checklist': return { sel: [], altro: '', note: sezione.testo || '', campi: {} };
      case 'traguardi': return { sel: [], altro: '' };
      case 'uda': return { unita: [] };
      case 'testo': return { note: sezione.testo || '' };
      default: return {};
    }
  }

  // Restituisce (creandolo se manca) il valore di una sezione nel piano,
  // garantendone la forma attesa.
  function valoreSezione(piano, sezione) {
    piano.valori = isOggetto(piano.valori) ? piano.valori : {};
    var v = piano.valori[sezione.id];
    if (!isOggetto(v)) v = piano.valori[sezione.id] = valoreVuoto(sezione);
    if ('sel' in valoreVuoto(sezione) && !Array.isArray(v.sel)) v.sel = [];
    if (sezione.tipo === 'uda' && !Array.isArray(v.unita)) v.unita = [];
    if (sezione.tipo === 'situazione') {
      if (!isOggetto(v.num)) v.num = {};
      if (!isOggetto(v.livelli)) v.livelli = {};
    }
    if (sezione.tipo === 'checklist' && !isOggetto(v.campi)) v.campi = {};
    return v;
  }

  function nuovaUnita() {
    return { titolo: '', periodo: '', ore: '', nuclei: [], obiettivi: '', contenuti: '' };
  }

  function nuovoPiano(modello, dati) {
    dati = dati || {};
    var disc = trovaDisciplina(modello, dati.disciplina);
    var ora = new Date().toISOString();
    var p = {
      id: uid(),
      creato: ora,
      modificato: ora,
      docente: testo(dati.docente).trim(),
      disciplina: testo(dati.disciplina).trim(),
      classe: testo(dati.classe).trim(),
      sezione: testo(dati.sezione).trim().toUpperCase(),
      anno: testo(dati.anno).trim() || modello.scuola.annoScolastico,
      ore: dati.ore != null && dati.ore !== '' ? testo(dati.ore) : (disc && disc.ore !== '' ? testo(disc.ore) : ''),
      libro: testo(dati.libro).trim(),
      plesso: testo(dati.plesso).trim(),
      valori: {}
    };
    modello.sezioni.forEach(function (s) {
      if (s.tipo !== 'dati') p.valori[s.id] = valoreVuoto(s);
    });
    return p;
  }

  // Crea un nuovo piano copiando i contenuti di uno esistente.
  // La situazione di partenza è specifica della classe e viene azzerata.
  function copiaContenuti(destinazione, sorgente, modello) {
    var valori = clona(sorgente.valori || {});
    modello.sezioni.forEach(function (s) {
      if (s.tipo === 'situazione') valori[s.id] = valoreVuoto(s);
    });
    destinazione.valori = valori;
    destinazione.extra = normalizzaExtra(clona(sorgente.extra || []));
    if (!destinazione.libro) destinazione.libro = testo(sorgente.libro);
    // La firma non si copia: il nuovo piano va firmato di nuovo.
    return destinazione;
  }

  // Se si chiede "A, B, C" si ottiene un piano per ogni sezione.
  function elencoSezioniClasse(s) {
    var parti = testo(s).toUpperCase().split(/[\s,;\/]+/).filter(Boolean);
    var uniche = [];
    parti.forEach(function (x) { if (uniche.indexOf(x) < 0) uniche.push(x); });
    return uniche;
  }

  function titoloPiano(p) {
    var classe = (testo(p.classe) + testo(p.sezione)).trim();
    return [p.disciplina || 'Disciplina non indicata', classe ? 'classe ' + classe : ''].filter(Boolean).join(' – ');
  }

  function nomeFile(p, estensione) {
    var parti = ['Piano di lavoro', p.disciplina, (testo(p.classe) + testo(p.sezione)), testo(p.anno).replace('/', '-')];
    return parti.filter(Boolean).join(' - ').replace(/[\\/:*?"<>|]+/g, '') + '.' + estensione;
  }

  function sezioneCompleta(piano, sezione) {
    if (sezione.tipo === 'dati') {
      return CAMPI_DATI.every(function (c) { return !c.obbligatorio || testo(piano[c.id]).trim(); });
    }
    var v = valoreSezione(piano, sezione);
    switch (sezione.tipo) {
      case 'situazione': return Number(v.num.alunni) > 0;
      case 'checklist':
      case 'traguardi': return v.sel.length > 0 || righe(v.altro).length > 0;
      case 'uda': return v.unita.some(function (u) { return testo(u.titolo).trim(); });
      case 'testo': return !!testo(v.note).trim();
      default: return true;
    }
  }

  function avanzamento(piano, modello) {
    var conteggiate = sezioniAttive(modello).filter(function (s) { return !s.facoltativa; });
    var fatte = conteggiate.filter(function (s) { return sezioneCompleta(piano, s); }).length;
    return {
      fatte: fatte,
      totale: conteggiate.length,
      percentuale: conteggiate.length ? Math.round(fatte * 100 / conteggiate.length) : 100
    };
  }

  // Controlli di coerenza sulla situazione di partenza (avvisi, non errori).
  function avvisiSituazione(v, livelli) {
    var n = v.num || {};
    var avvisi = [];
    var tot = Number(n.alunni) || 0;
    var m = Number(n.maschi) || 0;
    var f = Number(n.femmine) || 0;
    if (tot && (m || f) && m + f !== tot) {
      avvisi.push('Maschi + femmine (' + (m + f) + ') non corrisponde al numero di alunni (' + tot + ').');
    }
    var sommaLivelli = (livelli || []).reduce(function (acc, l) { return acc + (Number((v.livelli || {})[l]) || 0); }, 0);
    if (tot && sommaLivelli && sommaLivelli !== tot) {
      avvisi.push('La somma delle fasce di livello (' + sommaLivelli + ') non corrisponde al numero di alunni (' + tot + ').');
    }
    return avvisi;
  }

  function normalizzaPiano(p) {
    if (!isOggetto(p)) return null;
    var ora = new Date().toISOString();
    var out = {
      id: testo(p.id) || uid(),
      creato: testo(p.creato) || ora,
      modificato: testo(p.modificato) || ora,
      valori: isOggetto(p.valori) ? p.valori : {}
    };
    CAMPI_DATI.forEach(function (c) { out[c.id] = testo(p[c.id]); });
    var extra = normalizzaExtra(p.extra);
    if (extra.length) out.extra = extra;
    var firma = firmaValida(p.firma);
    if (firma) out.firma = firma;
    return out;
  }

  // Sezioni non previste dal modello d'istituto (di solito da un piano importato).
  function normalizzaExtra(extra) {
    return Array.isArray(extra) ? extra.filter(isOggetto).map(function (x) {
      return { id: testo(x.id) || uid(), titolo: testo(x.titolo), testo: testo(x.testo) };
    }) : [];
  }

  function nuovaSezioneExtra(titolo, contenuto) {
    return { id: uid(), titolo: testo(titolo), testo: testo(contenuto) };
  }

  // Accetta solo immagini PNG in base64: i piani possono arrivare da file di altri.
  function firmaValida(f) {
    if (!isOggetto(f) || !/^data:image\/png;base64,[A-Za-z0-9+\/]+=*$/.test(testo(f.immagine))) return null;
    return {
      immagine: f.immagine,
      data: /^\d{4}-\d{2}-\d{2}$/.test(testo(f.data)) ? f.data : new Date().toISOString().slice(0, 10),
      momento: testo(f.momento)
    };
  }

  // Vero se il piano è stato modificato dopo la firma.
  function modificatoDopoFirma(p) {
    var f = firmaValida(p.firma);
    return !!(f && f.momento && testo(p.modificato) > f.momento);
  }

  /* ---------- Import / export ---------- */

  function pacchettoModello(modello) {
    return { formato: FORMATO_MODELLO, versione: 1, modello: modello };
  }

  function pacchettoPiani(piani) {
    return { formato: FORMATO_PIANI, versione: 1, esportato: new Date().toISOString(), piani: piani };
  }

  // Riconosce il contenuto di un file importato.
  function leggiPacchetto(json) {
    var dati;
    try { dati = typeof json === 'string' ? JSON.parse(json) : json; } catch (e) {
      return { errore: 'Il file non è un file JSON valido.' };
    }
    if (!isOggetto(dati)) return { errore: 'Contenuto del file non riconosciuto.' };
    if (dati.formato === FORMATO_MODELLO && isOggetto(dati.modello)) {
      return { tipo: 'modello', modello: normalizzaModello(dati.modello) };
    }
    if (dati.formato === FORMATO_PIANI && Array.isArray(dati.piani)) {
      return { tipo: 'piani', piani: dati.piani.map(normalizzaPiano).filter(Boolean) };
    }
    return { errore: 'Il file non contiene né un modello né dei piani di lavoro.' };
  }

  // Unisce piani importati a quelli esistenti: a parità di id prevale il più recente.
  function unisciPiani(esistenti, importati) {
    var perId = {};
    var risultato = esistenti.slice();
    risultato.forEach(function (p, i) { perId[p.id] = i; });
    var stat = { aggiunti: 0, aggiornati: 0, ignorati: 0 };
    importati.forEach(function (p) {
      if (!(p.id in perId)) {
        perId[p.id] = risultato.length;
        risultato.push(p);
        stat.aggiunti++;
      } else if (p.modificato > risultato[perId[p.id]].modificato) {
        risultato[perId[p.id]] = p;
        stat.aggiornati++;
      } else {
        stat.ignorati++;
      }
    });
    return { piani: risultato, stat: stat };
  }

  var api = {
    CAMPI_DATI: CAMPI_DATI,
    CAMPI_SITUAZIONE: CAMPI_SITUAZIONE,
    MESI: MESI,
    clona: clona,
    uid: uid,
    righe: righe,
    normalizzaModello: normalizzaModello,
    trovaDisciplina: trovaDisciplina,
    sezioniAttive: sezioniAttive,
    opzioniSezione: opzioniSezione,
    valoreVuoto: valoreVuoto,
    valoreSezione: valoreSezione,
    nuovaUnita: nuovaUnita,
    nuovoPiano: nuovoPiano,
    copiaContenuti: copiaContenuti,
    elencoSezioniClasse: elencoSezioniClasse,
    titoloPiano: titoloPiano,
    nomeFile: nomeFile,
    sezioneCompleta: sezioneCompleta,
    avanzamento: avanzamento,
    avvisiSituazione: avvisiSituazione,
    normalizzaPiano: normalizzaPiano,
    normalizzaExtra: normalizzaExtra,
    nuovaSezioneExtra: nuovaSezioneExtra,
    firmaValida: firmaValida,
    modificatoDopoFirma: modificatoDopoFirma,
    pacchettoModello: pacchettoModello,
    pacchettoPiani: pacchettoPiani,
    leggiPacchetto: leggiPacchetto,
    unisciPiani: unisciPiani
  };

  root.PDL = root.PDL || {};
  root.PDL.model = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
