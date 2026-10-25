/*
 * Interfaccia dell'applicazione: elenco dei piani, compilazione,
 * anteprima/stampa, modello d'istituto e guida.
 */
(function () {
  'use strict';

  const { model: M, documento: D, store: S, word: W, importa: I, firma: F } = window.PDL;
  const app = document.getElementById('app');

  const stato = {
    modello: null,
    modelloSalvato: false,
    piani: [],
    selezionati: new Set(),
    filtri: { testo: '', anno: '', classe: '' },
    importazione: null
  };

  const FORMATI_PIANO = '.docx,.odt,.txt,.doc,.pdf';

  const NOMI_CLASSI = { 1: 'Prima', 2: 'Seconda', 3: 'Terza' };
  const ETICHETTE_TIPO = {
    dati: 'Intestazione',
    situazione: 'Situazione di partenza',
    checklist: 'Elenco di voci',
    traguardi: 'Traguardi della disciplina',
    uda: 'Tabella UdA',
    testo: 'Testo libero'
  };

  /* ================= Utilità DOM ================= */

  function h(tag, attrs, ...figli) {
    const el = document.createElement(tag);
    let valore;
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
      else if (k === 'class') el.className = v;
      else if (k === 'html') el.innerHTML = v;
      else if (k === 'value') valore = v;
      else if (k === 'checked') el.checked = true;
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    }
    aggiungi(el, figli);
    // Il valore va impostato dopo le <option>, altrimenti le select lo ignorano.
    if (valore != null) el.value = valore;
    return el;
  }

  function aggiungi(el, figli) {
    figli.flat(Infinity).forEach((f) => {
      if (f == null || f === false) return;
      el.append(f instanceof Node ? f : String(f));
    });
  }

  function bottone(testo, onclick, classe = 'btn', attrs = {}) {
    return h('button', Object.assign({ type: 'button', class: classe, onclick }, attrs), testo);
  }

  function campo(etichetta, controllo, classe = 'campo') {
    return h('label', { class: classe }, h('span', { class: 'etichetta' }, etichetta), controllo);
  }

  function areaTesto(valore, oninput, attrs = {}) {
    return h('textarea', Object.assign({ rows: 3, oninput: (e) => { adatta(e.target); oninput(e.target.value); } }, attrs), valore || '');
  }

  function adatta(ta) {
    ta.style.height = 'auto';
    ta.style.height = Math.min(ta.scrollHeight + 2, 600) + 'px';
  }

  function adattaTutte(radice = app) {
    radice.querySelectorAll('textarea').forEach(adatta);
  }

  function selezione(opzioni, valore, onchange, attrs = {}) {
    return h('select', Object.assign({ value: valore, onchange: (e) => onchange(e.target.value) }, attrs),
      opzioni.map((o) => (typeof o === 'string' ? h('option', { value: o }, o) : h('option', { value: o.valore }, o.testo))));
  }

  function toast(messaggio, tipo = 'ok') {
    const box = document.getElementById('avvisi');
    const el = h('div', { class: 'toast ' + tipo }, messaggio);
    box.append(el);
    setTimeout(() => el.classList.add('via'), 3800);
    setTimeout(() => el.remove(), 4300);
  }

  function scarica(nome, contenuto, tipo) {
    const url = URL.createObjectURL(new Blob([contenuto], { type: tipo }));
    const a = h('a', { href: url, download: nome });
    document.body.append(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }

  // Restituisce { nome, dati: Uint8Array } del file scelto.
  function chiediFile(accept) {
    return new Promise((risolvi) => {
      const input = h('input', { type: 'file', accept, class: 'nascosto' });
      input.addEventListener('change', () => {
        const file = input.files && input.files[0];
        input.remove();
        if (!file) return;
        file.arrayBuffer()
          .then((b) => risolvi({ nome: file.name, dati: new Uint8Array(b) }))
          .catch(() => toast('Impossibile leggere il file.', 'errore'));
      });
      document.body.append(input);
      input.click();
    });
  }

  function dataOggi() {
    const d = new Date();
    return new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
  }

  function dataBreve(iso) {
    const d = new Date(iso);
    if (isNaN(d)) return '';
    return d.toLocaleDateString('it-IT', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' ' +
      d.toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' });
  }

  function barraAvanzamento(av) {
    return h('div', { class: 'progresso', title: `${av.fatte} di ${av.totale} sezioni compilate` },
      h('div', { class: 'progresso-barra' }, h('span', { style: `width:${av.percentuale}%` })),
      h('span', { class: 'progresso-testo' }, av.percentuale + '%'));
  }

  function nomeClasse(p) {
    return (p.classe || '') + (p.sezione || '');
  }

  /* ================= Salvataggio ================= */

  let timerPiani = null;
  let timerModello = null;

  function salvaPianiOra() {
    clearTimeout(timerPiani);
    timerPiani = null;
    const ok = S.salvaPiani(stato.piani);
    const el = document.getElementById('stato-salvataggio');
    if (el) {
      el.textContent = ok ? 'Salvato alle ' + new Date().toLocaleTimeString('it-IT', { hour: '2-digit', minute: '2-digit' }) : 'Salvataggio non riuscito';
      el.classList.toggle('errore', !ok);
    }
    if (!ok) toast('Salvataggio non riuscito: la memoria del browser è piena o non disponibile. Esporta subito i piani.', 'errore');
    return ok;
  }

  function programmaSalvataggio() {
    const el = document.getElementById('stato-salvataggio');
    if (el) { el.textContent = 'Salvataggio…'; el.classList.remove('errore'); }
    clearTimeout(timerPiani);
    timerPiani = setTimeout(salvaPianiOra, 400);
  }

  function salvaModelloOra() {
    clearTimeout(timerModello);
    timerModello = null;
    stato.modelloSalvato = true;
    if (!S.salvaModello(stato.modello)) toast('Impossibile salvare il modello nel browser.', 'errore');
    aggiornaTestata();
  }

  function modelloModificato() {
    stato.modello.aggiornato = new Date().toISOString();
    clearTimeout(timerModello);
    timerModello = setTimeout(salvaModelloOra, 400);
  }

  window.addEventListener('pagehide', () => {
    if (timerPiani) salvaPianiOra();
    if (timerModello) salvaModelloOra();
  });

  /* ================= Navigazione ================= */

  function vai(hash) {
    if (location.hash === hash) render();
    else location.hash = hash;
  }

  function render(opzioni = {}) {
    if (timerModello) salvaModelloOra();
    const parti = location.hash.replace(/^#\/?/, '').split('/');
    const vista = parti[0] || 'piani';
    const sezioneNav = ['piano', 'anteprima', 'importa'].includes(vista) ? 'piani' : vista;
    document.querySelectorAll('[data-nav]').forEach((a) => {
      const attivo = a.dataset.nav === sezioneNav;
      a.classList.toggle('attivo', attivo);
      if (attivo) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    document.body.classList.toggle('modo-anteprima', vista === 'anteprima');
    const scroll = window.scrollY;
    app.replaceChildren();
    switch (vista) {
      case 'piano': vistaEditor(decodeURIComponent(parti[1] || '')); break;
      case 'anteprima': vistaAnteprima(decodeURIComponent(parti[1] || '').split(',')); break;
      case 'modello': vistaModello(); break;
      case 'guida': vistaGuida(); break;
      case 'importa': vistaImporta(); break;
      default: vistaElenco();
    }
    adattaTutte();
    window.scrollTo(0, opzioni.mantieniScroll ? scroll : 0);
  }

  function aggiornaTestata() {
    const el = document.getElementById('nome-scuola');
    if (el) el.textContent = stato.modello.scuola.nome || '';
  }

  /* ================= Elenco dei piani ================= */

  function intestazioneVista(titolo, sottotitolo, ...azioni) {
    return h('div', { class: 'intestazione-vista' },
      h('div', null, h('h1', null, titolo), sottotitolo ? h('p', { class: 'sottotitolo' }, sottotitolo) : null),
      azioni.length ? h('div', { class: 'azioni' }, azioni) : null);
  }

  function vistaElenco() {
    const n = stato.piani.length;
    app.append(intestazioneVista('I miei piani di lavoro',
      n ? `${n} ${n === 1 ? 'piano salvato' : 'piani salvati'} in questo browser` : null,
      n ? bottone('Backup di tutti i piani', () => esportaPiani(stato.piani, 'Backup piani di lavoro.json')) : null,
      bottone('Importa un piano già scritto', importaPianoEsistente),
      bottone('+ Nuovo piano', () => apriDialogoPiano(), 'btn primario')));

    if (!n) {
      app.append(benvenuto());
      return;
    }

    const anni = [...new Set(stato.piani.map((p) => p.anno).filter(Boolean))].sort().reverse();
    app.append(h('div', { class: 'filtri' },
      h('input', {
        type: 'search', placeholder: 'Cerca per disciplina, docente o classe…', 'aria-label': 'Cerca',
        value: stato.filtri.testo, oninput: (e) => { stato.filtri.testo = e.target.value; disegnaTabella(); }
      }),
      selezione([{ valore: '', testo: 'Tutti gli anni' }].concat(anni.map((a) => ({ valore: a, testo: 'A.S. ' + a }))),
        stato.filtri.anno, (v) => { stato.filtri.anno = v; disegnaTabella(); }, { 'aria-label': 'Anno scolastico' }),
      selezione([{ valore: '', testo: 'Tutte le classi' }, { valore: '1', testo: 'Classi prime' }, { valore: '2', testo: 'Classi seconde' }, { valore: '3', testo: 'Classi terze' }],
        stato.filtri.classe, (v) => { stato.filtri.classe = v; disegnaTabella(); }, { 'aria-label': 'Classe' })),
    h('div', { id: 'tabella-piani' }));
    disegnaTabella();
  }

  function pianiFiltrati() {
    const f = stato.filtri;
    const t = f.testo.trim().toLowerCase();
    return stato.piani
      .filter((p) => (!f.anno || p.anno === f.anno) && (!f.classe || p.classe === f.classe) &&
        (!t || [p.disciplina, p.docente, nomeClasse(p), p.libro].join(' ').toLowerCase().includes(t)))
      .sort((a, b) => b.anno.localeCompare(a.anno) || a.classe.localeCompare(b.classe) ||
        a.sezione.localeCompare(b.sezione) || a.disciplina.localeCompare(b.disciplina));
  }

  function disegnaTabella() {
    const box = document.getElementById('tabella-piani');
    if (!box) return;
    const ids = new Set(stato.piani.map((p) => p.id));
    stato.selezionati.forEach((id) => { if (!ids.has(id)) stato.selezionati.delete(id); });
    const piani = pianiFiltrati();
    const scelti = () => stato.piani.filter((p) => stato.selezionati.has(p.id));
    const tuttiScelti = piani.length > 0 && piani.every((p) => stato.selezionati.has(p.id));

    const barraSelezione = stato.selezionati.size ? h('div', { class: 'barra-selezione' },
      h('strong', null, `${stato.selezionati.size} selezionati`),
      bottone('Anteprima e stampa', () => vai('#/anteprima/' + scelti().map((p) => p.id).join(','))),
      bottone('Scarica in Word', () => scaricaWord(scelti())),
      bottone('Esporta file', () => esportaPiani(scelti(), `Piani di lavoro (${stato.selezionati.size}).json`)),
      bottone('Elimina', () => eliminaPiani(scelti()), 'btn pericolo'),
      bottone('Annulla selezione', () => { stato.selezionati.clear(); disegnaTabella(); }, 'btn-link')) : null;

    const righe = piani.map((p) => {
      const av = M.avanzamento(p, stato.modello);
      return h('tr', null,
        h('td', { class: 'col-check' }, h('input', {
          type: 'checkbox', 'aria-label': 'Seleziona ' + M.titoloPiano(p), checked: stato.selezionati.has(p.id),
          onchange: (e) => { e.target.checked ? stato.selezionati.add(p.id) : stato.selezionati.delete(p.id); disegnaTabella(); }
        })),
        h('td', null, h('a', { href: '#/piano/' + encodeURIComponent(p.id), class: 'link-piano' }, p.disciplina || 'Senza disciplina')),
        h('td', null, nomeClasse(p) || '—'),
        h('td', null, p.docente || '—'),
        h('td', null, p.anno),
        h('td', null, barraAvanzamento(av),
          M.firmaValida(p.firma) ? h('span', { class: 'badge-firma', title: 'Firmato il ' + D.dataLeggibile(p.firma.data) }, '✓ Firmato') : null),
        h('td', { class: 'data' }, dataBreve(p.modificato)),
        h('td', { class: 'col-azioni' },
          bottone('Anteprima', () => vai('#/anteprima/' + encodeURIComponent(p.id)), 'btn piccolo'),
          bottone('Duplica', () => apriDialogoPiano(p), 'btn piccolo'),
          bottone('Elimina', () => eliminaPiani([p]), 'btn piccolo pericolo')));
    });

    box.replaceChildren(
      barraSelezione || '',
      piani.length ? h('div', { class: 'tabella-scroll' }, h('table', { class: 'tabella' },
        h('thead', null, h('tr', null,
          h('th', { class: 'col-check' }, h('input', {
            type: 'checkbox', 'aria-label': 'Seleziona tutti', checked: tuttiScelti,
            onchange: (e) => { piani.forEach((p) => (e.target.checked ? stato.selezionati.add(p.id) : stato.selezionati.delete(p.id))); disegnaTabella(); }
          })),
          h('th', null, 'Disciplina'), h('th', null, 'Classe'), h('th', null, 'Docente'), h('th', null, 'A.S.'),
          h('th', null, 'Completamento'), h('th', null, 'Ultima modifica'), h('th', null, h('span', { class: 'sr' }, 'Azioni')))),
        h('tbody', null, righe))) : h('p', { class: 'nota' }, 'Nessun piano corrisponde ai filtri.'));
  }

  function benvenuto() {
    return h('div', { class: 'card benvenuto' },
      h('h2', null, 'Benvenuto!'),
      h('p', null, 'Questa app aiuta i docenti della scuola secondaria di primo grado a compilare il piano di lavoro annuale in pochi minuti, con un modello uguale per tutta la scuola.'),
      h('ol', { class: 'passi' },
        h('li', null, h('strong', null, 'Il referente prepara il modello d\'istituto'), ': intestazione della scuola, sezioni del piano, voci degli elenchi, discipline. Poi lo esporta e lo invia ai colleghi.'),
        h('li', null, h('strong', null, 'Ogni docente importa il modello'), ' con il pulsante «Importa file» e crea i propri piani: gran parte delle sezioni si compila con un clic.'),
        h('li', null, h('strong', null, 'Il piano si firma con il dito e si scarica in Word o PDF'), ', sempre con la stessa impaginazione.'),
        h('li', null, h('strong', null, 'Hai già un piano scritto?'), ' Caricalo: l\'app lo riporta nel modello della scuola.')),
      h('div', { class: 'azioni' },
        bottone('+ Crea il primo piano', () => apriDialogoPiano(), 'btn primario'),
        bottone('Importa un piano già scritto (Word, OpenOffice)', importaPianoEsistente),
        bottone('Importa un backup o un modello', importaFile),
        h('a', { class: 'btn', href: '#/modello' }, 'Configura il modello d\'istituto')));
  }

  function eliminaPiani(piani) {
    if (!piani.length) return;
    const msg = piani.length === 1
      ? `Eliminare definitivamente il piano «${M.titoloPiano(piani[0])}»?`
      : `Eliminare definitivamente ${piani.length} piani?`;
    if (!confirm(msg)) return;
    const via = new Set(piani.map((p) => p.id));
    stato.piani = stato.piani.filter((p) => !via.has(p.id));
    via.forEach((id) => stato.selezionati.delete(id));
    salvaPianiOra();
    toast(piani.length === 1 ? 'Piano eliminato.' : `${piani.length} piani eliminati.`);
    vai('#/piani');
  }

  /* ================= Nuovo piano / duplica ================= */

  function apriDialogoPiano(sorgente) {
    const pref = S.preferenze();
    const m = stato.modello;
    const dati = {
      docente: sorgente ? sorgente.docente : (pref.docente || ''),
      disciplina: sorgente ? sorgente.disciplina : '',
      classe: sorgente ? sorgente.classe : '',
      sezione: '',
      anno: m.scuola.annoScolastico,
      ore: sorgente ? sorgente.ore : '',
      libro: sorgente ? sorgente.libro : ''
    };
    let origine = sorgente ? sorgente.id : '';

    const discipline = m.discipline.map((d) => d.nome);
    if (dati.disciplina && !discipline.includes(dati.disciplina)) discipline.push(dati.disciplina);

    const inputOre = h('input', { type: 'number', min: 0, max: 40, value: dati.ore, oninput: (e) => { dati.ore = e.target.value; } });
    const selDisciplina = selezione([{ valore: '', testo: '— scegli —' }].concat(discipline.map((d) => ({ valore: d, testo: d }))),
      dati.disciplina, (v) => {
        dati.disciplina = v;
        const d = M.trovaDisciplina(m, v);
        if (d && d.ore !== '') { dati.ore = String(d.ore); inputOre.value = dati.ore; }
      }, { required: true });

    const altriPiani = stato.piani.slice().sort((a, b) => b.modificato.localeCompare(a.modificato));
    const selOrigine = selezione([{ valore: '', testo: '— nessuno: piano vuoto —' }].concat(altriPiani.map((p) => ({
      valore: p.id, testo: `${M.titoloPiano(p)} · ${p.docente || 'senza docente'} · ${p.anno}`
    }))), origine, (v) => {
      origine = v;
      const p = stato.piani.find((x) => x.id === v);
      if (p && !dati.disciplina) {
        dati.disciplina = p.disciplina;
        selDisciplina.value = p.disciplina;
        if (!dati.ore) { dati.ore = p.ore; inputOre.value = p.ore; }
      }
    });

    const errore = h('p', { class: 'errore-form', role: 'alert' });
    const dlg = h('dialog', { class: 'dialogo', 'aria-labelledby': 'titolo-dialogo' });
    const form = h('form', {
      onsubmit: (e) => {
        e.preventDefault();
        const sezioni = M.elencoSezioniClasse(dati.sezione);
        const mancanti = [];
        if (!dati.docente.trim()) mancanti.push('docente');
        if (!dati.disciplina) mancanti.push('disciplina');
        if (!dati.classe) mancanti.push('classe');
        if (!sezioni.length) mancanti.push('sezione');
        if (mancanti.length) { errore.textContent = 'Compila: ' + mancanti.join(', ') + '.'; return; }
        const src = stato.piani.find((p) => p.id === origine);
        const esistenti = sezioni.filter((sz) => stato.piani.some((p) => p.disciplina === dati.disciplina && p.classe === dati.classe &&
          p.sezione === sz && p.anno === dati.anno.trim() && p.docente.trim().toLowerCase() === dati.docente.trim().toLowerCase()));
        if (esistenti.length && !confirm(`Esiste già un piano di ${dati.disciplina} per ${esistenti.map((s) => dati.classe + s).join(', ')} nello stesso anno. Crearne comunque un altro?`)) return;
        const nuovi = sezioni.map((sz) => {
          const p = M.nuovoPiano(m, Object.assign({}, dati, { sezione: sz }));
          if (src) M.copiaContenuti(p, src, m);
          return p;
        });
        stato.piani.push(...nuovi);
        salvaPianiOra();
        S.salvaPreferenza('docente', dati.docente.trim());
        dlg.close();
        if (nuovi.length === 1) {
          vai('#/piano/' + encodeURIComponent(nuovi[0].id));
          toast(src ? 'Piano creato copiando i contenuti: controlla e completa la situazione di partenza.' : 'Piano creato.');
        } else {
          toast(`Creati ${nuovi.length} piani (${nuovi.map(nomeClasse).join(', ')}).`);
          vai('#/piani');
        }
      }
    },
    h('h2', { id: 'titolo-dialogo' }, sorgente ? 'Duplica piano' : 'Nuovo piano di lavoro'),
    sorgente ? h('p', { class: 'nota' }, `I contenuti di «${M.titoloPiano(sorgente)}» verranno copiati nel nuovo piano. Indica la classe di destinazione.`) : null,
    h('div', { class: 'griglia-2' },
      campo('Docente *', h('input', { type: 'text', value: dati.docente, required: true, autocomplete: 'name', oninput: (e) => { dati.docente = e.target.value; } })),
      campo('Disciplina *', selDisciplina),
      campo('Classe *', selezione([{ valore: '', testo: '— scegli —' }, { valore: '1', testo: 'Prima' }, { valore: '2', testo: 'Seconda' }, { valore: '3', testo: 'Terza' }],
        dati.classe, (v) => { dati.classe = v; }, { required: true })),
      campo('Sezione *', h('input', { type: 'text', value: dati.sezione, required: true, placeholder: 'es. A  oppure  A, B, C', oninput: (e) => { dati.sezione = e.target.value; } })),
      campo('Anno scolastico', h('input', { type: 'text', value: dati.anno, oninput: (e) => { dati.anno = e.target.value; } })),
      campo('Ore settimanali', inputOre)),
    campo('Libro di testo', h('input', { type: 'text', value: dati.libro, oninput: (e) => { dati.libro = e.target.value; } })),
    h('p', { class: 'nota' }, 'Suggerimento: scrivendo più sezioni (es. «A, B, C») si crea un piano per ciascuna classe in un colpo solo.'),
    altriPiani.length ? campo('Copia i contenuti da un piano esistente', selOrigine) : null,
    altriPiani.length ? h('p', { class: 'nota' }, 'Utile per ripartire dal piano dell\'anno scorso o da quello di un collega: vengono copiate tutte le sezioni tranne la situazione di partenza.') : null,
    errore,
    h('div', { class: 'azioni-dialogo' },
      bottone('Annulla', () => dlg.close()),
      h('button', { type: 'submit', class: 'btn primario' }, sorgente ? 'Crea copia' : 'Crea piano')));

    dlg.append(form);
    dlg.addEventListener('close', () => dlg.remove());
    document.body.append(dlg);
    dlg.showModal();
  }

  /* ================= Compilazione del piano ================= */

  function vistaEditor(id) {
    const piano = stato.piani.find((p) => p.id === id);
    if (!piano) {
      app.append(h('div', { class: 'card' }, h('p', null, 'Piano non trovato: potrebbe essere stato eliminato.'), h('a', { href: '#/piani', class: 'btn' }, 'Torna all\'elenco')));
      return;
    }
    const sezioni = M.sezioniAttive(stato.modello);
    let numero = 0;
    const voci = [];
    const schede = [];
    sezioni.forEach((s) => {
      const n = s.tipo === 'dati' ? null : ++numero;
      voci.push(h('li', null, h('a', {
        href: '#sez-' + s.id,
        onclick: (e) => { e.preventDefault(); document.getElementById('sez-' + s.id).scrollIntoView({ behavior: 'smooth', block: 'start' }); }
      }, h('span', { class: 'pallino', 'data-stato-indice': s.id, 'aria-hidden': 'true' }), (n ? n + '. ' : '') + s.titolo)));
      schede.push(schedaSezione(piano, s, n));
    });
    (piano.extra || []).forEach((x, i) => {
      const n = ++numero;
      voci.push(h('li', null, h('a', {
        href: '#sez-extra-' + x.id,
        onclick: (e) => { e.preventDefault(); document.getElementById('sez-extra-' + x.id).scrollIntoView({ behavior: 'smooth', block: 'start' }); }
      }, h('span', { class: 'pallino extra', 'aria-hidden': 'true' }), n + '. ' + (x.titolo || 'Sezione aggiuntiva'))));
      schede.push(schedaExtra(piano, x, n, i === 0));
    });
    voci.push(h('li', null, h('a', {
      href: '#sez-firma',
      onclick: (e) => { e.preventDefault(); document.getElementById('sez-firma').scrollIntoView({ behavior: 'smooth', block: 'start' }); }
    }, h('span', { class: 'pallino', id: 'pallino-firma', 'aria-hidden': 'true' }), 'Firma')));
    schede.push(schedaFirma(piano));

    const barra = h('div', { class: 'barra-editor' },
      h('div', { class: 'barra-titolo' },
        h('a', { href: '#/piani', class: 'indietro', title: 'Torna all\'elenco' }, '← Elenco'),
        h('h1', { id: 'titolo-piano' }, M.titoloPiano(piano)),
        h('span', { id: 'stato-salvataggio', class: 'stato-salvataggio' }, 'Salvato automaticamente')),
      h('div', { class: 'azioni' },
        bottone('Anteprima e stampa', () => { salvaPianiOra(); vai('#/anteprima/' + encodeURIComponent(piano.id)); }, 'btn primario'),
        bottone('Scarica Word', () => scaricaWord([piano])),
        bottone('Duplica', () => apriDialogoPiano(piano)),
        bottone('Esporta file', () => esportaPiani([piano], M.nomeFile(piano, 'json'))),
        bottone('Elimina', () => eliminaPiani([piano]), 'btn pericolo')));

    app.append(h('div', { class: 'editor' },
      h('aside', { class: 'indice' },
        h('div', { id: 'avanzamento-editor' }),
        h('nav', { 'aria-label': 'Sezioni del piano' }, h('ol', null, voci))),
      h('div', { class: 'editor-corpo' }, barra, schede)));
    aggiornaStatoEditor(piano);
  }

  function aggiornaStatoEditor(piano) {
    const av = M.avanzamento(piano, stato.modello);
    const box = document.getElementById('avanzamento-editor');
    if (box) box.replaceChildren(h('p', { class: 'avanzamento-testo' }, `${av.fatte} di ${av.totale} sezioni compilate`), barraAvanzamento(av));
    M.sezioniAttive(stato.modello).forEach((s) => {
      const ok = M.sezioneCompleta(piano, s);
      const sel = CSS.escape(s.id);
      document.querySelectorAll(`[data-stato-indice="${sel}"]`).forEach((el) => el.classList.toggle('completa', ok));
      document.querySelectorAll(`[data-stato="${sel}"]`).forEach((el) => {
        el.classList.toggle('completa', ok);
        el.textContent = ok ? '✓ Compilata' : (s.facoltativa ? 'Facoltativa' : 'Da compilare');
      });
    });
    const t = document.getElementById('titolo-piano');
    if (t) t.textContent = M.titoloPiano(piano);
    const pf = document.getElementById('pallino-firma');
    if (pf) pf.classList.toggle('completa', !!M.firmaValida(piano.firma));
    const af = document.getElementById('avviso-firma');
    if (af) af.hidden = !M.modificatoDopoFirma(piano);
  }

  function schedaExtra(piano, x, numero, prima) {
    return h('section', { class: 'card sezione extra', id: 'sez-extra-' + x.id },
      h('header', { class: 'card-testa' },
        h('h2', null, numero + '. ', h('input', {
          type: 'text', class: 'titolo-extra', value: x.titolo, 'aria-label': 'Titolo della sezione aggiuntiva',
          oninput: (e) => { x.titolo = e.target.value; modificato(piano); }
        })),
        h('span', { class: 'badge-stato' }, 'Sezione aggiuntiva')),
      prima ? h('p', { class: 'guida' }, 'Questa parte del piano importato non corrisponde a nessuna sezione del modello d\'istituto. Puoi lasciarla (comparirà nel documento con l\'impaginazione della scuola), spostarne il testo nella sezione giusta oppure eliminarla.') : null,
      areaTesto(x.testo, (t) => { x.testo = t; modificato(piano); }, { rows: 4, 'aria-label': x.titolo || 'Sezione aggiuntiva' }),
      h('div', { class: 'azioni' }, bottone('Elimina questa sezione', () => {
        if (x.testo.trim() && !confirm(`Eliminare la sezione «${x.titolo}» e il suo testo?`)) return;
        piano.extra = piano.extra.filter((y) => y !== x);
        modificato(piano);
        render({ mantieniScroll: true });
      }, 'btn piccolo pericolo')));
  }

  async function firmaPiano(piano) {
    if (!F) return false;
    const immagine = await F.apriFirma({ titolo: 'Firma il piano di lavoro', nome: piano.docente });
    if (!immagine) return false;
    piano.modificato = new Date().toISOString();
    piano.firma = { immagine, data: dataOggi(), momento: piano.modificato };
    salvaPianiOra();
    toast('Piano firmato.');
    return true;
  }

  function schedaFirma(piano) {
    const f = M.firmaValida(piano.firma);
    const corpo = f
      ? [
        h('div', { class: 'firma-attuale' },
          h('img', { src: f.immagine, alt: 'Firma di ' + (piano.docente || 'docente'), class: 'anteprima-firma' }),
          h('div', null,
            campo('Data della firma', h('input', {
              type: 'date', value: f.data,
              onchange: (e) => {
                if (!/^\d{4}-\d{2}-\d{2}$/.test(e.target.value)) return;
                piano.firma.data = e.target.value;
                piano.modificato = new Date().toISOString();
                piano.firma.momento = piano.modificato;
                salvaPianiOra();
              }
            })),
            h('p', { class: 'avviso-firma', id: 'avviso-firma', hidden: !M.modificatoDopoFirma(piano) },
              '⚠ Il piano è stato modificato dopo la firma: se le modifiche sono importanti, firmalo di nuovo.'))),
        h('div', { class: 'azioni' },
          bottone('Firma di nuovo', async () => { if (await firmaPiano(piano)) render({ mantieniScroll: true }); }),
          bottone('Rimuovi la firma', () => {
            if (!confirm('Rimuovere la firma da questo piano?')) return;
            delete piano.firma;
            modificato(piano);
            render({ mantieniScroll: true });
          }, 'btn pericolo'))
      ]
      : [
        h('p', null, 'Firma con il dito su tablet o smartphone, oppure con la penna o il mouse. La firma comparirà in fondo al documento, nel PDF e nel file Word.'),
        h('div', { class: 'azioni' }, bottone('✍ Firma con il dito', async () => { if (await firmaPiano(piano)) render({ mantieniScroll: true }); }, 'btn primario'))
      ];
    return h('section', { class: 'card sezione', id: 'sez-firma', 'aria-labelledby': 'tit-firma' },
      h('header', { class: 'card-testa' },
        h('h2', { id: 'tit-firma' }, 'Firma del docente'),
        h('span', { class: 'badge-stato' + (f ? ' completa' : '') }, f ? '✓ Firmato' : 'Da firmare')),
      corpo,
      h('p', { class: 'nota' }, 'È una firma grafica (firma elettronica semplice), adatta ai documenti interni. Se la scuola richiede la firma digitale, applicala al PDF con gli strumenti previsti.'));
  }

  function modificato(piano) {
    piano.modificato = new Date().toISOString();
    programmaSalvataggio();
    aggiornaStatoEditor(piano);
  }

  function schedaSezione(piano, s, numero) {
    return h('section', { class: 'card sezione', id: 'sez-' + s.id, 'aria-labelledby': 'tit-' + s.id },
      h('header', { class: 'card-testa' },
        h('h2', { id: 'tit-' + s.id }, (numero ? numero + '. ' : '') + s.titolo),
        h('span', { class: 'badge-stato', 'data-stato': s.id })),
      s.guida ? h('p', { class: 'guida' }, s.guida) : null,
      corpoSezione(piano, s));
  }

  function corpoSezione(piano, s) {
    switch (s.tipo) {
      case 'dati': return editorDati(piano);
      case 'situazione': return editorSituazione(piano, s);
      case 'checklist': return editorChecklist(piano, s);
      case 'traguardi': return editorChecklist(piano, s);
      case 'uda': return editorUda(piano, s);
      case 'testo': {
        const v = M.valoreSezione(piano, s);
        return areaTesto(v.note, (t) => { v.note = t; modificato(piano); }, { rows: 5, 'aria-labelledby': 'tit-' + s.id });
      }
      default: return null;
    }
  }

  function editorDati(piano) {
    const m = stato.modello;
    const testoCampo = (id, attrs = {}) => h('input', Object.assign({
      type: 'text', value: piano[id], oninput: (e) => { piano[id] = e.target.value; modificato(piano); }
    }, attrs));
    const discipline = m.discipline.map((d) => d.nome);
    if (piano.disciplina && !discipline.includes(piano.disciplina)) discipline.push(piano.disciplina);
    return h('div', { class: 'griglia-4' },
      campo('Docente', testoCampo('docente', { autocomplete: 'name' }), 'campo largo-2'),
      campo('Disciplina', selezione([{ valore: '', testo: '— scegli —' }].concat(discipline.map((d) => ({ valore: d, testo: d }))),
        piano.disciplina, (v) => {
          const prima = M.trovaDisciplina(m, piano.disciplina);
          const dopo = M.trovaDisciplina(m, v);
          piano.disciplina = v;
          if (dopo && dopo.ore !== '' && (!piano.ore || (prima && String(prima.ore) === piano.ore))) piano.ore = String(dopo.ore);
          modificato(piano);
          // Traguardi e nuclei tematici dipendono dalla disciplina.
          render({ mantieniScroll: true });
        }), 'campo largo-2'),
      campo('Classe', selezione([{ valore: '', testo: '—' }].concat(Object.keys(NOMI_CLASSI).map((k) => ({ valore: k, testo: NOMI_CLASSI[k] }))),
        piano.classe, (v) => { piano.classe = v; modificato(piano); })),
      campo('Sezione', testoCampo('sezione', { maxlength: 6 })),
      campo('Anno scolastico', testoCampo('anno')),
      campo('Ore settimanali', h('input', { type: 'number', min: 0, max: 40, value: piano.ore, oninput: (e) => { piano.ore = e.target.value; modificato(piano); } })),
      campo('Libro di testo', testoCampo('libro'), 'campo largo-2'),
      campo('Plesso / sede (facoltativo)', testoCampo('plesso'), 'campo largo-2'));
  }

  function numero(valore, oninput, attrs = {}) {
    return h('input', Object.assign({ type: 'number', min: 0, max: 60, inputmode: 'numeric', value: valore == null ? '' : valore, oninput: (e) => oninput(e.target.value) }, attrs));
  }

  function editorSituazione(piano, s) {
    const v = M.valoreSezione(piano, s);
    const avvisi = h('div', { class: 'avvisi-form', 'aria-live': 'polite' });
    const aggiornaAvvisi = () => avvisi.replaceChildren(...M.avvisiSituazione(v, stato.modello.livelli).map((a) => h('p', null, '⚠ ' + a)));
    const cambia = () => { aggiornaAvvisi(); modificato(piano); };
    aggiornaAvvisi();
    return [
      h('h3', null, 'Composizione della classe'),
      h('div', { class: 'griglia-numeri' }, M.CAMPI_SITUAZIONE.map((c) => campo(c.etichetta, numero(v.num[c.id], (x) => { v.num[c.id] = x; cambia(); })))),
      h('h3', null, 'Fasce di livello (numero di alunni)'),
      h('div', { class: 'griglia-numeri' }, stato.modello.livelli.map((l) => campo(l, numero(v.livelli[l], (x) => { v.livelli[l] = x; cambia(); })))),
      avvisi,
      h('h3', null, 'Strumenti di rilevazione'),
      listaVoci(piano, v, s.opzioni),
      h('h3', null, 'Osservazioni sulla classe'),
      areaTesto(v.note, (t) => { v.note = t; modificato(piano); }, {
        rows: 4, 'aria-label': 'Osservazioni sulla classe',
        placeholder: 'Partecipazione, interesse, comportamento, clima relazionale, metodo di studio…'
      })
    ];
  }

  function alterna(elenco, voce, attiva) {
    const i = elenco.indexOf(voce);
    if (attiva && i < 0) elenco.push(voce);
    if (!attiva && i >= 0) elenco.splice(i, 1);
  }

  // Elenco di caselle di spunta + voci aggiuntive scritte a mano.
  function listaVoci(piano, v, opzioni, { lunga = false, vuoto = '' } = {}) {
    const box = h('div', { class: 'checklist' + (lunga ? ' lunga' : '') });
    const disegna = () => {
      const extra = v.sel.filter((x) => !opzioni.includes(x));
      box.replaceChildren();
      if (!opzioni.length && !extra.length && vuoto) box.append(h('p', { class: 'nota' }, vuoto));
      opzioni.concat(extra).forEach((voce) => {
        const eExtra = extra.includes(voce);
        box.append(h('label', { class: 'voce' + (eExtra ? ' extra' : ''), title: eExtra ? 'Voce non più presente nel modello d\'istituto' : null },
          h('input', {
            type: 'checkbox', checked: v.sel.includes(voce),
            onchange: (e) => { alterna(v.sel, voce, e.target.checked); if (eExtra) disegna(); modificato(piano); }
          }),
          h('span', null, voce)));
      });
    };
    disegna();
    const comandi = opzioni.length > 2 ? h('div', { class: 'comandi-lista' },
      bottone('Seleziona tutto', () => { opzioni.forEach((o) => alterna(v.sel, o, true)); disegna(); modificato(piano); }, 'btn-link'),
      bottone('Deseleziona tutto', () => { v.sel.length = 0; disegna(); modificato(piano); }, 'btn-link')) : null;
    return [
      comandi, box,
      campo('Altre voci (una per riga)', areaTesto(v.altro, (t) => { v.altro = t; modificato(piano); }, { rows: 2 }), 'campo altro')
    ];
  }

  function editorChecklist(piano, s) {
    const v = M.valoreSezione(piano, s);
    const traguardi = s.tipo === 'traguardi';
    const opzioni = M.opzioniSezione(s, piano, stato.modello);
    const vuoto = traguardi
      ? (piano.disciplina
        ? 'Il modello d\'istituto non contiene traguardi per questa disciplina: aggiungili qui sotto oppure chiedi al referente di inserirli nel modello.'
        : 'Scegli prima la disciplina nei dati generali: qui compariranno i traguardi corrispondenti.')
      : 'Nessuna voce predefinita: aggiungile qui sotto.';
    return [
      listaVoci(piano, v, opzioni, { lunga: traguardi || opzioni.some((o) => o.length > 70), vuoto }),
      s.campi && s.campi.length ? h('div', { class: 'griglia-campi' }, s.campi.map((c) => campo(c.etichetta,
        c.tipo === 'numero'
          ? numero(v.campi[c.id], (x) => { v.campi[c.id] = x; modificato(piano); })
          : h('input', { type: 'text', value: v.campi[c.id] || '', oninput: (e) => { v.campi[c.id] = e.target.value; modificato(piano); } })))) : null,
      traguardi ? null : campo('Note e specificazioni', areaTesto(v.note, (t) => { v.note = t; modificato(piano); }, { rows: 2 }))
    ];
  }

  function editorUda(piano, s) {
    const v = M.valoreSezione(piano, s);
    const disc = M.trovaDisciplina(stato.modello, piano.disciplina);
    const nucleiDisc = disc ? disc.nuclei : [];
    const lista = h('div', { class: 'uda-lista' });
    const periodi = h('datalist', { id: 'periodi-uda' }, stato.modello.periodi.concat(M.MESI).map((p) => h('option', { value: p })));

    const disegna = () => {
      lista.replaceChildren(...v.unita.map(schedaUnita));
      if (!v.unita.length) lista.append(h('p', { class: 'nota' }, 'Nessuna unità inserita. Aggiungine una o copiale da un altro piano.'));
      adattaTutte(lista);
    };

    const sposta = (i, d) => {
      const [u] = v.unita.splice(i, 1);
      v.unita.splice(i + d, 0, u);
      disegna();
      modificato(piano);
    };

    function schedaUnita(u, i) {
      if (!Array.isArray(u.nuclei)) u.nuclei = [];
      const nuclei = nucleiDisc.concat(u.nuclei.filter((x) => !nucleiDisc.includes(x)));
      const cambia = (k) => (x) => { u[k] = x; modificato(piano); };
      return h('div', { class: 'uda' },
        h('div', { class: 'uda-testa' },
          h('strong', null, 'Unità ' + (i + 1)),
          h('div', { class: 'azioni' },
            bottone('↑', () => sposta(i, -1), 'btn piccolo', { disabled: i === 0, 'aria-label': 'Sposta su', title: 'Sposta su' }),
            bottone('↓', () => sposta(i, 1), 'btn piccolo', { disabled: i === v.unita.length - 1, 'aria-label': 'Sposta giù', title: 'Sposta giù' }),
            bottone('Duplica', () => { v.unita.splice(i + 1, 0, M.clona(u)); disegna(); modificato(piano); }, 'btn piccolo'),
            bottone('Elimina', () => {
              const piena = ['titolo', 'obiettivi', 'contenuti'].some((k) => String(u[k] || '').trim());
              if (piena && !confirm(`Eliminare l'unità ${i + 1}${u.titolo ? ' «' + u.titolo + '»' : ''}?`)) return;
              v.unita.splice(i, 1); disegna(); modificato(piano);
            }, 'btn piccolo pericolo'))),
        h('div', { class: 'griglia-uda' },
          campo('Titolo', h('input', { type: 'text', value: u.titolo, oninput: (e) => cambia('titolo')(e.target.value), 'data-titolo-uda': '' }), 'campo largo'),
          campo('Periodo', h('input', { type: 'text', list: 'periodi-uda', value: u.periodo, oninput: (e) => cambia('periodo')(e.target.value), placeholder: 'es. Settembre-Ottobre' })),
          campo('Ore', numero(u.ore, cambia('ore'), { max: 200 }))),
        nuclei.length ? h('fieldset', { class: 'nuclei' }, h('legend', null, 'Nuclei tematici'),
          nuclei.map((n) => h('label', { class: 'voce piccola' },
            h('input', { type: 'checkbox', checked: u.nuclei.includes(n), onchange: (e) => { alterna(u.nuclei, n, e.target.checked); modificato(piano); } }),
            h('span', null, n)))) : null,
        h('div', { class: 'griglia-2' },
          campo('Obiettivi di apprendimento', areaTesto(u.obiettivi, cambia('obiettivi'), { rows: 4, placeholder: 'Un obiettivo per riga' })),
          campo('Contenuti e attività', areaTesto(u.contenuti, cambia('contenuti'), { rows: 4, placeholder: 'Argomenti, attività, compiti di realtà…' }))));
    }

    const conUnita = stato.piani.filter((p) => p.id !== piano.id && p.valori && p.valori[s.id] && Array.isArray(p.valori[s.id].unita) && p.valori[s.id].unita.length);
    const copiaDa = conUnita.length ? selezione(
      [{ valore: '', testo: 'Copia le unità da un altro piano…' }].concat(conUnita.map((p) => ({
        valore: p.id, testo: `${M.titoloPiano(p)} · ${p.anno} (${p.valori[s.id].unita.length} unità)`
      }))), '', (id) => {
        const p = stato.piani.find((x) => x.id === id);
        if (!p) return;
        const copie = M.clona(p.valori[s.id].unita);
        v.unita.push(...copie);
        disegna();
        modificato(piano);
        copiaDa.value = '';
        toast(`Aggiunte ${copie.length} unità da «${M.titoloPiano(p)}».`);
      }, { 'aria-label': 'Copia le unità da un altro piano' }) : null;

    disegna();
    return [
      periodi,
      lista,
      h('div', { class: 'azioni' },
        bottone('+ Aggiungi unità', () => {
          v.unita.push(M.nuovaUnita());
          disegna();
          modificato(piano);
          const titoli = lista.querySelectorAll('[data-titolo-uda]');
          if (titoli.length) titoli[titoli.length - 1].focus();
        }, 'btn primario'),
        copiaDa)
    ];
  }

  /* ================= Anteprima, Word, esportazione ================= */

  function vistaAnteprima(ids) {
    const piani = ids.map((id) => stato.piani.find((p) => p.id === id)).filter(Boolean);
    if (!piani.length) {
      app.append(h('div', { class: 'card' }, h('p', null, 'Nessun piano da mostrare.'), h('a', { href: '#/piani', class: 'btn' }, 'Torna all\'elenco')));
      return;
    }
    const ritorno = piani.length === 1 ? '#/piano/' + encodeURIComponent(piani[0].id) : '#/piani';
    app.append(
      h('div', { class: 'barra-anteprima' },
        h('a', { href: ritorno, class: 'btn' }, piani.length === 1 ? '← Torna alla compilazione' : '← Torna all\'elenco'),
        h('div', { class: 'azioni' },
          piani.length === 1 ? bottone(M.firmaValida(piani[0].firma) ? '✍ Firma di nuovo' : '✍ Firma con il dito',
            async () => { if (await firmaPiano(piani[0])) render({ mantieniScroll: true }); }) : null,
          bottone('Stampa o salva PDF', () => window.print(), 'btn primario'),
          bottone('Scarica Word', () => scaricaWord(piani))),
        h('p', { class: 'nota' }, 'Per il PDF scegli «Salva come PDF» come stampante nella finestra di stampa.')),
      ...piani.map((p) => h('div', { class: 'foglio', html: D.renderPiano(p, stato.modello) })));
  }

  async function scaricaWord(piani) {
    if (!piani.length) return;
    salvaPianiOra();
    try {
      const dati = await W.creaDocx(piani, stato.modello);
      const nome = piani.length === 1 ? M.nomeFile(piani[0], 'docx') : `Piani di lavoro (${piani.length}).docx`;
      scarica(nome, dati, W.MIME);
      toast('Documento Word scaricato.');
    } catch (e) {
      toast('Non è stato possibile creare il documento Word: ' + e.message, 'errore');
    }
  }

  function esportaPiani(piani, nome) {
    salvaPianiOra();
    scarica(nome, JSON.stringify(M.pacchettoPiani(piani), null, 2), 'application/json');
    toast(piani.length === 1 ? 'Piano esportato.' : `${piani.length} piani esportati.`);
  }

  async function importaFile() {
    const file = await chiediFile('.json,application/json,' + FORMATI_PIANO);
    if (!/\.json$/i.test(file.nome)) { avviaImportazionePiano(file); return; }
    const r = M.leggiPacchetto(new TextDecoder().decode(file.dati));
    if (r.errore) { toast(r.errore, 'errore'); return; }
    if (r.tipo === 'modello') {
      if (!confirm('Il file contiene un modello d\'istituto. Sostituire il modello attuale? I piani già compilati non vengono modificati.')) return;
      stato.modello = r.modello;
      salvaModelloOra();
      toast('Modello d\'istituto importato.');
      render();
    } else {
      const { piani, stat } = M.unisciPiani(stato.piani, r.piani);
      stato.piani = piani;
      salvaPianiOra();
      toast(`Importazione completata: ${stat.aggiunti} nuovi, ${stat.aggiornati} aggiornati` + (stat.ignorati ? `, ${stat.ignorati} già presenti.` : '.'));
      vai('#/piani');
    }
  }

  /* ================= Importazione di un piano già scritto ================= */

  async function importaPianoEsistente() {
    avviaImportazionePiano(await chiediFile(FORMATI_PIANO));
  }

  async function avviaImportazionePiano(file) {
    try {
      const blocchi = await I.leggiFile(file.nome, file.dati);
      const analisi = I.analizza(blocchi, stato.modello);
      if (!analisi.segmenti.length) { toast('Il documento sembra vuoto: non c\'è testo da importare.', 'errore'); return; }
      const pref = S.preferenze();
      stato.importazione = {
        nome: file.nome,
        analisi,
        scelte: analisi.segmenti.map((x) => x.proposta),
        dati: Object.assign({
          docente: pref.docente || '', disciplina: '', classe: '', sezione: '',
          anno: stato.modello.scuola.annoScolastico, ore: '', libro: ''
        }, analisi.dati)
      };
      vai('#/importa');
    } catch (e) {
      toast(e.message || 'Impossibile leggere il documento.', 'errore');
    }
  }

  function vistaImporta() {
    const imp = stato.importazione;
    if (!imp) {
      app.append(h('div', { class: 'card' },
        h('p', null, 'Nessun documento in corso di importazione.'),
        h('div', { class: 'azioni' }, bottone('Scegli un file', importaPianoEsistente, 'btn primario'), h('a', { href: '#/piani', class: 'btn' }, 'Torna all\'elenco'))));
      return;
    }
    const m = stato.modello;
    const d = imp.dati;
    const segmenti = imp.analisi.segmenti;
    const riconosciute = segmenti.filter((x) => x.riconosciuta).length;
    const daImportare = segmenti.filter((x) => !x.intestazione).length;

    app.append(intestazioneVista('Importa un piano già scritto', `Documento: ${imp.nome}`,
      bottone('Scegli un altro file', importaPianoEsistente)));

    // Dati generali
    const discipline = m.discipline.map((x) => x.nome);
    if (d.disciplina && !discipline.includes(d.disciplina)) discipline.push(d.disciplina);
    const testoCampo = (k, attrs = {}) => h('input', Object.assign({ type: 'text', value: d[k] || '', oninput: (e) => { d[k] = e.target.value; } }, attrs));
    app.append(h('section', { class: 'card' },
      h('h2', null, '1. Controlla i dati generali'),
      h('p', { class: 'nota' }, 'Li ho letti dal documento: correggi o completa quello che manca.'),
      h('div', { class: 'griglia-4' },
        campo('Docente *', testoCampo('docente', { autocomplete: 'name' }), 'campo largo-2'),
        campo('Disciplina *', selezione([{ valore: '', testo: '— scegli —' }].concat(discipline.map((x) => ({ valore: x, testo: x }))),
          d.disciplina, (v) => { d.disciplina = v; }), 'campo largo-2'),
        campo('Classe *', selezione([{ valore: '', testo: '—' }, { valore: '1', testo: 'Prima' }, { valore: '2', testo: 'Seconda' }, { valore: '3', testo: 'Terza' }],
          d.classe, (v) => { d.classe = v; })),
        campo('Sezione *', testoCampo('sezione', { maxlength: 6 })),
        campo('Anno scolastico', testoCampo('anno')),
        campo('Ore settimanali', h('input', { type: 'number', min: 0, max: 40, value: d.ore || '', oninput: (e) => { d.ore = e.target.value; } })),
        campo('Libro di testo', testoCampo('libro'), 'campo largo-2'))));

    // Sezioni
    const destinazioni = [];
    let n = 0;
    M.sezioniAttive(m).forEach((sez) => { if (sez.tipo !== 'dati') destinazioni.push({ valore: sez.id, testo: `${++n}. ${sez.titolo}` }); });
    const righe = segmenti.map((seg, i) => {
      const anteprima = seg.testo.length > 260 ? seg.testo.slice(0, 260) + '…' : seg.testo;
      const opzioni = destinazioni.concat([
        { valore: I.EXTRA, testo: `➕ Sezione aggiuntiva «${seg.titolo}»` },
        { valore: I.IGNORA, testo: '✕ Non importare' }
      ]);
      const etichetta = seg.intestazione ? ['intestazione', 'Intestazione'] : seg.riconosciuta ? ['riconosciuta', 'Riconosciuta'] : ['nuova', 'Non prevista dal modello'];
      const riga = h('div', { class: 'seg-import' + (imp.scelte[i] === I.IGNORA ? ' ignorata' : '') },
        h('div', { class: 'seg-testo' },
          h('div', { class: 'seg-titolo' }, h('strong', null, seg.titolo), h('span', { class: 'badge-seg ' + etichetta[0] }, etichetta[1])),
          h('p', { class: 'seg-anteprima' }, anteprima),
          seg.testo.length > 260 ? h('details', null, h('summary', null, 'Mostra tutto il testo'), h('p', { class: 'seg-anteprima' }, seg.testo)) : null),
        campo('Inserisci in', selezione(opzioni, imp.scelte[i], (v) => {
          imp.scelte[i] = v;
          riga.classList.toggle('ignorata', v === I.IGNORA);
        }), 'campo seg-destinazione'));
      return riga;
    });
    app.append(h('section', { class: 'card' },
      h('h2', null, '2. Controlla dove va ogni parte'),
      h('p', { class: 'nota' }, `Ho diviso il documento in ${daImportare} parti e ne ho abbinate ${riconosciute} alle sezioni del modello d'istituto. ` +
        'Le parti non previste dal modello diventano sezioni aggiuntive: nessun testo va perso. Dove possibile le voci vengono spuntate automaticamente (metodologie, strumenti, verifiche…).'),
      righe));

    const errore = h('p', { class: 'errore-form', role: 'alert' });
    app.append(h('section', { class: 'card' },
      h('h2', null, '3. Crea il piano'),
      h('p', null, 'Il piano verrà creato con la struttura e l\'impaginazione del modello d\'istituto. Il documento originale non viene modificato.'),
      errore,
      h('div', { class: 'azioni' },
        bottone('Crea il piano uniformato', () => {
          const mancanti = ['docente', 'disciplina', 'classe', 'sezione'].filter((k) => !String(d[k] || '').trim());
          if (mancanti.length) { errore.textContent = 'Compila nei dati generali: ' + mancanti.join(', ') + '.'; errore.scrollIntoView({ block: 'center' }); return; }
          const piano = I.creaPiano(imp.analisi, imp.scelte, d, m);
          stato.piani.push(piano);
          salvaPianiOra();
          S.salvaPreferenza('docente', d.docente.trim());
          stato.importazione = null;
          vai('#/piano/' + encodeURIComponent(piano.id));
          toast('Piano importato nel modello della scuola. Controlla le sezioni e le voci spuntate automaticamente.');
        }, 'btn primario'),
        bottone('Annulla', () => { stato.importazione = null; vai('#/piani'); }))));
  }

  /* ================= Modello d'istituto ================= */

  function vistaModello() {
    const m = stato.modello;
    app.append(intestazioneVista('Modello d\'istituto',
      'La struttura comune a tutti i piani di lavoro della scuola. Il referente lo configura una volta, lo esporta e lo distribuisce ai colleghi.'));

    if (m.bloccato) {
      app.append(h('div', { class: 'banner-info' },
        h('p', null, h('strong', null, 'Modello distribuito dall\'istituto. '), 'È in sola lettura per garantire che tutti i docenti usino la stessa struttura.'),
        bottone('Modifica comunque', () => {
          if (!confirm('Modificando il modello i tuoi piani potrebbero non essere più uniformi a quelli dei colleghi. Continuare?')) return;
          m.bloccato = false;
          salvaModelloOra();
          render();
        })));
    }

    app.append(
      h('fieldset', { class: 'modello', disabled: m.bloccato },
        h('legend', { class: 'sr' }, 'Impostazioni del modello'),
        schedaScuola(m), schedaSezioniModello(m), schedaDiscipline(m)),
      schedaCondivisione(m));
  }

  function schedaScuola(m) {
    const t = (k, etichetta, attrs = {}) => campo(etichetta, h('input', Object.assign({
      type: 'text', value: m.scuola[k], oninput: (e) => { m.scuola[k] = e.target.value; modelloModificato(); }
    }, attrs)));
    const elenco = (k, etichetta) => campo(etichetta, areaTesto(m[k].join('\n'), (x) => { m[k] = M.righe(x); modelloModificato(); }, { rows: 3 }));
    return h('section', { class: 'card' },
      h('h2', null, 'Intestazione e impostazioni generali'),
      h('div', { class: 'griglia-2' },
        t('nome', 'Nome della scuola', { placeholder: 'es. Istituto Comprensivo «G. Rodari»' }),
        t('sottotitolo', 'Sottotitolo', { placeholder: 'es. Scuola secondaria di primo grado' }),
        t('citta', 'Città (per luogo e data)'),
        t('annoScolastico', 'Anno scolastico proposto per i nuovi piani')),
      h('div', { class: 'griglia-2' },
        elenco('periodi', 'Periodi didattici (uno per riga)'),
        elenco('livelli', 'Fasce di livello (una per riga)')));
  }

  function schedaSezioniModello(m) {
    const lista = h('div', { class: 'lista-sezioni' });

    const disegna = () => {
      lista.replaceChildren(...m.sezioni.map(rigaSezione));
      adattaTutte(lista);
    };
    const cambia = () => modelloModificato();
    const sposta = (i, d) => {
      const [s] = m.sezioni.splice(i, 1);
      m.sezioni.splice(i + d, 0, s);
      cambia();
      disegna();
    };

    function rigaSezione(s, i) {
      const fissa = s.tipo === 'dati';
      const riga = h('div', { class: 'riga-sezione' + (s.attiva ? '' : ' disattiva') });
      const dettagli = [];
      dettagli.push(campo('Indicazioni per i docenti (compaiono sotto il titolo)', areaTesto(s.guida, (x) => { s.guida = x; cambia(); }, { rows: 2 })));
      if (s.tipo === 'checklist' || s.tipo === 'situazione') {
        dettagli.push(campo(s.tipo === 'situazione' ? 'Strumenti di rilevazione proposti (uno per riga)' : 'Voci proposte (una per riga)',
          areaTesto(s.opzioni.join('\n'), (x) => { s.opzioni = M.righe(x); cambia(); }, { rows: 6 })));
      }
      if (s.tipo === 'checklist' || s.tipo === 'situazione' || s.tipo === 'testo') {
        dettagli.push(campo('Testo già inserito nei nuovi piani (modificabile dal docente)', areaTesto(s.testo, (x) => { s.testo = x; cambia(); }, { rows: 2 })));
      }
      if (s.tipo === 'traguardi' || s.tipo === 'uda') {
        dettagli.push(h('p', { class: 'nota' }, s.tipo === 'traguardi'
          ? 'I traguardi dipendono dalla disciplina: si modificano più sotto, nella sezione «Discipline».'
          : 'I nuclei tematici proposti dipendono dalla disciplina: si modificano più sotto, nella sezione «Discipline».'));
      }
      dettagli.push(h('label', { class: 'voce piccola' },
        h('input', { type: 'checkbox', checked: s.facoltativa, onchange: (e) => { s.facoltativa = e.target.checked; cambia(); } }),
        h('span', null, 'Facoltativa (non conta nel completamento del piano)')));

      aggiungi(riga, [
        h('div', { class: 'riga-testa' },
          h('input', {
            type: 'checkbox', checked: s.attiva, disabled: fissa, 'aria-label': 'Includi la sezione nel piano',
            title: fissa ? 'Sezione sempre presente' : 'Includi nel piano',
            onchange: (e) => { s.attiva = e.target.checked; riga.classList.toggle('disattiva', !s.attiva); cambia(); }
          }),
          h('input', { type: 'text', class: 'titolo-sezione', value: s.titolo, 'aria-label': 'Titolo della sezione', oninput: (e) => { s.titolo = e.target.value; cambia(); } }),
          h('span', { class: 'tipo' }, ETICHETTE_TIPO[s.tipo]),
          h('div', { class: 'azioni' },
            bottone('↑', () => sposta(i, -1), 'btn piccolo', { disabled: i <= 1, 'aria-label': 'Sposta su' }),
            bottone('↓', () => sposta(i, 1), 'btn piccolo', { disabled: fissa || i === m.sezioni.length - 1, 'aria-label': 'Sposta giù' }),
            s.personalizzata ? bottone('Elimina', () => {
              if (!confirm(`Eliminare la sezione «${s.titolo}» dal modello?`)) return;
              m.sezioni.splice(i, 1); cambia(); disegna();
            }, 'btn piccolo pericolo') : null)),
        fissa ? null : h('details', null, h('summary', null, 'Personalizza'), h('div', { class: 'dettagli' }, dettagli))
      ]);
      return riga;
    }

    const nuova = (tipo) => {
      m.sezioni.push({
        id: 'extra-' + M.uid(), tipo, titolo: tipo === 'testo' ? 'Nuova sezione di testo' : 'Nuova sezione',
        attiva: true, facoltativa: false, personalizzata: true, guida: '', opzioni: [], campi: [], testo: ''
      });
      cambia();
      disegna();
      const titoli = lista.querySelectorAll('.titolo-sezione');
      titoli[titoli.length - 1].select();
    };

    disegna();
    return h('section', { class: 'card' },
      h('h2', null, 'Sezioni del piano'),
      h('p', { class: 'nota' }, 'Scegli quali sezioni includere, in che ordine e con quali voci. Le modifiche valgono anche per i piani già creati.'),
      lista,
      h('div', { class: 'azioni' },
        bottone('+ Sezione con elenco di voci', () => nuova('checklist')),
        bottone('+ Sezione di testo libero', () => nuova('testo'))));
  }

  function schedaDiscipline(m) {
    const lista = h('div', { class: 'lista-discipline' });

    const rinomina = (d, input) => {
      const nuovo = input.value.trim();
      const vecchio = d.nome;
      if (nuovo === vecchio) return;
      if (!nuovo || m.discipline.some((x) => x !== d && x.nome.toLowerCase() === nuovo.toLowerCase())) {
        toast(nuovo ? 'Esiste già una disciplina con questo nome.' : 'Il nome non può essere vuoto.', 'errore');
        input.value = vecchio;
        return;
      }
      d.nome = nuovo;
      const coinvolti = stato.piani.filter((p) => p.disciplina === vecchio);
      coinvolti.forEach((p) => { p.disciplina = nuovo; });
      if (coinvolti.length) { salvaPianiOra(); toast(`Aggiornati ${coinvolti.length} piani con il nuovo nome.`); }
      modelloModificato();
      disegna();
    };

    function schedaDisciplina(d, i) {
      const usata = stato.piani.filter((p) => p.disciplina === d.nome).length;
      return h('details', { class: 'disciplina' },
        h('summary', null, h('strong', null, d.nome), h('span', { class: 'nota' }, ` ${d.traguardi.length} traguardi · ${d.nuclei.length} nuclei tematici` + (usata ? ` · usata in ${usata} piani` : ''))),
        h('div', { class: 'dettagli' },
          h('div', { class: 'griglia-2' },
            campo('Nome', h('input', { type: 'text', value: d.nome, onchange: (e) => rinomina(d, e.target) })),
            campo('Ore settimanali proposte', h('input', { type: 'number', min: 0, max: 40, value: d.ore, oninput: (e) => { d.ore = e.target.value === '' ? '' : Number(e.target.value); modelloModificato(); } }))),
          campo('Nuclei tematici (uno per riga)', areaTesto(d.nuclei.join('\n'), (x) => { d.nuclei = M.righe(x); modelloModificato(); }, { rows: 4 })),
          campo('Traguardi per lo sviluppo delle competenze (uno per riga)', areaTesto(d.traguardi.join('\n'), (x) => { d.traguardi = M.righe(x); modelloModificato(); }, { rows: 8 })),
          bottone('Elimina disciplina', () => {
            if (!confirm(`Eliminare «${d.nome}» dal modello?` + (usata ? ` I ${usata} piani che la usano restano, ma senza traguardi proposti.` : ''))) return;
            m.discipline.splice(i, 1); modelloModificato(); disegna();
          }, 'btn piccolo pericolo')));
    }

    const disegna = () => lista.replaceChildren(...m.discipline.map(schedaDisciplina));
    disegna();
    lista.addEventListener('toggle', (e) => adattaTutte(e.target), true);

    return h('section', { class: 'card' },
      h('h2', null, 'Discipline'),
      h('p', { class: 'nota' }, 'Per ogni disciplina: traguardi (dalle Indicazioni nazionali 2012) e nuclei tematici proposti ai docenti. Rinominando una disciplina si aggiornano anche i piani salvati in questo browser.'),
      lista,
      h('div', { class: 'azioni' }, bottone('+ Aggiungi disciplina', () => {
        let n = 1;
        while (m.discipline.some((d) => d.nome === 'Nuova disciplina ' + n)) n++;
        m.discipline.push({ nome: 'Nuova disciplina ' + n, ore: '', nuclei: [], traguardi: [] });
        modelloModificato();
        disegna();
        const ultimo = lista.lastElementChild;
        ultimo.open = true;
        ultimo.querySelector('input').select();
      })));
  }

  function schedaCondivisione(m) {
    const blocca = h('input', { type: 'checkbox', checked: true });
    return h('section', { class: 'card' },
      h('h2', null, 'Distribuzione ai docenti'),
      h('ol', { class: 'passi' },
        h('li', null, 'Esporta il modello e invia il file ai colleghi (e-mail, registro elettronico, cartella condivisa).'),
        h('li', null, 'Ogni docente apre l\'app e usa «Importa file»: da quel momento tutti compilano la stessa struttura.'),
        h('li', null, 'Se l\'app è pubblicata sul sito della scuola, salva il file come ', h('code', null, 'modello-istituto.json'), ' accanto a ', h('code', null, 'index.html'), ': i docenti lo riceveranno automaticamente, anche quando lo aggiorni.')),
      h('label', { class: 'voce piccola' }, blocca, h('span', null, 'Distribuisci in sola lettura (i docenti non lo modificano per errore)')),
      h('div', { class: 'azioni' },
        bottone('Esporta il modello', () => {
          const copia = M.clona(m);
          copia.bloccato = blocca.checked;
          copia.aggiornato = m.aggiornato || new Date().toISOString();
          scarica('modello-istituto.json', JSON.stringify(M.pacchettoModello(copia), null, 2), 'application/json');
          toast('Modello esportato.');
        }, 'btn primario'),
        bottone('Importa un modello', importaFile),
        bottone('Ripristina il modello predefinito', () => {
          if (!confirm('Ripristinare sezioni, voci e discipline predefinite? I dati della scuola vengono mantenuti.')) return;
          const nuovo = M.normalizzaModello(null);
          nuovo.scuola = m.scuola;
          nuovo.aggiornato = new Date().toISOString();
          stato.modello = nuovo;
          salvaModelloOra();
          render();
          toast('Modello predefinito ripristinato.');
        }, 'btn pericolo')));
  }

  /* ================= Guida ================= */

  function vistaGuida() {
    const blocco = (titolo, ...contenuto) => h('section', { class: 'card guida-testo' }, h('h2', null, titolo), contenuto);
    app.append(
      intestazioneVista('Guida rapida', 'Come usare l\'app in pochi minuti.'),
      blocco('Per il referente o il dirigente',
        h('ol', null,
          h('li', null, 'Apri «Modello d\'istituto» e inserisci il nome della scuola e la città.'),
          h('li', null, 'Scegli le sezioni da includere, il loro ordine e le voci proposte (metodologie, strumenti, verifiche…).'),
          h('li', null, 'Controlla le discipline: traguardi e nuclei tematici sono già inseriti e si possono adattare al curricolo d\'istituto.'),
          h('li', null, 'Premi «Esporta il modello» e distribuisci il file ai docenti.'))),
      blocco('Per i docenti',
        h('ol', null,
          h('li', null, 'Premi «Importa file» e scegli il modello ricevuto dalla scuola (una sola volta).'),
          h('li', null, 'Premi «+ Nuovo piano», indica disciplina, classe e sezione. Scrivendo «A, B, C» crei tre piani insieme.'),
          h('li', null, 'Compila le sezioni: quasi tutto si fa spuntando le voci. L\'indice a sinistra mostra cosa manca.'),
          h('li', null, 'Premi «Anteprima e stampa» per il PDF, oppure «Scarica Word» per il file modificabile.'))),
      blocco('Hai già un piano scritto?',
        h('ol', null,
          h('li', null, 'Premi «Importa un piano già scritto» e scegli il file (.docx di Word, .odt di LibreOffice o .txt).'),
          h('li', null, 'L\'app riconosce docente, disciplina e classe e divide il documento nelle sue parti, abbinandole alle sezioni del modello d\'istituto.'),
          h('li', null, 'Controlla gli abbinamenti e premi «Crea il piano uniformato»: il piano avrà la struttura e l\'aspetto comuni della scuola.')),
        h('p', { class: 'nota' }, 'Le voci di metodologie, strumenti, verifiche, competenze e traguardi vengono spuntate quando il testo le cita in modo riconoscibile; il resto viene conservato come «altre voci» o note. ' +
          'Le parti non previste dal modello diventano sezioni aggiuntive. I file .doc (Word 97-2003) vanno prima salvati come .docx; i PDF non sono supportati.')),
      blocco('Firma',
        h('p', null, 'In fondo alla compilazione, o dall\'anteprima, premi «Firma con il dito»: su tablet e smartphone si firma con il dito, al computer con il mouse o la penna. La firma compare nel PDF e nel file Word, con luogo e data.'),
        h('p', { class: 'nota' }, 'Si tratta di una firma grafica (firma elettronica semplice): va bene per i documenti interni. Se il piano viene modificato dopo la firma, l\'app lo segnala.')),
      blocco('Per fare prima',
        h('ul', null,
          h('li', null, h('strong', null, 'Riparti dall\'anno scorso: '), 'crea il nuovo piano scegliendo «Copia i contenuti da un piano esistente». Viene azzerata solo la situazione di partenza.'),
          h('li', null, h('strong', null, 'Stessa disciplina in più classi: '), 'compila un piano, poi «Duplica» e indica le altre sezioni.'),
          h('li', null, h('strong', null, 'Unità di apprendimento: '), 'usa «Copia le unità da un altro piano» per riprendere la programmazione di un collega o di un altro anno.'),
          h('li', null, h('strong', null, 'Consiglio di classe: '), 'importa i piani dei colleghi, filtra per classe, selezionali e stampali o scaricali in un unico documento.'))),
      blocco('Dove sono salvati i dati',
        h('p', null, 'I piani sono salvati automaticamente ', h('strong', null, 'solo in questo browser, su questo computer'), '. Nessun dato viene inviato a server esterni.'),
        h('ul', null,
          h('li', null, 'Fai periodicamente «Backup di tutti i piani» dall\'elenco e conserva il file.'),
          h('li', null, 'Per passare a un altro computer esporta i piani e importali lì con «Importa file».'),
          h('li', null, 'Cancellare i dati di navigazione del browser elimina anche i piani non esportati.'))),
      blocco('Riferimenti',
        h('ul', null,
          h('li', null, 'Indicazioni nazionali per il curricolo della scuola dell\'infanzia e del primo ciclo d\'istruzione (D.M. 254/2012) e Indicazioni nazionali e nuovi scenari (2018).'),
          h('li', null, 'Raccomandazione del Consiglio dell\'Unione europea del 22 maggio 2018 sulle competenze chiave per l\'apprendimento permanente.'),
          h('li', null, 'D.Lgs. 62/2017 sulla valutazione; L. 104/1992, L. 170/2010 e normativa sui BES per l\'inclusione.'),
          h('li', null, 'L. 92/2019 e Linee guida per l\'insegnamento dell\'educazione civica (D.M. 183/2024).')),
        h('p', { class: 'nota' }, 'I testi dei traguardi sono riportati dalle Indicazioni nazionali, in alcuni casi in forma sintetica: il referente può adattarli nel modello d\'istituto.')));
  }

  /* ================= Avvio ================= */

  function mostraAggiornamentoModello(remoto) {
    const banner = document.getElementById('banner');
    banner.replaceChildren(h('div', { class: 'banner-info' },
      h('p', null, h('strong', null, 'È disponibile un modello d\'istituto aggiornato. '), 'Applicalo per allinearti ai colleghi: i piani già compilati non vengono modificati.'),
      h('div', { class: 'azioni' },
        bottone('Aggiorna', () => {
          stato.modello = remoto;
          salvaModelloOra();
          banner.replaceChildren();
          render({ mantieniScroll: true });
          toast('Modello d\'istituto aggiornato.');
        }, 'btn primario'),
        bottone('Più tardi', () => banner.replaceChildren()))));
  }

  function avvia() {
    const salvato = S.caricaModello();
    stato.modello = salvato || M.normalizzaModello(null);
    stato.modelloSalvato = !!salvato;
    stato.piani = S.caricaPiani();

    document.getElementById('btn-importa').addEventListener('click', importaFile);
    const stile = document.createElement('style');
    stile.textContent = D.CSS_DOCUMENTO;
    document.head.append(stile);

    window.addEventListener('hashchange', () => render());
    aggiornaTestata();
    render();

    S.caricaModelloIstituto().then((remoto) => {
      if (!remoto) return;
      if (!stato.modelloSalvato) {
        stato.modello = remoto;
        salvaModelloOra();
        render({ mantieniScroll: true });
      } else if (remoto.aggiornato && remoto.aggiornato > (stato.modello.aggiornato || '')) {
        mostraAggiornamentoModello(remoto);
      }
    });
  }

  avvia();
})();
