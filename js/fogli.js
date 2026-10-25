/*
 * Dati comuni delle classi da un foglio condiviso (Google Fogli, Excel, CSV).
 *
 * Il foglio ha due schede:
 *   - "Classi": Classe, Sezione, Plesso, Alunni, Maschi, Femmine,
 *               Alunni con PEI, Alunni con PDP (DSA), Alunni con PDP (altri BES)
 *   - "Libri":  Classe, Sezione (vuota = tutte), Disciplina, Libro di testo
 * I nomi delle colonne vengono riconosciuti anche se scritti in modo diverso.
 */
(function (root) {
  'use strict';

  const PDL = root.PDL || {};
  const req = (m) => (typeof require === 'function' ? require(m) : null);
  const zip = PDL.zip || req('./zip.js');
  const I = PDL.importa || req('./importa.js');
  const M = PDL.model || req('./model.js');

  /* ---------- CSV ---------- */

  function leggiCsv(testo) {
    const t = String(testo).replace(/^﻿/, '');
    const prima = t.split(/\r?\n/)[0] || '';
    // Excel in italiano salva i CSV con il punto e virgola.
    const sep = (prima.match(/;/g) || []).length > (prima.match(/,/g) || []).length ? ';' : (prima.includes('\t') ? '\t' : ',');
    const righe = [];
    let riga = [];
    let cella = '';
    let virgolette = false;
    for (let i = 0; i < t.length; i++) {
      const c = t[i];
      if (virgolette) {
        if (c === '"' && t[i + 1] === '"') { cella += '"'; i++; } else if (c === '"') virgolette = false; else cella += c;
      } else if (c === '"') virgolette = true;
      else if (c === sep) { riga.push(cella); cella = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && t[i + 1] === '\n') i++;
        riga.push(cella); righe.push(riga); riga = []; cella = '';
      } else cella += c;
    }
    if (cella || riga.length) { riga.push(cella); righe.push(riga); }
    return righe.filter((r) => r.some((x) => String(x).trim()));
  }

  /* ---------- XLSX ---------- */

  function colonna(rif) {
    let n = 0;
    for (const ch of String(rif).replace(/\d+/g, '')) n = n * 26 + (ch.toUpperCase().charCodeAt(0) - 64);
    return n - 1;
  }

  /**
   * Legge tutte le schede di un file .xlsx.
   * @returns {Promise<Object<string, string[][]>>} nome scheda → righe
   */
  async function leggiXlsx(dati) {
    const z = zip.apri(dati);
    const workbook = await z.testo('xl/workbook.xml');
    if (!workbook) throw new Error('Il file non è un foglio Excel valido.');
    const relazioni = {};
    I.scorriXml((await z.testo('xl/_rels/workbook.xml.rels')) || '', {
      apri(n, a) { if (n === 'Relationship') relazioni[a.Id] = a.Target; }, chiudi() {}, testo() {}
    });
    const schede = [];
    I.scorriXml(workbook, {
      apri(n, a) { if (n === 'sheet') schede.push({ nome: a.name, rid: a['r:id'] }); }, chiudi() {}, testo() {}
    });
    const condivise = [];
    const sst = await z.testo('xl/sharedStrings.xml');
    if (sst) {
      let corrente = null;
      let inT = false;
      let inRph = false;
      I.scorriXml(sst, {
        apri(n, a, vuoto) {
          if (n === 'si') corrente = '';
          else if (n === 'rPh') inRph = !vuoto;
          else if (n === 't' && !vuoto) inT = true;
        },
        chiudi(n) {
          if (n === 'si') { condivise.push(corrente || ''); corrente = null; } else if (n === 't') inT = false; else if (n === 'rPh') inRph = false;
        },
        testo(s) { if (inT && !inRph && corrente != null) corrente += s; }
      });
    }
    const out = {};
    for (const sc of schede) {
      let percorso = relazioni[sc.rid] || '';
      percorso = percorso.startsWith('/') ? percorso.slice(1) : 'xl/' + percorso.replace(/^\.\//, '');
      const xml = await z.testo(percorso);
      if (!xml) continue;
      const righe = [];
      let riga = null;
      let cella = null;
      let campo = null;
      I.scorriXml(xml, {
        apri(n, a, vuoto) {
          if (n === 'row') { riga = []; if (vuoto) { righe.push(riga); riga = null; } } else if (n === 'c' && riga) {
            cella = { col: a.r ? colonna(a.r) : riga.length, tipo: a.t || '', valore: '' };
            if (vuoto) { riga[cella.col] = ''; cella = null; }
          } else if ((n === 'v' || n === 't') && cella && !vuoto) campo = n;
        },
        chiudi(n) {
          if (n === 'v' || n === 't') campo = null;
          else if (n === 'c' && cella && riga) {
            riga[cella.col] = cella.tipo === 's' ? (condivise[Number(cella.valore)] || '') : cella.valore;
            cella = null;
          } else if (n === 'row' && riga) {
            for (let i = 0; i < riga.length; i++) if (riga[i] == null) riga[i] = '';
            righe.push(riga);
            riga = null;
          }
        },
        testo(s) { if (cella && campo) cella.valore += s; }
      });
      out[sc.nome] = righe.filter((r) => r.some((x) => String(x).trim()));
    }
    return out;
  }

  /* ---------- Riconoscimento delle colonne ---------- */

  const COLONNE = [
    ['classe', /^(classe|classi|anno di corso|anno)$|^classe /],
    ['sezione', /^sez/],
    ['plesso', /plesso|sede|scuola$/],
    ['disciplina', /disciplin|materia|insegnamento/],
    ['libro', /libr|testo|titolo|adozion/],
    ['h104', /\bpei\b|104|disabil/],
    ['dsa', /\bdsa\b/],
    ['bes', /\bbes\b/],
    ['maschi', /maschi/],
    ['femmine', /femmine/],
    ['alunni', /alunni|studenti|iscritti|totale|^n$|^numero$/]
  ];

  function campoDi(intestazione) {
    const n = I.norm(intestazione);
    if (!n) return null;
    for (const [campo, re] of COLONNE) if (re.test(n)) return campo;
    return null;
  }

  const ROMANI = { i: '1', ii: '2', iii: '3', prima: '1', seconda: '2', terza: '3', '1': '1', '2': '2', '3': '3' };

  // "2", "II", "seconda", "2A", "2^A", "2 A" → { classe: '2', sezione: 'A' }
  function leggiClasse(valore) {
    const t = I.norm(String(valore).replace(/[ª°^]/g, ' ')).replace(/\s+/g, ' ');
    const m = /^(i{1,3}|[1-3]|prima|seconda|terza)(?: ?([a-z]))?$/.exec(t);
    return m ? { classe: ROMANI[m[1]], sezione: (m[2] || '').toUpperCase() } : null;
  }

  /**
   * Trasforma le schede lette in dati comuni.
   * @param {Object<string, string[][]>} schede
   * @returns {{classi: Array, libri: Array, avvisi: string[]}}
   */
  function interpreta(schede, modello) {
    const classi = [];
    const libri = [];
    const avvisi = [];
    Object.keys(schede).forEach((nome) => {
      const righe = schede[nome];
      // L'intestazione è la prima riga in cui si riconosce la colonna "Classe".
      const iTesta = righe.findIndex((r) => r.some((c) => campoDi(c) === 'classe'));
      if (iTesta < 0) return;
      const campi = righe[iTesta].map(campoDi);
      righe[iTesta].forEach((c, i) => {
        if (!campi[i] && /\bpdp\b/.test(I.norm(c))) {
          avvisi.push(`Scheda «${nome}»: la colonna «${c}» non distingue DSA e altri BES. Usa due colonne: «Alunni con PDP (DSA)» e «Alunni con PDP (altri BES)».`);
        }
      });
      const ha = (k) => campi.includes(k);
      const eLibri = ha('disciplina') && ha('libro');
      const eClassi = ['plesso', 'alunni', 'maschi', 'femmine', 'h104', 'dsa', 'bes'].some((k) => ha(k) && !(eLibri && k === 'plesso'));
      righe.slice(iTesta + 1).forEach((r, k) => {
        const val = (campo) => { const i = campi.indexOf(campo); return i >= 0 ? String(r[i] == null ? '' : r[i]).trim() : ''; };
        if (!val('classe')) return;
        const c = leggiClasse(val('classe'));
        if (!c) { avvisi.push(`Scheda «${nome}», riga ${iTesta + k + 2}: classe «${val('classe')}» non riconosciuta (usa 1, 2 o 3).`); return; }
        const sezione = (val('sezione') || c.sezione).toUpperCase();
        if (eClassi && sezione) {
          const riga = { classe: c.classe, sezione, plesso: val('plesso') };
          M.CAMPI_COMUNI_NUMERICI.forEach((n) => {
            const v = val(n).replace(/[,.]0+$/, '');
            if (v && !/^\d{1,3}$/.test(v)) avvisi.push(`Scheda «${nome}», classe ${c.classe}${sezione}: «${v}» non è un numero.`);
            riga[n] = /^\d{1,3}$/.test(v) ? v : '';
          });
          classi.push(riga);
        }
        if (eLibri && val('libro') && val('disciplina')) {
          const disciplina = I.riconosciDisciplina(val('disciplina'), modello) || val('disciplina');
          libri.push({ classe: c.classe, sezione, plesso: val('plesso'), disciplina, libro: val('libro') });
        }
      });
    });
    return { classi, libri, avvisi };
  }

  /* ---------- Google Fogli ---------- */

  /**
   * Indirizzi da cui scaricare un foglio Google a partire dal suo link.
   * Il foglio deve essere condiviso come "Chiunque abbia il link: visualizzatore"
   * oppure pubblicato sul web (File → Condividi → Pubblica sul web).
   */
  function indirizziGoogle(link) {
    const s = String(link || '').trim();
    const pubblicato = /^(https:\/\/docs\.google\.com\/spreadsheets\/d\/e\/[\w-]+)/.exec(s);
    if (pubblicato) return { tipo: 'xlsx', url: pubblicato[1] + '/pub?output=xlsx' };
    const id = /\/spreadsheets\/d\/([\w-]{20,})/.exec(s);
    if (id) {
      const base = 'https://docs.google.com/spreadsheets/d/' + id[1] + '/gviz/tq?tqx=out:csv&headers=1&sheet=';
      return { tipo: 'schede', url: { Classi: base + 'Classi', Libri: base + 'Libri' } };
    }
    return null;
  }

  async function scaricaTesto(url, f) {
    let r;
    try { r = await f(url, { cache: 'no-store', credentials: 'omit' }); } catch (e) {
      throw new Error('Il foglio non è raggiungibile: controlla la connessione e che il foglio sia condiviso con «Chiunque abbia il link».');
    }
    if (!r.ok) throw new Error('Il foglio non è accessibile (errore ' + r.status + '): controlla che sia condiviso con «Chiunque abbia il link».');
    return r;
  }

  /**
   * Scarica e interpreta il foglio comune da un link di Google Fogli.
   * @param {Function} [f] funzione fetch (sostituibile nei test)
   */
  async function daLink(link, modello, f = root.fetch && root.fetch.bind(root)) {
    const ind = indirizziGoogle(link);
    if (!ind) throw new Error('Il link non sembra un foglio Google: copialo dalla barra degli indirizzi del foglio.');
    if (ind.tipo === 'xlsx') {
      const r = await scaricaTesto(ind.url, f);
      return interpreta(await leggiXlsx(new Uint8Array(await r.arrayBuffer())), modello);
    }
    const schede = {};
    for (const nome of Object.keys(ind.url)) {
      const r = await scaricaTesto(ind.url[nome], f);
      const testo = await r.text();
      // Se il foglio non è condiviso Google risponde con la pagina di accesso.
      if (/^\s*</.test(testo)) throw new Error('Il foglio non è condiviso: in Google Fogli usa Condividi → «Chiunque abbia il link» → Visualizzatore.');
      schede[nome] = leggiCsv(testo);
    }
    // Se manca la scheda "Libri" Google restituisce la prima scheda: si scarta il doppione.
    if (JSON.stringify(schede.Libri) === JSON.stringify(schede.Classi)) delete schede.Libri;
    return interpreta(schede, modello);
  }

  /**
   * Interpreta un file caricato dal computer (.xlsx o .csv).
   */
  async function daFile(nome, dati, modello) {
    if (/\.xlsx$/i.test(nome)) return interpreta(await leggiXlsx(dati), modello);
    if (/\.(csv|tsv|txt)$/i.test(nome)) return interpreta({ [nome]: leggiCsv(new TextDecoder().decode(dati)) }, modello);
    if (/\.(xls|ods)$/i.test(nome)) throw new Error('Salva il foglio in formato .xlsx (o scaricalo da Google Fogli come Microsoft Excel) e riprova.');
    throw new Error('Formato non supportato: usa un file .xlsx o .csv.');
  }

  /* ---------- Modello del foglio da caricare su Drive ---------- */

  function cellaXml(valore, rif, stile) {
    const s = stile ? ` s="${stile}"` : '';
    if (typeof valore === 'number') return `<c r="${rif}"${s}><v>${valore}</v></c>`;
    const t = String(valore == null ? '' : valore).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return `<c r="${rif}" t="inlineStr"${s}><is><t xml:space="preserve">${t}</t></is></c>`;
  }

  function schedaXml(righe, larghezze) {
    const lettera = (i) => String.fromCharCode(65 + i);
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
      '<cols>' + larghezze.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>' +
      '<sheetData>' + righe.map((r, i) => `<row r="${i + 1}">` +
        r.map((v, j) => cellaXml(v, lettera(j) + (i + 1), i === 0 ? 1 : 0)).join('') + '</row>').join('') +
      '</sheetData></worksheet>';
  }

  /**
   * Crea il modello del foglio comune (.xlsx) già predisposto per la scuola.
   * @param {string[]} sezioni sezioni presenti (es. ['A','B','C'])
   */
  async function creaModelloFoglio(modello, sezioni = ['A', 'B', 'C']) {
    const classi = [['Classe', 'Sezione', 'Plesso', 'Alunni', 'Maschi', 'Femmine', 'Alunni con PEI', 'Alunni con PDP (DSA)', 'Alunni con PDP (altri BES)']];
    ['1', '2', '3'].forEach((c) => sezioni.forEach((s) => classi.push([Number(c), s, '', '', '', '', '', '', ''])));
    const libri = [['Classe', 'Sezione', 'Disciplina', 'Libro di testo']];
    ['1', '2', '3'].forEach((c) => modello.discipline.forEach((d) => libri.push([Number(c), '', d.nome, ''])));
    const istruzioni = [
      ['Foglio comune dei piani di lavoro'],
      ['Scheda «Classi»: una riga per classe. Plesso, numero di alunni e alunni con PEI/PDP valgono per tutti i docenti della classe.'],
      ['Scheda «Libri»: una riga per classe e disciplina. Lascia vuota la sezione se il libro è lo stesso per tutte le sezioni; aggiungi una riga con la sezione solo dove il libro è diverso.'],
      ['Scrivi solo numeri: mai nomi di alunni. Le celle vuote non vengono usate e restano compilabili dal docente.'],
      ['Condivisione: Condividi → «Chiunque abbia il link» → Visualizzatore, poi incolla il link nell\'app (Modello d\'istituto → Dati comuni delle classi).'],
      ['Non rinominare le schede «Classi» e «Libri».']
    ];
    return creaXlsx([['Classi', classi, [9, 9, 22, 9, 9, 9, 15, 20, 24]], ['Libri', libri, [9, 9, 28, 50]], ['Istruzioni', istruzioni, [120]]]);
  }

  /**
   * Crea un file .xlsx.
   * @param {Array<[string, Array<Array<string|number>>, number[]]>} fogli  [nome, righe, larghezze colonne]
   */
  function creaXlsx(fogli) {
    const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
    const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    return zip.crea([
      { nome: '[Content_Types].xml', dati: XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        fogli.map((f, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
        '</Types>' },
      { nome: '_rels/.rels', dati: XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>' },
      { nome: 'xl/workbook.xml', dati: XML + `<workbook ${NS}><sheets>` +
        fogli.map((f, i) => `<sheet name="${f[0]}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') + '</sheets></workbook>' },
      { nome: 'xl/_rels/workbook.xml.rels', dati: XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        fogli.map((f, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
        `<Relationship Id="rId${fogli.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
      { nome: 'xl/styles.xml', dati: XML + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFE8F0FA"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>' +
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>' },
      ...fogli.map((f, i) => ({ nome: `xl/worksheets/sheet${i + 1}.xml`, dati: schedaXml(f[1], f[2]) }))
    ]);
  }

  const api = { leggiCsv, leggiXlsx, interpreta, leggiClasse, indirizziGoogle, daLink, daFile, creaModelloFoglio, creaXlsx };
  root.PDL = root.PDL || {};
  root.PDL.fogli = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
