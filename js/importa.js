/*
 * Importazione di un piano di lavoro già scritto (.docx, .odt, .txt).
 *
 * 1. leggiFile(): estrae dal documento paragrafi (con titoli, grassetti,
 *    elenchi) e tabelle.
 * 2. analizza(): riconosce i dati del docente, divide il testo in sezioni
 *    e propone per ognuna la sezione corrispondente del modello d'istituto.
 * 3. creaPiano(): dopo la revisione del docente, trasforma ogni sezione nel
 *    formato del modello (voci spuntate, traguardi, numeri, tabella UdA...).
 *    Quello che non corrisponde a nulla diventa una "sezione aggiuntiva":
 *    nessun contenuto va perso e l'aspetto è comunque quello della scuola.
 */
(function (root) {
  'use strict';

  const PDL = root.PDL || {};
  const M = PDL.model || (typeof require === 'function' ? require('./model.js') : null);
  const zip = PDL.zip || (typeof require === 'function' ? require('./zip.js') : null);

  const EXTRA = 'extra';
  const IGNORA = 'ignora';

  /* ================= Testo ================= */

  function norm(s) {
    return String(s == null ? '' : s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
      .replace(/[^a-z0-9]+/g, ' ').trim();
  }

  const STOP = new Set(('di del della dello dei degli delle e ed il lo la i gli le un uno una per con in a ad da dal ' +
    'dalla dai dalle su sul sulla tra fra al allo alla ai agli alle che o od anche nel nello nella nei negli nelle come ' +
    'piu non si sono essere loro suo sua suoi sue questo questa quello quella ogni sia dall dell nell sull all coll').split(' '));

  // Radici delle parole significative (le prime 5 lettere bastano per plurali e varianti).
  function radici(s) {
    return [...new Set(norm(s).split(' ').filter((t) => t && !STOP.has(t) && (t.length >= 3 || /\d/.test(t))).map((t) => t.slice(0, 5)))];
  }

  function pulisciVoce(s) {
    return String(s).replace(/^[\s•●▪◦·*\-–—>]+/, '').replace(/^\(?([a-z]|\d{1,2})[.)]\s+/i, '').trim();
  }

  function frase(s) {
    // "METODOLOGIE DIDATTICHE" → "Metodologie didattiche"
    const t = String(s).trim();
    const lettere = t.replace(/[^A-Za-zÀ-ÿ]/g, '');
    if (lettere.length > 3 && lettere === lettere.toUpperCase()) return t.charAt(0) + t.slice(1).toLowerCase();
    return t;
  }

  function pulisciTitolo(s) {
    return frase(String(s).replace(/^\s*((\d{1,2}|[IVX]{1,4}|[A-Za-z])[.)]|\d{1,2}(\.\d{1,2})*)\s+/, '').replace(/[\s:.\-–]+$/, ''));
  }

  /* ================= Lettura XML ================= */

  const ENTITA = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
  function decodifica(s) {
    return s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => {
      if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
      return ENTITA[e] != null ? ENTITA[e] : m;
    });
  }

  function attributi(s) {
    const out = {};
    s.replace(/([\w:.-]+)\s*=\s*"([^"]*)"/g, (m, k, v) => { out[k] = decodifica(v); return m; });
    return out;
  }

  // Scorre l'XML chiamando apri(nome, attr, vuoto), chiudi(nome), testo(stringa).
  // Per i tag auto-chiudenti (vuoto = true) chiudi() non viene chiamata.
  function scorriXml(xml, { apri, chiudi, testo }) {
    const re = /<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!\[CDATA\[([\s\S]*?)\]\]>|<(\/?)([\w:.-]+)([^>]*?)(\/?)>|([^<]+)/g;
    let m;
    while ((m = re.exec(xml))) {
      if (m[1] != null) testo(m[1]);
      else if (m[3]) {
        if (m[2]) chiudi(m[3]);
        else apri(m[3], attributi(m[4]), !!m[5]);
      } else if (m[6] != null) testo(decodifica(m[6]));
    }
  }

  function nuovoParagrafo() {
    return { t: 'p', testo: '', forte: false, lista: false, caratteri: 0, grassetti: 0, prefissoGrassetto: '', finePrefisso: false, stile: '' };
  }

  // Chiude un paragrafo calcolando se è tutto in grassetto.
  function chiudiParagrafo(p) {
    p.testo = p.testo.replace(/[ \t ]+/g, ' ').replace(/ *\n */g, '\n').trim();
    p.grassetto = p.caratteri > 0 && p.grassetti / p.caratteri >= 0.9;
    return p;
  }

  /* ---------- DOCX ---------- */

  function leggiStiliDocx(xml) {
    const stili = {};
    if (!xml) return stili;
    let s = null;
    let inRpr = false;
    scorriXml(xml, {
      apri(n, a, vuoto) {
        if (n === 'w:style') s = vuoto ? null : (stili[a['w:styleId']] = { nome: '', base: '', grassetto: null, livello: null });
        else if (!s) return;
        else if (n === 'w:name') s.nome = a['w:val'] || '';
        else if (n === 'w:basedOn') s.base = a['w:val'] || '';
        else if (n === 'w:outlineLvl') s.livello = Number(a['w:val']);
        else if (n === 'w:rPr') inRpr = !vuoto;
        else if (n === 'w:b' && inRpr) s.grassetto = !/^(0|false|off)$/i.test(a['w:val'] || '');
      },
      chiudi(n) { if (n === 'w:style') s = null; else if (n === 'w:rPr') inRpr = false; },
      testo() {}
    });
    // Eredità dagli stili di base.
    const risolvi = (id, campo, profondita = 0) => {
      const st = stili[id];
      if (!st || profondita > 10) return null;
      return st[campo] != null ? st[campo] : risolvi(st.base, campo, profondita + 1);
    };
    for (const id of Object.keys(stili)) {
      const st = stili[id];
      st.grassettoEffettivo = !!risolvi(id, 'grassetto');
      const livello = risolvi(id, 'livello');
      const nome = norm(st.nome || id);
      st.titolo = (livello != null && livello <= 5) || /^(heading|titolo|title|intestazione)( \d)?$/.test(nome) || /^(heading|titolo)\d$/.test(norm(id));
    }
    return stili;
  }

  function leggiDocx(documentXml, stiliXml) {
    const stili = leggiStiliDocx(stiliXml);
    const blocchi = [];
    const pila = [];          // elementi aperti
    const paragrafi = [];     // paragrafi aperti (le caselle di testo ne annidano)
    const tabelle = [];       // tabelle aperte (possono essere annidate)
    let run = null;
    let inTesto = false;

    const dentro = (nome) => pila.includes(nome);
    const emetti = (p) => {
      chiudiParagrafo(p);
      if (!p.testo) return;
      const t = tabelle[tabelle.length - 1];
      if (t && t.cella) t.cella.push(p);
      else blocchi.push(p);
    };

    scorriXml(documentXml, {
      apri(n, a, vuoto) {
        const p = paragrafi[paragrafi.length - 1];
        switch (n) {
          case 'w:p': if (!vuoto) paragrafi.push(nuovoParagrafo()); break;
          case 'w:pStyle':
            if (p && dentro('w:pPr')) {
              const st = stili[a['w:val']];
              p.stile = a['w:val'] || '';
              if (st && st.titolo) p.forte = true;
              p.grassettoStile = !!(st && st.grassettoEffettivo);
            }
            break;
          case 'w:outlineLvl': if (p && dentro('w:pPr') && Number(a['w:val']) <= 5) p.forte = true; break;
          case 'w:numPr': if (p) p.lista = true; break;
          case 'w:r': run = vuoto ? null : { testo: '', grassetto: p ? !!p.grassettoStile : false }; break;
          case 'w:b': if (run && dentro('w:rPr')) run.grassetto = !/^(0|false|off)$/i.test(a['w:val'] || ''); break;
          case 'w:t': inTesto = !vuoto; break;
          case 'w:tab': if (run) run.testo += '\t'; break;
          case 'w:br': case 'w:cr': if (run && a['w:type'] !== 'page') run.testo += '\n'; break;
          case 'w:tbl': if (!vuoto) tabelle.push({ righe: [], riga: null, cella: null }); break;
          case 'w:tr': if (tabelle.length && !vuoto) tabelle[tabelle.length - 1].riga = []; break;
          case 'w:tc': if (tabelle.length && !vuoto) tabelle[tabelle.length - 1].cella = []; break;
          default: break;
        }
        if (!vuoto) pila.push(n);
      },
      chiudi(n) {
        const i = pila.lastIndexOf(n);
        if (i >= 0) pila.splice(i);
        const p = paragrafi[paragrafi.length - 1];
        switch (n) {
          case 'w:t': inTesto = false; break;
          case 'w:r':
            if (run && p) {
              const len = run.testo.trim().length;
              p.caratteri += len;
              if (run.grassetto) p.grassetti += len;
              if (!p.finePrefisso) {
                if (run.grassetto || !run.testo.trim()) p.prefissoGrassetto += run.testo;
                else p.finePrefisso = true;
              }
              p.testo += run.testo;
            }
            run = null;
            break;
          case 'w:p': if (paragrafi.length) emetti(paragrafi.pop()); break;
          case 'w:tc': {
            const t = tabelle[tabelle.length - 1];
            if (t && t.riga) t.riga.push(t.cella || []);
            if (t) t.cella = null;
            break;
          }
          case 'w:tr': {
            const t = tabelle[tabelle.length - 1];
            if (t && t.riga) t.righe.push(t.riga);
            if (t) t.riga = null;
            break;
          }
          case 'w:tbl': {
            const t = tabelle.pop();
            if (!t) break;
            const esterna = tabelle[tabelle.length - 1];
            if (esterna && esterna.cella) {
              // Tabella annidata: il suo testo confluisce nella cella esterna.
              t.righe.forEach((r) => r.forEach((c) => c.forEach((p) => esterna.cella.push(p))));
            } else if (t.righe.length) {
              blocchi.push({ t: 'tab', righe: t.righe });
            }
            break;
          }
          default: break;
        }
      },
      testo(s) { if (inTesto && run) run.testo += s; }
    });
    return blocchi;
  }

  /* ---------- ODT ---------- */

  function stiliGrassettoOdt(...xml) {
    const grassetti = new Set();
    const titoli = new Set();
    xml.filter(Boolean).forEach((x) => {
      let nome = null;
      scorriXml(x, {
        apri(n, a) {
          if (n === 'style:style') {
            nome = a['style:name'];
            const genitore = norm(a['style:parent-style-name'] || '');
            if (/^(heading|titolo)/.test(genitore) || /^(heading|titolo)/.test(norm(nome))) titoli.add(nome);
            if (grassetti.has(a['style:parent-style-name'])) grassetti.add(nome);
          } else if (n === 'style:text-properties' && nome && /bold|[6-9]00/.test(a['fo:font-weight'] || '')) {
            grassetti.add(nome);
          }
        },
        chiudi(n) { if (n === 'style:style') nome = null; },
        testo() {}
      });
    });
    return { grassetti, titoli };
  }

  function leggiOdt(contentXml, stylesXml) {
    const { grassetti, titoli } = stiliGrassettoOdt(stylesXml, contentXml);
    const blocchi = [];
    const paragrafi = [];
    const tabelle = [];
    const spanGrassetti = [];
    let inStili = false;
    let liste = 0;

    const aggiungiTesto = (s) => {
      const p = paragrafi[paragrafi.length - 1];
      if (!p || inStili) return;
      const g = spanGrassetti.length ? spanGrassetti[spanGrassetti.length - 1] : p.grassettoStile;
      const len = s.trim().length;
      p.caratteri += len;
      if (g) p.grassetti += len;
      if (!p.finePrefisso) { if (g || !s.trim()) p.prefissoGrassetto += s; else p.finePrefisso = true; }
      p.testo += s;
    };

    scorriXml(contentXml, {
      apri(n, a, vuoto) {
        switch (n) {
          case 'office:automatic-styles': if (!vuoto) inStili = true; break;
          case 'text:p': case 'text:h': {
            const p = nuovoParagrafo();
            p.forte = n === 'text:h' || titoli.has(a['text:style-name']);
            p.lista = liste > 0;
            p.grassettoStile = grassetti.has(a['text:style-name']);
            if (!vuoto) paragrafi.push(p);
            break;
          }
          case 'text:span': if (!vuoto) spanGrassetti.push(grassetti.has(a['text:style-name']) || (spanGrassetti[spanGrassetti.length - 1] || false)); break;
          case 'text:list': if (!vuoto) liste++; break;
          case 'text:s': aggiungiTesto(' '.repeat(Number(a['text:c']) || 1)); break;
          case 'text:tab': aggiungiTesto('\t'); break;
          case 'text:line-break': aggiungiTesto('\n'); break;
          case 'table:table': if (!vuoto) tabelle.push({ righe: [], riga: null, cella: null }); break;
          case 'table:table-row': if (tabelle.length && !vuoto) tabelle[tabelle.length - 1].riga = []; break;
          case 'table:table-cell': {
            const t = tabelle[tabelle.length - 1];
            if (t && vuoto) { if (t.riga) t.riga.push([]); } else if (t) t.cella = [];
            break;
          }
          default: break;
        }
      },
      chiudi(n) {
        switch (n) {
          case 'office:automatic-styles': inStili = false; break;
          case 'text:p': case 'text:h': if (paragrafi.length) emetti(paragrafi.pop()); break;
          case 'text:span': spanGrassetti.pop(); break;
          case 'text:list': liste = Math.max(0, liste - 1); break;
          case 'table:table-cell': chiudiCella(); break;
          case 'table:table-row': {
            const t = tabelle[tabelle.length - 1];
            if (t && t.riga) { while (t.riga.length && !t.riga[t.riga.length - 1].length) t.riga.pop(); if (t.riga.length) t.righe.push(t.riga); }
            if (t) t.riga = null;
            break;
          }
          case 'table:table': {
            const t = tabelle.pop();
            if (!t) break;
            const esterna = tabelle[tabelle.length - 1];
            if (esterna && esterna.cella) t.righe.forEach((r) => r.forEach((c) => c.forEach((p) => esterna.cella.push(p))));
            else if (t.righe.length) blocchi.push({ t: 'tab', righe: t.righe });
            break;
          }
          default: break;
        }
      },
      testo: aggiungiTesto
    });

    function chiudiCella() {
      const t = tabelle[tabelle.length - 1];
      if (t && t.riga && t.cella) t.riga.push(t.cella);
      if (t) t.cella = null;
    }
    function emetti(p) {
      chiudiParagrafo(p);
      if (!p.testo) return;
      const t = tabelle[tabelle.length - 1];
      if (t && t.cella) t.cella.push(p);
      else blocchi.push(p);
    }
    return blocchi;
  }

  /* ---------- Testo semplice ---------- */

  function leggiTesto(s) {
    return String(s).replace(/\r\n?/g, '\n').split('\n').map((riga) => {
      const p = nuovoParagrafo();
      p.lista = /^\s*([•●▪◦*\-–]|\d{1,2}[.)])\s+/.test(riga) && !/^\s*\d{1,2}[.)]\s+[A-ZÀ-Ý\s]{4,}$/.test(riga);
      p.testo = riga.trim();
      return chiudiParagrafo(p);
    }).filter((p) => p.testo);
  }

  /**
   * Legge un file e ne estrae i blocchi di testo.
   * @param {string} nome  nome del file (serve l'estensione)
   * @param {Uint8Array} dati
   */
  async function leggiFile(nome, dati) {
    const est = (String(nome).toLowerCase().match(/\.([a-z0-9]+)$/) || [])[1] || '';
    if (est === 'docx' || est === 'docm' || est === 'dotx') {
      const z = zip.apri(dati);
      const doc = await z.testo('word/document.xml');
      if (!doc) throw new Error('Il file non sembra un documento Word valido.');
      return leggiDocx(doc, await z.testo('word/styles.xml'));
    }
    if (est === 'odt' || est === 'ott') {
      const z = zip.apri(dati);
      const contenuto = await z.testo('content.xml');
      if (!contenuto) throw new Error('Il file non sembra un documento OpenDocument valido.');
      return leggiOdt(contenuto, await z.testo('styles.xml'));
    }
    if (est === 'txt' || est === 'md') return leggiTesto(new TextDecoder().decode(dati));
    if (est === 'doc') throw new Error('I file .doc (Word 97-2003) non si possono leggere: apri il file con Word o LibreOffice e salvalo come .docx, poi riprova.');
    if (est === 'pdf') throw new Error('I PDF non si possono leggere in modo affidabile: se hai il file originale in Word usa quello (.docx), altrimenti copia il testo in un file .txt.');
    throw new Error('Formato non supportato: usa un file .docx, .odt o .txt.');
  }

  /* ================= Riconoscimento delle sezioni ================= */

  // [parola chiave, peso]. "$" finale = parola intera.
  const PAROLE_CHIAVE = {
    situazione: [['situazione di partenza', 3], ['situazione iniziale', 3], ['situazione della classe', 3], ['analisi della situazione', 3],
      ['analisi della classe', 3], ['presentazione della classe', 3], ['profilo della classe', 3], ['composizione della classe', 3],
      ['livelli di partenza', 3], ['fasce di livello', 3], ['test d ingresso', 2], ['prove d ingresso', 2], ['situazione', 1], ['classe$', 1]],
    competenze: [['competenze chiave', 3], ['competenze europee', 3], ['competenze di cittadinanza', 3], ['competenze trasversali', 2], ['competenze$', 1]],
    traguardi: [['traguardi', 3]],
    uda: [['unita di apprendimento', 3], ['uda$', 3], ['obiettivi di apprendimento', 3], ['obiettivi specifici', 3], ['contenuti', 2],
      ['programmazione', 2], ['articolazione', 1], ['percorsi', 1], ['moduli', 1], ['nuclei tematici', 2], ['argomenti', 2], ['obiettivi', 1],
      ['conoscenze', 1], ['abilita', 1], ['scansione', 1]],
    metodologie: [['metodolog', 3], ['metodi$', 2], ['strategie', 2], ['approcci', 1], ['didattica$', 1]],
    strumenti: [['strumenti', 2], ['mezzi$', 2], ['sussidi', 2], ['materiali', 1], ['risorse', 1], ['spazi', 1]],
    verifiche: [['verifiche', 2], ['verifica$', 2], ['prove$', 1], ['tipologia delle prove', 3]],
    valutazione: [['criteri di valutazione', 3], ['valutazione', 2], ['griglie', 1]],
    recupero: [['recupero', 3], ['potenziamento', 3], ['consolidamento', 3], ['sostegno allo studio', 2]],
    inclusione: [['inclusione', 3], ['inclusiv', 3], ['bes$', 3], ['dsa$', 3], ['disabilit', 3], ['pdp$', 2], ['pei$', 2],
      ['bisogni educativi', 3], ['personalizzazione', 2], ['individualizzazione', 2]],
    edcivica: [['educazione civica', 4], ['cittadinanza e costituzione', 4], ['civica', 3]],
    famiglie: [['famiglie', 3], ['famiglia', 3], ['rapporti con', 2], ['colloqui', 2]],
    note: [['osservazioni', 2], ['note$', 2], ['considerazioni', 1]]
  };

  function contienePhrase(n, kw) {
    const intero = kw.endsWith('$');
    const k = norm(intero ? kw.slice(0, -1) : kw);
    const re = new RegExp('(^| )' + k.replace(/ /g, ' ') + (intero ? '( |$)' : ''));
    return re.test(n);
  }

  /**
   * Propone la sezione del modello per un titolo.
   * @returns {{id: string, punteggio: number} | null}
   */
  function mappaTitolo(titolo, modello) {
    const n = norm(titolo);
    if (!n) return null;
    let migliore = null;
    modello.sezioni.forEach((s, ordine) => {
      if (!s.attiva || s.tipo === 'dati') return;
      let punteggio = 0;
      (PAROLE_CHIAVE[s.id] || []).forEach(([kw, peso]) => { if (contienePhrase(n, kw)) punteggio += peso; });
      // Sezioni rinominate o personalizzate: confronto con le parole del titolo.
      const r = radici(s.titolo);
      if (r.length) {
        const t = new Set(radici(titolo));
        const comuni = r.filter((x) => t.has(x)).length;
        if (comuni === r.length) punteggio += 3;
        else if (comuni / r.length >= 0.6 && comuni >= 2) punteggio += 2;
      }
      if (punteggio > 0 && (!migliore || punteggio > migliore.punteggio)) migliore = { id: s.id, punteggio, ordine };
    });
    return migliore ? { id: migliore.id, punteggio: migliore.punteggio } : null;
  }

  // Riconosce i titoli scritti senza stile: 'numerato', 'maiuscolo', 'grassetto' o null.
  function tipoTitolo(p) {
    const t = p.testo;
    if (t.length > 90 || t.length < 3 || /[,;]$/.test(t) || /\n/.test(t)) return null;
    if (!/[a-zà-ÿ]{3}/i.test(t)) return null;
    const lettere = t.replace(/[^A-Za-zÀ-ÿ]/g, '');
    if (lettere.length >= 4 && lettere === lettere.toUpperCase()) return 'maiuscolo';
    if (/^\s*(\d{1,2}|[IVX]{1,4})[.)]\s+[A-ZÀ-Ý]/.test(t) && t.length <= 70 && !/[.]$/.test(t) && !p.lista) return 'numerato';
    if (p.grassetto && !/[.]$/.test(t)) return 'grassetto';
    return null;
  }

  /* ---------- Dati del docente ---------- */

  const DISCIPLINE_SINONIMI = [
    [/scienze motorie|educazione fisica|motoria/, 'Educazione fisica'],
    [/religione|\birc\b/, 'Religione cattolica'],
    [/arte e immagine|\barte\b|educazione artistica/, 'Arte e immagine'],
    [/seconda lingua|francese|spagnolo|tedesco/, 'Seconda lingua comunitaria'],
    [/inglese/, 'Lingua inglese'],
    [/tecnologia/, 'Tecnologia'],
    [/musica|musicale/, 'Musica'],
    [/matematica/, 'Matematica'],
    [/\bscienze\b/, 'Scienze'],
    [/geografia/, 'Geografia'],
    [/storia/, 'Storia'],
    [/italiano/, 'Italiano']
  ];

  function riconosciDisciplina(testo, modello) {
    const n = norm(testo);
    if (!n) return '';
    // Vince la corrispondenza più lunga: "scienze motorie" è Educazione fisica, non Scienze;
    // "strumento musicale" (se la scuola l'ha aggiunta) non è Musica.
    const candidati = [];
    modello.discipline.forEach((d) => { if (contienePhrase(n, d.nome + '$')) candidati.push({ nome: d.nome, lunghezza: norm(d.nome).length }); });
    DISCIPLINE_SINONIMI.forEach(([re, nome]) => {
      const m = re.exec(n);
      if (m && M.trovaDisciplina(modello, nome)) candidati.push({ nome, lunghezza: m[0].length });
    });
    candidati.sort((a, b) => b.lunghezza - a.lunghezza);
    return candidati.length ? candidati[0].nome : '';
  }

  const ROMANI = { i: '1', ii: '2', iii: '3', prima: '1', seconda: '2', terza: '3' };

  function riconosciDati(righe, modello) {
    const dati = {};
    const pezzi = [];
    righe.forEach((r) => {
      pezzi.push(r);
      r.split(/\t+| {3,}| \| /).forEach((x) => { if (x.trim() && x.trim() !== r.trim()) pezzi.push(x.trim()); });
    });
    for (const s of pezzi) {
      let m;
      if (!dati.docente && (m = /^(?:il\s+|la\s+)?(?:docente|insegnante|prof\.\s*ssa\.?|prof\.?ssa\b|prof\.|prof\b|professor(?:e|essa)\b)\s*[:\-–]?\s*(.+)$/i.exec(s))) {
        const v = m[1].replace(/^(prof\.?\s*(ssa)?\.?|professor(e|essa))\s+/i, '').split(/\s+[-–|]\s+|\t|\s{2,}|\s+(?:disciplina|materia|classe)\b/i)[0].trim();
        if (/[a-z]{2}/i.test(v) && v.length <= 60 && !/\d/.test(v) && !/^(di|del|coordinat|sostegno)/i.test(v)) dati.docente = v;
      }
      if (!dati.disciplina && (m = /^(?:disciplina|materia|insegnamento)\s*[:\-–]?\s*(.+)$/i.exec(s))) {
        dati.disciplina = riconosciDisciplina(m[1], modello) || frase(m[1].split(/\s+[-–|]\s+|\t|\s{2,}|\s+classe\b/i)[0]).slice(0, 60);
      }
      if (!dati.classe && (m = /\bclass[ei]\s*[:\-–]?\s*(i{1,3}|[1-3]|prima|seconda|terza)\s*(?:[ªa°^]|\s)?\s*(?:sez(?:ione)?\.?\s*)?([a-z])\b/i.exec(s))) {
        dati.classe = ROMANI[m[1].toLowerCase()] || m[1];
        dati.sezione = m[2].toUpperCase();
      }
      if (!dati.anno && (m = /(?:\ba\.\s*s\.|\banno\s+scolastico)\s*[:\-–]?\s*(20\d\d)\s*[\/\-–]\s*(?:20)?(\d\d)\b/i.exec(s))) {
        const a = Number(m[1]);
        if (Number(m[2]) === (a + 1) % 100) dati.anno = a + '/' + (a + 1);
      }
      if (!dati.ore && (m = /(\d{1,2})\s*ore\s*settimanal|ore\s*settimanal\w*\s*[:\-–]?\s*(\d{1,2})/i.exec(s))) dati.ore = m[1] || m[2];
      if (!dati.libro && (m = /^libr[oi]\s+di\s+testo(?:\s+in\s+adozione|\s+adottat[oi])?\s*[:\-–]\s*([^;].{2,})$/i.exec(s))) dati.libro = m[1].trim();
    }
    if (!dati.disciplina) {
      for (const r of righe.slice(0, 15)) {
        const d = riconosciDisciplina(r, modello);
        if (d) { dati.disciplina = d; break; }
      }
    }
    if (!dati.anno) {
      for (const r of righe.slice(0, 20)) {
        const m = /\b(20\d\d)\s*[\/\-–]\s*(20\d\d)\b/.exec(r);
        if (m && Number(m[2]) === Number(m[1]) + 1) { dati.anno = m[1] + '/' + m[2]; break; }
      }
    }
    return dati;
  }

  /* ---------- Segmentazione ---------- */

  function testoCella(c) { return c.map((p) => p.testo).join('\n').trim(); }

  // Le tabelle "etichetta | contenuto" (molto usate nei modelli scolastici)
  // diventano titoli + paragrafi; le tabelle a una colonna diventano paragrafi.
  function appiattisci(blocchi, modello) {
    const out = [];
    blocchi.forEach((b) => {
      if (b.t !== 'tab') { out.push(b); return; }
      const colonne = Math.max(...b.righe.map((r) => r.length));
      const testi = b.righe.map((r) => r.map(testoCella));
      // Tabelle di numeri con le etichette in testa (es. composizione della classe): restano tabelle.
      const numerica = testi.length >= 2 && testi.slice(1).every((r) => r.every((c) => /^\s*\d{0,3}\s*$/.test(c)));
      if (numerica) { out.push({ t: 'tab', righe: testi }); return; }
      if (colonne === 1) { b.righe.forEach((r) => r[0].forEach((p) => out.push(p))); return; }
      if (colonne === 2) {
        const sicura = (t) => { const m = mappaTitolo(t, modello); return m && m.punteggio >= 2; };
        const etichette = b.righe.filter((r) => r.length === 2 && testoCella(r[0]).length <= 80 && sicura(testoCella(r[0])));
        if (etichette.length && etichette.length >= Math.ceil(b.righe.length / 3)) {
          b.righe.forEach((r) => {
            const sx = testoCella(r[0] || []);
            const dx = r[1] || [];
            if (!sx) { dx.forEach((p) => out.push(p)); return; }
            if (!sicura(sx) && testoCella(dx).length < 80 && !/\n/.test(testoCella(dx))) {
              out.push(Object.assign(nuovoParagrafo(), { testo: sx.replace(/[:\s]+$/, '') + ': ' + testoCella(dx) }));
              return;
            }
            out.push(Object.assign(nuovoParagrafo(), { testo: sx, forte: true }));
            dx.forEach((p) => out.push(p));
          });
          return;
        }
      }
      out.push({ t: 'tab', righe: b.righe.map((r) => r.map(testoCella)) });
    });
    return out;
  }

  function lineeDi(blocchi) {
    const out = [];
    blocchi.forEach((b) => {
      if (b.t === 'tab') b.righe.forEach((r) => { const t = r.filter(Boolean).join(' – '); if (t) out.push(t); });
      else b.testo.split('\n').forEach((t) => { if (t.trim()) out.push(t.trim()); });
    });
    return out;
  }

  const TITOLO_INTESTAZIONE = /piano (di lavoro|annuale|didattico|individuale)|programmazione (annuale|didattica|disciplinare|individuale)|istituto|scuola secondaria|anno scolastico|^a s /;
  const RIGA_DATI = /^(docente|insegnante|prof|disciplina|materia|class[ei]|sezione|anno|a\.\s*s\.|libr[oi]|ore|plesso|sede|scuola|istituto|coordinat|dirigente)/i;

  /**
   * Analizza i blocchi estratti dal file.
   * @returns {{dati: object, segmenti: Array}}
   */
  function analizza(blocchiGrezzi, modello) {
    const blocchi = appiattisci(blocchiGrezzi, modello);
    const dati = riconosciDati(lineeDi(blocchi).slice(0, 80), modello);
    const segmenti = [];
    let corrente = { titolo: 'Intestazione del documento', blocchi: [], iniziale: true };
    segmenti.push(corrente);

    const apri = (titolo, mappa, tipo) => {
      corrente = { titolo: pulisciTitolo(titolo) || 'Sezione', blocchi: [], mappa, tipo };
      segmenti.push(corrente);
    };
    const disc = M.trovaDisciplina(modello, dati.disciplina);
    // Nella programmazione i sottotitoli sono di solito unità, nuclei o periodi.
    const eUnita = (b, tipo) => tipo !== corrente.tipo || INIZIO_UDA.test(b.testo) || MESI_RE.test(b.testo) ||
      (disc && corrispondenze(b.testo, disc.nuclei, { traguardi: true }).length > 0);
    const sottotitolo = (b) => corrente.blocchi.push(Object.assign({}, b, { sottotitolo: true }));

    blocchi.forEach((b) => {
      if (b.t !== 'p') { corrente.blocchi.push(b); return; }
      const tipo = b.forte ? 'forte' : tipoTitolo(b);
      if (tipo) {
        const mappa = mappaTitolo(b.testo, modello);
        const inUda = corrente.mappa && corrente.mappa.id === 'uda';
        // Regole: i titoli con stile o riconosciuti aprono sempre una sezione;
        // nella programmazione (UdA) gli altri titoli sono le singole unità;
        // altrove i titoli numerati o in maiuscolo aprono una sezione aggiuntiva.
        if (tipo === 'forte' || mappa) apri(b.testo, mappa, tipo);
        else if ((inUda && eUnita(b, tipo)) || tipo === 'grassetto') sottotitolo(b);
        else apri(b.testo, null, tipo);
        return;
      }
      // "Metodologie: lezione frontale, lavori di gruppo..." in un unico paragrafo.
      const due = /^([^:\n]{3,60}):\s*(\S[\s\S]*)$/.exec(b.testo);
      if (due) {
        const etichetta = due[1].trim();
        const lettere = etichetta.replace(/[^A-Za-zÀ-ÿ]/g, '');
        const evidenziata = b.prefissoGrassetto.replace(/[:\s]+$/, '').trim().length >= etichetta.length * 0.8 ||
          (lettere.length >= 4 && lettere === lettere.toUpperCase());
        const mappa = evidenziata ? mappaTitolo(etichetta, modello) : null;
        if (mappa && mappa.punteggio >= 2) {
          apri(etichetta, mappa, 'etichetta');
          corrente.blocchi.push(Object.assign(nuovoParagrafo(), { testo: due[2].trim(), lista: b.lista }));
          return;
        }
      }
      corrente.blocchi.push(b);
    });

    // Le parti prima della prima sezione riconosciuta che contengono solo
    // intestazione (scuola, titolo, docente, classe...) non vanno importate.
    let primaRiconosciuta = false;
    segmenti.forEach((s) => {
      if (primaRiconosciuta) return;
      const linee = lineeDi(s.blocchi);
      const titoloN = norm(s.titolo);
      const disciplinaNelTitolo = riconosciDisciplina(s.titolo, modello) && titoloN.split(' ').length <= 4;
      const quasiTuttiDati = linee.filter((l) => RIGA_DATI.test(l)).length >= Math.max(1, linee.length * 0.5);
      const titoloDocumento = TITOLO_INTESTAZIONE.test(titoloN) || /\b20\d\d\b/.test(titoloN);
      if (s.iniziale || titoloDocumento || disciplinaNelTitolo || !linee.length || quasiTuttiDati) {
        s.intestazione = true;
        s.mappa = null;
      } else if (s.mappa) {
        primaRiconosciuta = true;
      }
    });

    // Unisce le parti di intestazione consecutive in un solo blocco.
    const uniti = [];
    segmenti.forEach((s) => {
      const prec = uniti[uniti.length - 1];
      if (s.intestazione && prec && prec.intestazione) {
        if (!s.iniziale) prec.blocchi.push(Object.assign(nuovoParagrafo(), { testo: s.titolo, sottotitolo: true }));
        prec.blocchi.push(...s.blocchi);
      } else {
        uniti.push(s);
      }
    });
    if (uniti[0] && uniti[0].intestazione) uniti[0].titolo = 'Intestazione del documento';

    let risultato = uniti.map((s) => ({
      titolo: s.titolo,
      blocchi: s.blocchi,
      proposta: s.intestazione ? IGNORA : s.mappa ? s.mappa.id : EXTRA,
      riconosciuta: !!s.mappa,
      intestazione: !!s.intestazione,
      testo: lineeDi(s.blocchi).join('\n')
    })).filter((s) => s.testo);

    // Documento senza alcuna struttura riconoscibile: si conserva tutto il testo.
    if (!risultato.some((s) => !s.intestazione) && risultato.length) {
      risultato = risultato.map((s) => Object.assign(s, { titolo: 'Contenuto del piano', proposta: EXTRA, intestazione: false }));
    }
    return { dati, segmenti: risultato };
  }

  /* ================= Trasformazione nel formato del modello ================= */

  function parole(s) {
    return [...new Set(norm(s).split(' ').filter((t) => t && !STOP.has(t) && (t.length >= 3 || /\d/.test(t))))];
  }

  // "lavori"/"lavoro", "play"/"playing", "verifica"/"verifiche" sono la stessa parola.
  function simili(a, b) {
    if (a === b) return true;
    if (a.length >= 5 && b.length >= 5 && a.slice(0, 5) === b.slice(0, 5)) return true;
    return Math.min(a.length, b.length) >= 4 && (a.startsWith(b) || b.startsWith(a));
  }

  // "LIM / monitor interattivo" → ["LIM", "monitor interattivo"]; le parentesi non contano.
  function alternative(opzione) {
    return String(opzione).replace(/\([^)]*\)/g, ' ').split(/\s+\/\s+/).map((x) => x.trim()).filter(Boolean);
  }

  function somiglianza(pezzo, opzione) {
    const P = parole(pezzo);
    let migliore = { inter: 0, cO: 0, cP: 0 };
    if (!P.length) return migliore;
    alternative(opzione).forEach((alt) => {
      const O = parole(alt);
      if (!O.length) return;
      const inter = O.filter((t) => P.some((q) => simili(t, q))).length;
      const s = { inter, cO: inter / O.length, cP: inter / P.length };
      if (s.cO > migliore.cO || (s.cO === migliore.cO && s.cP > migliore.cP)) migliore = s;
    });
    return migliore;
  }

  const ALIAS = {
    competenze: [
      ['comunicazione nella madrelingua', 'Competenza alfabetica funzionale'],
      ['madrelingua', 'Competenza alfabetica funzionale'],
      ['lingue straniere', 'Competenza multilinguistica'],
      ['competenza matematica', 'Competenza matematica e competenza in scienze, tecnologie e ingegneria'],
      ['scienza e tecnologia', 'Competenza matematica e competenza in scienze, tecnologie e ingegneria'],
      ['imparare a imparare', 'Competenza personale, sociale e capacità di imparare a imparare'],
      ['sociali e civiche', 'Competenza in materia di cittadinanza'],
      ['spirito di iniziativa', 'Competenza imprenditoriale'],
      ['imprenditorial', 'Competenza imprenditoriale'],
      ['espressione cultural', 'Competenza in materia di consapevolezza ed espressione culturali'],
      ['digitale', 'Competenza digitale']
    ]
  };

  function corrispondenze(pezzo, opzioni, { traguardi = false, alias = [] } = {}) {
    if (traguardi) {
      let migliore = null;
      opzioni.forEach((o) => {
        const s = somiglianza(pezzo, o);
        const ok = s.cO >= 0.6 || (s.cP >= 0.7 && s.inter >= 4);
        if (ok && (!migliore || s.cO > migliore.cO)) migliore = { o, cO: s.cO };
      });
      return migliore ? [migliore.o] : [];
    }
    const trovate = opzioni.filter((o) => {
      const s = somiglianza(pezzo, o);
      return s.inter > 0 && (s.cO >= 0.75 || (s.cO >= 0.5 && s.cP >= 0.6));
    });
    const n = norm(pezzo);
    alias.forEach(([a, o]) => { if (opzioni.includes(o) && contienePhrase(n, a) && !trovate.includes(o)) trovate.push(o); });
    return trovate;
  }

  /**
   * Abbina righe di testo alle voci proposte dal modello.
   * @returns {{sel: string[], resto: string[]}}
   */
  function abbina(linee, opzioni, opz = {}) {
    const sel = [];
    const resto = [];
    const aggiungi = (o) => { if (!sel.includes(o)) sel.push(o); };
    linee.map(pulisciVoce).filter(Boolean).forEach((riga) => {
      if (opz.traguardi) {
        const m = corrispondenze(riga, opzioni, opz);
        if (m.length) m.forEach(aggiungi); else resto.push(riga);
        return;
      }
      const pezzi = riga.length <= 300
        ? riga.split(/\s*[;•]\s*|,\s+|\s+[-–]\s+|\.\s+/).map((x) => pulisciVoce(x).replace(/\.$/, '')).filter((x) => x.length > 1)
        : [riga];
      if (pezzi.length > 1) {
        const esiti = pezzi.map((p) => ({ p, m: corrispondenze(p, opzioni, opz) }));
        esiti.forEach((e) => e.m.forEach(aggiungi));
        const riconosciuti = esiti.filter((e) => e.m.length).length;
        if (riconosciuti / pezzi.length >= 0.6) {
          esiti.forEach((e) => { if (!e.m.length) resto.push(e.p); });
          return;
        }
        // Riga breve (un elenco): restano solo i pezzi non riconosciuti.
        // Paragrafo lungo: si conserva intero per non spezzare il discorso.
        if (riconosciuti && riga.length <= 200) { esiti.forEach((e) => { if (!e.m.length) resto.push(e.p); }); return; }
        if (riconosciuti) { resto.push(riga); return; }
      }
      const m = corrispondenze(riga, opzioni, opz);
      m.forEach(aggiungi);
      if (!m.length || riga.length > 80) resto.push(riga);
    });
    // Le frasi introduttive ("Si useranno i seguenti strumenti:") non servono.
    return { sel, resto: resto.filter((r) => !/:\s*$/.test(r)) };
  }

  function testoLibero(blocchi) {
    let out = '';
    blocchi.forEach((b) => {
      if (b.t === 'tab') {
        out += '\n' + b.righe.map((r) => r.filter(Boolean).join(' – ')).filter(Boolean).join('\n') + '\n\n';
      } else if (b.lista) {
        out += '• ' + b.testo.replace(/\n/g, ' ') + '\n';
      } else {
        out += '\n' + b.testo + '\n\n';
      }
    });
    return out.replace(/\n{3,}/g, '\n\n').trim();
  }

  function aggiungiTestoA(attuale, nuovo) {
    const a = String(attuale || '').trim();
    const n = String(nuovo || '').trim();
    if (!n) return a;
    return a ? a + '\n\n' + n : n;
  }

  function aggiungiRighe(attuale, righe) {
    const r = righe.filter(Boolean);
    if (!r.length) return attuale || '';
    return [String(attuale || '').trim(), ...r].filter(Boolean).join('\n');
  }

  // Le voci brevi vanno tra le "altre voci", le frasi lunghe nelle note.
  function smista(v, resto) {
    v.altro = aggiungiRighe(v.altro, resto.filter((r) => r.length <= 160));
    const lunghe = resto.filter((r) => r.length > 160);
    if (lunghe.length && 'note' in v) v.note = aggiungiTestoA(v.note, lunghe.join('\n\n'));
    else if (lunghe.length) v.altro = aggiungiRighe(v.altro, lunghe);
  }

  const NUMERI_LETTERE = { un: 1, uno: 1, una: 1, due: 2, tre: 3, quattro: 4, cinque: 5, sei: 6, sette: 7, otto: 8, nove: 9, dieci: 10 };

  function numeriSituazione(testo) {
    const t = (' ' + String(testo).toLowerCase().replace(/\s+/g, ' ') + ' ')
      .replace(/\b(un|uno|una|due|tre|quattro|cinque|sei|sette|otto|nove|dieci) (?=(alunn|student|allie|ragazz|maschi|femmine))/g, (m, n) => NUMERI_LETTERE[n] + ' ');
    const trova = (...re) => { for (const r of re) { const m = r.exec(t); if (m) return m[1]; } return ''; };
    return {
      alunni: trova(/(?:composta|formata) da (\d{1,2})\b/, /(\d{1,2}) (?:alunni|studenti|allievi|ragazzi e ragazze)/, /(?:alunni|studenti|allievi|totale)\s*[:\-–]\s*(\d{1,2})\b/),
      maschi: trova(/(\d{1,2}) (?:maschi|ragazzi)\b/, /maschi\s*[:\-–]?\s*(\d{1,2})\b/),
      femmine: trova(/(\d{1,2}) (?:femmine|ragazze)\b/, /femmine\s*[:\-–]?\s*(\d{1,2})\b/),
      ripetenti: trova(/(\d{1,2}) (?:alunn[oiae] )?ripetent/, /ripetenti\s*[:\-–]?\s*(\d{1,2})\b/),
      h104: trova(/(\d{1,2}) (?:alunn[oiae] )?(?:con )?(?:disabilit|diversamente abil|certificat|h\b|l\.? ?104)/,
        /(?:l\.? ?104(?:\/\d+)?\)?|disabilit\S*)\s*[:\-–]?\s*(\d{1,2})\b/),
      dsa: trova(/(\d{1,2}) (?:alunn[oiae] )?(?:con )?dsa\b/, /\bdsa\s*[:\-–]?\s*(\d{1,2})\b/),
      bes: trova(/(\d{1,2}) (?:alunn[oiae] )?(?:con )?(?:altri )?bes\b/, /\bbes\s*[:\-–]?\s*(\d{1,2})\b/),
      nai: trova(/(\d{1,2}) (?:alunn[oiae] )?(?:stranier|non italofon|nai\b)/, /(?:stranieri|nai)\s*[:\-–]?\s*(\d{1,2})\b/)
    };
  }

  // Tabelle con le etichette nella prima riga e i numeri sotto → "Etichetta: numero".
  function testoConTabelleNumeriche(blocchi) {
    return blocchi.map((b) => {
      if (b.t !== 'tab' || b.righe.length < 2) return b.t === 'tab' ? b.righe.map((r) => r.join(' – ')).join('\n') : b.testo;
      const [testa, ...corpo] = b.righe;
      if (!corpo.every((r) => r.every((c) => /^\s*\d{0,2}\s*$/.test(c)))) return b.righe.map((r) => r.join(' – ')).join('\n');
      return corpo.map((r) => testa.map((t, i) => t + ': ' + (r[i] || '')).join('\n')).join('\n');
    }).join('\n');
  }

  function applicaSituazione(v, sez, seg, modello) {
    const testoNumeri = testoConTabelleNumeriche(seg.blocchi);
    const num = numeriSituazione(testoNumeri);
    Object.keys(num).forEach((k) => { if (num[k] && !v.num[k]) v.num[k] = num[k]; });
    const t = testoNumeri.toLowerCase();
    modello.livelli.forEach((l) => {
      const m = new RegExp(norm(l).slice(0, 6) + '[^\\d\\n]{0,25}?(\\d{1,2})\\b').exec(norm(t));
      if (m && !v.livelli[l]) v.livelli[l] = m[1];
    });
    const { sel } = abbina(lineeDi(seg.blocchi), sez.opzioni);
    sel.forEach((o) => { if (!v.sel.includes(o)) v.sel.push(o); });
    v.note = aggiungiTestoA(v.note, testoLibero(seg.blocchi));
  }

  /* ---------- Unità di apprendimento ---------- */

  const CAMPI_UDA = {
    numero: ['n$', 'nr$', 'num$', 'n ro$'],
    titolo: ['unita', 'uda$', 'titolo', 'modulo', 'tema$', 'percorso', 'nome$'],
    periodo: ['periodo', 'tempi', 'tempo$', 'mese', 'mesi$', 'quadrimestre', 'trimestre', 'scansione', 'calendario'],
    ore: ['ore$', 'n ore', 'durata', 'monte ore'],
    nuclei: ['nucle', 'ambit', 'area$', 'aree$'],
    obiettivi: ['obiettiv', 'abilita', 'traguard', 'competenz'],
    contenuti: ['contenut', 'conoscenz', 'attivita', 'argoment']
  };

  function campoColonna(intestazione) {
    const n = norm(intestazione);
    let migliore = null;
    Object.keys(CAMPI_UDA).forEach((campo) => {
      CAMPI_UDA[campo].forEach((kw) => {
        const intero = kw.endsWith('$');
        const k = intero ? kw.slice(0, -1) : kw;
        const m = new RegExp('(^| )' + k + (intero ? '( |$)' : '')).exec(n);
        if (m && (!migliore || m.index < migliore.pos)) migliore = { campo, pos: m.index };
      });
    });
    return migliore ? migliore.campo : null;
  }

  const MESI_RE = /settembre|ottobre|novembre|dicembre|gennaio|febbraio|marzo|aprile|maggio|giugno|quadrimestre|trimestre|intero anno|tutto l.anno/i;

  function separaTitoloTempi(u, valore) {
    const righe = String(valore).split('\n').map((r) => r.trim()).filter(Boolean);
    u.titolo = u.titolo || righe.shift() || '';
    righe.forEach((r) => {
      const ore = /(\d{1,3})\s*ore\b/i.exec(r);
      if (ore && !u.ore) u.ore = ore[1];
      const resto = r.replace(/(\d{1,3})\s*ore\b/i, '').replace(/^[\s·,;\-–]+|[\s·,;\-–]+$/g, '');
      if (resto && MESI_RE.test(resto) && !u.periodo) u.periodo = resto;
      else if (resto) u.contenuti = aggiungiRighe(u.contenuti, [resto]);
    });
  }

  function nucleiDa(valore, nucleiDisc) {
    return String(valore).split(/[;\n]/).map(pulisciVoce).filter(Boolean).map((x) => {
      const m = corrispondenze(x, nucleiDisc, { traguardi: true });
      return m[0] || x;
    });
  }

  function unitaDaTabella(righe, nucleiDisc) {
    if (righe.length < 1) return [];
    const colonne = Math.max(...righe.map((r) => r.length));
    // Tabella verticale: una UdA per tabella con etichette nella prima colonna.
    if (colonne === 2) {
      const etichette = righe.map((r) => campoColonna(r[0] || ''));
      if (etichette.filter((e) => e && e !== 'numero').length >= 2) {
        const unita = [];
        let u = null;
        righe.forEach((r, i) => {
          const campo = etichette[i];
          if (!u || (campo === 'titolo' && u.titolo)) { u = M.nuovaUnita(); unita.push(u); }
          assegna(u, campo || 'contenuti', r[1] || '', nucleiDisc, campo ? '' : r[0]);
        });
        return unita.filter((x) => x.titolo || x.obiettivi || x.contenuti);
      }
    }
    const campi = (righe[0] || []).map(campoColonna);
    const haIntestazione = campi.filter((c) => c && c !== 'numero').length >= 1 && righe.length > 1;
    const corpo = haIntestazione ? righe.slice(1) : righe;
    const mappa = haIntestazione ? campi : righe[0].map((x, i) => (i === 0 ? 'titolo' : 'contenuti'));
    return corpo.map((r) => {
      const u = M.nuovaUnita();
      r.forEach((c, i) => {
        if (!String(c).trim()) return;
        const campo = mappa[i] || 'contenuti';
        assegna(u, campo, c, nucleiDisc, mappa[i] ? '' : (haIntestazione ? righe[0][i] : ''));
      });
      if (!u.titolo) {
        const prima = String(u.contenuti || u.obiettivi).split('\n')[0];
        u.titolo = prima.length > 80 ? prima.slice(0, 77) + '…' : prima;
      }
      return u;
    }).filter((u) => u.titolo || u.obiettivi || u.contenuti);
  }

  function assegna(u, campo, valore, nucleiDisc, etichetta) {
    const v = String(valore).trim();
    if (!v || campo === 'numero') return;
    switch (campo) {
      case 'titolo': separaTitoloTempi(u, v); break;
      case 'periodo': u.periodo = u.periodo ? u.periodo + ' ' + v : v.replace(/\n/g, ' '); break;
      case 'ore': { const m = /\d{1,3}/.exec(v); if (m) u.ore = m[0]; break; }
      case 'nuclei': nucleiDa(v, nucleiDisc).forEach((n) => { if (!u.nuclei.includes(n)) u.nuclei.push(n); }); break;
      case 'obiettivi': u.obiettivi = aggiungiRighe(u.obiettivi, [v]); break;
      default: u.contenuti = aggiungiRighe(u.contenuti, [etichetta ? etichetta + ': ' + v : v]);
    }
  }

  const INIZIO_UDA = /^(?:u\.?\s?d\.?\s?a\.?|unit[aà](?: di apprendimento| didattica)?|modulo)\s*(?:n\.?\s*)?(\d{1,2})?\s*[:\-–.)]?\s*(.*)$/i;

  function applicaUda(v, seg, contesto, nucleiDisc) {
    // Tabelle: una riga per unità.
    let trovate = false;
    seg.blocchi.filter((b) => b.t === 'tab').forEach((b) => {
      const u = unitaDaTabella(b.righe, nucleiDisc);
      if (u.length) { v.unita.push(...u); trovate = true; }
    });
    // Testo: "UdA 1 – Titolo", oppure sottotitoli in grassetto/maiuscolo.
    const nTitolo = norm(seg.titolo);
    const campoBase = /obiettiv|abilita/.test(nTitolo) && !/contenut/.test(nTitolo) ? 'obiettivi' : 'contenuti';
    let u = null;
    let campo = campoBase;
    let periodo = '';
    const nuova = (titolo) => {
      u = M.nuovaUnita();
      u.titolo = titolo;
      u.periodo = periodo;
      v.unita.push(u);
      campo = campoBase;
      trovate = true;
    };
    seg.blocchi.filter((b) => b.t === 'p').forEach((b) => {
      b.testo.split('\n').map((r) => r.trim()).filter(Boolean).forEach((riga, k) => {
        let m;
        if (k === 0 && b.sottotitolo) {
          if (MESI_RE.test(riga) && riga.length <= 40) { periodo = pulisciTitolo(riga); if (u && !u.periodo && !u.obiettivi && !u.contenuti) u.periodo = periodo; return; }
          nuova(pulisciTitolo(riga));
          return;
        }
        if ((m = INIZIO_UDA.exec(riga)) && riga.length <= 120 && !b.lista) {
          nuova(pulisciTitolo(m[2] || riga));
          return;
        }
        if ((m = /^(periodo|tempi|tempo|durata)\s*[:\-–]\s*(.+)$/i.exec(riga))) {
          if (!u) nuova('Programmazione');
          const ore = /(\d{1,3})\s*ore\b/i.exec(m[2]);
          if (ore) u.ore = ore[1];
          u.periodo = m[2].replace(/\(?\d{1,3}\s*ore\)?/i, '').replace(/[\s,;\-–]+$/, '').trim();
          return;
        }
        if ((m = /^(obiettiv\w*(?: di apprendimento)?|abilit[aà])\s*[:\-–]?\s*(.*)$/i.exec(riga))) {
          if (!u) nuova('Programmazione');
          campo = 'obiettivi';
          if (m[2]) u.obiettivi = aggiungiRighe(u.obiettivi, [pulisciVoce(m[2])]);
          return;
        }
        if ((m = /^(contenut\w*|conoscenz\w*|argoment\w*|attivit[aà])\s*[:\-–]?\s*(.*)$/i.exec(riga))) {
          if (!u) nuova('Programmazione');
          campo = 'contenuti';
          if (m[2]) u.contenuti = aggiungiRighe(u.contenuti, [pulisciVoce(m[2])]);
          return;
        }
        if (!u) {
          // Testo senza unità: confluisce in un'unica unità condivisa tra le sezioni
          // (es. "Obiettivi" e "Contenuti" scritti in due parti diverse del documento).
          if (!contesto.unitaComune) {
            contesto.unitaComune = M.nuovaUnita();
            contesto.unitaComune.titolo = 'Programmazione annuale';
            v.unita.push(contesto.unitaComune);
          }
          const c = contesto.unitaComune;
          c[campoBase] = aggiungiRighe(c[campoBase], [pulisciVoce(riga)]);
          trovate = true;
          return;
        }
        u[campo] = aggiungiRighe(u[campo], [pulisciVoce(riga)]);
      });
    });
    return trovate;
  }

  /* ---------- Creazione del piano ---------- */

  function applicaSegmento(piano, sezione, seg, modello, contesto) {
    const v = M.valoreSezione(piano, sezione);
    const linee = lineeDi(seg.blocchi);
    switch (sezione.tipo) {
      case 'situazione': applicaSituazione(v, sezione, seg, modello); break;
      case 'traguardi': {
        const opzioni = M.opzioniSezione(sezione, piano, modello);
        const { sel, resto } = abbina(linee, opzioni, { traguardi: true });
        sel.forEach((o) => { if (!v.sel.includes(o)) v.sel.push(o); });
        v.altro = aggiungiRighe(v.altro, resto);
        break;
      }
      case 'checklist': {
        const { sel, resto } = abbina(linee, sezione.opzioni, { alias: ALIAS[sezione.id] || [] });
        sel.forEach((o) => { if (!v.sel.includes(o)) v.sel.push(o); });
        smista(v, resto);
        const campi = (sezione.campi || []).map((c) => c.id);
        const ore = /(\d{1,3})\s*ore\b/i.exec(seg.testo);
        if (ore && campi.includes('ore') && !v.campi.ore) v.campi.ore = ore[1];
        ['scritte', 'orali', 'pratiche'].forEach((tipo) => {
          const re = new RegExp('(\\d|due|tre|quattro|cinque|sei)\\s+(?:verifiche|prove)\\s+' + tipo.slice(0, -1), 'i');
          const m = re.exec(seg.testo);
          if (m && campi.includes(tipo) && !v.campi[tipo]) v.campi[tipo] = String(NUMERI_LETTERE[m[1].toLowerCase()] || m[1]);
        });
        break;
      }
      case 'uda': {
        const disc = M.trovaDisciplina(modello, piano.disciplina);
        applicaUda(v, seg, contesto, disc ? disc.nuclei : []);
        break;
      }
      case 'testo': v.note = aggiungiTestoA(v.note, testoLibero(seg.blocchi)); break;
      default: break;
    }
  }

  /**
   * Crea il piano uniformato.
   * @param {{segmenti: Array}} analisi  risultato di analizza()
   * @param {string[]} scelte  destinazione di ogni segmento (id sezione, 'extra' o 'ignora')
   * @param {object} dati  dati generali confermati dal docente
   */
  function creaPiano(analisi, scelte, dati, modello) {
    const piano = M.nuovoPiano(modello, dati);
    const contesto = {};
    analisi.segmenti.forEach((seg, i) => {
      const dest = scelte[i] || seg.proposta;
      if (dest === IGNORA) return;
      if (dest === EXTRA) {
        const t = testoLibero(seg.blocchi);
        if (t) piano.extra = (piano.extra || []).concat(M.nuovaSezioneExtra(seg.titolo, t));
        return;
      }
      const sezione = modello.sezioni.find((s) => s.id === dest);
      if (sezione) applicaSegmento(piano, sezione, seg, modello, contesto);
    });
    return piano;
  }

  const api = {
    EXTRA, IGNORA,
    leggiFile, leggiDocx, leggiOdt, leggiTesto,
    analizza, mappaTitolo, riconosciDati, riconosciDisciplina,
    abbina, numeriSituazione, unitaDaTabella, creaPiano,
    norm
  };
  root.PDL = root.PDL || {};
  root.PDL.importa = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
