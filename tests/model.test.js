'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const defaults = require('../js/defaults.js');
const M = require('../js/model.js');
const D = require('../js/documento.js');

const modello = () => M.normalizzaModello(null);
const sezione = (m, id) => m.sezioni.find((s) => s.id === id);

test('anno scolastico: da agosto inizia il nuovo anno', () => {
  assert.equal(defaults.annoScolasticoCorrente(new Date(2026, 8, 29)), '2026/2027');
  assert.equal(defaults.annoScolasticoCorrente(new Date(2027, 2, 1)), '2026/2027');
  assert.equal(defaults.annoScolasticoCorrente(new Date(2027, 7, 1)), '2027/2028');
});

test('il modello predefinito contiene discipline con traguardi e nuclei', () => {
  const m = modello();
  assert.equal(m.sezioni[0].tipo, 'dati');
  assert.ok(m.discipline.length >= 12);
  m.discipline.forEach((d) => {
    assert.ok(d.traguardi.length > 0, d.nome + ' senza traguardi');
    assert.ok(d.nuclei.length > 0, d.nome + ' senza nuclei');
  });
  const ids = m.sezioni.map((s) => s.id);
  assert.equal(new Set(ids).size, ids.length, 'id delle sezioni duplicati');
});

test('normalizzaModello rimette la sezione dati in testa e scarta voci non valide', () => {
  const m = M.normalizzaModello({
    sezioni: [
      { id: 'x', tipo: 'checklist', titolo: 'X', opzioni: ['a', '', '  b '] },
      { id: 'y', tipo: 'sconosciuto' },
      { id: 'dati', tipo: 'dati', titolo: 'Dati', attiva: false },
      { id: 'x', tipo: 'testo' }
    ],
    discipline: [{ nome: 'Italiano' }, { nome: 'italiano' }, { nome: '' }]
  });
  assert.deepEqual(m.sezioni.map((s) => s.id), ['dati', 'x']);
  assert.equal(m.sezioni[0].attiva, true);
  assert.deepEqual(m.sezioni[1].opzioni, ['a', 'b']);
  assert.equal(m.discipline.length, 1);
});

test('nuovoPiano precompila ore dalla disciplina e testi predefiniti', () => {
  const m = modello();
  const p = M.nuovoPiano(m, { docente: 'Rossi', disciplina: 'Matematica', classe: '2', sezione: 'b' });
  assert.equal(p.ore, '4');
  assert.equal(p.sezione, 'B');
  assert.equal(p.anno, m.scuola.annoScolastico);
  assert.match(p.valori.valutazione.note, /D\.Lgs\. 62\/2017/);
  assert.equal(M.titoloPiano(p), 'Matematica – classe 2B');
});

test('elencoSezioniClasse separa e deduplica', () => {
  assert.deepEqual(M.elencoSezioniClasse('a, B;c b  /d'), ['A', 'B', 'C', 'D']);
  assert.deepEqual(M.elencoSezioniClasse('  '), []);
});

test('avanzamento e completamento delle sezioni', () => {
  const m = modello();
  const p = M.nuovoPiano(m, { docente: 'Rossi', disciplina: 'Italiano', classe: '1', sezione: 'A' });
  const iniziale = M.avanzamento(p, m);
  // Solo i dati generali: un testo predefinito non basta a considerare compilata una sezione.
  assert.equal(iniziale.fatte, 1);
  assert.equal(sezione(m, 'note').facoltativa, true, 'le osservazioni sono facoltative');
  p.valori.metodologie.sel.push('Lezione frontale');
  p.valori.situazione.num.alunni = '22';
  p.valori.uda.unita.push(Object.assign(M.nuovaUnita(), { titolo: 'Il testo narrativo' }));
  assert.equal(M.avanzamento(p, m).fatte, iniziale.fatte + 3);
  assert.equal(M.sezioneCompleta(p, sezione(m, 'traguardi')), false);
  p.valori.traguardi.altro = 'Traguardo personalizzato';
  assert.equal(M.sezioneCompleta(p, sezione(m, 'traguardi')), true);
});

test('copiaContenuti azzera la situazione di partenza', () => {
  const m = modello();
  const a = M.nuovoPiano(m, { docente: 'Rossi', disciplina: 'Storia', classe: '3', sezione: 'A' });
  a.valori.situazione.num.alunni = '20';
  a.valori.metodologie.sel = ['Debate'];
  const b = M.copiaContenuti(M.nuovoPiano(m, { docente: 'Rossi', disciplina: 'Storia', classe: '3', sezione: 'B' }), a, m);
  assert.deepEqual(b.valori.metodologie.sel, ['Debate']);
  assert.equal(b.valori.situazione.num.alunni, undefined);
  b.valori.metodologie.sel.push('CLIL');
  assert.deepEqual(a.valori.metodologie.sel, ['Debate'], 'la copia non deve condividere riferimenti');
});

test('avvisi di coerenza della situazione di partenza', () => {
  const v = { num: { alunni: 20, maschi: 10, femmine: 9 }, livelli: { Base: 5, Iniziale: 5 } };
  const avvisi = M.avvisiSituazione(v, ['Avanzato', 'Intermedio', 'Base', 'Iniziale']);
  assert.equal(avvisi.length, 2);
  assert.deepEqual(M.avvisiSituazione({ num: { alunni: 3, maschi: 1, femmine: 2 }, livelli: {} }, []), []);
});

test('leggiPacchetto riconosce modelli e piani e rifiuta file estranei', () => {
  const m = modello();
  const p = M.nuovoPiano(m, { disciplina: 'Musica' });
  assert.equal(M.leggiPacchetto(JSON.stringify(M.pacchettoModello(m))).tipo, 'modello');
  const letti = M.leggiPacchetto(JSON.stringify(M.pacchettoPiani([p])));
  assert.equal(letti.tipo, 'piani');
  assert.equal(letti.piani[0].id, p.id);
  assert.ok(M.leggiPacchetto('non json').errore);
  assert.ok(M.leggiPacchetto('{"a":1}').errore);
});

test('unisciPiani: a parità di id prevale il più recente', () => {
  const vecchio = { id: 'a', modificato: '2026-01-01', docente: 'x' };
  const nuovo = { id: 'a', modificato: '2026-02-01', docente: 'y' };
  const r = M.unisciPiani([vecchio], [nuovo, { id: 'b', modificato: '2026-01-01' }]);
  assert.equal(r.piani.length, 2);
  assert.equal(r.piani[0].docente, 'y');
  assert.deepEqual(r.stat, { aggiunti: 1, aggiornati: 1, ignorati: 0 });
  assert.equal(M.unisciPiani(r.piani, [vecchio]).stat.ignorati, 1);
});

test('il documento contiene i dati inseriti e ne esegue l\'escape', () => {
  const m = modello();
  m.scuola.nome = 'IC <Test>';
  const p = M.nuovoPiano(m, { docente: 'Anna "Bianchi"', disciplina: 'Tecnologia', classe: '1', sezione: 'C' });
  p.valori.metodologie.sel = ['Problem solving', 'Voce <script>alert(1)</script>'];
  p.valori.traguardi.sel = [M.trovaDisciplina(m, 'Tecnologia').traguardi[0]];
  p.valori.uda.unita.push({ titolo: 'Disegno tecnico', periodo: 'Ottobre', ore: '8', nuclei: ['Prevedere, immaginare e progettare'], obiettivi: 'Usare squadre\nUsare il compasso', contenuti: '' });
  const html = D.renderPiano(p, m);
  assert.match(html, /IC &lt;Test&gt;/);
  assert.match(html, /Anna &quot;Bianchi&quot;/);
  assert.doesNotMatch(html, /<script>/);
  assert.match(html, /Problem solving/);
  assert.match(html, /Disegno tecnico/);
  assert.match(html, /Usare squadre<br>Usare il compasso/);
  assert.match(html, /Sezione non compilata/);
});

test('firma: solo immagini PNG valide, non copiata nei duplicati, mostrata nel documento', () => {
  const m = modello();
  const png = 'data:image/png;base64,iVBORw0KGgo=';
  assert.equal(M.firmaValida({ immagine: 'javascript:alert(1)' }), null);
  assert.equal(M.firmaValida({ immagine: 'data:image/svg+xml;base64,PHN2Zz4=' }), null);
  const f = M.firmaValida({ immagine: png, data: '2026-10-25', momento: '2026-10-25T10:00:00.000Z' });
  assert.equal(f.data, '2026-10-25');

  const p = M.nuovoPiano(m, { docente: 'Rossi', disciplina: 'Musica', classe: '1', sezione: 'A' });
  p.firma = f;
  p.extra = [{ titolo: 'Concerto di Natale', testo: 'Prove a dicembre.' }];
  const letto = M.normalizzaPiano(JSON.parse(JSON.stringify(p)));
  assert.deepEqual(letto.firma, f);
  assert.equal(letto.extra[0].titolo, 'Concerto di Natale');

  const copia = M.copiaContenuti(M.nuovoPiano(m, { disciplina: 'Musica' }), p, m);
  assert.equal(copia.firma, undefined);
  assert.equal(copia.extra[0].testo, 'Prove a dicembre.');

  p.modificato = '2026-10-25T09:00:00.000Z';
  assert.equal(M.modificatoDopoFirma(p), false);
  p.modificato = '2026-10-26T09:00:00.000Z';
  assert.equal(M.modificatoDopoFirma(p), true);

  const html = D.renderPiano(p, m);
  assert.match(html, /<img class="doc-firma-img"/);
  assert.match(html, /25\/10\/2026/);
  assert.match(html, /14\. Concerto di Natale/);
});

test('le sezioni disattivate non compaiono nel documento', () => {
  const m = modello();
  sezione(m, 'famiglie').attiva = false;
  const p = M.nuovoPiano(m, { disciplina: 'Italiano' });
  assert.doesNotMatch(D.renderPiano(p, m), /Rapporti con le famiglie/);
});

test('nomeFile produce un nome valido', () => {
  const p = { disciplina: 'Arte e immagine', classe: '3', sezione: 'A', anno: '2026/2027' };
  assert.equal(M.nomeFile(p, 'doc'), 'Piano di lavoro - Arte e immagine - 3A - 2026-2027.doc');
});
