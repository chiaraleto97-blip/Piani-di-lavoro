'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
require('../js/defaults.js');
const M = require('../js/model.js');
require('../js/documento.js');
const Z = require('../js/zip.js');
const W = require('../js/word.js');
const I = require('../js/importa.js');

const modello = () => M.normalizzaModello(null);
const campione = (nome) => new Uint8Array(fs.readFileSync(path.join(__dirname, 'campioni', nome)));

async function importa(nome, dati, m = modello()) {
  const analisi = I.analizza(await I.leggiFile(nome, dati), m);
  const piano = I.creaPiano(analisi, analisi.segmenti.map((s) => s.proposta), Object.assign({ classe: '1', sezione: 'A' }, analisi.dati), m);
  return { analisi, piano, m };
}

test('zip: crea e rilegge file compressi e non compressi', async () => {
  const testo = 'Piano di lavoro '.repeat(200);
  const archivio = await Z.crea([{ nome: 'a.txt', dati: testo }, { nome: 'cartella/b.bin', dati: new Uint8Array([1, 2, 3]) }]);
  const z = Z.apri(archivio);
  assert.deepEqual(z.nomi, ['a.txt', 'cartella/b.bin']);
  assert.equal(await z.testo('a.txt'), testo);
  assert.deepEqual([...(await z.file('cartella/b.bin'))], [1, 2, 3]);
  assert.equal(await z.file('manca.txt'), null);
  assert.equal(Z.crc32(new TextEncoder().encode('123456789')), 0xCBF43926);
  assert.throws(() => Z.apri(new Uint8Array(40)), /non è un documento valido/);
});

test('word: il .docx contiene stili, numerazione e immagine della firma', async () => {
  const m = modello();
  const p = M.nuovoPiano(m, { docente: 'Anna & Co', disciplina: 'Storia', classe: '2', sezione: 'B' });
  p.valori.traguardi.sel = M.trovaDisciplina(m, 'Storia').traguardi.slice(0, 2);
  // PNG 1x1 valido
  p.firma = { immagine: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', data: '2026-10-25' };
  const z = Z.apri(await W.creaDocx([p], m));
  for (const n of ['[Content_Types].xml', '_rels/.rels', 'word/document.xml', 'word/styles.xml', 'word/numbering.xml', 'word/_rels/document.xml.rels', 'word/media/firma1.png']) {
    assert.ok(z.nomi.includes(n), 'manca ' + n);
  }
  const doc = await z.testo('word/document.xml');
  assert.match(doc, /Anna &amp; Co/);
  assert.match(doc, /<w:pStyle w:val="Heading2"\/>/);
  assert.match(doc, /r:embed="rIdFirma1"/);
  assert.match(doc, /25\/10\/2026/);
  assert.match(await z.testo('word/_rels/document.xml.rels'), /media\/firma1\.png/);
});

test('importazione: un piano esportato dall\'app ritorna nelle stesse sezioni', async () => {
  const m = modello();
  const p = M.nuovoPiano(m, { docente: 'Paolo Gialli', disciplina: 'Matematica', classe: '3', sezione: 'A', anno: '2026/2027' });
  p.valori.situazione.num = { alunni: '20', maschi: '9', femmine: '11', dsa: '2' };
  p.valori.metodologie.sel = ['Problem solving', 'Flipped classroom'];
  p.valori.strumenti.sel = ['Dizionari'];
  p.valori.traguardi.sel = M.trovaDisciplina(m, 'Matematica').traguardi.slice(0, 2);
  p.valori.uda.unita = [{ titolo: 'Equazioni', periodo: 'Ottobre', ore: '12', nuclei: ['Relazioni e funzioni'], obiettivi: 'Risolvere equazioni', contenuti: 'Principi di equivalenza' }];
  const { analisi, piano } = await importa('mio.docx', await W.creaDocx([p], m), m);
  assert.equal(analisi.dati.docente, 'Paolo Gialli');
  assert.equal(analisi.dati.disciplina, 'Matematica');
  assert.equal(analisi.dati.classe, '3');
  assert.equal(analisi.dati.sezione, 'A');
  assert.equal(analisi.dati.anno, '2026/2027');
  assert.equal(piano.valori.situazione.num.alunni, '20');
  assert.equal(piano.valori.situazione.num.dsa, '2');
  assert.equal(piano.valori.situazione.num.maschi, '9');
  assert.deepEqual(piano.valori.metodologie.sel.sort(), ['Flipped classroom', 'Problem solving']);
  assert.deepEqual(piano.valori.strumenti.sel, ['Dizionari']);
  assert.equal(piano.valori.traguardi.sel.length, 2);
  assert.equal(piano.valori.traguardi.altro, '');
  const u = piano.valori.uda.unita[0];
  assert.equal(u.titolo, 'Equazioni');
  assert.equal(u.periodo, 'Ottobre');
  assert.equal(u.ore, '12');
  assert.deepEqual(u.nuclei, ['Relazioni e funzioni']);
  assert.equal(u.obiettivi, 'Risolvere equazioni');
});

test('importazione: documento Word con titoli e tabella UdA', async () => {
  const { analisi, piano } = await importa('piano.docx', campione('piano-titoli-word.docx'));
  assert.deepEqual(analisi.dati, { anno: '2025/2026', docente: 'Laura Verdi', disciplina: 'Matematica', classe: '2', sezione: 'B', libro: 'Contaci! vol. 2' });
  const proposte = analisi.segmenti.map((s) => s.proposta);
  assert.deepEqual(proposte, ['ignora', 'situazione', 'competenze', 'traguardi', 'uda', 'metodologie', 'strumenti', 'verifiche', 'recupero', 'extra']);
  const v = piano.valori;
  assert.deepEqual(v.situazione.num, { alunni: '21', maschi: '11', femmine: '10', h104: '1', dsa: '2' });
  assert.deepEqual(v.situazione.livelli, { Avanzato: '4', Intermedio: '8', Base: '6', Iniziale: '3' });
  // Le competenze del 2006 vengono ricondotte a quelle del 2018.
  assert.ok(v.competenze.sel.includes('Competenza alfabetica funzionale'));
  assert.ok(v.competenze.sel.includes('Competenza in materia di cittadinanza'));
  assert.equal(v.traguardi.sel.length, 2);
  assert.match(v.traguardi.altro, /lavorare in gruppo/);
  assert.equal(v.uda.unita.length, 3);
  assert.equal(v.uda.unita[1].titolo, 'Radice quadrata');
  assert.ok(v.metodologie.sel.includes('Cooperative learning'));
  assert.ok(v.strumenti.sel.includes('LIM / monitor interattivo'));
  assert.equal(v.strumenti.altro, 'software GeoGebra');
  assert.equal(v.verifiche.campi.scritte, '3');
  assert.equal(piano.extra.length, 1);
  assert.equal(piano.extra[0].titolo, 'Uscite didattiche');
});

test('importazione: lo stesso documento in formato .odt dà lo stesso risultato', async () => {
  const docx = await importa('a.docx', campione('piano-titoli-word.docx'));
  const odt = await importa('a.odt', campione('piano-titoli-libreoffice.odt'));
  assert.deepEqual(odt.analisi.dati, docx.analisi.dati);
  assert.deepEqual(odt.analisi.segmenti.map((s) => s.proposta), docx.analisi.segmenti.map((s) => s.proposta));
  assert.deepEqual(odt.piano.valori.uda, docx.piano.valori.uda);
});

test('importazione: tabelle etichetta/contenuto e titoli in grassetto', async () => {
  const { analisi, piano } = await importa('b.docx', campione('piano-tabelle-grassetti.docx'));
  assert.equal(analisi.dati.docente, 'Marco Bianchi');
  assert.equal(analisi.dati.disciplina, 'Lingua inglese');
  assert.equal(analisi.dati.classe, '3');
  assert.equal(analisi.dati.sezione, 'C');
  assert.equal(analisi.dati.libro, undefined, 'gli strumenti non sono il libro di testo');
  assert.equal(analisi.segmenti[0].proposta, 'ignora');
  const v = piano.valori;
  assert.equal(v.situazione.num.alunni, '19');
  assert.equal(v.situazione.num.ripetenti, '1');
  assert.ok(v.metodologie.sel.includes('Role playing'));
  assert.ok(v.inclusione.sel.includes('Misure dispensative'));
  assert.equal(v.edcivica.campi.ore, '4');
  assert.deepEqual(v.uda.unita.map((u) => [u.titolo, u.periodo]), [['UNIT 1 - Holidays', 'settembre-ottobre'], ['UNIT 2 - Future plans', 'novembre-dicembre']]);
  assert.ok(v.famiglie.sel.includes('Registro elettronico'));
});

test('importazione: testo semplice con titoli in maiuscolo', async () => {
  const { analisi, piano } = await importa('c.txt', campione('piano-testo.txt'));
  assert.equal(analisi.dati.docente, 'Giulia Neri');
  assert.equal(analisi.dati.disciplina, 'Tecnologia');
  assert.equal(analisi.dati.classe, '1');
  assert.equal(analisi.dati.sezione, 'D');
  const v = piano.valori;
  assert.equal(v.situazione.num.alunni, '24');
  assert.equal(v.situazione.num.dsa, '3');
  assert.deepEqual(v.metodologie.sel.sort(), ['Compiti di realtà', 'Didattica laboratoriale', 'Lezione frontale']);
  assert.deepEqual(v.uda.unita.map((u) => u.titolo), ['Il disegno tecnico', 'I materiali']);
  assert.equal(v.uda.unita[0].periodo, 'ottobre-dicembre');
  assert.equal(v.uda.unita[0].obiettivi, 'usare gli strumenti del disegno');
  // "PROGETTO ORTO" è un titolo dello stesso livello di "CONTENUTI": nuova sezione, non un'unità.
  assert.equal(piano.extra.length, 1);
  assert.equal(piano.extra[0].titolo, 'Progetto orto');
  assert.match(piano.extra[0].testo, /orto scolastico/);
});

test('importazione: un documento senza struttura viene conservato per intero', async () => {
  const testo = 'Quest\'anno lavoreremo molto sulla lettura.\nFaremo anche uscite sul territorio.';
  const { analisi, piano } = await importa('x.txt', new TextEncoder().encode(testo));
  assert.equal(analisi.segmenti.length, 1);
  assert.equal(analisi.segmenti[0].proposta, 'extra');
  assert.match(piano.extra[0].testo, /uscite sul territorio/);
});

test('importazione: formati non supportati danno un messaggio chiaro', async () => {
  await assert.rejects(I.leggiFile('vecchio.doc', new Uint8Array(10)), /salvalo come \.docx/);
  await assert.rejects(I.leggiFile('scansione.pdf', new Uint8Array(10)), /PDF/);
  await assert.rejects(I.leggiFile('foto.jpg', new Uint8Array(10)), /Formato non supportato/);
});

test('abbinamento delle voci al modello', () => {
  const opzioni = ['Lezione frontale', 'Lavori di gruppo', 'Peer tutoring / peer education', 'Libro di testo (anche in versione digitale)'];
  const r = I.abbina(['Lezione frontale, lavoro di gruppo, peer education, uscite sul territorio'], opzioni);
  assert.deepEqual(r.sel, ['Lezione frontale', 'Lavori di gruppo', 'Peer tutoring / peer education']);
  assert.deepEqual(r.resto, ['uscite sul territorio']);
  assert.deepEqual(I.abbina(['Si useranno i seguenti strumenti:', 'libro di testo'], opzioni), { sel: ['Libro di testo (anche in versione digitale)'], resto: [] });
});

test('riconoscimento dei dati e dei numeri della classe', () => {
  const m = modello();
  assert.deepEqual(I.riconosciDati(['Prof. Mario Rossi', 'classe III sez. B', 'a.s. 2025-26', 'Materia: scienze motorie'], m),
    { docente: 'Mario Rossi', classe: '3', sezione: 'B', anno: '2025/2026', disciplina: 'Educazione fisica' });
  assert.equal(I.riconosciDati(['Profilo della classe'], m).docente, undefined);
  const n = I.numeriSituazione('La classe è composta da 18 alunni, tra cui un alunno con disabilità e due alunni con DSA.');
  assert.equal(n.alunni, '18');
  assert.equal(n.h104, '1');
  assert.equal(n.dsa, '2');
});
