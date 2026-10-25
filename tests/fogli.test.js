'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
require('../js/defaults.js');
const M = require('../js/model.js');
require('../js/zip.js');
require('../js/importa.js');
const F = require('../js/fogli.js');

const modello = () => M.normalizzaModello(null);
const campione = (nome) => new Uint8Array(fs.readFileSync(path.join(__dirname, 'campioni', nome)));

// Risposta simulata di fetch.
const risposta = (corpo, stato = 200) => ({
  ok: stato >= 200 && stato < 300, status: stato,
  text: async () => corpo, arrayBuffer: async () => (typeof corpo === 'string' ? new TextEncoder().encode(corpo) : corpo).buffer
});

test('CSV: virgole, punto e virgola, virgolette e righe vuote', () => {
  assert.deepEqual(F.leggiCsv('"Classe","Sezione","Libro"\n"1","A","Storia, vol. ""1"""\n\n'), [['Classe', 'Sezione', 'Libro'], ['1', 'A', 'Storia, vol. "1"']]);
  assert.deepEqual(F.leggiCsv('﻿Classe;Alunni\r\n2;21\r\n'), [['Classe', 'Alunni'], ['2', '21']]);
});

test('classe scritta in modi diversi', () => {
  assert.deepEqual(F.leggiClasse('2'), { classe: '2', sezione: '' });
  assert.deepEqual(F.leggiClasse('2^A'), { classe: '2', sezione: 'A' });
  assert.deepEqual(F.leggiClasse('III c'), { classe: '3', sezione: 'C' });
  assert.deepEqual(F.leggiClasse('prima'), { classe: '1', sezione: '' });
  assert.equal(F.leggiClasse('4'), null);
});

test('foglio Excel risalvato da LibreOffice (schede Classi e Libri)', async () => {
  const r = await F.daFile('foglio.xlsx', campione('foglio-comune.xlsx'), modello());
  assert.equal(r.classi.length, 4);
  assert.deepEqual(r.classi[0], { classe: '1', sezione: 'A', plesso: 'Sede di Almese', alunni: '22', maschi: '12', femmine: '10', h104: '1', dsa: '2', bes: '0' });
  assert.deepEqual(r.classi[2].classe + r.classi[2].sezione, '2A', 'classe "2^A" in una sola colonna');
  assert.equal(r.classi[1].h104, '', 'le celle vuote restano vuote');
  assert.deepEqual(r.libri.map((l) => l.disciplina), ['Matematica', 'Matematica', 'Lingua inglese', 'Educazione fisica']);
  assert.deepEqual(r.avvisi, []);
});

test('avvisi: colonna PDP unica e valori non numerici', () => {
  const r = F.interpreta({ Classi: [['Classe', 'Sezione', 'Alunni', 'PDP'], ['1', 'A', 'venti', '3']] }, modello());
  assert.equal(r.classi[0].alunni, '');
  assert.equal(r.avvisi.length, 2);
  assert.match(r.avvisi.join(' '), /PDP/);
  assert.match(r.avvisi.join(' '), /non è un numero/);
});

test('modello del foglio: si rilegge e contiene tutte le classi e discipline', async () => {
  const m = modello();
  const schede = await F.leggiXlsx(await F.creaModelloFoglio(m, ['A', 'B']));
  assert.deepEqual(Object.keys(schede), ['Classi', 'Libri', 'Istruzioni']);
  assert.equal(schede.Classi.length, 1 + 3 * 2);
  assert.equal(schede.Libri.length, 1 + 3 * m.discipline.length);
  const r = F.interpreta(schede, m);
  assert.equal(r.classi.length, 6);
  assert.equal(r.libri.length, 0, 'libri ancora da compilare');
});

test('link di Google Fogli: indirizzi di lettura', () => {
  const s = F.indirizziGoogle('https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789/edit?usp=sharing');
  assert.equal(s.tipo, 'schede');
  assert.match(s.url.Classi, /\/d\/1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789\/gviz\/tq\?tqx=out:csv&headers=1&sheet=Classi$/);
  const p = F.indirizziGoogle('https://docs.google.com/spreadsheets/d/e/2PACX-1vQabc_def-123/pubhtml');
  assert.deepEqual(p, { tipo: 'xlsx', url: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vQabc_def-123/pub?output=xlsx' });
  assert.equal(F.indirizziGoogle('https://example.com/foglio'), null);
});

test('lettura dal link: schede, scheda mancante, foglio non condiviso, errori', async () => {
  const m = modello();
  const link = 'https://docs.google.com/spreadsheets/d/1AbCdEfGhIjKlMnOpQrStUvWxYz0123456789/edit';
  const classi = '"Classe","Sezione","Plesso","Alunni","Alunni con PEI"\n"1","A","Sede","20","1"\n';
  const libri = '"Classe","Sezione","Disciplina","Libro di testo"\n"1","","Storia","Noi e la storia 1"\n';
  const r = await F.daLink(link, m, async (url) => risposta(/sheet=Libri/.test(url) ? libri : classi));
  assert.equal(r.classi[0].h104, '1');
  assert.equal(r.libri[0].libro, 'Noi e la storia 1');
  // Senza scheda "Libri" Google restituisce di nuovo la prima scheda.
  const senzaLibri = await F.daLink(link, m, async () => risposta(classi));
  assert.equal(senzaLibri.classi.length, 1);
  await assert.rejects(F.daLink(link, m, async () => risposta('<!DOCTYPE html><html>Accedi</html>')), /non è condiviso/);
  await assert.rejects(F.daLink(link, m, async () => risposta('', 404)), /errore 404/);
  await assert.rejects(F.daLink(link, m, async () => { throw new TypeError('Failed to fetch'); }), /non è raggiungibile/);
  await assert.rejects(F.daLink('ciao', m, async () => risposta('')), /non sembra un foglio Google/);
});

test('il modello predefinito ha i dati comuni vuoti', () => {
  assert.deepEqual(M.normalizzaModello(null).datiComuni, { url: '', aggiornato: '', anno: '', classi: [], libri: [] });
  assert.equal(M.datiComuniPer(M.normalizzaModello(null), { classe: '1', sezione: 'A' }), null);
});

test('dati comuni applicati ai piani della classe', () => {
  const m = modello();
  m.datiComuni = M.normalizzaDatiComuni({
    anno: m.scuola.annoScolastico,
    classi: [{ classe: '2', sezione: 'b', plesso: 'Rubiana', alunni: '18', maschi: '', femmine: '', h104: '2', dsa: '3', bes: '' }],
    libri: [
      { classe: '2', sezione: '', disciplina: 'Matematica', libro: 'Libro per tutte le sezioni' },
      { classe: '2', sezione: 'B', disciplina: 'Matematica', libro: 'Libro della 2B' },
      { classe: '2', sezione: '', disciplina: 'Storia', libro: 'Storia 2' }
    ]
  });
  const p = M.nuovoPiano(m, { docente: 'R', disciplina: 'Matematica', classe: '2', sezione: 'B', libro: 'vecchio' });
  p.valori.situazione.num.maschi = '9';
  p.valori.situazione.num.bes = '1';
  assert.equal(M.applicaDatiComuni(p, m), true);
  assert.equal(p.plesso, 'Rubiana');
  assert.equal(p.libro, 'Libro della 2B', 'il libro della sezione vince su quello generale');
  assert.deepEqual(p.valori.situazione.num, { maschi: '9', bes: '1', alunni: '18', h104: '2', dsa: '3' }, 'le celle vuote non sovrascrivono');
  assert.equal(M.applicaDatiComuni(p, m), false, 'già allineato');

  const a = M.nuovoPiano(m, { disciplina: 'Matematica', classe: '2', sezione: 'A' });
  M.applicaDatiComuni(a, m);
  assert.equal(a.libro, 'Libro per tutte le sezioni');
  assert.equal(a.plesso, '');

  const vecchio = M.nuovoPiano(m, { disciplina: 'Matematica', classe: '2', sezione: 'B', anno: '2019/2020' });
  assert.equal(M.applicaDatiComuni(vecchio, m), false, 'i piani di altri anni non si toccano');

  const letto = M.normalizzaModello(JSON.parse(JSON.stringify(m)));
  assert.equal(letto.datiComuni.classi[0].sezione, 'B');
  assert.equal(letto.datiComuni.libri.length, 3);
});
