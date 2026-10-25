/*
 * Lettura e scrittura di archivi ZIP (i file .docx e .odt sono archivi ZIP).
 * Usa la compressione integrata nel browser (CompressionStream), senza librerie.
 */
(function (root) {
  'use strict';

  const TABELLA_CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();

  function crc32(dati) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < dati.length; i++) c = TABELLA_CRC[(c ^ dati[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  const utf8 = (s) => new TextEncoder().encode(s);

  async function trasforma(dati, flusso) {
    const s = new Blob([dati]).stream().pipeThrough(flusso);
    return new Uint8Array(await new Response(s).arrayBuffer());
  }

  async function comprimi(dati) {
    if (typeof CompressionStream === 'undefined') return null;
    try { return await trasforma(dati, new CompressionStream('deflate-raw')); } catch (e) { return null; }
  }

  async function decomprimi(dati) {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('Il browser non supporta la lettura di file compressi: aggiornalo all\'ultima versione.');
    }
    return trasforma(dati, new DecompressionStream('deflate-raw'));
  }

  function dataDos(d) {
    return {
      ora: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
      giorno: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()
    };
  }

  /**
   * Crea un archivio ZIP.
   * @param {Array<{nome: string, dati: Uint8Array|string}>} file
   * @returns {Promise<Uint8Array>}
   */
  async function crea(file) {
    const { ora, giorno } = dataDos(new Date());
    const voci = [];
    let dimensione = 0;
    for (const f of file) {
      const nome = utf8(f.nome);
      const dati = typeof f.dati === 'string' ? utf8(f.dati) : f.dati;
      const compressi = await comprimi(dati);
      const usaCompressione = compressi && compressi.length < dati.length;
      const contenuto = usaCompressione ? compressi : dati;
      voci.push({ nome, dati, contenuto, metodo: usaCompressione ? 8 : 0, crc: crc32(dati), offset: dimensione });
      dimensione += 30 + nome.length + contenuto.length;
    }
    const inizioDirectory = dimensione;
    for (const v of voci) dimensione += 46 + v.nome.length;
    const out = new Uint8Array(dimensione + 22);
    const dv = new DataView(out.buffer);
    let p = 0;
    const intestazione = (v, centrale) => {
      dv.setUint32(p, centrale ? 0x02014b50 : 0x04034b50, true); p += 4;
      if (centrale) { dv.setUint16(p, 20, true); p += 2; }
      dv.setUint16(p, 20, true); p += 2;              // versione necessaria
      dv.setUint16(p, 0x0800, true); p += 2;          // nomi in UTF-8
      dv.setUint16(p, v.metodo, true); p += 2;
      dv.setUint16(p, ora, true); p += 2;
      dv.setUint16(p, giorno, true); p += 2;
      dv.setUint32(p, v.crc, true); p += 4;
      dv.setUint32(p, v.contenuto.length, true); p += 4;
      dv.setUint32(p, v.dati.length, true); p += 4;
      dv.setUint16(p, v.nome.length, true); p += 2;
      dv.setUint16(p, 0, true); p += 2;               // campo extra
      if (centrale) {
        dv.setUint16(p, 0, true); p += 2;             // commento
        dv.setUint16(p, 0, true); p += 2;             // disco
        dv.setUint16(p, 0, true); p += 2;             // attributi interni
        dv.setUint32(p, 0, true); p += 4;             // attributi esterni
        dv.setUint32(p, v.offset, true); p += 4;
      }
      out.set(v.nome, p); p += v.nome.length;
    };
    for (const v of voci) { intestazione(v, false); out.set(v.contenuto, p); p += v.contenuto.length; }
    for (const v of voci) intestazione(v, true);
    dv.setUint32(p, 0x06054b50, true); p += 4;
    dv.setUint16(p, 0, true); p += 2;
    dv.setUint16(p, 0, true); p += 2;
    dv.setUint16(p, voci.length, true); p += 2;
    dv.setUint16(p, voci.length, true); p += 2;
    dv.setUint32(p, dimensione - inizioDirectory, true); p += 4;
    dv.setUint32(p, inizioDirectory, true); p += 4;
    dv.setUint16(p, 0, true);
    return out;
  }

  /**
   * Apre un archivio ZIP. I file vengono decompressi solo quando richiesti.
   * @param {Uint8Array} dati
   */
  function apri(dati) {
    const dv = new DataView(dati.buffer, dati.byteOffset, dati.byteLength);
    let fine = -1;
    for (let i = dati.length - 22; i >= Math.max(0, dati.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { fine = i; break; }
    }
    if (fine < 0) throw new Error('Il file non è un documento valido (archivio non riconosciuto).');
    const totale = dv.getUint16(fine + 10, true);
    let p = dv.getUint32(fine + 16, true);
    const voci = new Map();
    const dec = new TextDecoder();
    for (let i = 0; i < totale; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('Archivio danneggiato.');
      const metodo = dv.getUint16(p + 10, true);
      const compressa = dv.getUint32(p + 20, true);
      const lNome = dv.getUint16(p + 28, true);
      const lExtra = dv.getUint16(p + 30, true);
      const lCommento = dv.getUint16(p + 32, true);
      const offset = dv.getUint32(p + 42, true);
      const nome = dec.decode(dati.subarray(p + 46, p + 46 + lNome));
      voci.set(nome, { metodo, compressa, offset });
      p += 46 + lNome + lExtra + lCommento;
    }
    async function file(nome) {
      const v = voci.get(nome);
      if (!v) return null;
      const inizio = v.offset + 30 + dv.getUint16(v.offset + 26, true) + dv.getUint16(v.offset + 28, true);
      const contenuto = dati.subarray(inizio, inizio + v.compressa);
      if (v.metodo === 0) return contenuto;
      if (v.metodo === 8) return decomprimi(contenuto);
      throw new Error('Formato di compressione non supportato.');
    }
    return {
      nomi: [...voci.keys()],
      file,
      testo: async (nome) => { const b = await file(nome); return b ? dec.decode(b) : null; }
    };
  }

  const api = { crea, apri, crc32 };
  root.PDL = root.PDL || {};
  root.PDL.zip = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
