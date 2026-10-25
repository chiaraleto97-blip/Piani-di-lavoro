/*
 * Esportazione in Word (.docx) a partire dai blocchi di documento.js.
 * Il file si apre allo stesso modo in Word, LibreOffice e Google Documenti
 * e può contenere l'immagine della firma.
 */
(function (root) {
  'use strict';

  const PDL = root.PDL || {};
  const documento = PDL.documento || (typeof require === 'function' ? require('./documento.js') : null);
  const zip = PDL.zip || (typeof require === 'function' ? require('./zip.js') : null);

  const LARGHEZZA_TESTO = 9638; // A4 con margini di 2 cm, in ventesimi di punto
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  function x(s) {
    return String(s == null ? '' : s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function run(testo, props = '') {
    const rpr = props ? `<w:rPr>${props}</w:rPr>` : '';
    return String(testo).split('\n').map((parte, i) =>
      `<w:r>${rpr}${i ? '<w:br/>' : ''}<w:t xml:space="preserve">${x(parte)}</w:t></w:r>`).join('');
  }

  function runs(elenco, extra = '') {
    return elenco.map((r) => run(r.testo, (r.b ? '<w:b/>' : '') + (r.i ? '<w:i/>' : '') + extra)).join('');
  }

  function paragrafo(contenuto, ppr = '') {
    return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${contenuto}</w:p>`;
  }

  const STILI_TESTATA = {
    scuola: { ppr: '<w:spacing w:after="0"/>', rpr: '<w:b/><w:caps/><w:sz w:val="28"/>' },
    sottotitolo: { ppr: '<w:spacing w:after="0"/>', rpr: '<w:sz w:val="22"/>' },
    titolo: { ppr: '<w:spacing w:before="240" w:after="40"/>', rpr: '<w:b/><w:caps/><w:spacing w:val="20"/><w:sz w:val="36"/>' },
    anno: { ppr: '<w:spacing w:after="280"/>', rpr: '' }
  };

  function bordi(colore = '666666') {
    return ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map((l) => `<w:${l} w:val="single" w:sz="4" w:space="0" w:color="${colore}"/>`).join('');
  }

  function cella(contenuto, larghezza, { sfondo = false, grassetto = false, dimensione = 22 } = {}) {
    const extra = (grassetto ? '<w:b/>' : '') + `<w:sz w:val="${dimensione}"/>`;
    const corpo = Array.isArray(contenuto) ? runs(contenuto, `<w:sz w:val="${dimensione}"/>`) : run(contenuto, extra);
    return `<w:tc><w:tcPr><w:tcW w:w="${larghezza}" w:type="dxa"/>` +
      (sfondo ? '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>' : '') +
      `</w:tcPr>${paragrafo(corpo, '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>')}</w:tc>`;
  }

  function tabella(b) {
    const colonne = b.intestazione ? b.intestazione.length : Math.max(...b.righe.map((r) => r.length));
    const perc = b.larghezze || Array(colonne).fill(100 / colonne);
    const larg = perc.map((p) => Math.round(LARGHEZZA_TESTO * p / 100));
    const dimensione = b.classe === 'uda' ? 20 : 22;
    let xml = '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>' +
      `<w:tblBorders>${bordi()}</w:tblBorders><w:tblLayout w:type="fixed"/>` +
      '<w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>' +
      `</w:tblPr><w:tblGrid>${larg.map((l) => `<w:gridCol w:w="${l}"/>`).join('')}</w:tblGrid>`;
    if (b.intestazione) {
      xml += '<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>' +
        b.intestazione.map((t, i) => cella(t, larg[i], { sfondo: true, grassetto: true, dimensione })).join('') + '</w:tr>';
    }
    b.righe.forEach((r) => {
      xml += '<w:tr>' + r.map((c, i) => {
        const colonnaIntestazione = b.intestazioneColonna && i === 0;
        return cella(c, larg[i], { sfondo: colonnaIntestazione, grassetto: colonnaIntestazione && b.classe !== 'campi', dimensione });
      }).join('') + '</w:tr>';
    });
    return xml + '</w:tbl>';
  }

  function base64InByte(dataUrl) {
    const bin = atob(dataUrl.split(',')[1]);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  function dimensioniPng(dati) {
    const dv = new DataView(dati.buffer, dati.byteOffset, dati.byteLength);
    return { w: dv.getUint32(16), h: dv.getUint32(20) };
  }

  function immagine(rid, id, dati) {
    const { w, h } = dimensioniPng(dati);
    const emuPx = 9525;
    const scala = Math.min(1, 1800000 / (w * emuPx), 680000 / (h * emuPx));
    const cx = Math.round(w * emuPx * scala);
    const cy = Math.round(h * emuPx * scala);
    return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
      `<wp:docPr id="${id}" name="Firma ${id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="firma${id}.png"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
      '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  }

  function firma(b, contesto) {
    const nessunBordo = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((l) => `<w:${l} w:val="nil"/>`).join('');
    const meta = Math.round(LARGHEZZA_TESTO / 2);
    const luogo = (b.luogo ? b.luogo + ', ' : 'Luogo e data: ') + (b.data || '____________________');
    let destra = paragrafo(run('Il/La docente'), '<w:jc w:val="center"/><w:spacing w:after="0"/>');
    if (b.immagine) {
      const n = contesto.immagini.length + 1;
      const rid = 'rIdFirma' + n;
      const dati = base64InByte(b.immagine);
      contesto.immagini.push({ rid, nome: `media/firma${n}.png`, dati });
      destra += paragrafo(immagine(rid, n, dati), '<w:jc w:val="center"/><w:spacing w:before="60" w:after="0"/>');
    } else {
      destra += paragrafo('', '<w:spacing w:before="480" w:after="0"/>');
    }
    destra += paragrafo(run(b.docente), '<w:pBdr><w:top w:val="single" w:sz="6" w:space="1" w:color="111111"/></w:pBdr><w:jc w:val="center"/><w:ind w:left="900" w:right="900"/>');
    return paragrafo('', '<w:spacing w:before="360" w:after="0"/>') +
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${nessunBordo}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="${meta}"/><w:gridCol w:w="${LARGHEZZA_TESTO - meta}"/></w:tblGrid>` +
      '<w:tr><w:trPr><w:cantSplit/></w:trPr>' +
      `<w:tc><w:tcPr><w:tcW w:w="${meta}" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>${paragrafo(run(luogo))}</w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="${LARGHEZZA_TESTO - meta}" w:type="dxa"/></w:tcPr>${destra}</w:tc>` +
      '</w:tr></w:tbl>';
  }

  function blocco(b, contesto) {
    switch (b.t) {
      case 'testata':
        return b.righe.map((r) => {
          const s = STILI_TESTATA[r.stile] || STILI_TESTATA.anno;
          return paragrafo(run(r.testo, s.rpr), '<w:jc w:val="center"/>' + s.ppr);
        }).join('');
      case 'h2': return paragrafo(run(b.testo), '<w:pStyle w:val="Heading2"/>');
      case 'p': {
        const stile = b.classe === 'sotto' ? '<w:pStyle w:val="Sottotitolo"/>' : b.classe === 'vuoto' ? '<w:pStyle w:val="NonCompilato"/>' : '';
        return paragrafo(runs(b.runs), stile);
      }
      case 'lista': {
        let numId = 1;
        if (b.numerata) { contesto.liste.push(contesto.liste.length + 2); numId = contesto.liste[contesto.liste.length - 1]; }
        return b.voci.map((v) => paragrafo(run(v),
          `<w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`)).join('');
      }
      case 'tabella': return tabella(b) + paragrafo('', '<w:spacing w:after="0" w:line="120" w:lineRule="auto"/>');
      case 'firma': return firma(b, contesto);
      default: return '';
    }
  }

  const STILI = XML + `<w:styles ${NS}>` +
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/>' +
    '<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="it-IT"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:keepLines/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="444444"/></w:pBdr>' +
    '<w:spacing w:before="320" w:after="120"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="Sottotitolo"><w:name w:val="Sottotitolo sezione"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:keepNext/><w:spacing w:before="160" w:after="40"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="NonCompilato"><w:name w:val="Sezione non compilata"/><w:basedOn w:val="Normal"/>' +
    '<w:rPr><w:i/><w:color w:val="777777"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/>' +
    '<w:pPr><w:spacing w:after="20"/><w:ind w:left="720"/></w:pPr></w:style>' +
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>' +
    '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
    '</w:styles>';

  function numerazione(liste) {
    const livello = (formato, testo) => '<w:lvl w:ilvl="0"><w:start w:val="1"/>' +
      `<w:numFmt w:val="${formato}"/><w:lvlText w:val="${testo}"/><w:lvlJc w:val="left"/>` +
      '<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>';
    return XML + `<w:numbering ${NS}>` +
      `<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>${livello('bullet', '•')}</w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/>${livello('decimal', '%1.')}</w:abstractNum>` +
      '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
      liste.map((id) => `<w:num w:numId="${id}"><w:abstractNumId w:val="1"/>` +
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>').join('') +
      '</w:numbering>';
  }

  /**
   * Crea il file .docx con uno o più piani (uno per pagina).
   * @returns {Promise<Uint8Array>}
   */
  async function creaDocx(piani, modello) {
    const contesto = { liste: [], immagini: [] };
    const corpo = piani.map((p, i) =>
      (i ? paragrafo('<w:r><w:br w:type="page"/></w:r>') : '') +
      documento.strutturaPiano(p, modello).map((b) => blocco(b, contesto)).join('')).join('');
    const doc = XML + `<w:document ${NS}><w:body>${corpo}<w:p/>` +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
      '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>' +
      '</w:body></w:document>';
    const relazioni = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rIdStili" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '<Relationship Id="rIdNumeri" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>' +
      contesto.immagini.map((im) => `<Relationship Id="${im.rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${im.nome}"/>`).join('') +
      '</Relationships>';
    const tipi = XML + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Default Extension="png" ContentType="image/png"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>' +
      '</Types>';
    const radice = XML + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>';
    return zip.crea([
      { nome: '[Content_Types].xml', dati: tipi },
      { nome: '_rels/.rels', dati: radice },
      { nome: 'word/document.xml', dati: doc },
      { nome: 'word/styles.xml', dati: STILI },
      { nome: 'word/numbering.xml', dati: numerazione(contesto.liste) },
      { nome: 'word/_rels/document.xml.rels', dati: relazioni },
      ...contesto.immagini.map((im) => ({ nome: 'word/' + im.nome, dati: im.dati }))
    ]);
  }

  const api = { creaDocx, MIME: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' };
  root.PDL = root.PDL || {};
  root.PDL.word = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
