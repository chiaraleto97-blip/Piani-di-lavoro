/*
 * Riquadro per firmare con il dito, la penna o il mouse.
 * apriFirma() mostra una finestra e restituisce l'immagine PNG della firma
 * (ritagliata attorno al tratto) oppure null se si annulla.
 */
(function (root) {
  'use strict';

  const INCHIOSTRO = '#0b1f44';

  function apriFirma({ titolo = 'Firma', nome = '' } = {}) {
    return new Promise((risolvi) => {
      const dlg = document.createElement('dialog');
      dlg.className = 'dialogo dialogo-firma';
      dlg.setAttribute('aria-labelledby', 'titolo-firma');
      dlg.innerHTML =
        '<h2 id="titolo-firma"></h2>' +
        '<p class="nota">Firma nel riquadro con il dito, la penna o il mouse.</p>' +
        '<div class="area-firma"><canvas aria-label="Area per la firma"></canvas>' +
        '<div class="linea-firma" aria-hidden="true"></div><span class="nome-firma"></span></div>' +
        '<div class="azioni-dialogo">' +
        '<button type="button" class="btn" data-azione="cancella">Cancella</button>' +
        '<span class="spazio"></span>' +
        '<button type="button" class="btn" data-azione="annulla">Annulla</button>' +
        '<button type="button" class="btn primario" data-azione="conferma" disabled>Conferma firma</button>' +
        '</div>';
      dlg.querySelector('#titolo-firma').textContent = titolo;
      dlg.querySelector('.nome-firma').textContent = nome;
      document.body.append(dlg);

      const canvas = dlg.querySelector('canvas');
      const ctx = canvas.getContext('2d');
      const conferma = dlg.querySelector('[data-azione="conferma"]');
      const tratti = [];
      let tratto = null;
      let esito = null;

      function dimensiona() {
        const r = canvas.getBoundingClientRect();
        const dpr = window.devicePixelRatio || 1;
        canvas.width = Math.round(r.width * dpr);
        canvas.height = Math.round(r.height * dpr);
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ridisegna();
      }

      function disegnaTratto(c, t) {
        c.strokeStyle = INCHIOSTRO;
        c.fillStyle = INCHIOSTRO;
        c.lineCap = 'round';
        c.lineJoin = 'round';
        const p = t.punti;
        if (p.length === 1) {
          c.beginPath();
          c.arc(p[0].x, p[0].y, p[0].w / 2, 0, Math.PI * 2);
          c.fill();
          return;
        }
        for (let i = 1; i < p.length; i++) {
          // Curve morbide passando per i punti medi.
          const a = p[i - 1];
          const b = p[i];
          const m0 = i > 1 ? { x: (p[i - 2].x + a.x) / 2, y: (p[i - 2].y + a.y) / 2 } : a;
          const m1 = { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
          c.lineWidth = b.w;
          c.beginPath();
          c.moveTo(m0.x, m0.y);
          c.quadraticCurveTo(a.x, a.y, m1.x, m1.y);
          c.stroke();
        }
      }

      function ridisegna() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        tratti.forEach((t) => disegnaTratto(ctx, t));
      }

      function punto(e) {
        const r = canvas.getBoundingClientRect();
        // La pressione (penna) rende il tratto più naturale; dito e mouse hanno spessore fisso.
        const pressione = e.pointerType === 'pen' && e.pressure > 0 ? e.pressure : 0.5;
        return { x: e.clientX - r.left, y: e.clientY - r.top, w: 1.4 + pressione * 2.6 };
      }

      canvas.addEventListener('pointerdown', (e) => {
        e.preventDefault();
        // Su alcuni dispositivi la cattura non è disponibile: si firma lo stesso.
        try { canvas.setPointerCapture(e.pointerId); } catch (err) { /* nessun problema */ }
        tratto = { punti: [punto(e)] };
        tratti.push(tratto);
        ridisegna();
      });
      canvas.addEventListener('pointermove', (e) => {
        if (!tratto) return;
        e.preventDefault();
        const eventi = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
        (eventi.length ? eventi : [e]).forEach((ev) => tratto.punti.push(punto(ev)));
        ridisegna();
      });
      const fine = () => {
        if (!tratto) return;
        tratto = null;
        conferma.disabled = !tratti.some((t) => t.punti.length > 2);
      };
      canvas.addEventListener('pointerup', fine);
      canvas.addEventListener('pointercancel', fine);

      // Immagine ritagliata attorno alla firma, con un piccolo margine.
      function esporta() {
        const tutti = tratti.flatMap((t) => t.punti);
        const margine = 8;
        const x0 = Math.max(0, Math.min(...tutti.map((p) => p.x)) - margine);
        const y0 = Math.max(0, Math.min(...tutti.map((p) => p.y)) - margine);
        const x1 = Math.max(...tutti.map((p) => p.x)) + margine;
        const y1 = Math.max(...tutti.map((p) => p.y)) + margine;
        const scala = Math.min(2, 600 / (x1 - x0));
        const out = document.createElement('canvas');
        out.width = Math.max(1, Math.round((x1 - x0) * scala));
        out.height = Math.max(1, Math.round((y1 - y0) * scala));
        const c = out.getContext('2d');
        c.setTransform(scala, 0, 0, scala, -x0 * scala, -y0 * scala);
        tratti.forEach((t) => disegnaTratto(c, t));
        return out.toDataURL('image/png');
      }

      dlg.addEventListener('click', (e) => {
        const azione = e.target.closest('[data-azione]');
        if (!azione) return;
        if (azione.dataset.azione === 'cancella') { tratti.length = 0; ridisegna(); conferma.disabled = true; }
        if (azione.dataset.azione === 'annulla') dlg.close();
        if (azione.dataset.azione === 'conferma') { esito = esporta(); dlg.close(); }
      });
      const ridimensiona = () => dimensiona();
      window.addEventListener('resize', ridimensiona);
      dlg.addEventListener('close', () => {
        window.removeEventListener('resize', ridimensiona);
        dlg.remove();
        risolvi(esito);
      });
      dlg.showModal();
      dimensiona();
    });
  }

  root.PDL = root.PDL || {};
  root.PDL.firma = { apriFirma };
})(window);
