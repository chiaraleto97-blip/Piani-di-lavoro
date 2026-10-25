# Piani di lavoro — istruzioni per Claude Code

App web statica (HTML, CSS, JavaScript senza dipendenze né build) per compilare i piani di
lavoro dei docenti della scuola secondaria di primo grado. Interfaccia e codice in italiano.

## Due persone lavorano in parallelo: regole per i commit

Sul progetto lavorano contemporaneamente due persone, ognuna con le proprie sessioni di
Claude Code. Queste regole evitano che una sovrascriva il lavoro dell'altra.

1. **Mai commit o push su `main`.** `main` è protetto: si aggiorna solo con pull request
   e solo se i test automatici passano.
2. **Un branch per ogni attività**, creato da `main` aggiornato:
   `git fetch origin main && git checkout -b <branch> origin/main`.
   Se la sessione ha un branch assegnato, usa quello. Se la pull request di quel branch è già
   stata unita, riparti da `origin/main` aggiornato (stesso nome di branch) invece di
   aggiungere commit alla storia già unita.
3. **Commit piccoli e chiari**: un argomento per commit, messaggio in italiano che dica cosa
   cambia e perché (es. «Importazione: riconosce le tabelle verticali delle UdA»). Non mescolare
   nello stesso commit modifiche non collegate o riformattazioni di file interi.
4. **Prima di ogni push** allinea il branch a `main` e fai girare i test:
   `git fetch origin main && git merge origin/main`, poi `npm test`.
   Non si fa push con i test rossi.
5. **Apri una pull request verso `main`** a lavoro finito, seguendo il modello
   (`.github/pull_request_template.md`). Una pull request = una funzione o una correzione:
   meglio due piccole che una enorme.
6. **Conflitti**: risolvili tenendo le modifiche di entrambi. Non scartare mai il lavoro
   dell'altra persona; se le due modifiche fanno cose incompatibili, fermati e chiedi.
7. **Mai riscrivere la storia condivisa**: niente `push --force` su branch altrui o su `main`,
   niente `rebase` di commit già pubblicati da altri. Su un branch tuo, se serve, solo
   `--force-with-lease`.
8. **File grandi condivisi** (`js/app.js`, `css/style.css`): modifica solo le parti che servono,
   senza spostare o riformattare blocchi esistenti, così le pull request parallele non vanno in
   conflitto.
9. **Niente dati reali di alunni** in campioni, test o commit.

## Struttura

| File | Contenuto |
| --- | --- |
| `index.html`, `css/style.css` | Pagina e stili (anche di stampa) |
| `js/defaults.js` | Modello d'istituto predefinito: sezioni, voci, discipline, traguardi |
| `js/model.js` | Dati: piani, completamento, dati comuni delle classi, import/export JSON |
| `js/documento.js` | Struttura del documento e HTML per anteprima e stampa |
| `js/word.js` | Esportazione .docx dalla stessa struttura |
| `js/zip.js` | Lettura/scrittura ZIP (formato di .docx, .odt, .xlsx) |
| `js/importa.js` | Importazione di piani già scritti (.docx, .odt, .txt) |
| `js/fogli.js` | Foglio comune delle classi (Google Fogli, .xlsx, .csv) |
| `js/firma.js` | Firma con dito, penna o mouse |
| `js/store.js` | Salvataggio nel browser (localStorage) |
| `js/app.js` | Interfaccia (viste, dialoghi, eventi) |
| `tests/` | Test con `node:test`; `tests/campioni/` documenti di esempio |

I moduli in `js/` (tranne `app.js`, `store.js`, `firma.js`) non usano il DOM e si caricano sia
nel browser (`window.PDL`) sia in Node (`require`): la logica nuova va lì, con i suoi test.

## Comandi

- Test: `npm test` (Node.js 22+). Girano anche su GitHub a ogni pull request.
- Prova nel browser: `python3 -m http.server` nella cartella e apri `http://localhost:8000`.

## Convenzioni

- Nomi di variabili e funzioni in italiano, come nel codice esistente.
- Nessuna dipendenza esterna e nessun passaggio di build: l'app deve funzionare anche aprendo
  `index.html` dal disco.
- Tutto il testo inserito dall'utente che finisce in HTML passa da `esc()` (documento.js) o viene
  inserito come testo, mai come HTML.
