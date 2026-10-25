# Piani di lavoro

App web per compilare in modo **rapido e uniforme** i piani di lavoro annuali dei docenti
della scuola secondaria di primo grado.

- **Un solo modello per tutta la scuola**: il referente configura intestazione, sezioni,
  voci degli elenchi e discipline, poi distribuisce il modello ai colleghi (anche in sola lettura).
- **Compilazione a colpi di spunta**: competenze chiave europee, traguardi delle Indicazioni
  nazionali per ogni disciplina, metodologie, strumenti, verifiche, criteri di valutazione,
  recupero e potenziamento, inclusione, educazione civica, rapporti con le famiglie.
- **Unità di apprendimento** con periodo, ore, nuclei tematici della disciplina, obiettivi e contenuti.
- **Scorciatoie**: più classi in un colpo solo («A, B, C»), duplicazione, copia dal piano
  dell'anno precedente o di un collega, copia delle UdA da un altro piano.
- **Dati comuni delle classi da un foglio Google**: plesso, libri di testo e numero di alunni
  (anche con PEI e PDP) si gestiscono una volta sola in un foglio condiviso e l'app li inserisce in
  sola lettura in tutti i piani della classe; se il foglio viene corretto, i piani si aggiornano.
- **Indice con avanzamento**: si vede subito quali sezioni mancano.
- **Firma con il dito** (o con penna e mouse): compare in fondo al documento con luogo e data;
  l'app segnala se il piano viene modificato dopo la firma.
- **Importazione di un piano già scritto** (.docx, .odt, .txt): l'app riconosce docente,
  disciplina e classe, divide il documento nelle sue parti e le riporta nelle sezioni del modello
  d'istituto (spuntando le voci riconosciute, i traguardi, i numeri della classe e la tabella delle
  UdA). Il docente controlla gli abbinamenti prima di creare il piano; le parti non previste dal
  modello diventano sezioni aggiuntive, così nessun testo va perso.
- **Uscite**: anteprima, stampa/PDF e documento Word (.docx) con impaginazione identica per tutti;
  stampa di più piani insieme (es. per il consiglio di classe).
- **Nessun server, nessun account**: i dati restano nel browser del docente; esportazione e
  importazione in file `.json` per backup, condivisione e cambio di computer.

## Come si usa

Non serve installare nulla: basta aprire `index.html` con un browser aggiornato
(Chrome, Edge, Firefox, Safari). In alternativa si può pubblicare la cartella su un
qualsiasi hosting statico, ad esempio GitHub Pages (Settings → Pages → Deploy from a branch).

1. **Referente**: *Modello d'istituto* → nome della scuola, sezioni, voci e discipline →
   *Esporta il modello* → invio del file `modello-istituto.json` ai docenti.
2. **Docenti**: *Importa file* → scelta del modello ricevuto → *+ Nuovo piano*.
3. Oppure, per chi ha già un piano: *Importa un piano già scritto* → revisione → *Crea il piano uniformato*.
4. *Firma con il dito*, poi *Anteprima e stampa* per il PDF o *Scarica Word* per il file modificabile.

La firma con il dito è una firma grafica (firma elettronica semplice), adatta ai documenti
interni; non sostituisce la firma digitale quando questa è richiesta.

Se l'app è pubblicata online, basta salvare `modello-istituto.json` accanto a `index.html`:
viene caricato automaticamente al primo avvio e, quando viene aggiornato, ai docenti compare
un avviso per applicare la nuova versione.

## Dati comuni delle classi (foglio Google)

1. In *Modello d'istituto → Dati comuni delle classi* scarica il modello del foglio (.xlsx) e
   caricalo su Google Drive aprendolo con Fogli Google. Ha due schede:
   - **Classi**: Classe, Sezione, Plesso, Alunni, Maschi, Femmine, Alunni con PEI,
     Alunni con PDP (DSA), Alunni con PDP (altri BES);
   - **Libri**: Classe, Sezione (vuota = tutte le sezioni), Disciplina, Libro di testo.
2. Compila il foglio (solo numeri, mai nomi di alunni) e condividilo con
   «Chiunque abbia il link → Visualizzatore».
3. Incolla il link nell'app ed esporta il modello d'istituto: ogni docente, aprendo l'app,
   legge i dati aggiornati dal foglio. Le celle vuote restano compilabili dal docente.

Se il dominio della scuola non consente la condivisione con link, scarica il foglio come
Microsoft Excel e usa *Carica il foglio da file*: i dati entrano nel modello d'istituto distribuito.
Il numero di alunni con PEI e PDP è un dato riservato: il link va condiviso solo con i docenti e
un modello con i dati delle classi non va pubblicato su un sito pubblico (all'esportazione si può
escluderli: i docenti li leggeranno dal foglio).

## Dati

I piani sono salvati nel `localStorage` del browser, quindi solo su quel computer e in quel
browser. È importante fare periodicamente *Backup di tutti i piani* dall'elenco.

## Struttura del codice

| File | Contenuto |
| --- | --- |
| `index.html`, `css/style.css` | Pagina e stili (anche di stampa) |
| `js/defaults.js` | Modello predefinito: sezioni, voci, discipline, traguardi e nuclei tematici |
| `js/model.js` | Logica dei dati (piani, completamento, import/export), senza DOM |
| `js/documento.js` | Struttura del documento e HTML per anteprima e stampa |
| `js/word.js` | Esportazione in Word (.docx) dalla stessa struttura |
| `js/zip.js` | Lettura e scrittura di archivi ZIP (formato di .docx e .odt) |
| `js/importa.js` | Lettura di piani già scritti e conversione nel modello d'istituto |
| `js/firma.js` | Riquadro per la firma con dito, penna o mouse |
| `js/fogli.js` | Lettura del foglio comune (Google Fogli, .xlsx, .csv) e modello del foglio |
| `js/store.js` | Salvataggio nel browser e caricamento del modello d'istituto |
| `js/app.js` | Interfaccia |
| `tests/campioni/` | Piani di esempio in vari formati usati dai test dell'importazione |

JavaScript senza dipendenze e senza passaggi di build; serve un browser aggiornato
(le funzioni di compressione usate per Word sono disponibili da Chrome/Edge 103,
Firefox 113 e Safari 16.4). Test (Node.js 22+):

```sh
npm test
```

## Riferimenti

Indicazioni nazionali per il curricolo (D.M. 254/2012) e Nuovi scenari (2018);
Raccomandazione UE del 22 maggio 2018 sulle competenze chiave; D.Lgs. 62/2017;
L. 104/1992, L. 170/2010 e normativa BES; L. 92/2019 e Linee guida per l'educazione civica
(D.M. 183/2024). I traguardi sono riportati dalle Indicazioni nazionali, talvolta in forma
sintetica, e possono essere adattati nel modello d'istituto.
