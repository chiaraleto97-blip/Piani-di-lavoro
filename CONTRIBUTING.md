# Lavorare in due sul progetto

Il codice stabile sta nel branch **`main`**, che è protetto: nessuno può scriverci
direttamente. Ogni modifica passa da una **pull request**, che entra in `main` solo quando i
test automatici sono verdi. Così due persone (e le loro sessioni di Claude Code) possono
lavorare in parallelo senza sovrascriversi.

## Il giro di una modifica

1. **Nuova sessione di Claude Code** sul repository, partendo da `main`.
   Una sessione per ogni attività (es. «aggiungi la scheda Uscite didattiche»).
2. Claude lavora su un **branch suo** e fa commit piccoli e descrittivi.
   Le regole che segue sono scritte in [`CLAUDE.md`](CLAUDE.md).
3. A lavoro finito chiedi a Claude di **aprire una pull request verso `main`**.
4. Su GitHub i **test partono da soli**: aspetta la spunta verde.
5. Premi **«Squash and merge»**: la modifica entra in `main` come un unico commit e il
   branch viene cancellato.

## Se GitHub dice che il branch non è aggiornato

Succede quando l'altra persona ha unito una sua modifica nel frattempo. Premi
**«Update branch»** nella pull request (oppure chiedi a Claude di allineare il branch a
`main`): i test ripartono sulla versione aggiornata e poi si può unire.

## Se ci sono conflitti

Vuol dire che avete modificato le stesse righe. Chiedi a Claude di risolverli tenendo le
modifiche di entrambi; se le due modifiche fanno cose incompatibili, decidete insieme quale
tenere.

## Per evitare conflitti

- **Dividetevi le attività** e ditevi su cosa state lavorando.
- **Pull request piccole e frequenti**: meglio unire spesso che accumulare.
- Prima di iniziare una nuova attività, parti sempre da `main` aggiornato.
- Nei file grandi (`js/app.js`, `css/style.css`) modificate solo le parti necessarie.

## Cosa non fare

- Non scrivere direttamente su `main` (GitHub comunque lo impedisce).
- Non usare «force push» sui branch dell'altra persona.
- Non mettere nel repository dati reali di alunni.
