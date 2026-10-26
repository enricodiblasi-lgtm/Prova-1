# Controllo anagrafica Snoots

Controllo periodico (ogni 15 giorni) della qualità dell'anagrafica clienti e animali nel gestionale Snoots:
doppioni, dati mancanti, dati errati.

## Procedura (per Claude, ogni 15 giorni)

1. **Scarica i dati da Snoots** (connettore Snoots MCP, sola lettura):
   - `search-customers` con `s=1000`, pagine 1..N finché `hasNextPage` è falso;
   - `search-patients` con `s=1000`, pagine 1..N;
   - `search-invoices` con `emitted="true"` (restituisce tutte le fatture emesse in un'unica
     risposta): serve a distinguere i clienti dormienti da quelli attivi senza animale.

   I risultati sono troppo grandi per la conversazione e vengono salvati in file
   `mcp-Snoots_MCP-search-customers-*.txt`, `...-search-patients-*.txt`, `...-search-invoices-*.txt`.
   Copiali in una cartella **fuori dal repository** (scratchpad), solo quelli del
   mese in corso, e verifica che gli `_id` unici siano pari a `totalDocs`.
2. **Recupera lo storico**: scarica dalla dashboard il file `storico.json` attuale
   e mettilo come `storico` in `output/dashboard_data.json`, così il nuovo controllo
   si aggiunge ai precedenti invece di ripartire da zero.
3. **Lancia il controllo**:
   ```
   python3 anagrafica/controllo_anagrafica.py --input <cartella_json> --out anagrafica/output
   ```
   Produce `Controllo_Anagrafica_<data>.xlsx` e i file per la dashboard in
   `output/dashboard/` (senza CF, telefoni ed email).
4. **Aggiorna la dashboard** "Qualità anagrafica Snoots"
   (https://claude.ai/artifact/CV8aTA7Xg26R41AgbYBnJc): carica ogni file di
   `output/dashboard/` come asset e aggiorna `url`, `name` e `updated` del dataset
   con lo stesso nome.
   Aggiungi anche `dormienti.json` (dataset `dormienti`). Elimina gli asset sostituiti.
5. Invia l'Excel all'utente e riassumi le variazioni rispetto al controllo precedente.

`output/` è in `.gitignore`: contiene dati personali e non va mai committato.

## Cosa controlla

**Clienti**
- Doppioni: stesso CF o P.IVA; CF quasi uguale (errore di battitura); stesso nome con
  stesso telefono, email o indirizzo; stesso nome e cognome; nome quasi identico; nome e
  cognome invertiti. Due schede con CF validi e diversi non sono mai doppioni (omonimi o
  familiari) e finiscono nel foglio "Contatti condivisi".
- Dati mancanti: nome, cognome, CF, telefono, email, indirizzo, città, CAP, provincia, sesso;
  per aziende ragione sociale e P.IVA.
- Dati errati: CF (formato, carattere di controllo, coerenza con nome e cognome e con il
  sesso), P.IVA, telefono, email, CAP, provincia, tipo cliente.
- Clienti senza animali, divisi in:
  - **dormienti da richiamare**: clienti storici (dal vecchio programma di fatturazione, 2023)
    senza animali e senza fatture in Snoots. Non sono un errore: vanno nel foglio
    "Clienti dormienti" e nella scheda "Da richiamare" della dashboard;
  - **attivi senza animale**: hanno fatture o sono stati inseriti in clinica, ma manca la scheda
    dell'animale (dato mancante).

**Animali**
- Doppioni: stesso microchip; stesso nome sotto lo stesso proprietario o sotto suoi doppioni;
  stesso nome e data di nascita con proprietari diversi.
- Dati mancanti o errati: nome (vuoto o generico), data di nascita (mancante, futura, età non
  plausibile), sesso, microchip (obbligatorio per i cani, 15 cifre), sterilizzazione, razza,
  data di decesso, proprietario inesistente.

**Problemi sistematici**: valori predefiniti messi dall'importazione (sesso, sterilizzazione),
razza solo nelle note, indirizzi con virgola finale, consenso marketing impostato a sì per tutti.
