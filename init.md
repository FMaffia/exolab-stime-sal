Devo progettare un'applicazione web che sostituisca il file Excel che uso per stimare progetti di sviluppo software. Il file è nel repository: `docs/Template_stima_attività.xlsx`.

Il tuo compito ora NON è scrivere codice, ma produrre il documento di analisi e specifiche funzionali `docs/SPECIFICHE.md`, che verrà pushato e userò come riferimento per lo sviluppo.

## Come procedere

1. **Analizza l'Excel**: leggi tutti i fogli, le celle e le formule (non solo i valori calcolati). Ricava la struttura della stima, le colonne, i calcoli e i parametri di configurazione. Segnalami eventuali incongruenze nelle formule (valori cablati invece che presi da CONFIG, riferimenti che escludono righe, ecc.).
2. **Fammi le domande** su tutto ciò che è ambiguo o non specificato qui sotto, raggruppate e numerate, PRIMA di scrivere il documento. Non fare assunzioni silenziose: se proponi una soluzione, scrivila come proposta da confermare.
3. **Scrivi `docs/SPECIFICHE.md`** solo dopo le mie risposte.

## Requisiti

### Struttura
- N **aree**; ogni area contiene N **progetti**; ogni progetto contiene N **stime**.
- Ogni progetto ha i propri **referenti**.

### Profili e autenticazione
- Login tramite **Microsoft Entra ID**.
- L'applicazione è multi-profilo:
  - **Admin**: vede tutte le stime di tutti i progetti, crea le aree, imposta i parametri globali. **Non** crea i progetti.
  - **Writer**: crea i progetti (associandoli ad aree già esistenti) e le stime dei progetti, ne modifica le configurazioni, associa utenti al progetto.

### Parametri configurabili
- Nell'Excel (foglio CONFIG) ci sono parametri **globali** usati nei calcoli.
- I parametri globali sono **sovrascrivibili per singolo progetto**. Una volta sovrascritti, per quel progetto fanno fede i valori sovrascritti.
- Oltre al valore, per progetto si possono sovrascrivere anche la **label** e la **visibilità** del campo.

### Stati della stima
- In attesa di requisiti
- In fase di accettazione
- Accettato

Una volta accettata, sarà possibile fare il **SAL in percentuale delle singole voci**. Le regole del SAL sono **fuori scope** per ora: nel documento indicalo solo come fase successiva. Concentriamoci sulle stime.

### Export
- Ogni stima deve essere esportabile in **Excel in 2 modalità**: con e senza colonna economica.

## Contenuto atteso del documento
- Obiettivo e glossario
- Modello gerarchico Area → Progetto → Stima → voci
- Profili, matrice dei permessi (azione × profilo)
- Funzionalità per area (aree, progetti, referenti/utenti, stime, parametri, export)
- Regole di calcolo ricavate dall'Excel, con formule esplicite e parametri coinvolti
- Regola di risoluzione parametri globali / override di progetto (valore, label, visibilità)
- Ciclo di vita della stima con transizioni consentite
- Modello dati logico
- Requisiti non funzionali (sicurezza lato API, audit, precisione dei calcoli)
- Fuori scope
- Criteri di accettazione
- Domande aperte residue

Scrivi in italiano, in Markdown, in modo che il documento sia usabile come input per lo sviluppo da parte di Claude Code.