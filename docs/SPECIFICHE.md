# Stime SAL — Analisi e specifiche funzionali

> Versione 1.0 — 2026-10-16
> Fonte di partenza: `docs/Template stima attività.xlsx` (fogli *Stima*, *Dettaglio intervento*, *CONFIG*).
> Documento di riferimento per lo sviluppo. Le voci marcate **[PROPOSTA]** sono decisioni proposte e non ancora confermate: sono riepilogate in §15.

---

## Indice

1. [Obiettivo](#1-obiettivo)
2. [Glossario](#2-glossario)
3. [Analisi del file Excel di partenza](#3-analisi-del-file-excel-di-partenza)
4. [Modello gerarchico](#4-modello-gerarchico)
5. [Profili e autorizzazioni](#5-profili-e-autorizzazioni)
6. [Funzionalità](#6-funzionalità)
7. [Regole di calcolo](#7-regole-di-calcolo)
8. [Parametri: risoluzione globale / progetto / stima](#8-parametri-risoluzione-globale--progetto--stima)
9. [Ciclo di vita della stima](#9-ciclo-di-vita-della-stima)
10. [Export Excel](#10-export-excel)
11. [Modello dati logico](#11-modello-dati-logico)
12. [Requisiti non funzionali](#12-requisiti-non-funzionali)
13. [Fuori scope](#13-fuori-scope)
14. [Criteri di accettazione](#14-criteri-di-accettazione)
15. [Domande aperte residue](#15-domande-aperte-residue)

---

## 1. Obiettivo

Sostituire il file Excel usato per stimare progetti di sviluppo software con un'applicazione web multi-utente che:

- organizza le stime in una gerarchia **Area → Progetto → Stima → Macro-attività → Voci**;
- applica in modo uniforme e verificabile le regole di calcolo oggi presenti nell'Excel (contingency, test, PM, risk budget, valorizzazione economica);
- centralizza i **parametri globali** e ne consente l'override per progetto e per singola stima;
- gestisce il **ciclo di vita** della stima fino all'accettazione, predisponendo la base dati per il **SAL** (fase successiva);
- esporta ogni stima in Excel, con o senza informazioni economiche e con o senza formule;
- autentica gli utenti con **Microsoft Entra ID** e traccia ogni modifica (audit campo per campo).

## 2. Glossario

| Termine | Definizione |
|---|---|
| **Area** | Raggruppamento organizzativo di progetti (es. business unit, cliente, linea di servizio). Creata dall'Admin. |
| **Progetto** | Iniziativa appartenente a un'area; contiene stime, referenti e utenti associati. Creato da un Writer. |
| **Stima** | Valutazione dell'effort di un progetto (o di una sua parte, es. una CR), con stato, versione e parametri. |
| **Macro-attività** | Riga del riepilogo della stima (es. *Setup ambienti*, *Analisi*, *Attività di sviluppo*, *Project Management*, *Risk budget*). |
| **Voce** | Riga di dettaglio di una macro-attività (es. *Attività 1*). |
| **gg/uomo (gg)** | Unità di misura dell'effort: giornate/uomo. |
| **Contingency** | Maggiorazione percentuale applicata alle voci di sviluppo per coprire incertezze tecniche. |
| **Stima netta** | Stima di una voce di sviluppo *senza* contingency (colonna "no contingency" dell'Excel). |
| **Test funzionali** | Effort di test calcolato in percentuale sul totale sviluppi. |
| **PM** | Project Management: percentuale sulla base di calcolo. |
| **Risk budget** | Riserva percentuale calcolata sulla base comprensiva di PM. |
| **Tariffa** | Importo in € per giornata/uomo, unico per stima. |
| **Parametro** | Uno dei 5 valori di configurazione (§7.1) con valore, label e visibilità. |
| **Override** | Ridefinizione di valore, label o visibilità di un parametro a livello di progetto o di stima. |
| **Forzatura** | Sostituzione manuale del valore calcolato di una riga (Test, PM, Risk) con un valore inserito dall'utente. |
| **Congelamento** | Copia dei parametri risolti nella stima al momento dell'accettazione, per renderla indipendente da modifiche successive ai globali o al progetto. |
| **CR** | Change Request: modifica di perimetro aggiunta a una stima già accettata o in una nuova stima. |
| **SAL** | Stato Avanzamento Lavori: percentuale di completamento per voce. **Fase successiva**, fuori scope. |
| **App Role** | Ruolo applicativo definito sulla registrazione dell'app in Entra ID e incluso nel token (`roles`). |
| **Membro di progetto** | Utente associato a un progetto con ruolo di progetto *Writer* o *Viewer*. |
| **Referente** | Contatto anagrafico di un progetto (interno o del cliente); non è necessariamente un utente dell'applicazione. |

## 3. Analisi del file Excel di partenza

### 3.1 Struttura

**CONFIG** — parametri globali

| Cella | Label originale | Valore | Formato |
|---|---|---|---|
| C2 | % testing | 0,10 | % |
| C3 | % contincency *(refuso)* | 0,30 | % |
| C4 | % PM | 0,15 | % |
| C5 | % Risk Budget | 0,20 | % |
| C6 | Tariffa | 442 | € |

**Dettaglio intervento** — voci di sviluppo. Colonne: `#`, Attività, Descrizione, Stima (gg/uomo) no contingency, Stima (gg/uomo), Note / punti aperti.

| Riga | Formula originale |
|---|---|
| Voce *i* (righe 3–8) | `E_i = ROUND(D_i + D_i * CONFIG!$C$3, 0)` |
| Totale sviluppi (9) | `D9 = SUM(D3:D8)` · `E9 = SUM(E3:E8)` |
| Test funzionali (10) | `D10 = ROUND(D9 * 0.1, 0)` · `E10 = ROUND(E9 * CONFIG!$C$2, 0)` |
| TOTALE (11) | `D11 = SUM(D9:D10)` · `E11 = SUM(E9:E10)` |

**Stima** — riepilogo. Colonne: `#`, Attività, Descrizione, Stima (gg/uomo), Note / punti aperti.

| Riga | Formula originale | Valore |
|---|---|---|
| Setup ambienti | input | 3 |
| Analisi | input | 10 |
| Attività di sviluppo | `='Dettaglio intervento'!E9` | 61 |
| Project Management | `=ROUND(SUM(D3:D5) * CONFIG!$C$4, 0)` | 11 |
| Risk budget | `=ROUND(SUM(D3:D6) * CONFIG!C5, 0)` | 17 |
| Totale | `=SUM(D3:D7)` | 102 |
| Totale economico | `=D8 * CONFIG!C6` | 45.084,00 € |

### 3.2 Incongruenze rilevate e decisioni

| # | Incongruenza | Decisione |
|---|---|---|
| I1 | `Stima!D5` referenzia `E9` (Totale sviluppi) invece di `E11` (TOTALE con test): i test funzionali non entrano nel totale. | **Errore.** Nell'app la riga *Attività di sviluppo* = Totale sviluppi + Test. Con i dati del template il totale passa da 102 a **110 gg** (48.620,00 €). |
| I2 | `Dettaglio!D10` usa `0.1` cablato invece di `CONFIG!$C$2`. | **Errore.** Anche il test netto usa il parametro *% testing*. |
| I3 | `Stima!D7` usa riferimento relativo `CONFIG!C5`. | Irrilevante nell'app (nessun riferimento a celle). Nell'export con formule si usano riferimenti assoluti/nomi. |
| I4 | Arrotondamento all'intero a ogni passaggio (voce, test, PM, risk). | **Confermato come Excel** (§7.3). |
| I5 | Il Risk budget è calcolato su una base che include il PM. | **Confermato.** |
| I6 | Contingency applicata solo alle voci di sviluppo; Setup e Analisi ne sono esenti ma entrano nella base di PM e Risk. | **Confermato.** |
| I7 | Colonna "no contingency" (D9:D11) non usata da altre formule. | Mantenuta come informazione (sola lettura sui totali). |
| I8 | Refusi ("contincecy", "contincency"), spazi finali ("Totale ", "TOTALE "), titolo duplicato, colonna `#` vuota nel riepilogo. | Corretti nelle label di default; `#` numerato automaticamente. |

## 4. Modello gerarchico

```
Area (1) ──< Progetto (N)
               ├──< Membro di progetto (N)   utente + ruolo WRITER | VIEWER
               ├──< Referente (N)            anagrafica, non necessariamente utente
               ├──< Override parametri progetto (0..5)
               └──< Stima (N)
                      ├──< Override parametri stima (0..5)
                      ├──< Forzature (0..3)    TEST | PM | RISK
                      ├──< Transizioni di stato (N)
                      └──< Macro-attività (N)
                             └──< Voce (N)
```

### 4.1 Macro-attività: tipi

| Tipo | Cardinalità per stima | Valore | Dettaglio a voci | Modificabile |
|---|---|---|---|---|
| `SVILUPPO` | esattamente 1 | calcolato dalle voci (con contingency) + Test | obbligatorio | nome e descrizione; non eliminabile |
| `MANUALE` | 0..N | input diretto **oppure** somma delle proprie voci (senza contingency, senza test) | facoltativo | aggiungibile, rinominabile, eliminabile, riordinabile |
| `PM` | esattamente 1 | calcolato (§7) o forzato | no | nome via label del parametro; non eliminabile |
| `RISK` | esattamente 1 | calcolato (§7) o forzato | no | nome via label del parametro; non eliminabile |

- Una nuova stima nasce con le macro-attività del template: *Setup ambienti* (MANUALE, descr. "Installazione progetti e startup"), *Analisi* (MANUALE, descr. "Analisi degli interventi"), *Attività di sviluppo* (SVILUPPO), *Project Management* (PM), *Risk budget* (RISK).
- L'utente può aggiungere macro-attività MANUALE nuove (es. "Formazione", "Migrazione dati"): sono attività **senza calcoli propri**.
- **[PROPOSTA]** Le macro-attività MANUALE e SVILUPPO sono liberamente ordinabili tra loro; PM e RISK sono sempre in coda, in quest'ordine, seguite da Totale e Totale economico.
- **[PROPOSTA]** Ogni macro-attività MANUALE ha il flag `inclusaInBase` (default `true`) che determina se entra nella base di calcolo di PM e Risk (§7.2).
- Se una macro-attività MANUALE ha almeno una voce, il suo valore è la somma delle voci e l'input diretto è disabilitato.

## 5. Profili e autorizzazioni

### 5.1 Ruoli globali (App Roles Entra ID)

| App Role | Descrizione |
|---|---|
| `Admin` | Vede e modifica tutto; gestisce aree, parametri globali e ruoli utenti. **Non crea progetti.** |
| `Writer` | Può creare progetti; opera sui progetti di cui è membro con ruolo WRITER. |
| `Viewer` **[PROPOSTA]** | Ruolo base per accedere all'applicazione; opera solo sui progetti di cui è membro con ruolo VIEWER. |

- I ruoli sono **App Roles** sulla registrazione dell'applicazione in Entra ID (tenant singolo) e sono letti dal claim `roles` del token.
- L'Admin assegna e revoca gli App Roles **dall'interfaccia dell'applicazione**; il backend li scrive in Entra ID tramite Microsoft Graph (`appRoleAssignedTo` sul service principal dell'app).
  - Prerequisito: permesso applicativo Graph **`AppRoleAssignment.ReadWrite.All`** (oltre a `User.Read.All` per la ricerca utenti), con consenso dell'amministratore del tenant. Da richiedere all'IT.
  - Le modifiche sono effettive al successivo rinnovo del token dell'utente (al più circa 1 ora); l'interfaccia lo segnala.
  - Non è possibile revocare l'ultimo Admin.
- **[PROPOSTA]** Un utente può avere più App Roles; i permessi sono l'unione. In particolare un utente Admin che deve creare un progetto riceve anche il ruolo `Writer`.
- **[PROPOSTA]** L'applicazione Entra ID è configurata con "assegnazione utente obbligatoria": chi non ha alcun App Role non può accedere.

### 5.2 Ruoli di progetto (gestiti nell'applicazione)

| Ruolo di progetto | Requisito globale | Capacità |
|---|---|---|
| `WRITER` | App Role `Writer` (o `Admin`) | Modifica il progetto, le sue stime, membri, referenti e override. |
| `VIEWER` | qualunque App Role | Sola lettura; **non vede importi né tariffa**; export solo senza colonna economica. |

- Chi crea un progetto ne diventa automaticamente membro `WRITER`.
- **[PROPOSTA]** Un progetto deve avere sempre almeno un membro `WRITER`.
- L'associazione utenti avviene tramite **ricerca nella directory del tenant** (Microsoft Graph, per nome o email), anche per utenti che non hanno mai effettuato l'accesso; al primo accesso l'utente ritrova i propri progetti (match su `objectId` Entra).
- Un Writer globale **non membro** di un progetto non lo vede.

### 5.3 Matrice dei permessi

Legenda: ✅ consentito · ❌ negato · W = membro WRITER del progetto · V = membro VIEWER del progetto.

| Azione | Admin | Writer (W) | Viewer (V) | Writer non membro |
|---|:-:|:-:|:-:|:-:|
| Accedere all'applicazione | ✅ | ✅ | ✅ | ✅ |
| Gestire App Roles degli utenti | ✅ | ❌ | ❌ | ❌ |
| Creare / modificare / archiviare aree | ✅ | ❌ | ❌ | ❌ |
| Eliminare un'area (solo se senza progetti) **[PROPOSTA]** | ✅ | ❌ | ❌ | ❌ |
| Vedere l'elenco aree attive | ✅ | ✅ | ✅ | ✅ |
| Modificare parametri globali | ✅ | ❌ | ❌ | ❌ |
| Consultare parametri globali | ✅ | ✅ | ❌ | ✅ |
| Creare un progetto | ❌ ¹ | — | — | ✅ |
| Vedere progetto e stime | ✅ | ✅ | ✅ | ❌ |
| Modificare anagrafica progetto | ✅ | ✅ | ❌ | ❌ |
| Spostare un progetto in un'altra area **[PROPOSTA]** | ✅ | ✅ | ❌ | ❌ |
| Archiviare / ripristinare un progetto | ✅ | ✅ | ❌ | ❌ |
| Eliminare un progetto | ✅ | ❌ | ❌ | ❌ |
| Gestire membri del progetto | ✅ | ✅ | ❌ | ❌ |
| Gestire referenti | ✅ | ✅ | ❌ | ❌ |
| Override parametri di progetto | ✅ | ✅ | ❌ | ❌ |
| Creare / duplicare / creare nuova versione di stima | ✅ | ✅ | ❌ | ❌ |
| Modificare stima, macro-attività, voci | ✅ | ✅ | ❌ | ❌ |
| Override parametri di stima, forzature | ✅ | ✅ | ❌ | ❌ |
| Cambiare stato della stima (incl. Accettato) | ✅ | ✅ | ❌ | ❌ |
| Eliminare una stima **[PROPOSTA]** | ✅ | ✅ ² | ❌ | ❌ |
| Vedere tariffa e importi | ✅ | ✅ | ❌ | — |
| Export senza colonna economica | ✅ | ✅ | ✅ | ❌ |
| Export con colonna economica | ✅ | ✅ | ❌ | ❌ |
| Consultare l'audit | ✅ (tutto) | ✅ (del progetto) | ❌ | ❌ |

¹ L'Admin non crea progetti per requisito; può farlo solo se ha anche il ruolo `Writer` (§5.1).
² Solo stime mai passate per lo stato *Accettato*.

## 6. Funzionalità

### 6.1 Aree (Admin)
- Creazione con: nome (obbligatorio, univoco), codice (facoltativo, univoco), descrizione.
- Modifica, archiviazione, ripristino. In un'area archiviata non si possono creare né spostare progetti; i progetti esistenti restano consultabili.
- Elenco con filtro attive/archiviate e conteggio progetti.

### 6.2 Progetti
- **Creazione** (Writer): area (tra quelle attive, obbligatoria), nome (obbligatorio), codice (univoco, facoltativo), cliente, descrizione.
- **Modifica** anagrafica, **spostamento** di area, **archiviazione / ripristino** (Admin, Writer membro). Un progetto archiviato è in sola lettura.
- **Eliminazione** (solo Admin): con conferma esplicita (digitazione del nome). **[PROPOSTA]** Eliminazione logica di progetto, stime, voci e override; il progetto scompare da tutte le viste e dagli export; l'evento è registrato in audit.
- **Membri**: aggiunta via ricerca Graph, scelta ruolo `WRITER`/`VIEWER`, modifica ruolo, rimozione.
- **Referenti**: anagrafica con nome, cognome, email, telefono, ruolo/funzione, tipo (`INTERNO` | `CLIENTE`), note. Nessun accesso all'app implicito.
- **Override parametri di progetto**: vedi §8.
- Elenco progetti filtrabile per area, stato (attivo/archiviato), testo libero; l'Admin vede tutti i progetti, gli altri solo quelli di cui sono membri.

### 6.3 Stime
- **Creazione**: titolo (obbligatorio), descrizione, codice (generato o inserito, univoco nel progetto). Nasce in stato *In attesa di requisiti*, versione 1, con le macro-attività di default (§4.1).
- **Duplicazione**: crea una nuova stima indipendente (nuovo codice, versione 1) copiando macro-attività, voci, override di stima e forzature; non copia stato, transizioni né SAL.
- **Nuova versione**: come la duplicazione, ma mantiene il codice e incrementa la versione; collega la nuova stima a quella d'origine. Le versioni precedenti restano consultabili e modificabili secondo il proprio stato.
- **Editor della stima** con due viste sincronizzate:
  - **Riepilogo** (equivalente al foglio *Stima*): righe per macro-attività, gg, descrizione, note, Totale, Totale economico (se visibile all'utente e se *Tariffa* è visibile).
  - **Dettaglio** per ogni macro-attività con voci (equivalente al foglio *Dettaglio intervento*): `#` automatico, attività, descrizione, gg netti (solo SVILUPPO), gg, note; righe di totale/test.
- **Ricalcolo immediato** a ogni modifica (anteprima lato client); il valore di riferimento è quello calcolato dal backend (§12.3).
- Voci: aggiunta, modifica, eliminazione, riordino (drag & drop). Campo *Note / punti aperti* su voci e macro-attività.
- **Forzature** su Test, PM, Risk: l'utente inserisce un valore che sostituisce quello calcolato, con **motivazione facoltativa**; l'interfaccia mostra sempre anche il valore calcolato e un indicatore di forzatura; la forzatura è rimovibile.
- **Override parametri di stima**: vedi §8.
- **Concorrenza**: blocco ottimistico; se due utenti modificano la stessa stima, il secondo salvataggio viene rifiutato con invito a ricaricare.

### 6.4 Parametri globali (Admin)
- Modifica di valore, label del parametro, label della riga collegata, visibilità per ciascuno dei 5 parametri (§7.1).
- Le modifiche si propagano subito a tutte le stime **non congelate** che non hanno override sul campo (§8).
- Ogni modifica è registrata in audit.

### 6.5 Export
Vedi §10.

### 6.6 Audit
- Consultazione dello storico per entità (area, progetto, stima) con filtri per utente, intervallo date, tipo di evento.

## 7. Regole di calcolo

### 7.1 Parametri

| Codice | Label parametro (default) | Label riga collegata (default) | Tipo | Default | Dominio |
|---|---|---|---|---|---|
| `TESTING` | % testing | Test funzionali | percentuale | 10% | 0–100% |
| `CONTINGENCY` | % contingency | Stima (gg/uomo) *(colonna con contingency)* | percentuale | 30% | 0–200% |
| `PM` | % PM | Project Management | percentuale | 15% | 0–100% |
| `RISK` | % Risk Budget | Risk budget | percentuale | 20% | 0–100% |
| `TARIFFA` | Tariffa | Totale economico | importo €/gg | 442,00 | ≥ 0 |

L'elenco è **fisso**: non è prevista la creazione di nuovi parametri. Notazione: `t`, `c`, `p`, `r`, `T` = valori risolti (§8) rispettivamente di TESTING, CONTINGENCY, PM, RISK, TARIFFA.

### 7.2 Formule

`R(x)` = arrotondamento all'intero, metà per eccesso (§7.3).

**Voci della macro-attività SVILUPPO** (input: `n_i` = gg netti della voce *i*)
```
g_i       = R(n_i + n_i × c)                          // gg con contingency
TotSvilN  = Σ n_i                                     // totale sviluppi netto
TotSvil   = Σ g_i                                     // totale sviluppi
TestN     = R(TotSvilN × t)                           // test netto
Test      = forzatura(TEST) ?? R(TotSvil × t)         // test
TotaleN   = TotSvilN + TestN                          // TOTALE netto (informativo)
Sviluppo  = TotSvil + Test                            // valore della macro-attività SVILUPPO (correzione I1)
```

**Macro-attività MANUALE** *m*
```
M_m = Σ voci_m.gg       se la macro-attività ha voci (nessun arrotondamento, nessuna contingency)
M_m = input_m           altrimenti
```

**Riepilogo**
```
Base     = Sviluppo + Σ M_m   (solo MANUALE con inclusaInBase = true)
PM       = forzatura(PM)   ?? R(Base × p)
Risk     = forzatura(RISK) ?? R((Base + PM) × r)
Totale   = Sviluppo + Σ M_m (tutte) + PM + Risk
TotEcon  = round2(Totale × T)
```

**Valorizzazione economica per riga** (solo quando richiesta, §10)
```
Econ(riga) = round2(gg(riga) × T)
```
Il *Totale economico* è sempre `round2(Totale × T)` e non la somma delle righe arrotondate; con gg interi le due quantità coincidono.

**Effetto della visibilità** (§8.3): un parametro non visibile vale **0** ai fini del calcolo e la riga collegata viene nascosta:

| Parametro nascosto | Effetto |
|---|---|
| `TESTING` | `Test = TestN = 0`; righe *Test funzionali* nascoste. |
| `CONTINGENCY` | `g_i = n_i`; colonna "senza contingency" nascosta (coincide con quella con contingency). |
| `PM` | `PM = 0`; riga nascosta. |
| `RISK` | `Risk = 0`; riga nascosta. |
| `TARIFFA` | Nessun valore economico visualizzato né esportabile. |

Una forzatura su una riga il cui parametro è nascosto è ignorata e non impostabile.

### 7.3 Arrotondamenti e precisione
- `R(x)` replica il `ROUND(x;0)` di Excel: arrotondamento all'intero con metà **lontano da zero** (per valori ≥ 0 equivale a `HALF_UP`; es. 6,5 → 7).
- `round2(x)`: 2 decimali, `HALF_UP`.
- Gli arrotondamenti si applicano **esattamente nei punti indicati in §7.2** (per voce, test, PM, risk), come nell'Excel.
- Tutti i calcoli in **aritmetica decimale** (es. `BigDecimal`), mai in virgola mobile. Nota: in rari casi limite Excel (che usa double) può arrotondare diversamente su valori al confine di x,5; l'applicazione è la fonte di verità.
- **Input gg** **[PROPOSTA]**: numeri ≥ 0 con al massimo 2 decimali.
- **Percentuali**: memorizzate come frazione con 4 decimali (es. 0,1500); visualizzate come percentuale.

### 7.4 Esempio di riferimento (dati del template, parametri di default)

| Voce | Netto | Con contingency 30% |
|---|--:|--:|
| Attività 1 | 6 | 8 |
| Attività 2 | 8 | 10 |
| Attività 3 | 7 | 9 |
| Attività 4 | 15 | 20 |
| Attività 5 | 8 | 10 |
| Attività 6 | 3 | 4 |
| Totale sviluppi | 47 | 61 |
| Test funzionali (10%) | 5 | 6 |
| TOTALE | 52 | 67 |

| Riepilogo | gg |
|---|--:|
| Setup ambienti | 3 |
| Analisi | 10 |
| Attività di sviluppo | 67 |
| Project Management = R(80 × 0,15) | 12 |
| Risk budget = R(92 × 0,20) | 18 |
| **Totale** | **110** |
| **Totale economico** = 110 × 442 | **48.620,00 €** |

## 8. Parametri: risoluzione globale / progetto / stima

### 8.1 Livelli
Ogni parametro ha quattro attributi sovrascrivibili in modo **indipendente**: `valore`, `label` (del parametro), `labelRiga` (della riga/colonna collegata), `visibile`.

```
risolto(stima, parametro, attributo) =
    overrideStima(stima, parametro).attributo     se non nullo
 ?? overrideProgetto(progetto, parametro).attributo se non nullo
 ?? globale(parametro).attributo
```

- Un override con un attributo nullo eredita quell'attributo dal livello superiore (es. si può cambiare solo la label lasciando ereditare il valore).
- L'interfaccia mostra per ogni attributo la **provenienza** (globale / progetto / stima) e permette di **ripristinare** l'ereditarietà (azzerando l'override).
- La `labelRiga` sovrascritta rinomina la riga (o la colonna, per CONTINGENCY) nell'editor e negli export.

### 8.2 Congelamento all'accettazione
- Al passaggio in stato **Accettato**, tutti gli attributi risolti dei 5 parametri vengono **materializzati come override di stima** e la stima è marcata `parametriCongelati = true`. Da quel momento modifiche ai parametri globali o di progetto **non** hanno effetto sulla stima.
- Viene inoltre salvata nella transizione di stato una fotografia dei totali (gg ed €) a scopo di tracciabilità.
- **[PROPOSTA]** Se la stima esce dallo stato Accettato (es. per aggiungere una CR) resta congelata: i parametri cambiano solo modificando esplicitamente gli override di stima. È disponibile l'azione esplicita **"Riallinea ai parametri del progetto"** (Writer/Admin) che rimuove gli override di stima e il congelamento, con conferma e audit.

### 8.3 Visibilità
- `visibile = false` ⇒ il parametro vale 0 nei calcoli e la riga/colonna collegata è nascosta in editor ed export (§7.2).
- Il parametro nascosto resta visibile, con indicazione "nascosto", nella pagina di configurazione per Admin e Writer, per poterlo riattivare.

## 9. Ciclo di vita della stima

### 9.1 Stati

| Stato | Codice | Significato |
|---|---|---|
| In attesa di requisiti | `IN_ATTESA_REQUISITI` | Stato iniziale; stima in elaborazione. |
| In fase di accettazione | `IN_ACCETTAZIONE` | Stima sottoposta al cliente/committente. |
| Accettato | `ACCETTATO` | Stima approvata; abilita (in futuro) il SAL. |

### 9.2 Transizioni
- **Tutte le transizioni tra i tre stati sono consentite**, in entrambe le direzioni, per consentire ad esempio l'aggiunta di CR a una stima già accettata.
- Possono cambiare stato: **Admin** e **membri WRITER** del progetto.
- In **tutti** gli stati la stima è modificabile (voci incluse, anche quelle già accettate).
- Ogni transizione registra: stato di partenza e di arrivo, utente, data/ora, **motivazione facoltativa**.
- Transizione verso **Accettato**: registra inoltre data di accettazione (default oggi) e un **riferimento facoltativo** (es. numero d'ordine), esegue il congelamento (§8.2).
- Un progetto può avere **più stime accettate** (es. lotti, CR).

### 9.3 Vincoli per il SAL futuro
Anche se il SAL è fuori scope, il modello deve garantire che:
- ogni voce e macro-attività abbia un **identificativo stabile** che non cambia con modifiche, riordini o cambi di stato;
- la modifica di una voce (anche dei suoi gg) **non azzeri** il SAL già registrato;
- una voce eliminata in una stima che è stata accettata almeno una volta venga **eliminata logicamente**, per preservare il SAL storico.

## 10. Export Excel

### 10.1 Opzioni
All'avvio dell'export l'utente sceglie:

| Opzione | Valori | Note |
|---|---|---|
| Colonna economica | `Nessuna` · `Solo totale` · `Per voce` | `Solo totale`: riga *Totale economico* come nel template. `Per voce`: colonna € per ogni riga di riepilogo e dettaglio, più il totale. Il Viewer può scegliere solo `Nessuna`. Non disponibile se *Tariffa* è nascosta. |
| Formule | `Sì` · `No` | `Sì` (uso interno): celle calcolate con formule Excel funzionanti e foglio parametri. `No` (invio al cliente): solo valori. |

### 10.2 Contenuto
- **Foglio "Stima"**: intestazione (area, progetto, stima, codice, versione, stato, data export), riepilogo per macro-attività con `#`, Attività, Descrizione, gg, [€], Note; Totale; [Totale economico].
- **Un foglio di dettaglio per ogni macro-attività con voci** (es. "Dettaglio sviluppo"): voci con `#`, Attività, Descrizione, [gg netti], gg, [€], Note; per SVILUPPO anche Totale sviluppi, Test funzionali, TOTALE.
- **Foglio "Parametri"** (solo con formule): parametri visibili con label e valore risolti; senza colonna economica la Tariffa è esclusa.
- Le label rispettano gli override; righe e colonne di parametri nascosti sono omesse.
- Con formule: le formule devono restituire **esattamente** i valori dell'applicazione (stessi punti di arrotondamento, riferimenti assoluti o nomi definiti). Le righe forzate sono esportate come **valori** con un commento di cella "Valore forzato: <motivazione>".
- Senza formule: solo valori, nessun foglio parametri, nessun commento interno.
- Nome file: `<codiceProgetto>_<codiceStima>_v<versione>_<AAAAMMGG>.xlsx`.
- I controlli di autorizzazione sull'export sono applicati dal backend (§12.1).

## 11. Modello dati logico

Tutte le entità hanno inoltre `creatoIl`, `creatoDa`, `modificatoIl`, `modificatoDa`; le entità modificabili hanno `versioneRecord` per il blocco ottimistico.

| Entità | Attributi principali | Vincoli |
|---|---|---|
| **Utente** | id, entraObjectId, nome, email, ultimoAccesso | entraObjectId univoco; i ruoli globali non sono memorizzati (letti dal token / Graph). |
| **Area** | id, nome, codice, descrizione, archiviata | nome univoco; codice univoco se presente. |
| **Progetto** | id, areaId, codice, nome, cliente, descrizione, archiviato, eliminato | codice univoco se presente. |
| **MembroProgetto** | progettoId, utenteId, ruolo (`WRITER`\|`VIEWER`) | PK (progettoId, utenteId); almeno un WRITER per progetto. |
| **Referente** | id, progettoId, nome, cognome, email, telefono, ruolo, tipo (`INTERNO`\|`CLIENTE`), note | — |
| **ParametroGlobale** | codice (`TESTING`\|`CONTINGENCY`\|`PM`\|`RISK`\|`TARIFFA`), tipo, valore, label, labelRiga, visibile | 5 righe fisse. |
| **OverrideParametroProgetto** | progettoId, codice, valore?, label?, labelRiga?, visibile? | PK (progettoId, codice); attributi nullabili = ereditati. |
| **OverrideParametroStima** | stimaId, codice, valore?, label?, labelRiga?, visibile? | PK (stimaId, codice). |
| **Stima** | id, progettoId, codice, versione, stimaOrigineId?, titolo, descrizione, stato, parametriCongelati, dataAccettazione?, riferimentoAccettazione?, eliminata | (progettoId, codice, versione) univoco. |
| **MacroAttivita** | id, stimaId, tipo (`SVILUPPO`\|`MANUALE`\|`PM`\|`RISK`), nome, descrizione, note, ordine, ggManuale?, inclusaInBase, eliminata | un solo SVILUPPO, PM, RISK per stima. |
| **Voce** | id, macroAttivitaId, ordine, attivita, descrizione, gg, note, eliminata | `gg` = netto per SVILUPPO, effettivo per MANUALE. |
| **Forzatura** | stimaId, riga (`TEST`\|`PM`\|`RISK`), valore, motivazione? | PK (stimaId, riga). |
| **TransizioneStato** | id, stimaId, statoDa, statoA, utenteId, timestamp, motivazione?, riferimento?, fotografiaTotali (JSON) | immutabile. |
| **EventoAudit** | id, timestamp, utenteId, entita, entitaId, progettoId?, azione (`CREATE`\|`UPDATE`\|`DELETE`\|`STATE`\|`EXPORT`\|`ROLE`), campo?, valorePrima?, valoreDopo?, motivazione? | immutabile; append-only. |

**Tipi numerici** (Oracle): gg `NUMBER(9,2)`; percentuali `NUMBER(7,4)`; tariffa e importi `NUMBER(14,2)`.

I valori calcolati (g_i, Test, PM, Risk, totali) **non sono persistiti** sulle entità: sono ricalcolati dal backend; fa eccezione la `fotografiaTotali` della transizione di stato.

## 12. Requisiti non funzionali

### 12.1 Sicurezza
- Autenticazione OIDC/OAuth2 con Microsoft Entra ID, **tenant singolo**; il frontend ottiene un access token per l'API.
- Il backend valida firma, issuer, audience, tenant e scadenza del token; i ruoli derivano dal claim `roles`.
- **Ogni endpoint applica le autorizzazioni lato server** secondo §5.3, inclusa l'appartenenza al progetto; nessuna decisione di autorizzazione è demandata al frontend.
- Per i Viewer i campi economici (tariffa, importi, parametro TARIFFA) sono **rimossi dalle risposte API**, non solo nascosti in interfaccia; l'export con colonna economica è rifiutato (HTTP 403).
- I valori calcolati inviati dal client sono ignorati: il backend ricalcola sempre.
- Validazione di tutti gli input (domini §7.1, lunghezze, gg ≥ 0); protezione da accessi diretti per id (verifica che l'oggetto appartenga a un progetto accessibile).
- Il secret/certificato per Microsoft Graph è custodito in un vault (es. Azure Key Vault), mai nel codice.

### 12.2 Audit
- Registrazione **campo per campo** (valore prima/dopo) di ogni creazione, modifica, eliminazione su: aree, progetti, membri, referenti, parametri globali, override, stime, macro-attività, voci, forzature; inoltre cambi di stato, assegnazioni di App Roles ed export (chi, quando, opzioni scelte).
- Il log è append-only e non modificabile da alcun profilo.
- Consultazione: Admin su tutto, Writer membro sul proprio progetto.
- **[PROPOSTA]** Conservazione: illimitata.

### 12.3 Precisione dei calcoli
- Motore di calcolo unico nel backend, in aritmetica decimale, con le regole di arrotondamento di §7.3; il frontend può replicarlo per l'anteprima ma la risposta del backend prevale.
- Test automatici obbligatori sul motore con almeno i casi di §14.4.

### 12.4 Altri
- Interfaccia **solo in italiano**; formati numerici italiani (virgola decimale, separatore delle migliaia, `€` in coda).
- Browser supportati: ultime due versioni di Edge e Chrome.
- Tempo di risposta del ricalcolo < 500 ms per stime fino a 500 voci.

### 12.5 Stack tecnologico **[PROPOSTA — da definire]**
- Backend: Java / Spring Boot, Spring Security OAuth2 Resource Server, Apache POI per l'export, migrazioni DB con Flyway o Liquibase.
- Frontend: React con MSAL (`@azure/msal-react`).
- Database: Oracle.
- Hosting: Azure; identità: Microsoft Entra ID; integrazione Microsoft Graph.

## 13. Fuori scope

- **SAL** (avanzamento percentuale per voce) e relative regole: **fase successiva**; il modello ne predispone solo i vincoli (§9.3).
- Import delle stime da file Excel esistenti.
- Accesso da tenant esterni (multi-tenant) e utenti guest.
- Interfaccia multilingua.
- Tariffe multiple per figura professionale; valute diverse da EUR.
- Creazione di nuovi parametri o formule configurabili.
- Notifiche (email, Teams) sui cambi di stato.
- Workflow di approvazione formale da parte del cliente.

## 14. Criteri di accettazione

### 14.1 Autenticazione e profili
- **CA-01** Un utente senza App Role non accede all'applicazione.
- **CA-02** Un Writer vede solo i progetti di cui è membro; un Admin vede tutti i progetti.
- **CA-03** Un Admin senza ruolo Writer non vede il comando "Nuovo progetto" e la relativa API risponde 403.
- **CA-04** Chi crea un progetto ne diventa membro WRITER.
- **CA-05** Un Viewer non riceve tariffa e importi in nessuna risposta API; l'export con colonna economica gli restituisce 403.
- **CA-06** L'Admin assegna un App Role da interfaccia; dopo il rinnovo del token l'utente ha i nuovi permessi; la revoca dell'ultimo Admin è rifiutata.
- **CA-07** Una chiamata API su un progetto di cui l'utente non è membro (e non è Admin) risponde 403/404 anche conoscendone l'id.

### 14.2 Struttura
- **CA-10** Una nuova stima contiene Setup ambienti, Analisi, Attività di sviluppo, Project Management, Risk budget.
- **CA-11** Si può aggiungere una macro-attività MANUALE con o senza voci; con voci, il suo valore è la somma delle voci.
- **CA-12** Le macro-attività SVILUPPO, PM e RISK non sono eliminabili.
- **CA-13** Duplicazione e nuova versione copiano macro-attività, voci, override di stima e forzature; lo stato della copia è *In attesa di requisiti*.

### 14.3 Parametri
- **CA-20** Con override di progetto `CONTINGENCY = 20%`, tutte le stime non congelate del progetto lo usano; le altre continuano a usare il globale.
- **CA-21** Un override di stima prevale su quello di progetto, che prevale sul globale, per ogni attributo indipendentemente.
- **CA-22** Rinominando `labelRiga` di RISK in "Imprevisti", riga ed export mostrano "Imprevisti".
- **CA-23** Dopo l'accettazione, la modifica di un parametro globale non cambia i totali della stima accettata.

### 14.4 Calcoli (dati del template, §7.4)
- **CA-30** Parametri di default ⇒ Sviluppo 67, PM 12, Risk 18, **Totale 110 gg**, **Totale economico 48.620,00 €**; colonna netta: 47 / 5 / 52.
- **CA-31** `CONTINGENCY = 20%` ⇒ voci 7, 10, 8, 18, 10, 4; Totale sviluppi 57; Test 6; Sviluppo 63; PM 11; Risk 17; **Totale 104**; **45.968,00 €**.
- **CA-32** `RISK` non visibile ⇒ riga Risk assente; **Totale 92**; **40.664,00 €**.
- **CA-33** Forzatura PM = 15 ⇒ Risk = R(95 × 0,20) = 19; **Totale 114**; **50.388,00 €**; il calcolato (12) resta visibile.
- **CA-34** Una voce di sviluppo di 5 gg netti con contingency 30% vale **7** (6,5 arrotondato per eccesso).
- **CA-35** Una macro-attività MANUALE "Formazione" da 5 gg con `inclusaInBase = true` porta la Base a 85 ⇒ PM 13, Risk 20, **Totale 118**, **52.156,00 €**.

### 14.5 Ciclo di vita
- **CA-40** Sono consentite tutte le transizioni tra i tre stati, per Admin e WRITER; non per VIEWER.
- **CA-41** Il passaggio ad Accettato registra data, riferimento facoltativo, fotografia dei totali e congela i parametri.
- **CA-42** Una stima accettata può essere riportata in *In attesa di requisiti*, modificata (aggiunta voci CR) e riaccettata; gli identificativi delle voci preesistenti non cambiano.

### 14.6 Export
- **CA-50** Export *con formule*: aprendo il file in Excel i totali coincidono con quelli dell'applicazione.
- **CA-51** Export *senza formule*: nessuna cella contiene formule; nessun foglio parametri.
- **CA-52** Export con colonna economica `Per voce` mostra l'importo per ogni riga; `Solo totale` mostra solo il Totale economico; `Nessuna` non contiene importi né tariffa.
- **CA-53** Righe di parametri nascosti e label sovrascritte sono rispettate nell'export.

### 14.7 Audit
- **CA-60** La modifica dei gg di una voce produce un evento con campo, valore prima e valore dopo, utente e data/ora.
- **CA-61** Ogni export produce un evento con le opzioni scelte.

## 15. Domande aperte residue

Proposte inserite nel documento e da confermare:

1. **Congelamento dopo la riapertura** (§8.2): una stima riportata indietro da Accettato resta congelata, con l'azione esplicita "Riallinea ai parametri del progetto"?
2. **Eliminazione progetto** (§6.2): logica (proposta) o fisica definitiva?
3. **Ruolo `Viewer` globale e assegnazione obbligatoria** (§5.1): confermi un terzo App Role base per l'accesso, con accesso negato a chi non ha ruoli?
4. **Admin che crea progetti** (§5.1): ok la soluzione "assegnarsi anche il ruolo Writer"?
5. **Base di PM/Risk per nuove macro-attività MANUALE** (§4.1): flag `inclusaInBase` con default `true`?
6. **Spostamento progetto di area** (§5.3): consentito anche al Writer membro o solo all'Admin?
7. **Eliminazione stima** (§5.3): consentita a Writer/Admin solo se mai accettata?
8. **Eliminazione area** (§5.3): consentita solo se vuota?
9. **Input gg decimali** (§7.3): ammessi fino a 2 decimali?
10. **Export per il cliente** (§10): nell'export senza formule vanno escluse anche la colonna "senza contingency" e/o le note?
11. **Ordinamento macro-attività** (§4.1): PM e Risk sempre in coda?
12. **Minimo un WRITER per progetto** (§5.2).
13. **Conservazione audit** (§12.2): illimitata?
14. **Permesso Graph `AppRoleAssignment.ReadWrite.All`** (§5.1): da validare con l'IT; in caso di diniego, fallback su ruoli gestiti nel database dell'applicazione.
15. **Stack tecnologico** (§12.5): da confermare.
