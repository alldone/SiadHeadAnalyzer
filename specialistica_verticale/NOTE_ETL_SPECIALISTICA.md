# ETL Specialistica - Stato Lavoro

## Scopo

Nella cartella `bancadati` sono presenti file Excel e alcuni PDF con l'elenco delle prestazioni da importare nel database.

L'ETL legge i file sorgente, normalizza i dati e produce un file di output per struttura nel formato richiesto da `FORMAT_ACQUAVIVA.xlsx`.

Le colonne del file output sono:

- `Codice STS.11`
- `BRANCA`
- `CODICE REGIONALE`
- `CODICE SSN`

## File di riferimento

- Template output: `FORMAT_ACQUAVIVA.xlsx`
- Tabella di codifica: `BRANCA_Codici regionali-Codici SSN.xlsx`
- Sorgenti input: cartella `bancadati`
- Script ETL: `etl_bancadati.py`
- Script validazione: `validate_output.py`

## Regole ETL generali

### Identificazione struttura

- Lo `STS` viene letto dalla colonna `CODICE STS11 (*)`.
- La denominazione viene letta dalla colonna `DENOMINAZIONE STRUTTURA`.
- Se in una cartella esistono piu file della stessa struttura e uno dei file non riporta lo `STS`, viene ereditato dal file fratello della stessa cartella se il valore e univoco.

### Identificazione branca

- La branca viene normalizzata a codice numerico a 2 cifre.
- Le informazioni vengono lette da:
  - `BRANCA SPECIALISTICA (*)`
  - `CODICE BRANCA DPCM 2017 (*)`
- Sono gestiti casi con:
  - `XX - DESCRIZIONE`
  - solo `XX`
  - sola descrizione
  - piu branche nella stessa cella

### Identificazione prestazioni

- Il codice prestazione viene letto principalmente da:
  - `CODICE PRESTAZIONE (*)`
  - in alcuni layout da `CODICE PRESTAZIONE (Regionale)`
- Se la cella contiene codice e descrizione, viene estratto solo il codice.
- Se la cella contiene piu codici separati da virgole o testo, vengono estratti tutti i codici validi.
- Se la cella indica che sono contrattualizzate tutte le prestazioni della branca, l'ETL espande automaticamente tutte le prestazioni della branca usando la tabella di codifica.

### Conversione regionale -> SSN

- La conversione standard usa `BRANCA_Codici regionali-Codici SSN.xlsx`.
- Il match considera:
  - il `CODICE REGIONALE`
  - la `BRANCA`
- Se il codice regionale non viene trovato ma il codice e gia un `CODICE SSN` valido per quella branca, il valore viene riscritto identico in `CODICE SSN`.

## Naming output

- I file output sono creati nella cartella `output`.
- Il nome e nel formato:
  - `STS_DENOMINAZIONE.xlsx`
  - oppure `STS_01_DENOMINAZIONE.xlsx`, `STS_02_DENOMINAZIONE.xlsx`, ecc. se la stessa struttura ha piu file input

## Gestione PDF

Sono stati inclusi nel flusso anche i PDF testuali:

- `ECORAD`
- `STUDIO RADIOLOGIA ARCERI`

## Regola speciale RONTGEN

Nel file input `bancadati/RONTGEN/RONTGEN.xlsx` e presente una casistica particolare nella colonna prestazioni:

- il codice fuori parentesi rappresenta un codice SSN dichiarato nel file input
- il codice tra parentesi rappresenta il codice regionale
- in alcuni casi una stessa porzione di testo contiene piu parentesi dopo lo stesso codice fuori parentesi

### Regola applicata

Per `RONTGEN` il file output viene compilato cosi:

1. si legge il `CODICE REGIONALE` dalle parentesi
2. si verifica se quel `CODICE REGIONALE` esiste nella tabella di codifica per la branca `08`
3. se esiste, nel file output si usa il `CODICE SSN` preso dalla tabella di codifica
4. se non esiste in tabella, il `CODICE SSN` nel file output resta vuoto

### Esempi RONTGEN

Casi valorizzati dalla tabella:

- `87.24.R2 -> 87.24`
- `87.29.R -> 87.29`
- `87.09.1.R -> 87.09.1`
- `87.62.R -> 87.62`
- `87.16.4.R -> 87.16.4`
- `87.43.1.R2 -> 87.43.1`

Casi lasciati vuoti perche il codice regionale non esiste nella tabella:

- `87.16.1.R2`
- `88.29.2.R1`
- `87.43.2.R1`
- `87.43.2.R3`
- `87.44.1.R1`

## Report prodotti

### Output principali

- Cartella output: `output`
- Manifest output: `output/_manifest.tsv`
- Log incrementali: `output/_incrementali.txt`
- Log anomalie ETL: `output/_anomalie.tsv`

### Report anomalie residue

- TSV riepilogo anomalie: `output/_prestazioni_non_ricavabili.tsv`
- Excel unico anomalie: `output/_anomalie_non_riconducibili_ssn.xlsx`

Il file Excel unico contiene 3 colonne:

- `CODICE STRUTTURA`
- `CARTELLA/FILE INPUT`
- `PRESTAZIONI NON RICONDUCIBILI A SSN`

Attualmente contiene 32 righe utili, una per ogni file input con anomalie residue.

### Report dettagliato match da parentesi

- `output/_match_parentesi_espliciti.tsv`

Questo report contiene:

- `codice_regionale`
- `codice_ssn_input`
- `codice_ssn_output`
- `verifica_catalogo`
- `regola_output`

Serve in particolare per verificare la logica applicata al caso `RONTGEN`.

## Validazione finale

Ultima validazione eseguita con `validate_output.py`:

- sorgenti supportati: `89`
- output attesi: `89`
- output presenti: `89`
- errori strutturali: `0`
- warning semantici: `34`

File di validazione:

- `output/_validation_summary.txt`
- `output/_validation_report.tsv`
- `output/_validation_per_file.tsv`

## Stato attuale

Il flusso ETL e operativo.

Gli output sono stati generati, i casi speciali principali sono stati gestiti, e resta disponibile un file Excel unico per richiamare le strutture che hanno ancora prestazioni non riconducibili a SSN.
