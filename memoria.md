# Memoria Progetto SiadHeadAnalyzer

## Verticale Mobilita Farmaci

### Obiettivo

Nuovo verticale dedicato alla mobilita infraregionale farmaci, con GUI `tkinter` separata e integrazione nel launcher principale.

Il verticale:

- acquisisce separatamente la cartella `attiva` e la cartella `passiva`;
- abbina per ogni azienda erogatrice i file `<codice> importi.csv` e `<codice> prestazioni.csv`;
- riconosce i file a partire dal codice azienda iniziale nel filename, con deduzione intelligente di `importi` o `prestazioni` dal nome e fallback sul contenuto del CSV;
- usa una listbox multi-selezione per scegliere le colonne da riportare;
- di default seleziona `FARMACEUTICA` e `SOMM. DIRETTA DI FARMACI`;
- scrive un workbook Excel con due fogli: `attiva` e `passiva`;
- esegue sempre un controllo automatico degli importi tra input e output.

### Convenzioni dataset rilevate

- L'azienda erogatrice e' nel nome file, ad esempio `201 importi.csv`.
- Nel flusso `attiva` la prima colonna del CSV e' `ASL`.
- Nel flusso `passiva` la prima colonna del CSV e' `AZIENDA SANITARIA CREDITRICE`.
- Nei CSV `passiva` la colonna `SOMM. DIRETTA DI  FARMACI` puo avere doppio spazio interno, da normalizzare in lettura.
- I numeri usano il formato italiano:
  `.` separatore delle migliaia, `,` separatore dei decimali.

### Ordinamento richiesto

Le aziende erogatrici nel report devono essere ordinate cosi:

- `201`
- `202`
- `203`
- `204`
- `205`
- `912`
- `914`
- `915`
- `916`

Eventuali codici non previsti vanno dopo quelli ordinati esplicitamente.

### Struttura report Excel

Ogni workbook contiene esattamente due fogli:

- `attiva`
- `passiva`

Ogni foglio espone le colonne:

- `AZIENDA EROGATRICE`
- `AZIENDA SANITARIA DEBITRICE` nel foglio `attiva`
- `AZIENDA SANITARIA CREDITRICE` nel foglio `passiva`
- `FARMACEUTICA - IMPORTI`
- `FARMACEUTICA - PRESTAZIONI`
- `SOMM. DIRETTA DI FARMACI - IMPORTI`
- `SOMM. DIRETTA DI FARMACI - PRESTAZIONI`

Ogni foglio deve includere anche:

- una riga di `SUBTOTALE` dopo ogni azienda erogatrice;
- una riga finale di `TOTALE GENERALE`.

Translator aziende erogatrici:

- `201` -> `180201 - ASP COSENZA`
- `202` -> `180202 - ASP CROTONE`
- `203` -> `180203 - ASP CATANZARO`
- `204` -> `180204 - ASP VIBO VALENTIA`
- `205` -> `180205 - ASP REGGIO CALABRIA`
- `912` -> `912 - AO ANNUNZIATA - COSENZA`
- `914` -> `914 - AOU RENATO DULBECCO - CATANZARO`
- `915` -> `915 - AO BIANCHI MELACRINO MORELLI GOM - REGGIO CALABRIA`
- `916` -> `916 - INRCA`

Le righe vengono prodotte accoppiando per ogni controparte:

- importi farmaceutica
- prestazioni farmaceutica
- importi somm. diretta di farmaci
- prestazioni somm. diretta di farmaci

### File principali del verticale

- `mobilita_verticale/mobilita_report.py`
  logica di validazione input, parsing CSV, normalizzazione numeri, ordinamento, scrittura Excel e controllo importi input/output.
- `mobilita_verticale/mobilita_gui.py`
  GUI del verticale con selezione cartella, output, validazione e lancio report.
- `mobilita_gui.py`
  launcher standalone del verticale.

## Verticale SIAD

## Obiettivo

Applicazione Python con GUI cross-platform per:

- selezionare `TRACCIATO 1 XSD`, `TRACCIATO 2 XSD`, cartella XML e file di output;
- cercare ricorsivamente i file `SIAD*.xml`;
- classificare i file in base agli XSD;
- mostrare in GUI l'elenco dei file letti e il riepilogo aggregato;
- generare un report Excel multi-sheet.

## File principali

- `siad_report_gui.py`
  Logica applicativa completa: parsing XML/XSD, GUI `tkinter`, aggregazioni, export Excel.
- `README.md`
  Istruzioni di uso, build e release.
- `requirements.txt`
  Dipendenze runtime.
- `requirements-dev.txt`
  Dipendenze build.
- `siad_report_gui.spec`
  Configurazione `PyInstaller`.
- `.github/workflows/release.yml`
  Workflow GitHub Actions per build/release macOS e Windows.

## Regole di conteggio attuali

### Unicita'

- L'unicita' del `CF` e' calcolata per azienda (`CodiceASL`), non globalmente.
- Le prese in carico sono deduplicate per `Tracciato + CodiceASL + Id_Rec`.

### Tracciato 1

- Conta le nuove prese in carico con data nell'anno di analisi.

### Tracciato 2

- Se la data presa in carico e' antecedente al `01/01/anno di analisi`, il record entra tra le prese precedenti ancora attive.
- Se la data e' nell'anno di analisi, il record entra nella colonna separata:
  `Prese in carico T2 ANNO con CF non ancora presente`
  solo se il `CF` non e' gia' presente tra i `CF` univoci conteggiati per la stessa azienda nei tracciati 1 e 2.

### Totali

- `TOT. PRESE IN CARICO attive` = prese conteggiate secondo le regole sopra.
- `TOT. CF non univoci attivi` = stesso numero delle prese conteggiate.
- Il riepilogo dei `CF` e' organizzato in tre gruppi con lo stesso set di metriche:
  totale, `>= 65 anni`, `cf ambigui`.
- Gruppo `CF per azienda`
  calcolato sui `CF` univoci per azienda.
- Gruppo `Teste singole globali`
  calcolato sui `CF` esclusivi azienda, cioe' non condivisi con altre aziende.
- Gruppo `Differenze`
  calcolato sui `CF` condivisi tra aziende.

## Gestione eta' e CF

- Il codice fiscale e' estratto dagli ultimi 16 caratteri di `Id_Rec`.
- L'eta' e' calcolata al `31/12` dell'anno di analisi.
- I `CF` ambigui a cavallo del secolo vengono disambiguati usando `AnnoNascita` del tracciato 1 quando disponibile.
- Se non disambiguabili, vengono conteggiati come `CF ambigui`.

## Struttura report Excel

### Sheet `Report`

- riepilogo per azienda;
- include riga finale `TOTALE`;
- include styling con header formattati, righe alternate e totale evidenziato.

Colonne attuali:

- `SEDE`
- `Prese in carico precedenti ancora attive al 01/01/ANNO`
- `Nuove Prese in carico ANNO`
- `Prese in carico T2 ANNO con CF non ancora presente`
- `TOT. PRESE IN CARICO attive nel ANNO`
- `TOT. CF non univoci attivi nel ANNO`
- `[CF per azienda] TOT. PAZIENTI* attivi nel ANNO`
- `[CF per azienda] di cui >= 65 anni`
- `[CF per azienda] numero cf ambigui`
- `[Teste singole globali] CF esclusivi azienda`
- `[Teste singole globali] di cui >= 65 anni`
- `[Teste singole globali] numero cf ambigui`
- `[Differenze] CF condivisi con altre aziende`
- `[Differenze] di cui >= 65 anni`
- `[Differenze] numero cf ambigui`

### Sheet `Dettaglio`

Una riga per ogni record letto, con:

- azienda di competenza;
- tracciato;
- trimestre;
- file XML;
- `Id_Rec`;
- `CF`;
- anno nascita;
- eta';
- `IsOver65`;
- flag `Incluso nel report`;
- nota operativa.

### Sheet `Dettaglio_<ASL>`

- un foglio separato per ogni azienda.

### Sheet `CF_Univoci`

- elenco dei `CF` univoci per azienda;
- trimestri riscontrati;
- numero occorrenze;
- numero occorrenze incluse;
- anno nascita usato;
- eta';
- ambiguita'.

### Sheet `CF_Esclusi`

Contiene solo i `CF` distinti per azienda che:

- non entrano in nessun elenco conteggiato;
- non sono meri duplicati tecnici.

Nota: con il dataset attuale questo foglio risulta vuoto.

## GUI attuale

La GUI usa `tkinter` ed e' organizzata con due tab:

- `File XML`
  elenco dei file trovati con path relativo e tipo tracciato;
- `Riepilogo`
  vista tabellare del riepilogo aggregato.

## Build locale

### Esecuzione

macOS:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python siad_report_gui.py
```

Windows:

```powershell
py -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python siad_report_gui.py
```

### Build eseguibile

```bash
python -m pip install -r requirements.txt -r requirements-dev.txt
pyinstaller --noconfirm siad_report_gui.spec
```

Output previsto:

- `dist/ReportSIAD/`

## GitHub Releases

Workflow presente:

- `.github/workflows/release.yml`

Fa build su:

- `macos-15-intel` per macOS Intel
- `macos-latest` per macOS Apple Silicon
- `windows-latest`

Pubblica asset `.zip` su release quando viene pushato un tag `v*`.

### Cosa eseguire per generare i deploy su GitHub

Dopo aver collegato il repository remoto:

```bash
git init
git add .
git commit -m "Initial SIAD report app"
git branch -M main
git remote add origin <URL_REPO_GITHUB>
git push -u origin main
```

Per generare una release con build automatiche macOS e Windows:

```bash
git tag v1.0.0
git push origin v1.0.0
```

Effetto:

- GitHub Actions esegue la workflow `Build Release`
- viene compilata l'app su macOS Intel, macOS Apple Silicon e Windows
- vengono generati gli artifact `.zip`
- gli artifact vengono allegati alla release del tag

Per una release successiva:

```bash
git add .
git commit -m "Aggiornamenti progetto"
git push origin main
git tag v1.0.1
git push origin v1.0.1
```

Se vuoi lanciare la build senza tag, puoi usare anche `workflow_dispatch` dalla pagina Actions di GitHub, ma in quel caso gli artifact vengono caricati come artifact della workflow e non come release assets, salvo pubblicazione da tag.

## Verticale FAR Recon (Riconciliazione NSIS / SISR)

### Obiettivo

Verticale dedicato alla riconciliazione tra dati FAR gia' acquisiti su NSIS e nuovi flussi SISR per le 5 ASP calabresi (201..205). Produce un report Excel con conteggi T1/T2, scarti, delta validi e proiezione post-upload.

### Regole di conteggio (DM 17/12/2008, Specifiche FAR v6.3)

- **Tracciato 1**: 1 record = 1 wrapper `<FlsResSemires_1>`.
  Chiave NSIS: `(CodASL, CodStruttura, ID_REC, Data, TipoPrestazione)`.

- **Tracciato 2**: 1 record = 1 evento (Tariffa, PrestazioniSR, Valutazione, Sospensione, Dimissione).
  Conteggio rapido: conta occorrenze tag XML (non dedup).
  Conteggio NSIS § 4.3 (dedup per chiave-evento):
    - Tariffa: `(CodASL, CodStruttura, ID_REC, "Tariffa", Data)`
    - PrestazioniSR: `(CodASL, CodStruttura, ID_REC, "PrestazioniSR", TempoPieno, TempoParziale)`
    - Valutazione: `(CodASL, CodStruttura, ID_REC, "Valutazione", Tipo, Data)`
    - Sospensione: `(CodASL, CodStruttura, ID_REC, "Sospensione", Data, DataFine)`
    - Dimissione: `(CodASL, CodStruttura, ID_REC, "Dimissione", Data)`

- **Scarti SISR**: conteggio righe FLAG_OK_KO per Azienda x Trimestre x Tracciato (logica Nicoletta — NO dedup globale per id_rec).

- **Proiezione post-upload**: unione delle chiavi NSIS § 4.3 tra acquisiti e nuovi, con calcolo overlap reale.

### File del verticale

- `far_recon_verticale/__init__.py`
- `far_recon_verticale/far_recon_core.py` — logica pura: scansione, parsing, chiavi NSIS, scarti, proiezione, export xlsx.
- `far_recon_verticale/far_recon_gui.py` — GUI `FarReconApp` (tkinter, embed_mode).
- `far_recon_gui.py` — launcher standalone.

### Sheet prodotti nel workbook

1. `acquisiti_NSIS` — conteggio rapido T1/T2 per azienda (se cartella fornita).
2. `nuovi_FLUSSI` — conteggio rapido T1/T2 per azienda + colonna netto scarti.
3. `scarti_nuovi_FLUSSI` — scarti dedup globale per id_rec.
4. `delta_validi` — inviati - scartati = validi per T1, T2 eventi, T2 id_rec, T2 wrapper.
5. `scarti_per_trimestre` — KO/OK/Totale per Azienda x Trimestre x Tracciato.
6. `proiezione_post_upload` — NSIS attuale + veri nuovi = post-upload (chiavi § 4.3).

### Dipendenza aggiuntiva

- `xlrd>=1.2,<3` (lettura report scarti `.xls`).

## Note operative

- Su macOS serve un Python con supporto `tkinter`.
- Il Python Homebrew trovato in questa sessione non aveva `_tkinter`.
- Se si usa un Python senza `tkinter`, la GUI non parte.

## Stato verifiche

Verificato sul dataset reale presente nella cartella di lavoro:

- parsing e classificazione XML ok;
- generazione workbook ok;
- sheet presenti:
  `Report`, `Dettaglio`, `Dettaglio_201`, `Dettaglio_202`, `Dettaglio_203`, `Dettaglio_204`, `Dettaglio_205`, `CF_Univoci`, `CF_Esclusi`.

Ultimo asset generato in workspace:

- `report_siad_6.xlsx`

## Promemoria per spostamento repo

Quando sposti il progetto nella cartella definitiva:

- porta con te almeno:
  `siad_report_gui.py`, `README.md`, `memoria.md`, `requirements.txt`, `requirements-dev.txt`, `siad_report_gui.spec`, `.github/workflows/release.yml`
- porta con te anche `.gitignore`
- evita di versionare:
  `.venv`, `.venv311`, `__pycache__`, file Excel temporanei e file `~$...`
