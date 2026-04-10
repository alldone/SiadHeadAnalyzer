# SiadHeadAnalyzer

Suite desktop Python con piu verticali operative: `SIAD`, `FAR D33Za`, `Specialistica`, `Mobilita Farmaci` e `SIND Detenuti`.

## Verticali nel repo

Il repository ospita piu verticali distinti, orchestrati da un launcher unico:

- `main_gui.py`
  launcher principale con ribbon e tab `Home`, `SIAD`, `FAR D33Za`, `Specialistica`, `Mobilita Farmaci`, `SIND Detenuti`

- `siad_report_gui.py`
  verticale desktop standalone `SIAD`
- `far_d33za_gui.py`
  verticale desktop standalone `FAR D33Za`
- `specialistica_gui.py`
  verticale desktop standalone `Specialistica`
- `mobilita_gui.py`
  verticale desktop standalone `Mobilita Farmaci`

I verticali separati vivono nelle rispettive cartelle dedicate e condividono solo il launcher principale.

## Verticale FAR D33Za

Il verticale `FAR D33Za` legge ricorsivamente i file `FAR*.xml` e calcola l'indicatore NSG:

- assistiti residenti con eta `>=75` anni;
- trattamenti residenziali `R1`, `R2`, `R3`;
- conteggio una sola volta nell'anno per paziente;
- attribuzione del paziente al livello di intensita piu elevato, come previsto dalla scheda `D33Za`.
- uso dei `txt` di anagrafe assistiti come fonte del denominatore e come lookup per residenza/eta, soprattutto sul `Tracciato 2`.

Il report Excel prodotto contiene:

- `D33Za_Riepilogo`
  riepilogo per `ASL Residenza` con popolazione residente `>=75`, numeratori e indicatori `x1000`
- `Dettaglio`
  tutte le righe FAR valutate, con motivo di inclusione o esclusione
- `Assistiti_Selezionati`
  un record finale per ogni assistito conteggiato nell'indicatore
- `Popolazione_Anagrafe`
  denominatore `>=75` per azienda sanitaria ricavato dai `txt`
- `Verifica_Tracciati`
  conteggi di controllo per `Tracciato 1` e `Tracciato 2`

Nota operativa:

- il `Tracciato 1` usa i dati FAR e, se disponibili, li integra con anagrafe;
- il `Tracciato 2` ricava residenza ed eta tramite il collegamento tra `ID_REC` e codice fiscale anagrafico;
- il denominatore dell'indicatore e' la popolazione residente `>=75` ricavata dai file `Export_Assistiti_*.txt`.

## Verticale Mobilita Farmaci

Il verticale `Mobilita Farmaci` lavora su due cartelle selezionate separatamente: `attiva` e `passiva`.

Per ogni azienda erogatrice, ricavata dal nome file (`201 importi.csv`, `201 prestazioni.csv`, ecc.), il programma:

- abbina `importi` e `prestazioni` dentro ciascuna delle due cartelle;
- riconosce i file in modo intelligente:
  parte sempre dal codice azienda iniziale nel nome file, poi deduce `importi` o `prestazioni`
  dai token del filename e, se necessario, dal contenuto numerico del CSV;
- preserva nel report la prima colonna sorgente (`ASL` per attiva, `AZIENDA SANITARIA CREDITRICE` per passiva);
- permette di scegliere da listbox le tipologie da riportare negli sheet;
- seleziona di default `FARMACEUTICA` e `SOMM. DIRETTA DI FARMACI`;
- interpreta i numeri nel formato italiano (`.` migliaia, `,` decimali);
- genera un workbook Excel con due fogli: `attiva` e `passiva`;
- esegue sempre un controllo automatico degli importi tra input CSV e output Excel.

Colonne del report:

- `AZIENDA EROGATRICE`
- `AZIENDA SANITARIA DEBITRICE` per il foglio `attiva`
- `AZIENDA SANITARIA CREDITRICE` per il foglio `passiva`
- `FARMACEUTICA - IMPORTI`
- `FARMACEUTICA - PRESTAZIONI`
- `SOMM. DIRETTA DI FARMACI - IMPORTI`
- `SOMM. DIRETTA DI FARMACI - PRESTAZIONI`

Ogni foglio include anche:

- una riga di `SUBTOTALE` dopo ogni blocco di azienda erogatrice;
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

Ordinamento aziende erogatrici:

- `201`, `202`, `203`, `204`, `205`, `912`, `914`, `915`, `916`

## Cosa fa

- legge i due `XSD` dei tracciati SIAD;
- cerca ricorsivamente i file `SIAD*.xml` in una cartella selezionata;
- riconosce automaticamente `tracciato 1` e `tracciato 2`;
- estrae `CodiceASL`, `Id_Rec`, `AnnoNascita` e `Codice Fiscale`;
- calcola eta' e ambiguita' dei `CF`;
- applica regole di conteggio per azienda;
- evidenzia i `CF` condivisi tra aziende e separa le statistiche in blocchi coerenti;
- genera un report Excel multi-sheet;
- mostra in GUI sia l'elenco dei file letti sia il riepilogo aggregato;
- consente di validare un XML selezionato rispetto allo XSD del relativo tracciato.

## Punti chiave

- Unicita' del `CF` calcolata per azienda `CodiceASL`
- Riepilogo organizzato in tre gruppi di statistiche:
  `CF per azienda`, `Teste singole globali`, `Differenze`
- Gestione dei `CF` ambigui con supporto a `AnnoNascita`
- Evidenza separata delle prese in carico da `tracciato 2` nell'anno
- Foglio dedicato ai `CF` univoci per azienda
- Foglio dedicato ai `CF` esclusi dai conteggi
- Output Excel con formattazione leggibile, header curati e righe alternate

## Output generato

Il file Excel contiene:

- `Report`
  riepilogo per azienda con totali finali
- `Dettaglio`
  tutte le righe lette, con motivazione di inclusione o esclusione e flag `IsOver65`
- `Dettaglio_<ASL>`
  un foglio per ogni azienda
- `CF_Univoci`
  elenco dei `CF` univoci per azienda con trimestri e occorrenze
- `CF_Esclusi`
  elenco dei `CF` distinti per azienda non utilizzati nei conteggi

## Interfaccia

La GUI consente di selezionare:

- `TRACCIATO 1 XSD`
- `TRACCIATO 2 XSD`
- cartella sorgente XML
- file Excel di output
- anno di analisi

La finestra include tre tab:

- `File XML`
  elenco dei file intercettati
- `Riepilogo`
  vista tabellare del summary prodotto
- `Validazione`
  elenco degli errori XSD dell'XML selezionato

Quando si clicca su un file nella tab `File XML`, l'app propone la validazione con lo `XSD` del tracciato corrispondente:

- se il file e' conforme, mostra un messaggio di esito positivo;
- se il file non e' conforme, mostra un messaggio di errore e popola la tab `Validazione` con il dettaglio.

## Regole di conteggio

- `Tracciato 1`
  nuove prese in carico dell'anno di analisi
- `Tracciato 2` con data antecedente al `01/01/anno`
  prese precedenti ancora attive
- `Tracciato 2` con data nell'anno
  conteggiato in colonna separata solo se il `CF` non e' gia' presente tra i `CF` univoci conteggiati per la stessa azienda
- Deduplica prese in carico per `Tracciato + CodiceASL + Id_Rec`
- Nel riepilogo le statistiche dei `CF` sono presentate in tre gruppi con lo stesso set di metriche:
  totale, `>= 65 anni`, `cf ambigui`

## Avvio locale

### macOS

```bash
cd /percorso/progetto
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python main_gui.py
```

### Windows

```powershell
cd C:\percorso\progetto
py -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python main_gui.py
```

Nota:

- Su macOS serve una build Python con supporto `tkinter`

## Build eseguibile

Installa le dipendenze di build:

```bash
python -m pip install -r requirements.txt -r requirements-dev.txt
```

Poi genera l'app con `PyInstaller`:

```bash
pyinstaller --noconfirm siad_report_gui.spec
```

Output:

- `dist/SiadHeadAnalyzer/`

## Release automatiche su GitHub

Il repository include la workflow:

- `.github/workflows/release.yml`

La workflow:

- builda su `macos-15-intel` per macOS Intel
- builda su `macos-latest` per macOS Apple Silicon
- builda su `windows-latest`
- crea un archivio `.zip`
- pubblica gli asset nella release GitHub quando viene pushato un tag `v*`

Esempio:

```bash
git tag v1.0.0
git push origin v1.0.0
```

## Struttura repository

- `main_gui.py`
  launcher principale della suite
- `siad_report_gui.py`
  verticale standalone `SIAD`
- `far_d33za_gui.py`
  verticale standalone `FAR D33Za`
- `specialistica_gui.py`
  verticale standalone `Specialistica`
- `mobilita_gui.py`
  verticale standalone `Mobilita Farmaci`
- `far_verticale/`
  codice del verticale `FAR D33Za`
- `mobilita_verticale/`
  codice del verticale `Mobilita Farmaci`
- `sind_verticale/`
  codice del verticale `SIND Detenuti`
- `requirements.txt`
  dipendenze runtime
- `requirements-dev.txt`
  dipendenze build
- `siad_report_gui.spec`
  configurazione PyInstaller
- `.github/workflows/release.yml`
  pipeline GitHub Actions
- `memoria.md`
  memoria operativa del progetto

## Roadmap possibile

- firma e notarizzazione macOS
- installer Windows
- miglioramento export con filtri Excel e freeze pane
- supporto a configurazioni di regole esterne

## Licenza

Questo progetto e' distribuito sotto licenza MIT. Vedi [LICENSE](/Users/aldo/Desktop/PROGETTI/python/SiadHeadAnalyzer/LICENSE).
