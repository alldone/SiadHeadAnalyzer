# Report SIAD

Applicazione desktop Python per analizzare i flussi `SIAD` a partire dai due tracciati XML/XSD, costruire riepiloghi per azienda e produrre un report Excel strutturato, leggibile e pronto per la verifica operativa.

## Cosa fa

- legge i due `XSD` dei tracciati SIAD;
- cerca ricorsivamente i file `SIAD*.xml` in una cartella selezionata;
- riconosce automaticamente `tracciato 1` e `tracciato 2`;
- estrae `CodiceASL`, `Id_Rec`, `AnnoNascita` e `Codice Fiscale`;
- calcola eta' e ambiguita' dei `CF`;
- applica regole di conteggio per azienda;
- evidenzia i `CF` condivisi tra aziende e il totale delle teste singole globali;
- genera un report Excel multi-sheet;
- mostra in GUI sia l'elenco dei file letti sia il riepilogo aggregato.

## Punti chiave

- Unicita' del `CF` calcolata per azienda `CodiceASL`
- Evidenza dei `CF` condivisi tra aziende
- Contatore finale delle teste singole globali
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
  tutte le righe lette, con motivazione di inclusione o esclusione
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

La finestra include due tab:

- `File XML`
  elenco dei file intercettati
- `Riepilogo`
  vista tabellare del summary prodotto

## Regole di conteggio

- `Tracciato 1`
  nuove prese in carico dell'anno di analisi
- `Tracciato 2` con data antecedente al `01/01/anno`
  prese precedenti ancora attive
- `Tracciato 2` con data nell'anno
  conteggiato in colonna separata solo se il `CF` non e' gia' presente tra i `CF` univoci conteggiati per la stessa azienda
- Deduplica prese in carico per `Tracciato + CodiceASL + Id_Rec`
- Nel riepilogo vengono indicati anche i `CF` condivisi con altre aziende e, sulla riga totale, il numero di teste singole globali

## Avvio locale

### macOS

```bash
cd /percorso/progetto
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python siad_report_gui.py
```

### Windows

```powershell
cd C:\percorso\progetto
py -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python siad_report_gui.py
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

- `dist/ReportSIAD/`

## Release automatiche su GitHub

Il repository include la workflow:

- `.github/workflows/release.yml`

La workflow:

- builda su `macos-latest`
- builda su `windows-latest`
- crea un archivio `.zip`
- pubblica gli asset nella release GitHub quando viene pushato un tag `v*`

Esempio:

```bash
git tag v1.0.0
git push origin v1.0.0
```

## Struttura repository

- `siad_report_gui.py`
  applicazione principale
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

Questo progetto e' distribuito sotto licenza MIT. Vedi [LICENSE](/Users/aldo/Desktop/PROGETTI/gisella/siad/rework/LICENSE).
