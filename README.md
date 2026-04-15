<p align="center">
  <strong style="font-size:2em">QuintoQuote</strong>
</p>

<p align="center">
  Generatore open source di preventivi PDF professionali<br>
  per <strong>Cessione del Quinto</strong> e <strong>Delega di Pagamento</strong>.
</p>

<p align="center">
  <strong>Locale. Veloce. Personalizzabile. Pronto da inviare al cliente.</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/python-3.10%2B-blue?logo=python&logoColor=white" alt="Python 3.10+"/>
  <img src="https://img.shields.io/badge/license-MIT-green" alt="MIT License"/>
  <img src="https://img.shields.io/badge/local--first-nessun%20SaaS-orange" alt="Local-first"/>
  <img src="https://img.shields.io/badge/open%20source-%E2%9C%93-brightgreen" alt="Open Source"/>
</p>

<p align="center">
  Web UI locale &nbsp;&middot;&nbsp; CLI &nbsp;&middot;&nbsp; Branding agente &nbsp;&middot;&nbsp; Bollino OAM cliccabile &nbsp;&middot;&nbsp; Multi-scenario
</p>

<p align="center">
  <img src="docs/screenshots/demo.gif" alt="QuintoQuote demo" width="720"/>
</p>

> *QuintoQuote e un progetto open source pensato per aiutare agenti, collaboratori e realta non strutturate a preparare preventivi PDF chiari, ordinati e professionali. E uno strumento di supporto operativo e non sostituisce mai la documentazione bancaria ufficiale, precontrattuale o contrattuale, ne le verifiche, le delibere o le condizioni definitive del soggetto finanziatore.*

---

## Provalo in 30 secondi

```powershell
irm https://raw.githubusercontent.com/OmarPrampolini/QuintoQuote/main/install.ps1 | iex
```

Lo script crea `.venv`, installa QuintoQuote, abilita il launcher locale e avvia la web UI.
Compili il form. Scarichi il PDF. Fine.

---

## Build EXE Windows

Per generare il pacchetto desktop Windows dal repository locale:

```powershell
cd C:\Users\pramp\Downloads\QuintoQuote
.\build_exe.ps1 -Clean
```

Output:
- `dist\QuintoQuote\QuintoQuote.exe`
- `dist\QuintoQuote-portable.zip`

La build include:
- eseguibile desktop con avvio diretto della web UI
- template PDF presenti in `docs/`
- bundle OCR locale in `dist\QuintoQuote\ocr\`
- salvataggio dati utente in `%LOCALAPPDATA%\QuintoQuote`

---

## Uso EXE Windows

Se vuoi usare QuintoQuote come programma Windows:

1. estrai `QuintoQuote-portable.zip` oppure apri la cartella `dist\QuintoQuote`
2. avvia `QuintoQuote.exe` con doppio click
3. il browser si apre automaticamente sulla web UI locale
4. lavori normalmente da interfaccia web, ma senza dover installare Python

Comportamento del file `.exe`:
- avvia direttamente la UI locale anche senza parametri
- sceglie una porta libera automaticamente se la `5000` e gia occupata
- include OCR locale bundle-friendly
- mostra `Chiudi App` nella barra di navigazione per terminare il programma in modo esplicito

Percorsi usati dalla versione `.exe`:
- configurazione: `%LOCALAPPDATA%\QuintoQuote\config.json`
- immagini utente: `%LOCALAPPDATA%\QuintoQuote\assets\`
- PDF generati: `%LOCALAPPDATA%\QuintoQuote\output_preventivi\`
- file temporanei: `%LOCALAPPDATA%\QuintoQuote\.quintoquote_tmp\`

Consiglio pratico:
- non lanciare l'app da dentro cartelle protette o di sistema
- lascia `QuintoQuote.exe` insieme alla cartella `ocr` e alla cartella `_internal`
- se distribuisci il pacchetto a terzi, distribuisci sempre l'intera cartella o lo ZIP completo

---

## GitHub Actions

Il repository include una pipeline Windows pronta in:

- `.github/workflows/build-windows-exe.yml`

Cosa fa:
- installa Python
- installa le dipendenze del progetto
- installa Tesseract OCR sul runner Windows
- genera `QuintoQuote.exe`
- crea `QuintoQuote-portable.zip`
- carica gli artefatti su GitHub Actions
- se apri una Release GitHub, allega automaticamente lo ZIP alla release

Flusso consigliato:
1. fai push su `main` per ottenere gli artifact nella tab `Actions`
2. crea una Release GitHub quando vuoi pubblicare una build stabile
3. scarica `QuintoQuote-portable.zip` direttamente dalla Release o dagli artifact del workflow

---

## Il risultato

| Pagina 1 | Pagina 2 |
| --- | --- |
| ![PDF esempio pagina 1](docs/screenshots/pdf-esempio-pagina-1.png) | ![PDF esempio pagina 2](docs/screenshots/pdf-esempio-pagina-2.png) |

Questo e il tipo di PDF che puoi inviare subito al cliente: chiaro, elegante, brandizzato e pronto all'uso. Hero card rata/TAEG, timeline finanziaria, tabella economica dettagliata, bollino OAM cliccabile. Tutto con i tuoi colori e il tuo nome.

---

## La web UI

![Nuovo preventivo](docs/screenshots/ui-nuovo-preventivo.png)

Anteprima live a destra, form a sinistra. Ogni campo che modifichi aggiorna l'anteprima in tempo reale. Quando sei soddisfatto, scarichi il PDF.

---

## Perche QuintoQuote

- **Gira in locale sul tuo PC.** Nessun server, nessun cloud, nessun dato che esce.
- **Nessun SaaS, nessun abbonamento.** Installi e usi. Per sempre.
- **PDF professionali gia pronti.** Non servono template, non serve Canva, non serve Word.
- **Branding agente completo.** Nome, rete, OAM, colori, logo, bollino: tutto tuo.
- **Bollino OAM cliccabile.** Chi riceve il PDF clicca e verifica la tua iscrizione.
- **Multi-scenario.** Piu opzioni nello stesso PDF, una pagina per scenario.
- **Adatto a uso reale.** Non e una demo. E quello che usi ogni giorno.

---

## Per chi e

- Agenti in attivita finanziaria
- Collaboratori OAM
- Mediatori creditizi
- Reti vendita e strutture commerciali
- Professionisti che vogliono generare preventivi in locale, senza SaaS

---

## Cosa fa, in pratica

| Funzione | Dettaglio |
| --- | --- |
| PDF premium | Hero card rata/TAEG, KPI, timeline, tabella, note legali |
| Anteprima live | Il preventivo si aggiorna mentre compili il form |
| Editing inline | Modifica disclaimer, note e closing direttamente nell'anteprima |
| Branding agente | Nome, rete mandante, codice OAM, telefono, colori primario/accento |
| Bollino OAM | JPG/PNG caricato dalle Impostazioni, cliccabile nel PDF |
| Logo agente | Opzionale, visibile nell'header del PDF |
| Multi-scenario | Piu combinazioni rata/durata nello stesso documento |
| Storico PDF | Tutti i preventivi generati, scaricabili dalla UI |
| Modulistica MEF | Compilatore guidato per Allegato E e Allegato C Flussi Finanziari |
| Dossier NO AI | Upload PDF, JPG, PNG, OCR locale e prefill moduli |
| CLI guidata | Prompt interattivo da terminale |
| CLI batch | `--non-interactive` per script e automazioni |

---

## Compilatore Allegati MEF

Nella web UI trovi anche la sezione **Moduli**.

- **Allegato E Delega MEF**: 43 campi mappati sulle prime 2 pagine del template originale.
- **Allegato C Flussi Finanziari MEF**: 20 campi mappati sul template Creditonet.
- **Output diretto in PDF**: compili i campi nella UI e scarichi il modulo gia popolato, mantenendo impaginazione e caselle originali.

I template PDF di partenza vengono letti dalla cartella `docs/`.

---

## Dossier Documenti (NO AI)

Nella web UI trovi anche la sezione **Dossier**.

- Accetta piu PDF testuali caricati insieme
- Accetta anche screenshot e scansioni in JPG/PNG tramite OCR locale
- Richiede la scelta della tipologia documento per ogni upload: `Busta paga NoiPA`, `Contratto di finanziamento`, `Carta di identità`, `Tessera sanitaria`
- Usa parser e regole dedicati per la tipologia scelta, senza AI
- Estrae i campi con regex e anchor text
- Permette analisi incrementale: aggiungi un documento, poi un altro, poi salvi la revisione e continui sullo stesso dossier
- Mostra i dati estratti in una schermata di revisione, modificabili prima del prefill
- Aggrega i dati trovati e precompila Allegato E e Allegato C
- Include parser dedicati per cedolino NoiPA/MEF e contratto di finanziamento/delega in PDF testuale
- Usa Tesseract OCR in locale come fallback sui PDF scannerizzati e come parser principale per immagini
- Riconosce anche pattern sintetici OCR tipo `300 euro x 120 mesi` come possibile `rata` + `durata`, con montante derivato automaticamente

Nota: per OCR serve `tesseract.exe`. In sviluppo puo stare nel sistema o in `PATH`; in ottica `.exe` il runtime cerca anche copie locali bundle-friendly vicino all'applicazione.

---

## Configurazione agente

![Profilo agente e branding](docs/screenshots/ui-impostazioni-profilo.png)

1. Avvia QuintoQuote
2. Vai in **Impostazioni**
3. Compila nome, rete, OAM, telefono
4. Scegli i colori
5. Carica bollino OAM e logo
6. Salva

Da quel momento ogni PDF che generi porta il tuo branding.

---

## Storico

![Storico preventivi](docs/screenshots/ui-storico.png)

Tutti i PDF generati restano disponibili per il download.

---

## Installazione

### One-liner GitHub

```powershell
irm https://raw.githubusercontent.com/OmarPrampolini/QuintoQuote/main/install.ps1 | iex
```

Per default installa il progetto in `~/QuintoQuote`, crea `.venv`, aggiorna `pip`, installa il package e avvia QuintoQuote.

### Quando usare installazione Python e quando usare EXE

Usa la versione Python se:
- stai sviluppando il progetto
- vuoi modificare il codice
- vuoi usare la CLI o fare build locali

Usa la versione `.exe` se:
- vuoi semplicemente lavorare
- vuoi distribuire QuintoQuote su PC Windows senza Python
- vuoi un avvio diretto, piu semplice e piu vicino a un programma desktop classico

### Requisiti

- Python 3.10+
- `reportlab`
- `flask`
- `werkzeug`
- `pymupdf`
- `pillow`
- `tesseract` opzionale ma consigliato per scansioni e screenshot

### Setup

```bash
python -m venv .venv
.venv\Scripts\python -m pip install --upgrade pip
.venv\Scripts\python -m pip install -e .
```

Questo isola QuintoQuote dal Python globale e forza una baseline di dipendenze sicure.

Dopo l'installazione puoi avviare QuintoQuote con l'interprete della virtualenv:

```bash
.venv\Scripts\python -m preventivo_generator_v2 start
```

Oppure direttamente con il launcher generato nella virtualenv:

```powershell
.venv\Scripts\quintoquote.exe start
```

Apre la web UI e prova ad aprire il browser automaticamente.
Se serve, vai a mano su [http://127.0.0.1:5000](http://127.0.0.1:5000).

### Avvio su porta diversa

```bash
QuintoQuote start --port 5010
```

### Fallback senza installazione

```bash
python quintoquote.py start
```

### Dove finiscono i file nella versione Python

Per default:
- config: `config.json` nella root del progetto
- assets: cartella `assets/`
- output PDF: cartella `output_preventivi/`
- temporanei OCR: cartella `.quintoquote_tmp/`

---

## Uso da CLI

### Avvio web locale da CLI

```powershell
quintoquote start
```

oppure:

```powershell
.venv\Scripts\python -m preventivo_generator_v2 start
```

Per usare cartelle personalizzate:

```powershell
quintoquote start --port 5010 --config-path C:\QuintoQuote\config.json --assets-dir C:\QuintoQuote\assets --out-dir C:\QuintoQuote\output
```

### CLI guidata

CLI guidata:

```bash
python quintoquote.py
```

### CLI non interattiva

CLI non interattiva:

```bash
python quintoquote.py --non-interactive \
  --cliente "Mario Rossi" \
  --data-nascita "15/05/1975" \
  --tipo-lavoro "Dipendente Statale" \
  --provincia "Milano" \
  --tipo-finanziamento "Cessione del Quinto" \
  --importo-rata 350 \
  --durata-mesi 120 \
  --tan 4.5 \
  --taeg 4.75 \
  --importo-erogato 30000
```

Multi-scenario:

```bash
python quintoquote.py --non-interactive \
  --cliente "Mario Rossi" \
  --data-nascita "15/05/1975" \
  --tipo-lavoro "Dipendente Statale" \
  --provincia "Milano" \
  --importo-rata 350 --durata-mesi 120 --tan 4.5 --taeg 4.75 --importo-erogato 30000 \
  --scenario "300;96;4.2;4.5;25000" \
  --scenario "400;120;4.8;5.1;35000"
```

Formato scenario: `rata;durata_mesi;tan;taeg;importo_erogato[;tipo_finanziamento]`

### Uso pratico consigliato

- per il lavoro quotidiano usa la web UI
- usa la CLI non interattiva quando vuoi integrazione, script o automazioni locali
- usa i `--scenario` se vuoi produrre un unico PDF con piu opzioni di rata/durata
- usa il Dossier dalla web UI quando devi estrarre dati da documenti e precompilare Allegato E/C

### Esempi utili

Avvio su porta diversa:

```powershell
quintoquote start --port 5055
```

Generazione PDF con output personalizzato:

```powershell
python quintoquote.py --non-interactive ^
  --cliente "Mario Rossi" ^
  --data-nascita "15/05/1975" ^
  --tipo-lavoro "Dipendente Statale" ^
  --provincia "Milano" ^
  --tipo-finanziamento "Cessione del Quinto" ^
  --importo-rata 350 ^
  --durata-mesi 120 ^
  --tan 4.5 ^
  --taeg 4.75 ^
  --importo-erogato 30000 ^
  --out-dir "C:\Preventivi"
```

---

## Risoluzione problemi

- Se il browser non si apre, copia a mano l'indirizzo mostrato nella console o nel processo di avvio.
- Se la porta `5000` e occupata, la versione `.exe` prova automaticamente una porta successiva.
- Se l'OCR non legge bene uno screenshot, usa immagini piu nitide e piu grandi.
- Se un documento non viene riconosciuto bene nel Dossier, scegli sempre la tipologia corretta prima dell'upload.
- Se Allegato C non ha campi compilabili nativi, QuintoQuote usa un fallback a coordinate sul modulo.
- Se distribuisci l'app, non separare `QuintoQuote.exe` dalla sua cartella `ocr`.

---

## Roadmap

- [ ] Refactor in moduli separati
- [ ] Template PDF aggiuntivi
- [ ] Export scenari comparativo
- [ ] Configurazioni branding avanzate
- [x] Packaging per distribuzione standalone Windows
- [ ] Firma digitale del binario Windows

---

## Note legali

QuintoQuote e un progetto open source pensato per aiutare agenti, collaboratori e realta non strutturate a preparare preventivi PDF chiari, ordinati e professionali. E uno strumento di supporto operativo e non sostituisce mai la documentazione bancaria ufficiale, precontrattuale o contrattuale, ne le verifiche, le delibere o le condizioni definitive del soggetto finanziatore. I valori economici riportati nei PDF generati hanno finalita esclusivamente illustrativa e devono essere verificati sul portale ufficiale dell'istituto mandante prima di qualsiasi utilizzo commerciale.

## Licenza

MIT
