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
- Classifica i file con parole chiave, senza AI
- Estrae i campi con regex e anchor text
- Permette analisi incrementale: aggiungi un documento, poi un altro, poi salvi la revisione e continui sullo stesso dossier
- Mostra i dati estratti in una schermata di revisione, modificabili prima del prefill
- Aggrega i dati trovati e precompila Allegato E e Allegato C
- Include parser dedicati per cedolino NoiPA/MEF e contratto di finanziamento/delega in PDF testuale
- Usa Tesseract OCR in locale come fallback sui PDF scannerizzati e come parser principale per immagini
- Riconosce anche pattern sintetici OCR tipo `300 euro x 120 mesi` come possibile `rata` + `durata`, con montante derivato automaticamente

Nota: per OCR serve `tesseract.exe` installato sul PC. Se non e in `PATH`, puoi indicarlo con la variabile d'ambiente `QUINTOQUOTE_TESSERACT_PATH`.

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

---

## Uso da CLI

CLI guidata:

```bash
python quintoquote.py
```

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

---

## Roadmap

- [ ] Refactor in moduli separati
- [ ] Template PDF aggiuntivi
- [ ] Export scenari comparativo
- [ ] Configurazioni branding avanzate
- [ ] Packaging per distribuzione standalone

---

## Note legali

QuintoQuote e un progetto open source pensato per aiutare agenti, collaboratori e realta non strutturate a preparare preventivi PDF chiari, ordinati e professionali. E uno strumento di supporto operativo e non sostituisce mai la documentazione bancaria ufficiale, precontrattuale o contrattuale, ne le verifiche, le delibere o le condizioni definitive del soggetto finanziatore. I valori economici riportati nei PDF generati hanno finalita esclusivamente illustrativa e devono essere verificati sul portale ufficiale dell'istituto mandante prima di qualsiasi utilizzo commerciale.

## Licenza

MIT
