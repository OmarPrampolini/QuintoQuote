# QuintoQuote

Generatore di preventivi PDF professionali per Cessione del Quinto e Delega di Pagamento.

QuintoQuote trasforma i dati di una simulazione in un PDF commerciale pronto da condividere con il cliente. Include:

- web UI locale con anteprima live
- CLI guidata
- CLI non interattiva per automazioni o batch
- branding agente completo
- storico dei PDF generati

## AVVIO IMMEDIATO POWERSHELL

### Copia E Incolla Questo

Prima installazione:

```powershell
Set-Location 'C:\percorso\alla\cartella\QuintoQuote'
pip install -e .
```

Sostituisci `C:\percorso\alla\cartella\QuintoQuote` con la cartella reale del progetto sul tuo PC.

Avvio normale da quel momento in poi:

```powershell
Set-Location 'C:\percorso\alla\cartella\QuintoQuote'
QuintoQuote start
```

Se preferisci `cd` invece di `Set-Location`:

```powershell
cd 'C:\percorso\alla\cartella\QuintoQuote'
QuintoQuote start
```

### CLICCA QUI PER L'APPLICAZIONE

Dopo `QuintoQuote start`, QuintoQuote prova ad aprire il browser da solo.
Se vuoi aprirla a mano, clicca questo link locale:

[http://127.0.0.1:5000](http://127.0.0.1:5000)

Con `start`, QuintoQuote:

- avvia la web UI locale
- prova ad aprire automaticamente il browser
- stampa nel terminale `Clicca qui per l'applicazione: http://127.0.0.1:5000`
- si chiude con `CTRL+C`

### Se `QuintoQuote` Non Viene Riconosciuto

Esegui una sola volta:

```powershell
Set-Location 'C:\percorso\alla\cartella\QuintoQuote'
pip install -e .
```

Se vuoi un fallback immediato senza installare il comando:

```powershell
Set-Location 'C:\percorso\alla\cartella\QuintoQuote'
python .\quintoquote.py start
```

## IL PDF GENERATO PUO INCLUDERE IL BOLLINO OAM CLICCABILE E IL PROFILO AGENTE

> [!IMPORTANT]
> **Il PDF puo mostrare il profilo agente e aggiungere il bollino OAM cliccabile.**
> Per attivarlo vai in **Impostazioni > Profilo Agente (Branding)**, compila:
> **Nome Agente**, **Rete Mandante**, **Codice OAM**, **Telefono**, colori e, nel blocco **Immagini**, carica il file nel campo **Bollino OAM**.

> [!TIP]
> Quando carichi il bollino OAM nelle Impostazioni, QuintoQuote lo inserisce nel PDF e lo rende cliccabile verso [organismo-am.it](https://www.organismo-am.it/).
> Lo stesso profilo agente viene usato anche nella testata del PDF.

## Screenshot

### Nuovo Preventivo

![Nuovo preventivo](docs/screenshots/ui-nuovo-preventivo.png)

### Profilo Agente E Branding

Qui configuri esattamente i dati che finiscono nel PDF: nome agente, rete mandante, codice OAM, telefono, colori, bollino OAM e logo.

![Profilo agente e branding](docs/screenshots/ui-impostazioni-profilo.png)

### Storico PDF

![Storico preventivi](docs/screenshots/ui-storico.png)

### PDF Di Esempio

| Pagina 1 | Pagina 2 |
| --- | --- |
| ![PDF esempio pagina 1](docs/screenshots/pdf-esempio-pagina-1.png) | ![PDF esempio pagina 2](docs/screenshots/pdf-esempio-pagina-2.png) |

## Cosa Personalizzi

- Profilo agente: nome, rete mandante, codice OAM, telefono
- Colori del PDF: primario e accento
- Bollino OAM: JPG/PNG caricato dalla pagina Impostazioni
- Logo agente: opzionale
- Testi finali: disclaimer, note e closing modificabili anche dall'anteprima

## Come Attivare Profilo Agente E Bollino OAM

1. Avvia la web UI.
2. Apri `Impostazioni`.
3. Compila `Nome Agente`, `Rete Mandante`, `Codice OAM` e `Telefono`.
4. Carica il file nel campo `Bollino OAM`.
5. Facoltativo: carica anche `Logo Agente`.
6. Salva.
7. Torna su `Nuovo`, compila il preventivo e scarica il PDF.

Risultato:

- il profilo agente compare nella testata del PDF
- il bollino OAM viene inserito nel PDF
- il bollino OAM e cliccabile

## Funzionalita

- PDF professionale con hero card rata/TAEG, timeline e tabella economica
- Profilo agente e branding configurabile
- Bollino OAM cliccabile nel PDF
- Multi-scenario: piu opzioni nello stesso PDF, una pagina per scenario
- Anteprima live nel browser durante la compilazione
- Editing inline di disclaimer, note e closing prima del download
- Storico dei PDF generati con download diretto

## Requisiti

- Python 3.10+
- `reportlab`
- `flask` per la modalita web

## Installazione

Installazione veloce:

```bash
pip install -r requirements.txt
```

Installazione come comando locale:

```bash
pip install -e .
```

Dopo `pip install -e .` puoi usare i comandi `QuintoQuote` e `quintoquote`.

## Avvio Rapido

Comando consigliato:

```bash
QuintoQuote start
```

oppure:

```bash
quintoquote start
```

fallback senza installazione del comando:

```bash
python quintoquote.py start
```

Con `start`, QuintoQuote prova anche ad aprire il browser da solo.
Se serve, apri manualmente `http://127.0.0.1:5000`.

Per tenere separata questa istanza da un altro progetto simile o personale, usa porta e file locali dedicati:

```bash
QuintoQuote start \
  --port 5010 \
  --config-path ./.demo/config-public.json \
  --assets-dir ./.demo/assets-public
```

In questo modo:

- non riusi per errore la stessa porta di un'altra istanza
- non condividi il `config.json` locale con un altro progetto
- non condividi la cartella `assets/` degli upload

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

Multi-scenario via CLI:

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

Formato scenario:

```text
rata;durata_mesi;tan;taeg;importo_erogato[;tipo_finanziamento]
```

## File Locali Generati

Questi file o cartelle sono locali e non vanno pubblicati nel repository:

- `config.json`
- `assets/`
- `output_preventivi/`
- eventuali percorsi personalizzati passati a `--config-path` o `--assets-dir`

Sono gia inclusi in `.gitignore`.

## Struttura Del Repository

```text
quintoquote.py              # entrypoint pubblico della CLI
preventivo_generator_v2.py  # core applicativo (CLI + web + PDF)
requirements.txt            # dipendenze runtime
pyproject.toml              # metadati progetto e script "quintoquote"
docs/screenshots/           # screenshot usati nel README
LICENSE
README.md
```

## Note Legali

I documenti generati hanno finalita illustrativa e non sostituiscono il SECCI ne la documentazione precontrattuale ufficiale del finanziatore. I valori economici devono essere verificati sul portale ufficiale dell'istituto mandante prima di qualsiasi utilizzo commerciale.

## Licenza

MIT
