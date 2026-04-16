# TEST

Cartella di prova con dati completamente fittizi per verificare cosa genera il compilatore MEF.

Contenuto:

- `input/pratica_fittizia.json`: dati base della pratica fake
- `values/*.json`: valori effettivamente usati per ciascun modulo
- `output/*.pdf`: PDF generati
- `output/RIEPILOGO.md`: riepilogo rapido dei file creati
- `genera_documenti_fittizi.py`: script per rigenerare tutto

Rigenerazione locale:

```powershell
cd C:\Users\pramp\Downloads\QuintoQuote
.\.venv\Scripts\python.exe .\TEST\genera_documenti_fittizi.py
```

I dati sono volutamente falsi e servono solo come smoke test funzionale dei moduli:

- Allegato E
- Allegato C
- Frontespizio Banche / Finanziarie MEF
- Frontespizio Integrativo Banche / Finanziarie MEF
