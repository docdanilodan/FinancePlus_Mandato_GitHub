# FinancePlus Mandato

Gestionale desktop offline per generazione mandato FinancePlus, compilazione anagrafica cliente da visura/report PDF e produzione DOCX/PDF.

## Funzioni principali

- Dashboard desktop Tkinter in locale.
- Database SQLite locale.
- Archivio automatico in `FinancePlus_Data/`.
- Import visura camerale PDF.
- Import report PDF tipo Creditime.
- Compilazione automatica di societa', amministratore, sede, PEC, P.IVA/C.F.
- Generazione mandato DOCX da modello originale.
- Conversione PDF tramite Microsoft Word o LibreOffice/soffice, se disponibile.
- Asset minimi incorporati: il programma puo' creare modello Word e logo se non presenti nella cartella.

## Struttura repository

```text
FinancePlus_Mandato_GitHub/
├── Mandato_aggiornato_report.py
├── requirements.txt
├── README.md
├── .gitignore
├── run_windows.bat
└── run_mac_linux.sh
```

## Requisiti

- Python 3.10 o superiore.
- Windows consigliato per uso desktop.
- Microsoft Word o LibreOffice consigliati per esportare il DOCX in PDF mantenendo il layout.

Dipendenze Python:

```bash
pip install -r requirements.txt
```

## Avvio su Windows

Metodo semplice:

```bat
run_windows.bat
```

Metodo manuale:

```bat
python -m pip install -r requirements.txt
python Mandato_aggiornato_report.py
```

## Avvio su Mac/Linux

```bash
chmod +x run_mac_linux.sh
./run_mac_linux.sh
```

Oppure:

```bash
python3 -m pip install -r requirements.txt
python3 Mandato_aggiornato_report.py
```

## Caricamento su GitHub

1. Crea un nuovo repository GitHub, ad esempio `financeplus-mandato`.
2. Carica tutti i file contenuti in questa cartella.
3. Non caricare `FinancePlus_Data/`: contiene dati locali, documenti clienti, mandati generati e database.
4. Dopo il clone su un altro PC, esegui `run_windows.bat`.

Comandi Git da terminale:

```bash
git init
git add .
git commit -m "Prima versione FinancePlus Mandato"
git branch -M main
git remote add origin https://github.com/TUO-UTENTE/financeplus-mandato.git
git push -u origin main
```

## Note operative importanti

- `FinancePlus_Data/` viene creata automaticamente all'avvio.
- Il database locale viene creato automaticamente.
- Per dati reali di clienti, usare repository GitHub privato.
- Per esportare in PDF con layout identico al DOCX, installare LibreOffice oppure usare un PC con Microsoft Word.

## Sicurezza dati

Questo progetto lavora in locale. Non pubblicare su GitHub file contenenti:

- documenti clienti;
- visure camerali reali;
- report bancari o Centrale Rischi;
- database `financeplus_mandato.db`;
- cartella `FinancePlus_Data/`.
