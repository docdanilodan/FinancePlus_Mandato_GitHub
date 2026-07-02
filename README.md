# FinancePlus Mandato - GitHub Streamlit OK

Pacchetto corretto per pubblicare **FinancePlus Mandato** su GitHub e Streamlit Cloud.

## File principali

| File/cartella | Uso |
|---|---|
| `streamlit_app.py` | Web app corretta per Streamlit Cloud |
| `Mandato_aggiornato_report.py` | Wrapper compatibile Streamlit, senza tkinter |
| `desktop/Mandato_aggiornato_report_DESKTOP_TKINTER.py` | Versione desktop originale Tkinter |
| `assets/Mandato_vuoto.docx` | Modello Word mandato |
| `assets/LOGO_FINANCE_2.PNG` | Logo FinancePlus |
| `requirements.txt` | Librerie Python |
| `packages.txt` | LibreOffice per conversione PDF su Cloud |
| `runtime.txt` | Python 3.11 per massima compatibilità |

## Deploy Streamlit Cloud

Main file path consigliato:

```text
streamlit_app.py
```

La causa dell'errore precedente era l'import di `tkinter` dentro il file desktop.
Questa versione web non importa `tkinter`.

## Funzioni incluse

- Upload visura camerale o report PDF.
- Estrazione dati società e amministratore.
- Correzione manuale dati.
- Compilazione `Mandato_vuoto.docx`.
- Download DOCX.
- Download PDF se LibreOffice è disponibile.
- Salvataggio anagrafica JSON.

## Avvio locale

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

