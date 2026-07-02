# Correzione errore Streamlit Cloud - FinancePlus Mandato

## Errore visto nello screenshot

Streamlit Cloud mostra:

```text
ImportError: This app has encountered an error
File "Mandato_aggiornato_report.py"
import tkinter as tk
```

## Perché succede

`tkinter` serve per programmi desktop Windows/Mac/Linux con finestre grafiche locali.
Streamlit Cloud invece è un server web Linux senza interfaccia desktop Tkinter.
Quindi un file `.py` desktop non può essere usato direttamente come app Streamlit.

## Cosa è stato corretto

In questo pacchetto:

- `streamlit_app.py` è la web app corretta per Streamlit Cloud;
- `Mandato_aggiornato_report.py` è diventato un wrapper web sicuro, senza tkinter;
- il vecchio programma desktop è stato spostato in `desktop/Mandato_aggiornato_report_DESKTOP_TKINTER.py`;
- `assets/Mandato_vuoto.docx` contiene il modello Word;
- `assets/LOGO_FINANCE_2.PNG` contiene il logo FinancePlus;
- `requirements.txt`, `runtime.txt` e `packages.txt` sono pronti per GitHub/Streamlit.

## Deploy corretto su Streamlit Cloud

1. Carica tutto su GitHub.
2. Vai su Streamlit Cloud.
3. Clicca **New app**.
4. Scegli repository e branch.
5. In **Main file path** scrivi:

```text
streamlit_app.py
```

6. Premi **Deploy**.

Se per errore lasci come main file path `Mandato_aggiornato_report.py`, ora non dovrebbe più comparire l'errore `tkinter`, perché quel file è stato trasformato in wrapper web.

## Uso locale Windows

Per app web locale:

```bat
run_streamlit_locale.bat
```

Per app desktop originale:

```bat
run_desktop_locale.bat
```

