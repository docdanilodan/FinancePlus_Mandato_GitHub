# -*- coding: utf-8 -*-
"""Wrapper Streamlit Cloud OK.

Il vecchio file desktop Tkinter e' in:
    desktop/Mandato_aggiornato_report_DESKTOP_TKINTER.py

Questo file resta con lo stesso nome per evitare l'errore su Streamlit Cloud
quando il main file path era impostato su Mandato_aggiornato_report.py.
"""

from streamlit_app import main

if __name__ == "__main__":
    main()
