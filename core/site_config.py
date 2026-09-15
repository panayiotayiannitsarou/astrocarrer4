"""
site_config.py
================
Το AstroCheck Analysis, το AstroCheck Career και η κεντρική σελίδα (Home)
γίνονται deploy ως ΤΡΙΑ ΞΕΧΩΡΙΣΤΑ Streamlit apps (τρία ξεχωριστά URLs), μέσα
όμως από το ΙΔΙΟ GitHub repo -- καθένα δείχνει σε διαφορετικό entry-point
αρχείο (Home.py / AstroCheck_Analysis.py / AstroCheck_Career.py) στις
ρυθμίσεις "Main file path" του Streamlit Community Cloud.

Επειδή είναι πραγματικά ξεχωριστά apps (όχι σελίδες μέσα στο ίδιο app),
δεν μπορούν να χρησιμοποιήσουν st.page_link μεταξύ τους -- χρειάζονται
απλά συνδέσμους (URL). Τα πραγματικά URLs δεν υπάρχουν πριν γίνει το πρώτο
deploy, οπότε μπαίνουν εδώ ΜΙΑ φορά, σε ένα σημείο, μετά το πρώτο deploy.

Πώς να το συμπληρώσεις:
1. Κάνε deploy και τα τρία apps στο Streamlit Community Cloud (ή όπου
   επιλέξεις), ένα-ένα, με "Main file path" αντίστοιχα Home.py,
   AstroCheck_Analysis.py, AstroCheck_Career.py.
2. Πάρε τα 3 URLs που θα σου δώσει το Streamlit Cloud.
3. Βάλε τα εδώ κάτω -- ή, προτιμότερο, ως Streamlit "Secrets" με τα ίδια
   ονόματα (HOME_URL, ANALYSIS_URL, CAREER_URL) στις ρυθμίσεις κάθε app,
   ώστε να μην χρειάζεται να ξανακάνεις commit για να αλλάξεις ένα URL.
"""
from __future__ import annotations

try:
    import streamlit as st
    _secrets = st.secrets
except Exception:
    _secrets = {}


def _url(key: str, fallback: str) -> str:
    try:
        value = _secrets.get(key)
    except Exception:
        value = None
    return value or fallback


# Placeholders -- αντικατέστησέ τα με τα πραγματικά URLs μετά το πρώτο deploy,
# ή (προτιμότερο) όρισέ τα ως Secrets με το ίδιο όνομα.
HOME_URL = _url("HOME_URL", "https://astrocheck-home.streamlit.app")
ANALYSIS_URL = _url("ANALYSIS_URL", "https://astrocheck-analysis.streamlit.app")
CAREER_URL = _url("CAREER_URL", "https://astrocheck-career.streamlit.app")
