import streamlit as st
st.set_page_config(layout="wide")
import Table_complete
import Mission_par_Ligne_analytique
import Mission_par_Scientifique
import presentation
PAGES = {
    " Accueil": presentation,
    "Traité par Ligne analytique": Mission_par_Ligne_analytique,
    "Traité par Scientifique": Mission_par_Scientifique,
    "Table étudiée": Table_complete
    }

st.sidebar.header('Liste des outils :gear:', )
selection = st.sidebar.radio("  ", list(PAGES.keys()))
page = PAGES[selection]
page.main()