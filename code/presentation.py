import streamlit as st
from main_script import budget
from main_script import depense
from main_script import budget_disponible
from main_script import nombre_mission
from main_script import date_expire
from numpy import *
from pandas import *
from openpyxl import *
import glob2
from os.path import splitext, basename



def saut_ligne(i):
    for i in range(int(i)):
        st.write("")
    return
def main():
    st.write('<img class="Logo-etat" src="https://www.inrae.fr/themes/custom/inrae_socle/public/images/etat_logo.svg" alt="République française" width="138" height="146">',
             '<img class="Logo-site" src="https://www.inrae.fr/themes/custom/inrae_socle/logo.svg" alt="INRAE">',
             unsafe_allow_html=True)
    '''
    st.header(":blue[Application d'aide à la lecture des tableaux Excels] 🌻")
    st.write("Cet outil a pour but d'aider à la lecture de tableau excel trop volumineux pour être étudier à la main!",)
    st.write("Attention cette application a été concue à partir de la page 2 : 'ps(12)' du fichier Excel : 'BPP Octobre 2022' !")
    '''
    st.write("Si vous rencontrez une erreur rafraîchissez la page !")
    saut_ligne(3)   
    st.title("Quelques chiffre clés")
    st.write("Le budget total est de", budget, "dont", depense, "ont deja ete depense")
    st.write("Le budget total restant à répartir est:", budget_disponible)
    st.write("Pour un nombre total de missions:", nombre_mission) 
    saut_ligne(3)
    st.title("Mission bientôt expirée")
    date_expire()
    #st.write("Will be soon delivered")
    saut_ligne(5)
    st.markdown(
        """
        ### Contact
        Si vous observer quelconques bugs ou avez des idées d'améliorations, contactez moi via mail :   -->  📧 lemoinedamien21@gmail.com 
    """ 
    )
    
    return 
