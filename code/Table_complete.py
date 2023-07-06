import streamlit as st
import streamlit as st
from numpy import *
from pandas import *
from openpyxl import *
import streamlit as st
import glob2
from main_script import table_complete
from os.path import splitext, basename
def saut_ligne(i):
    for i in range(int(i)):
        st.write("")
    return
def main():
    st.write('<img class="Logo-etat" src="https://www.inrae.fr/themes/custom/inrae_socle/public/images/etat_logo.svg" alt="République française" width="138" height="146">',
             '<img class="Logo-site" src="https://www.inrae.fr/themes/custom/inrae_socle/logo.svg" alt="INRAE">',
             "<h1 style='text-align: center; color : aqua'>Traitement de donnée Excel : </h1>",
             unsafe_allow_html=True)
    saut_ligne(2)
    '''
    file = glob2.glob('EXCEL\*.xlsx')
    names = []
    for i in file:
        names.append(splitext(basename(i))[0])
    Sc_dico = {}
    z = 0
    for i in names:
        Sc_dico[i] = z
        z += 1
    option = st.selectbox("Choisissez le fichier excel ⬇️", names)
    m = Sc_dico[option]
    saut_ligne(2)
    st.title(":blue[Table complète étudiée 📊] ")
    saut_ligne(3)
    classeur = load_workbook(file[m]) #file
    wb = classeur.sheetnames
    wa = wb[1]
    drap = classeur[wa]
    '''
    table_complete()
    return
