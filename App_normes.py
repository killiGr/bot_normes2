import pandas as pd
from bs4 import BeautifulSoup
import requests
import streamlit as st
from io import BytesIO


# Demander la norme ---------------------------------------------------------------------------------------------------------------
st.write('# Choisir la norme')
st.write('#### Utilisation :')
st.write('- Vous pouvez écrire une ou plusiurs normes dans l\'encadré séparées par une virgule')
st.write('- Il est possible de copier/coller une colonne excel')
st.write('- Les doublons ne sont pas dérangeants')
norme=st.text_input('Norme :')
norme=list(pd.Series(norme.split()).drop_duplicates()) # crée la liste d'élements


# Traitement
button=st.checkbox('Rechercher :')
if button:
    
    df_tot2=pd.DataFrame()
    
    for element in norme:
        df_tot=pd.DataFrame()
        num_page=0
        while num_page!=9:
            url='https://www.boutique.afnor.org/fr-fr/resultats?Keywords='+element+'&Culture=fr-FR&PageIndex='+str(num_page)+'&PageSize=100&SortPropertyName=_score&SortByDescending=True&StandardStateIds[0]=1&StandardStateIds[1]=2&StandardStateIds[2]=3&StandardStateIds[3]=4&MandatoryApplication=0&Harmonised=0&IsNovelty=0'
            page = requests.get(url)
            soup = BeautifulSoup(page.content, 'html.parser')

            # Nom
            nom = soup.find_all('h2', {'class': 'title'})
            for i,ele in enumerate(nom):
                nom[i]=ele.get_text()

            # Titre
            titre = soup.find_all('div', {'class': 'prod-intro'})
            for i,ele in enumerate(titre):
                titre[i]=ele.get_text()

            # Date
            date = soup.find_all('div', {'class': 'date'})
            for i,ele in enumerate(date):
                date[i]=ele.get_text()

            # En vigueur
            vigueur = soup.find_all('span', class_=['slib current', 'slib cancelled','slib project','slib cancelledproject'])
            for i,ele in enumerate(vigueur):
                if ele.get_text()=='En vigueur':
                    vigueur[i]='Oui'
                else:
                    vigueur[i]='Non'


            df=pd.DataFrame({'N° Référence': nom, 'Intitulé du document': titre, 'Date': date, 'En vigueur': vigueur, })
            df['N° Référence'] = df['N° Référence'].astype(str)
            df=df.loc[df['N° Référence'].str.contains(element)]
            df_tot=pd.concat([df_tot,df])
            num_page=num_page+1

        # Mise en forme des données
        df_tot['Date'] = df_tot['Date'].str.replace(r'\D+', '').astype(int) # Garder que l'année pas le mois
        df_tot=df_tot.sort_values(by=['N° Référence','Date'], ascending=False) # organiser par ref et date pr...
        df_tot=df_tot.drop_duplicates(subset='N° Référence', keep='first')# ensuite drop les dates ultérieures
        df_tot=df_tot.loc[(df_tot['Intitulé du document'].str.lower().str.contains('foudre', regex=False)) | \
                          (df_tot['Intitulé du document'].str.lower().str.contains('foudroiment', regex=False)) | \
                          (df_tot['Intitulé du document'].str.lower().str.contains('foudroyé', regex=False)) | \
                          (df_tot['Intitulé du document'].str.lower().str.contains('éclateur', regex=False)) | \
                          (df_tot['Intitulé du document'].str.lower().str.contains('parafoudre', regex=False))]# Select normes liées à la foudre
        df_tot2=pd.concat([df_tot2,df_tot])
    
    # Mise en forme total
    df_tot2[['1', 'N° Référence']] = df_tot2['N° Référence'].str.extract(r'([A-Za-z\s]+)\s*([\d-]+)') # scinder reference en 2
    df_tot2['Domaine']='Foudre' # Ajout colonne Domaine
    df_tot2.reset_index(inplace=True, drop=True)
    df_tot2 = df_tot2.reindex(columns=['N° Référence', '1', 'Intitulé du document','Date','En vigueur','Domaine'])
        
    # Affichage
    st.write(df_tot2)
    
    output = BytesIO()
    df_tot2.to_excel(output, index=False)
    output.seek(0)
    st.download_button(
        label="Télécharger le fichier",
        data=output.getvalue(),
        file_name='Output_norme_Afnor.xlsx',
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# streamlit run C:\Users\kyllian.gressier\Desktop\App_normes.py
