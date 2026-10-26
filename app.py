import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Configuration de la page
st.set_page_config(page_title="CRM - Hillside", page_icon="🏠", layout="wide")

st.title("🏠 Accueil & Hub Central - Hillside")
st.markdown("Tableau de bord de pilotage commercial connecté en direct avec votre Google Sheet.")

# Connexion sécurisée avec nettoyage automatique et blindé de la clé privée
@st.cache_resource
def get_gspread_client():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    # Récupération des secrets
    creds_dict = dict(st.secrets["connections"]["gsheets"])
    
    # Nettoyage et formatage blindé de la clé privée
    if "private_key" in creds_dict:
        pk = creds_dict["private_key"].strip()
        # Remplace les éventuels "\n" textuels par de vrais retours à la ligne
        pk = pk.replace("\\n", "\n")
        # S'assure que les en-têtes et pieds de page sont bien séparés par des sauts de ligne
        pk = pk.replace("-----BEGIN PRIVATE KEY-----", "-----BEGIN PRIVATE KEY-----\n")
        pk = pk.replace("-----END PRIVATE KEY-----", "\n-----END PRIVATE KEY-----")
        creds_dict["private_key"] = pk
        
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

try:
    client = get_gspread_client()
    
    # ⚠️ Mettez ici le nom exact de votre fichier Google Sheet sur votre Google Drive
    sh = client.open("Suivi Commercial & Pipeline Leads - Immobilier - Google Sheets")
    
    # Menu latéral pour choisir l'onglet à afficher
    page = st.sidebar.selectbox("Navigation", ["Tableau de Bord", "Pipeline Commercial"])
    
    if page == "Tableau de Bord":
        st.subheader("📊 Tableau de Bord")
        worksheet = sh.worksheet("Tableau de Bord")
        df_dashboard = pd.DataFrame(worksheet.get_all_records())
        st.dataframe(df_dashboard, use_container_width=True)
        
    elif page == "Pipeline Commercial":
        st.subheader("📈 Pipeline Commercial")
        worksheet = sh.worksheet("Pipeline Commercial")
        df_pipeline = pd.DataFrame(worksheet.get_all_records())
        st.dataframe(df_pipeline, use_container_width=True)

except Exception as e:
    st.error("⚠️ Impossible de récupérer les données du Google Sheet pour le moment.")
    st.info("Vérifiez que le nom de votre Google Sheet dans le code (`client.open(...)`) correspond exactement au nom réel du fichier sur Google Drive.")
    with st.expander("Voir les détails techniques de l'erreur"):
        st.write(e)
