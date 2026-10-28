import json
import gspread
from google.oauth2.service_account import Credentials
import streamlit as st
import pandas as pd

# Configuration de la page Streamlit
st.set_page_config(
    page_title="Hillside Residence - CRM",
    page_icon="🏠",
    layout="wide"
)

# Constantes et Scopes
SPREADSHEET_ID = "1-d-J5oB1BlT3JeUUPpann-nECzFw01cf-HBqeBjM1OA"
SCOPES = [
    "https://www.spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

def get_gspread_client():
    try:
        # Vérification de l'existence des secrets dans l'environnement Streamlit Cloud
        if not hasattr(st, "secrets") or "gcp_service_account" not in st.secrets:
            return None
            
        # Récupération sous forme de dictionnaire d'origine
        creds_dict = dict(st.secrets["gcp_service_account"])
        
        # Normalisation rigoureuse de la clé privée pour éviter tout problème de saut de ligne
        if "private_key" in creds_dict:
            pk = str(creds_dict["private_key"])
            # Remplace les éventuels doubles anti-slashes littéraux par de vrais retours à la ligne
            pk = pk.replace("\\n", "\n")
            creds_dict["private_key"] = pk

        # Autorisation et initialisation du client gspread
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        gc = gspread.authorize(creds)
        return gc
        
    except Exception:
        return None

@st.cache_data(ttl=30)
def synchroniser_google_sheet():
    try:
        gc = get_gspread_client()
        if not gc:
            return pd.DataFrame()
        
        sh = gc.open_by_key(SPREADSHEET_ID)
        
        # Recherche de l'onglet "pipeline commercial" ou repli sur le premier onglet
        titres_onglets = [w.title.strip().lower() for w in sh.worksheets()]
        if "pipeline commercial" in titres_onglets:
            worksheet = sh.worksheet("pipeline commercial")
        else:
            worksheet = sh.get_worksheet(0)

        lignes = worksheet.get_all_values()
        if len(lignes) > 1:
            en_tetes = [str(h).strip() if h != "" else f"Col_{i}" for i, h in enumerate(lignes[0])]
            df_gs = pd.DataFrame(lignes[1:], columns=en_tetes)
            df_gs = df_gs.loc[:, ~df_gs.columns.str.startswith("Col_")]
            df_gs = df_gs.fillna("").loc[(df_gs != "").any(axis=1)]
            return df_gs
        return pd.DataFrame()
    except Exception:
        return pd.DataFrame()

# --- Interface principale de l'application ---
st.title("🏠 Accueil & Hub Central - Hillside")
st.markdown("Tableau de bord de pilotage commercial connecté en direct avec votre Google Sheet.")

st.divider()

# Test de connexion et chargement des données
df_data = synchroniser_google_sheet()

if not df_data.empty:
    st.success("🟢 Connexion Google Sheets établie avec succès !")
    st.dataframe(df_data, use_container_width=True)
else:
    st.info("ℹ️ Mode autonome actif (Google Sheets non détecté ou non configuré, affichage des données locales).")
    st.warning("Aucune donnée disponible dans le pipeline.")
