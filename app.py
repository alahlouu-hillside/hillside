import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# Configuration de la page
st.set_page_config(page_title="CRM - Hillside", page_icon="🏠", layout="wide")

st.title("🏠 Accueil & Hub Central - Hillside")
st.markdown("Tableau de bord de pilotage commercial connecté en direct avec votre Google Sheet.")

@st.cache_resource
def get_gspread_client():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    secrets = st.secrets["connections"]["gsheets"]
    
    # Nettoyage ultra-robuste de la clé privée pour éliminer tout caractère parasite ou guillemet en trop
    pk = str(secrets["private_key"]).strip()
    pk = pk.strip('"').strip("'")  # Supprime les guillemets englobants si présents
    pk = pk.replace("\\n", "\n")
    
    if "-----BEGIN PRIVATE KEY-----" not in pk:
        pk = "-----BEGIN PRIVATE KEY-----\n" + pk
    if "-----END PRIVATE KEY-----" not in pk:
        pk = pk + "\n-----END PRIVATE KEY-----"

    creds_dict = {
        "type": "service_account",
        "project_id": str(secrets.get("project_id", "crm-hillside")),
        "private_key_id": str(secrets.get("private_key_id", "")),
        "private_key": pk,
        "client_email": str(secrets.get("client_email", "")),
        "client_id": str(secrets.get("client_id", "")),
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
        "client_x509_cert_url": str(secrets.get("client_x509_cert_url", ""))
    }
    
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
    return gspread.authorize(creds)

try:
    client = get_gspread_client()
    
    # Connexion à votre fichier exact sur Google Drive
    sh = client.open("Suivi Commercial & Pipeline Leads - Immobilier")
    
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
    with st.expander("Voir les détails techniques de l'erreur"):
        st.write(e)
