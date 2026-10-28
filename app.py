import json
import gspread
from google.oauth2.service_account import Credentials
import streamlit as st
import pandas as pd

# --- Configuration de la page Streamlit ---
st.set_page_config(
    page_title="Hillside Residence - CRM",
    page_icon="🏠",
    layout="wide"
)

# --- Constantes et Scopes ---
SPREADSHEET_ID = "1-d-J5oB1BlT3JeUUPpann-nECzFw01cf-HBqeBjM1OA"
SCOPES = [
    "https://www.spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

# --- Fonction d'authentification Google Sheets robuste ---
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

# --- Fonction de synchronisation Google Sheets avec cache ---
@st.cache_data(ttl=30)
def synchroniser_google_sheet():
    try:
        gc = get_gspread_client()
        if not gc:
            return pd.DataFrame()
        
        sh = gc.open_by_key(SPREADSHEET_ID)
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

# --- Barre latérale (Navigation & Compte) ---
st.sidebar.markdown("## HILLSIDE RESIDENCE")
st.sidebar.markdown("*By BEAURIVAGE*")
st.sidebar.divider()

with st.sidebar.expander("👤 Mon Compte Administrateur", expanded=False):
    st.markdown("**Statut :** Connecté")
    st.markdown("**Rôle :** Directeur Commercial")

st.sidebar.success("✔️ Administrateur Actif")

if st.sidebar.button("🔄 Rafraîchir les Données", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("### NAVIGATION PRINCIPALE")

# Choix de la page active via la sidebar
page = st.sidebar.radio(
    "Aller vers",
    [
        "🏠 Accueil & Hub Central",
        "🤖 Assistant AI & Communication",
        "👥 Portefeuille Prospects & Clients",
        "🏢 Catalogue & Stocks Immobiliers",
        "✍️ Espace Ventes & Quittances",
        "📊 Chiffre d'Affaires & Trésorerie",
        "⚙️ Paramètres & Analyses"
    ],
    label_visibility="collapsed"
)

# --- Chargement des données globales ---
df_data = synchroniser_google_sheet()
google_connecte = not df_data.empty

# --- ROUTAGE DES PAGES ---

if page == "🏠 Accueil & Hub Central":
    st.title("🏠 Accueil & Hub Central - Hillside")
    st.markdown("Tableau de bord de pilotage commercial connecté en direct avec votre Google Sheet.")
    st.divider()

    if google_connecte:
        st.success("🟢 Connexion Google Sheets établie avec succès !")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Prospects / Leads", len(df_data))
        with col2:
            st.metric("Projets Actifs", "Hillside Bouskoura")
        with col3:
            st.metric("Statut Pipeline", "Actif & Synchro")
            
        st.subheader("Aperçu des données du Pipeline Commercial")
        st.dataframe(df_data, use_container_width=True)
    else:
        st.info("ℹ️ Mode autonome actif (Google Sheets non détecté ou non configuré, affichage des données locales).")
        st.warning("Aucune donnée disponible dans le pipeline.")

elif page == "🤖 Assistant AI & Communication":
    st.title("🤖 Assistant AI & Communication")
    st.markdown("Générez des messages de suivi, qualifiez vos prospects et automatisez vos relances.")
    st.divider()
    st.info("Module IA opérationnel - Prêt pour la génération de scripts de vente et d'emails.")

elif page == "👥 Portefeuille Prospects & Clients":
    st.title("👥 Portefeuille Prospects & Clients")
    st.markdown("Suivi détaillé de vos leads, prospects chauds et réservations.")
    st.divider()
    if google_connecte:
        st.dataframe(df_data, use_container_width=True)
    else:
        st.warning("Aucune donnée client chargée depuis le Google Sheet.")

elif page == "🏢 Catalogue & Stocks Immobiliers":
    st.title("🏢 Catalogue & Stocks Immobiliers")
    st.markdown("État des unités (F3, F4, commerces) et disponibilités en temps réel.")
    st.divider()
    st.info("Gestion des stocks et inventaire des lots de la résidence Hillside.")

elif page == "✍️ Espace Ventes & Quittances":
    st.title("✍️ Espace Ventes & Quittances")
    st.markdown("Édition des contrats VEFA et des quittances de versement.")
    st.divider()
    st.info("Module d'édition documentaire et de génération de reçus.")

elif page == "📊 Chiffre d'Affaires & Trésorerie":
    st.title("📊 Chiffre d'Affaires & Trésorerie")
    st.markdown("Suivi des encaissements, échéanciers et prévisions financières.")
    st.divider()
    st.info("Tableaux de bord financiers et états des versements.")

elif page == "⚙️ Paramètres & Analyses":
    st.title("⚙️ Paramètres & Analyses")
    st.markdown("Configuration générale du CRM et statistiques avancées.")
    st.divider()
    st.markdown(f"- **ID Google Sheet configuré :** `{SPREADSHEET_ID}`")
    st.markdown(f"- **État de la liaison :** {'Connecté' if google_connecte else 'Mode Autonome'}")
