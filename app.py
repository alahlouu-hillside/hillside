import streamlit as st
import pandas as pd
import os

# --- 1. CONFIGURATION DE LA PAGE & CHARTE GRAPHIQUE ---
st.set_page_config(
    page_title="Hillside | CRM & Pilotage Immobilier d'Exception", 
    page_icon="🏠", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Injection de styles CSS de haute précision pour une esthétique executive
st.markdown("""
    <style>
        .main-header {
            font-size: 2.2rem;
            font-weight: 800;
            color: #1E3A8A;
            margin-bottom: 0.2rem;
        }
        .sub-header {
            font-size: 1.1rem;
            color: #4B5563;
            margin-bottom: 2rem;
        }
        .metric-card {
            background-color: #F8FAFC;
            border-left: 5px solid #3B82F6;
            padding: 15px;
            border-radius: 8px;
        }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "suivi_commercial_hillside.xlsx"
SHEET_NAME = "Pipeline Commercial"

# --- 2. ARCHITECTURE DE PERSISTENCE DES DONNÉES ---
def init_excel():
    df_init = pd.DataFrame(columns=[
        "ID", "Nom Client", "Téléphone", "Email", "Projet / Villa", 
        "Type de bien", "Statut", "Prix Vente (MAD)", "Montant Encaissé (MAD)", 
        "Reste à Payer (MAD)", "Lien Plan / Brochure", "Date de contact", "Notes"
    ])
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_init.to_excel(writer, sheet_name=SHEET_NAME, index=False)

if not os.path.exists(EXCEL_FILE):
    init_excel()

@st.cache_data(ttl=1)
def load_data():
    try:
        return pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    except Exception:
        init_excel()
        return pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)

def save_data(df):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

df_clients = load_data()

# --- 3. MOTEUR DE COLORISATION VISUELLE ---
def color_status(val):
    color_map = {
        "Validé / Réservé": "background-color: #D1FAE5; color: #065F46; font-weight: bold;",
        "Offre envoyée": "background-color: #DBEAFE; color: #1E40AF; font-weight: bold;",
        "Visite planifiée": "background-color: #FEF3C7; color: #92400E; font-weight: bold;",
        "En cours de relance": "background-color: #F3F4F6; color: #374151;",
        "Perdu": "background-color: #FEE2E2; color: #991B1B;"
    }
    return color_map.get(val, "")

# --- 4. NAVIGATION EXCLUSIVE ---
st.sidebar.markdown("## 🏢 HILLSIDE IMMOBILIER")
st.sidebar.markdown("---")
menu = st.sidebar.radio("Navigation Centrale", [
    "🏠 Accueil & Hub Central", 
    "📊 Tableau de Bord & Indicateurs", 
    "💰 Suivi Financier & Trésorerie",
    "📁 Gestion des Plans & Documents",
    "📈 Pipeline & Édition des Leads", 
    "✍️ Saisie / Ajout de Prospect",
    "⚙️ Paramètres & Export"
])

# ====================================================
# MODULE 1 : ACCUEIL & HUB CENTRAL
# ====================================================
if menu == "🏠 Accueil & Hub Central":
    st.markdown('<p class="main-header">🏠 Accueil & Hub Central - Hillside</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Plateforme intégrée de pilotage commercial et de gestion d’actifs immobiliers haut standing.</p>', unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Prospects", len(df_clients))
    with col2:
        total_v = df_clients["Prix Vente (MAD)"].sum() if not df_clients.empty and "Prix Vente (MAD)" in df_clients.columns else 0
        st.metric("CA Potentiel", f"{total_v:,.2f} MAD")
    with col3:
        total_e = df_clients["Montant Encaissé (MAD)"].sum() if not df_clients.empty and "Montant Encaissé (MAD)" in df_clients.columns else 0
        st.metric("Total Encaissé", f"{total_e:,.2f} MAD")
    with col4:
        valides = len(df_clients[df_clients["Statut"] == "Validé / Réservé"]) if not df_clients.empty else 0
        st.metric("Réservations Actives", valides)
        
    st.markdown("---")
    st.info("💡 **Sécurité et Autonomie :** Le système opère en mode local transactionnel. Les données et les documents sont synchronisés en temps réel.")

# ====================================================
# MODULE 2 : TABLEAU DE BORD & INDICATEURS
# ====================================================
elif menu == "📊 Tableau de Bord & Indicateurs":
    st.markdown('<p class="main-header">📊 Tableau de Bord Commercial</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Analyse analytique de la performance et de l’avancement des ventes.</p>', unsafe_allow_html=True)
    
    if df_clients.empty:
        st.warning("Aucune donnée enregistrée dans le système.")
    else:
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("📌 Répartition par Statut Commercial")
            if "Statut" in df_clients.columns:
                st.bar_chart(df_clients["Statut"].value_counts())
        with col_b:
            st.subheader("🏗️ Répartition par Programme / Projet")
            if "Projet / Villa" in df_clients.columns:
                st.bar_chart(df_clients["Projet / Villa"].value_counts())
            
        st.markdown("---")
        st.subheader("📋 Matrice Globale avec Codes Couleurs")
        st.dataframe(df_clients.style.applymap(color_status, subset=["Statut"]), use_container_width=True)

# ====================================================
# MODULE 3 : SUIVI FINANCIER & TRÉSORERIE
# ====================================================
elif menu == "💰 Suivi Financier & Trésorerie":
    st.markdown('<p class="main-header">💰 Suivi Financier & Reste à Payer</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Contrôle strict des encaissements et des créances clients.</p>', unsafe_allow_html=True)
    
    if df_clients.empty:
        st.info("Aucun dossier financier disponible.")
    else:
        df_clients["Reste à Payer (MAD)"] = df_clients["Prix Vente (MAD)"] - df_clients["Montant Encaissé (MAD)"]
        
        st.dataframe(df_clients[["Nom Client", "Projet / Villa", "Prix Vente (MAD)", "Montant Encaissé (MAD)", "Reste à Payer (MAD)", "Statut"]].style.applymap(color_status, subset=["Statut"]), use_container_width=True)
        
        total_vente = df_clients["Prix Vente (MAD)"].sum()
        total_encaisse = df_clients["Montant Encaissé (MAD)"].sum()
        total_reste = df_clients["Reste à Payer (MAD)"].sum()
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Chiffre d’Affaires Total", f"{total_vente:,.2f} MAD")
        c2.metric("Total Recouvré", f"{total_encaisse:,.2f} MAD")
        c3.metric("Encours Restant (Créances)", f"{total_reste:,.2f} MAD", delta_color="inverse")

# ====================================================
# MODULE 4 : GESTION DES PLANS & DOCUMENTS
# ====================================================
elif menu == "📁 Gestion des Plans & Documents":
    st.markdown('<p class="main-header">📁 Plans d’Architecture & Brochures</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Dépôt, indexation et téléchargement des supports techniques et commerciaux.</p>', unsafe_allow_html=True)
    
    st.subheader("📤 Téléversement de nouveaux plans ou documents")
    uploaded_file = st.file_uploader("Sélectionner un fichier (PDF, PNG, JPG, DWG)", type=["pdf", "png", "jpg", "dwg"])
    if uploaded_file is not None:
        file_path = os.path.join(".", uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success(f"Le fichier '{uploaded_file.name}' a été enregistré avec succès sur le serveur.")
        
    st.markdown("---")
    st.subheader("📥 Documents & Plans Disponibles")
    fichiers_disponibles = [f for f in os.listdir(".") if f.endswith((".pdf", ".png", ".jpg", ".dwg")) and f != EXCEL_FILE]
    
    if fichiers_disponibles:
        for f in fichiers_disponibles:
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"📄 **{f}**")
            with col2:
                with open(f, "rb") as file_download:
                    st.download_button(label="Télécharger", data=file_download, file_name=f, key=f"dl_{f}")
    else:
        st.info("Aucun document technique n’est actuellement indexé.")

# ====================================================
# MODULE 5 : PIPELINE & ÉDITION DES LEADS
# ====================================================
elif menu == "📈 Pipeline & Édition des Leads":
    st.markdown('<p class="main-header">📈 Pipeline Commercial & Édition en Direct</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Modification directe de la base de données opérationnelle.</p>', unsafe_allow_html=True)
    
    if not df_clients.empty:
        edited_df = st.data_editor(df_clients, num_rows="dynamic", use_container_width=True, key="master_editor")
        if st.button("💾 Synchroniser et Sauvegarder les modifications"):
            save_data(edited_df)
            st.success("Synchronisation effectuée avec succès !")
            st.rerun()
    else:
        st.info("Le pipeline est vide.")

# ====================================================
# MODULE 6 : SAISIE / AJOUT DE PROSPECT
# ====================================================
elif menu == "✍️ Saisie / Ajout de Prospect":
    st.markdown('<p class="main-header">✍️ Enregistrement d’un Nouveau Prospect</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Formulaire de qualification commerciale de haute précision.</p>', unsafe_allow_html=True)
    
    with st.form("form_prospect_expert", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            nom = st.text_input("Nom & Prénom du Client *")
            telephone = st.text_input("Téléphone")
            email = st.text_input("Email")
            projet = st.selectbox("Programme Immobilier", [
                "Hillside Residences (Bouskoura)", 
                "Villa KM", 
                "Les Collines de Dar Bouazza", 
                "Autre"
            ])
            type_bien = st.selectbox("Type de Bien", ["Appartement", "Villa", "Terrain", "Commercial"])
        with col2:
            statut = st.selectbox("Statut du Lead", [
                "Nouveau Prospect", 
                "En cours de relance", 
                "Visite planifiée", 
                "Offre envoyée", 
                "Validé / Réservé", 
                "Perdu"
            ])
            prix_vente = st.number_input("Prix de Vente Total (MAD)", min_value=0.0, step=50000.0)
            montant_encaisse = st.number_input("Montant Déjà Encaissé (MAD)", min_value=0.0, step=10000.0)
            date_contact = st.date_input("Date du Premier Contact")
            
        notes = st.text_area("Historique des Échanges & Remarques Commerciales")
        
        submitted = st.form_submit_button("Valider et Enregistrer le Prospect")
        if submitted:
            if nom:
                new_id = int(df_clients["ID"].max() + 1) if not df_clients.empty and "ID" in df_clients.columns and pd.notna(df_clients["ID"].max()) else 1
                reste_a_payer = prix_vente - montant_encaisse
                new_row = {
                    "ID": new_id,
                    "Nom Client": nom,
                    "Téléphone": telephone,
                    "Email": email,
                    "Projet / Villa": projet,
                    "Type de bien": type_bien,
                    "Statut": statut,
                    "Prix Vente (MAD)": prix_vente,
                    "Montant Encaissé (MAD)": montant_encaisse,
                    "Reste à Payer (MAD)": reste_a_payer,
                    "Lien Plan / Brochure": "N/A",
                    "Date de contact": str(date_contact),
                    "Notes": notes
                }
                df_clients = pd.concat([df_clients, pd.DataFrame([new_row])], ignore_index=True)
                save_data(df_clients)
                st.success(f"Le dossier de **{nom}** a été intégré au système avec succès !")
            else:
                st.error("Erreur critique : Le nom du client est obligatoire.")

# ====================================================
# MODULE 7 : PARAMÈTRES & EXPORT
# ====================================================
elif menu == "⚙️ Paramètres & Export":
    st.markdown('<p class="main-header">⚙️ Administration & Export des Données</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Gestion de la base de données et sécurisation des sauvegardes.</p>', unsafe_allow_html=True)
    
    if not df_clients.empty:
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="📥 Télécharger la Base de Données Complète (Excel Master)",
                data=f,
                file_name="suivi_commercial_hillside_master.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    st.markdown("---")
    st.subheader("⚠️ Zone de Maintenance Critique")
    if st.button("Réinitialiser l’Intégralité de la Base de Données"):
        if os.path.exists(EXCEL_FILE):
            os.remove(EXCEL_FILE)
            st.success("La base de données a été purgée et réinitialisée à zéro.")
            st.rerun()
