import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="CRM - Hillside & Immobilier", page_icon="🏠", layout="wide")

EXCEL_FILE = "suivi_commercial_hillside.xlsx"
SHEET_NAME = "Pipeline Commercial"

# Initialisation du fichier Excel avec toutes les colonnes avancées (Encaissements, Plans, etc.)
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

# Barre latérale de navigation
st.sidebar.title("🏢 Hillside - Pilotage")
menu = st.sidebar.radio("Navigation :", [
    "🏠 Accueil & Hub Central", 
    "📊 Tableau de Bord & Encaissements", 
    "💰 Suivi Financier & Reste à Payer",
    "📁 Gestion des Plans & Documents",
    "📈 Pipeline & Édition des Leads", 
    "✍️ Saisie / Ajout de Prospect",
    "⚙️ Paramètres & Export"
])

# Fonction utilitaire pour appliquer un code couleur selon le statut
def color_status(val):
    color_map = {
        "Validé / Réservé": "background-color: #d4edda; color: #155724;",
        "Offre envoyée": "background-color: #cce5ff; color: #004085;",
        "Visite planifiée": "background-color: #fff3cd; color: #856404;",
        "En cours de relance": "background-color: #e2e3e5; color: #383d41;",
        "Perdu": "background-color: #f8d7da; color: #721c24;"
    }
    return color_map.get(val, "")

# ----------------------------------------------------
# 1. ACCUEIL & HUB CENTRAL
# ----------------------------------------------------
if menu == "🏠 Accueil & Hub Central":
    st.title("🏠 Accueil & Hub Central - Hillside")
    st.markdown("Tableau de bord de pilotage commercial et suivi des encaissements.")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Total Prospects", len(df_clients))
    with col2:
        total_v = df_clients["Prix Vente (MAD)"].sum() if not df_clients.empty and "Prix Vente (MAD)" in df_clients.columns else 0
        st.metric("Chiffre d'Affaires Potentiel", f"{total_v:,.2f} MAD")
    with col3:
        total_e = df_clients["Montant Encaissé (MAD)"].sum() if not df_clients.empty and "Montant Encaissé (MAD)" in df_clients.columns else 0
        st.metric("Total Encaissé", f"{total_e:,.2f} MAD")
    with col4:
        valides = len(df_clients[df_clients["Statut"] == "Validé / Réservé"]) if not df_clients.empty else 0
        st.metric("Réservations", valides)

# ----------------------------------------------------
# 2. TABLEAU DE BORD & ENCAISSEMENTS
# ----------------------------------------------------
elif menu == "📊 Tableau de Bord & Encaissements":
    st.title("📊 Tableau de Bord & Performance Commerciale")
    
    if df_clients.empty:
        st.warning("Aucune donnée disponible.")
    else:
        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader("📌 Répartition par Statut (Code Couleur)")
            st.bar_chart(df_clients["Statut"].value_counts())
        with col_b:
            st.subheader("🏗️ Répartition par Projet / Villa")
            st.bar_chart(df_clients["Projet / Villa"].value_counts())
            
        st.markdown("---")
        st.subheader("📋 Vue d'ensemble avec codes couleurs")
        st.dataframe(df_clients.style.applymap(color_status, subset=["Statut"]), use_container_width=True)

# ----------------------------------------------------
# 3. SUIVI FINANCIER & RESTE À PAYER
# ----------------------------------------------------
elif menu == "💰 Suivi Financier & Reste à Payer":
    st.title("💰 Suivi des Encaissements & Reste à Payer")
    
    if df_clients.empty:
        st.info("Aucun client enregistré.")
    else:
        # Calcul automatique du reste à payer
        df_clients["Reste à Payer (MAD)"] = df_clients["Prix Vente (MAD)"] - df_clients["Montant Encaissé (MAD)"]
        
        st.dataframe(df_clients[["Nom Client", "Projet / Villa", "Prix Vente (MAD)", "Montant Encaissé (MAD)", "Reste à Payer (MAD)", "Statut"]], use_container_width=True)
        
        total_vente = df_clients["Prix Vente (MAD)"].sum()
        total_encaisse = df_clients["Montant Encaissé (MAD)"].sum()
        total_reste = df_clients["Reste à Payer (MAD)"].sum()
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Ventes", f"{total_vente:,.2f} MAD")
        c2.metric("Total Encaissé", f"{total_encaisse:,.2f} MAD")
        c3.metric("Total Reste à Recouvrer", f"{total_reste:,.2f} MAD", delta_color="inverse")

# ----------------------------------------------------
# 4. GESTION DES PLANS & DOCUMENTS
# ----------------------------------------------------
elif menu == "📁 Gestion des Plans & Documents":
    st.title("📁 Plans & Brochures des Lots / Villas")
    st.markdown("Téléchargez ou consultez les plans associés à chaque projet.")
    
    # Section de téléchargement de documents pour l'application
    st.subheader("📤 Mettre en ligne un plan (PDF / Image)")
    uploaded_file = st.file_uploader("Sélectionner un plan ou une brochure", type=["pdf", "png", "jpg"])
    if uploaded_file is not None:
        with open(uploaded_file.name, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success(f"Le fichier {uploaded_file.name} a été enregistré avec succès sur le serveur !")
        
    st.markdown("---")
    st.subheader("📥 Plans & Documents Disponibles")
    # Liste des fichiers présents dans le dossier
    fichiers_disponibles = [f for f in os.listdir(".") if f.endswith((".pdf", ".png", ".jpg", ".dwg")) and f != "suivi_commercial_hillside.xlsx"]
    
    if fichiers_disponibles:
        for f in fichiers_disponibles:
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"📄 **{f}**")
            with col2:
                with open(f, "rb") as file_download:
                    st.download_button(label="Télécharger", data=file_download, file_name=f, key=f)
    else:
        st.info("Aucun plan ou document n'a encore été mis en ligne. Utilisez le bouton ci-dessus pour en ajouter.")

# ----------------------------------------------------
# 5. PIPELINE & ÉDITION DES LEADS
# ----------------------------------------------------
elif menu == "📈 Pipeline & Édition des Leads":
    st.title("📈 Pipeline & Modification en direct")
    if not df_clients.empty:
        edited_df = st.data_editor(df_clients, num_rows="dynamic", use_container_width=True, key="editor_leads")
        if st.button("💾 Sauvegarder les modifications"):
            save_data(edited_df)
            st.success("Modifications enregistrées avec succès !")
            st.rerun()
    else:
        st.info("Aucune donnée à éditer.")

# ----------------------------------------------------
# 6. SAISIE / AJOUT DE PROSPECT
# ----------------------------------------------------
elif menu == "✍️ Saisie / Ajout de Prospect":
    st.title("✍️ Enregistrer un nouveau client / prospect")
    
    with st.form("form_prospect_complet", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            nom = st.text_input("Nom & Prénom du Client *")
            telephone = st.text_input("Téléphone")
            email = st.text_input("Email")
            projet = st.selectbox("Projet / Villa", ["Hillside Residences (Bouskoura)", "Villa KM", "Les Collines de Dar Bouazza", "Autre"])
            type_bien = st.selectbox("Type de bien", ["Appartement", "Villa", "Terrain", "Commercial"])
        with col2:
            statut = st.selectbox("Statut du Lead", ["Nouveau Prospect", "En cours de relance", "Visite planifiée", "Offre envoyée", "Validé / Réservé", "Perdu"])
            prix_vente = st.number_input("Prix de Vente Total (MAD)", min_value=0.0, step=50000.0)
            montant_encaisse = st.number_input("Montant Déjà Encaissé (MAD)", min_value=0.0, step=10000.0)
            date_contact = st.date_input("Date du contact")
            
        notes = st.text_area("Notes / Remarques commerciales & Historique")
        
        submitted = st.form_submit_button("Enregistrer dans le CRM")
        if submitted:
            if nom:
                new_id = len(df_clients) + 1
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
                st.success(f"Prospect **{nom}** ajouté avec succès !")
            else:
                st.error("Le nom du client est obligatoire.")

# ----------------------------------------------------
# 7. PARAMÈTRES & EXPORT
# ----------------------------------------------------
elif menu == "⚙️ Paramètres & Export":
    st.title("⚙️ Paramètres & Export de la Base")
    if not df_clients.empty:
        with open(EXCEL_FILE, "rb") as f:
            st.download_button("Télécharger le fichier Excel complet", data=f, file_name="suivi_commercial_hillside.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if st.button("⚠️ Réinitialiser toute la base de données"):
        if os.path.exists(EXCEL_FILE):
            os.remove(EXCEL_FILE)
            st.success("Base réinitialisée.")
            st.rerun()
