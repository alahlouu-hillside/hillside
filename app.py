import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="CRM - Hillside & Immobilier", page_icon="🏠", layout="wide")

EXCEL_FILE = "suivi_commercial.xlsx"

# Initialisation automatique du fichier Excel avec une structure complète si absent
if not os.path.exists(EXCEL_FILE):
    df_init = pd.DataFrame(columns=[
        "ID", "Nom Client", "Téléphone", "Email", "Projet / Villa", 
        "Type de bien", "Statut", "Budget (MAD)", "Date de contact", "Notes"
    ])
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_init.to_excel(writer, sheet_name="Pipeline Commercial", index=False)

@st.cache_data(ttl=1)
def load_data():
    return pd.read_excel(EXCEL_FILE, sheet_name="Pipeline Commercial")

def save_data(df):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="Pipeline Commercial", index=False)

df_clients = load_data()

# Barre latérale de navigation complète
st.sidebar.title("🏢 Navigation - Hillside")
menu = st.sidebar.radio("Aller vers :", [
    "🏠 Accueil & Hub Central", 
    "📊 Tableau de Bord Commercial", 
    "📈 Pipeline & Suivi des Leads", 
    "✍️ Saisie / Ajout de Prospect",
    "⚙️ Paramètres & Données"
])

# ----------------------------------------------------
# 1. ACCUEIL & HUB CENTRAL
# ----------------------------------------------------
if menu == "🏠 Accueil & Hub Central":
    st.title("🏠 Accueil & Hub Central - Hillside")
    st.markdown("Tableau de bord de pilotage commercial intégré et autonome pour vos projets immobiliers.")
    
    st.info("💡 **État du système :** L'application fonctionne en mode autonome avec un stockage local sécurisé (`suivi_commercial.xlsx`). Toutes vos données sont conservées en temps réel.")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Prospects", len(df_clients))
    with col2:
        if not df_clients.empty and "Budget (MAD)" in df_clients.columns:
            total_b = df_clients["Budget (MAD)"].sum()
            st.metric("Valeur Totale du Pipeline", f"{total_b:,.2f} MAD")
        else:
            st.metric("Valeur Totale du Pipeline", "0 MAD")
    with col3:
        if not df_clients.empty and "Statut" in df_clients.columns:
            valides = len(df_clients[df_clients["Statut"] == "Validé / Réservé"])
            st.metric("Réservations Confirmées", valides)
        else:
            st.metric("Réservations Confirmées", 0)

# ----------------------------------------------------
# 2. TABLEAU DE BORD COMMERCIAL
# ----------------------------------------------------
elif menu == "📊 Tableau de Bord Commercial":
    st.title("📊 Tableau de Bord & Indicateurs Clés")
    
    if df_clients.empty:
        st.warning("Aucune donnée disponible pour le moment. Veuillez ajouter des prospects depuis l'onglet de saisie.")
    else:
        col_a, col_b = st.columns(2)
        
        with col_a:
            st.subheader("📌 Répartition par Statut")
            if "Statut" in df_clients.columns:
                statut_counts = df_clients["Statut"].value_counts()
                st.bar_chart(statut_counts)
                
        with col_b:
            st.subheader("🏗️ Répartition par Projet")
            if "Projet / Villa" in df_clients.columns:
                projet_counts = df_clients["Projet / Villa"].value_counts()
                st.bar_chart(projet_counts)
                
        st.markdown("---")
        st.subheader("📋 Vue synthétique des leads actifs")
        st.dataframe(df_clients, use_container_width=True)

# ----------------------------------------------------
# 3. PIPELINE & SUIVI DES LEADS (TABLEAU ÉDITABLE)
# ----------------------------------------------------
elif menu == "📈 Pipeline & Suivi des Leads":
    st.title("📈 Pipeline & Édition des Données")
    st.markdown("Modifiez directement vos données ci-dessous (les changements s'enregistrent instantanément).")
    
    if not df_clients.empty:
        edited_df = st.data_editor(df_clients, num_rows="dynamic", use_container_width=True, key="pipeline_editor")
        if st.button("💾 Enregistrer les modifications du tableau"):
            save_data(edited_df)
            st.success("Modifications enregistrées avec succès !")
            st.rerun()
    else:
        st.info("Le tableau est actuellement vide.")

# ----------------------------------------------------
# 4. SAISIE / AJOUT DE PROSPECT
# ----------------------------------------------------
elif menu == "✍️ Saisie / Ajout de Prospect":
    st.title("✍️ Enregistrer un nouveau prospect")
    
    with st.form("form_nouveau_prospect", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            nom = st.text_input("Nom & Prénom du Client *")
            telephone = st.text_input("Téléphone")
            email = st.text_input("Email")
            projet = st.selectbox("Projet concerné", [
                "Hillside Residences (Bouskoura)", 
                "Villa KM", 
                "Les Collines de Dar Bouazza", 
                "Autre"
            ])
        with col2:
            type_bien = st.selectbox("Type de bien", ["Appartement", "Villa", "Terrain", "Commercial"])
            statut = st.selectbox("Statut du Lead", [
                "Nouveau Prospect", 
                "En cours de relance", 
                "Visite planifiée", 
                "Offre envoyée", 
                "Validé / Réservé", 
                "Perdu"
            ])
            budget = st.number_input("Budget estimé (MAD)", min_value=0.0, step=50000.0)
            date_contact = st.date_input("Date du premier contact")
            
        notes = st.text_area("Notes / Remarques commerciales & historiques des échanges")
        
        submitted = st.form_submit_button("Ajouter le prospect au CRM")
        
        if submitted:
            if nom:
                new_id = len(df_clients) + 1
                new_row = {
                    "ID": new_id,
                    "Nom Client": nom,
                    "Téléphone": telephone,
                    "Email": email,
                    "Projet / Villa": projet,
                    "Type de bien": type_bien,
                    "Statut": statut,
                    "Budget (MAD)": budget,
                    "Date de contact": str(date_contact),
                    "Notes": notes
                }
                df_clients = pd.concat([df_clients, pd.DataFrame([new_row])], ignore_index=True)
                save_data(df_clients)
                st.success(f"Le prospect **{nom}** a été ajouté avec succès !")
            else:
                st.error("Le nom du client est obligatoire.")

# ----------------------------------------------------
# 5. PARAMÈTRES & DONNÉES
# ----------------------------------------------------
elif menu == "⚙️ Paramètres & Données":
    st.title("⚙️ Gestion des données & Export")
    
    st.subheader("📥 Exporter vos données")
    if not df_clients.empty:
        # Bouton de téléchargement direct en Excel
        with open(EXCEL_FILE, "rb") as f:
            st.download_button(
                label="Télécharger la base de données (Excel)",
                data=f,
                file_name="suivi_commercial_hillside.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("Aucune donnée à exporter pour l'instant.")
        
    st.markdown("---")
    st.subheader("⚠️ Zone de danger")
    if st.button("Réinitialiser toute la base de données"):
        if os.path.exists(EXCEL_FILE):
            os.remove(EXCEL_FILE)
            st.success("La base de données a été réinitialisée.")
            st.rerun()
