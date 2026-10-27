import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="CRM - Hillside", page_icon="🏠", layout="wide")

st.title("🏠 Accueil & Hub Central - Hillside")
st.markdown("Tableau de bord de pilotage commercial et gestion intégrée des clients.")

EXCEL_FILE = "suivi_commercial.xlsx"

# Initialisation du fichier Excel s'il n'existe pas encore
if not os.path.exists(EXCEL_FILE):
    df_init = pd.DataFrame(columns=["Nom Client", "Téléphone", "Email", "Projet / Villa", "Statut", "Budget (MAD)", "Notes"])
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_init.to_excel(writer, sheet_name="Clients", index=False)

@st.cache_data(ttl=1)
def load_data():
    return pd.read_excel(EXCEL_FILE, sheet_name="Clients")

def save_data(df):
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name="Clients", index=False)

df_clients = load_data()

# Navigation dans l'application
menu = st.sidebar.selectbox("Navigation", ["📊 Tableau de Bord", "👥 Gestion des Clients & Leads"])

if menu == "📊 Tableau de Bord":
    st.subheader("📊 Indicateurs Clés de Performance")
    
    if not df_clients.empty:
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Prospects / Clients", len(df_clients))
        with col2:
            if "Budget (MAD)" in df_clients.columns:
                total_budget = df_clients["Budget (MAD)"].sum()
                st.metric("Pipeline Financier Total", f"{total_budget:,.2f} MAD")
        with col3:
            if "Statut" in df_clients.columns:
                clients_actifs = len(df_clients[df_clients["Statut"] == "Validé / Réservé"])
                st.metric("Ventes / Réservations", clients_actifs)
                
        st.markdown("---")
        st.subheader("📋 Vue d'ensemble de la base de données")
        st.dataframe(df_clients, use_container_width=True)
    else:
        st.info("Aucune donnée client pour le moment. Rendez-vous dans l'onglet de gestion pour ajouter vos premiers prospects.")

elif menu == "👥 Gestion des Clients & Leads":
    st.subheader("✍️ Ajouter un nouveau client ou lead")
    
    with st.form("ajout_client_form"):
        col1, col2 = st.columns(2)
        with col1:
            nom = st.text_input("Nom & Prénom du Client")
            telephone = st.text_input("Téléphone")
            email = st.text_input("Email")
        with col2:
            projet = st.selectbox("Projet / Bien concerné", ["Hillside Residences (Bouskoura)", "Villa KM", "Les Collines de Dar Bouazza", "Autre"])
            statut = st.selectbox("Statut du Lead", ["Nouveau Prospect", "En cours de relance", "Visite planifiée", "Offre envoyée", "Validé / Réservé", "Perdu"])
            budget = st.number_input("Budget estimé (en MAD)", min_value=0.0, step=50000.0)
            
        notes = st.text_area("Notes / Remarques commerciales")
        
        submit_button = st.form_submit_button(label="Enregistrer le client")
        
        if submit_button:
            if nom:
                new_row = {
                    "Nom Client": nom,
                    "Téléphone": telephone,
                    "Email": email,
                    "Projet / Villa": projet,
                    "Statut": statut,
                    "Budget (MAD)": budget,
                    "Notes": notes
                }
                # Ajout de la ligne au DataFrame existant
                df_clients = pd.concat([df_clients, pd.DataFrame([new_row])], ignore_index=True)
                save_data(df_clients)
                st.success(f"Le client **{nom}** a été ajouté avec succès !")
                st.rerun()
            else:
                st.warning("Veuillez renseigner au moins le nom du client.")

    st.markdown("---")
    st.subheader("📝 Modifier ou supprimer des enregistrements")
    if not df_clients.empty:
        edited_df = st.data_editor(df_clients, num_rows="dynamic", use_container_width=True)
        if st.button("💾 Sauvegarder les modifications"):
            save_data(edited_df)
            st.success("Modifications enregistrées dans le fichier Excel !")
            st.rerun()
