import streamlit as st

# Configuration de la page
st.set_page_config(page_title="CRM - Hillside", page_icon="🏠", layout="wide")

st.title("🏠 Accueil & Hub Central - Hillside")
st.markdown("Tableau de bord de pilotage commercial connecté en direct avec votre Google Sheet.")

# Connexion à Google Sheets via les secrets Streamlit
try:
    conn = st.connection("gsheets", type="gsheets")
    
    # Menu latéral pour choisir l'onglet à afficher
    page = st.sidebar.selectbox("Navigation", ["Tableau de Bord", "Pipeline Commercial"])
    
    if page == "Tableau de Bord":
        st.subheader("📊 Tableau de Bord")
        # Lecture de l'onglet "Tableau de Bord"
        df_dashboard = conn.read(worksheet="Tableau de Bord", ttl=5)
        st.dataframe(df_dashboard, use_container_width=True)
        
    elif page == "Pipeline Commercial":
        st.subheader("📈 Pipeline Commercial")
        # Lecture de l'onglet "Pipeline Commercial"
        df_pipeline = conn.read(worksheet="Pipeline Commercial", ttl=5)
        st.dataframe(df_pipeline, use_container_width=True)

except Exception as e:
    st.error("⚠️ Impossible de récupérer les données du Google Sheet pour le moment.")
    st.info("Vérifiez que vos secrets TOML sont bien configurés et que le robot a accès au fichier.")
    with st.expander("Voir les détails techniques de l'erreur"):
        st.write(e)
