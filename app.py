import streamlit as st
import pandas as pd
df = pd.read_excel("suivi_commercial.xlsx", sheet_name="Tableau de Bord")
import gspread

st.set_page_config(page_title="CRM - Hillside", page_icon="🏠", layout="wide")

st.title("🏠 Accueil & Hub Central - Hillside")
st.markdown("Tableau de bord de pilotage commercial connecté en direct avec votre Google Sheet.")

@st.cache_resource
def get_gspread_client():
    return gspread.service_account(filename="credentials.json")

try:
    client = get_gspread_client()
    
    sheet_name = "Suivi Commercial & Pipeline Leads - Immobilier"
    sh = client.open(sheet_name)
    
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
    st.error("⚠️ Impossible de récupérer les données du Google Sheet.")
    with st.expander("Voir les détails techniques de l'erreur"):
        st.write(e)
