import streamlit as st
import pandas as pd
import gspread
import json

st.set_page_config(page_title="CRM - Hillside", page_icon="🏠", layout="wide")

st.title("🏠 Accueil & Hub Central - Hillside")
st.markdown("Tableau de bord de pilotage commercial connecté en direct avec votre Google Sheet.")

@st.cache_resource
def get_gspread_client():
    # Définition directe des identifiants sous forme de dictionnaire pour contourner le décalage d'horloge serveur
    creds_dict = {
  "type": "service_account",
  "project_id": "crm-hillside",
  "private_key_id": "3d0cd04ce02a82e3e10c18628a514be18c133bbe",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCzvbXii/kFfcOR\n93OF1ZBdCz3g9vh44s4Ur6mk/oo0iqvr48ysj9ZiEI6aE965PZatpOGmleG8TrzG\nDGJ7dLDZGIoJyFxh80aBYFL3Qrw8OJcZ0D9l/K66QonujhYcwmFKUTQJW6Fh+EwJ\nQNrCaHftZlij+M/uk6DY56uXK52YyPjpZEvqeSFP+MhXNaqcE4EIV+66zAmqFD3P\nF3k4pYj3JkYc3HZAaYHq3bHaxoDNZb0HlRkzez3iDar7UnLE/L8TxKCVaKr5r1dS\nwnSEoZ80+xZIscJrQpQ36oCLz2IIcWOggk6w9gUuNps/IpqdWLjdovZtf+9rDCL+\nMw1FujOlAgMBAAECggEAKm/P3f9oycpd7pxjLtXTuhgRocgFtbrs55bq4Y0Ece/M\ntArKc/KY7fnVRV0MIKXBKdnMOhHYpGZdrQ/W/6l3KKahzrn6z9iJjLnClOTxfvFo\ngXx4v3pXgZv8p3FuGJevM0X1WIFFGtWuFCA3Fh/msssPghVYeI5r81JNYcl3Mc9a\nImo5AETNwoSDuy2hRvQgW6SagAmy4jGxDOL2u+qeIWM/3y4dN0103UY/WVeBq6rM\nfd58owxJIXlw6/4Ph/glgzYLhRGYwWglqvh65mO69dGMZ2O5Wyjrj+KVbETyBn0r\n4yya9zit7K6w3G0e73JjCP9rt2kTewUrpC7yzwUA8QKBgQDZYOtVa2xThlUnKSP+\nTG5moOKoibjf1Te6MuPpoB27Ge3svOjWmTl5k/DuWFmcS0xIZH5fxihrhnZb2UE2\neRSQWTFgDCCOozRkHr3fNVUpyZUEarp884lsVgWjZnWM3fbrtYPKNn6FhFotBYwg\nRSiC+4iYqqjul5FV0STf6KN69QKBgQDTrOofImHhWb4yPZPtDI/d9TN/mXKk4VSo\nVG0cDtory+ZagEJvmsSa7gznYpiFTqZuIujBPE2Ljf01o4m79GyMbpWjlEDFmzzJ\nSQkU1pyADHYHEHyzPUqIzNvx/69bBINr0F2cILODGUzMFLwdfp7UGhdG8dN39ncO\nOU0UiOPH8QKBgQCBMU7sV4nsIjxdp9/SdU5wBFJxZFjR/wC90x0NbrH8zy5Osv+W\n+z3jef+JSE0JzN6rb76D4VuIfEO/1uEhJNkJ6VilAtAetfSfRH4IOMRCr6G3HE91\nj7AvCPkU4gIfFJR6BkkrVzPe2ohVlaRwUWoAXsoSIeuA0ef4nvJReQwgOQKBgEMZ\nslHSTWLpVvCIMqdXbxe2Te/qDbyUrYh60uetEZ4Gkewocmb8kcxnH/INgDnp97o8\nGLqAyjgt95GUWOobDnnp6EeljG7Xygoou9c+PPgeex+8KzuKd2EsR6DMI6Nfp+V6\ndOI0P5OUZpLfE87pv6GO78gOirGhkJYStguHZWWRAoGAPZxBPJPc2iXB33r3t4hE\n5QICT78omcce+ArZUim9+PphcFCcLM4IUnc2T2jY+I+yxpI+YPC6UeaKhRyvZn4/\nMNezdLuJAv7BBN7IH1P6BSE4nwoDYqnoJfile755wTyPgocedJyHrwLJQl7TipjV\nOp8hH1S0afA+4wMfjN1gy7c=\n-----END PRIVATE KEY-----\n",
  "client_email": "hillside-bot@crm-hillside.iam.gserviceaccount.com",
  "client_id": "107770060658250416573",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/hillside-bot%40crm-hillside.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}

    return gspread.service_account_from_dict(creds_dict)

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
