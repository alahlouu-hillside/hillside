from datetime import datetime, date
import os
import sqlite3
import urllib.parse
from google import genai
import gspread
import openpyxl
import pandas as pd
import streamlit as st

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="CRM Hillside - Beaurivage & Madison Immobilier",
    layout="wide",
    page_icon="🏢",
    initial_sidebar_state="expanded",
)

ADMIN_PRINCIPAL = "a.lahlou@beaurivage.ma"

# --- CHARGE GRAPHIQUE (BEAURIVAGE & HILLSIDE) ---
st.markdown(
    """
    <style>
        .main {
            background-color: #F8FAFC;
            font-family: 'Calibri', sans-serif;
        }
        h1, h2, h3 {
            color: #0F172A;
            font-weight: 700;
        }
        [data-testid="stSidebar"] {
            background-color: #0F172A;
            color: #FFFFFF;
        }
        [data-testid="stSidebar"] .stMarkdown p, [data-testid="stSidebar"] label {
            color: #E2E8F0;
        }
        .stButton>button {
            background-color: #1E3A8A;
            color: white;
            border-radius: 6px;
            font-weight: bold;
            border: none;
            padding: 8px 16px;
        }
        .stButton>button:hover {
            background-color: #1E40AF;
        }
        [data-testid="stMetricValue"] {
            color: #1E3A8A;
            font-weight: 700;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# --- INITIALISATION DE LA SÉCURITÉ & CHARGEMENT DES 86 LOTS ---
def init_db():
    conn = sqlite3.connect("crm_hillside.db", timeout=10.0)
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS utilisateurs_autorises (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE
        )
    """)
    c.execute(
        "INSERT OR IGNORE INTO utilisateurs_autorises (email) VALUES (?)",
        (ADMIN_PRINCIPAL,),
    )
    c.execute("""
        CREATE TABLE IF NOT EXISTS lots (
            code TEXT PRIMARY KEY, immeuble TEXT, typologie TEXT, etage TEXT,
            statut TEXT DEFAULT 'DISPONIBLE', vue TEXT, surface_totale REAL,
            surface_interieure REAL, terrasse REAL, jardin REAL, parking INTEGER,
            prix_plancher REAL, prix_catalogue REAL, remise_max REAL
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS reservations (
            id INTEGER PRIMARY KEY AUTOINCREMENT, code_lot TEXT, nom_client TEXT,
            cin_client TEXT, telephone TEXT, email TEXT, adresse TEXT,
            nom_client2 TEXT, cin_client2 TEXT, part_client1 REAL, part_client2 REAL,
            prix_catalogue REAL, remise REAL, prix_net REAL, 
            pourcentage_avance REAL, acompte_verse REAL, date_ech_av1 TEXT,
            avance2 REAL, date_ech_av2 TEXT,
            avance3 REAL, date_ech_av3 TEXT,
            reliquat REAL, date_ech_reliquat TEXT,
            mode_paiement TEXT, ref_paiement TEXT, banque TEXT, date_reservation TEXT,
            contrat_fichier TEXT
        )
    """)
    
    colonnes_a_verifier = [
        ("pourcentage_avance", "REAL"),
        ("acompte_verse", "REAL"),
        ("date_ech_av1", "TEXT"),
        ("mode_paiement_av1", "TEXT"),
        ("ref_paiement_av1", "TEXT"),
        ("banque_av1", "TEXT"),
        ("avance2", "REAL"),
        ("date_ech_av2", "TEXT"),
        ("mode_paiement_av2", "TEXT"),
        ("ref_paiement_av2", "TEXT"),
        ("banque_av2", "TEXT"),
        ("avance3", "REAL"),
        ("date_ech_av3", "TEXT"),
        ("mode_paiement_av3", "TEXT"),
        ("ref_paiement_av3", "TEXT"),
        ("banque_av3", "TEXT"),
        ("reliquat", "REAL"),
        ("date_ech_reliquat", "TEXT")
    ]
    for col_nom, col_type in colonnes_a_verifier:
        try:
            c.execute(f"ALTER TABLE reservations ADD COLUMN {col_nom} {col_type}")
        except sqlite3.OperationalError:
            pass

    fichiers_possibles = ["TDB_HILLSIDE.xls.xlsx", "TDB_HILLSIDE.xlsx", "TDB_HILLSIDE.xls"]
    fichier_trouve = next((f for f in fichiers_possibles if os.path.exists(f)), None)

    if fichier_trouve:
        try:
            df_excel = pd.read_excel(fichier_trouve, sheet_name="Pricing", skiprows=3)
            for _, row in df_excel.iterrows():
                code_val = row.iloc[0] if len(row) > 0 else None
                code = str(code_val).strip() if pd.notna(code_val) else ""
                
                exclus = ["nan", "total", "type", "prix moyen/m2", "ca previsionel", "ca pvp", "magasins"]
                if not code or code.lower() in exclus or not code.startswith(("A", "B", "C", "D", "Magasin")):
                    continue
                
                immeuble = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "Immeuble A"
                typologie = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ("Magasin" if code.startswith("Magasin") else "F3")
                etage = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else "RDC"
                
                etat_brut = str(row.iloc[4]).strip().upper() if pd.notna(row.iloc[4]) else "DISPONIBLE"
                statut = "RESERVÉ" if "VENDU" in etat_brut or "RESERV" in etat_brut else "DISPONIBLE"
                
                vue = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ""
                
                try:
                    surface_totale = float(row.iloc[6]) if pd.notna(row.iloc[6]) else 0.0
                except:
                    surface_totale = 0.0

                try:
                    surface_interieure = float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0
                except:
                    surface_interieure = 0.0

                try:
                    terrasse = float(row.iloc[8]) if pd.notna(row.iloc[8]) else 0.0
                except:
                    terrasse = 0.0

                try:
                    jardin = float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0
                except:
                    jardin = 0.0

                try:
                    parking = int(float(row.iloc[10])) if pd.notna(row.iloc[10]) else (0 if code.startswith("Magasin") else 1)
                except:
                    parking = 1

                try:
                    prix_plancher = float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0.0
                except:
                    prix_plancher = 0.0

                try:
                    prix_catalogue = float(row.iloc[12]) if pd.notna(row.iloc[12]) else 0.0
                except:
                    prix_catalogue = 0.0

                try:
                    remise_max = float(row.iloc[15]) if pd.notna(row.iloc[15]) else 0.0
                except:
                    remise_max = 0.0

                c.execute("""
                    INSERT OR IGNORE INTO lots (
                        code, immeuble, typologie, etage, statut, vue, surface_totale, 
                        surface_interieure, terrasse, jardin, parking, prix_plancher, prix_catalogue, remise_max
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (code, immeuble, typologie, etage, statut, vue, surface_totale, surface_interieure, terrasse, jardin, parking, prix_plancher, prix_catalogue, remise_max))
            conn.commit()
        except Exception as e:
            print(f"Erreur d'importation automatique : {e}")

    conn.commit()
    conn.close()


init_db()


def verifier_acces(email):
    with sqlite3.connect("crm_hillside.db", timeout=10.0) as conn:
        res = conn.execute(
            "SELECT email FROM utilisateurs_autorises WHERE email = ?",
            (email.strip().lower(),),
        ).fetchone()
    return res is not None


if "user_email" not in st.session_state:
    st.session_state.user_email = ADMIN_PRINCIPAL

if "admin_authentifie" not in st.session_state:
    st.session_state.admin_authentifie = False


# --- MOTEUR DE SYNCHRONISATION GOOGLE SHEETS ---
@st.cache_data(ttl=60)
def synchroniser_google_sheet():
    credentials_file = "credentials.json"
    if os.path.exists(credentials_file):
        try:
            gc = gspread.service_account(filename=credentials_file)
            sh = gc.open("Suivi Commercial & Pipeline Leads - Immobilier")

            worksheet = None
            for ws in sh.worksheets():
                if ws.title.strip().lower() == "pipeline commercial":
                    worksheet = ws
                    break
            if worksheet is None:
                worksheet = sh.get_worksheet(0)

            lignes = worksheet.get_all_values()
            if len(lignes) > 1:
                en_tetes = [
                    str(h).strip() if h != "" else f"Col_{i}"
                    for i, h in enumerate(lignes[0])
                ]
                df_gs = pd.DataFrame(lignes[1:], columns=en_tetes)
                df_gs = df_gs.loc[:, ~df_gs.columns.str.startswith("Col_")]
                df_gs = df_gs.fillna("")
                df_gs = df_gs[(df_gs != "").any(axis=1)]
                return df_gs
        except Exception as e:
            print(f"Erreur Google Sheet : {e}")
            return pd.DataFrame()
    return pd.DataFrame()


def mettre_a_jour_google_sheet_complet(df_modifie):
    credentials_file = "credentials.json"
    if os.path.exists(credentials_file):
        try:
            gc = gspread.service_account(filename=credentials_file)
            sh = gc.open("Suivi Commercial & Pipeline Leads - Immobilier")
            
            worksheet = None
            for ws in sh.worksheets():
                if ws.title.strip().lower() == "pipeline commercial":
                    worksheet = ws
                    break
            if worksheet is None:
                worksheet = sh.get_worksheet(0)

            df_propre = df_modifie.copy()
            df_propre = df_propre.fillna("").astype(str)
            df_propre = df_propre.replace("nan", "").replace("NaT", "").replace("None", "")

            data_a_ecrire = [df_propre.columns.tolist()] + df_propre.values.tolist()
            worksheet.clear()
            worksheet.update(data_a_ecrire)
            st.cache_data.clear()
            return True, "Succès"
        except Exception as e:
            return False, str(e)
    return False, "Fichier credentials.json introuvable"


def ajouter_lead_google_sheet(nouvelles_donnees_dict):
    try:
        gc = gspread.service_account(filename="credentials.json")
        sh = gc.open("Suivi Commercial & Pipeline Leads - Immobilier")
        
        worksheet = None
        for ws in sh.worksheets():
            if ws.title.strip().lower() == "pipeline commercial":
                worksheet = ws
                break
        if worksheet is None:
            worksheet = sh.get_worksheet(0)

        en_tetes = worksheet.row_values(1)
        id_col = next((h for h in en_tetes if h.lower() in ["id", "n°", "code", "référence"]), en_tetes[0])
        id_prospect = str(nouvelles_donnees_dict.get(id_col, ""))
        
        if id_prospect:
            cellule_id = worksheet.find(id_prospect)
            if cellule_id:
                return False, "Doublon détecté : cet ID existe déjà !"

        nouvelle_ligne = [str(nouvelles_donnees_dict.get(h, "")) for h in en_tetes]
        worksheet.append_row(nouvelle_ligne)
        st.cache_data.clear()
        return True, "Succès"
    except Exception as e:
        return False, str(e)


def modifier_lead_par_id(id_cible, id_col_name, dictionnaire_modifications):
    try:
        gc = gspread.service_account(filename="credentials.json")
        sh = gc.open("Suivi Commercial & Pipeline Leads - Immobilier")
        
        worksheet = None
        for ws in sh.worksheets():
            if ws.title.strip().lower() == "pipeline commercial":
                worksheet = ws
                break
        if worksheet is None:
            worksheet = sh.get_worksheet(0)

        en_tetes = worksheet.row_values(1)
        if id_col_name in en_tetes:
            cell = worksheet.find(str(id_cible))
            if cell:
                row_num = cell.row
                for col_name, nouvelle_valeur in dictionnaire_modifications.items():
                    if col_name in en_tetes:
                        col_idx = en_tetes.index(col_name) + 1
                        worksheet.update_cell(row_num, col_idx, str(nouvelle_valeur))
                st.cache_data.clear()
                return True
    except Exception as e:
        st.error(f"Erreur de modification sur le Google Sheet : {e}")
    return False


def supprimer_lead_par_id(id_cible, id_col_name):
    try:
        gc = gspread.service_account(filename="credentials.json")
        sh = gc.open("Suivi Commercial & Pipeline Leads - Immobilier")
        
        worksheet = None
        for ws in sh.worksheets():
            if ws.title.strip().lower() == "pipeline commercial":
                worksheet = ws
                break
        if worksheet is None:
            worksheet = sh.get_worksheet(0)

        cell = worksheet.find(str(id_cible))
        if cell:
            worksheet.delete_rows(cell.row)
            st.cache_data.clear()
            return True
    except Exception as e:
        st.error(f"Erreur de suppression sur le Google Sheet : {e}")
    return False


def get_df(table):
    if table == "pipeline_leads_crm":
        df_gs = synchroniser_google_sheet()
        if df_gs is not None and not df_gs.empty:
            return df_gs.fillna("")
        else:
            return pd.DataFrame()
    
    with sqlite3.connect("crm_hillside.db", timeout=10.0) as conn:
        df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
    return df.fillna("")


def trouver_plan(code_lot):
    dossier = "Hillside Fiches Plan 2d"
    if not os.path.exists(dossier):
        return None
    code_net = str(code_lot).strip().upper()
    for root, _, files in os.walk(dossier):
        for f in files:
            if code_net in f.upper():
                return os.path.join(root, f)
    return None


# --- GESTION DES CLÉS API ---
if "saved_api_keys" not in st.session_state:
    cle_defaut = ""
    try:
        cle_defaut = st.secrets["api_keys"]["commercial"]
    except Exception:
        cle_defaut = ""
    st.session_state.saved_api_keys = {"Principale": cle_defaut}

if "active_api_key_name" not in st.session_state:
    st.session_state.active_api_key_name = "Principale"

active_api_key = st.session_state.saved_api_keys.get(st.session_state.active_api_key_name, "")


# --- BARRE LATÉRALE DE NAVIGATION ---
logo_hillside = "logo hillside bleu.png"
if os.path.exists(logo_hillside):
    st.sidebar.image(logo_hillside, use_container_width=True)
else:
    st.sidebar.markdown("<h3 style='color:white; text-align:center;'>HILLSIDE RESIDENCE</h3>", unsafe_allow_html=True)

st.sidebar.markdown(
    """
    <div style="text-align: center; margin: -5px 0 5px 0;">
        <span style="font-size: 13px; color: #94A3B8; font-style: italic;">By</span>
    </div>
    """,
    unsafe_allow_html=True
)

logo_beaurivage = "logo beaurivage blanc.png"
if os.path.exists(logo_beaurivage):
    st.sidebar.image(logo_beaurivage, use_container_width=True, width=140)
else:
    st.sidebar.markdown("<p style='color:#94A3B8; text-align:center; font-size:12px;'>BEAURIVAGE</p>", unsafe_allow_html=True)

st.sidebar.markdown("---")

with st.sidebar.expander("👤 Mon Compte Administrateur", expanded=False):
    user_actuel = st.text_input("E-mail connecté :", value=st.session_state.user_email)
    if user_actuel != st.session_state.user_email:
        st.session_state.user_email = user_actuel.strip().lower()
        st.rerun()

if verifier_acces(st.session_state.user_email):
    st.sidebar.markdown(
        """
        <div style="background-color: #064E3B; border: 1px solid #10B981; padding: 8px; border-radius: 6px; text-align: center; margin-bottom: 10px;">
            <span style="color: #6EE7B7; font-weight: bold; font-size: 12px;">✅ Administrateur Actif</span>
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.sidebar.markdown(
        """
        <div style="background-color: #7F1D1D; border: 1px solid #EF4444; padding: 8px; border-radius: 6px; text-align: center; margin-bottom: 10px;">
            <span style="color: #FCA5A5; font-weight: bold; font-size: 12px;">⛔ Accès non autorisé</span>
        </div>
        """,
        unsafe_allow_html=True
    )

if st.sidebar.button("🔄 Rafraîchir les Données", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.markdown("<p style='color: #94A3B8; font-weight: bold; font-size: 11px; text-transform: uppercase; margin-bottom: 5px;'>Navigation Principale</p>", unsafe_allow_html=True)

menu = st.sidebar.radio(
    "Aller vers :",
    [
        "🏠 Accueil & Hub Central",
        "🤖 Assistant AI & Communication",
        "👥 Portefeuille Prospects & Clients",
        "🏢 Catalogue & Stocks Immobiliers",
        "✍️ Espace Ventes & Quittances",
        "📊 Chiffre d'Affaires & Trésorerie",
        "⚙️ Paramètres & Analyses",
    ],
    label_visibility="collapsed"
)

df_lots = get_df("lots")


# --- ROUTAGE DES MODULES ---

if menu == "🏠 Accueil & Hub Central":
    st.title("🏠 Accueil & Hub Central - Hillside")
    st.caption("Tableau de bord de pilotage commercial connecté en direct avec votre Google Sheet[cite: 1].")
    st.markdown("---")

    with st.spinner("Chargement des indicateurs Google Sheets..."):
        df_leads = get_df("pipeline_leads_crm")

    if df_leads is None or df_leads.empty:
        st.warning("⚠️ Impossible de récupérer les données du Google Sheet pour le moment.")
    else:
        total_leads = len(df_leads)
        nb_bureau_vente = len(df_leads[df_leads.astype(str).apply(lambda x: x.str.contains("Bureau de vente", case=False)).any(axis=1)])
        nb_chauds_qualifies = len(df_leads[df_leads.astype(str).apply(lambda x: x.str.contains("Chaud|Qualifié", case=False)).any(axis=1)])
        nb_offres = len(df_leads[df_leads.astype(str).apply(lambda x: x.str.contains("Offre", case=False)).any(axis=1)])
        nb_relances_retard = len(df_leads[df_leads.astype(str).apply(lambda x: x.str.contains("retard|en retard", case=False)).any(axis=1)])

        st.markdown("##### 🔍 Sélectionnez une catégorie pour filtrer instantanément le tableau :")
        filtre_clique = st.radio(
            "Filtrer le tableau par :",
            ["Tous les prospects", "Bureau de vente", "Opportunités / Chauds", "Offres", "En retard"],
            horizontal=True,
            label_visibility="collapsed"
        )

        b1, b2, b3, b4, b5 = st.columns(5)
        b1.metric("PORTEFEUILLE", f"{total_leads}")
        b2.metric("BUREAU VENTE", f"{nb_bureau_vente}")
        b3.metric("OPPORTUNITÉS", f"{nb_chauds_qualifies}")
        b4.metric("OFFRES", f"{nb_offres}")
        b5.metric("EN RETARD", f"{nb_relances_retard}")

        st.markdown("---")

        col_chart1, col_chart2 = st.columns(2)
        with col_chart1:
            st.markdown("##### 📈 Volume par Statut / Étape")
            col_statut_candidate = next((c for c in df_leads.columns if any(k in c.lower() for k in ["statut", "étape", "etape", "qualification"])), None)
            if col_statut_candidate:
                counts_statut = df_leads[col_statut_candidate].replace("", "Non renseigné").value_counts()
                st.bar_chart(counts_statut, height=180)
            else:
                st.info("Colonne de statut non détectée.")

        with col_chart2:
            st.markdown("##### 📊 Répartition par Typologie / Source")
            col_type_candidate = next((c for c in df_leads.columns if any(k in c.lower() for k in ["typologie", "source", "type", "bien"])), None)
            if col_type_candidate:
                counts_type = df_leads[col_type_candidate].replace("", "Non renseigné").value_counts()
                st.bar_chart(counts_type, height=180)
            else:
                st.info("Colonne de typologie non détectée.")

        st.markdown("---")

        df_filtre = df_leads.copy()
        if filtre_clique == "Bureau de vente":
            df_filtre = df_filtre[df_filtre.astype(str).apply(lambda x: x.str.contains("Bureau de vente", case=False)).any(axis=1)]
        elif filtre_clique == "Opportunités / Chauds":
            df_filtre = df_filtre[df_filtre.astype(str).apply(lambda x: x.str.contains("Chaud|Qualifié", case=False)).any(axis=1)]
        elif filtre_clique == "Offres":
            df_filtre = df_filtre[df_filtre.astype(str).apply(lambda x: x.str.contains("Offre", case=False)).any(axis=1)]
        elif filtre_clique == "En retard":
            df_filtre = df_filtre[df_filtre.astype(str).apply(lambda x: x.str.contains("retard|en retard", case=False)).any(axis=1)]

        st.info(f"📁 **Affichage actif :** {filtre_clique} ({len(df_filtre)} client(s) trouvé(s))[cite: 1].")
        
        column_configurations = {}
        for col in df_leads.columns:
            col_lower = col.lower()
            if any(keyword in col_lower for keyword in ["commentaire", "note", "observation", "remarque", "description", "suivi"]):
                column_configurations[col] = st.column_config.TextColumn(col, required=False)
            elif any(keyword in col_lower for keyword in ["statut", "étape", "etape", "source", "type", "commercial", "assigné"]):
                valeurs_uniques = [str(v) for v in df_leads[col].unique() if str(v).strip() != ""]
                column_configurations[col] = st.column_config.SelectboxColumn(
                    col,
                    options=valeurs_uniques if valeurs_uniques else ["Option 1", "Option 2"],
                    required=False,
                )

        df_edited = st.data_editor(
            df_filtre,
            column_config=column_configurations,
            num_rows="dynamic",
            use_container_width=True,
            key="editor_pipeline_accueil_fixed"
        )

        if st.button("💾 Enregistrer et synchroniser les modifications sur Google Sheet", use_container_width=True):
            with st.spinner("Synchronisation avec Google Sheets en cours..."):
                df_complet_maj = df_leads.copy()
                for idx, row in df_edited.iterrows():
                    if idx in df_complet_maj.index:
                        df_complet_maj.loc[idx] = row
                
                succes, msg = mettre_a_jour_google_sheet_complet(df_complet_maj)
                if succes:
                    st.success("✅ C'est OK ! Les modifications ont été enregistrées et synchronisées avec succès sur le Google Sheet.")
                    st.balloons()
                else:
                    st.error(f"❌ Échec de la synchronisation : {msg}")


elif menu == "🤖 Assistant AI & Communication":
    st.title("🤖 Assistant Commercial Intelligent & Communication")
    st.caption("Discutez avec l'assistant IA et contactez directement vos prospects par WhatsApp ou E-mail.")
    st.markdown("---")

    if not verifier_acces(st.session_state.user_email):
        st.error("⛔ Accès non autorisé.")
    else:
        df_leads_chat = get_df("pipeline_leads_crm")
        
        if not df_leads_chat.empty:
            client_col = next((c for c in df_leads_chat.columns if "client" in c.lower() or "nom" in c.lower() or "prospect" in c.lower()), df_leads_chat.columns[0])
            tel_cols = [c for c in df_leads_chat.columns if "tele" in c.lower() or "phone" in c.lower() or "gsm" in c.lower()]
            tel_col_name = tel_cols[0] if tel_cols else df_leads_chat.columns[0]
            email_cols = [c for c in df_leads_chat.columns if "mail" in c.lower() or "courriel" in c.lower()]
            email_col_name = email_cols[0] if email_cols else None

            col_sel1, _ = st.columns(2)
            with col_sel1:
                client_cible = st.selectbox(
                    "🎯 Cibler un prospect par son Nom :",
                    ["Aucun (Discussion générale)"] + df_leads_chat[client_col].tolist(),
                )

            tel_cible = ""
            email_cible = ""
            if client_cible != "Aucun (Discussion générale)":
                row_c = df_leads_chat[df_leads_chat[client_col] == client_cible].iloc[0]
                tel_cible = str(row_c.get(tel_col_name, ""))
                email_cible = str(row_c.get(email_col_name, "")) if email_col_name else "Non renseigné"
                st.info(f"👤 **Prospect :** {client_cible} | 📞 **Téléphone :** {tel_cible} | ✉️ **E-mail :** {email_cible}")

    st.markdown("---")
    col_chat, col_action = st.columns([3, 2])
    with col_chat:
        st.subheader("💬 Discussion avec l'Assistant")
        if "messages_chat" not in st.session_state:
            st.session_state.messages_chat = [{
                "role": "assistant",
                "content": "Bonjour Directeur ! Assistant Beaurivage prêt.",
            }]

        for msg in st.session_state.messages_chat:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])

        user_prompt = st.chat_input("Votre consigne...")
        if user_prompt:
            if not active_api_key.strip():
                st.error("⚠️ Clé API requise (configurable dans les Paramètres).")
            else:
                st.session_state.messages_chat.append({"role": "user", "content": user_prompt})
                with st.chat_message("user"):
                    st.write(user_prompt)
                with st.chat_message("assistant"):
                    with st.spinner("Analyse..."):
                        try:
                            ai_client = genai.Client(api_key=active_api_key.strip())
                            response = ai_client.models.generate_content(
                                model="gemini-3.8-flash", contents=user_prompt
                            )
                            reponse_ia = response.text
                            st.write(reponse_ia)
                            st.session_state.messages_chat.append({"role": "assistant", "content": reponse_ia})
                            st.session_state.dernier_message_genere = reponse_ia
                        except Exception as e:
                            st.error(f"Erreur : {e}")

    with col_action:
        st.subheader("📱 Éditeur & Actions de Contact")
        msg_defaut = st.session_state.get(
            "dernier_message_genere",
            "Bonjour, suite à votre intérêt pour la Résidence Hillside...",
        )
        message_modifie = st.text_area("Message / Objet :", value=msg_defaut, height=180)

        if 'client_cible' in locals() and client_cible != "Aucun (Discussion générale)" and 'tel_cible' in locals() and tel_cible and tel_cible.lower() != "nan" and tel_cible.strip() != "":
            num_net = tel_cible.replace(" ", "").replace("+", "").replace("-", "")
            if num_net.startswith("0"):
                num_net = "212" + num_net[1:]
            whatsapp_url = f"https://wa.me/{num_net}?text={urllib.parse.quote(message_modifie)}"
            st.markdown(
                f"""
                    <a href="{whatsapp_url}" target="_blank">
                        <button style="width:100%; background-color:#25D366; color:white; padding:10px; border:none; border-radius:5px; font-weight:bold; margin-bottom: 8px;">
                            🚀 Envoyer sur WhatsApp
                        </button>
                    </a>
                """,
                unsafe_allow_html=True,
            )

        if 'client_cible' in locals() and client_cible != "Aucun (Discussion générale)" and 'email_cible' in locals() and email_cible and email_cible.lower() != "nan" and email_cible.strip() != "" and email_cible != "Non renseigné":
            email_subject = urllib.parse.quote("Résidence Hillside - Suivi Commercial")
            email_body = urllib.parse.quote(message_modifie)
            email_url = f"mailto:{email_cible}?subject={email_subject}&body={email_body}"
            st.markdown(
                f"""
                    <a href="{email_url}" target="_blank">
                        <button style="width:100%; background-color:#0078D4; color:white; padding:10px; border:none; border-radius:5px; font-weight:bold;">
                            📧 Envoyer un E-mail
                        </button>
                    </a>
                """,
                unsafe_allow_html=True,
            )


elif menu == "👥 Portefeuille Prospects & Clients":
    st.title("👥 Portefeuille Prospects & Clients")
    st.caption("Gérez votre pipeline commercial, ajoutez de nouveaux leads ou mettez à jour vos contacts.")
    st.markdown("---")

    if not verifier_acces(st.session_state.user_email):
        st.error("⛔ Accès non autorisé.")
    else:
        tab_ajout, tab_pipe = st.tabs([
            "➕ Ajouter un Lead", 
            "📊 Pipeline & Validation"
        ])

        with tab_ajout:
            st.subheader("Enregistrement d'un nouveau prospect")
            df_leads_init = get_df("pipeline_leads_crm")
            with st.form("form_ajout_lead", clear_on_submit=True):
                nouvelles_donnees_ajout = {}
                cols_ajout_ui = st.columns(3)
                
                if not df_leads_init.empty:
                    id_col = next((h for h in df_leads_init.columns if h.lower() in ["id", "n°", "code"]), df_leads_init.columns[0])
                    prochain_id = 1
                    try:
                        valeurs_id_numeriques = [int(v) for v in df_leads_init[id_col].unique() if str(v).strip().isdigit()]
                        if valeurs_id_numeriques:
                            prochain_id = max(valeurs_id_numeriques) + 1
                    except Exception:
                        prochain_id = len(df_leads_init) + 1

                    date_du_jour = datetime.now().strftime("%d/%m/%Y")

                    for idx, col_name in enumerate(df_leads_init.columns):
                        valeurs_possibles = sorted([str(v) for v in df_leads_init[col_name].unique() if str(v).strip() != ""])
                        with cols_ajout_ui[idx % 3]:
                            if col_name.lower() in ["id", "n°", "code"]:
                                nouvelles_donnees_ajout[col_name] = st.text_input(f"{col_name}", value=str(prochain_id), key=f"add_{col_name}")
                            elif any(k in col_name.lower() for k in ["date", "contact", "alerte", "relance"]):
                                nouvelles_donnees_ajout[col_name] = st.text_input(f"{col_name} (JJ/MM/AAAA)", value=date_du_jour, key=f"add_{col_name}")
                            elif 1 < len(valeurs_possibles) <= 15:
                                nouvelles_donnees_ajout[col_name] = st.selectbox(f"{col_name}", valeurs_possibles, key=f"add_sel_{col_name}")
                            else:
                                nouvelles_donnees_ajout[col_name] = st.text_input(f"{col_name}", value="", key=f"add_{col_name}")

                submit_ajout = st.form_submit_button("🚀 Enregistrer le lead sur le Google Sheet", use_container_width=True)
                if submit_ajout:
                    succes_ajout, msg_retour = ajouter_lead_google_sheet(nouvelles_donnees_ajout)
                    if succes_ajout:
                        st.success("✅ C'est OK ! Nouveau lead enregistré et synchronisé avec succès sur le Google Sheet.")
                    else:
                        st.error(f"❌ Échec de l'ajout : {msg_retour}")

        with tab_pipe:
            st.subheader("Gestion et modification des leads existants (Google Sheet)")
            df_leads = get_df("pipeline_leads_crm")
            if df_leads.empty:
                st.warning("Aucun lead trouvé.")
            else:
                id_col = next((h for h in df_leads.columns if h.lower() in ["id", "n°", "code"]), df_leads.columns[0])
                client_col_act = next((c for c in df_leads.columns if "client" in c.lower() or "nom" in c.lower()), df_leads.columns[1] if len(df_leads.columns) > 1 else df_leads.columns[0])

                try:
                    ids_disponibles = sorted([int(v) for v in df_leads[id_col].unique() if str(v).strip().isdigit()])
                    ids_disponibles = [str(v) for v in ids_disponibles]
                except Exception:
                    ids_disponibles = sorted([str(v) for v in df_leads[id_col].unique() if str(v).strip() != ""])
                
                selected_id = st.selectbox("Sélectionner le numéro ID du prospect à modifier :", ids_disponibles)
                if selected_id:
                    row_lead = df_leads[df_leads[id_col].astype(str) == str(selected_id)].iloc[0]
                    nom_client_associe = row_lead[client_col_act]
                    
                    st.info(f"📋 **Client sélectionné (ID : {selected_id}) :** {nom_client_associe}")
                    with st.form("form_validation_modifications"):
                        st.markdown(f"#### 📝 Modification de : **{nom_client_associe}**")
                        nouvelles_donnees = {}
                        cols_ui = st.columns(3)
                        
                        for idx, col_name in enumerate(df_leads.columns):
                            val_actuelle = str(row_lead[col_name]) if pd.notna(row_lead[col_name]) else ""
                            valeurs_possibles = sorted([str(v) for v in df_leads[col_name].unique() if str(v).strip() != ""])
                            with cols_ui[idx % 3]:
                                if 1 < len(valeurs_possibles) <= 15 and col_name.lower() != id_col.lower() and not any(k in col_name.lower() for k in ["date", "contact", "alerte", "relance"]):
                                    if val_actuelle not in valeurs_possibles:
                                        valeurs_possibles.insert(0, val_actuelle)
                                    try:
                                        index_defaut = valeurs_possibles.index(val_actuelle)
                                    except ValueError:
                                        index_defaut = 0
                                    nouvelles_donnees[col_name] = st.selectbox(f"{col_name}", valeurs_possibles, index=index_defaut, key=f"sel_{col_name}_{selected_id}")
                                else:
                                    nouvelles_donnees[col_name] = st.text_input(f"{col_name}", value=val_actuelle, key=f"txt_{col_name}_{selected_id}")
                        
                        st.markdown("---")
                        col_btn1, col_btn2 = st.columns(2)
                        with col_btn1:
                            submit_mod = st.form_submit_button("💾 Mettre à jour sur Google Sheet", use_container_width=True)
                        with col_btn2:
                            submit_del = st.form_submit_button("🗑️ Supprimer ce lead définitivement", use_container_width=True)
                            
                        if submit_mod:
                            succes = modifier_lead_par_id(selected_id, id_col, nouvelles_donnees)
                            if succes:
                                st.success("✅ C'est OK ! Modifications enregistrées et synchronisées avec succès.")
                            else:
                                st.error("❌ Échec de la mise à jour.")
                        if submit_del:
                            succes_sup = supprimer_lead_par_id(selected_id, id_col)
                            if succes_sup:
                                st.success(f"🗑️ C'est OK ! Le lead (ID : {selected_id}) a été supprimé avec succès.")
                            else:
                                st.error("❌ Échec de la suppression.")

                    with st.expander("👀 Afficher le tableau complet des leads en direct"):
                        st.dataframe(df_leads, use_container_width=True)


elif menu == "🏢 Catalogue & Stocks Immobiliers":
    st.title("🏢 Catalogue & Stocks Immobiliers")
    st.caption("Consultez l'état des lots (disponibles/réservés), filtrez par critères et téléchargez leurs plans 2D.")
    st.markdown("---")

    if not df_lots.empty:
        st.markdown("#### 🔍 Filtres Dynamiques du Catalogue")
        f_col1, f_col2, f_col3, f_col4 = st.columns(4)

        immeubles_dispos = ["Tous"] + sorted(df_lots["immeuble"].dropna().unique().tolist())
        typologies_dispos = ["Tous"] + sorted(df_lots["typologie"].dropna().unique().tolist())
        etages_dispos = ["Tous"] + sorted(df_lots["etage"].dropna().unique().tolist())
        statuts_dispos = ["Tous"] + sorted(df_lots["statut"].dropna().unique().tolist())

        with f_col1:
            filtre_immeuble = st.selectbox("Immeuble :", immeubles_dispos)
        with f_col2:
            filtre_typologie = st.selectbox("Typologie :", typologies_dispos)
        with f_col3:
            filtre_etage = st.selectbox("Étage :", etages_dispos)
        with f_col4:
            filtre_statut = st.selectbox("Statut :", statuts_dispos)

        df_f = df_lots.copy()
        if filtre_immeuble != "Tous":
            df_f = df_f[df_f["immeuble"] == filtre_immeuble]
        if filtre_typologie != "Tous":
            df_f = df_f[df_f["typologie"] == filtre_typologie]
        if filtre_etage != "Tous":
            df_f = df_f[df_f["etage"] == filtre_etage]
        if filtre_statut != "Tous":
            df_f = df_f[df_f["statut"] == filtre_statut]

        st.markdown("---")

        df_lots_affiche = df_f.copy()
        df_lots_affiche.index = range(1, len(df_lots_affiche) + 1)

        for col_fin in ["prix_plancher", "prix_catalogue", "remise_max"]:
            if col_fin in df_lots_affiche.columns:
                df_lots_affiche[col_fin] = df_lots_affiche[col_fin].apply(
                    lambda x: f"{int(float(x)):,}".replace(",", " ") if pd.notna(x) and str(x).strip() != "" else x
                )

        def colorer_statut(val):
            if str(val).upper() == "DISPONIBLE":
                return "background-color: #D1E7DD; color: #0F5132; font-weight: bold;"
            elif "RESERV" in str(val).upper() or "VENDU" in str(val).upper():
                return "background-color: #F8D7DA; color: #842029; font-weight: bold;"
            return ""

        df_styled = df_lots_affiche.style.map(colorer_statut, subset=["statut"])
        st.dataframe(df_styled, use_container_width=True)
    else:
        st.warning("⚠️ Aucun appartement trouvé dans la base locale.")

    st.markdown("---")
    st.subheader("📥 Téléchargement des Plans 2D des Lots")
    tous_les_lots = df_lots["code"].tolist() if not df_lots.empty else []
    if tous_les_lots:
        lot_choisi_plan = st.selectbox("🔍 Choisir un lot pour récupérer son plan 2D :", tous_les_lots)
        if lot_choisi_plan:
            chemin_plan = trouver_plan(lot_choisi_plan)
            if chemin_plan and os.path.exists(chemin_plan):
                st.success(f"✅ C'est OK ! Plan trouvé pour le lot **{lot_choisi_plan}**.")
                with open(chemin_plan, "rb") as f_plan:
                    st.download_button(
                        label=f"⬇️ Télécharger le plan du lot {lot_choisi_plan}",
                        data=f_plan,
                        file_name=os.path.basename(chemin_plan),
                        mime="application/octet-stream",
                    )
            else:
                st.warning(f"⚠️ Aucun fichier de plan 2D n'a été trouvé pour le lot **{lot_choisi_plan}**.")
    else:
        st.info("Aucun lot disponible dans la base de données locale.")


elif menu == "✍️ Espace Ventes & Quittances":
    st.title("✍️ Espace Ventes & Quittances")
    st.caption("Enregistrez, modifiez, désistez ou switchez vos biens vendus en toute simplicité[cite: 1].")
    st.markdown("---")

    tab_res_new, tab_res_modif, tab_res_suivi = st.tabs([
        "📝 Nouvelle Réservation / Vente", 
        "✏️ Modifier / Gérer une Vente Existante", 
        "📁 Dossiers Ventes & Suivi"
    ])

    with tab_res_new:
        conn_res = sqlite3.connect("crm_hillside.db")
        try:
            df_lots_dispo = pd.read_sql_query("SELECT * FROM lots WHERE statut = 'DISPONIBLE'", conn_res)
        except Exception:
            df_lots_dispo = pd.DataFrame()
        conn_res.close()

        if df_lots_dispo.empty:
            st.warning("⚠️ Aucun lot n'est actuellement marqué comme 'DISPONIBLE' dans la base de données locale.")
        else:
            st.markdown("#### 🏠 1. Sélection du Lot")
            codes_lots = df_lots_dispo["code"].tolist()
            lot_selectionne = st.selectbox("Code du Lot :", codes_lots, label_visibility="collapsed")
            
            infos_lot = df_lots_dispo[df_lots_dispo["code"] == lot_selectionne].iloc[0]
            prix_cat_defaut = float(infos_lot.get("prix_catalogue", 0.0))
            prix_plancher_lot = float(infos_lot.get("prix_plancher", 0.0))
            
            prix_plancher_aff = f"{int(prix_plancher_lot):,}".replace(",", " ")
            st.info(f"**Immeuble :** {infos_lot.get('immeuble', '')} | **Typologie :** {infos_lot.get('typologie', '')} | **Étage :** {infos_lot.get('etage', '')} | **Surface :** {infos_lot.get('surface_totale', '')} m² | **Prix Plancher :** {prix_plancher_aff} MAD")

            st.markdown("---")
            st.markdown("#### 👤 2. Informations du Client (Acquéreur)")
            c1, c2 = st.columns(2)
            with c1:
                nom_client = st.text_input("Nom & Prénom du Client :", key="res_nom1")
                cin_client = st.text_input("CIN / Passeport :", key="res_cin1")
                telephone = st.text_input("Numéro de Téléphone :", key="res_tel")
            with c2:
                email = st.text_input("Adresse E-mail :", key="res_email")
                adresse = st.text_area("Adresse complète :", key="res_adr", height=108)

            with st.expander("👥 Co-acquéreur (Client 2 optionnel)", expanded=False):
                c3, c4 = st.columns(2)
                with c3:
                    nom_client2 = st.text_input("Nom & Prénom du Client 2 :", key="res_nom2")
                    cin_client2 = st.text_input("CIN / Passeport Client 2 :", key="res_cin2")
                with c4:
                    part_client1 = st.slider("Part Client 1 (%)", min_value=0, max_value=100, value=50, key="res_part1")
                    part_client2 = 100 - part_client1
                    st.info(f"Répartition : Client 1 ({part_client1}%) / Client 2 ({part_client2}%)")

            st.markdown("---")
            st.markdown("#### 💰 3. Conditions Financières & Remise")
            
            palier_remise = st.radio("Pas d'incrémentation de la remise :", [50, 100, 1000], horizontal=True, key="res_palier_remise")

            c5, c6, c7 = st.columns(3)
            with c5:
                st.markdown(f"**Prix Catalogue Initial (Fixe)**")
                prix_catalogue = st.number_input("Prix Catalogue (MAD) :", value=prix_cat_defaut, disabled=True, format="%.2f", key="res_px_cat_disabled", label_visibility="collapsed")
            with c6:
                st.markdown(f"**Remise accordée (MAD)**")
                remise = st.number_input("Remise accordée (MAD) :", value=0.0, step=float(palier_remise), min_value=0.0, format="%.2f", key="res_remise", label_visibility="collapsed")
            with c7:
                prix_net = prix_catalogue - remise
                prix_net_aff = f"{int(prix_net):,}".replace(",", " ")
                st.markdown(f"**Prix Net Vente (MAD)**")
                st.markdown(f"### {prix_net_aff}")

            alerte_plancher = False
            if prix_net < prix_plancher_lot:
                alerte_plancher = True
                st.error(f"🚨 **ALERTE PRIX PLANCHER DÉPASSÉ !** Le prix net ({prix_net_aff} MAD) est inférieur au prix plancher autorisé ({prix_plancher_aff} MAD).")

            st.markdown("---")
            st.markdown("#### 🧮 4. Ventilation des Paiements, Échéances & Rappels (Détails par Chèque / Avance)")
            
            ca1, ca2 = st.columns(2)
            with ca1:
                pourcentage_avance = st.selectbox("Choisir le % d'Avance Rapide :", [10, 20, 30, 50, 100, "Personnalisé"], key="res_select_pct")
            with ca2:
                if pourcentage_avance == "Personnalisé":
                    pct_val = st.number_input("Entrer le pourcentage (%) :", min_value=0.0, max_value=100.0, value=20.0, step=1.0, key="res_pct_perso")
                else:
                    pct_val = float(pourcentage_avance)

            acompte_calcule = (prix_net * pct_val) / 100.0

            if "last_pct" not in st.session_state or st.session_state.last_pct != pct_val or "last_prix_net" not in st.session_state or st.session_state.last_prix_net != prix_net:
                st.session_state.res_acompte_input = float(acompte_calcule)
                st.session_state.last_pct = pct_val
                st.session_state.last_prix_net = prix_net

            st.markdown("##### 📌 Avance 1 / Réservation")
            col_av1_1, col_av1_2, col_av1_3, col_av1_4 = st.columns(4)
            with col_av1_1:
                acompte_verse = st.number_input(f"Montant [{pct_val}%] (MAD) :", format="%.2f", key="res_acompte_input")
            with col_av1_2:
                date_ech_av1 = st.date_input("Date échéance :", value=datetime.now(), key="ech_av1")
            with col_av1_3:
                mode_paiement_av1 = st.selectbox("Mode :", ["Chèque", "Virement Bancaire", "Espèces", "Crédit Bancaire"], key="mode_av1")
            with col_av1_4:
                ref_av1 = st.text_input("N° de Chèque / Réf :", key="ref_av1")
            banque_av1 = st.text_input("Banque émettrice (Avance 1) :", key="banque_av1")

            st.markdown("##### 📌 Avance 2 (Optionnelle)")
            col_av2_1, col_av2_2, col_av2_3, col_av2_4 = st.columns(4)
            with col_av2_1:
                avance2 = st.number_input("Montant (MAD) :", value=0.0, step=1000.0, format="%.2f", key="res_av2")
            with col_av2_2:
                date_ech_av2 = st.date_input("Date échéance :", value=datetime.now(), key="ech_av2")
            with col_av2_3:
                mode_paiement_av2 = st.selectbox("Mode :", ["Chèque", "Virement Bancaire", "Espèces", "Crédit Bancaire"], key="mode_av2")
            with col_av2_4:
                ref_av2 = st.text_input("N° de Chèque / Réf :", key="ref_av2")
            banque_av2 = st.text_input("Banque émettrice (Avance 2) :", key="banque_av2")

            st.markdown("##### 📌 Avance 3 (Optionnelle)")
            col_av3_1, col_av3_2, col_av3_3, col_av3_4 = st.columns(4)
            with col_av3_1:
                avance3 = st.number_input("Montant (MAD) :", value=0.0, step=1000.0, format="%.2f", key="res_av3")
            with col_av3_2:
                date_ech_av3 = st.date_input("Date échéance :", value=datetime.now(), key="ech_av3")
            with col_av3_3:
                mode_paiement_av3 = st.selectbox("Mode :", ["Chèque", "Virement Bancaire", "Espèces", "Crédit Bancaire"], key="mode_av3")
            with col_av3_4:
                ref_av3 = st.text_input("N° de Chèque / Réf :", key="ref_av3")
            banque_av3 = st.text_input("Banque émettrice (Avance 3) :", key="banque_av3")

            total_verse = acompte_verse + avance2 + avance3
            reliquat = prix_net - total_verse

            st.markdown("---")
            col_rel_1, col_rel_2 = st.columns(2)
            with col_rel_1:
                st.markdown(f"**Reliquat à payer (Calculé)**")
                reliquat_aff = f"{int(reliquat):,}".replace(",", " ")
                st.markdown(f"### {reliquat_aff} MAD")
            with col_rel_2:
                date_ech_reliquat = st.date_input("Date échéance Reliquat :", value=datetime.now(), key="ech_reliquat")

            date_reservation = datetime.now().strftime("%d/%m/%Y")

            st.markdown("---")
            if st.button("💾 Valider et Enregistrer la Vente & Programmer les Rappels", use_container_width=True):
                if not nom_client.strip() or not cin_client.strip():
                    st.error("⚠️ Veuillez renseigner au moins le nom et le CIN du client.")
                elif alerte_plancher:
                    st.error("⛔ Enregistrement bloqué : Impossible de valider une vente en dessous du prix plancher sans validation de la direction.")
                else:
                    try:
                        conn_ins = sqlite3.connect("crm_hillside.db")
                        conn_ins.execute("""
                            INSERT INTO reservations (
                                code_lot, nom_client, cin_client, telephone, email, adresse,
                                nom_client2, cin_client2, part_client1, part_client2,
                                prix_catalogue, remise, prix_net, 
                                pourcentage_avance, acompte_verse, date_ech_av1, mode_paiement_av1, ref_paiement_av1, banque_av1,
                                avance2, date_ech_av2, mode_paiement_av2, ref_paiement_av2, banque_av2,
                                avance3, date_ech_av3, mode_paiement_av3, ref_paiement_av3, banque_av3,
                                reliquat, date_ech_reliquat, date_reservation
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, (
                            lot_selectionne, nom_client, cin_client, telephone, email, adresse,
                            nom_client2 if 'nom_client2' in locals() else "", 
                            cin_client2 if 'cin_client2' in locals() else "", 
                            part_client1 if 'part_client1' in locals() else 100.0, 
                            part_client2 if 'part_client2' in locals() else 0.0,
                            prix_catalogue, remise, prix_net, 
                            pct_val, 
                            acompte_verse, date_ech_av1.strftime("%Y-%m-%d"), mode_paiement_av1, ref_av1, banque_av1,
                            avance2, date_ech_av2.strftime("%Y-%m-%d"), mode_paiement_av2, ref_av2, banque_av2,
                            avance3, date_ech_av3.strftime("%Y-%m-%d"), mode_paiement_av3, ref_av3, banque_av3,
                            reliquat, date_ech_reliquat.strftime("%Y-%m-%d"), date_reservation
                        ))
                        conn_ins.execute("UPDATE lots SET statut = 'RESERVÉ' WHERE code = ?", (lot_selectionne,))
                        conn_ins.commit()
                        conn_ins.close()
                        st.success(f"✅ C'est OK ! Vente enregistrée avec succès pour l'appartement **{lot_selectionne}**. Le stock a basculé en RESERVÉ.")
                        st.balloons()
                    except Exception as e:
                        st.error(f"Erreur lors de l'enregistrement en base de données : {e}")

  with tab_res_modif:
        st.subheader("✏️ Modifier les chèques, dates, switcher de bien, ou gérer un désistement")
        df_ventes_existantes = get_df("reservations")
        if df_ventes_existantes.empty:
            st.info("Aucune vente enregistrée à modifier.")
        else:
            liste_ventes_labels = [f"ID: {r['id']} | Lot: {r['code_lot']} - Client: {r['nom_client']}" for _, r in df_ventes_existantes.iterrows()]
            vente_choisie_str = st.selectbox("Sélectionner la vente à modifier / gérer :", liste_ventes_labels)
            
            if vente_choisie_str:
                id_vente_sel = int(vente_choisie_str.split("|")[0].replace("ID:", "").strip())
                row_vente = df_ventes_existantes[df_ventes_existantes["id"] == id_vente_sel].iloc[0]
                
                st.markdown(f"---")
                st.info(f"📋 **Dossier sélectionné :** Lot **{row_vente['code_lot']}** | Client : **{row_vente['nom_client']}**")
                
                choix_action_admin = st.radio(
                    "Que souhaitez-vous faire pour ce dossier ?",
                    ["Modifier les chèques / dates de paiement", "Switcher de bien (Changer d'appartement)", "Désistement / Supprimer le client acheteur"],
                    horizontal=True
                )
                
                if choix_action_admin == "Modifier les chèques / dates de paiement":
                    with st.form("form_modif_paiement"):
                        st.markdown("#### 💳 Modification complète des détails de paiement et échéances")
                        
                        c_m1, c_m2 = st.columns(2)
                        with c_m1:
                            nouvel_acompte = st.number_input("Montant Avance 1 (MAD) :", value=float(row_vente['acompte_verse']), format="%.2f")
                            ref_m1 = st.text_input("Réf / N° Chèque Avance 1 :", value=str(row_vente.get('ref_paiement_av1', '')))
                            banque_m1 = st.text_input("Banque Avance 1 :", value=str(row_vente.get('banque_av1', '')))
                            date_m_av1 = st.text_input("Date échéance Avance 1 (YYYY-MM-DD) :", value=str(row_vente.get('date_ech_av1', '')))
                            
                            nouvelle_avance2 = st.number_input("Montant Avance 2 (MAD) :", value=float(row_vente.get('avance2', 0.0)), format="%.2f")
                            ref_m2 = st.text_input("Réf / N° Chèque Avance 2 :", value=str(row_vente.get('ref_paiement_av2', '')))
                            banque_m2 = st.text_input("Banque Avance 2 :", value=str(row_vente.get('banque_av2', '')))
                            date_m_av2 = st.text_input("Date échéance Avance 2 (YYYY-MM-DD) :", value=str(row_vente.get('date_ech_av2', '')))
                        with c_m2:
                            nouvelle_avance3 = st.number_input("Montant Avance 3 (MAD) :", value=float(row_vente.get('avance3', 0.0)), format="%.2f")
                            ref_m3 = st.text_input("Réf / N° Chèque Avance 3 :", value=str(row_vente.get('ref_paiement_av3', '')))
                            banque_m3 = st.text_input("Banque Avance 3 :", value=str(row_vente.get('banque_av3', '')))
                            date_m_av3 = st.text_input("Date échéance Avance 3 (YYYY-MM-DD) :", value=str(row_vente.get('date_ech_av3', '')))
                            
                            prix_net_actuel = float(row_vente.get('prix_net', 0.0))
                            nouveau_reliquat = prix_net_actuel - (nouvel_acompte + nouvelle_avance2 + nouvelle_avance3)
                            st.markdown(f"**Reliquat recalculé automatiquement : {int(nouveau_reliquat):,} MAD**".replace(",", " "))
                            
                            date_m_rel = st.text_input("Date échéance Reliquat (YYYY-MM-DD) :", value=str(row_vente['date_ech_reliquat']))
                            
                        submit_edit_paiement = st.form_submit_button("💾 Enregistrer les modifications de paiement", use_container_width=True)
                        if submit_edit_paiement:
                            conn_up = sqlite3.connect("crm_hillside.db")
                            conn_up.execute("""
                                UPDATE reservations SET acompte_verse = ?, ref_paiement_av1 = ?, banque_av1 = ?, date_ech_av1 = ?, 
                                avance2 = ?, ref_paiement_av2 = ?, banque_av2 = ?, date_ech_av2 = ?, 
                                avance3 = ?, ref_paiement_av3 = ?, banque_av3 = ?, date_ech_av3 = ?, 
                                reliquat = ?, date_ech_reliquat = ? WHERE id = ?
                            """, (nouvel_acompte, ref_m1, banque_m1, date_m_av1, nouvelle_avance2, ref_m2, banque_m2, date_m_av2, nouvelle_avance3, ref_m3, banque_m3, date_m_av3, nouveau_reliquat, date_m_rel, id_vente_sel))
                            conn_up.commit()
                            conn_up.close()
                            st.success("✅ C'est OK ! Modifications enregistrées avec succès et répercutées sur la trésorerie.")

                elif choix_action_admin == "Switcher de bien (Changer d'appartement)":
                    st.warning("⚠️ Le bien actuel sera libéré (remis en DISPONIBLE) et remplacé par un nouveau lot disponible.")
                    conn_l = sqlite3.connect("crm_hillside.db")
                    df_dispo_switch = pd.read_sql_query("SELECT * FROM lots WHERE statut = 'DISPONIBLE'", conn_l)
                    conn_l.close()
                    
                    if df_dispo_switch.empty:
                        st.error("Aucun autre lot disponible pour effectuer le switch.")
                    else:
                        nouveau_lot_choisi = st.selectbox("Sélectionner le nouveau lot d'attribution :", df_dispo_switch["code"].tolist())
                        if st.button("🔄 Confirmer le changement de bien (Switch)", use_container_width=True):
                            ancien_lot = row_vente['code_lot']
                            conn_sw = sqlite3.connect("crm_hillside.db")
                            conn_sw.execute("UPDATE lots SET statut = 'DISPONIBLE' WHERE code = ?", (ancien_lot,))
                            conn_sw.execute("UPDATE lots SET statut = 'RESERVÉ' WHERE code = ?", (nouveau_lot_choisi,))
                            conn_sw.execute("UPDATE reservations SET code_lot = ? WHERE id = ?", (nouveau_lot_choisi, id_vente_sel))
                            conn_sw.commit()
                            conn_sw.close()
                            st.success(f"✅ C'est OK ! Le client a bien été basculé de l'ancien lot **{ancien_lot}** vers le nouveau lot **{nouveau_lot_choisi}**.")

                elif choix_action_admin == "Désistement / Supprimer le client acheteur":
                    st.error("🚨 **Attention :** Cette action est irréversible. Elle va supprimer le dossier de vente, libérer l'appartement (qui redeviendra DISPONIBLE) et supprimer l'acheteur du CRM.")
                    if st.button("🗑️ Confirmer le Désistement et Supprimer le Dossier", use_container_width=True):
                        lot_a_liberer = row_vente['code_lot']
                        conn_del = sqlite3.connect("crm_hillside.db", timeout=10.0)
                        conn_del.execute("UPDATE lots SET statut = 'DISPONIBLE' WHERE code = ?", (lot_a_liberer,))
                        conn_del.execute("DELETE FROM reservations WHERE id = ?", (id_vente_sel,))
                        conn_del.commit()
                        conn_del.close()
                        st.success(f"🗑️ C'est OK ! Le dossier a été supprimé et l'appartement **{lot_a_liberer}** est de nouveau DISPONIBLE.")
                        st.rerun()

  with tab_res_suivi:
        st.subheader("📁 Tableau de Suivi des Ventes & Encaissements")
        df_r = get_df("reservations")
        if not df_r.empty:
            df_r_affiche = df_r.copy()
            df_r_affiche.index = range(1, len(df_r_affiche) + 1)
            for col_f in ["prix_catalogue", "remise", "prix_net", "acompte_verse", "avance2", "avance3", "reliquat"]:
                if col_f in df_r_affiche.columns:
                    df_r_affiche[col_f] = df_r_affiche[col_f].apply(
                        lambda x: f"{int(float(x)):,}".replace(",", " ") if pd.notna(x) and str(x).strip() != "" else x
                    )
            st.dataframe(df_r_affiche.astype(str), use_container_width=True)
        else:
            st.dataframe(df_r, use_container_width=True)


elif menu == "📊 Chiffre d'Affaires & Trésorerie":
    st.title("📊 Chiffre d'Affaires & Trésorerie")
    st.caption("Pilotez votre chiffre d'affaires, suivez les encaissements et consultez les chèques futurs.")
    st.markdown("---")

    df_sales = get_df("reservations")
    if df_sales.empty:
        st.info("Aucune vente ou réservation enregistrée pour le moment.")
    else:
        ca_potentiel = df_sales["prix_catalogue"].astype(float).sum()
        ca_net_valide = df_sales["prix_net"].astype(float).sum()
        total_encaisse = (
            df_sales["acompte_verse"].astype(float) +
            df_sales["avance2"].astype(float) +
            df_sales["avance3"].astype(float)
        ).sum()
        reste_a_encaisser = df_sales["reliquat"].astype(float).sum()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("CA POTENTIEL", f"{int(ca_potentiel):,}".replace(",", " ") + " MAD")
        k2.metric("CA NET SIGNÉ", f"{int(ca_net_valide):,}".replace(",", " ") + " MAD")
        k3.metric("TOTAL ENCAISSÉ", f"{int(total_encaisse):,}".replace(",", " ") + " MAD")
        k4.metric("RESTE À ENCAISSER", f"{int(reste_a_encaisser):,}".replace(",", " ") + " MAD")

        st.markdown("---")
        st.subheader("💳 Suivi des Chèques Futurs (À encaisser dans l'avenir)")
        
        date_jour_str = datetime.now().strftime("%Y-%m-%d")
        liste_cheques_A_venir = []
        
        for _, row in df_sales.iterrows():
            lot = row["code_lot"]
            client = row["nom_client"]
            
            for av_col, ref_col, banque_col, ech_col in [
                ("acompte_verse", "ref_paiement_av1", "banque_av1", "date_ech_av1"),
                ("avance2", "ref_paiement_av2", "banque_av2", "date_ech_av2"),
                ("avance3", "ref_paiement_av3", "banque_av3", "date_ech_av3")
            ]:
                montant_av = float(row.get(av_col, 0.0))
                date_ech = str(row.get(ech_col, ""))
                ref_chq = str(row.get(ref_col, "N/A"))
                banque_chq = str(row.get(banque_col, "N/A"))
                
                if montant_av > 0 and date_ech >= date_jour_str:
                    liste_cheques_A_venir.append({
                        "Lot": lot, "Client": client, "Banque": banque_chq, "N° Chèque / Réf": ref_chq, 
                        "Montant (MAD)": int(montant_av), "Date d'échéance": date_ech
                    })

        if liste_cheques_A_venir:
            df_cq_affiche = pd.DataFrame(liste_cheques_A_venir)
            df_cq_affiche.index = range(1, len(df_cq_affiche) + 1)
            df_cq_affiche["Montant (MAD)"] = df_cq_affiche["Montant (MAD)"].apply(lambda x: f"{int(x):,}".replace(",", " "))
            st.dataframe(df_cq_affiche, use_container_width=True)
        else:
            st.info("Aucun chèque futur à encaisser trouvé pour le moment.")

        st.markdown("---")
        st.subheader("🚨 Alertes & Relances de Paiement (Créances en cours)")
        
        df_creances = df_sales[df_sales["reliquat"].astype(float) > 0]
        if df_creances.empty:
            st.success("🎉 Aucun reliquat en attente ! Tous les clients sont à jour.")
        else:
            for _, row in df_creances.iterrows():
                lot = row["code_lot"]
                client = row["nom_client"]
                tel = row["telephone"]
                email = row["email"]
                rel = float(row["reliquat"])
                ech = row.get("date_ech_reliquat", "Non renseignée")

                with st.container():
                    col_a, col_b, col_c = st.columns([3, 2, 2])
                    with col_a:
                        st.markdown(f"**Appartement : {lot}** | Client : **{client}**")
                        st.caption(f"📞 Tél : {tel} | ✉️ E-mail : {email} | Échéance : {ech}")
                    with col_b:
                        st.markdown(f"**Reliquat : {int(rel):,} MAD**".replace(",", " "))
                    with col_c:
                        if tel and str(tel).lower() != "nan" and str(tel).strip() != "":
                            num_net = str(tel).replace(" ", "").replace("+", "").replace("-", "")
                            if num_net.startswith("0"):
                                num_net = "212" + num_net[1:]
                            wa_text = urllib.parse.quote(f"Bonjour {client}, suite à votre acquisition du lot {lot} à la Résidence Hillside, nous vous rappelons qu'un reliquat de {int(rel):,} MAD est en attente. Merci de régulariser votre situation.")
                            st.markdown(f'<a href="https://wa.me/{num_net}?text={wa_text}" target="_blank"><button style="background-color:#25D366; color:white; padding:6px 12px; border:none; border-radius:4px; font-weight:bold; width:100%;">🚀 Relancer WhatsApp</button></a>', unsafe_allow_html=True)
                    st.markdown("---")


elif menu == "⚙️ Paramètres & Analyses":
    st.title("⚙️ Paramètres & Analyses Sécurisés")
    st.caption("Espace réservé exclusivement à l'administrateur du CRM.")
    st.markdown("---")

    # --- VÉRIFICATION DE L'ACCÈS ADMINISTRATEUR ---
    if not st.session_state.admin_authentifie:
        st.warning("🔒 Cet espace est protégé. Veuillez entrer le code d'accès administrateur.")
        
        with st.form("form_code_admin"):
            code_saisi = st.text_input("Code d'accès administrateur :", type="password")
            btn_valider_code = st.form_submit_button("🔓 Valider l'accès")
            
            if btn_valider_code:
                if code_saisi.strip() == "Hillside2026" or st.session_state.user_email.lower() == ADMIN_PRINCIPAL.lower():
                    st.session_state.admin_authentifie = True
                    st.success("✅ Accès autorisé !")
                    st.rerun()
                else:
                    st.error("❌ Code d'accès incorrect.")
    else:
        if st.button("🔒 Verrouiller / Quitter l'espace sécurisé"):
            st.session_state.admin_authentifie = False
            st.rerun()

        tab_acc, tab_import, tab_api, tab_excel = st.tabs([
            "🔐 Gestion des Accès", 
            "📥 Importer votre tableau TDB", 
            "🔑 Clés API", 
            "🤖 Analyse Excel"
        ])

        with tab_acc:
            st.subheader("Liste des administrateurs autorisés")
            with sqlite3.connect("crm_hillside.db", timeout=10.0) as conn:
                df_emails = pd.read_sql_query("SELECT * FROM utilisateurs_autorises", conn)
            df_emails.index = range(1, len(df_emails) + 1)
            st.dataframe(df_emails, use_container_width=True)

            col1, col2 = st.columns(2)
            with col1:
                with st.form("form_add_mail_admin"):
                    nouveau_mail = st.text_input("Ajouter un e-mail autorisé :").strip().lower()
                    if st.form_submit_button("Autoriser"):
                        if nouveau_mail and "@" in nouveau_mail:
                            try:
                                with sqlite3.connect("crm_hillside.db", timeout=10.0) as conn:
                                    conn.execute("INSERT INTO utilisateurs_autorises (email) VALUES (?)", (nouveau_mail,))
                                    conn.commit()
                                st.success(f"✅ C'est OK ! E-mail '{nouveau_mail}' autorisé.")
                                st.rerun()
                            except sqlite3.IntegrityError:
                                st.warning("⚠️ Cet e-mail est déjà dans la liste.")
                        else:
                            st.error("E-mail invalide.")
            with col2:
                if not df_emails.empty:
                    mail_a_supprimer = st.selectbox("Révoquer un e-mail :", df_emails["email"].tolist())
                    if st.button("Supprimer cet accès"):
                        if mail_a_supprimer.lower() == ADMIN_PRINCIPAL.lower():
                            st.error("⛔ Impossible de supprimer l'administrateur principal.")
                        else:
                            with sqlite3.connect("crm_hillside.db", timeout=10.0) as conn:
                                conn.execute("DELETE FROM utilisateurs_autorises WHERE email = ?", (mail_a_supprimer,))
                                conn.commit()
                            st.success("🗑️ C'est OK ! Accès révoqué.")
                            st.rerun()

        with tab_import:
            st.subheader("📥 Importer ou remplacer le tableau de stocks")
            st.info("Le système détecte automatiquement le fichier TDB dans votre dossier, mais vous pouvez également en uploader un nouveau ci-dessous.")

            uploaded_tdb = st.file_uploader("Choisissez votre fichier TDB (.xls ou .xlsx)", type=["xls", "xlsx"], key="uploader_nouveau_tdb")
            if uploaded_tdb is not None:
                if st.button("🚀 Remplacer et Importer le nouveau tableau", use_container_width=True):
                    try:
                        uploaded_tdb.seek(0)
                        df_excel = pd.read_excel(uploaded_tdb, sheet_name="Pricing", skiprows=3)
                        
                        with sqlite3.connect("crm_hillside.db", timeout=10.0) as conn:
                            c = conn.cursor()
                            c.execute("DELETE FROM lots")
                            
                            nb_importes = 0
                            for _, row in df_excel.iterrows():
                                code_val = row.iloc[0] if len(row) > 0 else None
                                code = str(code_val).strip() if pd.notna(code_val) else ""
                                
                                exclus = ["nan", "total", "type", "prix moyen/m2", "ca previsionel", "ca pvp", "magasins"]
                                if not code or code.lower() in exclus or not code.startswith(("A", "B", "C", "D", "Magasin")):
                                    continue
                                
                                immeuble = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "Immeuble A"
                                typologie = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ("Magasin" if code.startswith("Magasin") else "F3")
                                etage = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else "RDC"
                                
                                etat_brut = str(row.iloc[4]).strip().upper() if pd.notna(row.iloc[4]) else "DISPONIBLE"
                                statut = "RESERVÉ" if "VENDU" in etat_brut or "RESERV" in etat_brut else "DISPONIBLE"
                                
                                vue = str(row.iloc[5]).strip() if pd.notna(row.iloc[5]) else ""
                                
                                try:
                                    surface_totale = float(row.iloc[6]) if pd.notna(row.iloc[6]) else 0.0
                                except:
                                    surface_totale = 0.0

                                try:
                                    surface_interieure = float(row.iloc[7]) if pd.notna(row.iloc[7]) else 0.0
                                except:
                                    surface_interieure = 0.0

                                try:
                                    terrasse = float(row.iloc[8]) if pd.notna(row.iloc[8]) else 0.0
                                except:
                                    terrasse = 0.0

                                try:
                                    jardin = float(row.iloc[9]) if pd.notna(row.iloc[9]) else 0.0
                                except:
                                    jardin = 0.0

                                try:
                                    parking = int(float(row.iloc[10])) if pd.notna(row.iloc[10]) else (0 if code.startswith("Magasin") else 1)
                                except:
                                    parking = 1

                                try:
                                    prix_plancher = float(row.iloc[11]) if pd.notna(row.iloc[11]) else 0.0
                                except:
                                    prix_plancher = 0.0

                                try:
                                    prix_catalogue = float(row.iloc[12]) if pd.notna(row.iloc[12]) else 0.0
                                except:
                                    prix_catalogue = 0.0

                                try:
                                    remise_max = float(row.iloc[15]) if pd.notna(row.iloc[15]) else 0.0
                                except:
                                    remise_max = 0.0

                                c.execute("""
                                    INSERT OR REPLACE INTO lots (
                                        code, immeuble, typologie, etage, statut, vue, surface_totale, 
                                        surface_interieure, terrasse, jardin, parking, prix_plancher, prix_catalogue, remise_max
                                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                                """, (code, immeuble, typologie, etage, statut, vue, surface_totale, surface_interieure, terrasse, jardin, parking, prix_plancher, prix_catalogue, remise_max))
                                nb_importes += 1
                            
                            conn.commit()

                        st.success(f"✅ C'est OK ! Ancien stock supprimé et {nb_importes} nouveaux lots importés avec succès.")
                        st.balloons()
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Erreur lors de l'importation : {e}")

        with tab_api:
            st.subheader("Configuration de la Clé API (Gemini / Assistant)")
            st.info("Enregistrez ou basculez facilement entre vos différentes clés API pour l'assistant intelligent.")
            
            with st.form("form_config_api"):
                nom_k = st.text_input("Nom de la clé (ex: Principale, Personnel, etc.) :")
                val_k = st.text_input("Valeur de la clé API :", type="password")
                submit_cle = st.form_submit_button("💾 Enregistrer cette clé")
                if submit_cle:
                    if nom_k.strip() and val_k.strip():
                        st.session_state.saved_api_keys[nom_k.strip()] = val_k.strip()
                        st.success(f"✅ C'est OK ! Clé '{nom_k.strip()}' enregistrée avec succès.")
                    else:
                        st.error("Veuillez remplir à la fois le nom et la valeur de la clé.")

            st.markdown("---")
            st.subheader("Sélection de la Clé Active")
            cles_dispos = list(st.session_state.saved_api_keys.keys())
            try:
                index_actuel = cles_dispos.index(st.session_state.active_api_key_name)
            except ValueError:
                index_actuel = 0

            choix_cle = st.selectbox("Clé API active pour les fonctionnalités d'IA :", cles_dispos, index=index_actuel)
            if choix_cle != st.session_state.active_api_key_name:
                st.session_state.active_api_key_name = choix_cle
                st.success(f"✅ C'est OK ! Clé active mise à jour : **{choix_cle}**")
                st.rerun()

        with tab_excel:
            st.subheader("🤖 Analyse Intelligente & Rendu Excel Identique")
            uploaded_excel = st.file_uploader("Choisissez votre fichier Excel :", type=["xlsx", "xls"], key="uploader_analyse_excel")
            if uploaded_excel:
                try:
                    wb = openpyxl.load_workbook(uploaded_excel, data_only=True)
                    feuilles = wb.sheetnames
                    
                    if len(feuilles) > 1:
                        feuille_choisie = st.selectbox("Sélectionner la feuille du classeur :", feuilles)
                        sheet = wb[feuille_choisie]
                    else:
                        sheet = wb.active
                        
                    st.success("✅ Fichier Excel chargé avec succès[cite: 1] !")
                    
                    merged_map = {}
                    for cr in sheet.merged_cells.ranges:
                        top_row, top_col = cr.min_row, cr.min_col
                        rowspan = cr.max_row - cr.min_row + 1
                        colspan = cr.max_col - cr.min_col + 1
                        merged_map[(top_row, top_col)] = {"rowspan": rowspan, "colspan": colspan}
                        for r in range(cr.min_row, cr.max_row + 1):
                            for c in range(cr.min_col, cr.max_col + 1):
                                if r == top_row and c == top_col:
                                    continue
                                merged_map[(r, c)] = "skip"

                    html_table = "<div style='overflow-x:auto; max-height: 600px; border: 1px solid #ccc;'><table style='border-collapse:collapse; width:100%; font-family:Calibri, sans-serif; font-size:13px;'>"
                    
                    for r_idx in range(1, sheet.max_row + 1):
                        html_table += "<tr>"
                        for c_idx in range(1, sheet.max_column + 1):
                            coord_status = merged_map.get((r_idx, c_idx), None)
                            if coord_status == "skip":
                                continue
                            
                            cell = sheet.cell(row=r_idx, column=c_idx)
                            val = cell.value if cell.value is not None else ""
                            if isinstance(val, str) and val.startswith("="):
                                val = ""
                            
                            cell_styles = "border: 1px solid #d3d3d3; padding: 6px 10px; white-space: nowrap;"
                            span_attrs = ""
                            if isinstance(coord_status, dict):
                                if coord_status["rowspan"] > 1:
                                    span_attrs += f" rowspan='{coord_status['rowspan']}'"
                                if coord_status["colspan"] > 1:
                                    span_attrs += f" colspan='{coord_status['colspan']}'"
                            
                            if cell.fill and cell.fill.fill_type and cell.fill.start_color:
                                rgb = cell.fill.start_color.rgb
                                if isinstance(rgb, str):
                                    if len(rgb) == 8:
                                        rgb = rgb[2:]
                                    if rgb and rgb not in ["000000", "FFFFFFFF", "None", "00000000"]:
                                        cell_styles += f" background-color: #{rgb};"
                            
                            if cell.font:
                                if cell.font.bold:
                                    cell_styles += " font-weight: bold;"
                                if cell.font.color and cell.font.color.rgb:
                                    f_rgb = cell.font.color.rgb
                                    if isinstance(f_rgb, str):
                                        if len(f_rgb) == 8:
                                            f_rgb = f_rgb[2:]
                                        if f_rgb and f_rgb not in ["000000", "00000000"]:
                                            cell_styles += f" color: #{f_rgb};"
                            
                            tag = "th" if r_idx == 1 else "td"
                            if tag == "th":
                                cell_styles += " background-color: #f1f3f4; font-weight: bold; position: sticky; top: 0; z-index: 2;"
                            
                            html_table += f"<{tag} style='{cell_styles}'{span_attrs}>{val}</{tag}>"
                        html_table += "</tr>"
                    html_table += "</table></div>"
                    
                    st.markdown("### 📊 Aperçu fidèle du fichier Excel (Mise en page complète)")
                    st.markdown(html_table, unsafe_allow_html=True)
                    
                    data_list = list(sheet.values)
                    if data_list:
                        raw_headers = data_list[0] if len(data_list) > 0 else []
                        en_tetes = [str(h).strip() if h is not None and str(h).strip() != "" else f"Col_{i+1}" for i, h in enumerate(raw_headers)]
                        df_u = pd.DataFrame(data_list[1:], columns=en_tetes).fillna("") if len(data_list) > 1 else pd.DataFrame()
                        
                        st.markdown("---")
                        st.subheader("💬 Poser une question à l'IA concernant ce tableau de bord")
                        
                        if "messages_excel_chat" not in st.session_state:
                            st.session_state.messages_excel_chat = [{
                                "role": "assistant",
                                "content": "Bonjour ! Je suis prêt à analyser votre fichier Excel. Que souhaitez-vous savoir ?",
                            }]

                        for msg in st.session_state.messages_excel_chat:
                            with st.chat_message(msg["role"]):
                                st.write(msg["content"])

                        user_excel_prompt = st.chat_input("Posez votre question sur les données du tableau...")
                        if user_excel_prompt:
                            if not active_api_key.strip():
                                st.error("⚠️ Clé API requise (configurable dans l'onglet 'Clés API').")
                            else:
                                st.session_state.messages_excel_chat.append({"role": "user", "content": user_excel_prompt})
                                with st.chat_message("user"):
                                    st.write(user_excel_prompt)
                                with st.chat_message("assistant"):
                                    with st.spinner("Analyse du tableau par l'IA..."):
                                        try:
                                            ai_client = genai.Client(api_key=active_api_key.strip())
                                            response = ai_client.models.generate_content(
                                                model="gemini-3.8-flash", contents=user_excel_prompt
                                            )
                                            reponse_ia = response.text
                                            st.write(reponse_ia)
                                            st.session_state.messages_excel_chat.append({"role": "assistant", "content": reponse_ia})
                                        except Exception as e:
                                            st.error(f"Erreur lors de l'analyse : {e}")
                                    
                except Exception as e:
                    st.error(f"Erreur lors de la lecture du fichier Excel : {e}")
