import streamlit as st
import pandas as pd
import os, random, io, smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime, timedelta
from supabase import create_client
import time
# ======================================================================================
# 1. CONFIGURATION & MÉMOIRE
# ======================================================================================
TITRE_OFFICIEL = "Plateforme de gestion des EDTs de surveillances du 11 au 21 Mai 2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
NOM_SOURCE = "dataEDT-ELT-S2-2026.xlsx"
FILE_EMAILS = "Permanents-Vacataires-ELT-2025-2026.xlsx"
TABLE_NAME = "surveillances_2026"

# Définition cruciale pour le filtrage des tableaux et exports
COLS_ORDRE = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']

S_URL = "https://ajcbkidmcjtyomknijwa.supabase.co"
S_KEY = "sb_publishable_otn3XM8LPLV0OGw74LRhDw_F446jkpw"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
MAIL_USER = "milouafarid@gmail.com"
MAIL_PASS = "kmtk zmkd kwpd cqzz" 

DATA_AUTO = {
    "ING1": {"Effectif": 118, "Horaire": "13h30 – 15h30", "Salles": ["S10", "S12", "S14", "S16"]},
    "ING2": {"Effectif": 16, "Horaire": "11h00 – 13h00", "Salles": ["S08 BIS (Promotion Entière)"]},
    "ING3EI": {"Effectif": 40, "Horaire": "08h30 – 10h30", "Salles": ["SN (Promotion Entière)"]},
    "ING3RSE": {"Effectif": 16, "Horaire": "08h30 – 10h30", "Salles": ["A10 (Promotion Entière)"]},
    "ING4": {"Effectif": 15, "Horaire": "11h00 – 13h00", "Salles": ["S14 (Promotion Entière)"]},
    "L2ELT": {"Effectif": 90, "Horaire": "11h00 – 13h00", "Salles": ["A08 (G1)", "A09 (G2)", "A12 (G3)"]},
    "L3ELT": {"Effectif": 70, "Horaire": "08h30 – 10h30", "Salles": ["A08 (G1)", "A09 (G2)", "A12 (G3)"]},
    "L1MCIL": {"Effectif": 288, "Horaire": "13h30 – 15h30", "Salles": ["S01", "S02", "S03", "S04", "S05", "S06", "SN", "S08"]},
    "L2MCIL": {"Effectif": 109, "Horaire": "11h00 – 13h00", "Salles": ["A10 (G1+G2)", "S02 (G3)", "S06 (G4)"]},
    "MCIL3": {"Effectif": 23, "Horaire": "08h30 – 10h30", "Salles": ["S01 (Promotion Entière)"]},
    "M1CE": {"Effectif": 12, "Horaire": "08h30 – 10h30", "Salles": ["S08 BIS (Promotion Entière)"]},
    "M1ER": {"Effectif": 15, "Horaire": "08h30 – 10h30", "Salles": ["S06 (Promotion Entière)"]},
    "M1ME": {"Effectif": 15, "Horaire": "08h30 – 10h30", "Salles": ["S08 (Promotion Entière)"]},
    "M1MCIL": {"Effectif": 34, "Horaire": "08h30 – 10h30", "Salles": ["S02 (Promotion Entière)"]},
    "M1RE": {"Effectif": 15, "Horaire": "08h30 – 10h30", "Salles": ["S04 (Promotion Entière)"]}
}

LISTE_SALLES = [f"S{i:02d}" for i in range(1, 19)] + ["SN", "S08 BIS"]
LISTE_AMPHIS = [f"A{i:02d}" for i in range(1, 13)]

st.set_page_config(page_title="Gestion EDT ELT 2026", layout="wide")
supabase = create_client(S_URL, S_KEY)
# --- FONCTION TECHNIQUE (À placer en haut du fichier) ---
# --- LISTE DES CRÉNEAUX HORAIRES (14 CRÉNEAUX) ---
HORAIRES_LIST = [
    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
    "10h - 11h", "11h - 12h", "11h - 12h30", 
    "12h - 13h", "12h30 - 14h", "13h - 14h", "14h - 15h30", "14h - 16h", "15h30 - 17h"
]
def charger_donnees_locales(path):
    """Charge un fichier Excel localement et nettoie les colonnes."""
    if os.path.exists(path):
        try:
            df = pd.read_excel(path)
            # Nettoyage des noms de colonnes (suppression des espaces invisibles)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            st.error(f"Erreur de lecture du fichier {path} : {e}")
            return pd.DataFrame()
    return pd.DataFrame()

# Variables de noms de fichiers pour T6
FILE_DATA_A = "DATA-ASSUIDUITE-2026.xlsx"
FILE_LISTE_A = "Liste des étudiants-2025-2026.xlsx"
# ======================================================================================
# 2. FONCTIONS TECHNIQUES
# ======================================================================================
def get_db():
    try:
        res = supabase.table(TABLE_NAME).select("*").execute()
        return pd.DataFrame(res.data) if res.data else pd.DataFrame(columns=COLS_ORDRE)
    except: return pd.DataFrame(columns=COLS_ORDRE)

def generer_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Planning')
    return output.getvalue()

@st.cache_data
def charger_fichiers():
    df_s = pd.DataFrame()
    map_nom_complet = {}
    d_em = {}

    # NOM DU FICHIER TEL QU'IL APPARAIT SUR VOTRE GITHUB
    FILE_CONTACTS = "Permanents-Vacataires-ELT2-2025-2026.xlsx"

    # 1. CHARGEMENT DES CONTACTS (Lecture locale)
    if os.path.exists(FILE_CONTACTS):
        try:
            df_c = pd.read_excel(FILE_CONTACTS)
            # Normalisation des noms de colonnes
            df_c.columns = [str(c).strip().upper() for c in df_c.columns]
            
            # On identifie les colonnes par rapport à votre JSON
            # NOM, PRÉNOM, Email
            for _, row in df_c.iterrows():
                n = str(row.get('NOM', '')).strip().upper()
                p = str(row.get('PRÉNOM', '')).strip().upper()
                
                # On cherche 'EMAIL' ou 'Email'
                m_val = row.get('EMAIL') if 'EMAIL' in df_c.columns else row.get('Email')
                m = str(m_val).strip().lower() if pd.notna(m_val) else ""
                
                if n and n != "NAN":
                    nom_complet = f"{n} {p}".strip()
                    map_nom_complet[n] = nom_complet
                    
                    if "@" in m:
                        # ON INDEXE PAR NOM SEUL ET NOM COMPLET
                        d_em[n] = m                # Clé: "MILOUA"
                        d_em[nom_complet] = m      # Clé: "MILOUA FARID"
        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier {FILE_CONTACTS} : {e}")
    else:
        st.error(f"Le fichier {FILE_CONTACTS} est introuvable à la racine du dépôt.")

    # 2. CHARGEMENT DE L'EDT (dataEDT-ELT-S2-2026.xlsx)
    if os.path.exists("dataEDT-ELT-S2-2026.xlsx"):
        try:
            df_f = pd.read_excel("dataEDT-ELT-S2-2026.xlsx")
            df_f.columns = [str(c).strip() for c in df_f.columns]
            mask = df_f["Enseignements"].str.contains("Cours", case=False, na=False)
            df_s = df_f[mask].copy()
            for c in ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']:
                if c not in df_s.columns: df_s[c] = ""
            df_s = df_s[['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']]
        except Exception as e:
            st.error(f"Erreur source EDT : {e}")
            
    return df_s, map_nom_complet, d_em
# --- APPEL ET INITIALISATION ---
df_src, map_noms, dict_emails = charger_fichiers()

if not df_src.empty:
    noms_famille = df_src["Enseignants"].unique()
    LISTE_PROFS = sorted([map_noms.get(str(n).strip().upper(), str(n).strip().upper()) for n in noms_famille])
else:
    LISTE_PROFS = []

def envoyer_mail(ens_nom, email_dest, df_perso):
    try:
        msg = MIMEMultipart()
        msg['From'] = MAIL_USER
        msg['To'] = email_dest
        msg['Subject'] = f"Convocation Surveillance S2 - {ens_nom}"
        
        rows = "".join([
            f"<tr>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Enseignements']}</td>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Horaire']}</td>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Jours']}</td>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Lieu']}</td>"
            f"<td style='border:1px solid #ddd;padding:8px;'>{r['Promotion']}</td>"
            f"</tr>" 
            for _, r in df_perso.iterrows()
        ])
        
        body = f"""
        <html>
        <body style='font-family: Arial, sans-serif; color: #333;'>
            <p>Salam M./Mme <b>{ens_nom}</b>,</p>
            <p>Veuillez recevoir votre planning de surveillance du semestre 02.</p>
            <div style='background-color: #fff3cd; padding: 15px; border-left: 5px solid #ffecb5; margin: 15px 0;'>
                <h4 style='margin-top:0;'>Rappels importants :</h4>
                <ul>
                    <li>Il est strictement interdit de rajouter un étudiant sur la liste sans aviser l'administration.</li>
                    <li>Le portable est strictement interdit et doit être en position éteinte.</li>
                    <li>Aucun étudiant ne peut accéder à la salle au-delà de 30 minutes après le début de l'épreuve.</li>
                </ul>
            </div>
            <table style='border-collapse: collapse; width: 100%; border: 1px solid #ddd;'>
                <thead style='background-color: #f2f2f2;'>
                    <tr><th>Matière</th><th>Heure</th><th>Jour</th><th>Lieu</th><th>Promo</th></tr>
                </thead>
                <tbody>{rows}</tbody>
            </table>
            <p><br>Cordialement,<br><b>Le chef de Département-ELT-FGE</b></p>
        </body>
        </html>
        """
        msg.attach(MIMEText(body, 'html'))
        xlsx = generer_excel_bytes(df_perso[COLS_ORDRE])
        part = MIMEApplication(xlsx, Name=f"Convocation_{ens_nom}.xlsx")
        part['Content-Disposition'] = f'attachment; filename="Convocation_{ens_nom}.xlsx"'
        msg.attach(part)
        
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as s:
            s.login(MAIL_USER, MAIL_PASS)
            s.sendmail(MAIL_USER, email_dest, msg.as_string())
        return True
    except Exception as e:
        print(f"Erreur envoi mail : {e}")
        return False

def generer_html_pv_pack(df_session):
    html = """<html><head><style>
        @media print {.pb {page-break-after:always;}} 
        body {font-family:Arial; font-size:12px; margin:10px;} 
        table {width:100%; border-collapse:collapse; margin-top:8px;} 
        td,th {border:1px solid black; padding:5px; text-align:left;}
        .header-pv {text-align:center; font-weight:bold; border-bottom:2px solid black; padding-bottom:5px; margin-bottom:10px;}
        .infraction-box {border: 2px solid #000; margin-top: 10px; height: 80px; padding: 5px; position: relative;}
        .infraction-box:before {content: "SIGNALEMENTS / INFRACTIONS :"; font-weight: bold; font-size: 10px;}
        .sep {border-top: 2px dashed #000; margin: 30px 0; padding-top: 10px; position: relative; text-align:center;}
        .sep:after {content: "✂ DÉCOUPE ICI ✂"; position: absolute; top: -12px; left: 42%; background: white; padding: 0 10px; font-size: 10px; font-weight:bold;}
    </style></head><body>"""
    
    groupes = df_session.groupby(['Enseignements', 'Lieu', 'Jours'])
    for idx, ((mat, lieu, jour), data) in enumerate(groupes):
        if idx > 0: html += "<div class='sep'></div>"
        promo = str(data['Promotion'].iloc[0])
        eff_prevu = 0
        for p_key in DATA_AUTO:
            if p_key in promo: eff_prevu += DATA_AUTO[p_key]['Effectif']
        
        html += f"""
        <div class='pb'>
            <div class='header-pv'>
                DÉPARTEMENT D'ÉLECTROTECHNIQUE<br>
                FACULTÉ DE GÉNIE ÉLECTRIQUE
            </div>
            <h3 style='text-align:center; margin:5px;'>PROCES-VERBAL DE SURVEILLANCE</h3>
            <p><b>:</b> {mat} | <b>Lieu:</b> {lieu} | <b>Date:</b> {jour}</p>
            <p><b>Promotion:</b> {promo} | <b>Chargé de matière:</b> {data.get('Responsable', pd.Series(['N/A'])).iloc[0]}</p>
            <table>
                <tr style='background-color:#f2f2f2;'>
                    <th>Étudiants prévus</th><th>Absences (Nombre)</th><th>Copies rendues (Nombre)</th>
                </tr>
                <tr style='height:30px; text-align:center;'>
                    <td>{eff_prevu if eff_prevu > 0 else '...'}</td><td></td><td></td>
                </tr>
            </table>
            <table>
                <tr style='background-color:#f2f2f2;'>
                    <th width='70%'>Surveillant (Nom & Prénom)</th><th>Signature</th>
                </tr>"""
        for s in data['Enseignants'].tolist():
            html += f"<tr><td>{s}</td><td style='height:35px;'></td></tr>"
        html += f"""</table>
            <div class='infraction-box'></div>
            <div style='display:flex; justify-content:space-between; margin-top:15px;'>
                <div style='text-align:center; width:45%;'>
                    <b>Signature du chargé de la matière</b><br><br><br>__________________________
                </div>
                <div style='text-align:center; width:45%;'>
                    <b>Signature Chef de salle / Amphi</b><br><br><br>__________________________
                </div>
            </div>
        </div>"""
    html += "</body></html>"
    return html

# ======================================================================================
# 3. LOGIQUE D'AFFECTATION
# ======================================================================================
def affecter_enseignants_dynamique(batch_temp, df_global, q_psup, q_vac, q_def, p_sup, vacs):
    df_full = pd.concat([df_global, pd.DataFrame(batch_temp)], ignore_index=True) if not df_global.empty else pd.DataFrame(batch_temp)
    for i, row in enumerate(batch_temp):
        charge = df_full[~df_full["Enseignants"].str.contains("⚠️|TEMP", na=False)]["Enseignants"].value_counts().to_dict()
        occupes = df_full[(df_full['Jours'] == row['Jours']) & (df_full['Horaire'] == row['Horaire'])]['Enseignants'].tolist()
        candidats = []
        for p in LISTE_PROFS:
            quota = q_psup if p in p_sup else (q_vac if p in vacs else q_def)
            seances = charge.get(p, 0)
            if seances < quota and p not in occupes:
                candidats.append((p, seances))
        if candidats:
            candidats.sort(key=lambda x: x[1])
            min_c = candidats[0][1]
            meilleurs = [n for n, c in candidats if c == min_c]
            elu = random.choice(meilleurs)
            batch_temp[i]["Enseignants"] = elu
            batch_temp[i]["email"] = dict_emails.get(elu, "")
            idx = len(df_global) + i if not df_global.empty else i
            df_full.at[idx, "Enseignants"] = elu
        else:
            batch_temp[i]["Enseignants"] = "⚠️ BESOIN (Déficit)"
    return batch_temp

# ======================================================================================
# 4. INTERFACE & CHARGEMENT DES DONNÉES (SÉCURISÉE)
# ======================================================================================

# Initialisation cruciale pour éviter le NameError dans tout le script
df_db_global = get_db()

with st.sidebar:
    # --- AJOUT SÉCURITÉ POUR LA CONFIGURATION ---
    st.header("🔐 Administration")
    pwd_side = st.text_input("Code d'accès Configuration :", type="password", key="pwd_sidebar_cfg")

    if pwd_side == "0000": # Remplacez 0000 par votre code secret
        st.success("Accès autorisé")
        st.header("⚙️ Configuration")
        
        # Paramètres de surveillance
        nb_amphi = st.number_input("👥 Surveillants / AMPHI", 1, 10, 3, key="nb_amp_cfg")
        nb_salle = st.number_input("👥 Surveillants / SALLE", 1, 10, 1, key="nb_sal_cfg")
        st.divider()
        
        # Gestion des Quotas
        st.subheader("⚖️ Quotas")
        # Récupération sécurisée de LISTE_PROFS
        list_p_cfg = LISTE_PROFS if 'LISTE_PROFS' in locals() or 'LISTE_PROFS' in globals() else []
        
        p_sup_list = st.multiselect("🎓 Postes Supérieurs", list_p_cfg, key="p_sup_cfg")
        q_psup = st.number_input("Seuil Max (Poste)", 0, 20, 2, key="q_psup_cfg")
        
        # Filtrage pour ne pas avoir de doublons dans les listes
        vac_list = st.multiselect("📝 Vacataires", [p for p in list_p_cfg if p not in p_sup_list], key="vac_list_cfg")
        q_vac = st.number_input("Seuil Max (Vac)", 0, 20, 6, key="q_vac_cfg")
        q_defaut = st.number_input("Seuil (Autres)", 0, 20, 3, key="q_def_cfg")
        st.divider()
        
        # Gestion des jours fériés
        feries = st.multiselect("🚫 Jours Fériés", [datetime(2026, 5, i).date() for i in range(1, 32)], key="fer_cfg")
        st.divider()
        
        # Section de téléchargement des données globales
        st.subheader("📥 Téléchargements")
        if not df_db_global.empty:
            # Export Excel
            st.download_button(
                "📊 Excel Global", 
                data=generer_excel_bytes(df_db_global[COLS_ORDRE]), 
                file_name="Planning_S2_2026.xlsx", 
                use_container_width=True,
                key="dl_xlsx_cfg"
            )
            
            # Export HTML avec rappel du TITRE_OFFICIEL
            html_g = f"<html><body><h2 style='text-align:center;'>{TITRE_OFFICIEL}</h2>{df_db_global[COLS_ORDRE].to_html(index=False)}</body></html>"
            st.download_button(
                "🌐 HTML Global", 
                data=html_g, 
                file_name="Planning_S2_2026.html", 
                mime="text/html", 
                use_container_width=True,
                key="dl_html_cfg"
            )

        # --- ZONE CRITIQUE : NETTOYAGE DE LA BASE DE DONNÉES ---
        st.divider()
        st.warning("⚠️ Zone de Maintenance Critique")
        
        # 1. Ajout de la case de confirmation (Sécurité supplémentaire)
        confirm_delete = st.checkbox("Confirmer la suppression totale des données", key="chk_confirm_del")
        
        # 2. Le bouton est désactivé (disabled) si la case n'est pas cochée
        if st.button("🧨 VIDER LA BASE", use_container_width=True, key="btn_wipe_db_sec", disabled=not confirm_delete):
            try:
                # Exécution de la suppression sur Supabase
                supabase.table(TABLE_NAME).delete().neq("Promotion", "X").execute()
                st.success("✅ La base de données a été réinitialisée avec succès.")
                
                # Petite pause pour confirmation visuelle
                import time
                time.sleep(2)
                
                # Rechargement de l'application
                st.rerun()
            except Exception as e:
                st.error(f"❌ Erreur lors de la suppression : {e}")

    # Gestion des messages liés au code d'accès de la barre latérale
    elif pwd_side != "":
        st.error("❌ Code d'accès incorrect.")
    else:
        st.info("🔒 Saisissez le code dans la barre latérale pour modifier la configuration.")

# --- RAPPEL DU TITRE OFFICIEL AU CENTRE ---
st.markdown(f"<h2 style='text-align:center; color:#1f4e79;'>{TITRE_OFFICIEL}</h2>", unsafe_allow_html=True)
# ======================================================================================
# 5. SYSTÈME DE SÉCURITÉ DYNAMIQUE
# ======================================================================================

CODE_ADMIN_EDT = "ELT2026"

if 'auth_admin_edt' not in st.session_state:
    st.session_state.auth_admin_edt = False

def verifier_admin_edt_dynamique(suffixe):
    cle_saisie = f"pwd_admin_edt_{suffixe}"
    if st.session_state[cle_saisie] == CODE_ADMIN_EDT:
        st.session_state.auth_admin_edt = True
        st.success("✅ Accès Administrateur accordé !")
        time.sleep(1)
        st.rerun()
    else:
        st.error("❌ Code incorrect.")

def afficher_verrou(suffixe):
    st.warning("🔒 Section réservée à l'administration du Département.")
    st.text_input(
        "Veuillez saisir le code secret pour modifier les EDTs :", 
        type="password", 
        key=f"pwd_admin_edt_{suffixe}", 
        on_change=verifier_admin_edt_dynamique,
        args=(suffixe,)
    )
    st.info("Note : Les onglets Assiduité (T6) et Requêtes (T7) restent accessibles aux étudiants.")

# ======================================================================================
# 6. CRÉATION DES ONGLETS ET LOGIQUE D'AFFICHAGE
# ======================================================================================

t1, t2, t3, t4, t5, t6, t7 = st.tabs([
    "🚀 SESSION COMMUNE", 
    "📅 PLANNING AUTO", 
    "📄 GÉNÉRER PV", 
    "📋 RÉCAPITULATIF", 
    "🔧 MAINTENANCE", 
    "📝 ASSIDUITÉ", 
    "📩 Soumettre un justificatif (Etudiants)"
])

# --- BLOC T1 : SESSION COMMUNE ---
with t1:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t1")
    else:
        st.subheader("📢 Session Commune")
        LISTE_COMPLÈTE_LIEUX = [
            "S01", "S02", "S03", "S04", "S05", "S06", "S07", "S08", "S08 BIS", "SN",
            "A01", "A02", "A03", "A04", "A05", "A06", "A07", "A08", "A09", "A10", "A11", "A12"
        ]
        c1, c2 = st.columns(2)
        with c1:
            resp_c = st.selectbox("Chargé de matière", LISTE_PROFS, key="sc_resp")
            nom_famille_seul = str(resp_c).split(" ")[0].strip().upper()
            mats_dispo = []
            if not df_src.empty:
                mask_mats = df_src["Enseignants"].str.upper().str.strip() == nom_famille_seul
                mats_dispo = sorted(df_src[mask_mats]["Enseignements"].unique().tolist())
            mat_sel_sc = st.selectbox("Matière", ["Toutes les matières"] + mats_dispo)
            p_c = st.multiselect("Promotions concernées", list(DATA_AUTO.keys()))

        h_defaut, s_defaut = "08h30 – 10h30", []
        if p_c:
            h_defaut = DATA_AUTO[p_c[0]]["Horaire"]
            s_auto = DATA_AUTO[p_c[0]]["Salles"]
            s_defaut = [s for s in s_auto if s in LISTE_COMPLÈTE_LIEUX]

        with c2:
            d_c = st.date_input("Date", datetime(2026, 5, 2), key="sc_d")
            creneaux_possibles = sorted(list(set([v["Horaire"] for v in DATA_AUTO.values()])))
            idx_h = creneaux_possibles.index(h_defaut) if h_defaut in creneaux_possibles else 0
            h_c = st.selectbox("Horaire", creneaux_possibles, index=idx_h, key="sc_h_select")
            s_c = st.multiselect("Lieux", LISTE_COMPLÈTE_LIEUX, default=s_defaut, key="sc_s")

        st.divider()
        if st.button("🔥 GÉNÉRER SESSION", use_container_width=True):
            if not p_c or not s_c:
                st.error("Veuillez sélectionner au moins une promotion et un lieu.")
            else:
                mats = mats_dispo if mat_sel_sc == "Toutes les matières" else [mat_sel_sc]
                batch = []
                for m in mats:
                    inf_rows = df_src[df_src["Enseignements"] == m]
                    if not inf_rows.empty:
                        inf = inf_rows.iloc[0]
                        for l in s_c:
                            q = nb_amphi if any(x in str(l) for x in ['A', 'SN']) else nb_salle
                            for _ in range(q):
                                batch.append({
                                    "Enseignements": m, "Code": str(inf["Code"]), "Enseignants": "TEMP", 
                                    "Horaire": h_c, "Jours": d_c.strftime("%d/%m/%Y"), "Lieu": l, 
                                    "Promotion": " / ".join(p_c), "Responsable": resp_c, "email": ""
                                })
                if batch:
                    batch = affecter_enseignants_dynamique(batch, df_db_global, q_psup, q_vac, q_defaut, p_sup_list, vac_list)
                    supabase.table(TABLE_NAME).insert(batch).execute()
                    st.success(f"✅ Session générée pour {resp_c}")
                    st.rerun()

# --- BLOC T2 : PLANNING AUTO ---
with t2:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t2")
    else:
        st.subheader("📅 Planning Automatique")
        promos_existantes = df_db_global["Promotion"].unique() if not df_db_global.empty else []
        st.metric("🎓 Promotions déjà générées", len(promos_existantes))
        p_sel = st.selectbox("Sélectionner Promotion", [""] + list(DATA_AUTO.keys()))
        
        if p_sel:
            cfg = DATA_AUTO[p_sel]
            st.info(f"👥 Effectif : **{cfg['Effectif']}** étudiants")
            if p_sel in promos_existantes:
                st.warning(f"⚠️ La promotion {p_sel} sera remplacée.")
            
            mats_p = sorted(df_src[df_src["Promotion"]==p_sel]["Enseignements"].unique().tolist()) if not df_src.empty else []
            ca, cb = st.columns(2)
            with ca:
                s_f = st.multiselect("Lieux", cfg["Salles"], default=cfg["Salles"])
                h_f = st.text_input("Heure", cfg["Horaire"])
            with cb: 
                m_f_sel = st.multiselect("Modules", ["Toutes les matières"] + mats_p, default=["Toutes les matières"])
                d_f = st.date_input("Date début", datetime(2026, 5, 11))
                
            if st.button("🚀 GÉNÉRER / REMPLACER"):
                if p_sel in promos_existantes:
                    supabase.table(TABLE_NAME).delete().eq("Promotion", p_sel).execute()
                
                mats_final = mats_p if "Toutes les matières" in m_f_sel else m_f_sel
                batch_pa, d_t = [], d_f
                for mod in mats_final:
                    while d_t.weekday() in [4, 5] or d_t in feries: 
                        d_t += timedelta(days=1)
                    inf_rows_pa = df_src[df_src["Enseignements"] == mod]
                    if not inf_rows_pa.empty:
                        inf = inf_rows_pa.iloc[0]
                        for lieu in s_f:
                            q = nb_amphi if any(x in lieu for x in ['A', 'SN']) else nb_salle
                            for _ in range(q):
                                batch_pa.append({
                                    "Enseignements": mod, "Code": str(inf["Code"]), "Enseignants": "TEMP", 
                                    "Horaire": h_f, "Jours": d_t.strftime("%d/%m/%Y"), "Lieu": lieu, 
                                    "Promotion": p_sel, "Responsable": inf["Enseignants"], "email": ""
                                })
                    d_t += timedelta(days=1)
                
                if batch_pa:
                    batch_pa = affecter_enseignants_dynamique(batch_pa, df_db_global, q_psup, q_vac, q_defaut, p_sup_list, vac_list)
                    supabase.table(TABLE_NAME).insert(batch_pa).execute()
                    st.rerun()

# --- BLOC T3 : GÉNÉRER PV ---
with t3:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t3")
    else:
        st.subheader("📄 Impression des PV")
        p_pv = st.selectbox("Promotion pour Pack PV", [""] + list(DATA_AUTO.keys()))
        if p_pv and not df_db_global.empty:
            df_p = df_db_global[df_db_global["Promotion"].str.contains(p_pv, na=False)]
            if not df_p.empty:
                pv_html = generer_html_pv_pack(df_p)
                st.download_button("🖨️ Télécharger Pack PV", data=pv_html, file_name=f"PV_{p_pv}.html", mime="text/html", use_container_width=True)
                st.components.v1.html(pv_html, height=600, scrolling=True)

# --- BLOC T4 : RÉCAPITULATIF ---
with t4:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t4")
    else:
        st.subheader("📋 Récapitulatif & Convocations")
        if not df_db_global.empty:
            st.dataframe(df_db_global[COLS_ORDRE], use_container_width=True)
            st.divider()
            ce1, ce2 = st.columns(2)
            with ce1:
                st.markdown("### 👤 Envoi Individuel")
                profs_actifs = sorted(df_db_global[~df_db_global["Enseignants"].str.contains("⚠️", na=False)]["Enseignants"].unique())
                ens_sel = st.selectbox("Sélectionner Enseignant", profs_actifs)
                if st.button("📨 Lancer l'envoi unique"):
                    nom_a_chercher = str(ens_sel).strip().upper()
                    email_dest = dict_emails.get(nom_a_chercher) or dict_emails.get(nom_a_chercher.split(" ")[0])
                    if email_dest:
                        df_sessions = df_db_global[df_db_global["Enseignants"] == ens_sel]
                        if envoyer_mail(ens_sel, email_dest, df_sessions):
                            st.success(f"✅ Envoyé à {email_dest}")
                        else: st.error("❌ Erreur SMTP.")
                    else: st.error(f"❌ Email introuvable pour {nom_a_chercher}")
            with ce2:
                st.markdown("### 📢 Campagne Massive")
                if st.button("📧 LANCER LA CAMPAGNE GLOBALE"):
                    profs_v = [p for p in df_db_global["Enseignants"].unique() if "⚠️" not in str(p)]
                    bar = st.progress(0)
                    for i, p in enumerate(profs_v):
                        n_v = str(p).strip().upper()
                        m_v = dict_emails.get(n_v) or dict_emails.get(n_v.split(" ")[0])
                        if m_v: envoyer_mail(p, m_v, df_db_global[df_db_global["Enseignants"] == p])
                        bar.progress((i + 1) / len(profs_v))
                    st.success("Campagne terminée !")
        else:
            st.warning("Aucune donnée disponible dans la base.")

# --- BLOC T5 : MAINTENANCE ---
with t5:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t5")
    else:
        st.subheader("🔧 Maintenance & Édition")
        if not df_db_global.empty:
            df_ed = st.data_editor(df_db_global, num_rows="dynamic", use_container_width=True)
            if st.button("💾 SAUVEGARDER LES MODIFICATIONS"):
                supabase.table(TABLE_NAME).delete().neq("Promotion", "X").execute()
                clean = df_ed.drop(columns=['id', 'created_at'], errors='ignore')
                supabase.table(TABLE_NAME).insert(clean.to_dict(orient='records')).execute()
                st.success("Modifications enregistrées !")
                st.rerun()

# --- T6 : Suivi de l’assiduité et liste des étudiants éligibles ---
with t6:
    # --- AJOUT DE LA SÉCURITÉ D'ACCÈS (Identique à T5) ---
    pwd_t6 = st.text_input("🔑 Code d'accès (T6) :", type="password", key="pwd_tab6")
    
    if pwd_t6 == "1234": # Remplacez par votre code réel
        st.markdown(f"### 📝 Suivi de l'Assiduité")
        
        df_aff_a = charger_donnees_locales(FILE_DATA_A)
        df_etud_m = charger_donnees_locales(FILE_LISTE_A)

        if df_aff_a.empty or df_etud_m.empty:
            st.error("⚠️ Fichiers sources (.xlsx) introuvables.")
        else:
            c1a, c2a = st.columns(2)
            with c1a:
                list_p_t6 = LISTE_PROFS if 'LISTE_PROFS' in locals() or 'LISTE_PROFS' in globals() else []
                sel_prof = st.selectbox("👤 Sélectionnez l'Enseignant :", [""] + list_p_t6, key="ens_T6")

            if sel_prof:
                nom_famille_a = str(sel_prof).split(" ")[0].strip().upper()
                with c2a:
                    mask_mats_a = df_aff_a["Enseignants"].str.upper().str.strip() == nom_famille_a
                    liste_mats = sorted(df_aff_a[mask_mats_a]["Enseignements"].unique().tolist())
                    sel_mat = st.selectbox("📚 Sélectionnez la Matière :", [""] + liste_mats, key="mat_T6")

                if sel_mat:
                    info_rows = df_aff_a[(df_aff_a["Enseignants"].str.upper().str.strip() == nom_famille_a) & (df_aff_a["Enseignements"] == sel_mat)]
                    if not info_rows.empty:
                        promo_c = str(info_rows.iloc[0]["Promotion"]).strip()
                        df_p = df_etud_m[df_etud_m["Promotion"].astype(str).str.strip() == promo_c].copy()
                        
                        if not df_p.empty:
                            df_p["Nom_Complet"] = df_p["Nom"].str.upper() + " " + df_p["Prénom"].str.title()
                            noms_e = sorted(df_p["Nom_Complet"].tolist())
                            
                            st.divider()
                            st.info(f"📍 Promotion détectée : **{promo_c}**")
                            
                            st.markdown("#### 🚫 Gestion de la Non-Éligibilité (Retrait)")
                            cn1, cn2 = st.columns(2)
                            with cn1:
                                etud_non = st.selectbox("👤 Étudiant concerné (Exclusion) :", [""] + noms_e, key="ne_et_t6")
                            with cn2:
                                causes = ["Hospitalisation", "Congé Académique", "Matière Acquise", "Exclu de matière", "Absence en TD", "Absence en TP"]
                                cause_s = st.selectbox("❓ Motif du retrait :", causes, key="ne_ca_t6")

                            # --- AJOUT DES CHAMPS DATE, JOUR ET HORAIRE ---
                            c_d1, c_d2, c_d3 = st.columns(3)
                            with c_d1:
                                date_abs = st.date_input("📅 Date de l'absence :", key="date_abs_t6")
                            with c_d2:
                                jours_semaine = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
                                jour_abs = st.selectbox("🗓️ Jour :", jours_semaine, key="jour_abs_t6")
                            with c_d3:
                                # Utilisation de la liste déroulante des 14 créneaux
                                horaire_abs = st.selectbox("🕒 Horaire :", options=HORAIRES_LIST, key="horaire_abs_t6")

                            if st.button("💾 ENREGISTRER DANS SUPABASE", use_container_width=True):
                                try:
                                    payload = {
                                        "enseignant": sel_prof,
                                        "matiere": sel_mat,
                                        "promotion": promo_c,
                                        "etud_non_eligible": etud_non if etud_non else "",
                                        "cause_non_eligibilite": cause_s if cause_s else "",
                                        "date_absence": str(date_abs),
                                        "jour_absence": jour_abs,
                                        "horaire_absence": horaire_abs,
                                        "date_saisie": datetime.now().strftime("%d/%m/%Y %H:%M")
                                    }
                                    
                                    supabase.table("suivi_assiduite_2026").insert(payload).execute()
                                    st.success(f"✅ Données enregistrées pour {sel_mat} !")
                                    time.sleep(1) 
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"❌ Erreur lors de l'enregistrement : {e}")

                            st.divider()
                            st.subheader("📥 Extraction des Rapports Officiels")

                            if sel_prof and sel_mat:
                                if 'promo_c' in locals() or 'promo_c' in globals():
                                    try:
                                        res_excl = supabase.table("suivi_assiduite_2026").select("etud_non_eligible").eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                        noms_exclus = [r['etud_non_eligible'] for r in res_excl.data if r.get('etud_non_eligible')]
                                    except Exception:
                                        noms_exclus = []

                                    if not df_p.empty and "Nom_Complet" in df_p.columns:
                                        import io
                                        output = io.BytesIO()

                                        df_eligible_final = df_p[~df_p["Nom_Complet"].isin(noms_exclus)].copy()
                                        export_eli = pd.DataFrame({
                                            "Nom et Prénom": df_eligible_final["Nom_Complet"],
                                            "Matière": sel_mat,
                                            "Chargé": sel_prof,
                                            "Promotion": promo_c
                                        })

                                        try:
                                            res_full = supabase.table("suivi_assiduite_2026").select("*").eq("matiere", sel_mat).eq("promotion", promo_c).execute()
                                            df_db_full = pd.DataFrame(res_full.data) if res_full.data else pd.DataFrame()
                                            
                                            if not df_db_full.empty and "etud_non_eligible" in df_db_full.columns:
                                                mask_non_eli = (df_db_full["etud_non_eligible"].notna()) & (df_db_full["etud_non_eligible"] != "")
                                                df_non_eligible = df_db_full[mask_non_eli].copy()
                                                
                                                # Mise à jour des colonnes exportées pour inclure date/jour/horaire
                                                cols_export = ["etud_non_eligible", "cause_non_eligibilite", "date_absence", "jour_absence", "horaire_absence", "matiere", "enseignant", "promotion"]
                                                export_non = df_non_eligible[cols_export].rename(
                                                    columns={
                                                        "etud_non_eligible": "Nom et Prénom", 
                                                        "cause_non_eligibilite": "Motif du Retrait",
                                                        "date_absence": "Date Absence",
                                                        "jour_absence": "Jour",
                                                        "horaire_absence": "Horaire",
                                                        "matiere": "Matière", 
                                                        "enseignant": "Chargé", 
                                                        "promotion": "Promotion"
                                                    }
                                                )
                                            else:
                                                export_non = pd.DataFrame()
                                        except Exception:
                                            export_non = pd.DataFrame()

                                        try:
                                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                                workbook = writer.book
                                                fmt_title = workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'})
                                                fmt_sub = workbook.add_format({'italic': True, 'font_size': 11, 'align': 'center'})
                                                fmt_bold = workbook.add_format({'bold': True})

                                                def appliquer_entete_officiel(sheet_obj, titre_liste):
                                                    sheet_obj.merge_range('A1:G1', "UNIVERSITÉ DJILLALI LIABÈS - SIDI BEL ABBÈS", fmt_title)
                                                    sheet_obj.merge_range('A2:G2', "Faculté de Génie Électrique - Département d'Électrotechnique", fmt_sub)
                                                    sheet_obj.merge_range('A3:G3', f"LISTE DES ÉTUDIANTS : {titre_liste}", fmt_title)
                                                    sheet_obj.write('A5', "Matière :", fmt_bold); sheet_obj.write('B5', sel_mat)
                                                    sheet_obj.write('A6', "Enseignant :", fmt_bold); sheet_obj.write('B6', sel_prof)
                                                    sheet_obj.write('D5', "Promotion :", fmt_bold); sheet_obj.write('E5', promo_c)
                                                    sheet_obj.write('D6', "Date export :", fmt_bold); sheet_obj.write('E6', datetime.now().strftime('%d/%m/%Y'))

                                                export_eli.to_excel(writer, sheet_name='Éligibles', startrow=8, index=False)
                                                appliquer_entete_officiel(writer.sheets['Éligibles'], "ÉLIGIBLES À L'EXAMEN")
                                                writer.sheets['Éligibles'].set_column('A:G', 22)

                                                if not export_non.empty:
                                                    export_non.to_excel(writer, sheet_name='Non-Éligibles', startrow=8, index=False)
                                                    appliquer_entete_officiel(writer.sheets['Non-Éligibles'], "NON-ÉLIGIBLES (RETRAIT)")
                                                    writer.sheets['Non-Éligibles'].set_column('A:G', 22)
                                                else:
                                                    ws2 = workbook.add_worksheet('Non-Éligibles')
                                                    appliquer_entete_officiel(ws2, "AUCUN ÉTUDIANT EXCLU")

                                            st.success(f"✅ Rapport généré ({len(export_eli)} étudiants éligibles).")
                                            st.download_button(label="📥 TÉLÉCHARGER LE RAPPORT COMPLET (XLSX)", data=output.getvalue(), file_name=f"Rapport_Officiel_{sel_mat}_{promo_c}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                                        except Exception as e:
                                            st.error(f"❌ Erreur Excel : {e}")
                                    else:
                                        st.warning("⚠️ Liste étudiants introuvable.")
                                else:
                                    st.error("⚠️ Promotion indéterminée.")
                            else:
                                st.info("ℹ️ Sélectionnez un enseignant et une matière.")
    elif pwd_t6 != "":
        st.error("❌ Code incorrect.")
    # --- T7 : GESTION DES REQUÊTES ET JUSTIFICATIFS PDF ---
with t7:
    import base64
    import time
    
    st.header("📩 Système de Gestion des Justificatifs")
    st.write("Cet espace permet aux étudiants de soumettre un justificatif et à l'administration de valider leur éligibilité.")

    choix_vue = st.radio("Sélectionnez votre profil :", ["Étudiant (Dépôt)", "Administration (Décision)"], horizontal=True)
    st.divider()

    # --- A. VUE ÉTUDIANT : FORMULAIRE DE DÉPÔT & SUIVI ---
    if choix_vue == "Étudiant (Dépôt)":
        st.subheader("📤 Soumettre une demande de réhabilitation")
        
        with st.form("form_depot_pdf_etudiant", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                l_etud = noms_e if 'noms_e' in locals() else ["Veuillez charger une promotion en T6"]
                etudiant_select = st.selectbox("Votre Nom et Prénom :", l_etud, key="sel_nom_etud_t7")
                
                l_mat = liste_mats if 'liste_mats' in locals() else ["Veuillez charger un enseignant en T6"]
                matiere_select = st.selectbox("Matière concernée :", l_mat, key="sel_mat_etud_t7")
            
            with col2:
                motif_abs = st.text_input("Motif de l'absence :", placeholder="ex: Certificat médical...", key="txt_motif_t7")
                fichier_pdf = st.file_uploader("Joindre le justificatif (PDF)", type=["pdf"], key="file_pdf_t7")

            submit_valider = st.form_submit_button("🚀 ENVOYER MA DEMANDE")

        if submit_valider:
            if not fichier_pdf:
                st.error("❌ Vous devez joindre un fichier PDF.")
            elif etudiant_select == "Veuillez charger une promotion en T6":
                st.warning("⚠️ Chargez d'abord une promotion en T6.")
            else:
                try:
                    pdf_bytes = fichier_pdf.read()
                    pdf_encoded = base64.b64encode(pdf_bytes).decode('utf-8')

                    data_req = {
                        "date_demande": datetime.now().strftime("%d/%m/%Y"),
                        "nom_etudiant": etudiant_select,
                        "matiere": matiere_select,
                        "promotion": promo_c if 'promo_c' in locals() else "N/A",
                        "motif": motif_abs,
                        "justificatif_pdf": pdf_encoded,
                        "statut": "En attente"
                    }

                    supabase.table("requetes_absences").insert(data_req).execute()
                    st.success(f"✅ Merci **{etudiant_select}**, votre demande a été enregistrée !")
                    st.balloons()
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur lors de l'envoi : {e}")

    # --- B. VUE ADMINISTRATION : TRAITEMENT DES DOSSIERS (SÉCURISÉE) ---
    else:
        # Ajout du champ pour le code d'accès
        pwd_admin = st.text_input("🔑 Code d'accès Administration :", type="password", key="pwd_admin_t7_sec")

        # Vérification du code pour déverrouiller la vue
        if pwd_admin == "1234":
            st.subheader("⚖️ Examen des dossiers en attente")
            try:
                # Récupération des dossiers "En attente" via Supabase
                query_admin = supabase.table("requetes_absences").select("*").eq("statut", "En attente").execute()
                resultats = query_admin.data

                if not resultats:
                    st.info("📭 Aucune demande en attente de validation.")
                else:
                    for req in resultats:
                        with st.expander(f"📄 Dossier de {req['nom_etudiant']} ({req['matiere']})"):
                            st.write(f"**Promotion :** {req['promotion']}")
                            st.write(f"**Motif :** {req['motif']}")
                            st.write(f"**Date de dépôt :** {req['date_demande']}")

                            # Décodage et bouton de visualisation du PDF
                            import base64
                            pdf_data_dec = base64.b64decode(req['justificatif_pdf'])
                            st.download_button(
                                label="👁️ Visualiser le justificatif (PDF)",
                                data=pdf_data_dec,
                                file_name=f"Justif_{req['nom_etudiant']}.pdf",
                                mime="application/pdf",
                                key=f"view_{req['id']}"
                            )

                            st.markdown("---")
                            c_fav, c_def = st.columns(2)

                            # ACTION : ACCORDER
                            if c_fav.button("✅ ACCORDER", key=f"btn_ok_{req['id']}", use_container_width=True):
                                supabase.table("requetes_absences").update({"statut": "Favorable"}).eq("id", req['id']).execute()
                                # Suppression automatique de l'exclusion dans la table d'assiduité
                                supabase.table("suivi_assiduite_2026").delete().eq("etud_non_eligible", req['nom_etudiant']).eq("matiere", req['matiere']).execute()
                                st.success(f"L'étudiant {req['nom_etudiant']} est réhabilité.")
                                import time
                                time.sleep(1)
                                st.rerun()

                            # ACTION : REJETER
                            if c_def.button("❌ REJETER", key=f"btn_no_{req['id']}", use_container_width=True):
                                supabase.table("requetes_absences").update({"statut": "Défavorable"}).eq("id", req['id']).execute()
                                st.warning("Demande rejetée.")
                                import time
                                time.sleep(1)
                                st.rerun()

            except Exception as e:
                st.error(f"Erreur lors de la récupération des données : {e}")
        
        # Message si le code n'est pas encore saisi ou incorrect
        elif pwd_admin != "":
            st.error("❌ Code d'accès incorrect.")

        # --- TABLEAU DE SUIVI (Visible pour tous ou selon votre choix) ---
        st.divider()
        st.subheader("📊 État d'avancement des demandes")
        try:
            p_f = promo_c if 'promo_c' in locals() else ""
            res_tab = supabase.table("requetes_absences").select("date_demande, nom_etudiant, matiere, motif, statut").eq("promotion", p_f).execute()
            
            if res_tab.data:
                df_tab = pd.DataFrame(res_tab.data)
                df_tab.columns = ["Date", "Étudiant", "Matière", "Motif", "Statut"]
                st.dataframe(df_tab, use_container_width=True, hide_index=True)
                
                import io
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
                    df_tab.to_excel(writer, index=False, sheet_name='Suivi')
                
                st.download_button(
                    label="📥 TÉLÉCHARGER LE TABLEAU (EXCEL)",
                    data=buf.getvalue(),
                    file_name=f"Suivi_Requetes_{p_f}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info(f"Aucune demande enregistrée pour la promotion : {p_f}")
        except Exception as e:
            st.error(f"Erreur d'affichage du tableau : {e}")
