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
TITRE_OFFICIEL = "Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"
NOM_SOURCE = "dataEDT-ELT-S2-2026.xlsx"
FILE_EMAILS = "Permanents-Vacataires-ELT-2025-2026.xlsx"
TABLE_NAME = "surveillances_2026"

COLS_ORDRE = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']

S_URL = "https://ajcbkidmcjtyomknijwa.supabase.co"
S_KEY = "sb_publishable_otn3XM8LPLV0OGw74LRhDw_F446jkpw"

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465
MAIL_USER = "milouafarid@gmail.com"
MAIL_PASS = "kmtk zmkd kwpd cqzz" 

DATA_AUTO = {
    "ING1": {"Effectif": 61, "Horaire": "13h30 – 15h30", "Salles": ["S10", "S12", "S14", "S16"]},
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

HORAIRES_LIST = [
    "8h - 9h", "8h - 9h30", "8h - 10h", "9h - 10h", "9h30 - 11h", 
    "10h - 11h", "11h - 12h", "11h - 12h30", 
    "12h - 13h", "12h30 - 14h", "13h - 14h", "14h - 15h30", "14h - 16h", "15h30 - 17h"
]

def charger_donnees_locales(path):
    if os.path.exists(path):
        try:
            df = pd.read_excel(path)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            st.error(f"Erreur de lecture du fichier {path} : {e}")
            return pd.DataFrame()
    return pd.DataFrame()

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
    df_out = df.copy()
    # Renommer et reorganiser pour l'affichage
    if 'Enseignants' in df_out.columns:
        df_out['Surveillants'] = df_out['Enseignants']
    cols_wanted = ['Enseignements', 'Code', 'Responsable', 'Surveillants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
    for c in cols_wanted:
        if c not in df_out.columns:
            df_out[c] = ""
    df_out = df_out[cols_wanted]
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_out.to_excel(writer, index=False, sheet_name='Planning')
    return output.getvalue()

@st.cache_data
def charger_fichiers():
    df_s = pd.DataFrame()
    map_nom_complet = {}
    d_em = {}
    is_vacataire = {}

    FILE_CONTACTS = "Permanents-Vacataires-ELT2-2025-2026.xlsx"

    if os.path.exists(FILE_CONTACTS):
        try:
            df_c = pd.read_excel(FILE_CONTACTS)
            df_c.columns = [str(c).strip().upper() for c in df_c.columns]
            
            for _, row in df_c.iterrows():
                n = str(row.get('NOM', '')).strip().upper()
                p = str(row.get('PRÉNOM', '')).strip().upper()
                
                m_val = row.get('EMAIL') if 'EMAIL' in df_c.columns else row.get('Email')
                m = str(m_val).strip().lower() if pd.notna(m_val) else ""

                # Détection du statut (Permanent vs Vacataire)
                # Détection robuste de la colonne Qualité (avec ou sans accent, maj/min)
                cols_upper = [str(c).strip().upper() for c in df_c.columns]
                qualite_col = None
                for c in df_c.columns:
                    c_upper = str(c).strip().upper()
                    if c_upper in ['QUALITE', 'QUALITÉ', 'QUALITÉ', 'QUALITY']:
                        qualite_col = c
                        break

                if qualite_col:
                    cat_val = str(row.get(qualite_col, '')).strip().lower()
                else:
                    cat_val = str(row.get('CATEGORIE', row.get('STATUT', ''))).strip().lower()
                vac_flag = True if ('vacataire' in cat_val or 'vac' in cat_val or 'externe' in cat_val or 'assoc' in cat_val) else False
                
                if n and n != "NAN":
                    nom_complet = f"{n} {p}".strip()
                    map_nom_complet[n] = nom_complet
                    is_vacataire[nom_complet] = vac_flag
                    is_vacataire[n] = vac_flag
                    
                    if "@" in m:
                        d_em[n] = m
                        d_em[nom_complet] = m
        except Exception as e:
            st.error(f"Erreur lors de la lecture du fichier {FILE_CONTACTS} : {e}")
    else:
        st.error(f"Le fichier {FILE_CONTACTS} est introuvable à la racine du dépôt.")

    if os.path.exists("dataEDT-ELT-S2-2026.xlsx"):
        try:
            df_f = pd.read_excel("dataEDT-ELT-S2-2026.xlsx")
            df_f.columns = [str(c).strip() for c in df_f.columns]
            mask = df_f["Enseignements"].str.contains("Cours", case=False, na=False)
            df_s = df_f[mask].copy()
            for c in COLS_ORDRE:
                if c not in df_s.columns: df_s[c] = ""
            df_s = df_s[COLS_ORDRE]
        except Exception as e:
            st.error(f"Erreur source EDT : {e}")
            
    # Création des listes séparées Permanent / Vacataire depuis le fichier source
    liste_permanents = []
    liste_vacataires = []

    for nom_complet in map_nom_complet.values():
        if is_vacataire.get(nom_complet, False):
            liste_vacataires.append(nom_complet)
        else:
            liste_permanents.append(nom_complet)

    return df_s, map_nom_complet, d_em, is_vacataire, sorted(liste_permanents), sorted(liste_vacataires)

df_src, map_noms, dict_emails, dict_vacataires, LISTE_PERMANENTS, LISTE_VACATAIRES = charger_fichiers()

def get_nom_complet(nom_famille):
    """Retourne le nom complet (NOM PRENOM) a partir du nom de famille."""
    if not nom_famille or str(nom_famille).upper() == "NAN":
        return ""
    nom_upper = str(nom_famille).strip().upper()
    return map_noms.get(nom_upper, nom_upper)

# Liste complète pour compatibilité
LISTE_PROFS = sorted(set(LISTE_PERMANENTS + LISTE_VACATAIRES))

# Pour l'affichage dans les selectbox, on garde aussi les noms du fichier EDT si absents du fichier contacts
if not df_src.empty:
    noms_famille = df_src["Enseignants"].unique()
    for n in noms_famille:
        nom_upper = str(n).strip().upper()
        nom_complet = map_noms.get(nom_upper, nom_upper)
        if nom_complet not in LISTE_PROFS:
            LISTE_PROFS.append(nom_complet)
            # Par défaut, on le met dans les permanents si non classé
            if not dict_vacataires.get(nom_complet, False):
                LISTE_PERMANENTS.append(nom_complet)
            else:
                LISTE_VACATAIRES.append(nom_complet)
    LISTE_PROFS = sorted(set(LISTE_PROFS))
    LISTE_PERMANENTS = sorted(set(LISTE_PERMANENTS))
    LISTE_VACATAIRES = sorted(set(LISTE_VACATAIRES))

def creer_grille_edt(df):
    """Cree une grille Jours (lignes) x Horaires (colonnes) a partir du DataFrame."""
    if df.empty:
        return pd.DataFrame()
    df = df.copy()
    # Trier chronologiquement
    df['__sort'] = pd.to_datetime(df['Jours'], format='%d/%m/%Y', errors='coerce')
    df = df.sort_values('__sort')
    # Construire le detail de chaque case
    df['Details'] = df.apply(
        lambda r: f"📚 {r['Enseignements']}\n👨‍🏫 {r.get('Responsable','N/A')}\n🏫 {r['Lieu']}\n👤 {r['Enseignants']}", axis=1
    )
    grille = df.pivot_table(
        index='Jours', columns='Horaire', values='Details',
        aggfunc=lambda x: '\n---\n'.join(x)
    ).fillna('')
    return grille

def grille_to_html(grille_df, titre_sous=""):
    """Convertit la grille en HTML stylise pour telechargement."""
    if grille_df.empty:
        return f"<html><body><h2>{TITRE_OFFICIEL}</h2><p>Aucune donnee.</p></body></html>"
    grille_html = grille_df.copy()
    grille_html = grille_html.map(lambda x: x.replace('\n', '<br>') if x else '')
    html = f"""<html><head><meta charset='utf-8'><style>
        @page {{ size: landscape; }}
        body {{ font-family: Arial, sans-serif; margin: 20px; color: #333; }}
        h2 {{ text-align: center; color: #1f4e79; margin-bottom: 5px; }}
        h3 {{ text-align: center; color: #555; margin-top: 0; margin-bottom: 20px; }}
        table {{ border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 11px; }}
        th, td {{ border: 1px solid #444; padding: 8px; text-align: center; vertical-align: middle; }}
        th {{ background: #1f4e79; color: white; font-weight: bold; font-size: 12px; }}
        tr:nth-child(even) td {{ background: #f2f2f2; }}
        tr:nth-child(odd) td {{ background: #ffffff; }}
        td {{ line-height: 1.4; }}
        .footer {{ text-align: center; margin-top: 20px; font-size: 10px; color: #777; }}
    </style></head><body>
        <h2>{TITRE_OFFICIEL}</h2>
        <h3>{titre_sous}</h3>
        {grille_html.to_html(escape=False)}
        <div class='footer'>Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')}</div>
    </body></html>"""
    return html

def generer_html_impression(grille_df, titre):
    """Genere un HTML pret pour impression PDF (Ctrl+P -> Enregistrer en PDF)."""
    if grille_df.empty:
        return f"""<html><head><meta charset='utf-8'></head><body>
            <h2>{TITRE_OFFICIEL}</h2><p>Aucune donnee.</p></body></html>"""
    grille_html = grille_df.copy()
    grille_html = grille_html.map(lambda x: x.replace('\n', '<br>') if x else '')
    html = f"""<html><head><meta charset='utf-8'><style>
        @page {{ size: landscape; }}
        @media print {{
            .no-print {{ display: none !important; }}
            body {{ margin: 0; padding: 10px; }}
        }}
        body {{ font-family: Arial, sans-serif; margin: 20px; color: #333; }}
        h2 {{ text-align: center; color: #1f4e79; margin-bottom: 5px; }}
        h3 {{ text-align: center; color: #555; margin-top: 0; margin-bottom: 20px; }}
        table {{ border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 11px; }}
        th, td {{ border: 1px solid #444; padding: 8px; text-align: center; vertical-align: middle; }}
        th {{ background: #1f4e79; color: white; font-weight: bold; font-size: 12px; }}
        tr:nth-child(even) td {{ background: #f2f2f2; }}
        tr:nth-child(odd) td {{ background: #ffffff; }}
        td {{ line-height: 1.4; }}
        .footer {{ text-align: center; margin-top: 20px; font-size: 10px; color: #777; }}
        .print-btn {{ background: #1f4e79; color: white; padding: 12px 24px; border: none; 
                      border-radius: 6px; font-size: 14px; cursor: pointer; margin: 20px auto; display: block; }}
        .print-btn:hover {{ background: #163a5c; }}
        .notice {{ text-align: center; color: #666; font-size: 12px; margin-bottom: 10px; }}
    </style></head><body>
        <div class='no-print'>
            <button class='print-btn' onclick="window.print()">🖨️ Imprimer / Enregistrer en PDF</button>
            <p class='notice'>💡 Astuce : Cliquez sur le bouton ci-dessus, puis choisissez "Enregistrer au format PDF" dans les options d'impression.</p>
        </div>
        <h2>{TITRE_OFFICIEL}</h2>
        <h3>{titre}</h3>
        {grille_html.to_html(escape=False)}
        <div class='footer'>Genere le {datetime.now().strftime('%d/%m/%Y a %H:%M')}</div>
    </body></html>"""
    return html

def generer_excel_bytes_vertical(df, titre_sheet="Planning"):
    """Genere un fichier Excel VERTICAL (ligne par ligne) avec grille coloree."""
    output = io.BytesIO()

    # Colonnes a afficher et ordre
    cols_display = ['Jours', 'Horaire', 'Enseignements', 'Responsable', 'Lieu', 'Promotion', 'Surveillants']
    df_out = df.copy()
    # Copier Enseignants -> Surveillants
    if 'Enseignants' in df_out.columns:
        df_out['Surveillants'] = df_out['Enseignants']
    for c in cols_display:
        if c not in df_out.columns:
            df_out[c] = ""
    df_out = df_out[cols_display]

    # Trier par date puis horaire
    df_out = df_out.copy()
    df_out['__sort_date'] = pd.to_datetime(df_out['Jours'], format='%d/%m/%Y', errors='coerce')
    df_out = df_out.sort_values(['__sort_date', 'Horaire']).drop(columns=['__sort_date'])

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Ecrire les donnees a partir de la ligne 3 (sans en-tete, on le fait manuellement)
        df_out.to_excel(writer, sheet_name=titre_sheet, index=False, startrow=3, header=False)
        workbook = writer.book
        worksheet = writer.sheets[titre_sheet]

        # Formats
        fmt_header = workbook.add_format({
            'bold': True, 'bg_color': '#1f4e79', 'font_color': 'white',
            'align': 'center', 'valign': 'vcenter', 'border': 2,
            'font_size': 12, 'text_wrap': True
        })
        fmt_cell_white = workbook.add_format({
            'text_wrap': True, 'align': 'left', 'valign': 'vcenter',
            'border': 1, 'font_size': 10, 'bg_color': '#ffffff'
        })
        fmt_cell_alt = workbook.add_format({
            'text_wrap': True, 'align': 'left', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#e7f3ff', 'font_size': 10
        })
        fmt_cell_perm = workbook.add_format({
            'text_wrap': True, 'align': 'left', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#d4edda', 'font_size': 10, 'bold': True
        })
        fmt_cell_vac = workbook.add_format({
            'text_wrap': True, 'align': 'left', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#fff3cd', 'font_size': 10
        })
        fmt_title = workbook.add_format({
            'bold': True, 'font_size': 14, 'align': 'center',
            'valign': 'vcenter', 'bg_color': '#1f4e79', 'font_color': 'white'
        })
        fmt_subtitle = workbook.add_format({
            'bold': True, 'font_size': 12, 'align': 'center',
            'valign': 'vcenter', 'bg_color': '#d9e2f3', 'font_color': '#1f4e79'
        })

        # Largeurs des colonnes
        worksheet.set_column(0, 0, 15)   # Jours
        worksheet.set_column(1, 1, 15)   # Horaire
        worksheet.set_column(2, 2, 35)   # Enseignements
        worksheet.set_column(3, 3, 25)   # Responsable
        worksheet.set_column(4, 4, 20)   # Lieu
        worksheet.set_column(5, 5, 18)   # Promotion
        worksheet.set_column(6, 6, 40)   # Surveillants

        n_rows = len(df_out)
        n_cols = len(cols_display)

        # Titre ligne 0
        worksheet.merge_range(0, 0, 0, n_cols - 1, TITRE_OFFICIEL, fmt_title)
        worksheet.set_row(0, 30)

        # Sous-titre ligne 1
        worksheet.merge_range(1, 0, 1, n_cols - 1, titre_sheet, fmt_subtitle)
        worksheet.set_row(1, 25)

        # En-tete ligne 2
        for col in range(n_cols):
            worksheet.write(2, col, cols_display[col], fmt_header)
        worksheet.set_row(2, 25)

        # Donnees avec alternance de couleurs + couleur selon type enseignant
        for row_idx in range(n_rows):
            actual_row = row_idx + 3
            fmt_base = fmt_cell_alt if row_idx % 2 == 0 else fmt_cell_white
            worksheet.set_row(actual_row, 35)

            for col_idx in range(n_cols):
                val = df_out.iloc[row_idx, col_idx]
                cell_val = str(val) if pd.notna(val) else ""

                # Colonne Surveillants (index 6) : colorer selon permanent/vacataire
                if col_idx == 6 and cell_val:
                    noms_ens = [n.strip() for n in cell_val.split(" / ") if n.strip()]
                    has_perm = any(n in LISTE_PERMANENTS for n in noms_ens)
                    has_vac = any(n in LISTE_VACATAIRES for n in noms_ens)

                    if has_perm and not has_vac:
                        fmt_use = fmt_cell_perm
                    elif has_vac and not has_perm:
                        fmt_use = fmt_cell_vac
                    else:
                        fmt_use = fmt_base
                else:
                    fmt_use = fmt_base

                worksheet.write(actual_row, col_idx, cell_val, fmt_use)
    return output.getvalue()


def exploser_enseignants(df):
    """Transforme un DataFrame ou 'Enseignants' peut contenir 'Nom1 / Nom2'
    en un DataFrame ou chaque ligne n'a qu'un seul enseignant."""
    if df.empty:
        return df.copy()
    rows = []
    for _, row in df.iterrows():
        noms_bruts = str(row.get("Enseignants", "")).split(" / ")
        for nom in noms_bruts:
            nom = nom.strip()
            if not nom or "⚠️" in nom or "TEMP" in nom:
                continue
            new_row = row.copy()
            new_row["Enseignants"] = nom
            rows.append(new_row)
    return pd.DataFrame(rows).reset_index(drop=True)

def extraire_planning_individuel(df, nom_enseignant):
    """Retourne un DataFrame contenant UNIQUEMENT les seances de l'enseignant demande,
    avec son nom seul dans la colonne Enseignants (format ligne)."""
    df_explose = exploser_enseignants(df)
    nom_clean = str(nom_enseignant).strip()
    mask = df_explose["Enseignants"].str.strip().str.upper() == nom_clean.upper()
    return df_explose[mask].copy()

def get_liste_enseignants_individuels(df):
    """Retourne la liste triee de tous les enseignants individuels presents dans le planning."""
    df_explose = exploser_enseignants(df)
    noms = df_explose["Enseignants"].dropna().unique()
    return sorted([n for n in noms if n and "⚠️" not in n and "TEMP" not in n])

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
            <p><b>Matière :</b> {mat} | <b>Lieu:</b> {lieu} | <b>Date:</b> {jour}</p>
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
# 3. LOGIQUE D'AFFECTATION OPTIMISÉE (2 ENSEIGNANTS : 1 PERMANENT + 1 VACATAIRE)
# ======================================================================================
def affecter_enseignants_dynamique(batch_temp, df_global, q_psup, q_vac, q_def, p_sup, vacs, nb_total_requis=2):
    """
    Nouvelle logique d'affectation :
    - 1 ou 2 surveillants : UNIQUEMENT des permanents
    - 3 surveillants : 2 permanents + 1 vacataire
    - Le responsable de matiere est PRIORITAIRE au PREMIER lieu de sa matiere
      (s'il n'a pas de chevauchement et que son quota le permet).
    - Si la matiere est en plusieurs lieux, le responsable n'est affecte qu'au premier.
    Les permanents sont affiches en premier, suivis des vacataires.
    Evite les chevauchements (meme personne, meme jour, meme horaire).
    """
    df_full = pd.concat([df_global, pd.DataFrame(batch_temp)], ignore_index=True) if not df_global.empty else pd.DataFrame(batch_temp)

    set_vacataires = set(LISTE_VACATAIRES) | set(vacs)
    set_permanents = set(LISTE_PERMANENTS) - set(vacs)

    for i, row in enumerate(batch_temp):
        jour_seance = row['Jours']
        horaire_seance = row['Horaire']
        responsable_matiere = str(row.get('Responsable', '')).strip().upper()
        matiere_actuelle = str(row.get('Enseignements', '')).strip()
        lieu_actuel = str(row.get('Lieu', '')).strip()

        hor_seance = horaire_seance
        occupes_global = df_full[(df_full['Jours'] == jour_seance) & (df_full['Horaire'] == hor_seance)]['Enseignants'].tolist()
        occupes = set()
        for occ in occupes_global:
            for sub_occ in str(occ).split('/'):
                occupes.add(sub_occ.strip().upper())

        charges_actuelles = df_full[~df_full["Enseignants"].str.contains("⚠️|TEMP", na=False)]["Enseignants"].value_counts().to_dict()

        # === VERIFICATION : le responsable est-il DEJA affecte a cette matiere sur ce creneau ? ===
        resp_deja_affecte_ici = False
        if responsable_matiere and matiere_actuelle:
            df_resp_matiere = df_full[
                (df_full['Jours'] == jour_seance) & 
                (df_full['Horaire'] == hor_seance) & 
                (df_full['Enseignements'] == matiere_actuelle)
            ]
            for occ in df_resp_matiere['Enseignants'].tolist():
                for sub_occ in str(occ).split('/'):
                    sub = sub_occ.strip().upper()
                    if responsable_matiere in sub or sub in responsable_matiere:
                        resp_deja_affecte_ici = True
                        break
                if resp_deja_affecte_ici:
                    break

        candidats_permanents = []
        candidats_vacataires = []

        for p in LISTE_PROFS:
            p_upper = p.strip().upper()
            if p_upper in occupes or "⚠️" in p or "TEMP" in p:
                continue
            quota = q_psup if p in p_sup else (q_vac if p in set_vacataires else q_def)
            seances = charges_actuelles.get(p, 0)
            if seances >= quota:
                continue
            if p in set_vacataires:
                candidats_vacataires.append((p, seances))
            else:
                candidats_permanents.append((p, seances))

        assignes_cette_ligne = []

        # 1. PRIORITE AU RESPONSABLE DE MATIERE (uniquement au PREMIER lieu de la matiere)
        if responsable_matiere and not resp_deja_affecte_ici:
            for p in LISTE_PERMANENTS:
                if responsable_matiere in p.upper() or p.upper() in responsable_matiere:
                    p_upper = p.strip().upper()
                    if p_upper not in occupes:
                        quota = q_psup if p in p_sup else q_def
                        if charges_actuelles.get(p, 0) < quota:
                            assignes_cette_ligne.append(p)
                            occupes.add(p_upper)
                            break

        # 2. NOUVELLE LOGIQUE selon nb_total_requis
        if nb_total_requis <= 2:
            nb_perm_requis = nb_total_requis
            nb_vac_requis = 0
        elif nb_total_requis == 3:
            nb_perm_requis = 2
            nb_vac_requis = 1
        else:
            nb_perm_requis = nb_total_requis - 1
            nb_vac_requis = 1

        nb_perm_actuel = sum(1 for p in assignes_cette_ligne if p not in set_vacataires)
        nb_vac_actuel = sum(1 for p in assignes_cette_ligne if p in set_vacataires)

        # Completer avec des permanents
        while nb_perm_actuel < nb_perm_requis and candidats_permanents:
            candidats_permanents.sort(key=lambda x: x[1])
            min_c = candidats_permanents[0][1]
            meilleurs = [n for n, c in candidats_permanents if c == min_c]
            elu = random.choice(meilleurs)
            if elu not in assignes_cette_ligne:
                assignes_cette_ligne.append(elu)
                occupes.add(elu.strip().upper())
                nb_perm_actuel += 1
            candidats_permanents = [item for item in candidats_permanents if item[0] != elu]

        # Completer avec des vacataires si necessaire
        while nb_vac_actuel < nb_vac_requis and candidats_vacataires:
            candidats_vacataires.sort(key=lambda x: x[1])
            min_c = candidats_vacataires[0][1]
            meilleurs = [n for n, c in candidats_vacataires if c == min_c]
            elu = random.choice(meilleurs)
            if elu not in assignes_cette_ligne:
                assignes_cette_ligne.append(elu)
                occupes.add(elu.strip().upper())
                nb_vac_actuel += 1
            candidats_vacataires = [item for item in candidats_vacataires if item[0] != elu]

        # Si on n'a pas atteint le total, on complete avec ce qui reste
        tous_restants = candidats_permanents + candidats_vacataires
        tous_restants = [item for item in tous_restants if item[0] not in assignes_cette_ligne]

        while len(assignes_cette_ligne) < nb_total_requis and tous_restants:
            tous_restants.sort(key=lambda x: x[1])
            min_c = tous_restants[0][1]
            meilleurs = [n for n, c in tous_restants if c == min_c]
            elu = random.choice(meilleurs)
            assignes_cette_ligne.append(elu)
            occupes.add(elu.strip().upper())
            tous_restants = [item for item in tous_restants if item[0] != elu]

        if assignes_cette_ligne:
            # Ordonner : PERMANENTS d'abord, puis VACATAIRES
            perm_part = [p for p in assignes_cette_ligne if p not in set_vacataires]
            vac_part = [p for p in assignes_cette_ligne if p in set_vacataires]
            assignes_ordonnes = perm_part + vac_part

            nom_final_affiche = " / ".join(assignes_ordonnes)
            batch_temp[i]["Enseignants"] = nom_final_affiche
            batch_temp[i]["email"] = ", ".join([dict_emails.get(e, "") for e in assignes_ordonnes if dict_emails.get(e, "")])
            idx = len(df_global) + i if not df_global.empty else i
            df_full.at[idx, "Enseignants"] = nom_final_affiche
        else:
            batch_temp[i]["Enseignants"] = "⚠️ BESOIN (Déficit)"

    return batch_temp

# ======================================================================================
# 4. INTERFACE & SIDEBAR EN HAUT DE LA PAGE
# ======================================================================================

df_db_global = get_db()

# Initialisation des jours feries en session_state pour accessibilite globale
if 'feries' not in st.session_state:
    st.session_state.feries = []

# SIDEBAR EN HAUT DE LA PAGE (Disposition personnalisée demandée)
st.markdown("### 🎛️ Panneau de Contrôle & Configuration")
with st.expander("📂 Ouvrir la Sidebar / Administration", expanded=False):
    st.header("🔐 Administration")
    pwd_side = st.text_input("Code d'accès Configuration :", type="password", key="pwd_sidebar_cfg")

    if pwd_side == "0000":
        st.success("Accès autorisé")
        st.header("⚙️ Configuration")
        
        nb_amphi = st.number_input("👥 Surveillants / AMPHI", 1, 10, 2, key="nb_amp_cfg") # Par défaut 2 (1 perm + 1 vac)
        nb_salle = st.number_input("👥 Surveillants / SALLE", 1, 10, 2, key="nb_sal_cfg") # Par défaut 2 (1 perm + 1 vac)
        st.divider()
        
        st.subheader("⚖️ Quotas")
        list_p_cfg = LISTE_PROFS if 'LISTE_PROFS' in locals() or 'LISTE_PROFS' in globals() else []
        
        st.markdown("**👔 PERMANENTS détectés dans le fichier source :**")
        st.write(LISTE_PERMANENTS if LISTE_PERMANENTS else "Aucun permanent détecté")
        
        st.markdown("**📝 VACATAIRES détectés dans le fichier source :**")
        st.write(LISTE_VACATAIRES if LISTE_VACATAIRES else "Aucun vacataire détecté")
        
        st.divider()
        
        p_sup_list = st.multiselect("🎓 Postes Supérieurs (choisir parmi tous)", list_p_cfg, key="p_sup_cfg")
        q_psup = st.number_input("Seuil Max (Poste)", 0, 20, 2, key="q_psup_cfg")
        
        vac_list = st.multiselect("📝 Vacataires prioritaires (sélection manuelle parmi la liste)", LISTE_VACATAIRES, key="vac_list_cfg")
        q_vac = st.number_input("Seuil Max (Vac)", 0, 20, 6, key="q_vac_cfg")
        q_defaut = st.number_input("Seuil (Autres)", 0, 20, 3, key="q_def_cfg")
        st.divider()
        
        feries_sel = st.multiselect("🚫 Jours Fériés", [datetime(2026, 5, i).date() for i in range(1, 32)], default=st.session_state.feries, key="fer_cfg")
        st.session_state.feries = feries_sel
        st.divider()
        
        st.subheader("📥 Téléchargements")
        if not df_db_global.empty:
            st.download_button(
                "📊 Excel Global", 
                data=generer_excel_bytes(df_db_global[COLS_ORDRE]), 
                file_name="Planning_S2_2026.xlsx", 
                use_container_width=True,
                key="dl_xlsx_cfg"
            )
            
            df_html_g = df_db_global.copy()
            df_html_g['Surveillants'] = df_html_g['Enseignants']
            cols_html = ['Enseignements', 'Code', 'Responsable', 'Surveillants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
            for c in cols_html:
                if c not in df_html_g.columns:
                    df_html_g[c] = ""
            html_g = f"<html><body><h2 style='text-align:center;'>{TITRE_OFFICIEL}</h2>{df_html_g[cols_html].to_html(index=False)}</body></html>"
            st.download_button(
                "🌐 HTML Global", 
                data=html_g, 
                file_name="Planning_S2_2026.html", 
                mime="text/html", 
                use_container_width=True,
                key="dl_html_cfg"
            )

        st.divider()
        st.warning("⚠️ Zone de Maintenance Critique")
        confirm_delete = st.checkbox("Confirmer la suppression totale des données", key="chk_confirm_del")
        if st.button("🧨 VIDER LA BASE", use_container_width=True, key="btn_wipe_db_sec", disabled=not confirm_delete):
            try:
                supabase.table(TABLE_NAME).delete().neq("Promotion", "X").execute()
                st.success("✅ La base de données a été réinitialisée avec succès.")
                time.sleep(2)
                st.rerun()
            except Exception as e:
                st.error(f"❌ Erreur lors de la suppression : {e}")

    elif pwd_side != "":
        st.error("❌ Code d'accès incorrect.")
    else:
        st.info("🔒 Saisissez le code pour accéder aux paramètres avancés.")

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
    "📩 Justificatifs"
])

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
                            # 2 enseignants par salle/amphi par défaut
                            q = 2
                            for _ in range(1): # Une seule ligne par salle regroupant les 2 surveillants
                                batch.append({
                                    "Enseignements": m, "Code": str(inf["Code"]), "Enseignants": "TEMP", 
                                    "Horaire": h_c, "Jours": d_c.strftime("%d/%m/%Y"), "Lieu": l, 
                                    "Promotion": " / ".join(p_c), "Responsable": resp_c, "email": ""
                                })
                if batch:
                    # Determiner le nombre de surveillants selon le type de lieu
                    nb_surv_t1 = nb_salle
                    for lieu_test in s_c:
                        if str(lieu_test).startswith("A"):
                            nb_surv_t1 = max(nb_surv_t1, nb_amphi)
                            break
                    batch = affecter_enseignants_dynamique(batch, df_db_global, q_psup, q_vac, q_defaut, p_sup_list, vac_list, nb_total_requis=nb_surv_t1)
                    supabase.table(TABLE_NAME).insert(batch).execute()
                    st.success(f"✅ Session générée pour {resp_c}")
                    st.rerun()

with t2:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t2")
    else:
        st.subheader("📅 Planning Automatique")
        promos_existantes = df_db_global["Promotion"].unique() if not df_db_global.empty else []
        st.metric("🎓 Promotions déjà générées", len(promos_existantes))
        p_sel = st.selectbox("Sélectionner Promotion", [""] + list(DATA_AUTO.keys()), key="t2_promo")

        if p_sel:
            cfg = DATA_AUTO[p_sel]
            st.info(f"👥 Effectif : **{cfg['Effectif']}** étudiants")
            if p_sel in promos_existantes:
                st.warning(f"⚠️ La promotion {p_sel} sera remplacée.")

            mats_p = sorted(df_src[df_src["Promotion"]==p_sel]["Enseignements"].unique().tolist()) if not df_src.empty else []
            m_f_sel = st.multiselect("Modules à programmer", ["Toutes les matières"] + mats_p, default=["Toutes les matières"], key="t2_modules")

            if m_f_sel:
                mats_final = mats_p if "Toutes les matières" in m_f_sel else m_f_sel

                # === PLANNING ÉDITABLE PAR MATIÈRE ===
                st.markdown("### 📋 Ordre, dates et horaires des examens")
                st.caption("Modifiez l'ordre, la date et l'horaire de chaque matière ci-dessous.")

                edt_key = f"edt_plan_{p_sel}"

                if edt_key not in st.session_state or len(st.session_state[edt_key]) != len(mats_final):
                    dates_init = []
                    d_tmp = datetime(2026, 5, 11)
                    for _ in mats_final:
                        while d_tmp.weekday() in [4, 5] or d_tmp.date() in [d.date() for d in st.session_state.feries]:
                            d_tmp += timedelta(days=1)
                        dates_init.append(d_tmp.date())
                        d_tmp += timedelta(days=1)

                    st.session_state[edt_key] = pd.DataFrame({
                        "Matière": mats_final,
                        "Ordre": list(range(1, len(mats_final) + 1)),
                        "Date": dates_init,
                        "Horaire": [cfg["Horaire"]] * len(mats_final)
                    })

                df_edited = st.data_editor(
                    st.session_state[edt_key],
                    num_rows="fixed",
                    use_container_width=True,
                    key=f"edt_editor_{p_sel}",
                    column_config={
                        "Matière": st.column_config.TextColumn("Matière", disabled=True),
                        "Ordre": st.column_config.NumberColumn("Ordre", min_value=1, step=1),
                        "Date": st.column_config.DateColumn("Date examen"),
                        "Horaire": st.column_config.TextColumn("Horaire")
                    }
                )
                st.session_state[edt_key] = df_edited

                s_f = st.multiselect("Lieux", cfg["Salles"], default=cfg["Salles"], key=f"t2_lieux_{p_sel}")

                st.divider()
                if st.button("🚀 GÉNÉRER / REMPLACER", key="t2_generer"):
                    if p_sel in promos_existantes:
                        supabase.table(TABLE_NAME).delete().eq("Promotion", p_sel).execute()

                    df_planning = st.session_state[edt_key].sort_values("Ordre").reset_index(drop=True)

                    batch_pa = []
                    for _, row_plan in df_planning.iterrows():
                        mod = row_plan["Matière"]
                        d_exam = row_plan["Date"]
                        h_exam = row_plan["Horaire"]

                        inf_rows_pa = df_src[df_src["Enseignements"] == mod]
                        if not inf_rows_pa.empty:
                            inf = inf_rows_pa.iloc[0]
                            resp_nom = get_nom_complet(str(inf["Enseignants"]))
                            for lieu in s_f:
                                batch_pa.append({
                                    "Enseignements": mod, "Code": str(inf["Code"]), "Enseignants": "TEMP",
                                    "Horaire": h_exam, "Jours": d_exam.strftime("%d/%m/%Y"), "Lieu": lieu,
                                    "Promotion": p_sel, "Responsable": resp_nom, "email": ""
                                })

                    if batch_pa:
                        nb_surv_t2 = nb_salle
                        for lieu_test in s_f:
                            if str(lieu_test).startswith("A"):
                                nb_surv_t2 = max(nb_surv_t2, nb_amphi)
                                break
                        batch_pa = affecter_enseignants_dynamique(batch_pa, df_db_global, q_psup, q_vac, q_defaut, p_sup_list, vac_list, nb_total_requis=nb_surv_t2)
                        supabase.table(TABLE_NAME).insert(batch_pa).execute()
                        st.success(f"✅ Planning généré pour {p_sel} ({len(batch_pa)} séance(s))")
                        time.sleep(1)
                        st.rerun()

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

with t4:
    if not st.session_state.auth_admin_edt:
        afficher_verrou("t4")
    else:
        st.subheader("📋 Récapitulatif & Convocations")

        if df_db_global.empty:
            st.warning("Aucune donnée disponible dans la base.")
        else:
            # --- PLANNING GLOBAL ---
            st.markdown("### 🌍 Planning Global (Grille Jours x Horaires)")
            grille_global = creer_grille_edt(df_db_global)
            if not grille_global.empty:
                st.dataframe(grille_global, use_container_width=True)
                cgx, cgh, cgp = st.columns(3)
                with cgx:
                    st.download_button("📊 Excel Global", data=generer_excel_bytes_vertical(df_db_global, "Planning Global"), file_name="Planning_Global.xlsx", use_container_width=True, key="dl_xl_g")
                with cgh:
                    st.download_button("🌐 HTML Global", data=grille_to_html(grille_global, "Planning Global"), file_name="Planning_Global.html", mime="text/html", use_container_width=True, key="dl_ht_g")
                with cgp:
                    st.download_button("📄 PDF (Impression)", data=generer_html_impression(grille_global, "Planning Global - Surveillances S2 2026"), file_name="Planning_Global.html", mime="text/html", use_container_width=True, key="dl_pd_g")
            else:
                st.info("Données insuffisantes pour générer la grille globale.")

            st.divider()

            # --- PAR PROMOTION ---
            st.markdown("### 🎓 Plannings par Promotion")
            promos_uniques = set()
            for p in df_db_global['Promotion'].dropna().unique():
                for sub in str(p).split(' / '):
                    promos_uniques.add(sub.strip())
            promos_uniques = sorted(promos_uniques)

            for promo in promos_uniques:
                df_promo = df_db_global[df_db_global['Promotion'].str.contains(promo, na=False)]
                if df_promo.empty:
                    continue
                with st.expander(f"🎓 {promo} ({len(df_promo)} séance(s))"):
                    grille_p = creer_grille_edt(df_promo)
                    if not grille_p.empty:
                        st.dataframe(grille_p, use_container_width=True)
                        cp1, cp2, cp3 = st.columns(3)
                        with cp1:
                            st.download_button(f"📊 Excel {promo}", data=generer_excel_bytes_vertical(df_promo, f"Planning {promo}"), file_name=f"Planning_{promo}.xlsx", use_container_width=True, key=f"dl_xl_{promo}")
                        with cp2:
                            st.download_button(f"🌐 HTML {promo}", data=grille_to_html(grille_p, f"Planning {promo}"), file_name=f"Planning_{promo}.html", mime="text/html", use_container_width=True, key=f"dl_ht_{promo}")
                        with cp3:
                            st.download_button(f"📄 PDF (Impression) {promo}", data=generer_html_impression(grille_p, f"Planning {promo} - Surveillances S2 2026"), file_name=f"Planning_{promo}.html", mime="text/html", use_container_width=True, key=f"dl_pd_{promo}")
                    else:
                        st.info("Données insuffisantes pour cette promotion.")

            st.divider()

            # --- ENVOIS D'EMAILS ---
            st.markdown("### 📧 Convocations par Email")
            ce1, ce2 = st.columns(2)
            with ce1:
                st.markdown("#### 👤 Envoi Individuel")
                profs_actifs = get_liste_enseignants_individuels(df_db_global)
                ens_sel = st.selectbox("Sélectionner Enseignant", profs_actifs, key="sel_ens_t4")
                if st.button("📨 Lancer l'envoi unique", key="btn_envoi_unique"):
                    nom_a_chercher = str(ens_sel).strip().upper()
                    email_dest = dict_emails.get(nom_a_chercher) or dict_emails.get(nom_a_chercher.split(" ")[0])
                    if email_dest:
                        df_sessions = extraire_planning_individuel(df_db_global, ens_sel)
                        if not df_sessions.empty:
                            if envoyer_mail(ens_sel, email_dest, df_sessions):
                                st.success(f"✅ Convocation individuelle envoyée à {email_dest}")
                            else:
                                st.error("❌ Erreur SMTP.")
                        else:
                            st.warning(f"⚠️ Aucune séance trouvée pour {ens_sel}.")
                    else:
                        st.error(f"❌ Email introuvable pour {nom_a_chercher}")
            with ce2:
                st.markdown("#### 📢 Campagne Massive")
                if st.button("📧 LANCER LA CAMPAGNE GLOBALE", key="btn_campagne"):
                    profs_v = get_liste_enseignants_individuels(df_db_global)
                    bar = st.progress(0)
                    for i, p in enumerate(profs_v):
                        n_v = str(p).strip().upper()
                        m_v = dict_emails.get(n_v) or dict_emails.get(n_v.split(" ")[0])
                        if m_v:
                            df_indiv = extraire_planning_individuel(df_db_global, p)
                            if not df_indiv.empty:
                                envoyer_mail(p, m_v, df_indiv)
                        bar.progress((i + 1) / len(profs_v))
                    st.success("Campagne terminée !")

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

with t6:
    pwd_t6 = st.text_input("🔑 Code d'accès (T6) :", type="password", key="pwd_tab6")
    
    if pwd_t6 == "1234":
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

                            c_d1, c_d2, c_d3 = st.columns(3)
                            with c_d1:
                                date_abs = st.date_input("📅 Date de l'absence :", key="date_abs_t6")
                            with c_d2:
                                jours_semaine = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi"]
                                jour_abs = st.selectbox("🗓️ Jour :", jours_semaine, key="jour_abs_t6")
                            with c_d3:
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
    elif pwd_t6 != "":
        st.error("❌ Code incorrect.")

with t7:
    import base64
    st.header("📩 Système de Gestion des Justificatifs")
    choix_vue = st.radio("Sélectionnez votre profil :", ["Étudiant (Dépôt)", "Administration (Décision)"], horizontal=True)
    st.divider()

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
                    st.success(f"✅ Demande enregistrée avec succès !")
                    st.balloons()
                    time.sleep(1)
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur lors de l'envoi : {e}")
    else:
        pwd_admin = st.text_input("🔑 Code d'accès Administration :", type="password", key="pwd_admin_t7_sec")
        if pwd_admin == "1234":
            st.subheader("⚖️ Examen des dossiers en attente")
            try:
                query_admin = supabase.table("requetes_absences").select("*").eq("statut", "En attente").execute()
                resultats = query_admin.data
                if not resultats:
                    st.info("📭 Aucune demande en attente.")
                else:
                    for req in resultats:
                        with st.expander(f"📄 Dossier de {req['nom_etudiant']} ({req['matiere']})"):
                            st.write(f"**Motif :** {req['motif']}")
                            pdf_data_dec = base64.b64decode(req['justificatif_pdf'])
                            st.download_button("👁️ Visualiser le justificatif (PDF)", data=pdf_data_dec, file_name=f"Justif_{req['nom_etudiant']}.pdf", mime="application/pdf", key=f"view_{req['id']}")
                            c_fav, c_def = st.columns(2)
                            if c_fav.button("✅ ACCORDER", key=f"btn_ok_{req['id']}", use_container_width=True):
                                supabase.table("requetes_absences").update({"statut": "Favorable"}).eq("id", req['id']).execute()
                                supabase.table("suivi_assiduite_2026").delete().eq("etud_non_eligible", req['nom_etudiant']).eq("matiere", req['matiere']).execute()
                                st.success("Étudiant réhabilité.")
                                time.sleep(1)
                                st.rerun()
                            if c_def.button("❌ REJETER", key=f"btn_no_{req['id']}", use_container_width=True):
                                supabase.table("requetes_absences").update({"statut": "Défavorable"}).eq("id", req['id']).execute()
                                st.warning("Demande rejetée.")
                                time.sleep(1)
                                st.rerun()
            except Exception as e:
                st.error(f"Erreur : {e}")
        elif pwd_admin != "":
            st.error("❌ Code incorrect.")
