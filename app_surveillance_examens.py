
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
import os
import re

# ============================================================
# CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Surveillances Examens S1 2026-2027",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .main-header { font-size: 2.2rem; font-weight: bold; color: #1f4e79; text-align: center;
        padding: 1rem; background: linear-gradient(90deg, #e3f2fd 0%, #bbdefb 100%);
        border-radius: 10px; margin-bottom: 1.5rem; }
    .sub-header { font-size: 1.4rem; font-weight: bold; color: #1565c0;
        margin-top: 1rem; margin-bottom: 0.5rem; border-bottom: 2px solid #1565c0; padding-bottom: 0.3rem; }
    .info-box { background-color: #e3f2fd; padding: 1rem; border-radius: 8px;
        border-left: 4px solid #1565c0; margin: 0.5rem 0; }
    .success-box { background-color: #e8f5e9; padding: 1rem; border-radius: 8px;
        border-left: 4px solid #2e7d32; margin: 0.5rem 0; }
    .warning-box { background-color: #fff3e0; padding: 1rem; border-radius: 8px;
        border-left: 4px solid #f57c00; margin: 0.5rem 0; }
    .card { background-color: #fafafa; padding: 1rem; border-radius: 8px;
        border: 1px solid #e0e0e0; margin: 0.5rem 0; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] { background-color: #f5f5f5; border-radius: 8px 8px 0 0;
        padding: 10px 20px; font-weight: 600; }
    .stTabs [aria-selected="true"] { background-color: #1565c0 !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

# Constantes
SALLES = [f"S{i:02d}" for i in range(1, 18)]   # S01 a S17
AMPHIS = [f"A{i:02d}" for i in range(1, 13)]   # A01 a A12
CRENEAUX = ["08h30 - 10h30", "11h00 - 13h00", "13h30 - 15h30"]
JOURS_FR = {"Monday": "Lundi", "Tuesday": "Mardi", "Wednesday": "Mercredi",
            "Thursday": "Jeudi", "Friday": "Vendredi", "Saturday": "Samedi", "Sunday": "Dimanche"}
FICHIER_SOURCE = "DATA-ENS-2026-2027_surveillances.xlsx"

# ============================================================
# SESSION STATE
# ============================================================
def init_session_state():
    defaults = {
        'enseignants_df': None, 'examens_df': None, 'planning_df': None,
        'surveillance_df': None, 'nb_surv_permanent': 3, 'nb_surv_vacataire': 2,
        'nb_surv_autre': 1, 'nb_surv_par_lieu': 2, 'exclus_manuels': [],
        'date_debut_val': date(2026, 11, 1), 'jours_feries': [],
        'promo_selected': None, 'data_loaded': False, 'promotions_list': [],
        'permanents_list': [], 'vacataires_list': [], 'all_enseignants_list': [],
        'ordre_matieres': {}, 'lieux_par_promo': {}
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

# ============================================================
# FONCTIONS
# ============================================================

def normaliser_qualite(val):
    val = str(val).strip().lower()
    mapping = {'permanent': 'Permanent', 'vacataire': 'Vacataire',
               'contractuel': 'Contractuel', 'autre': 'Autre',
               'professeur': 'Permanent', 'charge de cours': 'Vacataire',
               'charge_de_cours': 'Vacataire', 'doctorant': 'Vacataire',
               'maitre de conferences': 'Permanent', 'mc': 'Permanent', 'prof': 'Permanent'}
    return mapping.get(val, 'Permanent')

def est_cours(enseignement_str):
    """Verifie si l'enseignement est un Cours (exclut TD et TP)"""
    val = str(enseignement_str).strip()
    return bool(re.match(r'^[Cc][Oo][Uu][Rr][Ss]', val))

def extraire_nom_cours(enseignement_str):
    """Extrait le nom du cours apres le prefixe Cours-"""
    val = str(enseignement_str).strip()
    nom = re.sub(r'^[Cc][Oo][Uu][Rr][Ss][ \-_:]+', '', val).strip()
    return nom if nom else val

def charger_fichier_source_auto():
    """Charge le fichier et filtre uniquement les Cours"""
    try:
        paths_to_try = [FICHIER_SOURCE, os.path.join(os.getcwd(), FICHIER_SOURCE),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), FICHIER_SOURCE),
            os.path.join("/mnt/agents/upload/", FICHIER_SOURCE),
            os.path.join("/mount/src/planning-surveillances-s1-2026-2027/", FICHIER_SOURCE)]

        file_path = None
        for p in paths_to_try:
            if os.path.exists(p):
                file_path = p
                break

        if file_path is None:
            return None, f"Fichier {FICHIER_SOURCE} non trouve"

        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        ens_sheet = None

        for sheet in sheet_names:
            df_test = pd.read_excel(file_path, sheet_name=sheet, nrows=5)
            cols_lower = [str(c).lower().strip() for c in df_test.columns]
            has_qualite = any('qualite' in c or 'quality' in c or 'statut' in c or 'grade' in c for c in cols_lower)
            has_enseignements = any('enseignement' in c or 'cours' in c or 'matiere' in c or 'module' in c for c in cols_lower)
            has_nom = any('nom' in c or 'name' in c or 'enseignant' in c for c in cols_lower)
            if has_qualite and has_enseignements and has_nom:
                ens_sheet = sheet
                break

        if ens_sheet is None and len(sheet_names) > 0:
            ens_sheet = sheet_names[0]

        df_ens = pd.read_excel(file_path, sheet_name=ens_sheet)
        df_ens.columns = [str(col).strip() for col in df_ens.columns]
        cols_orig = list(df_ens.columns)
        cols_lower = [c.lower().strip().replace(' ', '_').replace('-', '_') for c in cols_orig]

        col_map = {}
        for i, c in enumerate(cols_lower):
            if any(x in c for x in ['nom', 'name', 'enseignant', 'prenom_nom', 'nom_prenom', 'professeur']):
                col_map['nom'] = cols_orig[i]
            elif any(x in c for x in ['qualite', 'quality', 'type', 'statut', 'grade', 'categorie', 'situation']):
                col_map['qualite'] = cols_orig[i]
            elif any(x in c for x in ['enseignement', 'cours', 'matiere', 'module', 'discipline', 'ue', 'matieres', 'enseignements']):
                col_map['enseignements'] = cols_orig[i]
            elif any(x in c for x in ['promotion', 'niveau', 'annee', 'class', 'promo', 'niveaux']):
                col_map['promotion'] = cols_orig[i]

        rename_map = {}
        if 'nom' in col_map: rename_map[col_map['nom']] = 'nom'
        if 'qualite' in col_map: rename_map[col_map['qualite']] = 'qualite'
        if 'enseignements' in col_map: rename_map[col_map['enseignements']] = 'enseignements'
        if 'promotion' in col_map: rename_map[col_map['promotion']] = 'promotion'
        df_ens = df_ens.rename(columns=rename_map)

        if 'nom' not in df_ens.columns:
            for col in df_ens.columns:
                if df_ens[col].dtype == 'object':
                    sample = df_ens[col].dropna().astype(str)
                    if len(sample) > 0 and sample.str.len().mean() > 3:
                        df_ens['nom'] = df_ens[col]
                        break

        for col in ['qualite', 'enseignements', 'promotion']:
            if col not in df_ens.columns:
                df_ens[col] = ''

        # Nettoyer
        df_ens = df_ens[df_ens['nom'].notna() & (df_ens['nom'].astype(str).str.strip() != '')].copy()
        df_ens['qualite'] = df_ens['qualite'].apply(normaliser_qualite)

        # === FILTRER UNIQUEMENT LES COURS ===
        examens_data = []
        for _, row in df_ens.iterrows():
            raw_ens = str(row.get('enseignements', ''))
            items = re.split(r'[,;/]+', raw_ens)
            for item in items:
                item = item.strip()
                if item and est_cours(item):
                    nom_cours = extraire_nom_cours(item)
                    examens_data.append({
                        'matiere': nom_cours,
                        'promotion': str(row.get('promotion', '')).strip(),
                        'enseignant': str(row.get('nom', '')).strip(),
                        'qualite_ens': row.get('qualite', 'Permanent'),
                        'date': None, 'creneau': None, 'lieu': None,
                        'ordre': 999
                    })

        df_exam = pd.DataFrame(examens_data)
        df_exam = df_exam.drop_duplicates(subset=['matiere', 'promotion', 'enseignant']).copy()
        df_exam = df_exam.sort_values('enseignant').drop_duplicates(subset=['matiere', 'promotion'], keep='first').copy()

        df_ens['nb_surveillance'] = 0
        df_ens['surveillance_attribuee'] = 0

        promotions = sorted(df_exam['promotion'].dropna().astype(str).str.strip().unique().tolist())
        promotions = [p for p in promotions if p != '']

        permanents = df_ens[df_ens['qualite'] == 'Permanent']['nom'].dropna().unique().tolist()
        vacataires = df_ens[df_ens['qualite'] == 'Vacataire']['nom'].dropna().unique().tolist()
        all_ens = df_ens['nom'].dropna().unique().tolist()

        return {
            'enseignants': df_ens, 'examens': df_exam, 'promotions': promotions,
            'permanents': permanents, 'vacataires': vacataires,
            'all_enseignants': all_ens, 'sheet_used': ens_sheet
        }, None

    except Exception as e:
        return None, str(e)

def est_jour_travaille(date_obj, jours_feries):
    """Vendredi et Samedi exclus. Dimanche est travaillable."""
    if isinstance(date_obj, datetime):
        date_obj = date_obj.date()
    elif isinstance(date_obj, str):
        try:
            date_obj = datetime.strptime(date_obj, "%Y-%m-%d").date()
        except:
            return True

    jour_semaine = date_obj.strftime("%A")
    jour_fr = JOURS_FR.get(jour_semaine, jour_semaine)

    if jour_fr in ["Vendredi", "Samedi"]:
        return False

    for jf in jours_feries:
        if isinstance(jf, str):
            try:
                jf = datetime.strptime(jf, "%d/%m/%Y").date()
            except:
                continue
        if isinstance(jf, datetime):
            jf = jf.date()
        if date_obj == jf:
            return False
    return True

def generer_planning_promo(examens_df, promotion, date_debut, jours_feries, creneaux, lieux, ordre_matieres=None):
    """Genere le planning pour une promotion avec ordre des matieres"""
    if examens_df is None or examens_df.empty:
        return None

    promo_df = examens_df[examens_df['promotion'].astype(str).str.strip() == str(promotion).strip()].copy()
    if promo_df.empty:
        return None

    if ordre_matieres and promotion in ordre_matieres:
        ordre_map = {m: i for i, m in enumerate(ordre_matieres[promotion])}
        promo_df['ordre'] = promo_df['matiere'].map(ordre_map).fillna(999).astype(int)
        promo_df = promo_df.sort_values('ordre')
    else:
        promo_df = promo_df.sort_values('matiere')

    date_courante = date_debut
    nb_creneaux = len(creneaux)
    nb_lieux = len(lieux)

    if nb_creneaux == 0 or nb_lieux == 0:
        st.error("Veuillez selectionner au moins un creneau et un lieu.")
        return None

    idx = 0
    for i in promo_df.index:
        while not est_jour_travaille(date_courante, jours_feries):
            date_courante += timedelta(days=1)

        creneau_idx = idx % nb_creneaux
        lieu_idx = idx % nb_lieux

        examens_df.at[i, 'date'] = date_courante
        examens_df.at[i, 'creneau'] = creneaux[creneau_idx]
        examens_df.at[i, 'lieu'] = lieux[lieu_idx]

        idx += 1
        if creneau_idx == nb_creneaux - 1:
            date_courante += timedelta(days=1)

    return examens_df

def attribuer_surveillants(planning_df, enseignants_df, nb_par_lieu=2):
    if planning_df is None or enseignants_df is None:
        return None, enseignants_df

    surveillants = enseignants_df.copy()
    surveillants['surveillance_attribuee'] = 0
    exclus = st.session_state.get('exclus_manuels', [])

    permanents = surveillants[(surveillants['qualite'] == 'Permanent') & (~surveillants['nom'].isin(exclus))].copy().sort_values('surveillance_attribuee')
    vacataires = surveillants[(surveillants['qualite'] == 'Vacataire') & (~surveillants['nom'].isin(exclus))].copy().sort_values('surveillance_attribuee')
    autres = surveillants[(~surveillants['qualite'].isin(['Permanent', 'Vacataire'])) & (~surveillants['nom'].isin(exclus))].copy().sort_values('surveillance_attribuee')

    attributions = []

    for idx, examen in planning_df.iterrows():
        date_examen = examen.get('date', None)
        creneau_examen = examen.get('creneau', '')
        matiere_examen = examen.get('matiere', '')
        enseignant_matiere = examen.get('enseignant', '')
        lieu_examen = examen.get('lieu', 'S01')

        if date_examen is None or pd.isna(date_examen):
            continue

        surveillants_occupes = set()
        for attr in attributions:
            attr_date = attr.get('date', None)
            if attr_date is not None and attr.get('creneau') == creneau_examen:
                d1 = attr_date.date() if hasattr(attr_date, 'date') else attr_date
                d2 = date_examen.date() if hasattr(date_examen, 'date') else date_examen
                if isinstance(d1, str):
                    try:
                        d1 = datetime.strptime(d1, "%Y-%m-%d").date()
                    except:
                        d1 = None
                if isinstance(d2, str):
                    try:
                        d2 = datetime.strptime(d2, "%Y-%m-%d").date()
                    except:
                        d2 = None
                if d1 and d2 and d1 == d2:
                    surveillants_occupes.update(attr.get('surveillants', []))

        liste_surveillants = []

        # 1. Enseignant charge de la matiere (prioritaire)
        if enseignant_matiere and str(enseignant_matiere) not in ['nan', '', 'None']:
            ens_info = surveillants[surveillants['nom'] == enseignant_matiere]
            if not ens_info.empty:
                nom_ens = ens_info.iloc[0]['nom']
                qualite_ens = ens_info.iloc[0]['qualite']
                if nom_ens not in surveillants_occupes and nom_ens not in exclus:
                    quota_key = f"nb_surv_{qualite_ens.lower()}"
                    quota = st.session_state.get(quota_key, 3)
                    current_count = surveillants.loc[surveillants['nom'] == nom_ens, 'surveillance_attribuee'].values[0]
                    if current_count < quota:
                        liste_surveillants.append({
                            'nom': nom_ens, 'qualite': qualite_ens, 'priorite': 'Charge de matiere'
                        })
                        surveillants_occupes.add(nom_ens)
                        surveillants.loc[surveillants['nom'] == nom_ens, 'surveillance_attribuee'] += 1

        # 2. Permanents (toujours en tete)
        for _, perm in permanents.iterrows():
            if len(liste_surveillants) >= nb_par_lieu:
                break
            if perm['nom'] in surveillants_occupes or perm['nom'] in exclus:
                continue
            quota_perm = st.session_state.get('nb_surv_permanent', 3)
            current_count = surveillants.loc[surveillants['nom'] == perm['nom'], 'surveillance_attribuee'].values[0]
            if current_count < quota_perm:
                liste_surveillants.append({'nom': perm['nom'], 'qualite': 'Permanent', 'priorite': 'Permanent'})
                surveillants_occupes.add(perm['nom'])
                surveillants.loc[surveillants['nom'] == perm['nom'], 'surveillance_attribuee'] += 1

        # 3. Vacataires
        for _, vac in vacataires.iterrows():
            if len(liste_surveillants) >= nb_par_lieu:
                break
            if vac['nom'] in surveillants_occupes or vac['nom'] in exclus:
                continue
            quota_vac = st.session_state.get('nb_surv_vacataire', 2)
            current_count = surveillants.loc[surveillants['nom'] == vac['nom'], 'surveillance_attribuee'].values[0]
            if current_count < quota_vac:
                liste_surveillants.append({'nom': vac['nom'], 'qualite': 'Vacataire', 'priorite': 'Vacataire'})
                surveillants_occupes.add(vac['nom'])
                surveillants.loc[surveillants['nom'] == vac['nom'], 'surveillance_attribuee'] += 1

        # 4. Autres
        for _, aut in autres.iterrows():
            if len(liste_surveillants) >= nb_par_lieu:
                break
            if aut['nom'] in surveillants_occupes or aut['nom'] in exclus:
                continue
            quota_aut = st.session_state.get('nb_surv_autre', 1)
            current_count = surveillants.loc[surveillants['nom'] == aut['nom'], 'surveillance_attribuee'].values[0]
            if current_count < quota_aut:
                liste_surveillants.append({'nom': aut['nom'], 'qualite': aut['qualite'], 'priorite': 'Autre'})
                surveillants_occupes.add(aut['nom'])
                surveillants.loc[surveillants['nom'] == aut['nom'], 'surveillance_attribuee'] += 1

        attributions.append({
            'date': date_examen, 'creneau': creneau_examen, 'matiere': matiere_examen,
            'promotion': examen.get('promotion', ''), 'lieu': lieu_examen,
            'surveillants': [s['nom'] for s in liste_surveillants],
            'details_surveillants': liste_surveillants
        })

    return attributions, surveillants

def generer_tableau_html(attributions):
    if not attributions:
        return "<p>Aucune attribution</p>"

    planning_par_jour = {}
    for attr in attributions:
        date_val = attr.get('date', None)
        if date_val is None:
            continue
        if isinstance(date_val, str):
            try:
                date_val = datetime.strptime(date_val, "%Y-%m-%d").date()
            except:
                continue
        elif isinstance(date_val, datetime):
            date_val = date_val.date()

        date_str = date_val.strftime('%d/%m/%Y')
        jour_fr = JOURS_FR.get(date_val.strftime('%A'), date_val.strftime('%A'))
        cle = f"{jour_fr} {date_str}"
        if cle not in planning_par_jour:
            planning_par_jour[cle] = {}
        creneau = attr.get('creneau', '')
        if creneau not in planning_par_jour[cle]:
            planning_par_jour[cle][creneau] = []
        planning_par_jour[cle][creneau].append(attr)

    html = """
    <style>
        .planning-table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 11px; }
        .planning-table th { background-color: #1565C0; color: white; padding: 8px; text-align: center; border: 2px solid #0D47A1; }
        .planning-table td { padding: 6px; border: 1px solid #90CAF9; vertical-align: top; }
        .creneau-cell { background-color: #E3F2FD; font-weight: bold; text-align: center; width: 120px; }
        .examen-cell { background-color: #FFF8E1; margin: 2px 0; padding: 5px; border-radius: 3px; border-left: 3px solid #FFA000; font-size: 10px; }
        .surv-permanent { color: #1565C0; font-weight: bold; }
        .surv-vacataire { color: #2E7D32; }
        .surv-autre { color: #E65100; }
        .surv-charge { color: #6A1B9A; font-weight: bold; }
    </style>
    <table class="planning-table">
    """
    jours = sorted(planning_par_jour.keys())
    html += "<tr><th>Creneau</th>"
    for jour in jours:
        html += f"<th>{jour}</th>"
    html += "</tr>"

    for creneau in CRENEAUX:
        html += f"<tr><td class='creneau-cell'>{creneau}</td>"
        for jour in jours:
            html += "<td>"
            if creneau in planning_par_jour.get(jour, {}):
                for examen in planning_par_jour[jour][creneau]:
                    survs = examen.get('details_surveillants', [])
                    surv_html = "<br>".join([
                        f"<span class='surv-{s['qualite'].lower() if s.get('priorite') != 'Charge de matiere' else 'charge'}'>{s['nom']} ({s['qualite']}{'*' if s.get('priorite') == 'Charge de matiere' else ''})</span>"
                        for s in survs
                    ])
                    html += f"""
                    <div class='examen-cell'>
                        <strong>{examen.get('matiere', '')}</strong><br>
                        <small>Promo: {examen.get('promotion', '')} | {examen.get('lieu', '')}</small><br>
                        <small>{surv_html}</small>
                    </div>
                    """
            html += "</td>"
        html += "</tr>"
    html += "</table>"
    return html

def generer_excel_colore(attributions):
    wb = Workbook()
    ws = wb.active
    ws.title = "Planning"

    data = []
    for attr in attributions:
        date_val = attr.get('date', None)
        if hasattr(date_val, 'strftime'):
            date_str = date_val.strftime('%d/%m/%Y')
            jour = JOURS_FR.get(date_val.strftime('%A'), date_val.strftime('%A'))
        else:
            date_str = str(date_val)
            jour = ''
        surv_str = ", ".join([f"{s['nom']} ({s['qualite']})" for s in attr.get('details_surveillants', [])])
        data.append({'Date': date_str, 'Jour': jour, 'Creneau': attr.get('creneau', ''),
                     'Matiere': attr.get('matiere', ''), 'Promotion': attr.get('promotion', ''),
                     'Lieu': attr.get('lieu', ''), 'Surveillants': surv_str})

    df = pd.DataFrame(data)
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                cell.border = Border(left=Side(style='thin', color='90CAF9'), right=Side(style='thin', color='90CAF9'),
                                     top=Side(style='thin', color='90CAF9'), bottom=Side(style='thin', color='90CAF9'))

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[column].width = min(max_length + 2, 60)

    ws2 = wb.create_sheet("Resume Enseignants")
    ens_count = {}
    for attr in attributions:
        for surv in attr.get('details_surveillants', []):
            nom = surv['nom']
            if nom not in ens_count:
                ens_count[nom] = {'count': 0, 'qualite': surv['qualite'], 'examens': []}
            ens_count[nom]['count'] += 1
            date_val = attr.get('date', '')
            date_str = date_val.strftime('%d/%m') if hasattr(date_val, 'strftime') else str(date_val)
            ens_count[nom]['examens'].append(f"{attr.get('matiere', '')} ({date_str})")

    resume_data = []
    for nom, info in sorted(ens_count.items(), key=lambda x: x[1]['count'], reverse=True):
        resume_data.append({'Enseignant': nom, 'Qualite': info['qualite'],
                           'Nombre': info['count'], 'Examens': "; ".join(info['examens'][:5]) + ("..." if len(info['examens']) > 5 else "")})

    df_resume = pd.DataFrame(resume_data)
    for r_idx, row in enumerate(dataframe_to_rows(df_resume, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws2.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font

    for col in ws2.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws2.column_dimensions[column].width = min(max_length + 2, 60)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generer_pdf(attributions):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                           rightMargin=1*cm, leftMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'],
                                 fontSize=16, textColor=colors.HexColor('#1565C0'), spaceAfter=15, alignment=1)

    elements.append(Paragraph("PLANNING DES SURVEILLANCES", title_style))
    elements.append(Paragraph("Annee 2026-2027 - Semestre 1", styles['Heading2']))
    elements.append(Spacer(1, 0.3*cm))

    table_data = [['Date', 'Jour', 'Creneau', 'Matiere', 'Promotion', 'Lieu', 'Surveillants']]
    for attr in attributions:
        date_val = attr.get('date', None)
        if hasattr(date_val, 'strftime'):
            date_str = date_val.strftime('%d/%m/%Y')
            jour = JOURS_FR.get(date_val.strftime('%A'), date_val.strftime('%A'))
        else:
            date_str = str(date_val)
            jour = ''
        surv_str = ", ".join([f"{s['nom']} ({s['qualite']})" for s in attr.get('details_surveillants', [])])
        table_data.append([date_str, jour, attr.get('creneau', ''), attr.get('matiere', ''),
                          attr.get('promotion', ''), attr.get('lieu', ''), surv_str])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#90CAF9')),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#E3F2FD'), colors.HexColor('#FFFFFF')]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ============================================================
# MAIN
# ============================================================

def main():
    init_session_state()

    st.markdown('<div class="main-header">📋 Surveillances Examens S1 2026-2027</div>', unsafe_allow_html=True)

    # Chargement auto
    if not st.session_state.data_loaded:
        with st.spinner("Chargement du fichier source..."):
            result, error = charger_fichier_source_auto()
            if result is not None:
                st.session_state.enseignants_df = result['enseignants']
                st.session_state.examens_df = result['examens']
                st.session_state.promotions_list = result['promotions']
                st.session_state.permanents_list = result['permanents']
                st.session_state.vacataires_list = result['vacataires']
                st.session_state.all_enseignants_list = result['all_enseignants']
                st.session_state.data_loaded = True
                if result['promotions']:
                    st.session_state.promo_selected = result['promotions'][0]
                st.success(f"✅ Charge: {result['sheet_used']} | {len(result['enseignants'])} ens. | {len(result['examens'])} cours | Promos: {', '.join(result['promotions'])}")
            else:
                st.error(f"❌ {error}")
                st.info("Placez le fichier dans le meme dossier que l'application.")

    # Sidebar
    with st.sidebar:
        st.markdown("## ⚙️ Configuration")
        if st.session_state.data_loaded:
            st.markdown("### 📁 Source")
            st.success(f"{FICHIER_SOURCE} charge")
            st.markdown(f"- Enseignants: {len(st.session_state.enseignants_df)}")
            st.markdown(f"- **Cours uniquement**: {len(st.session_state.examens_df)}")
            st.markdown(f"- Promotions: {', '.join(st.session_state.promotions_list)}")
            if st.button("🔄 Recharger", key="btn_reload"):
                st.session_state.data_loaded = False
                st.rerun()
        else:
            st.warning("Fichier non charge")
            fu = st.file_uploader("Charger manuellement", type=['xlsx', 'xls'], key="manual_up")
            if fu is not None:
                with open(FICHIER_SOURCE, "wb") as f:
                    f.write(fu.getvalue())
                st.session_state.data_loaded = False
                st.rerun()

        st.markdown("---")
        st.markdown("### 📊 Quotas")
        st.session_state.nb_surv_permanent = st.number_input("Permanent", 0, 20, st.session_state.nb_surv_permanent, key="w_qp")
        st.session_state.nb_surv_vacataire = st.number_input("Vacataire", 0, 20, st.session_state.nb_surv_vacataire, key="w_qv")
        st.session_state.nb_surv_autre = st.number_input("Autre", 0, 20, st.session_state.nb_surv_autre, key="w_qa")
        st.session_state.nb_surv_par_lieu = st.number_input("Surv. par lieu", 1, 5, st.session_state.nb_surv_par_lieu, key="w_nl")

        st.markdown("---")
        st.markdown("### 📅 Date de Debut")
        st.session_state.date_debut_val = st.date_input("Date debut", st.session_state.date_debut_val, key="w_dd")

        st.markdown("### 🎉 Jours Feries")
        jf = st.text_area("JJ/MM/AAAA, separes par virgules", placeholder="11/11/2026, 25/12/2026...", key="w_jf")
        if jf:
            try:
                st.session_state.jours_feries = [datetime.strptime(d.strip(), "%d/%m/%Y").date() for d in jf.split(",") if d.strip()]
            except:
                st.warning("Format invalide")

        st.markdown("---")
        st.info("💡 Vendredi et Samedi exclus. Dimanche est travaillable.")

    # Onglets
    tabs = st.tabs(["🏠 Accueil", "👥 Enseignants", "📚 Planning par Promotion", "🎯 Attributions", "📊 Export"])

    # ==================== ACCUEIL ====================
    with tabs[0]:
        st.markdown("""
        <div class="info-box">
            <h3>Gestion des Surveillances d'Examens</h3>
            <p><strong>S1 - 2026-2027</strong> | Filtre automatique: <b>Cours uniquement</b> (exclut TD et TP)</p>
            <ul>
                <li>📁 Chargement auto depuis <code>DATA-ENS-2026-2027_surveillances.xlsx</code></li>
                <li>📚 Uniquement les enseignements commencant par <b>Cours-</b></li>
                <li>👥 Permanents / Vacataires separes + Exclusion manuelle</li>
                <li>🏫 Salles (S01-S17) et Amphis (A01-A12) par promotion</li>
                <li>📅 <b>Ordre des matieres personnalisable</b></li>
                <li>🎯 Attribution intelligente (charge de matiere prioritaire)</li>
                <li>📊 Export HTML, Excel, PDF</li>
            </ul>
            <p><small>ℹ️ Seuls Vendredi et Samedi sont exclus automatiquement. Dimanche est travaillable.</small></p>
        </div>
        """, unsafe_allow_html=True)

        if st.session_state.data_loaded:
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.metric("Enseignants", len(st.session_state.enseignants_df))
            with c2: st.metric("Cours filtres", len(st.session_state.examens_df))
            with c3: st.metric("Promotions", len(st.session_state.promotions_list))
            with c4: st.metric("Permanents", len(st.session_state.permanents_list))

            with st.expander("👁️ Apercu des Cours extraits"):
                preview_df = st.session_state.examens_df[['matiere', 'promotion', 'enseignant']].head(20)
                if not preview_df.empty:
                    st.dataframe(preview_df, use_container_width=True, hide_index=True)
                else:
                    st.info("Aucun cours a afficher.")

    # ==================== ENSEIGNANTS ====================
    with tabs[1]:
        st.markdown('<div class="sub-header">Gestion des Enseignants</div>', unsafe_allow_html=True)

        if st.session_state.data_loaded:
            df_ens = st.session_state.enseignants_df

            # EXCLUSION MANUELLE
            st.markdown("### 🚫 Enseignants Exclus (0 surveillance)")
            st.markdown("<div class='warning-box'>Ces enseignants ne seront <b>JAMAIS</b> assignes a une surveillance.</div>", unsafe_allow_html=True)

            all_ens = st.session_state.all_enseignants_list
            exclus = st.multiselect("Selectionner les enseignants a EXCLURE", sorted(all_ens),
                                    default=st.session_state.exclus_manuels, key="w_exclus",
                                    help="0 surveillance pour ces enseignants")
            st.session_state.exclus_manuels = exclus
            if exclus:
                st.error(f"❌ Exclus: {', '.join(exclus)}")

            st.markdown("---")

            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### 👔 Permanents")
                perm_df = df_ens[df_ens['qualite'] == 'Permanent'][['nom', 'enseignements', 'promotion']].copy()
                perm_df['Exclu'] = perm_df['nom'].apply(lambda x: '❌ OUI' if x in exclus else '✅ Non')
                perm_df = perm_df.sort_values('Exclu', ascending=False)
                st.markdown(f"<div class='card'><b>Total:</b> {len(perm_df)} | <b>Actifs:</b> {len(perm_df[perm_df['Exclu']=='✅ Non'])} | <b>Exclus:</b> {len(perm_df[perm_df['Exclu']=='❌ OUI'])}</div>", unsafe_allow_html=True)
                if not perm_df.empty:
                    st.dataframe(perm_df, use_container_width=True, hide_index=True)
                else:
                    st.info("Aucun permanent.")

            with col2:
                st.markdown("#### 📝 Vacataires")
                vac_df = df_ens[df_ens['qualite'] == 'Vacataire'][['nom', 'enseignements', 'promotion']].copy()
                vac_df['Exclu'] = vac_df['nom'].apply(lambda x: '❌ OUI' if x in exclus else '✅ Non')
                vac_df = vac_df.sort_values('Exclu', ascending=False)
                st.markdown(f"<div class='card'><b>Total:</b> {len(vac_df)} | <b>Actifs:</b> {len(vac_df[vac_df['Exclu']=='✅ Non'])} | <b>Exclus:</b> {len(vac_df[vac_df['Exclu']=='❌ OUI'])}</div>", unsafe_allow_html=True)
                if not vac_df.empty:
                    st.dataframe(vac_df, use_container_width=True, hide_index=True)
                else:
                    st.info("Aucun vacataire.")

            st.markdown("---")
            st.markdown("#### 📈 Resume")
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1: st.metric("Perm. actifs", len(perm_df[perm_df['Exclu']=='✅ Non']) if not perm_df.empty else 0)
            with c2: st.metric("Vac. actifs", len(vac_df[vac_df['Exclu']=='✅ Non']) if not vac_df.empty else 0)
            with c3: st.metric("Autres", len(df_ens[~df_ens['qualite'].isin(['Permanent', 'Vacataire'])]))
            with c4: st.metric("Total actifs", len(df_ens) - len(exclus))
            with c5: st.metric("Exclus", len(exclus))
        else:
            st.warning("Donnees non chargees.")

    # ==================== PLANNING PAR PROMOTION ====================
    with tabs[2]:
        st.markdown('<div class="sub-header">Planification par Promotion</div>', unsafe_allow_html=True)

        if st.session_state.data_loaded and st.session_state.promotions_list:
            promo_selected = st.selectbox("📚 Selectionner une promotion", st.session_state.promotions_list,
                                          index=0, key="w_promo_sel")
            st.session_state.promo_selected = promo_selected

            if promo_selected:
                st.markdown(f"### ⚙️ Configuration: **{promo_selected}**")

                # LIEUX PAR PROMOTION
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.markdown("#### 🕐 Creneaux")
                    creneaux_sel = st.multiselect("Creneaux", CRENEAUX, default=CRENEAUX, key=f"w_cren_{promo_selected}")
                with col2:
                    st.markdown("#### 🏫 Salles S01-S17")
                    salles_sel = st.multiselect("Salles", SALLES, default=['S01', 'S02', 'S03', 'S04', 'S05'],
                                                key=f"w_sal_{promo_selected}")
                with col3:
                    st.markdown("#### 🎓 Amphis A01-A12")
                    amphis_sel = st.multiselect("Amphis", AMPHIS, default=['A01', 'A02', 'A03'],
                                                key=f"w_amp_{promo_selected}")

                lieux_sel = salles_sel + amphis_sel
                st.session_state.lieux_par_promo[promo_selected] = lieux_sel

                nb_examens = len(st.session_state.examens_df[st.session_state.examens_df['promotion'].astype(str).str.strip() == str(promo_selected).strip()])
                st.markdown(f"<div class='info-box'>📊 <b>{nb_examens}</b> cours | Lieux: <b>{len(lieux_sel)}</b> selectionnes ({', '.join(lieux_sel) if lieux_sel else 'Aucun'})</div>", unsafe_allow_html=True)

                # === ORDRE DES MATIERES ===
                st.markdown("---")
                st.markdown("#### 📋 Ordre des Matieres dans le Calendrier")
                st.markdown("<div class='info-box'>Definissez l'ordre des matieres. La matiere n°1 sera la premiere, n°2 la deuxieme, etc. Les matieres sans ordre seront placees a la fin.</div>", unsafe_allow_html=True)

                promo_examens = st.session_state.examens_df[st.session_state.examens_df['promotion'].astype(str).str.strip() == str(promo_selected).strip()].copy()

                if not promo_examens.empty:
                    ordre_df = promo_examens[['matiere', 'enseignant']].copy()
                    ordre_df['Ordre'] = range(1, len(ordre_df) + 1)

                    if promo_selected in st.session_state.ordre_matieres:
                        existing_ordre = st.session_state.ordre_matieres[promo_selected]
                        ordre_map = {m: i+1 for i, m in enumerate(existing_ordre)}
                        ordre_df['Ordre'] = ordre_df['matiere'].map(ordre_map).fillna(999).astype(int)

                    ordre_df = ordre_df.sort_values('Ordre')

                    edited_ordre = st.data_editor(
                        ordre_df,
                        column_config={
                            "matiere": st.column_config.TextColumn("Matiere", disabled=True),
                            "enseignant": st.column_config.TextColumn("Enseignant", disabled=True),
                            "Ordre": st.column_config.NumberColumn("Ordre (1= premiere)", min_value=1, max_value=999, step=1)
                        },
                        use_container_width=True,
                        num_rows="fixed",
                        key=f"ordre_editor_{promo_selected}"
                    )

                    edited_ordre = edited_ordre.sort_values('Ordre')
                    st.session_state.ordre_matieres[promo_selected] = edited_ordre['matiere'].tolist()

                    st.markdown("<div class='success-box'>✅ Ordre enregistre: " + " → ".join(edited_ordre['matiere'].tolist()[:5]) + ("..." if len(edited_ordre) > 5 else "") + "</div>", unsafe_allow_html=True)

                # Bouton generation
                st.markdown("---")
                if st.button(f"🚀 Generer le Planning pour {promo_selected}", type="primary", key=f"btn_gen_{promo_selected}"):
                    if len(lieux_sel) == 0:
                        st.error("❌ Selectionnez au moins un lieu.")
                    elif len(creneaux_sel) == 0:
                        st.error("❌ Selectionnez au moins un creneau.")
                    else:
                        with st.spinner(f"Generation pour {promo_selected}..."):
                            ordre = st.session_state.ordre_matieres.get(promo_selected, None)
                            planning = generer_planning_promo(
                                st.session_state.examens_df.copy(),
                                promo_selected,
                                st.session_state.date_debut_val,
                                st.session_state.jours_feries,
                                creneaux_sel,
                                lieux_sel,
                                ordre
                            )
                            if planning is not None:
                                st.session_state.planning_df = planning
                                st.success(f"✅ Planning genere pour {promo_selected}!")
                            else:
                                st.error("❌ Erreur de generation.")

                # Affichage planning
                if st.session_state.planning_df is not None:
                    st.markdown("---")
                    st.markdown(f"#### 📝 Planning de {promo_selected} (modifiable)")

                    planning_display = st.session_state.planning_df[
                        st.session_state.planning_df['promotion'].astype(str).str.strip() == str(promo_selected).strip()
                    ].copy()

                    if not planning_display.empty:
                        edited = st.data_editor(
                            planning_display,
                            column_config={
                                "date": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
                                "creneau": st.column_config.SelectboxColumn("Creneau", options=CRENEAUX),
                                "lieu": st.column_config.SelectboxColumn("Lieu", options=SALLES + AMPHIS),
                                "matiere": st.column_config.TextColumn("Matiere", disabled=True),
                                "promotion": st.column_config.TextColumn("Promotion", disabled=True),
                                "enseignant": st.column_config.TextColumn("Enseignant", disabled=True),
                            },
                            use_container_width=True,
                            num_rows="fixed",
                            key=f"plan_ed_{promo_selected}"
                        )

                        for idx in edited.index:
                            for col in ['date', 'creneau', 'lieu']:
                                if col in edited.columns:
                                    st.session_state.planning_df.at[idx, col] = edited.at[idx, col]

                        st.markdown("#### 📅 Vue Calendrier")
                        cal_data = []
                        for _, row in edited.iterrows():
                            d = row.get('date', None)
                            if d is not None and not pd.isna(d):
                                ds = d.strftime('%d/%m/%Y') if hasattr(d, 'strftime') else str(d)
                                jr = JOURS_FR.get(d.strftime('%A'), d.strftime('%A')) if hasattr(d, 'strftime') else ''
                                cal_data.append({'Date': ds, 'Jour': jr, 'Creneau': row.get('creneau', ''),
                                                'Matiere': row.get('matiere', ''), 'Lieu': row.get('lieu', ''),
                                                'Enseignant': row.get('enseignant', '')})
                        if cal_data:
                            st.dataframe(pd.DataFrame(cal_data), use_container_width=True, hide_index=True)
                    else:
                        st.info("Aucun examen pour cette promotion.")
        else:
            st.warning("Aucune promotion detectee.")

    # ==================== ATTRIBUTIONS ====================
    with tabs[3]:
        st.markdown('<div class="sub-header">Attribution des Surveillants</div>', unsafe_allow_html=True)

        if st.session_state.planning_df is not None and st.session_state.enseignants_df is not None:
            if st.session_state.exclus_manuels:
                st.markdown(f"<div class='warning-box'>🚫 Exclus: {', '.join(st.session_state.exclus_manuels)}</div>", unsafe_allow_html=True)

            if st.button("🎯 Attribuer les Surveillants", type="primary", key="btn_attrib"):
                with st.spinner("Attribution en cours..."):
                    attributions, ens_maj = attribuer_surveillants(
                        st.session_state.planning_df,
                        st.session_state.enseignants_df,
                        st.session_state.nb_surv_par_lieu
                    )
                    if attributions is not None:
                        st.session_state.surveillance_df = attributions
                        st.session_state.enseignants_df = ens_maj
                        st.success(f"✅ {len(attributions)} attributions effectuees!")
                    else:
                        st.error("❌ Erreur.")

            if st.session_state.surveillance_df is not None:
                st.markdown("---")
                st.markdown("#### 📋 Tableau des Attributions")

                attr_data = []
                for attr in st.session_state.surveillance_df:
                    d = attr.get('date', None)
                    ds = d.strftime('%d/%m/%Y') if hasattr(d, 'strftime') else str(d) if d else ''
                    for surv in attr.get('details_surveillants', []):
                        attr_data.append({'Date': ds, 'Creneau': attr.get('creneau', ''),
                                         'Matiere': attr.get('matiere', ''), 'Promotion': attr.get('promotion', ''),
                                         'Lieu': attr.get('lieu', ''), 'Surveillant': surv['nom'],
                                         'Qualite': surv['qualite'],
                                         'Role': 'Charge de matiere' if surv.get('priorite') == 'Charge de matiere' else surv['qualite']})

                if attr_data:
                    attr_df = pd.DataFrame(attr_data)
                    # CORRECTION: utiliser map au lieu de applymap (pandas 2.1+)
                    def color_q(val):
                        cmap = {'Permanent': 'background-color: #E3F2FD', 'Vacataire': 'background-color: #E8F5E9', 'Autre': 'background-color: #FFF3E0'}
                        return cmap.get(val, '')

                    try:
                        styled = attr_df.style.map(color_q, subset=['Qualite'])
                    except AttributeError:
                        # Fallback pour anciennes versions
                        styled = attr_df.style.applymap(color_q, subset=['Qualite'])

                    st.dataframe(styled, use_container_width=True, hide_index=True)
                else:
                    st.info("Aucune attribution a afficher.")

                # Chevauchements
                st.markdown("---")
                st.markdown("#### 🔍 Verification des Chevauchements")
                chevauchements = []
                ens_seances = {}
                for attr in st.session_state.surveillance_df:
                    for surv in attr.get('details_surveillants', []):
                        nom = surv['nom']
                        if nom not in ens_seances:
                            ens_seances[nom] = []
                        d = attr.get('date', None)
                        if hasattr(d, 'date'):
                            d = d.date()
                        elif isinstance(d, str):
                            try:
                                d = datetime.strptime(d, "%Y-%m-%d").date()
                            except:
                                d = None
                        ens_seances[nom].append({'date': d, 'creneau': attr.get('creneau', '')})

                for nom, seances in ens_seances.items():
                    seen = {}
                    for s in seances:
                        k = (s['date'], s['creneau'])
                        if k in seen:
                            chevauchements.append({'Enseignant': nom, 'Date': str(s['date']), 'Creneau': s['creneau']})
                        seen[k] = True

                if chevauchements:
                    st.error(f"⚠️ {len(chevauchements)} chevauchement(s)!")
                    st.dataframe(pd.DataFrame(chevauchements), use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Aucun chevauchement!")

                # Resume
                st.markdown("---")
                st.markdown("#### 📊 Resume par Enseignant")
                resume = {}
                for attr in st.session_state.surveillance_df:
                    for surv in attr.get('details_surveillants', []):
                        nom = surv['nom']
                        if nom not in resume:
                            resume[nom] = {'count': 0, 'qualite': surv['qualite']}
                        resume[nom]['count'] += 1

                if resume:
                    resume_df = pd.DataFrame([{'Enseignant': k, 'Qualite': v['qualite'], 'Surveillances': v['count']}
                                             for k, v in sorted(resume.items(), key=lambda x: x[1]['count'], reverse=True)])
                    st.dataframe(resume_df, use_container_width=True, hide_index=True)
                    if not resume_df.empty:
                        chart_data = resume_df.groupby('Qualite')['Surveillances'].sum().reset_index()
                        st.bar_chart(chart_data.set_index('Qualite'))
        else:
            st.warning("Generez d'abord le planning.")

    # ==================== EXPORT ====================
    with tabs[4]:
        st.markdown('<div class="sub-header">Export des Plannings</div>', unsafe_allow_html=True)

        if st.session_state.surveillance_df is not None:
            attributions = st.session_state.surveillance_df

            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown("#### 📄 HTML")
                html_content = generer_tableau_html(attributions)
                st.download_button("⬇️ Telecharger HTML", html_content, "planning_surveillances.html", "text/html", key="dl_html")
                with st.expander("Apercu"):
                    st.components.v1.html(html_content, height=600, scrolling=True)
            with col2:
                st.markdown("#### 📊 Excel")
                excel_buffer = generer_excel_colore(attributions)
                st.download_button("⬇️ Telecharger Excel", excel_buffer, "planning_surveillances.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")
            with col3:
                st.markdown("#### 📑 PDF")
                pdf_buffer = generer_pdf(attributions)
                st.download_button("⬇️ Telecharger PDF", pdf_buffer, "planning_surveillances.pdf", "application/pdf", key="dl_pdf")

            st.markdown("---")
            st.markdown("#### 🎨 Apercu du Planning Final")
            st.markdown(generer_tableau_html(attributions), unsafe_allow_html=True)
        else:
            st.warning("Aucune attribution a exporter.")

if __name__ == "__main__":
    main()
