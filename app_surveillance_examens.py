import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta, date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import io
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
import os
import re

st.set_page_config(page_title="Surveillances Examens S1 2026-2027", page_icon="📋", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.main-header { font-size: 2.2rem; font-weight: bold; color: #1f4e79; text-align: center; padding: 1rem; background: linear-gradient(90deg, #e3f2fd 0%, #bbdefb 100%); border-radius: 10px; margin-bottom: 1.5rem; }
.sub-header { font-size: 1.4rem; font-weight: bold; color: #1565c0; margin-top: 1rem; margin-bottom: 0.5rem; border-bottom: 2px solid #1565c0; padding-bottom: 0.3rem; }
.info-box { background-color: #e3f2fd; padding: 1rem; border-radius: 8px; border-left: 4px solid #1565c0; margin: 0.5rem 0; }
.success-box { background-color: #e8f5e9; padding: 1rem; border-radius: 8px; border-left: 4px solid #2e7d32; margin: 0.5rem 0; }
.warning-box { background-color: #fff3e0; padding: 1rem; border-radius: 8px; border-left: 4px solid #f57c00; margin: 0.5rem 0; }
.card { background-color: #fafafa; padding: 1rem; border-radius: 8px; border: 1px solid #e0e0e0; margin: 0.5rem 0; }
.stTabs [data-baseweb="tab-list"] { gap: 8px; }
.stTabs [data-baseweb="tab"] { background-color: #f5f5f5; border-radius: 8px 8px 0 0; padding: 10px 20px; font-weight: 600; }
.stTabs [aria-selected="true"] { background-color: #1565c0 !important; color: white !important; }
</style>
""", unsafe_allow_html=True)

SALLES = [f"S{i:02d}" for i in range(1, 18)]
AMPHIS = [f"A{i:02d}" for i in range(1, 13)]
CRENEAUX = ["08h30 - 10h30", "11h00 - 13h00", "13h30 - 15h30"]
JOURS_FR = {"Monday": "Lundi", "Tuesday": "Mardi", "Wednesday": "Mercredi", "Thursday": "Jeudi", "Friday": "Vendredi", "Saturday": "Samedi", "Sunday": "Dimanche"}
FICHIER_SOURCE = "DATA-ENS-2026-2027_surveillances.xlsx"

def init_session_state():
    defaults = {
        'enseignants_df': None, 'examens_df': None, 'planning_df': None,
        'surveillance_df': None, 'nb_surv_permanent': 3, 'nb_surv_vacataire': 2,
        'nb_surv_autre': 1, 'nb_surv_par_lieu': 2, 'exclus_manuels': [],
        'date_debut_val': date(2026, 11, 1), 'jours_feries': [],
        'promo_selected': None, 'data_loaded': False, 'promotions_list': [],
        'permanents_list': [], 'vacataires_list': [], 'all_enseignants_list': [],
        'ordre_matieres': {}, 'lieux_par_promo': {},
        'horaires_par_matiere': {}, 'edt_par_promo': {},
        'charges_matiere_assignees': {}  # NOUVEAU : Suivi des chargés assignés par matière/promo
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def normaliser_qualite(val):
    val = str(val).strip().lower()
    mapping = {'permanent': 'Permanent', 'vacataire': 'Vacataire', 'contractuel': 'Contractuel', 'autre': 'Autre',
               'professeur': 'Permanent', 'charge de cours': 'Vacataire', 'charge_de_cours': 'Vacataire',
               'doctorant': 'Vacataire', 'maitre de conferences': 'Permanent', 'mc': 'Permanent', 'prof': 'Permanent'}
    return mapping.get(val, 'Permanent')

def est_cours(enseignement_str):
    val = str(enseignement_str).strip()
    return bool(re.match(r'^[Cc][Oo][Uu][Rr][Ss]', val))

def extraire_nom_cours(enseignement_str):
    val = str(enseignement_str).strip()
    nom = re.sub(r'^[Cc][Oo][Uu][Rr][Ss][ \-_:]+', '', val).strip()
    return nom if nom else val

def charger_fichier_source_auto():
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
        df_ens = df_ens[df_ens['nom'].notna() & (df_ens['nom'].astype(str).str.strip() != '')].copy()
        df_ens['qualite'] = df_ens['qualite'].apply(normaliser_qualite)
        examens_data = []
        for _, row in df_ens.iterrows():
            raw_ens = str(row.get('enseignements', ''))
            items = re.split(r'[,;/]+', raw_ens)
            for item in items:
                item = item.strip()
                if item and est_cours(item):
                    nom_cours = extraire_nom_cours(item)
                    examens_data.append({'matiere': nom_cours, 'promotion': str(row.get('promotion', '')).strip(),
                        'enseignant': str(row.get('nom', '')).strip(), 'qualite_ens': row.get('qualite', 'Permanent'),
                        'date': None, 'creneau': None, 'lieu': None, 'ordre': 999})
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
        return {'enseignants': df_ens, 'examens': df_exam, 'promotions': promotions,
                'permanents': permanents, 'vacataires': vacataires, 'all_enseignants': all_ens, 'sheet_used': ens_sheet}, None
    except Exception as e:
        return None, str(e)

def est_jour_travaille(date_obj, jours_feries):
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

def generer_planning_promo(examens_df, promotion, date_debut, jours_feries, creneaux, lieux, ordre_matieres=None, horaires_matiere=None):
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
    if horaires_matiere and promotion in horaires_matiere:
        hmap = horaires_matiere[promotion]
        promo_df['creneau_pref'] = promo_df['matiere'].map(hmap)
    else:
        promo_df['creneau_pref'] = None
    date_courante = date_debut
    nb_lieux = len(lieux)
    if nb_lieux == 0:
        st.error("Veuillez selectionner au moins un lieu.")
        return None
    creneaux_dispo = creneaux if creneaux else CRENEAUX
    nb_creneaux = len(creneaux_dispo)
    if nb_creneaux == 0:
        st.error("Veuillez selectionner au moins un creneau.")
        return None
    idx = 0
    lieu_idx = 0
    for i in promo_df.index:
        while not est_jour_travaille(date_courante, jours_feries):
            date_courante += timedelta(days=1)
        creneau_pref = promo_df.at[i, 'creneau_pref']
        if creneau_pref and creneau_pref in creneaux_dispo:
            creneau = creneau_pref
        else:
            creneau = creneaux_dispo[idx % nb_creneaux]
        lieu = lieux[lieu_idx % nb_lieux]
        examens_df.at[i, 'date'] = date_courante
        examens_df.at[i, 'creneau'] = creneau
        examens_df.at[i, 'lieu'] = lieu
        idx += 1
        lieu_idx += 1
        if creneau == creneaux_dispo[-1]:
            date_courante += timedelta(days=1)
    return examens_df

def attribuer_surveillants(planning_df, enseignants_df, nb_par_lieu=2):
    """
    MODIFIÉ: 
    - Evite les doublons du chargé de matière pour la même matière/promotion
    - Classe les surveillants: Chargé de matière > Permanents > Vacataires > Autres
    """
    if planning_df is None or enseignants_df is None:
        return None, enseignants_df
    
    # Initialiser le suivi des chargés assignés
    if 'charges_matiere_assignees' not in st.session_state:
        st.session_state['charges_matiere_assignees'] = {}
    
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
        promotion_examen = examen.get('promotion', '')
        
        if date_examen is None or pd.isna(date_examen):
            continue
        
        surveillants_occupes = set()
        for attr in attributions:
            attr_date = attr.get('date', None)
            if attr_date is not None and attr.get('creneau') == creneau_examen:
                d1 = attr_date.date() if hasattr(attr_date, 'date') else attr_date
                d2 = date_examen.date() if hasattr(date_examen, 'date') else date_examen
                if isinstance(d1, str):
                    try: d1 = datetime.strptime(d1, "%Y-%m-%d").date()
                    except: d1 = None
                if isinstance(d2, str):
                    try: d2 = datetime.strptime(d2, "%Y-%m-%d").date()
                    except: d2 = None
                if d1 and d2 and d1 == d2:
                    surveillants_occupes.update(attr.get('surveillants', []))
        
        liste_surveillants = []
        clé_matiere = f"{promotion_examen}_{matiere_examen}"
        charge_deja_assigne = st.session_state['charges_matiere_assignees'].get(clé_matiere)
        
        # NOUVEAU: Assigner le chargé de matière UNE SEULE FOIS par matière/promotion
        if enseignant_matiere and str(enseignant_matiere) not in ['nan', '', 'None']:
            ens_info = surveillants[surveillants['nom'] == enseignant_matiere]
            if not ens_info.empty:
                nom_ens = ens_info.iloc[0]['nom']
                qualite_ens = ens_info.iloc[0]['qualite']
                
                # Vérifier que le chargé n'a pas déjà été assigné pour cette matière/promo
                if charge_deja_assigne != nom_ens and nom_ens not in surveillants_occupes and nom_ens not in exclus:
                    quota_key = f"nb_surv_{qualite_ens.lower()}"
                    quota = st.session_state.get(quota_key, 3)
                    current_count = surveillants.loc[surveillants['nom'] == nom_ens, 'surveillance_attribuee'].values[0]
                    if current_count < quota:
                        liste_surveillants.append({'nom': nom_ens, 'qualite': qualite_ens, 'priorite': 'Chargé de matière'})
                        surveillants_occupes.add(nom_ens)
                        surveillants.loc[surveillants['nom'] == nom_ens, 'surveillance_attribuee'] += 1
                        # Marquer que ce chargé a été assigné
                        st.session_state['charges_matiere_assignees'][clé_matiere] = nom_ens
        
        # NOUVEAU: Permanents d'abord (priorité après chargé de matière)
        for _, perm in permanents.iterrows():
            if len(liste_surveillants) >= nb_par_lieu: break
            if perm['nom'] in surveillants_occupes or perm['nom'] in exclus: continue
            quota_perm = st.session_state.get('nb_surv_permanent', 3)
            current_count = surveillants.loc[surveillants['nom'] == perm['nom'], 'surveillance_attribuee'].values[0]
            if current_count < quota_perm:
                liste_surveillants.append({'nom': perm['nom'], 'qualite': 'Permanent', 'priorite': 'Permanent'})
                surveillants_occupes.add(perm['nom'])
                surveillants.loc[surveillants['nom'] == perm['nom'], 'surveillance_attribuee'] += 1
        
        for _, vac in vacataires.iterrows():
            if len(liste_surveillants) >= nb_par_lieu: break
            if vac['nom'] in surveillants_occupes or vac['nom'] in exclus: continue
            quota_vac = st.session_state.get('nb_surv_vacataire', 2)
            current_count = surveillants.loc[surveillants['nom'] == vac['nom'], 'surveillance_attribuee'].values[0]
            if current_count < quota_vac:
                liste_surveillants.append({'nom': vac['nom'], 'qualite': 'Vacataire', 'priorite': 'Vacataire'})
                surveillants_occupes.add(vac['nom'])
                surveillants.loc[surveillants['nom'] == vac['nom'], 'surveillance_attribuee'] += 1
        
        for _, aut in autres.iterrows():
            if len(liste_surveillants) >= nb_par_lieu: break
            if aut['nom'] in surveillants_occupes or aut['nom'] in exclus: continue
            quota_aut = st.session_state.get('nb_surv_autre', 1)
            current_count = surveillants.loc[surveillants['nom'] == aut['nom'], 'surveillance_attribuee'].values[0]
            if current_count < quota_aut:
                liste_surveillants.append({'nom': aut['nom'], 'qualite': aut['qualite'], 'priorite': 'Autre'})
                surveillants_occupes.add(aut['nom'])
                surveillants.loc[surveillants['nom'] == aut['nom'], 'surveillance_attribuee'] += 1
        
        attributions.append({'date': date_examen, 'creneau': creneau_examen, 'matiere': matiere_examen,
            'promotion': promotion_examen, 'lieu': lieu_examen, 'enseignant': enseignant_matiere,
            'surveillants': [s['nom'] for s in liste_surveillants], 'details_surveillants': liste_surveillants})
    
    return attributions, surveillants

# ============================================================
# FONCTIONS EDT GRILLE - MODIFIÉES
# ============================================================

def construire_grille_edt(attributions, creneaux_liste):
    """
    MODIFIÉ:
    - Affiche Date + Horaire explicitement
    - Classe les surveillants: Chargé > Permanents > autres
    - Texte centré et structuré
    - Tri chronologique
    """
    if not attributions:
        return None, None, None
    
    # Trier les attributions par date et créneau (CHRONOLOGIQUE)
    attr_sorted = sorted(attributions, key=lambda x: (
        x.get('date') if x.get('date') is not None else datetime.max,
        creneaux_liste.index(x.get('creneau', '')) if x.get('creneau', '') in creneaux_liste else 999
    ))
    
    grille = {}
    jours_ordre = []
    
    for attr in attr_sorted:
        date_val = attr.get('date', None)
        if date_val is None: continue
        
        if isinstance(date_val, str):
            try: date_val = datetime.strptime(date_val, "%Y-%m-%d").date()
            except: continue
        elif isinstance(date_val, datetime): date_val = date_val.date()
        
        jour_nom = JOURS_FR.get(date_val.strftime('%A'), date_val.strftime('%A'))
        date_str = date_val.strftime('%d/%m/%Y')
        cle_jour = f"{jour_nom}\n{date_str}"
        
        if cle_jour not in grille:
            grille[cle_jour] = {}
            jours_ordre.append(cle_jour)
        
        creneau = attr.get('creneau', '')
        if creneau not in grille[cle_jour]: 
            grille[cle_jour][creneau] = []
        
        survs = attr.get('details_surveillants', [])
        # NOUVEAU: Trier les surveillants - Chargé de matière en premier, puis Permanents
        survs_sorted = sorted(survs, key=lambda s: (
            0 if s['priorite'] == 'Chargé de matière' else 1 if s['qualite'] == 'Permanent' else 2
        ))
        
        surv_text = "\n".join([
            f"{'👑 ' if s['priorite'] == 'Chargé de matière' else '⭐ ' if s['qualite'] == 'Permanent' else '• '}{s['nom']} ({s['qualite']})"
            for s in survs_sorted
        ])
        
        grille[cle_jour][creneau].append({
            'matiere': attr.get('matiere', ''),
            'enseignant': attr.get('enseignant', ''),
            'lieu': attr.get('lieu', ''),
            'creneau': creneau,
            'surveillants': surv_text,
            'promotion': attr.get('promotion', '')
        })
    
    creneaux_utilises = creneaux_liste if creneaux_liste else CRENEAUX
    data = []
    
    for creneau in creneaux_utilises:
        row = {'Creneau': creneau}
        for jour in jours_ordre:
            exams = grille.get(jour, {}).get(creneau, [])
            if exams:
                cellules = []
                for ex in exams:
                    # NOUVEAU: Format amélioré avec date/horaire explicite et texte centré
                    cell_text = (
                        f"📖 {ex['matiere']}\n"
                        f"🕐 {ex['creneau']}\n"
                        f"👤 Chargé: {ex['enseignant']}\n"
                        f"🏫 {ex['lieu']}\n"
                        f"👮 Surveillance:\n{ex['surveillants']}"
                    )
                    cellules.append(cell_text)
                row[jour] = "\n{'─'*40}\n".join(cellules)
            else:
                row[jour] = ""
        data.append(row)
    
    df_grille = pd.DataFrame(data)
    return df_grille, jours_ordre, grille

def generer_excel_edt(df_grille, promotion):
    """MODIFIÉ: Texte centré, hauteur de ligne ajustée"""
    wb = Workbook()
    ws = wb.active
    ws.title = f"EDT {promotion}"
    
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    creneau_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
    creneau_font = Font(bold=True, size=10)
    cell_fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
    cell_font = Font(size=9)
    thin_border = Border(
        left=Side(style='thin', color='90CAF9'),
        right=Side(style='thin', color='90CAF9'),
        top=Side(style='thin', color='90CAF9'),
        bottom=Side(style='thin', color='90CAF9')
    )
    
    headers = ['Creneau'] + [c for c in df_grille.columns if c != 'Creneau']
    
    for c_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
    
    for r_idx, row in df_grille.iterrows():
        for c_idx, col_name in enumerate(headers, 1):
            val = row.get(col_name, '')
            cell = ws.cell(row=r_idx + 2, column=c_idx, value=val)
            cell.border = thin_border
            # NOUVEAU: Texte centré
            if col_name == 'Creneau':
                cell.fill = creneau_fill
                cell.font = creneau_font
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            else:
                cell.fill = cell_fill
                cell.font = cell_font
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        max_lines = 1
        for cell in row:
            if cell.value:
                lines = str(cell.value).count('\n') + 1
                max_lines = max(max_lines, lines)
        ws.row_dimensions[row[0].row].height = max(80, max_lines * 18)
    
    ws.column_dimensions['A'].width = 18
    for col_idx in range(2, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = 40
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generer_html_edt(df_grille, promotion):
    """MODIFIÉ: Texte centré, structure améliorée, horaire explicite"""
    jours_cols = [c for c in df_grille.columns if c != 'Creneau']
    html = """
    <style>
        .edt-table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 11px; }
        .edt-table th { background-color: #1565C0; color: white; padding: 12px; text-align: center; border: 2px solid #0D47A1; font-size: 12px; font-weight: bold; }
        .edt-table td { padding: 12px; border: 1px solid #90CAF9; vertical-align: top; min-width: 220px; text-align: center; }
        .creneau-cell { background-color: #E3F2FD; font-weight: bold; text-align: center; font-size: 12px; width: 120px; }
        .exam-cell { background-color: #FFF8E1; text-align: center; }
        .matiere { font-weight: bold; color: #1565C0; font-size: 12px; text-align: center; }
        .horaire { color: #E65100; font-size: 11px; font-weight: bold; text-align: center; }
        .charge { color: #6A1B9A; font-weight: bold; text-align: center; font-size: 11px; }
        .lieu { color: #E65100; font-size: 10px; font-weight: bold; text-align: center; }
        .surv-header { color: #2E7D32; font-weight: bold; text-align: center; font-size: 11px; }
        .surv { color: #2E7D32; font-size: 10px; text-align: center; }
        .sep { border-top: 2px solid #ccc; margin: 8px 0; padding: 8px 0; }
    </style>
    <h2 style="color:#1565C0; text-align:center;">📅 EDT EXAMENS - Promotion {promotion}</h2>
    <table class="edt-table">
    """
    html += "<tr><th>Créneau</th>"
    for jour in jours_cols:
        html += f"<th>{jour.replace(chr(10), '<br>')}</th>"
    html += "</tr>"
    
    for _, row in df_grille.iterrows():
        html += f"<tr><td class='creneau-cell'>{row['Creneau']}</td>"
        for jour in jours_cols:
            val = row.get(jour, '')
            if val:
                parts = val.split('\n' + '─'*40 + '\n')
                cells_html = []
                for part in parts:
                    lines = part.split('\n')
                    formatted = []
                    for line in lines:
                        if line.startswith('📖 '): 
                            formatted.append(f"<div class='matiere'>{line[2:]}</div>")
                        elif line.startswith('🕐 '): 
                            formatted.append(f"<div class='horaire'>{line[2:]}</div>")
                        elif line.startswith('👤 Chargé: '): 
                            formatted.append(f"<div class='charge'>{line[11:]}</div>")
                        elif line.startswith('🏫 '): 
                            formatted.append(f"<div class='lieu'>{line[2:]}</div>")
                        elif line.startswith('👮 Surveillance:'): 
                            formatted.append(f"<div class='surv-header'>Surveillance :</div>")
                        elif line.startswith('👑 '): 
                            formatted.append(f"<div class='surv'><strong>👑 {line[2:]}</strong> (Chargé)</div>")
                        elif line.startswith('⭐ '): 
                            formatted.append(f"<div class='surv'><strong>⭐ {line[2:]}</strong> (Permanent)</div>")
                        elif line.startswith('• '): 
                            formatted.append(f"<div class='surv'>{line}</div>")
                        elif line.strip():
                            formatted.append(f"<div>{line}</div>")
                    cells_html.append("".join(formatted))
                content = "<div class='sep'></div>".join(cells_html)
                html += f"<td class='exam-cell'>{content}</td>"
            else:
                html += "<td></td>"
        html += "</tr>"
    html += "</table>"
    return html

def generer_pdf_edt(attributions, promotion, creneaux_liste):
    """MODIFIÉ: Affichage date/horaire, texte centré"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=1*cm, leftMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#1565C0'), spaceAfter=15, alignment=1)
    
    elements.append(Paragraph(f"📅 EDT EXAMENS - {promotion}", title_style))
    elements.append(Paragraph("Année 2026-2027 - Semestre 1", styles['Normal']))
    elements.append(Spacer(1, 0.3*cm))
    
    df_grille, jours_ordre, _ = construire_grille_edt(attributions, creneaux_liste)
    if df_grille is None:
        elements.append(Paragraph("Aucune donnée", styles['Normal']))
        doc.build(elements)
        buffer.seek(0)
        return buffer
    
    jours_cols = [c for c in df_grille.columns if c != 'Creneau']
    table_data = [['Créneau'] + [j.replace('\n', ' ') for j in jours_cols]]
    
    for _, row in df_grille.iterrows():
        row_data = [row['Creneau']]
        for jour in jours_cols:
            val = row.get(jour, '')
            if val:
                # Nettoyage et formatage pour PDF
                val = (val.replace('📖 ', 'Matière: ')
                         .replace('🕐 ', 'Horaire: ')
                         .replace('👤 Chargé: ', 'Chargé de matière: ')
                         .replace('🏫 ', 'Lieu: ')
                         .replace('👮 Surveillance:', 'Surveillance:')
                         .replace('👑 ', '★ ')
                         .replace('⭐ ', '● ')
                         .replace('• ', '○ '))
                row_data.append(val)
            else:
                row_data.append('')
        table_data.append(row_data)
    
    nb_cols = len(jours_cols) + 1
    available_width = 25 * cm
    col_widths = [3 * cm] + [(available_width - 3 * cm) / len(jours_cols)] * len(jours_cols)
    table = Table(table_data, repeatRows=1, colWidths=col_widths)
    
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#90CAF9')),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#E3F2FD'), colors.HexColor('#FFFFFF')]),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ============================================================
# FONCTIONS EXPORT EXISTANTES (inchangées mais peuvent être améliorées)
# ============================================================

def generer_tableau_html(attributions):
    if not attributions: return "<p>Aucune attribution</p>"
    planning_par_jour = {}
    for attr in attributions:
        date_val = attr.get('date', None)
        if date_val is None: continue
        if isinstance(date_val, str):
            try: date_val = datetime.strptime(date_val, "%Y-%m-%d").date()
            except: continue
        elif isinstance(date_val, datetime): date_val = date_val.date()
        date_str = date_val.strftime('%d/%m/%Y')
        jour_fr = JOURS_FR.get(date_val.strftime('%A'), date_val.strftime('%A'))
        cle = f"{jour_fr} {date_str}"
        if cle not in planning_par_jour: planning_par_jour[cle] = {}
        creneau = attr.get('creneau', '')
        if creneau not in planning_par_jour[cle]: planning_par_jour[cle][creneau] = []
        planning_par_jour[cle][creneau].append(attr)
    html = """
    <style>
        .planning-table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 11px; }
        .planning-table th { background-color: #1565C0; color: white; padding: 8px; text-align: center; border: 2px solid #0D47A1; }
        .planning-table td { padding: 6px; border: 1px solid #90CAF9; vertical-align: top; text-align: center; }
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
    html += "<tr><th>Créneau</th>"
    for jour in jours: html += f"<th>{jour}</th>"
    html += "</tr>"
    for creneau in CRENEAUX:
        html += f"<tr><td class='creneau-cell'>{creneau}</td>"
        for jour in jours:
            html += "<td>"
            if creneau in planning_par_jour.get(jour, {}):
                for examen in planning_par_jour[jour][creneau]:
                    survs = examen.get('details_surveillants', [])
                    surv_html = "<br>".join([f"<span class='surv-{s['qualite'].lower() if s.get('priorite') != 'Chargé de matière' else 'charge'}'>{s['nom']} ({s['qualite']}){'👑' if s.get('priorite') == 'Chargé de matière' else ''}</span>" for s in survs])
                    html += f"<div class='examen-cell'><strong>{examen.get('matiere', '')}</strong><br><small>Promo: {examen.get('promotion', '')} | {examen.get('lieu', '')}<br>Chargé: {examen.get('enseignant', '')}</small><br><small>{surv_html}</small></div>"
            html += "</td>"
        html += "</tr>"
    html += "</table>"
    return html

def generer_excel_colore(attributions):
    wb = Workbook()
    ws = wb.active
    ws.title = "Planning"
    
    # MODIFIÉ: Ajouter la colonne "Chargé de matière"
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
        data.append({
            'Date': date_str, 
            'Jour': jour, 
            'Créneau': attr.get('creneau', ''), 
            'Matière': attr.get('matiere', ''),
            'Chargé de matière': attr.get('enseignant', ''),
            'Promotion': attr.get('promotion', ''), 
            'Lieu': attr.get('lieu', ''), 
            'Surveillants': surv_str
        })
    
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
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = Border(left=Side(style='thin', color='90CAF9'), right=Side(style='thin', color='90CAF9'), top=Side(style='thin', color='90CAF9'), bottom=Side(style='thin', color='90CAF9'))
    
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value: max_length = max(max_length, len(str(cell.value)))
            except: pass
        ws.column_dimensions[column].width = min(max_length + 2, 60)
    
    ws2 = wb.create_sheet("Résumé Enseignants")
    ens_count = {}
    for attr in attributions:
        for surv in attr.get('details_surveillants', []):
            nom = surv['nom']
            if nom not in ens_count: ens_count[nom] = {'count': 0, 'qualite': surv['qualite'], 'priorite': surv['priorite'], 'examens': []}
            ens_count[nom]['count'] += 1
            date_val = attr.get('date', '')
            date_str = date_val.strftime('%d/%m') if hasattr(date_val, 'strftime') else str(date_val)
            ens_count[nom]['examens'].append(f"{attr.get('matiere', '')} ({date_str})")
    
    resume_data = []
    # NOUVEAU: Trier par priorité (Chargé > Permanent > autres) puis par count
    for nom, info in sorted(ens_count.items(), key=lambda x: (
        0 if x[1]['priorite'] == 'Chargé de matière' else 1 if x[1]['qualite'] == 'Permanent' else 2,
        -x[1]['count']
    )):
        resume_data.append({
            'Enseignant': nom, 
            'Rôle': info['priorite'],
            'Qualité': info['qualite'], 
            'Nombre': info['count'], 
            'Examens': "; ".join(info['examens'][:5]) + ("..." if len(info['examens']) > 5 else "")
        })
    
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
                if cell.value: max_length = max(max_length, len(str(cell.value)))
            except: pass
        ws2.column_dimensions[column].width = min(max_length + 2, 60)
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generer_pdf(attributions):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=1*cm, leftMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#1565C0'), spaceAfter=15, alignment=1)
    
    elements.append(Paragraph("📋 PLANNING DES SURVEILLANCES", title_style))
    elements.append(Paragraph("Année 2026-2027 - Semestre 1", styles['Heading2']))
    elements.append(Spacer(1, 0.3*cm))
    
    # Trier les attributions par date et créneau
    attr_sorted = sorted(attributions, key=lambda x: (
        x.get('date') if x.get('date') is not None else datetime.max,
        x.get('creneau', '')
    ))
    
    table_data = [['Date', 'Jour', 'Créneau', 'Matière', 'Chargé', 'Promotion', 'Lieu', 'Surveillants']]
    for attr in attr_sorted:
        date_val = attr.get('date', None)
        if hasattr(date_val, 'strftime'):
            date_str = date_val.strftime('%d/%m/%Y')
            jour = JOURS_FR.get(date_val.strftime('%A'), date_val.strftime('%A'))
        else:
            date_str = str(date_val)
            jour = ''
        surv_str = ", ".join([f"{s['nom']} ({s['qualite']})" for s in attr.get('details_surveillants', [])])
        table_data.append([
            date_str, 
            jour, 
            attr.get('creneau', ''), 
            attr.get('matiere', ''),
            attr.get('enseignant', ''),
            attr.get('promotion', ''), 
            attr.get('lieu', ''), 
            surv_str
        ])
    
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
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#E3F2FD'), colors.HexColor('#FFFFFF')]),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

def main():
    init_session_state()
    st.markdown('<div class="main-header">📋 GESTION DES SURVEILLANCES - S1 2026-2027</div>', unsafe_allow_html=True)
    
    # ... [reste du code main() inchangé] ...

if __name__ == "__main__":
    main()
