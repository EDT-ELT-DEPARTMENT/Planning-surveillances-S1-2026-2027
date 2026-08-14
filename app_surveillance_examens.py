from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
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
from reportlab.lib.units import cm
import os
import re
import time

# Titre officiel rappelé conformément aux consignes
TITRE_PLATEFORME = "Plateforme de gestion des EDTs-S2-2026-Département d'Électrotechnique-Faculté de génie électrique-UDL-SBA"

st.set_page_config(page_title=TITRE_PLATEFORME, page_icon="📋", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.main-header { font-size: 1.8rem; font-weight: bold; color: #1f4e79; text-align: center; padding: 1rem; background: linear-gradient(90deg, #e3f2fd 0%, #bbdefb 100%); border-radius: 10px; margin-bottom: 1.5rem; }
.sub-header { font-size: 1.4rem; font-weight: bold; color: #1565c0; margin-top: 1rem; margin-bottom: 0.5rem; border-bottom: 2px solid #1565c0; padding-bottom: 0.3rem; }
.info-box { background-color: #e3f2fd; padding: 1rem; border-radius: 8px; border-left: 4px solid #1565c0; margin: 0.5rem 0; text-align: center; }
.stTabs [data-baseweb="tab-list"] { gap: 8px; justify-content: center; }
.stTabs [data-baseweb="tab"] { background-color: #f5f5f5; border-radius: 8px 8px 0 0; padding: 10px 20px; font-weight: 600; text-align: center; }
.stTabs [aria-selected="true"] { background-color: #1565c0 !important; color: white !important; }
p, div, span, th, td { text-align: center !important; }
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
        'horaires_par_matiere': {}, 'jours_par_matiere': {}, 'edt_par_promo': {},
        'historique_edt': {}
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def normaliser_qualite(val):
    val = str(val).strip().lower()
    if 'vacataire' in val or 'charg' in val or 'contractuel' in val or 'doctorant' in val:
        return 'Vacataire'
    mapping = {
        'permanent': 'Permanent', 'vacataire': 'Vacataire', 'contractuel': 'Contractuel', 'autre': 'Autre',
        'professeur': 'Permanent', 'maitre de conferences': 'Permanent', 'mc': 'Permanent', 'prof': 'Permanent', 'mca': 'Permanent', 'mcb': 'Permanent'
    }
    for k, v in mapping.items():
        if k in val:
            return v
    return 'Permanent'

def corriger_nom_enseignant(nom_str):
    nom_clean = str(nom_str).strip()
    nom_lower = nom_clean.lower()
    corrections = {
        'belhadj': 'Belabed',
        'zeghdoudi': 'ZEGHOUDI',
        'babali': 'Bahlil'
    }
    for k, v in corrections.items():
        if k in nom_lower:
            return v
    return nom_clean

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
        
        # Application de la logique spécifique pour les noms de feuilles ('matières' et 'EDTCE')
        ens_sheet = None
        edt_sheet = None
        for sheet in sheet_names:
            s_lower = sheet.strip().lower()
            if s_lower in ['matières', 'matieres']:
                ens_sheet = sheet
            elif s_lower in ['edtce', 'ce', 'emploi_du_temps']:
                edt_sheet = sheet

        if ens_sheet is None:
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
        df_ens['nom'] = df_ens['nom'].apply(corriger_nom_enseignant)
        
        examens_data = []
        for _, row in df_ens.iterrows():
            raw_ens = str(row.get('enseignements', ''))
            items = re.split(r'[,;/]+', raw_ens)
            for item in items:
                item = item.strip()
                if item and est_cours(item):
                    nom_cours = extraire_nom_cours(item)
                    code_cours = f"CODE-{abs(hash(nom_cours)) % 9000 + 1000}"
                    examens_data.append({
                        'Enseignements': nom_cours, 
                        'Code': code_cours,
                        'Enseignants': str(row.get('nom', '')).strip(), 
                        'qualite_ens': row.get('qualite', 'Permanent'),
                        'Horaire': None, 'Jours': None, 'Lieu': None,
                        'Promotion': str(row.get('promotion', '')).strip(),
                        'ordre': 999
                    })
        df_exam = pd.DataFrame(examens_data)
        colonnes_attendues = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion']
        for col in colonnes_attendues:
            if col not in df_exam.columns:
                df_exam[col] = ''
                
        df_exam = df_exam.drop_duplicates(subset=['Enseignements', 'Promotion', 'Enseignants']).copy()
        df_exam = df_exam.sort_values('Enseignants').drop_duplicates(subset=['Enseignements', 'Promotion'], keep='first').copy()
        promotions = sorted(df_exam['Promotion'].dropna().astype(str).str.strip().unique().tolist()) if not df_exam.empty else []
        promotions = [p for p in promotions if p != '']
        
        df_ens['is_perm'] = df_ens['qualite'].apply(lambda x: 0 if x == 'Permanent' else 1)
        df_ens = df_ens.sort_values(['is_perm', 'nom']).drop(columns=['is_perm'])
        
        permanents = df_ens[df_ens['qualite'] == 'Permanent']['nom'].dropna().unique().tolist()
        vacataires = df_ens[df_ens['qualite'] == 'Vacataire']['nom'].dropna().unique().tolist()
        all_ens = df_ens['nom'].dropna().unique().tolist()
        return {'enseignants': df_ens, 'examens': df_exam, 'promotions': promotions,
                'permanents': permanents, 'vacataires': vacataires, 'all_enseignants': all_ens, 'sheet_used': ens_sheet, 'edt_sheet': edt_sheet}, None
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

def generer_planning_promo(examens_df, promotion, date_debut, jours_feries, creneaux, lieux, ordre_matieres=None, horaires_matiere=None, jours_matiere=None):
    if examens_df is None or examens_df.empty:
        return None
    promo_df = examens_df[examens_df['Promotion'].astype(str).str.strip() == str(promotion).strip()].copy()
    if promo_df.empty:
        return None
    if ordre_matieres and promotion in ordre_matieres:
        ordre_map = {m: i for i, m in enumerate(ordre_matieres[promotion])}
        promo_df['ordre'] = promo_df['Enseignements'].map(ordre_map).fillna(999).astype(int)
        promo_df = promo_df.sort_values('ordre')
    else:
        promo_df = promo_df.sort_values('Enseignements')
        
    nb_lieux = len(lieux)
    if nb_lieux == 0:
        st.error("Veuillez selectionner au moins un lieu.")
        return None
    creneaux_dispo = creneaux if creneaux else CRENEAUX
    if len(creneaux_dispo) == 0:
        st.error("Veuillez selectionner au moins un creneau.")
        return None
        
    creneaux_occupes = set()
    lieu_idx = 0
    date_courante = date_debut

    for i in promo_df.index:
        matiere_nom = promo_df.at[i, 'Enseignements']
        
        date_pref = jours_matiere.get(promotion, {}).get(matiere_nom) if jours_matiere else None
        if date_pref:
            if isinstance(date_pref, datetime):
                d_ex = date_pref.date()
            elif isinstance(date_pref, date):
                d_ex = date_pref
            else:
                try:
                    d_ex = datetime.strptime(str(date_pref), "%Y-%m-%d").date()
                except:
                    d_ex = date_debut
        else:
            d_ex = None
            
        creneau_pref = horaires_matiere.get(promotion, {}).get(matiere_nom) if horaires_matiere else None
        
        if d_ex and creneau_pref:
            date_examen = d_ex
            creneau = creneau_pref
        elif d_ex and not creneau_pref:
            creneau = None
            for c in creneaux_dispo:
                if (d_ex, c) not in creneaux_occupes:
                    creneau = c
                    break
            if not creneau:
                creneau = creneaux_dispo[0]
            date_examen = d_ex
        elif not d_ex and creneau_pref:
            d_test = date_debut
            while not est_jour_travaille(d_test, jours_feries) or (d_test, creneau_pref) in creneaux_occupes:
                d_test += timedelta(days=1)
            date_examen = d_test
            creneau = creneau_pref
        else:
            d_test = date_courante
            found = False
            while not found:
                while not est_jour_travaille(d_test, jours_feries):
                    d_test += timedelta(days=1)
                for c in creneaux_dispo:
                    if (d_test, c) not in creneaux_occupes:
                        date_examen = d_test
                        creneau = c
                        found = True
                        break
                if not found:
                    d_test += timedelta(days=1)
            date_courante = date_examen

        creneaux_occupes.add((date_examen, creneau))
        lieu = lieux[lieu_idx % nb_lieux]
        
        examens_df.loc[(examens_df['Promotion'].astype(str).str.strip() == str(promotion).strip()) & (examens_df['Enseignements'] == matiere_nom), 'date'] = date_examen
        examens_df.loc[(examens_df['Promotion'].astype(str).str.strip() == str(promotion).strip()) & (examens_df['Enseignements'] == matiere_nom), 'Horaire'] = creneau
        examens_df.loc[(examens_df['Promotion'].astype(str).str.strip() == str(promotion).strip()) & (examens_df['Enseignements'] == matiere_nom), 'Jours'] = JOURS_FR.get(date_examen.strftime("%A"), date_examen.strftime("%A"))
        examens_df.loc[(examens_df['Promotion'].astype(str).str.strip() == str(promotion).strip()) & (examens_df['Enseignements'] == matiere_nom), 'Lieu'] = lieu
        
        lieu_idx += 1
            
    planning_promo_result = examens_df[examens_df['Promotion'].astype(str).str.strip() == str(promotion).strip()].sort_values(by=['date', 'Horaire', 'Lieu'])
    return planning_promo_result

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
    charges_affectees_creneau = set()

    for idx, examen in planning_df.iterrows():
        date_examen = examen.get('date', None)
        creneau_examen = examen.get('Horaire', '')
        matiere_examen = examen.get('Enseignements', '')
        enseignant_matiere = examen.get('Enseignants', '')
        lieu_examen = examen.get('Lieu', 'S01')
        promotion_examen = examen.get('Promotion', '')
        
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
        
        cle_multi_lieux = (date_examen, creneau_examen, matiere_examen, enseignant_matiere)
        if enseignant_matiere and str(enseignant_matiere) not in ['nan', '', 'None']:
            if cle_multi_lieux not in charges_affectees_creneau:
                ens_info = surveillants[surveillants['nom'] == enseignant_matiere]
                if not ens_info.empty:
                    nom_ens = ens_info.iloc[0]['nom']
                    qualite_ens = ens_info.iloc[0]['qualite']
                    if nom_ens not in surveillants_occupes and nom_ens not in exclus:
                        quota_key = f"nb_surv_{qualite_ens.lower()}"
                        quota = st.session_state.get(quota_key, 3)
                        current_count = surveillants.loc[surveillants['nom'] == nom_ens, 'surveillance_attribuee'].values[0]
                        if current_count < quota:
                            liste_surveillants.append({'nom': nom_ens, 'qualite': qualite_ens, 'priorite': 'Charge de matiere'})
                            surveillants_occupes.add(nom_ens)
                            surveillants.loc[surveillants['nom'] == nom_ens, 'surveillance_attribuee'] += 1
                            charges_affectees_creneau.add(cle_multi_lieux)
                            
        if len(liste_surveillants) < 1:
            for _, perm in permanents.iterrows():
                if perm['nom'] in surveillants_occupes or perm['nom'] in exclus: continue
                quota_perm = st.session_state.get('nb_surv_permanent', 3)
                current_count = surveillants.loc[surveillants['nom'] == perm['nom'], 'surveillance_attribuee'].values[0]
                if current_count < quota_perm:
                    liste_surveillants.append({'nom': perm['nom'], 'qualite': 'Permanent', 'priorite': 'Permanent'})
                    surveillants_occupes.add(perm['nom'])
                    surveillants.loc[surveillants['nom'] == perm['nom'], 'surveillance_attribuee'] += 1
                    break
                    
        if nb_par_lieu >= 2 and len(liste_surveillants) == 1:
            vacataire_trouve = False
            for _, vac in vacataires.iterrows():
                if vac['nom'] in surveillants_occupes or vac['nom'] in exclus: continue
                quota_vac = st.session_state.get('nb_surv_vacataire', 2)
                current_count = surveillants.loc[surveillants['nom'] == vac['nom'], 'surveillance_attribuee'].values[0]
                if current_count < quota_vac:
                    liste_surveillants.append({'nom': vac['nom'], 'qualite': 'Vacataire', 'priorite': 'Vacataire'})
                    surveillants_occupes.add(vac['nom'])
                    surveillants.loc[surveillants['nom'] == vac['nom'], 'surveillance_attribuee'] += 1
                    vacataire_trouve = True
                    break
            if not vacataire_trouve:
                for _, perm in permanents.iterrows():
                    if perm['nom'] in surveillants_occupes or perm['nom'] in exclus: continue
                    current_count = surveillants.loc[surveillants['nom'] == perm['nom'], 'surveillance_attribuee'].values[0]
                    if current_count < st.session_state.get('nb_surv_permanent', 3):
                        liste_surveillants.append({'nom': perm['nom'], 'qualite': 'Permanent', 'priorite': 'Permanent'})
                        surveillants_occupes.add(perm['nom'])
                        surveillants.loc[surveillants['nom'] == perm['nom'], 'surveillance_attribuee'] += 1
                        break

        while len(liste_surveillants) < nb_par_lieu:
            assigned_added = False
            for _, aut in autres.iterrows():
                if aut['nom'] in surveillants_occupes or aut['nom'] in exclus: continue
                current_count = surveillants.loc[surveillants['nom'] == aut['nom'], 'surveillance_attribuee'].values[0]
                if current_count < st.session_state.get('nb_surv_autre', 1):
                    liste_surveillants.append({'nom': aut['nom'], 'qualite': aut['qualite'], 'priorite': 'Autre'})
                    surveillants_occupes.add(aut['nom'])
                    surveillants.loc[surveillants['nom'] == aut['nom'], 'surveillance_attribuee'] += 1
                    assigned_added = True
                    break
            if not assigned_added:
                break
                
        attributions.append({
            'date': date_examen, 'creneau': creneau_examen, 'matiere': matiere_examen,
            'enseignant': enseignant_matiere, 'promotion': promotion_examen, 'lieu': lieu_examen,
            'surveillants': [s['nom'] for s in liste_surveillants], 'details_surveillants': liste_surveillants
        })
        
    attributions = sorted(attributions, key=lambda x: (x.get('date', datetime.min), x.get('creneau', ''), x.get('promotion', '')))
    return attributions, surveillants

def construire_grille_edt(attributions, creneaux_liste):
    if not attributions:
        return None, None, None
    grille = {}
    jours_ordre = []
    for attr in attributions:
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
        if creneau not in grille[cle_jour]: grille[cle_jour][creneau] = []
        survs = attr.get('details_surveillants', [])
        surv_text = "\n".join([f"• {s['nom']} ({s['qualite']})" for s in survs])
        grille[cle_jour][creneau].append({
            'matiere': attr.get('matiere', ''), 
            'enseignant': attr.get('enseignant', ''),
            'lieu': attr.get('lieu', ''), 
            'surveillants': surv_text, 
            'promotion': attr.get('promotion', ''),
            'creneau': creneau,
            'date': date_str
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
                    cell_text = f"📖 {ex['matiere']}\n👤 Chargé: {ex['enseignant']}\n🏫 {ex['lieu']}\n👮\n{ex['surveillants']}"
                    cellules.append(cell_text)
                row[jour] = "\n---\n".join(cellules)
            else:
                row[jour] = ""
        data.append(row)
    df_grille = pd.DataFrame(data)
    return df_grille, jours_ordre, grille

def generer_excel_edt(df_grille, promotion):
    wb = Workbook()
    ws = wb.active
    ws.title = f"EDT {promotion}"
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    creneau_fill = PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid")
    creneau_font = Font(bold=True, size=10)
    cell_fill = PatternFill(start_color="FFF8E1", end_color="FFF8E1", fill_type="solid")
    cell_font = Font(size=9)
    thin_border = Border(left=Side(style='thin', color='90CAF9'), right=Side(style='thin', color='90CAF9'),
                         top=Side(style='thin', color='90CAF9'), bottom=Side(style='thin', color='90CAF9'))
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
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if col_name == 'Creneau':
                cell.fill = creneau_fill
                cell.font = creneau_font
            else:
                cell.fill = cell_fill
                cell.font = cell_font
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generer_html_edt(df_grille, promotion):
    jours_cols = [c for c in df_grille.columns if c != 'Creneau']
    html = f"""
    <style>
        .edt-table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 11px; text-align: center; }}
        .edt-table th {{ background-color: #1565C0; color: white; padding: 10px; text-align: center; border: 2px solid #0D47A1; font-size: 12px; }}
        .edt-table td {{ padding: 8px; border: 1px solid #90CAF9; vertical-align: middle; text-align: center; min-width: 200px; }}
        .creneau-cell {{ background-color: #E3F2FD; font-weight: bold; text-align: center; font-size: 12px; width: 120px; }}
        .exam-cell {{ background-color: #FFF8E1; text-align: center; }}
        .matiere {{ font-weight: bold; color: #1565C0; font-size: 12px; text-align: center; }}
        .ens {{ color: #333; font-size: 10px; text-align: center; }}
        .lieu {{ color: #E65100; font-size: 10px; font-weight: bold; text-align: center; }}
        .surv {{ color: #2E7D32; font-size: 10px; text-align: center; }}
        .sep {{ border-top: 1px dashed #ccc; margin: 4px 0; }}
    </style>
    <h2 style="color:#1565C0; text-align:center;">{TITRE_PLATEFORME}</h2>
    <h3 style="color:#333; text-align:center;">EDT Chronologique - Promotion {promotion}</h3>
    <table class="edt-table">
    """
    html += "<tr><th>Creneau / Horaire</th>"
    for jour in jours_cols:
        html += f"<th>{jour.replace(chr(10), '<br>')}</th>"
    html += "</tr>"
    for _, row in df_grille.iterrows():
        html += f"<tr><td class='creneau-cell'>{row['Creneau']}</td>"
        for jour in jours_cols:
            val = row.get(jour, '')
            if val:
                parts = val.split('\n---\n')
                cells_html = []
                for part in parts:
                    lines = part.split('\n')
                    formatted = []
                    for line in lines:
                        if line.startswith('📖 '): formatted.append(f"<div class='matiere'>{line[2:]}</div>")
                        elif line.startswith('👤 '): formatted.append(f"<div class='ens'>{line[2:]}</div>")
                        elif line.startswith('🏫 '): formatted.append(f"<div class='lieu'>{line[2:]}</div>")
                        elif line.startswith('👮'): formatted.append(f"<div class='surv'>{line}</div>")
                        elif line.startswith('• '): formatted.append(f"<div class='surv'>{line}</div>")
                        else: formatted.append(f"<div>{line}</div>")
                    cells_html.append("".join(formatted))
                content = "<div class='sep'></div>".join(cells_html)
                html += f"<td class='exam-cell'>{content}</td>"
            else:
                html += "<td></td>"
        html += "</tr>"
    html += "</table>"
    return html

def generer_pdf_edt(attributions, promotion, creneaux_liste):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=1*cm, leftMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    elements = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=13, textColor=colors.HexColor('#1565C0'), spaceAfter=6, alignment=1)
    subtitle_style = ParagraphStyle('SubTitle', parent=styles['Normal'], alignment=1, fontSize=10, textColor=colors.HexColor('#333333'), spaceAfter=10)
    
    elements.append(Paragraph(TITRE_PLATEFORME, title_style))
    elements.append(Paragraph(f"EDT Chronologique - Promotion {promotion}", subtitle_style))
    elements.append(Spacer(1, 0.2*cm))
    
    df_grille, jours_ordre, _ = construire_grille_edt(attributions, creneaux_liste)
    if df_grille is None:
        elements.append(Paragraph("Aucune donnee", styles['Normal']))
        doc.build(elements)
        buffer.seek(0)
        return buffer
        
    jours_cols = [c for c in df_grille.columns if c != 'Creneau']
    
    header_style = ParagraphStyle('TableHeader', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.whitesmoke, alignment=1, fontName='Helvetica-Bold')
    creneau_header_style = ParagraphStyle('CreneauHeader', parent=header_style, fontSize=8, leading=10)
    
    table_data = [[Paragraph('Creneau / Horaire', creneau_header_style)] + [Paragraph(j.replace('\n', '<br/>'), header_style) for j in jours_cols]]
    
    cell_style = ParagraphStyle('TableCell', parent=styles['Normal'], fontSize=7, leading=9, alignment=1, textColor=colors.HexColor('#333333'))
    creneau_cell_style = ParagraphStyle('CreneauCell', parent=cell_style, fontName='Helvetica-Bold', textColor=colors.HexColor('#1565C0'))
    
    for _, row in df_grille.iterrows():
        row_data = [Paragraph(str(row['Creneau']), creneau_cell_style)]
        for jour in jours_cols:
            val = row.get(jour, '')
            if val:
                parts = val.split('\n---\n')
                formatted_parts = []
                for part in parts:
                    lines = part.split('\n')
                    f_lines = []
                    for line in lines:
                        if line.startswith('📖 '):
                            f_lines.append(f"<b>{line}</b>")
                        elif line.startswith('👤 '):
                            f_lines.append(f"{line.replace('👤 ', '')}")
                        elif line.startswith('🏫 '):
                            f_lines.append(f"<font color='#E65100'><b>{line}</b></font>")
                        elif line.startswith('• '):
                            f_lines.append(f"<font color='#2E7D32'>{line}</font>")
                        elif line == '👮':
                            continue
                        else:
                            f_lines.append(line)
                    formatted_parts.append("<br/>".join(f_lines))
                cell_content = "<br/><br/>".join(formatted_parts)
                row_data.append(Paragraph(cell_content, cell_style))
            else:
                row_data.append(Paragraph("", cell_style))
        table_data.append(row_data)
        
    available_width = 27.7 * cm
    col_widths = [3 * cm] + [(available_width - 3 * cm) / len(jours_cols)] * len(jours_cols)
    
    table = Table(table_data, repeatRows=1, colWidths=col_widths)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#90CAF9')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#E3F2FD'), colors.HexColor('#FFFFFF')]),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
    ]))
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

def generer_tableau_html(attributions):
    if not attributions: return "<p style='text-align:center;'>Aucune attribution</p>"
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
    html = f"""
    <style>
        .planning-table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 11px; text-align: center; }}
        .planning-table th {{ background-color: #1565C0; color: white; padding: 8px; text-align: center; border: 2px solid #0D47A1; }}
        .planning-table td {{ padding: 6px; border: 1px solid #90CAF9; vertical-align: middle; text-align: center; }}
        .creneau-cell {{ background-color: #E3F2FD; font-weight: bold; text-align: center; width: 120px; }}
        .examen-cell {{ background-color: #FFF8E1; margin: 2px auto; padding: 5px; border-radius: 3px; border-left: 3px solid #FFA000; font-size: 10px; text-align: center; }}
    </style>
    <h3 style="color:#1565C0; text-align:center;">{TITRE_PLATEFORME}</h3>
    <table class="planning-table">
    """
    jours = sorted(planning_par_jour.keys())
    html += "<tr><th>Creneau / Horaire</th>"
    for jour in jours: html += f"<th>{jour}</th>"
    html += "</tr>"
    for creneau in CRENEAUX:
        html += f"<tr><td class='creneau-cell'>{creneau}</td>"
        for jour in jours:
            html += "<td>"
            if creneau in planning_par_jour.get(jour, {}):
                for examen in planning_par_jour[jour][creneau]:
                    survs = examen.get('details_surveillants', [])
                    surv_html = "<br>".join([f"<span>{s['nom']} ({s['qualite']}{'*' if s.get('priorite') == 'Charge de matiere' else ''})</span>" for s in survs])
                    html += f"<div class='examen-cell'><strong>{examen.get('matiere', '')}</strong><br><small>Promo: {examen.get('promotion', '')} | Lieu: {examen.get('lieu', '')}</small><br><small>Chargé: {examen.get('enseignant', '')}</small><br><small>{surv_html}</small></div>"
            html += "</td>"
        html += "</tr>"
    html += "</table>"
    return html

def generer_excel_colore(attributions):
    wb = Workbook()
    ws = wb.active
    ws.title = "Planning Global"
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
            'Enseignements': attr.get('matiere', ''), 
            'Code': f"CODE-{abs(hash(attr.get('matiere', ''))) % 9000 + 1000}", 
            'Enseignants': attr.get('enseignant', ''), 
            'Horaire': attr.get('creneau', ''), 
            'Jours': jour, 
            'Lieu': attr.get('lieu', ''), 
            'Promotion': attr.get('promotion', ''),
            'Date': date_str,
            'Surveillants': surv_str
        })
    df = pd.DataFrame(data)
    # Respect strict de la disposition demandée
    cols_order = ['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion', 'Date', 'Surveillants']
    df = df[[c for c in cols_order if c in df.columns]]
    
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
            else:
                cell.border = Border(left=Side(style='thin', color='90CAF9'), right=Side(style='thin', color='90CAF9'), top=Side(style='thin', color='90CAF9'), bottom=Side(style='thin', color='90CAF9'))
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value: max_length = max(max_length, len(str(cell.value)))
            except: pass
        ws.column_dimensions[column].width = min(max(max_length + 2, 15), 50)
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generer_pdf(attributions):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=1*cm, leftMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('CustomTitle', parent=styles['Heading1'], fontSize=12, textColor=colors.HexColor('#1565C0'), spaceAfter=10, alignment=1)
    elements.append(Paragraph(TITRE_PLATEFORME, title_style))
    elements.append(Paragraph("PLANNING CHRONOLOGIQUE DES SURVEILLANCES", styles['Heading2']))
    elements.append(Spacer(1, 0.3*cm))
    table_data = [['Enseignements', 'Code', 'Enseignants', 'Horaire', 'Jours', 'Lieu', 'Promotion', 'Date', 'Surveillants']]
    for attr in attributions:
        date_val = attr.get('date', None)
        if hasattr(date_val, 'strftime'):
            date_str = date_val.strftime('%d/%m/%Y')
            jour = JOURS_FR.get(date_val.strftime('%A'), date_val.strftime('%A'))
        else:
            date_str = str(date_val)
            jour = ''
        surv_str = ", ".join([f"{s['nom']} ({s['qualite']})" for s in attr.get('details_surveillants', [])])
        table_data.append([
            attr.get('matiere', ''), 
            f"CODE-{abs(hash(attr.get('matiere', ''))) % 9000 + 1000}", 
            attr.get('enseignant', ''), 
            attr.get('creneau', ''), 
            jour, 
            attr.get('lieu', ''), 
            attr.get('promotion', ''), 
            date_str, 
            surv_str
        ])
    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#90CAF9')),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 7),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#E3F2FD'), colors.HexColor('#FFFFFF')]),
    ]))
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

def main():
    init_session_state()
    st.markdown(f'<div class="main-header">{TITRE_PLATEFORME}</div>', unsafe_allow_html=True)

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
                if result['promotions']: st.session_state.promo_selected = result['promotions'][0]
                st.success(f"✅ Chargé: {result['sheet_used']} | {len(result['enseignants'])} ens. ({len(result['vacataires'])} vacataires) | {len(result['examens'])} cours | Promos: {', '.join(result['promotions'])}")
            else:
                st.error(f"❌ {error}")

    with st.sidebar:
        st.markdown("## ⚙️ Configuration")
        if st.session_state.data_loaded:
            st.markdown("### 📁 Source")
            st.success(f"{FICHIER_SOURCE} chargé")
            st.markdown(f"- Enseignants: {len(st.session_state.enseignants_df)}")
            st.markdown(f"- Vacataires détectés: {len(st.session_state.vacataires_list)}")
            st.markdown(f"- **Cours uniquement**: {len(st.session_state.examens_df)}")
            st.markdown(f"- Promotions: {', '.join(st.session_state.promotions_list)}")
            if st.button("🔄 Recharger", key="btn_reload"):
                st.session_state.data_loaded = False
                st.rerun()
        else:
            st.warning("Fichier non chargé")
            fu = st.file_uploader("Charger manuellement", type=['xlsx', 'xls'], key="manual_up")
            if fu is not None:
                with open(FICHIER_SOURCE, "wb") as f: f.write(fu.getvalue())
                st.session_state.data_loaded = False
                st.rerun()
        st.markdown("---")
        st.markdown("### 📊 Quotas")
        st.session_state.nb_surv_permanent = st.number_input("Permanent", 0, 20, st.session_state.nb_surv_permanent, key="w_qp")
        st.session_state.nb_surv_vacataire = st.number_input("Vacataire", 0, 20, st.session_state.nb_surv_vacataire, key="w_qv")
        st.session_state.nb_surv_autre = st.number_input("Autre", 0, 20, st.session_state.nb_surv_autre, key="w_qa")
        st.session_state.nb_surv_par_lieu = st.number_input("Surv. par lieu", 1, 5, st.session_state.nb_surv_par_lieu, key="w_nl")
        st.markdown("---")
        st.markdown("### 📅 Date de Début")
        st.session_state.date_debut_val = st.date_input("Date début", st.session_state.date_debut_val, key="w_dd")
        
        st.markdown("### 🎉 Jours Fériés (Calendrier)")
        if 'jours_feries_list' not in st.session_state:
            st.session_state.jours_feries_list = []
            
        col_pick, col_add = st.columns([2, 1])
        with col_pick:
            date_ferie_choisie = st.date_input("Choisir un férié", key="picker_ferie_unique")
        with col_add:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Ajouter", key="btn_add_ferie"):
                if date_ferie_choisie not in st.session_state.jours_feries_list:
                    st.session_state.jours_feries_list.append(date_ferie_choisie)
                    st.session_state.jours_feries = sorted(st.session_state.jours_feries_list)
                    st.rerun()
                    
        if st.session_state.jours_feries_list:
            st.write("Jours fériés enregistrés :")
            for jf_item in list(st.session_state.jours_feries_list):
                c_jf1, c_jf2 = st.columns([3, 1])
                with c_jf1:
                    st.text(jf_item.strftime("%d/%m/%Y"))
                with c_jf2:
                    if st.button("❌", key=f"del_ferie_{jf_item}"):
                        st.session_state.jours_feries_list.remove(jf_item)
                        st.session_state.jours_feries = sorted(st.session_state.jours_feries_list)
                        st.rerun()

        st.markdown("---")
        st.info("💡 Vendredi et Samedi exclus. Dimanche est travaillable.")

    tabs = st.tabs(["🏠 Accueil", "👥 Enseignants", "📚 Planning par Promotion", "🎯 Attributions", "📅 EDT par Promotion", "📂 Répertoire des EDTs", "📊 Export Global"])

    with tabs[0]:
        st.markdown(f"""
        <div class="info-box">
            <h3>{TITRE_PLATEFORME}</h3>
            <p><strong>Gestion des plannings d'examens et surveillances</strong> | Filtre: <b>Cours uniquement</b></p>
            <ul>
                <li>📁 Chargement automatique depuis <code>{FICHIER_SOURCE}</code> (Feuille <code>matières</code> prioritaire)</li>
                <li>📚 Uniquement les enseignements commençant par <b>Cours-</b></li>
                <li>🎯 Vacataire configuré en <b>deuxième position</b> pour chaque lieu</li>
                <li>🛠️ <b>Sélection manuelle par matière</b> activée</li>
                <li>🎉 <b>Sélection des jours fériés depuis le calendrier interactif</b></li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.session_state.data_loaded:
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.metric("Enseignants", len(st.session_state.enseignants_df))
            with c2: st.metric("Vacataires", len(st.session_state.vacataires_list))
            with c3: st.metric("Cours filtrés", len(st.session_state.examens_df))
            with c4: st.metric("Promotions", len(st.session_state.promotions_list))

    with tabs[1]:
        st.markdown('<div class="sub-header">Gestion des Enseignants et Qualité</div>', unsafe_allow_html=True)
        if st.session_state.data_loaded:
            df_ens = st.session_state.enseignants_df
            all_ens = st.session_state.all_enseignants_list
            exclus = st.multiselect("Sélectionner les enseignants à EXCLURE", sorted(all_ens), default=st.session_state.exclus_manuels, key="w_exclus")
            st.session_state.exclus_manuels = exclus
            if not df_ens.empty:
                disp_ens = df_ens[['nom', 'qualite', 'enseignements', 'promotion']].copy()
                disp_ens['Exclu'] = disp_ens['nom'].apply(lambda x: '❌ OUI' if x in exclus else '✅ Non')
                disp_ens = disp_ens.sort_values(by=['qualite', 'nom'], ascending=[True, True])
                st.dataframe(disp_ens, use_container_width=True, hide_index=True)
        else: st.warning("Données non chargées.")

    with tabs[2]:
        st.markdown('<div class="sub-header">Planification par Promotion (Sélection manuelle par matière)</div>', unsafe_allow_html=True)
        
        col_h1, col_h2 = st.columns([1, 4])
        with col_h1:
            if st.button("🗑️ Supprimer l'historique", type="secondary", key="btn_suppr_hist"):
                st.session_state.historique_edt = {}
                st.session_state.planning_df = None
                st.session_state.surveillance_df = None
                st.success("L'historique des EDTs a été supprimé avec succès.")
                st.rerun()
                
        if st.session_state.data_loaded and st.session_state.promotions_list:
            promo_selected = st.selectbox("📚 Sélectionner une promotion", st.session_state.promotions_list, index=0, key="w_promo_sel")
            st.session_state.promo_selected = promo_selected
            if promo_selected:
                col1, col2 = st.columns(2)
                with col1: salles_sel = st.multiselect("Salles", SALLES, default=['S01', 'S02', 'S03', 'S04', 'S05'], key=f"w_sal_{promo_selected}")
                with col2: amphis_sel = st.multiselect("Amphis", AMPHIS, default=['A01', 'A02', 'A03'], key=f"w_amp_{promo_selected}")
                lieux_sel = salles_sel + amphis_sel
                st.session_state.lieux_par_promo[promo_selected] = lieux_sel
                
                st.markdown("#### 🕒 Personnalisation de la Date et de l'Horaire par Matière")
                df_promo_matieres = st.session_state.examens_df[st.session_state.examens_df['Promotion'].astype(str).str.strip() == str(promo_selected).strip()]
                
                if not df_promo_matieres.empty:
                    if promo_selected not in st.session_state.horaires_par_matiere:
                        st.session_state.horaires_par_matiere[promo_selected] = {}
                    if promo_selected not in st.session_state.jours_par_matiere:
                        st.session_state.jours_par_matiere[promo_selected] = {}
                        
                    for _, row_mat in df_promo_matieres.iterrows():
                        m_nom = row_mat['Enseignements']
                        col_m1, col_m2, col_m3 = st.columns([2, 1, 1])
                        with col_m1:
                            st.text(f"📖 {m_nom} (Resp: {row_mat['Enseignants']})")
                        with col_m2:
                            date_actuelle = st.session_state.jours_par_matiere.get(promo_selected, {}).get(m_nom, None)
                            activer_date_fixe = st.checkbox(f"Date fixe - {m_nom}", value=(date_actuelle is not None), key=f"chk_date_{promo_selected}_{m_nom}")
                            if activer_date_fixe:
                                d_val = date_actuelle if isinstance(date_actuelle, date) else st.session_state.date_debut_val
                                d_choisie = st.date_input(f"Date - {m_nom}", value=d_val, key=f"d_choisie_{promo_selected}_{m_nom}")
                                st.session_state.jours_par_matiere.setdefault(promo_selected, {})[m_nom] = d_choisie
                            else:
                                if promo_selected in st.session_state.jours_par_matiere and m_nom in st.session_state.jours_par_matiere[promo_selected]:
                                    del st.session_state.jours_par_matiere[promo_selected][m_nom]
                        with col_m3:
                            options_h = ["Sélection manuelle"] + CRENEAUX
                            horaire_actuel = st.session_state.horaires_par_matiere.get(promo_selected, {}).get(m_nom, "Sélection manuelle")
                            try:
                                idx_h = options_h.index(horaire_actuel)
                            except:
                                idx_h = 0
                            h_choisi = st.selectbox(f"Horaire - {m_nom}", options_h, index=idx_h, key=f"h_{promo_selected}_{m_nom}")
                            if h_choisi != "Sélection manuelle":
                                st.session_state.horaires_par_matiere.setdefault(promo_selected, {})[m_nom] = h_choisi
                            else:
                                if promo_selected in st.session_state.horaires_par_matiere and m_nom in st.session_state.horaires_par_matiere[promo_selected]:
                                    del st.session_state.horaires_par_matiere[promo_selected][m_nom]
                                
                if st.button(f"🚀 Générer l'EDT pour {promo_selected}", type="primary", key=f"btn_gen_{promo_selected}"):
                    if len(lieux_sel) == 0: st.error("❌ Sélectionnez au moins un lieu.")
                    else:
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        total_promos = len(st.session_state.promotions_list)
                        generated_count = 0
                        
                        for i, p_item in enumerate(st.session_state.promotions_list):
                            time.sleep(0.2)
                            ordre = st.session_state.ordre_matieres.get(p_item, None)
                            horaires = st.session_state.horaires_par_matiere.get(p_item, None)
                            jours_m = st.session_state.jours_par_matiere.get(p_item, None)
                            
                            if st.session_state.planning_df is None:
                                st.session_state.planning_df = st.session_state.examens_df.copy()
                                
                            planning_promo = generer_planning_promo(st.session_state.planning_df, p_item, st.session_state.date_debut_val, st.session_state.jours_feries, CRENEAUX, lieux_sel, ordre, horaires, jours_m)
                            if planning_promo is not None:
                                st.session_state.planning_df = planning_promo
                                st.session_state.historique_edt[p_item] = planning_promo[planning_promo['Promotion'].astype(str).str.strip() == str(p_item).strip()].to_dict('records')
                            
                            generated_count += 1
                            progress_bar.progress(generated_count / total_promos)
                            status_text.text(f"Progression : {generated_count} / {total_promos} promotions générées")
                            
                        st.success(f"✅ Génération par sélection manuelle effectuée avec succès !")

                if st.session_state.planning_df is not None:
                    st.markdown("---")
                    st.markdown(f"#### 📝 Planning actuel de la promotion sélectionnée : {promo_selected}")
                    planning_display = st.session_state.planning_df[st.session_state.planning_df['Promotion'].astype(str).str.strip() == str(promo_selected).strip()].copy()
                    if not planning_display.empty:
                        # Respect de la disposition demandée : Enseignements, Code, Enseignants, Horaire, Jours, Lieu, Promotion
                        colonnes_ordre = ["Enseignements", "Code", "Enseignants", "Horaire", "Jours", "Lieu", "Promotion"]
                        st.dataframe(planning_display[[c for c in colonnes_ordre if c in planning_display.columns]], use_container_width=True, hide_index=True)
                    else:
                        st.info(f"Aucun examen planifié pour {promo_selected} pour le moment. Cliquez sur le bouton de génération ci-dessus.")
        else: st.warning("Aucune promotion détectée.")

    with tabs[3]:
        st.markdown('<div class="sub-header">Attribution des Surveillants (Vacataire en 2ème position)</div>', unsafe_allow_html=True)
        if st.session_state.planning_df is not None and st.session_state.enseignants_df is not None:
            if st.button("🎯 Attribuer les Surveillants", type="primary", key="btn_attrib"):
                with st.spinner("Attribution intelligente en cours..."):
                    attributions, ens_maj = attribuer_surveillants(st.session_state.planning_df, st.session_state.enseignants_df, st.session_state.nb_surv_par_lieu)
                    if attributions is not None:
                        st.session_state.surveillance_df = attributions
                        st.session_state.enseignants_df = ens_maj
                        st.success(f"✅ {len(attributions)} attributions effectuées (Vacataire en 2e position garanti)!")
                    else: st.error("❌ Erreur.")
            if st.session_state.surveillance_df is not None:
                st.markdown("---")
                attr_data = []
                for attr in st.session_state.surveillance_df:
                    d = attr.get('date', None)
                    ds = d.strftime('%d/%m/%Y') if hasattr(d, 'strftime') else str(d) if d else ''
                    jour = JOURS_FR.get(d.strftime('%A'), d.strftime('%A')) if hasattr(d, 'strftime') else ''
                    for surv in attr.get('details_surveillants', []):
                        attr_data.append({
                            'Enseignements': attr.get('matiere', ''),
                            'Code': f"CODE-{abs(hash(attr.get('matiere', ''))) % 9000 + 1000}",
                            'Enseignants': attr.get('enseignant', ''),
                            'Horaire': attr.get('creneau', ''),
                            'Jours': jour,
                            'Lieu': attr.get('lieu', ''),
                            'Promotion': attr.get('promotion', ''),
                            'Date': ds,
                            'Surveillant': surv['nom'],
                            'Qualité': surv['qualite'],
                            'Rôle': 'Chargé de matière' if surv.get('priorite') == 'Charge de matiere' else surv['qualite']
                        })
                if attr_data:
                    st.dataframe(pd.DataFrame(attr_data), use_container_width=True, hide_index=True)
        else: st.warning("Générez d'abord le planning.")

    with tabs[4]:
        st.markdown('<div class="sub-header">📅 EDT Chronologique en Grille par Promotion</div>', unsafe_allow_html=True)
        if st.session_state.surveillance_df is not None:
            promo_sel_choisie = st.selectbox("🎯 Choisir la promotion à afficher et télécharger :", st.session_state.promotions_list, key="select_promo_unique_edt")
            if promo_sel_choisie:
                st.markdown(f"#### 🎓 Promotion: **{promo_sel_choisie}**")
                attr_promo = [a for a in st.session_state.surveillance_df if str(a.get('promotion', '')).strip() == str(promo_sel_choisie).strip()]
                if attr_promo:
                    df_grille, _, _ = construire_grille_edt(attr_promo, CRENEAUX)
                    if df_grille is not None and not df_grille.empty:
                        st.dataframe(df_grille, use_container_width=True, hide_index=True)
                        c1, c2, c3 = st.columns(3)
                        with c1: st.download_button(f"⬇️ HTML - {promo_sel_choisie}", generer_html_edt(df_grille, promo_sel_choisie), f"EDT_{promo_sel_choisie}.html", "text/html", key=f"dl_html_{promo_sel_choisie}")
                        with c2: st.download_button(f"⬇️ Excel - {promo_sel_choisie}", generer_excel_edt(df_grille, promo_sel_choisie), f"EDT_{promo_sel_choisie}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"dl_xlsx_{promo_sel_choisie}")
                        with c3: st.download_button(f"⬇️ PDF - {promo_sel_choisie}", generer_pdf_edt(attr_promo, promo_sel_choisie, CRENEAUX), f"EDT_{promo_sel_choisie}.pdf", "application/pdf", key=f"dl_pdf_{promo_sel_choisie}")
                else:
                    st.info(f"Aucune attribution trouvée pour la promotion {promo_sel_choisie}. Veuillez lancer l'attribution dans l'onglet 'Attributions'.")
        else: st.warning("⚠️ Veuillez d'abord générer les attributions.")

    with tabs[5]:
        st.markdown('<div class="sub-header">📂 Répertoire des EDTs & Téléchargements (Individuel & Groupé)</div>', unsafe_allow_html=True)
        
        nb_generes = len(st.session_state.historique_edt) if 'historique_edt' in st.session_state else 0
        total_promos_attendu = len(st.session_state.promotions_list) if st.session_state.promotions_list else 23
        
        st.markdown(f"""
        <div style="background-color: #e3f2fd; padding: 15px; border-radius: 8px; text-align: center; margin-bottom: 20px;">
            <h3 style="color: #1565c0; margin: 0;">Afficheur Numérique des EDTs</h3>
            <p style="font-size: 1.5rem; font-weight: bold; color: #0d47a1; margin: 5px 0;">
                {nb_generes} / {total_promos_attendu} EDTs générés
            </p>
            <p style="font-size: 0.9rem; color: #555; margin: 0;">Nombre d'emplois du temps enregistrés dans le répertoire sur le total de vos promotions.</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.session_state.surveillance_df is not None and st.session_state.promotions_list:
            st.markdown("### 📥 Téléchargement Groupé de tous les EDTs")
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.download_button("⬇️ Télécharger le Planning Global Excel", generer_excel_colore(st.session_state.surveillance_df), "Planning_Global_Toutes_Promotions.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="btn_group_excel")
            with col_g2:
                st.download_button("⬇️ Télécharger le Planning Global PDF", generer_pdf(st.session_state.surveillance_df), "Planning_Global_Toutes_Promotions.pdf", "application/pdf", key="btn_group_pdf")
                
            st.markdown("---")
            st.markdown("### 📑 Répertoire Individuel par Promotion")
            for promo in st.session_state.promotions_list:
                attr_promo = [a for a in st.session_state.surveillance_df if str(a.get('promotion', '')).strip() == str(promo).strip()]
                if attr_promo:
                    df_g, _, _ = construire_grille_edt(attr_promo, CRENEAUX)
                    if df_g is not None and not df_g.empty:
                        with st.expander(f"🎓 Promotion : {promo}"):
                            st.dataframe(df_g, use_container_width=True, hide_index=True)
                            b_ind1, b_ind2, b_ind3 = st.columns(3)
                            with b_ind1:
                                st.download_button(f"HTML - {promo}", generer_html_edt(df_g, promo), f"EDT_{promo}.html", "text/html", key=f"rep_html_{promo}")
                            with b_ind2:
                                st.download_button(f"Excel - {promo}", generer_excel_edt(df_g, promo), f"EDT_{promo}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"rep_xlsx_{promo}")
                            with b_ind3:
                                st.download_button(f"PDF - {promo}", generer_pdf_edt(attr_promo, promo, CRENEAUX), f"EDT_{promo}.pdf", "application/pdf", key=f"rep_pdf_{promo}")
                else:
                    st.info(f"Promotion {promo} : Pas encore générée.")
        else:
            st.warning("⚠️ Veuillez d'abord générer les plannings et les attributions pour alimenter le répertoire.")

    with tabs[6]:
        st.markdown('<div class="sub-header">Export Global Chronologique</div>', unsafe_allow_html=True)
        if st.session_state.surveillance_df is not None:
            attributions = st.session_state.surveillance_df
            col1, col2, col3 = st.columns(3)
            with col1: st.download_button("⬇️ Télécharger HTML", generer_tableau_html(attributions), "planning_surveillances.html", "text/html", key="dl_gh")
            with col2: st.download_button("⬇️ Télécharger Excel", generer_excel_colore(attributions), "planning_surveillances.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.spreadsheet", key="dl_gx")
            with col3: st.download_button("⬇️ Télécharger PDF", generer_pdf(attributions), "planning_surveillances.pdf", "application/pdf", key="dl_gp")
            st.markdown("---")
            st.markdown(generer_tableau_html(attributions), unsafe_allow_html=True)
        else: st.warning("Aucune attribution à exporter.")

if __name__ == "__main__":
    main()
