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

SALLES = [f"S{i:02d}" for i in range(1, 19)]  # S01 à S18
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
        'historique_edt': {}, 'planning_manuel': {}  # NOUVEAU : stockage du planning manuel
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
                    examens_data.append({'Enseignements': nom_cours, 'Code': f"CODE-{abs(hash(nom_cours)) % 9000 + 1000}", 
                        'Promotion': str(row.get('promotion', '')).strip(),
                        'Enseignants': str(row.get('nom', '')).strip(), 'qualite_ens': row.get('qualite', 'Permanent'),
                        'date': None, 'Horaire': None, 'Jours': None, 'Lieu': None, 'ordre': 999})
        df_exam = pd.DataFrame(examens_data)
        df_exam = df_exam.drop_duplicates(subset=['Enseignements', 'Promotion', 'Enseignants']).copy()
        df_exam = df_exam.sort_values('Enseignants').drop_duplicates(subset=['Enseignements', 'Promotion'], keep='first').copy()
        df_ens['nb_surveillance'] = 0
        df_ens['surveillance_attribuee'] = 0
        promotions = sorted(df_exam['Promotion'].dropna().astype(str).str.strip().unique().tolist())
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

def generer_dates_travail(date_debut, nb_jours=30, jours_feries=None):
    """Génère une liste de dates de travail disponibles"""
    if jours_feries is None:
        jours_feries = []
    dates = []
    date_courante = date_debut
    while len(dates) < nb_jours:
        if est_jour_travaille(date_courante, jours_feries):
            dates.append(date_courante)
        date_courante += timedelta(days=1)
    return dates

def creer_planning_manuel_df(examens_df, promotion, dates_disponibles):
    """Crée un DataFrame pour l'édition manuelle du planning"""
    promo_df = examens_df[examens_df['Promotion'].astype(str).str.strip() == str(promotion).strip()].copy()
    if promo_df.empty:
        return None
    
    planning_data = []
    for _, row in promo_df.iterrows():
        planning_data.append({
            'Enseignements': row.get('Enseignements', ''),
            'Enseignants': row.get('Enseignants', ''),
            'Code': row.get('Code', ''),
            'Date': None,  # À remplir
            'Horaire': None,  # À remplir
            'Lieu': None,  # À remplir
        })
    
    return pd.DataFrame(planning_data)

def appliquer_planning_manuel(examens_df, promotion, planning_edited):
    """Applique le planning édité manuellement aux examens"""
    if planning_edited is None or planning_edited.empty:
        st.error("❌ Le planning est vide")
        return None
    
    # Vérifier que toutes les dates/horaires/lieux sont remplis
    if planning_edited.isnull().any().any():
        st.warning("⚠️ Veuillez remplir TOUS les champs (Date, Horaire, Lieu)")
        return None
    
    result = examens_df.copy()
    
    for _, row in planning_edited.iterrows():
        matiere = row.get('Enseignements', '')
        date_exam = row.get('Date', None)
        horaire = row.get('Horaire', None)
        lieu = row.get('Lieu', None)
        
        # Convertir la date si nécessaire
        if isinstance(date_exam, str):
            try:
                date_exam = datetime.strptime(date_exam, "%Y-%m-%d").date()
            except:
                date_exam = datetime.strptime(date_exam, "%d/%m/%Y").date()
        
        # Obtenir le jour de la semaine
        jour = JOURS_FR.get(date_exam.strftime("%A"), date_exam.strftime("%A"))
        
        # Appliquer à toutes les lignes correspondant à cette matière/promo
        mask = (result['Promotion'].astype(str).str.strip() == str(promotion).strip()) & \
               (result['Enseignements'] == matiere)
        
        result.loc[mask, 'date'] = date_exam
        result.loc[mask, 'Horaire'] = horaire
        result.loc[mask, 'Jours'] = jour
        result.loc[mask, 'Lieu'] = lieu
    
    return result

def construire_grille_edt(attributions, creneaux_liste):
    if not attributions:
        return None, None, None
    
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
        
        grille[cle_jour][creneau].append(attr)
    
    creneaux_utilises = creneaux_liste if creneaux_liste else CRENEAUX
    data = []
    
    for creneau in creneaux_utilises:
        row = {'Creneau': creneau}
        for jour in jours_ordre:
            exams = grille.get(jour, {}).get(creneau, [])
            if exams:
                cellules = []
                for ex in exams:
                    cell_text = (
                        f"📖 {ex.get('matiere', '')}\n"
                        f"👤 {ex.get('enseignant', '')}\n"
                        f"🏫 {ex.get('lieu', '')}\n"
                        f"👮 Surveillance:\n" + ", ".join([s['nom'] for s in ex.get('details_surveillants', [])])
                    )
                    cellules.append(cell_text)
                row[jour] = "\n---\n".join(cellules)
            else:
                row[jour] = ""
        data.append(row)
    
    df_grille = pd.DataFrame(data)
    return df_grille, jours_ordre, grille

def generer_html_edt(df_grille, promotion):
    jours_cols = [c for c in df_grille.columns if c != 'Creneau']
    html = f"""
    <style>
        .edt-table {{ border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 11px; }}
        .edt-table th {{ background-color: #1565C0; color: white; padding: 12px; text-align: center; border: 2px solid #0D47A1; font-size: 12px; font-weight: bold; }}
        .edt-table td {{ padding: 12px; border: 1px solid #90CAF9; vertical-align: top; min-width: 220px; text-align: center; }}
        .creneau-cell {{ background-color: #E3F2FD; font-weight: bold; text-align: center; font-size: 12px; width: 120px; }}
        .exam-cell {{ background-color: #FFF8E1; text-align: center; }}
    </style>
    <h2 style="color:#1565C0; text-align:center;">📅 EDT EXAMENS - Promotion {promotion}</h2>
    <table class="edt-table">
    <tr><th>Créneau</th>"""
    for jour in jours_cols:
        html += f"<th>{jour.replace(chr(10), '<br>')}</th>"
    html += "</tr>"
    
    for _, row in df_grille.iterrows():
        html += f"<tr><td class='creneau-cell'>{row['Creneau']}</td>"
        for jour in jours_cols:
            val = row.get(jour, '')
            html += f"<td class='exam-cell'>{val.replace(chr(10), '<br>')}</td>"
        html += "</tr>"
    html += "</table>"
    return html

def generer_excel_edt(df_grille, promotion):
    wb = Workbook()
    ws = wb.active
    ws.title = f"EDT {promotion}"
    
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    
    headers = ['Creneau'] + [c for c in df_grille.columns if c != 'Creneau']
    
    for c_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    for r_idx, row in df_grille.iterrows():
        for c_idx, col_name in enumerate(headers, 1):
            val = row.get(col_name, '')
            cell = ws.cell(row=r_idx + 2, column=c_idx, value=val)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    ws.column_dimensions['A'].width = 18
    for col_idx in range(2, len(headers) + 1):
        ws.column_dimensions[f'{chr(64+col_idx)}'.replace('64', 'A')].width = 40
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generer_pdf_edt(attributions, promotion, creneaux_liste):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=1*cm, leftMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    elements = []
    styles = getSampleStyleSheet()
    
    elements.append(Paragraph(f"📅 EDT EXAMENS - {promotion}", styles['Heading1']))
    elements.append(Spacer(1, 0.3*cm))
    
    df_grille, _, _ = construire_grille_edt(attributions, creneaux_liste)
    if df_grille is None:
        elements.append(Paragraph("Aucune donnée", styles['Normal']))
    else:
        table_data = [['Créneau'] + [c.replace('\n', ' ') for c in df_grille.columns if c != 'Creneau']]
        for _, row in df_grille.iterrows():
            table_data.append([row.get(col, '') for col in df_grille.columns])
        
        table = Table(table_data, repeatRows=1)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#90CAF9')),
        ]))
        elements.append(table)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer

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
    html = """<table border='1' style='border-collapse:collapse;width:100%;'>
    <tr style='background-color:#1565C0;color:white;'><th>Créneau</th>"""
    jours = sorted(planning_par_jour.keys())
    for jour in jours: html += f"<th>{jour}</th>"
    html += "</tr>"
    for creneau in CRENEAUX:
        html += f"<tr><td style='background-color:#E3F2FD;font-weight:bold;'>{creneau}</td>"
        for jour in jours:
            html += "<td>"
            if creneau in planning_par_jour.get(jour, {}):
                for examen in planning_par_jour[jour][creneau]:
                    html += f"<div style='background-color:#FFF8E1;padding:5px;margin:2px;border-left:3px solid #FFA000;'><strong>{examen.get('matiere', '')}</strong><br>{examen.get('lieu', '')}</div>"
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
        surv_str = ", ".join([f"{s['nom']}" for s in attr.get('details_surveillants', [])])
        data.append({
            'Date': date_str, 'Jour': jour, 'Créneau': attr.get('creneau', ''), 
            'Matière': attr.get('matiere', ''), 'Enseignant': attr.get('enseignant', ''),
            'Promotion': attr.get('promotion', ''), 'Lieu': attr.get('lieu', ''), 
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
    
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generer_pdf(attributions):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=1*cm, leftMargin=1*cm, topMargin=1*cm, bottomMargin=1*cm)
    elements = []
    styles = getSampleStyleSheet()
    
    elements.append(Paragraph("📋 PLANNING DES SURVEILLANCES", styles['Heading1']))
    elements.append(Spacer(1, 0.3*cm))
    
    attr_sorted = sorted(attributions, key=lambda x: x.get('date', datetime.max))
    table_data = [['Date', 'Jour', 'Créneau', 'Matière', 'Enseignant', 'Promotion', 'Lieu', 'Surveillants']]
    
    for attr in attr_sorted:
        date_val = attr.get('date', None)
        if hasattr(date_val, 'strftime'):
            date_str = date_val.strftime('%d/%m/%Y')
            jour = JOURS_FR.get(date_val.strftime('%A'), date_val.strftime('%A'))
        else:
            date_str = str(date_val)
            jour = ''
        surv_str = ", ".join([f"{s['nom']}" for s in attr.get('details_surveillants', [])])
        table_data.append([date_str, jour, attr.get('creneau', ''), attr.get('matiere', ''),
            attr.get('enseignant', ''), attr.get('promotion', ''), attr.get('lieu', ''), surv_str])
    
    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#90CAF9')),
    ]))
    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

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

        attributions.append({'date': date_examen, 'creneau': creneau_examen, 'matiere': matiere_examen,
            'promotion': promotion_examen, 'lieu': lieu_examen, 'enseignant': enseignant_matiere,
            'surveillants': [s['nom'] for s in liste_surveillants], 'details_surveillants': liste_surveillants})
    
    return attributions, surveillants

def main():
    init_session_state()
    st.markdown('<div class="main-header">📋 Gestion des EDTs - S2 2026</div>', unsafe_allow_html=True)

    if not st.session_state.data_loaded:
        with st.spinner("Chargement du fichier source..."):
            result, error = charger_fichier_source_auto()
            if result is not None:
                st.session_state.enseignants_df = result['enseignants']
                st.session_state.examens_df = result['examens']
                st.session_state.promotions_list = result['promotions']
                st.session_state.data_loaded = True
                st.success(f"✅ Chargé: {len(result['enseignants'])} ens. | {len(result['examens'])} cours | Promos: {', '.join(result['promotions'])}")
            else:
                st.error(f"❌ {error}")

    with st.sidebar:
        st.markdown("## ⚙️ Configuration")
        st.session_state.nb_surv_permanent = st.number_input("Permanents", 1, 10, st.session_state.nb_surv_permanent)
        st.session_state.nb_surv_vacataire = st.number_input("Vacataires", 1, 10, st.session_state.nb_surv_vacataire)
        st.session_state.nb_surv_par_lieu = st.number_input("Surv. par lieu", 1, 5, st.session_state.nb_surv_par_lieu)
        st.markdown("---")
        st.markdown("### 📅 Dates")
        st.session_state.date_debut_val = st.date_input("Début", st.session_state.date_debut_val)
        jf = st.text_area("Jours fériés (JJ/MM/AAAA)", placeholder="11/11/2026, 25/12/2026...")
        if jf:
            try: st.session_state.jours_feries = [datetime.strptime(d.strip(), "%d/%m/%Y").date() for d in jf.split(",") if d.strip()]
            except: st.warning("Format invalide")

    tabs = st.tabs(["🏠 Accueil", "👥 Enseignants", "🗓️ Planning Manuel", "🎯 Attributions", "📅 EDT Grille", "📊 Export"])

    with tabs[0]:
        st.markdown("""
        <div class="info-box">
            <h3>✨ Plateforme de Gestion des EDTs</h3>
            <p><strong>S2 - 2026</strong></p>
            <ul style="text-align:left;">
                <li>✅ Sélection MANUELLE des dates et horaires</li>
                <li>✅ Pas de répartition automatique</li>
                <li>✅ Possible d'avoir 2 examens le même jour à des créneaux différents</li>
                <li>✅ Attribution intelligente des surveillants</li>
                <li>✅ Export HTML, Excel, PDF</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        if st.session_state.data_loaded:
            c1, c2, c3 = st.columns(3)
            with c1: st.metric("Enseignants", len(st.session_state.enseignants_df))
            with c2: st.metric("Cours", len(st.session_state.examens_df))
            with c3: st.metric("Promotions", len(st.session_state.promotions_list))

    with tabs[1]:
        st.markdown('<div class="sub-header">Gestion des Enseignants</div>', unsafe_allow_html=True)
        if st.session_state.data_loaded:
            df_ens = st.session_state.enseignants_df
            # Compter les noms UNIQUES (pas les doublons)
            all_ens_uniq = df_ens['nom'].unique().tolist()
            perm_uniq = df_ens[df_ens['qualite'] == 'Permanent']['nom'].unique()
            vac_uniq = df_ens[df_ens['qualite'] == 'Vacataire']['nom'].unique()
            
            st.markdown(f"#### 📊 Résumé")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("👔 Permanents (uniques)", len(perm_uniq))
            with c2:
                st.metric("📝 Vacataires (uniques)", len(vac_uniq))
            with c3:
                st.metric("👥 Total (uniques)", len(all_ens_uniq))
            
            st.markdown("---")
            exclus = st.multiselect("🚫 Exclure de la surveillance", sorted(all_ens_uniq), default=st.session_state.exclus_manuels)
            st.session_state.exclus_manuels = exclus
            
            if exclus:
                st.warning(f"❌ Exclus: {', '.join(exclus)}")

    with tabs[2]:
        st.markdown('<div class="sub-header">🗓️ Planning Manuel - Sélection Date + Horaire + Lieu</div>', unsafe_allow_html=True)
        st.markdown("<div class='info-box'><b>✨ NOUVEAU</b> : Choisissez MANUELLEMENT la date (calendrier), l'horaire (3 créneaux) et le lieu (Salles ou Amphithéâtres). Aucune répartition automatique!</div>", unsafe_allow_html=True)
        
        if st.session_state.data_loaded and st.session_state.promotions_list:
            promo_selected = st.selectbox("📚 Promotion", st.session_state.promotions_list, key="promo_planning")
            
            if promo_selected:
                # Créer le DataFrame pour édition manuelle
                promo_df = st.session_state.examens_df[st.session_state.examens_df['Promotion'].astype(str).str.strip() == str(promo_selected).strip()].copy()
                
                if promo_df is not None and not promo_df.empty:
                    st.markdown(f"#### Configurer les examens de {promo_selected}")
                    st.markdown("**Choisissez pour chaque examen: 📅 Date (Calendrier) | ⏰ Horaire (3 créneaux) | 📍 Lieu (Salle OU Amphi)**")
                    st.markdown("---")
                    
                    # Stockage des sélections
                    planning_selections = {}
                    
                    # Créer une ligne par examen avec inputs personnalisés
                    for idx, (_, exam) in enumerate(promo_df.iterrows()):
                        enseignement = exam.get('Enseignements', '')
                        enseignant = exam.get('Enseignants', '')
                        code = exam.get('Code', '')
                        
                        # Créer un expander pour chaque examen
                        with st.expander(f"📚 {enseignement} - {enseignant}", expanded=(idx==0)):
                            col_info, col_config = st.columns([2, 3])
                            
                            # Infos (lecture seule)
                            with col_info:
                                st.markdown(f"**Enseignement**: {enseignement}")
                                st.markdown(f"**Enseignant**: {enseignant}")
                                st.markdown(f"**Code**: {code}")
                            
                            # Configuration (inputs)
                            with col_config:
                                st.markdown("**Configuration**")
                                
                                # Calendrier pour la date
                                date_exam = st.date_input(
                                    "📅 Sélectionner la date",
                                    value=st.session_state.date_debut_val,
                                    min_value=st.session_state.date_debut_val,
                                    max_value=st.session_state.date_debut_val + timedelta(days=60),
                                    key=f"date_{promo_selected}_{idx}"
                                )
                                
                                # Horaires (3 créneaux)
                                horaire = st.selectbox(
                                    "⏰ Sélectionner l'horaire",
                                    CRENEAUX,
                                    key=f"horaire_{promo_selected}_{idx}"
                                )
                                
                                # Lieu: Deux colonnes (Salles + Amphis)
                                lieu_col1, lieu_col2 = st.columns(2)
                                
                                with lieu_col1:
                                    salle = st.selectbox(
                                        "🏫 Salle (S01-S18)",
                                        [""] + SALLES,
                                        key=f"salle_{promo_selected}_{idx}"
                                    )
                                
                                with lieu_col2:
                                    amphi = st.selectbox(
                                        "🎓 Amphithéâtre (A01-A12)",
                                        [""] + AMPHIS,
                                        key=f"amphi_{promo_selected}_{idx}"
                                    )
                                
                                # Vérifier qu'au moins un lieu est sélectionné
                                if salle == "" and amphi == "":
                                    st.warning("⚠️ Sélectionnez au moins un lieu (Salle OU Amphi)")
                                    lieu_final = None
                                elif salle != "" and amphi != "":
                                    st.error("❌ Choisissez SOIT une Salle SOIT un Amphi, pas les deux!")
                                    lieu_final = None
                                else:
                                    lieu_final = salle if salle != "" else amphi
                                    st.success(f"✅ Lieu sélectionné: {lieu_final}")
                                
                                # Stocker les sélections
                                planning_selections[enseignement] = {
                                    'date': date_exam,
                                    'horaire': horaire,
                                    'lieu': lieu_final
                                }
                    
                    st.markdown("---")
                    
                    if st.button("✅ Appliquer et Générer", type="primary", key="btn_apply_planning", use_container_width=True):
                        # Vérifier que toutes les sélections sont valides
                        invalides = [m for m, sel in planning_selections.items() if sel['lieu'] is None]
                        
                        if invalides:
                            st.error(f"❌ Veuillez sélectionner un lieu pour: {', '.join(invalides)}")
                        else:
                            # Créer le planning édité
                            planning_edited = []
                            for enseignement, sels in planning_selections.items():
                                planning_edited.append({
                                    'Enseignements': enseignement,
                                    'Date': sels['date'],
                                    'Horaire': sels['horaire'],
                                    'Lieu': sels['lieu']
                                })
                            
                            planning_edited_df = pd.DataFrame(planning_edited)
                            
                            # Appliquer le planning
                            planning_updated = appliquer_planning_manuel(st.session_state.examens_df, promo_selected, planning_edited_df)
                            
                            if planning_updated is not None:
                                st.session_state.planning_df = planning_updated
                                st.session_state.planning_manuel[promo_selected] = planning_edited_df.to_dict('records')
                                st.success(f"✅ Planning de {promo_selected} configuré avec succès!")
                                
                                # Afficher le planning appliqué
                                st.markdown("---")
                                st.markdown("#### 📋 Planning Appliqué")
                                result_df = planning_updated[planning_updated['Promotion'] == promo_selected][['Enseignements', 'Enseignants', 'date', 'Horaire', 'Jours', 'Lieu']].copy()
                                result_df.columns = ['Enseignement', 'Enseignant', 'Date', 'Horaire', 'Jour', 'Lieu']
                                st.dataframe(result_df, use_container_width=True, hide_index=True)
                else:
                    st.warning(f"Aucun examen pour {promo_selected}")

    with tabs[3]:
        st.markdown('<div class="sub-header">🎯 Attribution des Surveillants</div>', unsafe_allow_html=True)
        if st.session_state.planning_df is not None and st.session_state.enseignants_df is not None:
            if st.button("🎯 Attribuer les Surveillants", type="primary"):
                with st.spinner("Attribution en cours..."):
                    attributions, ens_maj = attribuer_surveillants(st.session_state.planning_df, st.session_state.enseignants_df, st.session_state.nb_surv_par_lieu)
                    if attributions is not None:
                        st.session_state.surveillance_df = attributions
                        st.session_state.enseignants_df = ens_maj
                        st.success(f"✅ {len(attributions)} attributions effectuées!")
            
            if st.session_state.surveillance_df is not None:
                st.markdown("---")
                attr_data = []
                for attr in st.session_state.surveillance_df:
                    for surv in attr.get('details_surveillants', []):
                        d = attr.get('date', None)
                        ds = d.strftime('%d/%m/%Y') if hasattr(d, 'strftime') else str(d) if d else ''
                        jour = JOURS_FR.get(d.strftime('%A'), d.strftime('%A')) if hasattr(d, 'strftime') else ''
                        attr_data.append({
                            'Enseignement': attr.get('matiere', ''),
                            'Enseignant': attr.get('enseignant', ''),
                            'Horaire': attr.get('creneau', ''),
                            'Jour': jour,
                            'Lieu': attr.get('lieu', ''),
                            'Promotion': attr.get('promotion', ''),
                            'Date': ds,
                            'Surveillant': surv['nom'],
                            'Qualité': surv['qualite']
                        })
                if attr_data:
                    st.dataframe(pd.DataFrame(attr_data), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ Configurez d'abord le planning manuel")

    with tabs[4]:
        st.markdown('<div class="sub-header">📅 EDT Chronologique en Grille</div>', unsafe_allow_html=True)
        if st.session_state.surveillance_df is not None:
            promo_sel = st.selectbox("🎓 Promotion", st.session_state.promotions_list, key="sel_promo_edt")
            attr_promo = [a for a in st.session_state.surveillance_df if str(a.get('promotion', '')).strip() == str(promo_sel).strip()]
            
            if attr_promo:
                df_grille, _, _ = construire_grille_edt(attr_promo, CRENEAUX)
                if df_grille is not None:
                    st.dataframe(df_grille, use_container_width=True, hide_index=True)
                    c1, c2, c3 = st.columns(3)
                    with c1: st.download_button("⬇️ HTML", generer_html_edt(df_grille, promo_sel), f"EDT_{promo_sel}.html", "text/html", key="dl_html")
                    with c2: st.download_button("⬇️ Excel", generer_excel_edt(df_grille, promo_sel), f"EDT_{promo_sel}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")
                    with c3: st.download_button("⬇️ PDF", generer_pdf_edt(attr_promo, promo_sel, CRENEAUX), f"EDT_{promo_sel}.pdf", "application/pdf", key="dl_pdf")
            else:
                st.info("Aucune attribution pour cette promotion")
        else:
            st.warning("⚠️ Effectuez d'abord l'attribution")

    with tabs[5]:
        st.markdown('<div class="sub-header">📊 Export Global</div>', unsafe_allow_html=True)
        if st.session_state.surveillance_df is not None:
            attributions = st.session_state.surveillance_df
            c1, c2, c3 = st.columns(3)
            with c1: st.download_button("⬇️ HTML", generer_tableau_html(attributions), "planning.html", "text/html", key="dl_gh")
            with c2: st.download_button("⬇️ Excel", generer_excel_colore(attributions), "planning.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_gx")
            with c3: st.download_button("⬇️ PDF", generer_pdf(attributions), "planning.pdf", "application/pdf", key="dl_gp")
            st.markdown("---")
            st.markdown(generer_tableau_html(attributions), unsafe_allow_html=True)
        else:
            st.warning("Aucune attribution à exporter")

if __name__ == "__main__":
    main()
