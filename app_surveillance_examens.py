
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import base64
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import json
import os

# ============================================================
# CONFIGURATION DE LA PAGE
# ============================================================
st.set_page_config(
    page_title="Gestion des Surveillances d'Examens",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# STYLES CSS PERSONNALISÉS
# ============================================================
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f4e79;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #e3f2fd 0%, #bbdefb 100%);
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #1565c0;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
    }
    .info-box {
        background-color: #e3f2fd;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #1565c0;
        margin: 0.5rem 0;
    }
    .warning-box {
        background-color: #fff3e0;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #f57c00;
        margin: 0.5rem 0;
    }
    .success-box {
        background-color: #e8f5e9;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #2e7d32;
        margin: 0.5rem 0;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #f5f5f5;
        border-radius: 8px 8px 0 0;
        padding: 10px 20px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1565c0 !important;
        color: white !important;
    }
    .dataframe-container {
        border: 1px solid #ddd;
        border-radius: 8px;
        overflow: hidden;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# CONSTANTES ET CONFIGURATIONS
# ============================================================
SALLES = [f"S{i:02d}" for i in range(1, 19)]  # S01 à S18
AMPHIS = [f"A{i:02d}" for i in range(1, 13)]  # A01 à A12
LIEUX = SALLES + AMPHIS

CRENEAUX_HORAIRES = [
    "08h30 – 10h30",
    "11h00 – 13h00",
    "13h30 – 15h30"
]

JOURS_SEMAINE = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]

COULEURS = {
    "Lundi": "#E3F2FD",
    "Mardi": "#E8F5E9",
    "Mercredi": "#FFF3E0",
    "Jeudi": "#F3E5F5",
    "Dimanche": "#E0F2F1",
    "header": "#1565C0",
    "subheader": "#1976D2",
    "border": "#90CAF9"
}

# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

def init_session_state():
    """Initialise les variables de session"""
    defaults = {
        'enseignants_df': None,
        'examens_df': None,
        'planning_df': None,
        'surveillance_df': None,
        'nb_surv_permanent': 3,
        'nb_surv_vacataire': 2,
        'nb_surv_autre': 1,
        'nb_surv_par_lieu': 2,
        'exclus_permanents': [],
        'exclus_vacataires': [],
        'date_debut': datetime(2026, 11, 1),
        'jours_feries': [],
        'promotions': {},
        'config_creneaux': {},
        'surveillants_attribues': {},
        'matiere_modifiee': {}
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def est_jour_travaille(date_obj, jours_feries):
    """Vérifie si un jour est travaillable (exclut vendredi, samedi, dimanche et jours fériés)"""
    jour_semaine = date_obj.strftime("%A")
    # Conversion français
    jours_fr = {
        "Monday": "Lundi", "Tuesday": "Mardi", "Wednesday": "Mercredi",
        "Thursday": "Jeudi", "Friday": "Vendredi", "Saturday": "Samedi", "Sunday": "Dimanche"
    }
    jour_fr = jours_fr.get(jour_semaine, jour_semaine)

    if jour_fr in ["Vendredi", "Samedi"]:
        return False
    if date_obj in jours_feries:
        return False
    return True

def prochain_jour_travaille(date_debut, jours_feries):
    """Trouve le prochain jour travaillable"""
    date_courante = date_debut
    while not est_jour_travaille(date_courante, jours_feries):
        date_courante += timedelta(days=1)
    return date_courante

def charger_fichier_enseignants(uploaded_file):
    """Charge et traite le fichier des enseignants"""
    try:
        df = pd.read_excel(uploaded_file)
        # Normalisation des noms de colonnes
        df.columns = [col.strip().lower().replace(' ', '_') for col in df.columns]

        # Détection automatique des colonnes
        col_mapping = {}
        for col in df.columns:
            if any(x in col for x in ['nom', 'name', 'enseignant']):
                col_mapping['nom'] = col
            elif any(x in col for x in ['qualite', 'quality', 'type', 'statut', 'grade']):
                col_mapping['qualite'] = col
            elif any(x in col for x in ['matiere', 'course', 'module', 'discipline']):
                col_mapping['matiere'] = col
            elif any(x in col for x in ['promotion', 'niveau', 'annee', 'class']):
                col_mapping['promotion'] = col

        # Renommage standard
        if 'nom' in col_mapping:
            df = df.rename(columns={col_mapping['nom']: 'nom'})
        if 'qualite' in col_mapping:
            df = df.rename(columns={col_mapping['qualite']: 'qualite'})
        else:
            # Si pas de colonne qualité, demander à l'utilisateur de la créer
            st.warning("Colonne 'qualite' non détectée. Veuillez vérifier votre fichier.")
            df['qualite'] = 'Permanent'

        if 'matiere' not in df.columns:
            df['matiere'] = ''
        if 'promotion' not in df.columns:
            df['promotion'] = ''

        # Normalisation de la qualité
        df['qualite'] = df['qualite'].astype(str).str.strip().str.lower()
        df['qualite'] = df['qualite'].map({
            'permanent': 'Permanent',
            'vacataire': 'Vacataire',
            'contractuel': 'Contractuel',
            'autre': 'Autre'
        }).fillna('Permanent')

        df['nb_surveillance'] = 0
        df['surveillance_attribuee'] = 0

        return df
    except Exception as e:
        st.error(f"Erreur lors du chargement du fichier: {str(e)}")
        return None

def charger_fichier_examens(uploaded_file):
    """Charge et traite le fichier des examens"""
    try:
        df = pd.read_excel(uploaded_file, sheet_name=None)

        # Si plusieurs feuilles, on les combine
        if isinstance(df, dict):
            all_data = []
            for sheet_name, sheet_df in df.items():
                sheet_df['feuille'] = sheet_name
                all_data.append(sheet_df)
            df = pd.concat(all_data, ignore_index=True)

        # Normalisation
        df.columns = [col.strip().lower().replace(' ', '_') for col in df.columns]

        # Détection des colonnes
        col_mapping = {}
        for col in df.columns:
            if any(x in col for x in ['matiere', 'course', 'module', 'discipline', 'ue']):
                col_mapping['matiere'] = col
            elif any(x in col for x in ['promotion', 'niveau', 'annee', 'class', 'niveau']):
                col_mapping['promotion'] = col
            elif any(x in col for x in ['enseignant', 'prof', 'charge', 'responsable']):
                col_mapping['enseignant'] = col
            elif any(x in col for x in ['duree', 'duration', 'temps']):
                col_mapping['duree'] = col

        for std_name, orig_name in col_mapping.items():
            if std_name != orig_name:
                df = df.rename(columns={orig_name: std_name})

        # Colonnes par défaut si manquantes
        for col in ['matiere', 'promotion', 'enseignant', 'duree']:
            if col not in df.columns:
                df[col] = ''

        df['date'] = None
        df['creneau'] = None
        df['lieu'] = None
        df['nb_surveillants'] = 2

        return df
    except Exception as e:
        st.error(f"Erreur lors du chargement du fichier examens: {str(e)}")
        return None

def generer_planning(examens_df, date_debut, jours_feries, config_creneaux):
    """Génère le planning automatique des examens"""
    if examens_df is None or examens_df.empty:
        return None

    planning = examens_df.copy()
    date_courante = date_debut

    # Regrouper par promotion
    promotions = planning['promotion'].unique()

    for promo in promotions:
        if pd.isna(promo) or promo == '':
            continue

        promo_data = planning[planning['promotion'] == promo].copy()

        # Configuration par défaut pour la promotion
        if promo not in config_creneaux:
            config_creneaux[promo] = {
                'creneaux': CRENEAUX_HORAIRES.copy(),
                'lieux': ['S01', 'S02'],
                'duree_jours': len(promo_data)
            }

        config = config_creneaux[promo]
        nb_creneaux = len(config['creneaux'])

        idx = 0
        for i, row in promo_data.iterrows():
            # Avancer jusqu'au prochain jour travaillable
            while not est_jour_travaille(date_courante, jours_feries):
                date_courante += timedelta(days=1)

            # Attribution créneau et lieu
            creneau_idx = idx % nb_creneaux
            lieu_idx = idx % len(config['lieux'])

            planning.at[i, 'date'] = date_courante
            planning.at[i, 'creneau'] = config['creneaux'][creneau_idx]
            planning.at[i, 'lieu'] = config['lieux'][lieu_idx]

            idx += 1
            if creneau_idx == nb_creneaux - 1:
                date_courante += timedelta(days=1)

    return planning

def attribuer_surveillants(planning_df, enseignants_df, nb_par_lieu=2):
    """Attribue les surveillants aux examens"""
    if planning_df is None or enseignants_df is None:
        return None

    surveillants = enseignants_df.copy()

    # Réinitialiser les compteurs
    surveillants['surveillance_attribuee'] = 0

    # Séparer les catégories
    permanents = surveillants[
        (surveillants['qualite'] == 'Permanent') & 
        (~surveillants['nom'].isin(st.session_state.exclus_permanents))
    ].copy()

    vacataires = surveillants[
        (surveignants['qualite'] == 'Vacataire') & 
        (~surveillants['nom'].isin(st.session_state.exclus_vacataires))
    ].copy()

    autres = surveillants[
        (~surveillants['qualite'].isin(['Permanent', 'Vacataire']))
    ].copy()

    # Trier par nombre de surveillance déjà attribuée (équilibrage)
    permanents = permanents.sort_values('surveillance_attribuee')
    vacataires = vacataires.sort_values('surveillance_attribuee')

    attributions = []

    for idx, examen in planning_df.iterrows():
        date_examen = examen['date']
        creneau_examen = examen['creneau']
        matiere_examen = examen.get('matiere', '')
        enseignant_matiere = examen.get('enseignant', '')
        lieu_examen = examen.get('lieu', 'S01')

        # Trouver les surveillants déjà occupés à ce créneau
        surveillants_occupes = set()
        for attr in attributions:
            if attr['date'] == date_examen and attr['creneau'] == creneau_examen:
                surveillants_occupes.update(attr['surveillants'])

        liste_surveillants = []

        # 1. Enseignant chargé de la matière (prioritaire)
        if enseignant_matiere and str(enseignant_matiere) != 'nan':
            ens_info = surveillants[surveillants['nom'] == enseignant_matiere]
            if not ens_info.empty:
                nom_ens = ens_info.iloc[0]['nom']
                qualite_ens = ens_info.iloc[0]['qualite']
                if nom_ens not in surveillants_occupes:
                    # Vérifier le quota
                    quota = st.session_state.get(f'nb_surv_{qualite_ens.lower()}', 3)
                    if ens_info.iloc[0]['surveillance_attribuee'] < quota:
                        liste_surveillants.append({
                            'nom': nom_ens,
                            'qualite': qualite_ens,
                            'priorite': 'Chargé de matière'
                        })
                        surveillants_occupes.add(nom_ens)
                        surveillants.loc[surveillants['nom'] == nom_ens, 'surveillance_attribuee'] += 1

        # 2. Compléter avec les permanents (toujours en tête)
        for _, perm in permanents.iterrows():
            if len(liste_surveillants) >= nb_par_lieu:
                break
            if perm['nom'] in surveillants_occupes:
                continue
            quota_perm = st.session_state.nb_surv_permanent
            if perm['surveillance_attribuee'] < quota_perm:
                liste_surveillants.append({
                    'nom': perm['nom'],
                    'qualite': 'Permanent',
                    'priorite': 'Permanent'
                })
                surveillants_occupes.add(perm['nom'])
                surveillants.loc[surveillants['nom'] == perm['nom'], 'surveillance_attribuee'] += 1
                permanents.loc[permanents['nom'] == perm['nom'], 'surveillance_attribuee'] += 1

        # 3. Compléter avec les vacataires
        for _, vac in vacataires.iterrows():
            if len(liste_surveillants) >= nb_par_lieu:
                break
            if vac['nom'] in surveillants_occupes:
                continue
            quota_vac = st.session_state.nb_surv_vacataire
            if vac['surveillance_attribuee'] < quota_vac:
                liste_surveillants.append({
                    'nom': vac['nom'],
                    'qualite': 'Vacataire',
                    'priorite': 'Vacataire'
                })
                surveillants_occupes.add(vac['nom'])
                surveillants.loc[surveillants['nom'] == vac['nom'], 'surveillance_attribuee'] += 1
                vacataires.loc[vacataires['nom'] == vac['nom'], 'surveillance_attribuee'] += 1

        # 4. Si toujours pas assez, autres
        for _, aut in autres.iterrows():
            if len(liste_surveillants) >= nb_par_lieu:
                break
            if aut['nom'] in surveillants_occupes:
                continue
            quota_aut = st.session_state.nb_surv_autre
            if aut['surveillance_attribuee'] < quota_aut:
                liste_surveillants.append({
                    'nom': aut['nom'],
                    'qualite': aut['qualite'],
                    'priorite': 'Autre'
                })
                surveillants_occupes.add(aut['nom'])
                surveillants.loc[surveillants['nom'] == aut['nom'], 'surveillance_attribuee'] += 1

        attributions.append({
            'date': date_examen,
            'creneau': creneau_examen,
            'matiere': matiere_examen,
            'promotion': examen.get('promotion', ''),
            'lieu': lieu_examen,
            'surveillants': [s['nom'] for s in liste_surveillants],
            'details_surveillants': liste_surveillants
        })

    return attributions, surveillants

def generer_tableau_html(attributions):
    """Génère un tableau HTML coloré des surveillances"""
    if not attributions:
        return "<p>Aucune attribution à afficher</p>"

    # Regrouper par date et créneau
    planning_par_jour = {}
    for attr in attributions:
        date_str = attr['date'].strftime('%d/%m/%Y') if hasattr(attr['date'], 'strftime') else str(attr['date'])
        jour_semaine = attr['date'].strftime('%A') if hasattr(attr['date'], 'strftime') else ''
        jours_fr = {"Monday": "Lundi", "Tuesday": "Mardi", "Wednesday": "Mercredi", 
                    "Thursday": "Jeudi", "Friday": "Vendredi", "Saturday": "Samedi", "Sunday": "Dimanche"}
        jour_fr = jours_fr.get(jour_semaine, jour_semaine)

        cle = f"{jour_fr} {date_str}"
        if cle not in planning_par_jour:
            planning_par_jour[cle] = {}

        creneau = attr['creneau']
        if creneau not in planning_par_jour[cle]:
            planning_par_jour[cle][creneau] = []

        planning_par_jour[cle][creneau].append(attr)

    # Construction du HTML
    html = """
    <style>
        .planning-table { border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; }
        .planning-table th { background-color: #1565C0; color: white; padding: 12px; text-align: center; border: 2px solid #0D47A1; }
        .planning-table td { padding: 10px; border: 1px solid #90CAF9; vertical-align: top; }
        .creneau-cell { background-color: #E3F2FD; font-weight: bold; text-align: center; }
        .examen-cell { background-color: #FFF8E1; margin: 4px 0; padding: 8px; border-radius: 4px; border-left: 3px solid #FFA000; }
        .surv-permanent { color: #1565C0; font-weight: bold; }
        .surv-vacataire { color: #2E7D32; }
        .surv-autre { color: #E65100; }
        .jour-header { background-color: #1976D2; color: white; font-weight: bold; }
    </style>
    <table class="planning-table">
    """

    # En-têtes
    jours = sorted(planning_par_jour.keys())
    html += "<tr><th>Créneau Horaire</th>"
    for jour in jours:
        html += f"<th>{jour}</th>"
    html += "</tr>"

    # Lignes par créneau
    for creneau in CRENEAUX_HORAIRES:
        html += f"<tr><td class='creneau-cell'>{creneau}</td>"
        for jour in jours:
            html += "<td>"
            if creneau in planning_par_jour[jour]:
                for examen in planning_par_jour[jour][creneau]:
                    surv_html = "<br>".join([
                        f"<span class='surv-{s['qualite'].lower()}'>{s['nom']} ({s['qualite']})</span>"
                        for s in examen['details_surveillants']
                    ])
                    html += f"""
                    <div class='examen-cell'>
                        <strong>{examen['matiere']}</strong><br>
                        <small>Promo: {examen['promotion']} | Lieu: {examen['lieu']}</small><br>
                        <small>Surveillants:<br>{surv_html}</small>
                    </div>
                    """
            html += "</td>"
        html += "</tr>"

    html += "</table>"
    return html

def generer_excel_coloré(attributions, filename="planning_surveillances.xlsx"):
    """Génère un fichier Excel coloré"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Planning Surveillances"

    # Préparation des données
    data = []
    for attr in attributions:
        date_str = attr['date'].strftime('%d/%m/%Y') if hasattr(attr['date'], 'strftime') else str(attr['date'])
        surveillants_str = ", ".join([f"{s['nom']} ({s['qualite']})" for s in attr['details_surveillants']])
        data.append({
            'Date': date_str,
            'Jour': attr['date'].strftime('%A') if hasattr(attr['date'], 'strftime') else '',
            'Créneau': attr['creneau'],
            'Matière': attr['matiere'],
            'Promotion': attr['promotion'],
            'Lieu': attr['lieu'],
            'Surveillants': surveillants_str
        })

    df = pd.DataFrame(data)

    # Écriture avec styles
    header_fill = PatternFill(start_color="1565C0", end_color="1565C0", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)

            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                # Coloration alternée par jour
                if c_idx == 1:  # Colonne date
                    cell.alignment = Alignment(horizontal='center')
                cell.border = Border(
                    left=Side(style='thin', color='90CAF9'),
                    right=Side(style='thin', color='90CAF9'),
                    top=Side(style='thin', color='90CAF9'),
                    bottom=Side(style='thin', color='90CAF9')
                )

    # Ajustement des largeurs
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column].width = adjusted_width

    # Deuxième feuille : Résumé par enseignant
    ws2 = wb.create_sheet("Résumé Enseignants")

    # Calculer le nombre de surveillances par enseignant
    ens_count = {}
    for attr in attributions:
        for surv in attr['details_surveillants']:
            nom = surv['nom']
            if nom not in ens_count:
                ens_count[nom] = {'count': 0, 'qualite': surv['qualite'], 'examens': []}
            ens_count[nom]['count'] += 1
            ens_count[nom]['examens'].append(f"{attr['matiere']} ({attr['date'].strftime('%d/%m')})")

    resume_data = []
    for nom, info in sorted(ens_count.items(), key=lambda x: x[1]['count'], reverse=True):
        resume_data.append({
            'Enseignant': nom,
            'Qualité': info['qualite'],
            'Nombre de Surveillances': info['count'],
            'Examens Surveillés': "; ".join(info['examens'][:5]) + ("..." if len(info['examens']) > 5 else "")
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
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws2.column_dimensions[column].width = min(max_length + 2, 60)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def generer_pdf(attributions, filename="planning_surveillances.pdf"):
    """Génère un PDF avec le planning"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), 
                           rightMargin=1*cm, leftMargin=1*cm, 
                           topMargin=1*cm, bottomMargin=1*cm)

    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#1565C0'),
        spaceAfter=20,
        alignment=1  # Centre
    )

    elements.append(Paragraph("PLANNING DES SURVEILLANCES D'EXAMENS", title_style))
    elements.append(Paragraph(f"Année Académique 2026-2027 - Semestre 1", styles['Heading2']))
    elements.append(Spacer(1, 0.5*cm))

    # Données du tableau
    table_data = [['Date', 'Jour', 'Créneau', 'Matière', 'Promotion', 'Lieu', 'Surveillants']]

    for attr in attributions:
        date_str = attr['date'].strftime('%d/%m/%Y') if hasattr(attr['date'], 'strftime') else str(attr['date'])
        jour = attr['date'].strftime('%A') if hasattr(attr['date'], 'strftime') else ''
        surveillants_str = ", ".join([f"{s['nom']} ({s['qualite']})" for s in attr['details_surveillants']])

        table_data.append([
            date_str,
            jour,
            attr['creneau'],
            attr['matiere'],
            attr['promotion'],
            attr['lieu'],
            surveillants_str
        ])

    # Création du tableau
    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#E3F2FD')),
        ('GRID', (0, 0), (-1, -1), 1, colors.HexColor('#90CAF9')),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.HexColor('#E3F2FD'), colors.HexColor('#FFFFFF')]),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))

    elements.append(table)
    doc.build(elements)
    buffer.seek(0)
    return buffer

# ============================================================
# INTERFACE PRINCIPALE
# ============================================================

def main():
    init_session_state()
    st.markdown('📋 Gestion des Surveillances d\'Examens 2026-2027', unsafe_allow_html=True)
    
    # Sidebar - Configuration
    with st.sidebar:
        st.markdown("## ⚙️ Configuration")

        # Upload des fichiers
        st.markdown("### 📁 Fichiers Source")
        col1, col2 = st.columns(2)
        with col1:
            fichier_ens = st.file_uploader("Enseignants", type=['xlsx', 'xls'], key="ens_file")
        with col2:
            fichier_exam = st.file_uploader("Examens", type=['xlsx', 'xls'], key="exam_file")

        if fichier_ens is not None:
            st.session_state.enseignants_df = charger_fichier_enseignants(fichier_ens)
            if st.session_state.enseignants_df is not None:
                st.success(f"✅ {len(st.session_state.enseignants_df)} enseignants chargés")

        if fichier_exam is not None:
            st.session_state.examens_df = charger_fichier_examens(fichier_exam)
            if st.session_state.examens_df is not None:
                st.success(f"✅ {len(st.session_state.examens_df)} examens chargés")

        st.markdown("---")

        # Configuration des quotas
        st.markdown("### 📊 Quotas de Surveillance")
        st.session_state.nb_surv_permanent = st.number_input("Permanent", min_value=0, max_value=20, value=3, key="quota_perm")
        st.session_state.nb_surv_vacataire = st.number_input("Vacataire", min_value=0, max_value=20, value=2, key="quota_vac")
        st.session_state.nb_surv_autre = st.number_input("Autre", min_value=0, max_value=20, value=1, key="quota_autre")
        st.session_state.nb_surv_par_lieu = st.number_input("Surveillants par lieu", min_value=1, max_value=5, value=2, key="nb_surv_lieu")

        st.markdown("---")

        # Date de début
        st.markdown("### 📅 Date de Début")
        st.session_state.date_debut = st.date_input("Date début session", value=datetime(2026, 11, 1), key="date_debut")

        # Jours fériés
        st.markdown("### 🎉 Jours Fériés")
        jours_feries_input = st.text_area("Dates (JJ/MM/AAAA, séparées par des virgules)", 
                                          placeholder="11/11/2026, 25/12/2026...", key="jours_feries_str")
        if jours_feries_input:
            try:
                st.session_state.jours_feries = [
                    datetime.strptime(d.strip(), "%d/%m/%Y").date() 
                    for d in jours_feries_input.split(",") if d.strip()
                ]
            except:
                st.warning("Format invalide. Utilisez JJ/MM/AAAA")

        st.markdown("---")
        st.info("💡 Les vendredis et samedis sont exclus automatiquement.")

    # Onglets principaux
    tabs = st.tabs(["🏠 Accueil", "👥 Enseignants", "📚 Examens & Planning", "🎯 Attributions", "📊 Export"])

    # ==================== ONGLET ACCUEIL ====================
    with tabs[0]:
        st.markdown("""
        <div class="info-box">
            <h3>Bienvenue dans l'application de gestion des surveillances d'examens</h3>
            <p>Cette application vous permet de :</p>
            <ul>
                <li>📁 Charger vos fichiers Excel d'enseignants et d'examens</li>
                <li>👥 Gérer les listes de permanents et vacataires</li>
                <li>📅 Planifier automatiquement les examens sur la session</li>
                <li>🎯 Attribuer intelligemment les surveillants</li>
                <li>📊 Exporter les plannings en HTML, Excel et PDF</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Enseignants", len(st.session_state.enseignants_df) if st.session_state.enseignants_df is not None else 0)
        with col2:
            st.metric("Examens", len(st.session_state.examens_df) if st.session_state.examens_df is not None else 0)
        with col3:
            if st.session_state.surveillance_df is not None:
                st.metric("Attributions", len(st.session_state.surveillance_df))
            else:
                st.metric("Attributions", 0)

    # ==================== ONGLET ENSEIGNANTS ====================
    with tabs[1]:
        st.markdown('<div class="sub-header">Gestion des Enseignants</div>', unsafe_allow_html=True)

        if st.session_state.enseignants_df is not None:
            df_ens = st.session_state.enseignants_df

            # Affichage par catégorie
            col1, col2 = st.columns(2)

            with col1:
                st.markdown("#### 👔 Permanents")
                permanents = df_ens[df_ens['qualite'] == 'Permanent']['nom'].tolist()
                exclus_perm = st.multiselect(
                    "Exclure de la surveillance (0 surveillance)",
                    options=permanents,
                    default=st.session_state.exclus_permanents,
                    key="exclus_perm_select"
                )
                st.session_state.exclus_permanents = exclus_perm

                # Affichage avec statut
                perm_df = df_ens[df_ens['qualite'] == 'Permanent'][['nom', 'matiere', 'promotion']].copy()
                perm_df['Statut'] = perm_df['nom'].apply(lambda x: '❌ Exclu' if x in exclus_perm else '✅ Actif')
                st.dataframe(perm_df, use_container_width=True)

            with col2:
                st.markdown("#### 📝 Vacataires")
                vacataires = df_ens[df_ens['qualite'] == 'Vacataire']['nom'].tolist()
                exclus_vac = st.multiselect(
                    "Exclure de la surveillance (0 surveillance)",
                    options=vacataires,
                    default=st.session_state.exclus_vacataires,
                    key="exclus_vac_select"
                )
                st.session_state.exclus_vacataires = exclus_vac

                vac_df = df_ens[df_ens['qualite'] == 'Vacataire'][['nom', 'matiere', 'promotion']].copy()
                vac_df['Statut'] = vac_df['nom'].apply(lambda x: '❌ Exclu' if x in exclus_vac else '✅ Actif')
                st.dataframe(vac_df, use_container_width=True)

            # Résumé
            st.markdown("---")
            st.markdown("#### 📈 Résumé des Effectifs")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Permanents actifs", len(permanents) - len(exclus_perm))
            with col2:
                st.metric("Vacataires actifs", len(vacataires) - len(exclus_vac))
            with col3:
                autres = len(df_ens[~df_ens['qualite'].isin(['Permanent', 'Vacataire'])])
                st.metric("Autres", autres)
            with col4:
                st.metric("Total", len(df_ens))
        else:
            st.warning("⚠️ Veuillez charger le fichier des enseignants dans la barre latérale.")

    # ==================== ONGLET EXAMENS & PLANNING ====================
    with tabs[2]:
        st.markdown('<div class="sub-header">Planification des Examens</div>', unsafe_allow_html=True)

        if st.session_state.examens_df is not None:
            df_exam = st.session_state.examens_df

            # Configuration par promotion
            st.markdown("#### ⚙️ Configuration par Promotion")
            promotions = df_exam['promotion'].dropna().unique()

            for promo in promotions:
                if promo == '':
                    continue

                with st.expander(f"📚 Promotion: {promo}"):
                    col1, col2, col3 = st.columns(3)

                    with col1:
                        creneaux_selection = st.multiselect(
                            f"Créneaux horaires ({promo})",
                            options=CRENEAUX_HORAIRES,
                            default=CRENEAUX_HORAIRES,
                            key=f"creneaux_{promo}"
                        )

                    with col2:
                        lieux_options = st.multiselect(
                            f"Lieux ({promo})",
                            options=LIEUX,
                            default=['S01', 'S02', 'A01'],
                            key=f"lieux_{promo}"
                        )

                    with col3:
                        nb_jours = st.number_input(
                            f"Durée estimée (jours) ({promo})",
                            min_value=1, max_value=30,
                            value=len(df_exam[df_exam['promotion'] == promo]),
                            key=f"duree_{promo}"
                        )

                    st.session_state.config_creneaux[promo] = {
                        'creneaux': creneaux_selection,
                        'lieux': lieux_options,
                        'duree_jours': nb_jours
                    }

            # Génération du planning
            if st.button("🚀 Générer le Planning Automatique", type="primary", key="gen_planning"):
                with st.spinner("Génération en cours..."):
                    planning = generer_planning(
                        df_exam,
                        st.session_state.date_debut,
                        st.session_state.jours_feries,
                        st.session_state.config_creneaux
                    )
                    st.session_state.planning_df = planning
                    st.success("✅ Planning généré avec succès!")

            # Affichage et modification du planning
            if st.session_state.planning_df is not None:
                st.markdown("---")
                st.markdown("#### 📝 Planning Généré (Modifiable)")

                planning = st.session_state.planning_df.copy()

                # Conversion des dates pour l'affichage
                if 'date' in planning.columns:
                    planning['date_affichage'] = planning['date'].apply(
                        lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else str(x)
                    )

                # Édition du planning
                edited_planning = st.data_editor(
                    planning,
                    column_config={
                        "date": st.column_config.DateColumn("Date", format="DD/MM/YYYY"),
                        "creneau": st.column_config.SelectboxColumn("Créneau", options=CRENEAUX_HORAIRES),
                        "lieu": st.column_config.SelectboxColumn("Lieu", options=LIEUX),
                        "matiere": st.column_config.TextColumn("Matière", disabled=True),
                        "promotion": st.column_config.TextColumn("Promotion", disabled=True),
                    },
                    use_container_width=True,
                    num_rows="fixed",
                    key="planning_editor"
                )

                st.session_state.planning_df = edited_planning

                # Visualisation du planning
                st.markdown("#### 📅 Vue Calendrier")

                # Créer un résumé par jour
                if 'date' in edited_planning.columns:
                    cal_data = []
                    for date_val in sorted(edited_planning['date'].unique()):
                        if pd.isna(date_val):
                            continue
                        day_exams = edited_planning[edited_planning['date'] == date_val]
                        for _, row in day_exams.iterrows():
                            cal_data.append({
                                'Date': date_val.strftime('%d/%m/%Y') if hasattr(date_val, 'strftime') else str(date_val),
                                'Jour': date_val.strftime('%A') if hasattr(date_val, 'strftime') else '',
                                'Créneau': row.get('creneau', ''),
                                'Matière': row.get('matiere', ''),
                                'Promotion': row.get('promotion', ''),
                                'Lieu': row.get('lieu', '')
                            })

                    if cal_data:
                        cal_df = pd.DataFrame(cal_data)
                        st.dataframe(cal_df, use_container_width=True)
        else:
            st.warning("⚠️ Veuillez charger le fichier des examens dans la barre latérale.")

    # ==================== ONGLET ATTRIBUTIONS ====================
    with tabs[3]:
        st.markdown('<div class="sub-header">Attribution des Surveillants</div>', unsafe_allow_html=True)

        if st.session_state.planning_df is not None and st.session_state.enseignants_df is not None:
            if st.button("🎯 Attribuer les Surveillants", type="primary", key="attrib_surv"):
                with st.spinner("Attribution en cours..."):
                    attributions, ens_maj = attribuer_surveillants(
                        st.session_state.planning_df,
                        st.session_state.enseignants_df,
                        st.session_state.nb_surv_par_lieu
                    )
                    st.session_state.surveillance_df = attributions
                    st.session_state.enseignants_df = ens_maj
                    st.success("✅ Attributions effectuées!")

            if st.session_state.surveillance_df is not None:
                st.markdown("---")

                # Tableau détaillé
                st.markdown("#### 📋 Tableau des Attributions")

                attr_data = []
                for attr in st.session_state.surveillance_df:
                    date_str = attr['date'].strftime('%d/%m/%Y') if hasattr(attr['date'], 'strftime') else str(attr['date'])
                    for i, surv in enumerate(attr['details_surveillants']):
                        attr_data.append({
                            'Date': date_str,
                            'Créneau': attr['creneau'],
                            'Matière': attr['matiere'],
                            'Promotion': attr['promotion'],
                            'Lieu': attr['lieu'],
                            'Surveillant': surv['nom'],
                            'Qualité': surv['qualite'],
                            'Rôle': 'Chargé de matière' if surv.get('priorite') == 'Chargé de matière' else surv['qualite']
                        })

                attr_df = pd.DataFrame(attr_data)

                # Style conditionnel
                def color_qualite(val):
                    colors = {'Permanent': 'background-color: #E3F2FD', 
                             'Vacataire': 'background-color: #E8F5E9',
                             'Autre': 'background-color: #FFF3E0'}
                    return colors.get(val, '')

                styled_df = attr_df.style.applymap(color_qualite, subset=['Qualité'])
                st.dataframe(styled_df, use_container_width=True)

                # Vérification des chevauchements
                st.markdown("---")
                st.markdown("#### 🔍 Vérification des Chevauchements")

                chevauchements = []
                ens_surveillances = {}
                for attr in st.session_state.surveillance_df:
                    key = (attr['date'], attr['creneau'])
                    for surv in attr['details_surveillants']:
                        nom = surv['nom']
                        if nom not in ens_surveillances:
                            ens_surveillances[nom] = []
                        ens_surveillances[nom].append({
                            'date': attr['date'],
                            'creneau': attr['creneau'],
                            'matiere': attr['matiere']
                        })

                for nom, seances in ens_surveillances.items():
                    seen = set()
                    for seance in seances:
                        key = (seance['date'], seance['creneau'])
                        if key in seen:
                            chevauchements.append({
                                'Enseignant': nom,
                                'Date': seance['date'].strftime('%d/%m/%Y') if hasattr(seance['date'], 'strftime') else str(seance['date']),
                                'Créneau': seance['creneau'],
                                'Problème': 'Double attribution'
                            })
                        seen.add(key)

                if chevauchements:
                    st.error("⚠️ Chevauchements détectés!")
                    st.dataframe(pd.DataFrame(chevauchements), use_container_width=True)
                else:
                    st.success("✅ Aucun chevauchement détecté!")

                # Résumé par enseignant
                st.markdown("---")
                st.markdown("#### 📊 Résumé par Enseignant")

                resume = {}
                for attr in st.session_state.surveillance_df:
                    for surv in attr['details_surveillants']:
                        nom = surv['nom']
                        if nom not in resume:
                            resume[nom] = {'count': 0, 'qualite': surv['qualite']}
                        resume[nom]['count'] += 1

                resume_df = pd.DataFrame([
                    {'Enseignant': k, 'Qualité': v['qualite'], 'Surveillances': v['count']}
                    for k, v in sorted(resume.items(), key=lambda x: x[1]['count'], reverse=True)
                ])

                st.dataframe(resume_df, use_container_width=True)

                # Graphique
                if not resume_df.empty:
                    chart_data = resume_df.groupby('Qualité')['Surveillances'].sum().reset_index()
                    st.bar_chart(chart_data.set_index('Qualité'))
        else:
            st.warning("⚠️ Veuillez d'abord générer le planning et charger les enseignants.")

    # ==================== ONGLET EXPORT ====================
    with tabs[4]:
        st.markdown('<div class="sub-header">Export des Plannings</div>', unsafe_allow_html=True)

        if st.session_state.surveillance_df is not None:
            attributions = st.session_state.surveillance_df

            col1, col2, col3 = st.columns(3)

            with col1:
                st.markdown("#### 📄 HTML")
                html_content = generer_tableau_html(attributions)
                st.download_button(
                    label="⬇️ Télécharger HTML",
                    data=html_content,
                    file_name="planning_surveillances.html",
                    mime="text/html",
                    key="download_html"
                )
                with st.expander("👁️ Aperçu HTML"):
                    st.components.v1.html(html_content, height=600, scrolling=True)

            with col2:
                st.markdown("#### 📊 Excel")
                excel_buffer = generer_excel_coloré(attributions)
                st.download_button(
                    label="⬇️ Télécharger Excel",
                    data=excel_buffer,
                    file_name="planning_surveillances.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel"
                )

            with col3:
                st.markdown("#### 📑 PDF")
                pdf_buffer = generer_pdf(attributions)
                st.download_button(
                    label="⬇️ Télécharger PDF",
                    data=pdf_buffer,
                    file_name="planning_surveillances.pdf",
                    mime="application/pdf",
                    key="download_pdf"
                )

            # Aperçu du tableau final
            st.markdown("---")
            st.markdown("#### 🎨 Aperçu du Planning Final")
            st.markdown(generer_tableau_html(attributions), unsafe_allow_html=True)
        else:
            st.warning("⚠️ Aucune attribution à exporter. Veuillez d'abord attribuer les surveillants.")

if __name__ == "__main__":
    main()
