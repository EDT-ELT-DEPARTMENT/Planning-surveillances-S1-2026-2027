import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
import io
import os

# ═══════════════════════════════════════════════════════════════
# CONFIGURATION
# ═══════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Gestion des Surveillances d'Examens 2026-2027",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown('''
<style>
    .main-header { font-size: 2rem; font-weight: 700; color: #1f77b4; margin-bottom: 1rem; }
    .sub-header { font-size: 1.2rem; font-weight: 600; color: #333; margin-top: 1rem; margin-bottom: 0.5rem; }
    .info-box { background-color: #e8f4f8; border-left: 4px solid #1f77b4; padding: 1rem; border-radius: 4px; margin-bottom: 1rem; }
    .success-box { background-color: #e8f5e9; border-left: 4px solid #4caf50; padding: 1rem; border-radius: 4px; margin-bottom: 1rem; }
</style>
''', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# FONCTIONS UTILITAIRES
# ═══════════════════════════════════════════════════════════════
@st.cache_data
def load_excel_path(filepath):
    try:
        df = pd.read_excel(filepath, engine="openpyxl")
        if len(df) > 0 and (df.iloc[0].astype(str) == df.columns).all():
            df = df.iloc[1:].reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return None


@st.cache_data
def load_excel_file(file):
    try:
        df = pd.read_excel(file, engine="openpyxl")
        if len(df) > 0 and (df.iloc[0].astype(str) == df.columns).all():
            df = df.iloc[1:].reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Erreur lors du chargement : {e}")
        return None


def parse_horaire(horaire_str):
    if pd.isna(horaire_str):
        return (0, 0)
    h = str(horaire_str).strip().replace(" ", "").replace("–", "-")
    try:
        parts = h.split("-")
        if len(parts) != 2:
            return (0, 0)

        def to_dec(t):
            t = t.strip().lower().replace("h", ":").replace("H", ":")
            if ":" in t:
                hh, mm = t.split(":")[:2]
                return int(hh) + int(mm) / 60
            else:
                return int(t) if t.isdigit() else 0

        return (to_dec(parts[0]), to_dec(parts[1]))
    except Exception:
        return (0, 0)


def chevauchement(h1, h2):
    return not (h1[1] <= h2[0] or h2[1] <= h1[0])


def get_jour_index(jour_str):
    jours = {"lundi": 0, "mardi": 1, "mercredi": 2, "jeudi": 3,
             "vendredi": 4, "samedi": 5, "dimanche": 6}
    return jours.get(str(jour_str).strip().lower(), 0)


def detecter_qualite(qualite_str):
    q = str(qualite_str).strip().lower()
    if "permanent" in q or "permenant" in q or "associé" in q:
        return "Permanent"
    elif "vacataire" in q:
        return "Vacataire"
    elif "retraité" in q:
        return "Retraité"
    elif "disponibilité" in q:
        return "Disponibilité"
    else:
        return "Autre"


def charger_fichiers_auto():
    chemins_ens = [
        "Liste des enseignants-2026-2027.xlsx",
        "data/Liste des enseignants-2026-2027.xlsx",
        "assets/Liste des enseignants-2026-2027.xlsx",
        "./Liste des enseignants-2026-2027.xlsx"
    ]
    chemins_exam = [
        "DATA-ENS-2026-2027.xlsx",
        "data/DATA-ENS-2026-2027.xlsx",
        "assets/DATA-ENS-2026-2027.xlsx",
        "./DATA-ENS-2026-2027.xlsx"
    ]

    df_ens = None
    df_exam = None
    ens_path = None
    exam_path = None

    for p in chemins_ens:
        if os.path.exists(p):
            df_ens = load_excel_path(p)
            if df_ens is not None and not df_ens.empty:
                ens_path = p
                break

    for p in chemins_exam:
        if os.path.exists(p):
            df_exam = load_excel_path(p)
            if df_exam is not None and not df_exam.empty:
                exam_path = p
                break

    return df_ens, df_exam, ens_path, exam_path


def attribuer_surveillants(df_examens, df_enseignants, quota_perm,
                             quota_vac, quota_autre, nb_surv_lieu):
    if df_examens is None or df_enseignants is None:
        return None

    col_ens = None
    for c in df_enseignants.columns:
        if "enseignant" in c.lower():
            col_ens = c
            break
    if col_ens and col_ens != "Enseignants":
        df_enseignants = df_enseignants.rename(columns={col_ens: "Enseignants"})

    surveillants = []
    for _, row in df_enseignants.iterrows():
        nom = str(row.get("Enseignants", "")).strip()
        if not nom or nom.lower() in ["non défini", "nan", "", "none"]:
            continue

        qualite_brute = str(row.get("Qualité", "Autre")).strip()
        cat = detecter_qualite(qualite_brute)
        grade = str(row.get("Grade", "")).strip()
        email = str(row.get("Email", "")).strip()
        tel = str(row.get("N°/TEL", "")).strip()

        quota = quota_autre
        if cat == "Permanent":
            quota = quota_perm
        elif cat == "Vacataire":
            quota = quota_vac

        surveillants.append({
            "nom": nom,
            "qualite": qualite_brute,
            "categorie": cat,
            "grade": grade,
            "email": email,
            "tel": tel,
            "quota": quota,
            "assignations": 0
        })

    if not surveillants:
        st.warning("Aucun enseignant valide trouvé dans le fichier.")
        return None

    df_ex = df_examens.copy()
    df_ex["jour_idx"] = df_ex["Jours"].apply(get_jour_index)
    df_ex["horaire_tuple"] = df_ex["Horaire"].apply(parse_horaire)
    df_ex = df_ex.sort_values(["jour_idx", "horaire_tuple"]).reset_index(drop=True)

    resultats = []

    for idx, examen in df_ex.iterrows():
        jour = examen.get("Jours", "")
        horaire = examen.get("Horaire", "")
        h_tuple = parse_horaire(horaire)
        lieu = examen.get("Lieu", "")
        promotion = examen.get("Promotion", "")
        code = examen.get("Code", "")
        enseignant_cours = str(examen.get("Enseignants", "")).strip()

        nb_salles = max(1, len(str(lieu).split("/")))
        nb_needed = nb_surv_lieu * nb_salles

        disponibles = []
        for s in surveillants:
            if s["nom"].lower() == enseignant_cours.lower():
                continue
            if s["assignations"] >= s["quota"]:
                continue
            conflit = False
            for r in resultats:
                if r["Surveillant"] == s["nom"] and r["Jour"] == jour:
                    if chevauchement(parse_horaire(r["Horaire"]), h_tuple):
                        conflit = True
                        break
            if not conflit:
                disponibles.append(s)

        choisis = []
        np.random.shuffle(disponibles)
        disponibles.sort(key=lambda x: x["assignations"])

        for s in disponibles:
            if len(choisis) >= nb_needed:
                break
            choisis.append(s)
            s["assignations"] += 1

        if not choisis:
            choisis = [{"nom": "NON ATTRIBUÉ", "qualite": "-", "categorie": "-",
                        "grade": "-", "email": "-", "tel": "-"}]

        for s in choisis:
            resultats.append({
                "Jour": jour,
                "Date": "",
                "Horaire": horaire,
                "Promotion": promotion,
                "Code": code,
                "Enseignement": examen.get("Enseignements", ""),
                "Lieu": lieu,
                "Enseignant_Cours": enseignant_cours,
                "Surveillant": s["nom"],
                "Qualité": s["qualite"],
                "Catégorie": s["categorie"],
                "Grade": s["grade"],
                "Email": s["email"],
                "Téléphone": s["tel"]
            })

    return pd.DataFrame(resultats), surveillants


def calculer_dates(df_planning, date_debut, jours_feries):
    if df_planning is None or df_planning.empty:
        return df_planning

    df = df_planning.copy()
    df["Date"] = ""

    jours_map = {"lundi": 0, "mardi": 1, "mercredi": 2, "jeudi": 3,
                 "vendredi": 4, "samedi": 5, "dimanche": 6}

    dates_calculees = {}
    for jour in df["Jour"].unique():
        idx = jours_map.get(str(jour).strip().lower(), 0)
        delta = (idx - date_debut.weekday()) % 7
        current = date_debut + timedelta(days=delta)
        dates_calculees[jour] = current

    for i, row in df.iterrows():
        jour = str(row["Jour"]).strip()
        if jour in dates_calculees:
            d = dates_calculees[jour]
            if d in jours_feries:
                d = d + timedelta(days=1)
            df.at[i, "Date"] = d.strftime("%d/%m/%Y")

    return df


def to_excel_download(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Planning")
    return output.getvalue()


# ═══════════════════════════════════════════════════════════════
# INTERFACE PRINCIPALE
# ═══════════════════════════════════════════════════════════════
def main():
    st.markdown(
        '<div class="main-header">📋 Gestion des Surveillances d\'Examens 2026-2027</div>',
        unsafe_allow_html=True
    )

    df_ens_auto, df_exam_auto, ens_path, exam_path = charger_fichiers_auto()

    with st.sidebar:
        st.header("⚙️ Configuration")

        if df_ens_auto is not None and df_exam_auto is not None:
            st.markdown(
                '<div class="success-box">✅ Fichiers chargés automatiquement</div>',
                unsafe_allow_html=True
            )
            st.caption(f"📁 {ens_path}")
            st.caption(f"📁 {exam_path}")
        else:
            st.markdown(
                '<div class="info-box">⚠️ Fichiers non trouvés en auto. Utilisez l\'upload ci-dessous.</div>',
                unsafe_allow_html=True
            )

        st.markdown('<div class="sub-header">📁 Fichiers Source (fallback)</div>',
                    unsafe_allow_html=True)
        file_ens = st.file_uploader("Enseignants", type=["xlsx", "xls"], key="up_ens")
        file_exam = st.file_uploader("Examens", type=["xlsx", "xls"], key="up_exam")

        st.markdown('<div class="sub-header">📊 Quotas de Surveillance</div>',
                    unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            quota_perm = st.number_input("Permanent", min_value=0,
                                          max_value=50, value=5, step=1)
            quota_autre = st.number_input("Autre", min_value=0,
                                           max_value=50, value=2, step=1)
        with col2:
            quota_vac = st.number_input("Vacataire", min_value=0,
                                         max_value=50, value=4, step=1)
            nb_surv_lieu = st.number_input("Surveillants par lieu",
                                            min_value=1, max_value=10,
                                            value=2, step=1)

        st.markdown('<div class="sub-header">📅 Date de Début</div>',
                    unsafe_allow_html=True)
        if "date_debut" not in st.session_state:
            st.session_state.date_debut = date(2026, 11, 1)
        date_debut = st.date_input("Date début session",
                                    value=st.session_state.date_debut,
                                    key="date_debut")

        st.markdown('<div class="sub-header">🎉 Jours Fériés</div>',
                    unsafe_allow_html=True)
        jours_feries_input = st.date_input("Ajouter un jour férié",
                                            value=[], key="jours_feries_input")
        if not isinstance(jours_feries_input, list):
            jours_feries = [jours_feries_input] if jours_feries_input else []
        else:
            jours_feries = jours_feries_input

    if file_ens is not None:
        df_ens = load_excel_file(file_ens)
    else:
        df_ens = df_ens_auto

    if file_exam is not None:
        df_exam = load_excel_file(file_exam)
    else:
        df_exam = df_exam_auto

    if df_ens is not None:
        with st.expander("👁️ Aperçu Enseignants (avec Qualité)"):
            st.dataframe(df_ens.head(20), use_container_width=True)
            if "Qualité" in df_ens.columns:
                st.markdown("**Répartition par Qualité :**")
                st.write(df_ens["Qualité"].value_counts())

    if df_exam is not None:
        with st.expander("👁️ Aperçu Examens"):
            st.dataframe(df_exam.head(20), use_container_width=True)

    if df_ens is not None and df_exam is not None:
        st.markdown("---")
        st.markdown('<div class="sub-header">🚀 Génération du Planning</div>',
                    unsafe_allow_html=True)

        if st.button("Générer le planning de surveillance", type="primary",
                     use_container_width=True):
            with st.spinner("Attribution des surveillants en cours..."):
                result = attribuer_surveillants(
                    df_exam, df_ens,
                    quota_perm, quota_vac, quota_autre, nb_surv_lieu
                )

                if result is not None:
                    df_planning, stats_surv = result
                    df_planning = calculer_dates(df_planning, date_debut,
                                                  jours_feries)

                    st.session_state["planning"] = df_planning
                    st.session_state["stats_surv"] = stats_surv
                    st.success(
                        f"Planning généré avec succès ({len(df_planning)} lignes)"
                    )
                else:
                    st.error(
                        "Impossible de générer le planning. Vérifiez vos fichiers."
                    )
    else:
        st.info(
            "👈 Les fichiers ne sont pas encore disponibles. "
            "Placez-les à la racine du projet ou uploadez-les manuellement."
        )

    if "planning" in st.session_state and st.session_state["planning"] is not None:
        df_planning = st.session_state["planning"]
        stats_surv = st.session_state.get("stats_surv", [])

        st.markdown("---")
        st.markdown('<div class="sub-header">📊 Statistiques</div>',
                    unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Examens couverts", df_planning["Code"].nunique())
        with c2:
            st.metric("Total surveillances", len(df_planning))
        with c3:
            st.metric("Surveillants actifs",
                      df_planning["Surveillant"].nunique())
        with c4:
            non_attrib = len(
                df_planning[df_planning["Surveillant"] == "NON ATTRIBUÉ"]
            )
            st.metric("Non attribués", non_attrib)

        if stats_surv:
            st.markdown(
                '<div class="sub-header">📈 Répartition par Qualité</div>',
                unsafe_allow_html=True
            )
            stats_df = pd.DataFrame([
                {"Nom": s["nom"], "Qualité": s["qualite"],
                 "Catégorie": s["categorie"], "Grade": s["grade"],
                 "Quota": s["quota"], "Assignations": s["assignations"],
                 "Reste": s["quota"] - s["assignations"]}
                for s in stats_surv
            ])
            stats_df = stats_df.sort_values("Assignations", ascending=False)
            st.dataframe(stats_df, use_container_width=True, hide_index=True)

            st.markdown("**Synthèse par Catégorie :**")
            cat_summary = stats_df.groupby("Catégorie").agg(
                Total=("Nom", "count"),
                Assignations=("Assignations", "sum"),
                Quota_total=("Quota", "sum")
            ).reset_index()
            cat_summary["Utilisation %"] = (
                cat_summary["Assignations"] / cat_summary["Quota_total"] * 100
            ).round(1)
            st.dataframe(cat_summary, use_container_width=True, hide_index=True)

        st.markdown('<div class="sub-header">🗓️ Planning Complet</div>',
                    unsafe_allow_html=True)
        st.dataframe(df_planning, use_container_width=True, hide_index=True)

        st.markdown('<div class="sub-header">💾 Export</div>',
                    unsafe_allow_html=True)
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            excel_data = to_excel_download(df_planning)
            st.download_button(
                label="📥 Télécharger le planning (Excel)",
                data=excel_data,
                file_name="Planning_Surveillances_2026-2027.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_dl2:
            csv_data = df_planning.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label="📥 Télécharger le planning (CSV)",
                data=csv_data,
                file_name="Planning_Surveillances_2026-2027.csv",
                mime="text/csv",
                use_container_width=True
            )


if __name__ == "__main__":
    main()
