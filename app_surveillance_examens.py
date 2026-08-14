def charger_fichier_source_auto():
    try:
        paths_to_try = [
            FICHIER_SOURCE, 
            os.path.join(os.getcwd(), FICHIER_SOURCE),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), FICHIER_SOURCE),
            os.path.join("/mnt/agents/upload/", FICHIER_SOURCE),
            os.path.join("/mount/src/planning-surveillances-s1-2026-2027/", FICHIER_SOURCE)
        ]
        file_path = None
        for p in paths_to_try:
            if os.path.exists(p):
                file_path = p
                break
        if file_path is None:
            return None, f"Fichier {FICHIER_SOURCE} non trouvé"
        
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        ens_sheet = sheet_names[0] if len(sheet_names) > 0 else None
        
        df_ens = pd.read_excel(file_path, sheet_name=ens_sheet)
        df_ens.columns = [str(col).strip() for col in df_ens.columns]
        
        col_qualite_trouvee = None
        for col in df_ens.columns:
            echantillon = df_ens[col].dropna().astype(str).str.lower().tolist()
            if any('vacataire' in x for x in echantillon):
                col_qualite_trouvee = col
                break
                
        cols_orig = list(df_ens.columns)
        cols_lower = [enlever_accents(c).lower().strip().replace(' ', '_').replace('-', '_') for c in cols_orig]
        
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
                
        if col_qualite_trouvee and ('qualite' not in col_map or not col_map['qualite']):
            col_map['qualite'] = col_qualite_trouvee

        rename_map = {}
        if 'nom' in col_map: rename_map[col_map['nom']] = 'nom'
        if 'qualite' in col_map: rename_map[col_map['qualite']] = 'qualite'
        if 'enseignements' in col_map: rename_map[col_map['enseignements']] = 'enseignements'
        if 'promotion' in col_map: rename_map[col_map['promotion']] = 'promotion'
        
        df_ens = df_ens.rename(columns=rename_map)
        
        for col in ['qualite', 'enseignements', 'promotion']:
            if col not in df_ens.columns:
                df_ens[col] = ''
                
        df_ens = df_ens[df_ens['nom'].notna() & (df_ens['nom'].astype(str).str.strip() != '')].copy()
        df_ens['qualite'] = df_ens['qualite'].apply(normaliser_qualite)
        
        return df_ens, None
        
    except Exception as e:
        return None, str(e)
