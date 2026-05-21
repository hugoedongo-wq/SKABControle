import streamlit as st
import pandas as pd
import io
import numpy as np
from datetime import datetime

# Configuration stricte de la page pour Streamlit Cloud
st.set_page_config(
    page_title="SKAB Nutrition — Consolidation CI",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -----------------------------------------------------------------------------
# MOTEUR DE CONSOLIDATION (TRAITEMENT EN MÉMOIRE RAM)
# -----------------------------------------------------------------------------
def extraire_metadonnees_fichier(nom_fichier):
    """
    Analyse le nom du fichier pour extraire le code contrôleur et le pays.
    Format attendu : SKAB_Controleur_CI-001_Cameroun_2026.xlsx
    """
    nom_brut = nom_fichier.replace(".xlsx", "")
    parties = nom_brut.split("_")
    
    # Valeurs par défaut si le nom ne respecte pas le pattern
    code_ctrl = "Inconnu"
    pays_filiale = "Non spécifié"
    
    if len(parties) >= 4:
        code_ctrl = parties[2]     # ex: CI-001
        pays_filiale = parties[3]  # ex: Cameroun
    elif len(parties) == 3:
        code_ctrl = parties[2]
        
    return code_ctrl, pays_filiale

def consolider_les_fichiers(fichiers_charges):
    """
    Boucle sur tous les fichiers injectés, lit les onglets correspondants,
    fusionne les lignes et ajoute les colonnes de traçabilité.
    """
    # Onglets officiels définis dans le template SKAB
    onglets_cibles = {
        "MES_MISSIONS": "CONSO_MISSIONS",
        "POINTS_CONTROLE": "CONSO_POINTS_CONTROLE",
        "ANOMALIES": "CONSO_ANOMALIES",
        "PLANS_ACTION": "CONSO_PLANS_ACTION"
    }
    
    # Initialisation des listes pour stocker les DataFrames de chaque contrôleur
    donnies_assimilees = {onglet: [] for onglet in onglets_cibles.keys()}
    
    for fichier in fichiers_charges:
        try:
            # Lecture du fichier Excel en mémoire (BytesIO requis sur Streamlit Cloud)
            contenu_fichier = fichier.read()
            excel_object = pd.ExcelFile(io.BytesIO(contenu_fichier))
            
            # Extraction des infos de traçabilité via le nom du fichier
            code_ctrl, pays_filiale = extraire_metadonnees_fichier(fichier.name)
            
            for onglet in onglets_cibles.keys():
                if onglet in excel_object.sheet_names:
                    # Lecture de l'onglet
                    df = pd.read_excel(excel_object, sheet_name=onglet)
                    
                    # Nettoyage : suppression des lignes complètement vides
                    df = df.dropna(how='all')
                    
                    if not df.empty:
                        # Ajout des colonnes de suivi pour le reporting DAF
                        df["Contrôleur Source"] = code_ctrl
                        df["Zone / Filiale"] = pays_filiale
                        df["Nom Fichier Origine"] = fichier.name
                        
                        donnies_assimilees[onglet].append(df)
                        
            # Réinitialiser le pointeur du fichier pour les lectures suivantes
            fichier.seek(0)
            
        except Exception as e:
            st.error(f"❌ Impossible de lire le fichier `{fichier.name}`. Vérifiez sa structure. Erreur : {e}")

    # --- ÉCRITURE DU FICHIER MAÎTRE FINAL ---
    output_stream = io.BytesIO()
    
    with pd.ExcelWriter(output_stream, engine='xlsxwriter') as writer:
        # 1. Onglet d'Accueil et de Validation pour le Chef de Département
        df_accueil = pd.DataFrame({
            "CONTRÔLE INTERNE GROUPE SKAB": [
                "Livrable destiné à",
                "Généré par",
                "Date et Heure de consolidation",
                "Nombre de fichiers contrôleurs inclus",
                "Statut du document"
            ],
            "DONNÉES DE TRAÇABILITÉ": [
                "M. Élie DIGNOU (DAF)",
                "Chef de Département Contrôle Interne",
                datetime.now().strftime("%d/%m/%Y à %H:%M:%S"),
                len(fichiers_charges),
                "VALIDÉ — Prêt pour mise à jour du Dashboard Exécutif"
            ]
        })
        df_accueil.to_excel(writer, sheet_name="ACCUEIL_CD", index=False)
        
        # Outils de stylisation Excel via XlsxWriter
        workbook = writer.book
        worksheet_accueil = writer.sheets["ACCUEIL_CD"]
        worksheet_accueil.set_column('A:B', 40)
        
        # 2. Fusion et injection des onglets consolidés
        for onglet_template, nom_onglet_conso in onglets_cibles.items():
            list_dfs = donnies_assimilees[onglet_template]
            
            if list_dfs:
                # Fusionner tous les fichiers pour cet onglet précis
                df_structure_total = pd.concat(list_dfs, ignore_index=True)
            else:
                # Créer un tableau vide avec les colonnes de traçabilité de base si aucun contrôleur n'a de lignes
                df_structure_total = pd.DataFrame(columns=["ID", "Statut", "Contrôleur Source", "Zone / Filiale", "Nom Fichier Origine"])
                
            # Écriture dans le classeur final
            df_structure_total.to_excel(writer, sheet_name=nom_onglet_conso, index=False)
            
            # Ajustement automatique SÉCURISÉ de la largeur des colonnes
            worksheet = writer.sheets[nom_onglet_conso]
            for idx, col in enumerate(df_structure_total.columns):
                series = df_structure_total[col]
                
                # --- CORRECTION BLINDÉE DE L'ERREUR DE TYPE ET DE LONGUEUR ---
                # On nettoie les valeurs nulles pour éviter de fausser les calculs
                clean_series = series.dropna()
                
                if not clean_series.empty:
                    # Utilisation de .apply(str) au lieu de .astype(str).map(len) pour contourner le bug de l'interpréteur de type pandas
                    lengths = clean_series.apply(lambda x: len(str(x)))
                    max_cells_len = lengths.max()
                    
                    # Double sécurité : si le max calculé est indéterminé ou invalide (NaN)
                    if pd.isna(max_cells_len) or not isinstance(max_cells_len, (int, float, np.integer)):
                        max_cells_len = 0
                else:
                    max_cells_len = 0
                
                # Comparaison sécurisée avec la longueur du titre de la colonne
                max_len = max(int(max_cells_len), len(str(col))) + 3
                max_len = min(max_len, 50)  # Limitation pour éviter les colonnes géantes de commentaires
                
                worksheet.set_column(idx, idx, max_len)
                
    return output_stream.getvalue()

# -----------------------------------------------------------------------------
# INTERFACE UTILISATEUR (STREAMLIT UI)
# -----------------------------------------------------------------------------
st.title("🌾 SKAB Nutrition — Direction de l'Audit et du Contrôle Interne")
st.markdown("### `Espace de Consolidation Multi-Filiales`")
st.write(
    "Cet outil Web unifié permet de compiler instantanément les journaux de contrôle, "
    "les listes d'anomalies et les plans d'action terrain reçus des différents pays (Cameroun, Tchad, Gabon, RCA...)."
)
st.markdown("---")

st.subheader("👑 Session du Chef de Département")
st.info(
    "💡 **Rappel de la procédure :** Collectez les fichiers `.xlsx` de vos contrôleurs "
    "via vos canaux habituels, puis glissez-les tous ensemble ci-dessous."
)

fichiers_recus = st.file_uploader(
    "Sélectionnez ou déposez les fichiers des contrôleurs (Format .xlsx uniquement) :",
    type=["xlsx"],
    accept_multiple_files=True,
    key="uploader_maitre"
)

if fichiers_recus:
    st.markdown("### 📊 Fichiers prêts pour la fusion")
    
    details_chargement = []
    for f in fichiers_recus:
        code, zone = extraire_metadonnees_fichier(f.name)
        details_chargement.append({
            "Nom du fichier": f.name,
            "ID Contrôleur extrait": code,
            "Zone / Filiale": zone,
            "Taille": f"{f.size / 1024:.2f} KB"
        })
        
    st.dataframe(pd.DataFrame(details_chargement), use_container_width=True)
    
    st.markdown("---")
    
    if st.button("🔄 Lancer la consolidation globale", type="primary", use_container_width=True):
        with st.spinner("Analyse des structures Excel, injection des traçabilités et fusion des onglets..."):
            
            fichier_maitre_bytes = consolider_les_fichiers(fichiers_recus)
            
            st.success("🎉 Fichier MAÎTRE compilé avec succès en mémoire vive !")
            
            date_iso = datetime.now().strftime("%Y%m%d")
            nom_maitre_final = f"SKAB_MAITRE_CD_CONSOLIDE_{date_iso}.xlsx"
            
            zone_telechargement, zone_instructions = st.columns([1, 1])
            
            with zone_telechargement:
                st.write("📥 **Votre livrable est prêt :**")
                st.download_button(
                    label="💾 Télécharger le fichier MAÎTRE (.xlsx)",
                    data=fichier_maitre_bytes,
                    file_name=nom_maitre_final,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            with zone_instructions:
                st.warning(
                    f"📧 **Action DAF :** Vous pouvez maintenant récupérer ce fichier téléchargé et "
                    f"le transmettre directement à **M. Élie DIGNOU (DAF)** pour l'arbitrage budgétaire des "
                    f"plans de remédiation."
                )
else:
    st.info("⏳ En attente de fichiers pour démarrer le processus de traitement.")

st.markdown("---")
st.caption(
    f"© {datetime.now().year} Groupe SKAB Nutrition — "
    f"Application de centralisation du Contrôle Interne hébergée sur Streamlit Cloud."
)
