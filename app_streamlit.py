import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import openpyxl

# Configuration globale de la page
st.set_page_config(
    page_title="SKAB Nutrition — Consolidation CI",
    page_icon="📊",
    layout="wide"
)

st.title("📊 SKAB Nutrition — Plateforme de Consolidation du Contrôle Interne")
st.markdown("### 🚀 Remplacement de Power Query en environnement Cloud")
st.markdown("Déposez simultanément les fichiers Excel de vos contrôleurs pour fusionner instantanément les données et mettre à jour les indicateurs du Groupe.")

# --- MOTEUR DE TRAITEMENT PANDAS (OPTIMISÉ CLOUD) ---
def extraire_donnees_controleurs(fichiers_charges):
    all_missions = []
    all_points = []
    all_anomalies = []
    all_plans = []
    journal_import = []

    for fichier in fichiers_charges:
        nom_fichier = fichier.name
        try:
            # 1. Lecture sécurisée de l'identité du contrôleur (Feuille PARAMETRES)
            df_param = pd.read_excel(fichier, sheet_name="PARAMETRES", header=None)
            nom_ctrl = df_param.iloc[4, 1] if len(df_param) > 4 else "Inconnu"
            code_ctrl = df_param.iloc[5, 1] if len(df_param) > 5 else "INCONNU"
            pays_ctrl = df_param.iloc[6, 1] if len(df_param) > 6 else "Inconnu"

            # 2. Consolidation de l'onglet MES_MISSIONS (En-tête ligne 5 -> skiprows=4)
            df_mis = pd.read_excel(fichier, sheet_name="MES_MISSIONS", skiprows=4)
            df_mis = df_mis.dropna(subset=["N° Mission"])
            df_mis["Contrôleur"] = nom_ctrl
            df_mis["Code Contrôleur"] = code_ctrl
            df_mis["Pays Originel"] = pays_ctrl
            df_mis["Fichier Source"] = nom_fichier
            all_missions.append(df_mis)

            # 3. Consolidation de l'onglet POINTS_CONTROLE
            df_pts = pd.read_excel(fichier, sheet_name="POINTS_CONTROLE", skiprows=4)
            df_pts = df_pts.dropna(subset=["N° Mission"])
            df_pts["Contrôleur"] = nom_ctrl
            df_pts["Fichier Source"] = nom_fichier
            all_points.append(df_pts)

            # 4. Consolidation de l'onglet ANOMALIES
            df_anom = pd.read_excel(fichier, sheet_name="ANOMALIES", skiprows=4)
            df_anom = df_anom.dropna(subset=["Date détection"])
            df_anom["Contrôleur"] = nom_ctrl
            df_anom["Pays"] = pays_ctrl
            df_anom["Fichier Source"] = nom_fichier
            all_anomalies.append(df_anom)

            # 5. Consolidation de l'onglet PLANS_ACTION
            df_pln = pd.read_excel(fichier, sheet_name="PLANS_ACTION", skiprows=4)
            df_pln = df_pln.dropna(subset=["ID Plan"])
            df_pln["Contrôleur"] = nom_ctrl
            df_pln["Fichier Source"] = nom_fichier
            all_plans.append(df_pln)

            journal_import.append({
                "Fichier": nom_fichier, "Contrôleur": nom_ctrl, "Code": code_ctrl, "Pays": pays_ctrl, "Statut": "✅ Validé & Intégré"
            })

        except Exception as e:
            journal_import.append({
                "Fichier": nom_fichier, "Contrôleur": "Erreur", "Code": "-", "Pays": "-", "Statut": f"❌ Échec de lecture : {str(e)}"
            })

    # Regroupement final
    conso_mis = pd.concat(all_missions, ignore_index=True) if all_missions else pd.DataFrame()
    conso_pts = pd.concat(all_points, ignore_index=True) if all_points else pd.DataFrame()
    conso_anom = pd.concat(all_anomalies, ignore_index=True) if all_anomalies else pd.DataFrame()
    conso_pln = pd.concat(all_plans, ignore_index=True) if all_plans else pd.DataFrame()
    df_journal = pd.DataFrame(journal_import)

    return conso_mis, conso_pts, conso_anom, conso_pln, df_journal

# --- INTERFACE UTISATEUR (UI) ---
fichiers_transmis = st.file_uploader(
    "Téléversez les fichiers Excel des contrôleurs (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Sélectionnez plusieurs fichiers simultanément (Ctrl+A dans votre explorateur)"
)

if fichiers_transmis:
    # Exécution du traitement de données
    conso_mis, conso_pts, conso_anom, conso_pln, df_journal = extraire_donnees_controleurs(fichiers_transmis)
    
    # Affichage du journal d'intégration
    st.subheader("📋 Suivi de l'intégration des filiales")
    st.dataframe(df_journal, use_container_width=True)

    # --- ZONE STATISTIQUE INTERACTIVE ---
    st.divider()
    st.subheader("🎯 Indicateurs de Performance Consolidés (Vue Direction)")
    
    c1, c2, c3, c4 = st.columns(4)
    nb_missions = len(conso_mis) if not conso_mis.empty else 0
    nb_anomalies = len(conso_anom) if not conso_anom.empty else 0
    
    nb_critiques_actives = 0
    if not conso_anom.empty and "Niveau criticité" in conso_anom.columns and "Statut" in conso_anom.columns:
        nb_critiques_actives = len(conso_anom[
            (conso_anom["Niveau criticité"].str.contains("Critique", na=False)) & 
            (conso_anom["Statut"].str.contains("Ouvert|En cours", na=False, case=False))
        ])
        
    perte_financiere = 0
    if not conso_anom.empty and "Impact estimé (FCFA)" in conso_anom.columns:
        perte_financiere = pd.to_numeric(conso_anom["Impact estimé (FCFA)"], errors='coerce').sum()

    c1.metric("Missions enregistrées", nb_missions)
    c2.metric("Anomalies relevées", nb_anomalies)
    c3.metric("Urgences Critiques", nb_critiques_actives)
    c4.metric("Impact Financier Global", f"{perte_financiere:,.0f} FCFA")

    # --- GRAPHIQUES EXÉCUTIFS ---
    g1, g2 = st.columns(2)
    with g1:
        if not conso_anom.empty and "Niveau criticité" in conso_anom.columns:
            st.markdown("**Gravité des risques identifiés**")
            df_g1 = conso_anom["Niveau criticité"].value_counts().reset_index()
            df_g1.columns = ["Criticité", "Volume"]
            fig1 = px.pie(df_g1, values="Volume", names="Criticité", hole=0.4,
                          color_discrete_map={"🔴 Critique":"#D32F2F", "🟠 Majeur":"#F57C00", "🟡 Mineur":"#FBC02D", "🟢 Faible":"#388E3C"})
            st.plotly_chart(fig1, use_container_width=True)
            
    with g2:
        if not conso_anom.empty and "Pays" in conso_anom.columns:
            st.markdown("**Répartition géographique des anomalies**")
            df_g2 = conso_anom["Pays"].value_counts().reset_index()
            df_g2.columns = ["Pays", "Anomalies"]
            fig2 = px.bar(df_g2, x="Pays", y="Anomalies", text_auto=True, color="Anomalies", color_continuous_scale="Oranges")
            st.plotly_chart(fig2, use_container_width=True)

    # --- COMPILATION ET EXPORT DE DONNÉES STRUCTURÉES ---
    st.divider()
    st.subheader("📥 Génération des Données Consolidées")
    st.info("Le système génère un classeur structuré prêt à être injecté ou lu par votre Tableau de bord Excel Principal.")
    
    if st.button("⚙️ Lancer la compilation finale des données"):
        buffer_excel = BytesIO()
        with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
            # Structuration automatique conforme aux onglets attendus
            if not conso_mis.empty:
                conso_mis.to_excel(writer, sheet_name="CONSO_MISSIONS", index=False, startrow=4)
            if not conso_pts.empty:
                conso_pts.to_excel(writer, sheet_name="CONSO_POINTS_CONTROLE", index=False, startrow=4)
            if not conso_anom.empty:
                conso_anom.to_excel(writer, sheet_name="CONSO_ANOMALIES", index=False, startrow=4)
            if not conso_pln.empty:
                conso_pln.to_excel(writer, sheet_name="CONSO_PLANS_ACTION", index=False, startrow=4)
            
            df_journal.to_excel(writer, sheet_name="LOG_CONSOLIDATION", index=False)

        donnees_finales = buffer_excel.getvalue()
        
        st.download_button(
            label="💾 Télécharger les Données Consolidées Groupe (.xlsx)",
            data=donnees_finales,
            file_name="SKAB_DATA_CONSOLIDATION_GROUPE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.warning("📥 En attente de fichiers d'audit pour générer la synthèse et l'export.")
