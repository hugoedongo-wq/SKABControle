import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import numpy as np
from datetime import datetime

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="SKAB - Dashboard de Consolidation CI",
    page_icon="🛡️",
    layout="wide"
)

# Style CSS pour uniformiser l'affichage
st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 28px; }
    .stAlert { margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)

# --- MOTEUR DE TRAITEMENT ET DE LECTURE DES ONGLETS ---
def load_and_clean(file, sheet):
    """
    Lit un onglet spécifique. Détecte la vraie ligne d'en-tête (Header) 
    pour éviter les erreurs dues aux lignes de titre ou décoratives.
    """
    try:
        # Lecture globale de l'onglet
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
        if df_raw.empty:
            return pd.DataFrame()
            
        # Recherche de la ligne qui contient les vrais titres de colonnes
        # On cherche par exemple 'ID Mission' ou 'ID Anomalie' ou 'ID Point'
        header_idx = 0
        for idx, row in df_raw.iterrows():
            row_str = row.astype(str).values
            if any('ID ' in s or 'N° ' in s or 'Date' in s for s in row_str):
                header_idx = idx
                break
                
        # Re-lecture propre avec le bon en-tête
        df = pd.read_excel(file, sheet_name=sheet, skiprows=header_idx)
        df = df.dropna(how='all')
        
        if not df.empty:
            df['Fichier Source'] = file.name
        return df
    except:
        return pd.DataFrame()

def process_consolidation(files):
    """
    Parcourt l'ensemble des fichiers reçus et assemble les blocs.
    """
    all_data = {"MISSIONS": [], "POINTS": [], "ANOMALIES": [], "PLANS": []}
    
    for f in files:
        f.seek(0)
        all_data["MISSIONS"].append(load_and_clean(f, "MES_MISSIONS"))
        f.seek(0)
        all_data["POINTS"].append(load_and_clean(f, "POINTS_CONTROLE"))
        f.seek(0)
        all_data["ANOMALIES"].append(load_and_clean(f, "ANOMALIES"))
        f.seek(0)
        all_data["PLANS"].append(load_and_clean(f, "PLANS_ACTION"))

    # Fusion (concaténation) de toutes les listes en DataFrames uniques
    combined = {k: pd.concat(v, ignore_index=True) if v else pd.DataFrame() for k, v in all_data.items()}
    return combined

def get_safe_len(series, col_name):
    """Calcule proprement la largeur de colonne pour l'export Excel"""
    clean = series.dropna()
    if not clean.empty:
        max_c = int(clean.apply(lambda x: len(str(x))).max())
    else:
        max_c = 0
    return min(max(max_c, len(str(col_name))) + 3, 50)

# --- INTERFACE UTILISATEUR (STREAMLIT) ---
st.title("🛡️ Espace Chef de Département CI — Groupe SKAB")
st.subheader("Pilotage, Validation Métier et Consolidation des Missions 2026")

# 1. BARRE LATÉRALE - CHARGEMENT MULTIPLE
with st.sidebar:
    st.header("📥 Importation Terrain")
    uploaded_files = st.file_uploader(
        "Déposez les fichiers des contrôleurs (.xlsx) :", 
        type="xlsx", 
        accept_multiple_files=True
    )
    st.divider()
    st.caption("Direction de l'Audit & Contrôle Interne")
    st.caption("© 2026 Groupe SKAB Nutrition")

if not uploaded_files:
    st.info("👋 En attente des fichiers de contrôle des filiales pour initialiser le Dashboard de supervision.")
    st.stop()

# 2. APPEL DU MOTEUR DE FUSION
data = process_consolidation(uploaded_files)
df_mis = data["MISSIONS"]
df_pts = data["POINTS"]
df_anom = data["ANOMALIES"]

# --- DÉTECTION DYNAMIQUE DES COLONNES DU TEMPLATE SKAB ---
# Pour éviter les futures KeyError si un espace s'ajoute, on cherche par mots-clés
col_impact = next((c for c in df_anom.columns if 'Impact' in c), None)
col_crit = next((c for c in df_anom.columns if 'critic' in c or 'Critic' in c), None)
col_domaine = next((c for c in df_anom.columns if 'Domaine' in c or 'domaine' in c), None)
col_pays = next((c for c in df_anom.columns if 'Pays' in c), None)
col_tx_conf = next((c for c in df_mis.columns if 'conform' in c or 'Conform' in c), None)

# 3. AFFICHAGE DES INDICATEURS CLÉS (KPI)
st.markdown("### 📊 Indicateurs de Risques Pré-Consolidation")
k1, k2, k3, k4 = st.columns(4)

with k1:
    impact_total = 0
    if col_impact and not df_anom.empty:
        impact_total = pd.to_numeric(df_anom[col_impact], errors='coerce').fillna(0).sum()
    st.metric("Risque Financier Cumulé", f"{impact_total:,.0f} FCFA")

with k2:
    nb_critiques = 0
    if col_crit and not df_anom.empty:
        nb_critiques = df_anom[df_anom[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
    st.metric("Anomalies Critiques", nb_critiques, delta="Action urgente" if nb_critiques > 0 else None, delta_color="inverse")

with k3:
    conformite_moyenne = 0
    if col_tx_conf and not df_mis.empty:
        # Si le taux est saisi en pourcentage (ex: 75 au lieu de 0.75), on redresse le calcul
        raw_mean = pd.to_numeric(df_mis[col_tx_conf], errors='coerce').mean()
        conformite_moyenne = raw_mean * 100 if raw_mean <= 1.0 else raw_mean
    st.metric("Taux de Conformité Moyen", f"{conformite_moyenne:.1f}%")

with k4:
    st.metric("Classeurs à Fusionner", len(uploaded_files))

st.divider()

# 4. ANALYSES ET GRAPHES INTERACTIFS
g1, g2 = st.columns(2)

with g1:
    st.markdown("**🔍 Volume d'Anomalies par Domaine**")
    if not df_anom.empty and col_domaine:
        color_opt = {'🔴 Critique':'#FF4B4B', '🟠 Majeur':'#FFA500', '🟡 Mineur':'#FFD700', '🟢 Faible':'#2ECC71'}
        # Sécurité si la colonne criticité existe pour la légende colorée
        color_col = col_crit if col_crit else None
        fig = px.bar(df_anom, x=col_domaine, color=color_col, barmode='group', color_discrete_map=color_opt)
        fig.update_layout(height=350, margin=dict(l=0, r=0, t=20, b=0), xaxis_title=None, yaxis_title="Nombre d'écarts")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Pas de données graphiques disponibles pour les anomalies.")

with g2:
    st.markdown("**🌍 Cartographie des Alertes par Pays / Entité**")
    if not df_anom.empty and col_pays:
        df_pays_summary = df_anom.groupby(col_pays).size().reset_index(name="Nombre d'Anomalies")
        fig = px.pie(df_pays_summary, values="Nombre d'Anomalies", names=col_pays, hole=.4, color_discrete_sequence=px.colors.qualitative.Safe)
        fig.update_layout(height=350, margin=dict(l=0, r=0, t=20, b=0))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Aucun découpage géographique disponible.")

# 5. CONTRÔLE QUALITÉ (SAS DE VALIDATION)
st.divider()
st.markdown("### 🛠️ Diagnostic Qualité des Données Reçues")
q1, q2 = st.columns([1, 1])

with q1:
    st.markdown("**🚨 Incohérences détectées (À corriger par les contrôleurs)**")
    alertes_qualite = []
    
    # Validation 1 : Vérification des impacts financiers oubliés sur les anomalies lourdes
    if not df_anom.empty and col_crit and col_impact:
        lignes_anormales = df_anom[
            (df_anom[col_crit].astype(str).str.contains('Critique|🔴|Majeur|🟠', na=False)) & 
            (pd.to_numeric(df_anom[col_impact], errors='coerce').fillna(0) == 0)
        ]
        for _, row in lignes_anormales.iterrows():
            alertes_qualite.append(f"⚠️ **{row['Fichier Source']}** : Une anomalie d'ordre [{row[col_crit]}] a un impact financier à 0 ou non renseigné.")

    # Validation 2 : Croisement orphelin Missions <=> Points de contrôle
    col_num_mis_m = next((c for c in df_mis.columns if 'Mission' in c or 'N°' in c), None)
    col_num_mis_p = next((c for c in df_pts.columns if 'Mission' in c or 'N°' in c), None)
    
    if col_num_mis_m and not df_mis.empty:
        for m_id in df_mis[col_num_mis_m].dropna().unique():
            if df_pts.empty or col_num_mis_p not in df_pts.columns or m_id not in df_pts[col_num_mis_p].values:
                alertes_qualite.append(f"❌ La mission **{m_id}** est listée dans le journal mais n'a aucun point de contrôle rattaché.")

    if alertes_qualite:
        for alerte in alertes_qualite:
            st.warning(alerte)
    else:
        st.success("✅ Diagnostic validé : Les structures et chiffrages des fichiers importés sont cohérents.")

with q2:
    st.markdown("**📌 Top 5 des Alertes Majeures à transmettre au DAF**")
    if not df_anom.empty:
        cols_to_show = [c for c in [col_crit, col_pays, 'Description', col_impact] if c is not None]
        # Tri décroissant sur l'impact financier si disponible
        sort_col = col_impact if col_impact else df_anom.columns[0]
        df_top = df_anom.sort_values(by=sort_col, ascending=False).head(5)
        st.dataframe(df_top[cols_to_show], hide_index=True, use_container_width=True)
    else:
        st.info("Aucune anomalie à lister.")

# 6. ACTION DE SCELLÉ ET COMPILATION DU FICHIER MAÎTRE
st.divider()
st.header("📤 Clôture de Session & Génération")
st.write("Après vérification visuelle des anomalies et des graphiques, cliquez ci-dessous pour matérialiser le Fichier Maître.")

if st.button("🏗️ Compiler et figer le Fichier Maître Consolidé", type="primary", use_container_width=True):
    output_buffer = io.BytesIO()
    
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        # Onglet d'en-tête pour Monsieur le DAF
        df_meta = pd.DataFrame({
            "RAPPORT GLOBAL CI SKAB NUTRITION": ["Destinataire Principal", "Généré par", "Horodatage de fusion", "Volume d'aspiration", "Niveau de validation"],
            "MÉTADONNÉES DE LIVRAISON": ["M. Élie DIGNOU (DAF)", "Chef de Département Contrôle Interne", datetime.now().strftime("%d/%m/%Y à %H:%M"), f"{len(uploaded_files)} Rapports filiales intégrés", "VÉRIFIÉ ET SCELLÉ"]
        })
        df_meta.to_excel(writer, sheet_name="ACCUEIL_CONSO", index=False)
        writer.sheets["ACCUEIL_CONSO"].set_column('A:B', 35)
        
        # Injection des onglets de consolidation métiers
        onglets_export = {
            "CONSO_MISSIONS": data["MISSIONS"],
            "CONSO_POINTS_CONTROLE": data["POINTS"],
            "CONSO_ANOMALIES": data["ANOMALIES"],
            "CONSO_PLANS_ACTION": data["PLANS"]
        }
        
        for sheet_name, dataframe in onglets_export.items():
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            # Ajustement dynamique et robuste de la taille des cellules
            for i, column_name in enumerate(dataframe.columns):
                ws.set_column(i, i, get_safe_len(dataframe[column_name], column_name))
                
    st.success("🎉 Le Fichier Maître unique vient d'être structuré en mémoire de l'instance Cloud !")
    
    # Bouton de téléchargement local pour envoi
    date_fichier = datetime.now().strftime("%Y%m%d")
    st.download_button(
        label=f"💾 Récupérer le Fichier MAÎTRE : SKAB_MAITRE_CONSO_{date_fichier}.xlsx",
        data=output_buffer.getvalue(),
        file_name=f"SKAB_MAITRE_CONSO_{date_fichier}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("---")
st.caption("Direction Générale SKAB Nutrition — Application d'interfaçage d'Audit Interne hébergée sur Streamlit Cloud.")
