import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import numpy as np
from datetime import datetime

# --- CONFIGURATION ---
st.set_page_config(
    page_title="SKAB - Dashboard de Consolidation CI",
    page_icon="🛡️",
    layout="wide"
)

# Style CSS personnalisé pour les métriques
st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 28px; }
    .stAlert { margin-top: 10px; }
    </style>
""", unsafe_allow_html=True)

# --- MOTEUR DE TRAITEMENT ---
def load_and_clean(file, sheet, skip=4):
    try:
        df = pd.read_excel(file, sheet_name=sheet, skiprows=skip)
        df = df.dropna(how='all')
        if not df.empty:
            df['Fichier Source'] = file.name
        return df
    except:
        return pd.DataFrame()

def process_consolidation(files):
    all_data = {"MISSIONS": [], "POINTS": [], "ANOMALIES": [], "PLANS": []}
    
    for f in files:
        # On rembobine le fichier pour chaque lecture d'onglet
        f.seek(0)
        all_data["MISSIONS"].append(load_and_clean(f, "MES_MISSIONS"))
        f.seek(0)
        all_data["POINTS"].append(load_and_clean(f, "POINTS_CONTROLE"))
        f.seek(0)
        all_data["ANOMALIES"].append(load_and_clean(f, "ANOMALIES"))
        f.seek(0)
        all_data["PLANS"].append(load_and_clean(f, "PLANS_ACTION"))

    # Concaténation
    combined = {k: pd.concat(v, ignore_index=True) if v else pd.DataFrame() for k, v in all_data.items()}
    return combined

def get_safe_len(series, col_name):
    clean = series.dropna()
    if not clean.empty:
        max_c = int(clean.apply(lambda x: len(str(x))).max())
    else:
        max_c = 0
    return min(max(max_c, len(str(col_name))) + 3, 50)

# --- INTERFACE ---
st.title("🛡️ Espace Chef de Département CI")
st.subheader("Pilotage, Validation et Consolidation des Missions 2026")

# 1. CHARGEMENT
with st.sidebar:
    st.header("📥 Importation")
    uploaded_files = st.file_uploader(
        "Fichiers Contrôleurs (.xlsx)", 
        type="xlsx", 
        accept_multiple_files=True
    )
    st.divider()
    st.caption("© 2026 Groupe SKAB Nutrition")

if not uploaded_files:
    st.info("👋 Veuillez charger les fichiers des contrôleurs dans la barre latérale pour activer le dashboard.")
    st.stop()

# 2. TRAITEMENT
data = process_consolidation(uploaded_files)
df_anom = data["ANOMALIES"]
df_mis = data["MISSIONS"]
df_pts = data["POINTS"]

# 3. DASHBOARD PRÉ-CONSOLIDATION
# --- RANGÉE 1 : KPI FLASH ---
st.markdown("### 📊 État Global du Risque (Données en attente)")
k1, k2, k3, k4 = st.columns(4)

with k1:
    impact = 0
    if not df_anom.empty and 'Impact (FCFA)' in df_anom.columns:
        impact = pd.to_numeric(df_anom['Impact (FCFA)'], errors='coerce').sum()
    st.metric("Risque Financier Brut", f"{impact:,.0f} FCFA")

with k2:
    critiques = 0
    if not df_anom.empty and 'Niveau criticité' in df_anom.columns:
        critiques = df_anom[df_anom['Niveau criticité'].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
    st.metric("Anomalies Critiques", critiques, delta="Action Immédiate" if critiques > 0 else None, delta_color="inverse")

with k3:
    taux = 0
    if not df_mis.empty and 'Taux conformité' in df_mis.columns:
        taux = pd.to_numeric(df_mis['Taux conformité'], errors='coerce').mean() * 100
    st.metric("Conformité Moyenne", f"{taux:.1f}%")

with k4:
    st.metric("Fichiers Reçus", len(uploaded_files))

st.divider()

# --- RANGÉE 2 : ANALYSES GRAPHIQUES ---
g1, g2 = st.columns(2)

with g1:
    st.markdown("**🔍 Répartition des Risques par Domaine**")
    if not df_anom.empty:
        # Utiliser 'Type / Domaine' ou 'Domaine' selon votre template
        dom_col = 'Type / Domaine' if 'Type / Domaine' in df_anom.columns else 'Domaine'
        if dom_col in df_anom.columns:
            fig = px.bar(df_anom, x=dom_col, color='Niveau criticité', 
                         color_discrete_map={'🔴 Critique':'#FF4B4B', '🟠 Majeur':'#FFA500'},
                         barmode='group')
            fig.update_layout(height=350, margin=dict(l=0, r=0, t=20, b=0))
            st.plotly_chart(fig, use_container_width=True)

with g2:
    st.markdown("**🌍 Analyse Comparative par Pays**")
    if not df_anom.empty and 'Pays' in df_anom.columns:
        df_pays = df_anom.groupby('Pays').size().reset_index(name='Nombre')
        fig = px.pie(df_pays, values='Nombre', names='Pays', hole=.4, color_discrete_sequence=px.colors.qualitative.Pastel)
        fig.update_layout(height=350, margin=dict(l=0, r=0, t=20, b=0))
        st.plotly_chart(fig, use_container_width=True)

# --- RANGÉE 3 : QUALITÉ DES DONNÉES & ALERTES ---
st.divider()
st.markdown("### 🛠️ Sas de Validation (Contrôle Qualité)")

q1, q2 = st.columns([1, 1])

with q1:
    st.markdown("**🔴 Incohérences de saisie détectées**")
    errors = []
    
    # Check 1: Anomalies critiques sans impact financier
    if not df_anom.empty:
        crit_no_impact = df_anom[
            (df_anom['Niveau criticité'].astype(str).str.contains('Critique', na=False)) & 
            (pd.to_numeric(df_anom['Impact (FCFA)'], errors='coerce').fillna(0) == 0)
        ]
        for idx, row in crit_no_impact.iterrows():
            errors.append(f"⚠️ **{row['Fichier Source']}** : Anomalie Critique (ID {idx}) sans impact financier chiffré.")

    # Check 2: Missions sans points de contrôle
    if not df_mis.empty:
        for m in df_mis['N° Mission'].unique():
            if df_pts.empty or m not in df_pts['N° Mission'].values:
                errors.append(f"❌ Mission **{m}** déclarée mais aucun point de contrôle saisi.")

    if errors:
        for e in errors: st.warning(e)
    else:
        st.success("✅ Aucune incohérence majeure détectée. Les fichiers sont propres.")

with q2:
    st.markdown("**📌 Top 5 des Anomalies à Arbitrer (DAF)**")
    if not df_anom.empty:
        df_prior = df_anom.sort_values(by=['Impact (FCFA)'], ascending=False).head(5)
        st.dataframe(df_prior[['Niveau criticité', 'Pays', 'Description', 'Impact (FCFA)']], hide_index=True)

# --- RANGÉE 4 : GÉNÉRATION DU LIVRABLE ---
st.divider()
st.header("📤 Finalisation du Livrable")
st.write("Si les indicateurs ci-dessus sont cohérents, vous pouvez générer le fichier Maître Consolidé.")

if st.button("🏗️ Générer et Télécharger le Fichier Maître Consolidé", type="primary", use_container_width=True):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Page de garde
        df_info = pd.DataFrame({
            "METADONNEES": ["Généré par", "Date", "Source", "Statut"],
            "VALEURS": ["Chef CI (App Streamlit)", datetime.now().strftime("%d/%m/%Y %H:%M"), f"{len(uploaded_files)} fichiers", "Consolidé & Validé"]
        })
        df_info.to_excel(writer, sheet_name="INFO_CONSO", index=False)
        
        # Onglets consolidés
        sheets = {
            "CONSO_MISSIONS": data["MISSIONS"],
            "CONSO_POINTS": data["POINTS"],
            "CONSO_ANOMALIES": data["ANOMALIES"],
            "CONSO_PLANS": data["PLANS"]
        }
        
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            for i, col in enumerate(df.columns):
                ws.set_column(i, i, get_safe_len(df[col], col))
                
    st.success("✨ Consolidation terminée avec succès !")
    st.download_button(
        label="💾 Télécharger SKAB_MAITRE_CONSOLIDE.xlsx",
        data=output.getvalue(),
        file_name=f"SKAB_MAITRE_CONSO_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("---")
st.caption("Direction de l'Audit et du Contrôle Interne - Groupe SKAB Nutrition")
