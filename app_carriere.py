import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import io
import math
import re
from uuid import uuid4
import html
from typing import Optional
from pathlib import Path
from datetime import datetime
import streamlit.components.v1 as components
import auth
import projects
import importlib
projects = importlib.reload(projects)

# Bump this value to invalidate Streamlit cache when cleaning rules change.
DATA_CLEAN_VERSION = "2026-03-02-01"

def normalize_text_series(series: pd.Series, kind: Optional[str] = None) -> pd.Series:
    s = series.astype("string")
    s = s.str.normalize("NFKC")
    s = s.str.replace(r"[\u0000-\u001F\u007F-\u009F]", "", regex=True)  # control chars
    s = s.str.replace(
        r"[\u200B-\u200F\u202A-\u202E\u2060-\u2069\u061C\uFEFF]",
        "",
        regex=True,
    )  # invisible/format chars (bidi, zero-width, etc.)
    s = s.str.replace("\u00A0", " ", regex=False)  # NBSP
    s = s.str.replace("\u202F", " ", regex=False)  # narrow NBSP
    s = s.str.replace(r"[’‘´`ʹʻ]", "'", regex=True)  # apostrophes
    s = s.str.replace(r"[‐‑‒–—―−﹣－]", "-", regex=True)  # hyphens/minus

    if kind == "client":
        # Harmonise common separators so the same client doesn't appear twice
        # (e.g. "RB TRAVAUX" vs "RB-TRAVAUX", "ITB/RDP" vs "ITBRDP").
        s = s.str.replace(r"[-/_]", " ", regex=True)
        s = s.str.replace(r"[.,]", " ", regex=True)
        s = s.str.replace(r"\bINDUTRIES\b", "INDUSTRIES", regex=True)

    if kind == "mois":
        # Standardise months labels (including common mojibake seen in some Excels)
        s = s.str.casefold()
        month_map = {
            "<na>": pd.NA,
            "nan": pd.NA,
            "janvier": "janvier",
            "février": "février",
            "fevrier": "février",
            "fã©vrier": "février",
            "fÃ©vrier": "février",
            "mars": "mars",
            "avril": "avril",
            "mai": "mai",
            "juin": "juin",
            "juillet": "juillet",
            "août": "août",
            "aout": "août",
            "aoã»t": "août",
            "aoÃ»t": "août",
            "septembre": "septembre",
            "octobre": "octobre",
            "novembre": "novembre",
            "décembre": "décembre",
            "decembre": "décembre",
            "dã©cembre": "décembre",
            "dÃ©cembre": "décembre",
        }
        s = s.replace(month_map)

    s = s.str.replace(r"\s+", " ", regex=True).str.strip()

    if kind == "client":
        # Canonicalise remaining variants by picking the most frequent label per key
        # (key ignores separators/spaces/case).
        key = s.fillna("").str.replace(r"[\s\-_/\.,']", "", regex=True).str.casefold()
        tmp = pd.DataFrame({"key": key, "val": s})
        tmp = tmp.dropna()
        tmp = tmp[tmp["val"].astype(str).str.strip().ne("")]
        if not tmp.empty:
            counts = tmp.groupby(["key", "val"]).size().reset_index(name="n")
            counts = counts.sort_values(["key", "n", "val"], ascending=[True, False, True])
            rep = counts.drop_duplicates(subset=["key"], keep="first").set_index("key")["val"].to_dict()
            canonical = key.map(rep)
            s = canonical.where(canonical.notna(), s)

    return s

def safe_progress_max(value: object) -> float:
    try:
        v = float(value)
    except Exception:
        return 1.0
    if not math.isfinite(v) or v <= 0:
        return 1.0
    return v


PLOTLY_FONT_FAMILY = 'Inter, Roboto, system-ui, -apple-system, "Segoe UI", Arial, sans-serif'


def apply_plotly_style(fig, *, kind: str = "default"):
    try:
        fig.update_layout(
            template="plotly_white",
            font=dict(family=PLOTLY_FONT_FAMILY, size=13, color="#111827"),
            title_font_size=16,
            legend=dict(font=dict(size=12)),
            hoverlabel=dict(font_size=13),
        )
        if kind == "pie":
            fig.update_layout(
                height=380,
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.06,
                    xanchor="left",
                    x=0,
                    font=dict(size=11),
                ),
            )
        fig.update_xaxes(title_font=dict(size=13), tickfont=dict(size=11))
        fig.update_yaxes(title_font=dict(size=13), tickfont=dict(size=11))
    except Exception:
        return fig
    return fig

THEMES = [
    {
        "name": "Orange",
        "primary": "#FF8C00",
        "secondary": "#FFA07A",
        "accent": "#FF4500",
        "pie": ["#FF8C00", "#FFB347", "#FFA07A", "#FF4500", "#FFD1A1", "#FF6F61"],
        "qualitative": ["#FF8C00", "#FF4500", "#FFA07A", "#FFB347", "#6A5ACD", "#2E8B57", "#1E90FF"],
        "table_cmap": "YlOrRd",
    },
    {
        "name": "Bleu",
        "primary": "#1f77b4",
        "secondary": "#17becf",
        "accent": "#2ca02c",
        "pie": ["#1f77b4", "#17becf", "#2ca02c", "#9467bd", "#8c564b", "#7f7f7f"],
        "qualitative": ["#1f77b4", "#17becf", "#2ca02c", "#9467bd", "#e377c2", "#8c564b", "#7f7f7f"],
        "table_cmap": "PuBu",
    },
]

# Configuration de la page
st.set_page_config(
    page_title="Tableau de Bord - Carrière",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Auth ---
APP_DIR = Path(__file__).resolve().parent
AUTH_DB_PATH = APP_DIR / "data" / "users.sqlite3"
auth.init_db(AUTH_DB_PATH)

if "auth_user" not in st.session_state:
    st.session_state.auth_user = None

with st.sidebar.expander("Compte", expanded=True):
    user = st.session_state.auth_user
    if user is not None:
        st.write(f"Connecté: **{user.username}**")
        st.caption(f"Rôle: {user.role}")
        if st.button("Déconnexion", use_container_width=True):
            st.session_state.auth_user = None
            st.rerun()
    else:
        auth_mode = st.radio("Accès", ["Connexion", "Créer un compte"], horizontal=True)
        if auth_mode == "Connexion":
            with st.form("login_form", clear_on_submit=False):
                login_username = st.text_input("Nom d'utilisateur", key="login_username")
                login_password = st.text_input("Mot de passe", type="password", key="login_password")
                login_submit = st.form_submit_button("Se connecter", use_container_width=True)
            if login_submit:
                try:
                    st.session_state.auth_user = auth.authenticate(AUTH_DB_PATH, login_username, login_password)
                    st.rerun()
                except auth.AuthError as e:
                    st.error(str(e))
        else:
            with st.form("signup_form", clear_on_submit=False):
                signup_username = st.text_input("Nom d'utilisateur", key="signup_username")
                signup_password = st.text_input("Mot de passe (min 8)", type="password", key="signup_password")
                signup_password2 = st.text_input("Confirmer", type="password", key="signup_password2")
                signup_submit = st.form_submit_button("Créer le compte", use_container_width=True)
            if signup_submit:
                if signup_password != signup_password2:
                    st.error("Les mots de passe ne correspondent pas.")
                else:
                    try:
                        st.session_state.auth_user = auth.create_user(AUTH_DB_PATH, signup_username, signup_password)
                        st.rerun()
                    except auth.AuthError as e:
                        st.error(str(e))

if st.session_state.auth_user is None:
    components.html(
        """
        <script>
        try {
          const btn = window.parent?.document?.getElementById("pdf-gen-btn");
          if (btn) btn.remove();
        } catch (e) {}
        </script>
        """,
        height=0,
    )
    st.title(" Tableau de Bord - Suivi des Livraisons Carrière")
    st.info("Veuillez vous connecter (ou créer un compte) dans la barre latérale.")
    st.stop()

# Bandeau (visible par tous les utilisateurs connectés)
st.markdown(
    """
    <div style="
        background:#F2F2F2;
        border:1px solid #E0E0E0;
        padding:10px 12px;
        border-radius:8px;
        font-weight:700;
        margin: 4px 0 10px 0;
        display:flex;
        align-items:center;
        justify-content:space-between;
        gap:12px;
    ">
      <div id="top-banner" style="display:flex; align-items:center; justify-content:space-between; gap:12px; width:100%;">
        <div>Gloria GBAYE - CARRIDEC</div>
        <div id="top-banner-actions"></div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- Design system (lisible à 100% de zoom, desktop-first) ---
st.markdown(
    """
    <style>
      :root {
        --bg: #fafafa;
        --surface: #ffffff;
        --text: #111827;
        --muted: #475467;
        --border: #e5e7eb;
        --shadow1: 0 1px 2px rgba(16,24,40,.06);
        --shadow2: 0 6px 18px rgba(16,24,40,.10);
        --radius: 12px;
        --accent: #1e3a8a;
        --success: #10b981;
        --warning: #f59e0b;
      }

      html { font-size: 15px; }
      body, .stApp {
        background: var(--bg);
        color: var(--text);
        font-family: Inter, Roboto, system-ui, -apple-system, "Segoe UI", Arial, sans-serif;
        line-height: 1.45;
      }
      *, *::before, *::after { box-sizing: border-box; }

      /* Main container: max width + comfortable padding */
      div.block-container {
        max-width: 1400px;
        padding-left: 24px;
        padding-right: 24px;
      }
      @media (max-width: 1366px) {
        div.block-container { padding-left: 16px; padding-right: 16px; }
      }

      /* Typography */
      h1 { font-size: 1.6rem; line-height: 1.2; }
      h2 { font-size: 1.35rem; }
      h3 { font-size: 1.15rem; }
      p, li, label, .stMarkdown, .stText { font-size: 1rem; }
      small, .stCaption, .stCaptionContainer { font-size: 0.875rem; color: var(--muted); }

      /* Click targets */
      div[data-testid="stButton"] > button,
      div[data-testid="stDownloadButton"] > button,
      div[data-testid="stPopover"] > button {
        min-height: 44px;
        border-radius: var(--radius);
      }
      div[data-testid="stDownloadButton"] > button {
        width: 100%;
        background: var(--success);
        border: 1px solid var(--success);
        color: #ffffff;
        font-weight: 800;
      }
      div[data-testid="stDownloadButton"] > button:hover {
        background: #0ea371;
        border-color: #0ea371;
        color: #ffffff;
      }

      /* Section spacing */
      hr { margin: 18px 0; border-color: var(--border); }

      /* KPI visuals */
      .kpi-title { font-size: 1.25rem; font-weight: 800; color: var(--text); margin-top: 1rem; margin-bottom: .5rem; }
      .kpi-val { font-size: 1.6rem; color: var(--accent); font-weight: 800; }
      .kpi-box {
        background-color: var(--surface);
        padding: 16px;
        border-radius: var(--radius);
        border: 1px solid var(--border);
        box-shadow: var(--shadow1);
        margin-bottom: 12px;
      }

      /* KPI grid (custom, avoids truncation) */
      .kpi-grid {
        display: grid;
        grid-template-columns: repeat(6, minmax(0, 1fr));
        gap: 12px;
        margin: 10px 0 6px 0;
      }
      @media (max-width: 1400px) {
        .kpi-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); }
      }
      .kpi-card {
        background: var(--surface);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 12px 12px 10px 12px;
        box-shadow: var(--shadow1);
        min-width: 0;
      }
      .kpi-card .kpi-label {
        font-size: 0.875rem;
        color: var(--muted);
        font-weight: 700;
      }
      .kpi-card .kpi-value {
        margin-top: 6px;
        font-size: 1.15rem;
        line-height: 1.15;
        font-weight: 900;
        color: var(--text);
        white-space: normal;
        overflow-wrap: anywhere;
      }

      /* DataFrames: keep inside container (no page-level horizontal scroll) */
      div[data-testid="stDataFrame"] { max-width: 100%; width: 100%; }
      div[data-testid="stDataFrame"] > div { max-width: 100%; width: 100%; overflow-x: auto; }
      div[data-testid="stDataFrame"] [role="gridcell"] > div {
        white-space: normal !important;
        overflow-wrap: anywhere;
        line-height: 1.2;
      }

      /* Screen vs print helpers */
      .print-only { display: none; }
      .screen-only { display: inline; }
      @media print {
        .print-only { display: block !important; }
        .screen-only { display: none !important; }
        .kpi-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); }
        div[data-testid="stDataFrame"] > div { overflow: visible !important; }
      }

      /* Print-friendly tables */
      .print-table { margin-top: 10px; }
      .print-table table { width: 100%; border-collapse: collapse; table-layout: fixed; }
      .print-table th, .print-table td { border: 1px solid #d1d5db; padding: 4px 6px; font-size: 11px; }
      .print-table th { background: #f3f4f6; font-weight: 800; }
      .print-table td { word-break: break-word; }

      /* Metric cards */
      div[data-testid="metric-container"] {
        background-color: var(--surface);
        border: 1px solid var(--border);
        border-left: 4px solid var(--accent);
        padding: 12px 12px 12px 14px;
        border-radius: var(--radius);
        box-shadow: var(--shadow1);
      }
      div[data-testid="metric-container"] > label { color: var(--muted); font-weight: 700; }
      div[data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: var(--text);
        font-weight: 800;
        font-size: 1.1rem;
        line-height: 1.15;
        white-space: normal !important;
        overflow-wrap: anywhere;
      }

      /* Avoid accidental horizontal scroll */
      .stApp { overflow-x: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Fonction pour générer le fichier Excel en mémoire
@st.cache_data
def generate_excel_download(df, stats_df, mapping):
    output = io.BytesIO()
    produit_col = mapping.get('produit')
    client_col = mapping.get('client')
    poids_col = mapping.get('poids')
    ca_col = mapping.get('ca')
    date_col = mapping.get('date')
    mois_col = mapping.get('mois')

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # --- FORMATS ---
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF8C00', 'font_color': 'white', 'border': 1, 'align': 'center'})
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#FF8C00'})
        kpi_label_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFF5EC', 'font_color': '#666666', 'border': 1})
        kpi_val_fmt = workbook.add_format({'bold': True, 'bg_color': '#FFF5EC', 'font_color': '#FF8C00', 'font_size': 12, 'border': 1, 'num_format': '#,##0'})
        num_fmt = workbook.add_format({'num_format': '#,##0.00'})
        money_fmt = workbook.add_format({'num_format': '#,##0 "€"'})
        
        # 1. --- ONGLET SYNTHESE ---
        ws_synth = workbook.add_worksheet(' Synthèse Dashboard')
        ws_synth.write(0, 0, "TABLEAU DE BORD - SYNTHÈSE DES LIVRAISONS", title_fmt)
        
        # Layout des KPIs (2 lignes, 3 colonnes)
        for i, row in stats_df.iterrows():
            r = 2 + (i // 3) * 2
            c = (i % 3) * 2
            ws_synth.write(r, c, row['Indicateur'], kpi_label_fmt)
            ws_synth.write(r + 1, c, row['Valeur'], kpi_val_fmt)
            ws_synth.set_column(c, c, 20)
            
        # 2. --- ONGLET ANALYSE PRODUITS ---
        if produit_col:
            df_prod = df.groupby(produit_col).agg({poids_col: 'sum', ca_col: 'sum'}).reset_index()
            df_prod = df_prod.sort_values(by=poids_col, ascending=False)
            df_prod.to_excel(writer, sheet_name=' Analyse Produits', index=False, startrow=2)
            ws = writer.sheets[' Analyse Produits']
            ws.write(0, 0, "ANALYSE PAR PRODUIT (Tonnage & CA)", title_fmt)
            for col_num, value in enumerate(df_prod.columns.values):
                ws.write(2, col_num, value, header_fmt)
            
            # Data Bars for Tonnage
            ws.conditional_format(3, 1, 3 + len(df_prod), 1, {
                'type': 'data_bar', 'bar_color': '#FFA07A'
            })
            # Data Bars for CA
            ws.conditional_format(3, 2, 3 + len(df_prod), 2, {
                'type': 'data_bar', 'bar_color': '#FFB347'
            })
            ws.set_column('A:A', 25)
            ws.set_column('B:C', 18, num_fmt)

        # 3. --- ONGLET PERFORMANCE CLIENTS ---
        if client_col:
            df_client = df.groupby(client_col).agg({poids_col: 'sum', ca_col: 'sum'}).reset_index()
            df_client = df_client.sort_values(by=poids_col, ascending=False).head(20)
            df_client.to_excel(writer, sheet_name='👥 Performance Clients', index=False, startrow=2)
            ws = writer.sheets['👥 Performance Clients']
            ws.write(0, 0, "TOP 20 CLIENTS (Tonnage & CA)", title_fmt)
            for col_num, value in enumerate(df_client.columns.values):
                ws.write(2, col_num, value, header_fmt)
            
            # Data Bars
            ws.conditional_format(3, 1, 3 + len(df_client), 1, {'type': 'data_bar', 'bar_color': '#87CEEB'}) # Blue like Streamlit
            ws.conditional_format(3, 2, 3 + len(df_client), 2, {'type': 'data_bar', 'bar_color': '#FF8C00'}) # Orange
            ws.set_column('A:A', 30)
            ws.set_column('B:C', 18, num_fmt)

        # 4. --- ONGLET MATRICE ---
        if client_col and produit_col:
            pivot_df = pd.pivot_table(df, values=poids_col, index=client_col, columns=produit_col, aggfunc='sum', fill_value=0)
            pivot_df['Total'] = pivot_df.sum(axis=1)
            pivot_df = pivot_df.sort_values(by='Total', ascending=False)
            pivot_df.to_excel(writer, sheet_name=' Matrice Client-Produit', startrow=2)
            ws = writer.sheets[' Matrice Client-Produit']
            ws.write(0, 0, "MATRICE DE LIVRAISON (en tonnes)", title_fmt)
            
            # Heatmap color scale
            ws.conditional_format(3, 1, 3 + len(pivot_df), len(pivot_df.columns), {
                'type': '2_color_scale', 'min_color': "#FFFFFF", 'max_color': "#FF4500"
            })
            ws.set_column(0, 0, 30)
            ws.set_column(1, len(pivot_df.columns), 12, num_fmt)

        # 5. --- ONGLET EVOLUTION ---
        if date_col and produit_col:
            months_order = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']
            df_ev = df.copy()
            if mois_col in df_ev.columns:
                df_ev[mois_col] = df_ev[mois_col].str.lower()
                pivot_ev = pd.pivot_table(df_ev, values=poids_col, index=mois_col, columns=produit_col, aggfunc='sum', fill_value=0)
                # Reorder index
                existing_months = [m for m in months_order if m in pivot_ev.index]
                pivot_ev = pivot_ev.reindex(existing_months)
                pivot_ev.to_excel(writer, sheet_name=' Évolution Mensuelle', startrow=2)
                ws = writer.sheets[' Évolution Mensuelle']
                ws.write(0, 0, "ÉVOLUTION DU TONNAGE PAR MOIS ET PRODUIT", title_fmt)
                ws.set_column(0, 0, 15)
                ws.set_column(1, 15, 12, num_fmt)

        # 6. --- DONNEES BRUTES ---
        df.to_excel(writer, sheet_name=' Données Filtrées', index=False)
        ws = writer.sheets[' Données Filtrées']
        ws.freeze_panes(1, 0)
        ws.set_column('A:Z', 15)

    processed_data = output.getvalue()
    return processed_data


# Fonction de chargement et de nettoyage des données
@st.cache_data
def load_data(file_path, clean_version: str = DATA_CLEAN_VERSION, sheet_name=0):
    file_name = file_path if isinstance(file_path, str) else getattr(file_path, "name", "")
    suffix = Path(str(file_name)).suffix.lower()

    def _detect_csv_sep(fp) -> str:
        seps = [",", ";", "\t"]
        sample = ""
        try:
            if isinstance(fp, str):
                with open(fp, "rb") as f:
                    sample = f.read(4096).decode("utf-8", errors="ignore")
            else:
                fp.seek(0)
                sample = fp.read(4096).decode("utf-8", errors="ignore")
                fp.seek(0)
        except Exception:
            return ","
        counts = {sep: sample.count(sep) for sep in seps}
        best = max(counts, key=counts.get)
        return best if counts[best] > 0 else ","

    def _read_table(header, nrows):
        if suffix == ".csv":
            sep = _detect_csv_sep(file_path)
            try:
                if not isinstance(file_path, str):
                    file_path.seek(0)
                return pd.read_csv(file_path, header=header, nrows=nrows, sep=sep, engine="python")
            except UnicodeDecodeError:
                if not isinstance(file_path, str):
                    file_path.seek(0)
                return pd.read_csv(file_path, header=header, nrows=nrows, sep=sep, engine="python", encoding="latin1")
        else:
            try:
                file_path.seek(0)
            except Exception:
                pass
            return pd.read_excel(file_path, header=header, nrows=nrows, sheet_name=sheet_name)

    raw_df = _read_table(header=None, nrows=50)
    header_idx = 0
    for i, row in raw_df.iterrows():
        row_str = " ".join([str(x).lower() for x in row.values])
        if "ticket" in row_str and "client" in row_str:
            header_idx = i
            break
            
    df = _read_table(header=header_idx, nrows=None)
    df.columns = df.columns.astype(str).str.strip()
    
    col_mapping = {}
    for c in df.columns:
        clow = c.lower()
        if "affaire" in clow: col_mapping['ca'] = c
        elif "date" in clow and 'date' not in col_mapping: col_mapping['date'] = c
        elif "net" == clow.strip(): col_mapping['poids'] = c
        elif "client" in clow: col_mapping['client'] = c
        elif "produit" in clow: col_mapping['produit'] = c
        elif "lettres" in clow: col_mapping['mois'] = c
        elif "ticket" in clow: col_mapping['ticket'] = c

    for key in ("client", "produit", "mois"):
        col = col_mapping.get(key)
        if col and col in df.columns:
            df[col] = normalize_text_series(df[col], kind=key)
        
    if 'date' in col_mapping:
        df[col_mapping['date']] = pd.to_datetime(df[col_mapping['date']], errors='coerce')
        df['Année'] = df[col_mapping['date']].dt.year
        df['Trimestre'] = 'T' + df[col_mapping['date']].dt.quarter.astype(str)
        # Add week number
        df['Semaine'] = df[col_mapping['date']].dt.isocalendar().week
        
        months_fr = {
            1: "janvier",
            2: "février",
            3: "mars",
            4: "avril",
            5: "mai",
            6: "juin",
            7: "juillet",
            8: "août",
            9: "septembre",
            10: "octobre",
            11: "novembre",
            12: "décembre",
        }
        mois_calc = df[col_mapping['date']].dt.month.map(months_fr)

        if 'mois' not in col_mapping:
            df['Mois'] = mois_calc
            col_mapping['mois'] = 'Mois'
        else:
            # Force le mois à partir de la date pour éviter les valeurs corrompues
            # (ex: "Dã©cembre" au lieu de "décembre").
            mois_col = col_mapping['mois']
            if mois_col in df.columns:
                has_date = df[col_mapping['date']].notna()
                df.loc[has_date, mois_col] = mois_calc[has_date]
        
        # Jours ouvrés
        df['Jour_de_semaine'] = df[col_mapping['date']].dt.dayofweek
        df['Est_Ouvre'] = df['Jour_de_semaine'] < 5 # 0-4 c'est Lundi-Vendredi
            
    if 'ca' in col_mapping:
        df[col_mapping['ca']] = pd.to_numeric(df[col_mapping['ca']], errors='coerce').fillna(0)
    if 'poids' in col_mapping:
        df[col_mapping['poids']] = pd.to_numeric(df[col_mapping['poids']], errors='coerce').fillna(0)
        
    return df, col_mapping

dossier_app = os.path.dirname(os.path.abspath(__file__))
fichier_defaut = os.path.join(dossier_app, "extraction pont bsacule retraité.xlsx")

st.sidebar.title(" Paramètres")
st.sidebar.markdown("---")
st.sidebar.caption(f"Données: {DATA_CLEAN_VERSION} • Script: {os.path.abspath(__file__)}")

PROJECT_DB_PATH = Path(dossier_app) / "data" / "projects.sqlite3"
PROJECT_FILES_DIR = Path(dossier_app) / "data" / "project_files"
projects.init_db(PROJECT_DB_PATH)

user_id = int(st.session_state.auth_user.id)

if "active_project_id" not in st.session_state:
    st.session_state.active_project_id = None
if "rename_project_id" not in st.session_state:
    st.session_state.rename_project_id = None
if "sidebar_upload_counter" not in st.session_state:
    st.session_state.sidebar_upload_counter = 0
if "top_upload_counter" not in st.session_state:
    st.session_state.top_upload_counter = 0
if "show_top_uploader" not in st.session_state:
    st.session_state.show_top_uploader = False
if "project_search" not in st.session_state:
    st.session_state.project_search = ""


def _safe_stem(filename: str) -> str:
    stem = Path(str(filename or "")).stem.strip() or "extraction"
    stem = re.sub(r"\s+", " ", stem)
    stem = re.sub(r"[^A-Za-z0-9 _.\-]", "", stem).strip()
    return stem[:60] if len(stem) > 60 else stem


def _human_money(value: float) -> str:
    try:
        v = float(value)
    except Exception:
        return "0€"
    if abs(v) >= 1_000_000_000:
        return f"{v/1_000_000_000:,.1f} Md€".replace(",", " ").replace(".", ",")
    if abs(v) >= 1_000_000:
        return f"{v/1_000_000:,.1f} M€".replace(",", " ").replace(".", ",")
    return f"{v:,.0f} €".replace(",", " ")


def _fmt_number_fr(value: float, decimals: int = 0) -> str:
    try:
        v = float(value)
    except Exception:
        v = 0.0
    s = f"{v:,.{decimals}f}"
    s = s.replace(",", " ").replace(".", ",")
    if decimals == 0:
        s = s.split(",")[0]
    return s


def _fmt_eur_full(value: float, decimals: int = 0) -> str:
    return f"{_fmt_number_fr(value, decimals=decimals)} €"


def _kpi_card_html(label: str, screen_value: str, print_value: str, *, tooltip: str) -> str:
    return (
        f"<div class='kpi-card' title='{html.escape(tooltip)}'>"
        f"<div class='kpi-label'>{html.escape(label)}</div>"
        f"<div class='kpi-value'>"
        f"<span class='screen-only'>{html.escape(screen_value)}</span>"
        f"<span class='print-only'>{html.escape(print_value)}</span>"
        f"</div></div>"
    )


def _compute_project_stats(df: pd.DataFrame, mapping: dict) -> dict:
    ca_col = mapping.get("ca")
    poids_col = mapping.get("poids")
    date_col = mapping.get("date")
    ca_total = float(df[ca_col].sum()) if ca_col and ca_col in df.columns else 0.0
    tonnage_total = float(df[poids_col].sum()) if poids_col and poids_col in df.columns else 0.0
    nb_livraisons = int(len(df))

    date_min = None
    date_max = None
    if date_col and date_col in df.columns:
        try:
            dmin = pd.to_datetime(df[date_col], errors="coerce").min()
            dmax = pd.to_datetime(df[date_col], errors="coerce").max()
            if pd.notna(dmin):
                date_min = pd.Timestamp(dmin).date().isoformat()
            if pd.notna(dmax):
                date_max = pd.Timestamp(dmax).date().isoformat()
        except Exception:
            pass

    return {
        "nb_livraisons": nb_livraisons,
        "tonnage_total": tonnage_total,
        "ca_total": ca_total,
        "date_min": date_min,
        "date_max": date_max,
    }


def _create_project_from_upload(uploaded, theme_idx: int) -> str:
    project_id = str(uuid4())
    original_name = getattr(uploaded, "name", "") or "extraction"
    suffix = Path(original_name).suffix.lower() or ".xlsx"

    stem = _safe_stem(original_name)
    project_dir = PROJECT_FILES_DIR / f"user_{user_id}" / project_id
    project_dir.mkdir(parents=True, exist_ok=True)
    saved_path = project_dir / f"{stem}{suffix}"

    data = uploaded.getvalue() if hasattr(uploaded, "getvalue") else bytes(uploaded.read())
    saved_path.write_bytes(data)

    df_tmp, mapping_tmp = load_data(str(saved_path), clean_version=DATA_CLEAN_VERSION)
    stats = _compute_project_stats(df_tmp, mapping_tmp)

    default_name = f"Extraction - {stem}"
    if stats.get("date_min"):
        try:
            d0 = datetime.fromisoformat(stats["date_min"])
            default_name = f"Extraction du {d0:%d-%m-%Y} - {stem}"
        except Exception:
            default_name = f"Extraction du {stats['date_min']} - {stem}"

    projects.create_project(
        PROJECT_DB_PATH,
        project_id=project_id,
        user_id=user_id,
        name=default_name,
        data_path=str(saved_path),
        source_filename=original_name,
        date_min=stats.get("date_min"),
        date_max=stats.get("date_max"),
        nb_livraisons=stats.get("nb_livraisons"),
        tonnage_total=stats.get("tonnage_total"),
        ca_total=stats.get("ca_total"),
        theme_idx=int(theme_idx),
    )
    return project_id


def _ensure_default_project_if_needed() -> None:
    existing = projects.list_projects(PROJECT_DB_PATH, user_id)
    if existing:
        return
    if not os.path.exists(fichier_defaut):
        return

    try:
        df_tmp, mapping_tmp = load_data(str(fichier_defaut), clean_version=DATA_CLEAN_VERSION)
        stats = _compute_project_stats(df_tmp, mapping_tmp)
    except Exception:
        stats = {"date_min": None, "date_max": None, "nb_livraisons": None, "tonnage_total": None, "ca_total": None}

    project_id = str(uuid4())
    projects.create_project(
        PROJECT_DB_PATH,
        project_id=project_id,
        user_id=user_id,
        name="Extraction (par défaut)",
        data_path=str(fichier_defaut),
        source_filename=os.path.basename(fichier_defaut),
        date_min=stats.get("date_min"),
        date_max=stats.get("date_max"),
        nb_livraisons=stats.get("nb_livraisons"),
        tonnage_total=stats.get("tonnage_total"),
        ca_total=stats.get("ca_total"),
        theme_idx=0,
    )
    st.session_state.active_project_id = project_id


_ensure_default_project_if_needed()

# Sidebar upload -> creates a new "dossier" for this user
uploaded_file = st.sidebar.file_uploader(
    " Charger une extraction (Excel ou CSV)",
    type=["xls", "xlsx", "xlsm", "csv"],
    key=f"sidebar_uploader_{st.session_state.sidebar_upload_counter}",
)
if uploaded_file is not None:
    current = projects.get_project(PROJECT_DB_PATH, user_id, st.session_state.active_project_id) if st.session_state.active_project_id else None
    next_theme_idx = (int(current.theme_idx) + 1) % len(THEMES) if current else 0
    new_id = _create_project_from_upload(uploaded_file, theme_idx=next_theme_idx)
    st.session_state.active_project_id = new_id
    st.session_state.rename_project_id = new_id
    st.session_state.sidebar_upload_counter += 1
    st.rerun()

# Projects list (for the top bar)
all_projects = projects.list_projects(PROJECT_DB_PATH, user_id)
if not st.session_state.active_project_id and all_projects:
    st.session_state.active_project_id = all_projects[0].id
active_project = projects.get_project(PROJECT_DB_PATH, user_id, st.session_state.active_project_id) if st.session_state.active_project_id else None

# --- Barre de dossiers (zone principale, en haut) ---
st.markdown(
    """
    <style>
      div[data-testid="stMain"] .project-bar-title {
        font-weight: 800;
        color: #1f2937;
        margin: 4px 0 0 0;
      }
      div[data-testid="stMain"] .project-active-pill {
        display: inline-flex;
        gap: 10px;
        align-items: center;
        padding: 8px 10px;
        border-radius: 12px;
        border: 1px solid #d0d5dd;
        background: #ffffff;
        color: #111827;
        font-weight: 700;
        max-width: 100%;
        overflow: hidden;
        box-shadow: var(--shadow1);
      }
      div[data-testid="stMain"] .project-active-name {
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        max-width: 58vw;
        display: inline-block;
      }
      div[data-testid="stMain"] .project-active-meta {
        font-weight: 600;
        color: #475467;
        white-space: nowrap;
      }
      @media (max-width: 1100px) {
        div[data-testid="stMain"] .project-active-meta { display: none; }
        div[data-testid="stMain"] .project-active-name { max-width: 72vw; }
      }
      div[data-testid="stMain"] div[data-testid="stPopover"] > button {
        background: #111827;
        border: 1px solid #111827;
        color: #ffffff;
        border-radius: 12px;
        font-weight: 800;
      }
      div[data-testid="stMain"] div[data-testid="stPopover"] > button:hover {
        background: #0b3d91;
        border-color: #0b3d91;
        color: #ffffff;
      }
      div[data-testid="stMain"] div[data-testid="stPopover"] [role="dialog"] {
        min-width: 420px;
        max-width: 520px;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

bar_left, bar_right = st.columns([7, 3], vertical_alignment="center")
with bar_left:
    st.markdown("<div class='project-bar-title'>Extraction active</div>", unsafe_allow_html=True)
    if active_project:
        nb = int(active_project.nb_livraisons or 0)
        ton = float(active_project.tonnage_total or 0.0)
        ca = float(active_project.ca_total or 0.0)
        meta = f"{nb} liv • {ton:,.0f} T • {_human_money(ca)}".replace(",", " ")
        st.markdown(
            f"<div class='project-active-pill'><span class='project-active-name'>{active_project.name}</span> <span class='project-active-meta'>— {meta}</span></div>",
            unsafe_allow_html=True,
        )
        st.caption(f"Source : {active_project.source_filename or Path(active_project.data_path).name}")
    else:
        st.caption("Aucune extraction disponible pour ce compte.")
with bar_right:
    btn_c1, btn_c2 = st.columns(2, vertical_alignment="center")
    with btn_c1:
        with st.popover("Voir les extractions", use_container_width=True):
            st.markdown("**Sélectionner une extraction**")
            search_val = st.text_input(
                "Rechercher",
                value=st.session_state.project_search,
                placeholder="Nom du dossier…",
                label_visibility="collapsed",
                key="project_search_popover",
            )
            st.session_state.project_search = search_val

            q = (search_val or "").strip().casefold()
            filtered_projects = all_projects
            if q:
                filtered_projects = [p for p in all_projects if q in (p.name or "").casefold()]

            try:
                list_box = st.container(height=420)
            except TypeError:
                list_box = st.container()

            with list_box:
                if not filtered_projects:
                    st.info("Aucune extraction ne correspond à cette recherche.")
                for p in filtered_projects:
                    is_active = bool(active_project and p.id == active_project.id)
                    with st.container(border=True):
                        nb = int(p.nb_livraisons or 0)
                        ton = float(p.tonnage_total or 0.0)
                        ca = float(p.ca_total or 0.0)
                        st.caption(f"{nb} liv • {ton:,.0f} T • {_human_money(ca)}".replace(",", " "))
                        if st.button(
                            (p.name or "(Sans nom)") + (" (actif)" if is_active else ""),
                            key=f"pick_project_{p.id}",
                            type="primary" if is_active else "secondary",
                            use_container_width=True,
                        ):
                            st.session_state.active_project_id = p.id
                            st.session_state.rename_project_id = None
                            st.rerun()

            st.divider()
            if active_project and st.button("Renommer l'extraction active", use_container_width=True):
                st.session_state.rename_project_id = active_project.id
                st.rerun()

    with btn_c2:
        if st.button("＋ Nouvelle extraction", type="primary", use_container_width=True):
            st.session_state.show_top_uploader = not bool(st.session_state.show_top_uploader)

if st.session_state.show_top_uploader:
    top_file = st.file_uploader(
        "Importer une nouvelle extraction (Excel ou CSV)",
        type=["xls", "xlsx", "xlsm", "csv"],
        key=f"top_uploader_{st.session_state.top_upload_counter}",
        label_visibility="collapsed",
    )
    if top_file is not None:
        current = active_project
        next_theme_idx = (int(current.theme_idx) + 1) % len(THEMES) if current else 0
        new_id = _create_project_from_upload(top_file, theme_idx=next_theme_idx)
        st.session_state.active_project_id = new_id
        st.session_state.rename_project_id = new_id
        st.session_state.top_upload_counter += 1
        st.session_state.show_top_uploader = False
        st.rerun()

active_project = projects.get_project(PROJECT_DB_PATH, user_id, st.session_state.active_project_id) if st.session_state.active_project_id else None

if active_project and st.session_state.rename_project_id == active_project.id:
    with st.form("rename_project_form", clear_on_submit=False):
        new_name = st.text_input("Nom du dossier", value=active_project.name)
        c1, c2 = st.columns(2)
        with c1:
            ok = st.form_submit_button("Enregistrer", use_container_width=True)
        with c2:
            cancel = st.form_submit_button("Annuler", use_container_width=True)
    if ok:
        try:
            saved = projects.rename_project(PROJECT_DB_PATH, user_id=user_id, project_id=active_project.id, new_name=new_name)
            st.session_state.rename_project_id = None
            st.toast(f"Dossier renommé : {saved}")
            st.rerun()
        except projects.ProjectError as e:
            st.error(str(e))
    elif cancel:
        st.session_state.rename_project_id = None
        st.rerun()

if "local_data_source" not in st.session_state:
    st.session_state.local_data_source = active_project.data_path if active_project else fichier_defaut
else:
    st.session_state.local_data_source = active_project.data_path if active_project else st.session_state.local_data_source

data_source = st.session_state.local_data_source

theme_idx = int(active_project.theme_idx) if active_project else 0
st.session_state.theme_idx = theme_idx
theme = THEMES[theme_idx]
components.html(
    f"""
    <script>
    (function() {{
      const BTN_ID = "pdf-gen-btn";
      const STYLE_ID = "pdf-gen-style";
      const label = "Générer PDF";
      const bg = "{theme["primary"]}";
      const actionsId = "top-banner-actions";

      function ensureStyle() {{
        let style = window.parent.document.getElementById(STYLE_ID);
        if (!style) {{
          style = window.parent.document.createElement("style");
          style.id = STYLE_ID;
          window.parent.document.head.appendChild(style);
        }}
        style.textContent = `
          @media print {{
            @page {{ size: A4 landscape; margin: 10mm; }}
            #${{BTN_ID}} {{ display: none !important; }}
            section[data-testid="stSidebar"] {{ display: none !important; }}
            header[data-testid="stHeader"] {{ display: none !important; }}
            div[data-testid="stToolbar"] {{ display: none !important; }}
            div[data-testid="stDecoration"] {{ display: none !important; }}
            div.block-container {{ padding-top: 0 !important; }}
            * {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
            div[data-testid="stDataFrame"] {{ width: 100% !important; max-width: 100% !important; }}
            div[data-testid="stDataFrame"] > div {{ overflow: visible !important; }}
            div[data-testid="stPlotlyChart"], .stPlotlyChart, .js-plotly-plot {{
              break-inside: avoid !important;
              page-break-inside: avoid !important;
            }}
            div[data-testid="stPlotlyChart"] > div, .stPlotlyChart > div {{
              overflow: visible !important;
              height: auto !important;
              max-height: none !important;
            }}
          }}
        `;
      }}

      function ensureButton() {{
        let btn = window.parent.document.getElementById(BTN_ID);
        if (!btn) {{
          btn = window.parent.document.createElement("button");
          btn.id = BTN_ID;
          btn.type = "button";
          btn.innerText = label;
          btn.title = "Imprimer / enregistrer en PDF (état actuel de la page)";
          btn.onclick = function() {{
            try {{ window.parent.print(); }} catch (e) {{ try {{ window.print(); }} catch (e2) {{}} }}
          }};
        }}
        btn.style.position = "relative";
        btn.style.background = bg;
        btn.style.border = `1px solid ${{bg}}`;
        btn.style.color = "white";
        btn.style.padding = "8px 12px";
        btn.style.fontSize = "14px";
        btn.style.fontWeight = "700";
        btn.style.borderRadius = "10px";
        btn.style.boxShadow = "0 8px 18px rgba(0,0,0,0.18)";
        btn.style.cursor = "pointer";
        btn.style.whiteSpace = "nowrap";

        const actions = window.parent.document.getElementById(actionsId);
        if (actions && !actions.contains(btn)) {{
          actions.appendChild(btn);
        }} else if (!btn.isConnected) {{
          window.parent.document.body.appendChild(btn);
        }}
      }}

      try {{
        ensureStyle();
        ensureButton();
      }} catch (e) {{}}
    }})();
    </script>
    """,
    height=0,
)

if isinstance(data_source, str):
    st.sidebar.caption(f"Source: {data_source}")
else:
    try:
        st.sidebar.caption(f"Source: {data_source.name}")
    except Exception:
        pass

try:
    if uploaded_file is not None or (isinstance(data_source, str) and os.path.exists(data_source)):
        raw_full_df, mapping = load_data(data_source, clean_version=DATA_CLEAN_VERSION)
        df = raw_full_df.copy()

        if 'client' in mapping and mapping['client'] in df.columns:
            df[mapping['client']] = normalize_text_series(df[mapping['client']], kind="client")
        if 'produit' in mapping and mapping['produit'] in df.columns:
            df[mapping['produit']] = normalize_text_series(df[mapping['produit']], kind="produit")
        if 'mois' in mapping and mapping['mois'] in df.columns:
            df[mapping['mois']] = normalize_text_series(df[mapping['mois']], kind="mois")
        
        st.sidebar.header("Filtres")
        
        # Add Semaine Filter to match Excel visual
        if 'Semaine' in df.columns:
            semaines = ['Toutes'] + sorted(list(df['Semaine'].dropna().unique()))
            semaine_choisie = st.sidebar.selectbox("Numéro semaine", semaines)
            if semaine_choisie != 'Toutes':
                df = df[df['Semaine'] == semaine_choisie]
                
        if 'Année' in df.columns:
            annees = ['Toutes'] + list(df['Année'].dropna().unique())
            annee_choisie = st.sidebar.selectbox("Année", annees)
            if annee_choisie != 'Toutes':
                df = df[df['Année'] == annee_choisie]
                
        if 'mois' in mapping:
            mois_col = mapping['mois']
            mois_options = df[mois_col].dropna().unique().tolist()
            mois_options = [
                m
                for m in mois_options
                if str(m).strip() and str(m).casefold() not in ("<na>", "nan")
            ]
            months_order = [
                'janvier', 'février', 'mars', 'avril', 'mai', 'juin',
                'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'
            ]
            month_rank = {m: i for i, m in enumerate(months_order)}
            mois_options = sorted(mois_options, key=lambda m: month_rank.get(str(m).casefold(), 999))

            mois_key = "filter_mois"
            if mois_key in st.session_state:
                st.session_state[mois_key] = [m for m in st.session_state[mois_key] if m in mois_options]
            mois_choisis = st.sidebar.multiselect(
                "Mois en lettres",
                options=mois_options,
                key=mois_key,
                help="Sélection multiple. Laisser vide = tous les mois."
            )
            if mois_choisis:
                df = df[df[mois_col].isin(mois_choisis)]

        if 'client' in mapping:
            client_col = mapping['client']
            client_options = df[client_col].dropna().unique().tolist()
            client_options = [c for c in client_options if str(c).strip()]
            client_options = sorted(client_options, key=lambda v: str(v).casefold())

            clients_key = "filter_clients"
            if clients_key in st.session_state:
                st.session_state[clients_key] = [c for c in st.session_state[clients_key] if c in client_options]
            clients_choisis = st.sidebar.multiselect(
                "Clients",
                options=client_options,
                key=clients_key,
                help="Sélection multiple. Laisser vide = tous les clients."
            )
            if clients_choisis:
                df = df[df[client_col].isin(clients_choisis)]
                
        if 'produit' in mapping:
            produit_col = mapping['produit']
            produit_options = df[produit_col].dropna().unique().tolist()
            produit_options = sorted(produit_options, key=lambda v: str(v).casefold())

            produits_key = "filter_produits"
            if produits_key in st.session_state:
                st.session_state[produits_key] = [p for p in st.session_state[produits_key] if p in produit_options]
            produits_choisis = st.sidebar.multiselect(
                "Produits",
                options=produit_options,
                key=produits_key,
                help="Sélection multiple. Laisser vide = tous les produits."
            )
            if produits_choisis:
                df = df[df[produit_col].isin(produits_choisis)]

        # EN-TÊTE AVEC BOUTON D'EXPORT
        col_title, col_export = st.columns([4, 1])
        with col_title:
            st.title(" Tableau de Bord - Suivi des Livraisons Carrière")

        ca_col = mapping.get('ca')
        poids_col = mapping.get('poids')
        client_col = mapping.get('client')
        produit_col = mapping.get('produit')
        date_col = mapping.get('date')

        # --- KPIs Globaux ---
        st.markdown("<div class='kpi-title'>KPIs Globaux</div>", unsafe_allow_html=True)
        ca_total = df[ca_col].sum() if ca_col else 0
        ton_total = df[poids_col].sum() if poids_col else 0
        nb_livraisons = len(df)
        nb_clients = df[client_col].nunique() if client_col else 0
        prix_moyen = ca_total / ton_total if ton_total > 0 else 0
        
        jours_ouvres = 0
        if date_col and 'Est_Ouvre' in df.columns:
            jours_ouvres = df[df['Est_Ouvre']][date_col].dt.date.nunique()
        livraisons_jour = nb_livraisons / jours_ouvres if jours_ouvres > 0 else 0

        kpi_html = "<div class='kpi-grid'>"
        kpi_html += _kpi_card_html(
            "CA total",
            _human_money(ca_total),
            _fmt_eur_full(ca_total, decimals=0),
            tooltip=_fmt_eur_full(ca_total, decimals=0),
        )
        kpi_html += _kpi_card_html(
            "Tonnage total",
            f"{_fmt_number_fr(ton_total, decimals=0)} T",
            f"{_fmt_number_fr(ton_total, decimals=0)} T",
            tooltip=f"{_fmt_number_fr(ton_total, decimals=0)} T",
        )
        kpi_html += _kpi_card_html(
            "Nombre de livraisons",
            f"{_fmt_number_fr(nb_livraisons, decimals=0)}",
            f"{_fmt_number_fr(nb_livraisons, decimals=0)}",
            tooltip=f"{_fmt_number_fr(nb_livraisons, decimals=0)}",
        )
        kpi_html += _kpi_card_html(
            "Nombre de clients",
            f"{_fmt_number_fr(nb_clients, decimals=0)}",
            f"{_fmt_number_fr(nb_clients, decimals=0)}",
            tooltip=f"{_fmt_number_fr(nb_clients, decimals=0)}",
        )
        kpi_html += _kpi_card_html(
            "Prix moyen / tonne",
            _fmt_eur_full(prix_moyen, decimals=2),
            _fmt_eur_full(prix_moyen, decimals=2),
            tooltip=_fmt_eur_full(prix_moyen, decimals=2),
        )
        kpi_html += _kpi_card_html(
            "Livraisons / jour",
            _fmt_number_fr(livraisons_jour, decimals=1),
            _fmt_number_fr(livraisons_jour, decimals=1),
            tooltip=_fmt_number_fr(livraisons_jour, decimals=1),
        )
        kpi_html += "</div>"
        st.markdown(kpi_html, unsafe_allow_html=True)
        
        # EXPORT BOUTON (Placé en haut à droite visuellement via les colonnes)
        with col_export:
            st.markdown("<br>", unsafe_allow_html=True) # Alignement vertical
            
            if 'excel_data' not in st.session_state:
                st.session_state.excel_data = None
            
            # Bouton pour déclencher la génération lourde
            if st.button(" Préparer l'extraction"):
                with st.spinner("Génération du fichier Excel en cours..."):
                    # Créer un DF pour la synthèse
                    synth_data = {
                        'Indicateur': ['CA Total', 'Tonnage Total', 'Nb Livraisons', 'Nb Clients', 'Prix Moyen / T', 'Livraisons/Jour'],
                        'Valeur': [ca_total, ton_total, nb_livraisons, nb_clients, prix_moyen, livraisons_jour]
                    }
                    df_synth = pd.DataFrame(synth_data)
                    
                    # Générer le fichier et stocker en session
                    st.session_state.excel_data = generate_excel_download(df, df_synth, mapping)
                    st.toast("Fichier prêt !")
            
            # Afficher le bouton de téléchargement seulement si les données sont prêtes
            if st.session_state.excel_data is not None:
                st.download_button(
                    label=" Télécharger l'Excel",
                    data=st.session_state.excel_data,
                    file_name="Dashboard_Extract_Carriere.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        st.markdown("---")
        
        # --- VIEWS (Excel vs Dash) ---
        main_tab_dash, main_tab_excel, main_tab_update = st.tabs(
            ["Tableau de bord", "Analyses détaillées", "✏️ Mise à jour du tableur"]
        )
        
        # --- VUE GRAPHIQUES ORIGINALE ---
        with main_tab_dash:
            st.subheader(" Indicateurs & graphiques")
            
            # Reprise des analyses avancées par KPIs
            adv_col1, adv_col2, adv_col3 = st.columns(3)
            with adv_col1:
                st.markdown("<div class='kpi-title'> Par Période</div>", unsafe_allow_html=True)
                if date_col and ca_col:
                    df_monthly = df.groupby(df[date_col].dt.to_period('M'))[ca_col].sum().reset_index()
                    if len(df_monthly) > 1:
                        df_monthly = df_monthly.sort_values(date_col)
                        last_month_ca = df_monthly.iloc[-1][ca_col]
                        prev_month_ca = df_monthly.iloc[-2][ca_col]
                        evol_pct = ((last_month_ca - prev_month_ca) / prev_month_ca) * 100 if prev_month_ca > 0 else 0
                        evol_str = f"+{evol_pct:.1f}%" if evol_pct >= 0 else f"{evol_pct:.1f}%"
                        
                        meilleur_mois = df_monthly.loc[df_monthly[ca_col].idxmax()][date_col].strftime('%m/%Y')
                        pire_mois = df_monthly.loc[df_monthly[ca_col].idxmin()][date_col].strftime('%m/%Y')
                    else:
                        evol_str = "N/A"
                        meilleur_mois = df_monthly.iloc[0][date_col].strftime('%m/%Y') if len(df_monthly)>0 else "N/A"
                        pire_mois = meilleur_mois
                        
                    st.markdown(f"<div class='kpi-box'>Évolution CA vs M-1 : <span class='kpi-val'>{evol_str}</span><br>Meilleur mois : <b>{meilleur_mois}</b> | Pire mois : <b>{pire_mois}</b></div>", unsafe_allow_html=True)

            with adv_col2:
                st.markdown("<div class='kpi-title'> Par Produit</div>", unsafe_allow_html=True)
                if produit_col and ca_col and poids_col:
                    prod_stats = df.groupby(produit_col).agg({ca_col: 'sum', poids_col: 'sum'})
                    best_prod_ca = prod_stats[ca_col].idxmax() if not prod_stats.empty else "N/A"
                    best_prod_poids = prod_stats[poids_col].idxmax() if not prod_stats.empty else "N/A"
                    st.markdown(f"<div class='kpi-box'>Top Produit (CA) : <b>{best_prod_ca}</b><br>Top Produit (Tonnage) : <b>{best_prod_poids}</b></div>", unsafe_allow_html=True)

            with adv_col3:
                st.markdown("<div class='kpi-title'> Par Client</div>", unsafe_allow_html=True)
                if client_col and ca_col:
                    client_ca = df.groupby(client_col)[ca_col].sum().sort_values(ascending=False)
                    if len(client_ca) > 0:
                        top3_ca = client_ca.head(3).sum()
                        part_top3 = (top3_ca / ca_total * 100) if ca_total > 0 else 0
                        st.markdown(
                            f"<div class='kpi-box'>Part de CA (Top 3) : <span class='kpi-val'>{part_top3:.1f}%</span><br>Top Client : <b>{client_ca.index[0]}</b></div>",
                            unsafe_allow_html=True,
                        )

            try:
                row1_col1, row1_col2 = st.columns(2)
                
                df_evol_m = pd.DataFrame()
                if date_col and ca_col and poids_col:
                    df_evol_m = df.groupby(df[date_col].dt.to_period('M')).agg({ca_col: 'sum', poids_col: 'sum'}).reset_index()
                    df_evol_m[date_col] = df_evol_m[date_col].astype(str)

                # Graphique 1 : Évolution mensuelle du CA (Courbe)
                with row1_col1:
                    if not df_evol_m.empty:
                         fig1 = px.line(df_evol_m, x=date_col, y=ca_col, 
                                        title="Graphique 1: Évolution Mensuelle du CA (€)", markers=True,
                                        color_discrete_sequence=[theme["primary"]])
                         fig1.update_layout(xaxis_title="Mois", yaxis_title="CA (€)")
                         st.plotly_chart(apply_plotly_style(fig1), use_container_width=True)

                # Graphique 2 : Évolution mensuelle du Tonnage (Histogramme)
                with row1_col2:
                    if not df_evol_m.empty:
                         fig2 = px.bar(df_evol_m, x=date_col, y=poids_col, 
                                       title="Graphique 2: Évolution Mensuelle du Tonnage (T)",
                                       color_discrete_sequence=[theme["secondary"]])
                         fig2.update_layout(xaxis_title="Mois", yaxis_title="Tonnage (T)")
                         st.plotly_chart(apply_plotly_style(fig2), use_container_width=True)

                row2_col1, row2_col2 = st.columns(2)

                # Graphique 3 : Répartition par Nature de Produit (Camembert)
                with row2_col1:
                    if produit_col and ca_col and poids_col:
                        st.markdown("<b>Graphique 3: Répartition par Produit</b>", unsafe_allow_html=True)
                        tab_p1, tab_p2 = st.tabs(["Tonnage", "Chiffre d'Affaires"])
                        df_prod_pie = df.groupby(produit_col).agg({ca_col: 'sum', poids_col: 'sum'}).reset_index()
                        
                        with tab_p1:
                            fig3_t = px.pie(
                                df_prod_pie,
                                values=poids_col,
                                names=produit_col,
                                hole=0.4,
                                color_discrete_sequence=theme["pie"],
                            )
                            fig3_t.update_traces(textposition="inside", textinfo="percent+label")
                            fig3_t.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                            st.plotly_chart(apply_plotly_style(fig3_t), use_container_width=True)
                        with tab_p2:
                            fig3_ca = px.pie(
                                df_prod_pie,
                                values=ca_col,
                                names=produit_col,
                                hole=0.4,
                                color_discrete_sequence=theme["pie"],
                            )
                            fig3_ca.update_traces(textposition="inside", textinfo="percent+label")
                            fig3_ca.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                            st.plotly_chart(apply_plotly_style(fig3_ca), use_container_width=True)

                # Graphique 4 : Top 10 Clients (Barre horizontale)
                with row2_col2:
                    if client_col and ca_col and poids_col:
                        df_client_top = df.groupby(client_col).agg({ca_col: 'sum', poids_col: 'sum'}).reset_index()
                        df_client_top = df_client_top.sort_values(by=ca_col, ascending=False).head(10)
                        df_client_top = df_client_top.sort_values(by=ca_col, ascending=True) # Ascending for horizontal bar
                        
                        fig4 = px.bar(df_client_top, x=ca_col, y=client_col, orientation='h',
                                      title="Graphique 4: Top 10 Clients (CA)", text=ca_col,
                                      color_discrete_sequence=[theme["accent"]],
                                      hover_data=[poids_col])
                        fig4.update_traces(texttemplate='%{text:.2s} €', textposition='outside')
                        fig4.update_layout(xaxis_title="CA (€)", yaxis_title="", margin=dict(l=150))
                        st.plotly_chart(apply_plotly_style(fig4), use_container_width=True)

                row3_col1, row3_col2 = st.columns(2)

                # Graphique 5 : CA Journalier
                with row3_col1:
                    if date_col and ca_col:
                        df_daily = df.groupby(df[date_col].dt.date)[ca_col].sum().reset_index()
                        fig5 = px.area(
                            df_daily,
                            x=date_col,
                            y=ca_col,
                            title="Graphique 5: CA Journalier (Activité)",
                            color_discrete_sequence=[theme["primary"]],
                        )
                        fig5.update_layout(
                            xaxis_title="Jour",
                            yaxis_title="CA (€)",
                            xaxis=dict(showgrid=False),
                            yaxis=dict(showgrid=False),
                        )
                        st.plotly_chart(apply_plotly_style(fig5), use_container_width=True)

                # Graphique 6 : Comparatif Mensuel CA vs Tonnage (Double axe)
                with row3_col2:
                    if not df_evol_m.empty:
                        fig6 = go.Figure()
                        fig6.add_trace(go.Bar(x=df_evol_m[date_col], y=df_evol_m[poids_col], name="Tonnage (T)", marker_color=theme["secondary"], yaxis='y1'))
                        fig6.add_trace(go.Scatter(x=df_evol_m[date_col], y=df_evol_m[ca_col], name="CA (€)", marker_color=theme["primary"], mode='lines+markers', yaxis='y2'))
                        
                        fig6.update_layout(
                            title="Graphique 6: Comparatif CA vs Tonnage",
                            xaxis=dict(title="Mois"),
                            yaxis=dict(
                                title=dict(text="Tonnage (T)", font=dict(color=theme["secondary"])),
                                tickfont=dict(color=theme["secondary"]),
                            ),
                            yaxis2=dict(
                                title=dict(text="CA (€)", font=dict(color=theme["primary"])),
                                tickfont=dict(color=theme["primary"]),
                                overlaying="y",
                                side="right",
                            ),
                            legend=dict(x=0, y=1.1, orientation="h"),
                        )
                        st.plotly_chart(apply_plotly_style(fig6), use_container_width=True)
            except Exception as e:
                st.write(f"Erreur d'affichage des graphiques : {e}")

        # --- VUE MISE À JOUR DU TABLEUR ---
        with main_tab_update:
            st.subheader(" Mise à jour du tableur (ajout de nouvelles lignes)")
            st.info(
                "Importez un Excel contenant de nouveaux enregistrements. "
                "En mode fichier local (par défaut), le fichier source sera mis à jour et les graphiques se recalculeront automatiquement."
            )
            if isinstance(data_source, str):
                lock_path = Path(data_source).with_name("~$" + Path(data_source).name)
                if lock_path.exists():
                    st.warning(
                        "Le fichier source semble actuellement ouvert dans Excel (fichier de verrouillage détecté). "
                        "Fermez Excel pour pouvoir écraser le fichier, ou l'application enregistrera une nouvelle copie."
                    )

            with st.form("update_form", clear_on_submit=False):
                update_file = st.file_uploader(
                    "Fichier à ajouter (Excel ou CSV)",
                    type=["xlsx", "xls", "xlsm", "csv"],
                    key="update_uploader",
                )

                sheet_name = 0
                if update_file is not None and Path(str(getattr(update_file, "name", ""))).suffix.lower() != ".csv":
                    try:
                        update_file.seek(0)
                        xl = pd.ExcelFile(update_file)
                        if len(xl.sheet_names) > 1:
                            sheet_name = st.selectbox(
                                "Onglet du fichier à importer",
                                options=xl.sheet_names,
                                index=0,
                                key="update_sheet_select",
                            )
                    except Exception:
                        sheet_name = 0

                dedup_by_ticket = st.checkbox(
                    "Supprimer les doublons par Ticket",
                    value=True,
                    help="Recommandé si le ticket identifie une livraison unique. Désactivez si vous pensez que le ticket peut se répéter légitimement.",
                    key="update_dedup_ticket",
                )

                submitted = st.form_submit_button("🚀 Mettre à jour", type="primary", use_container_width=True)

            if submitted:
                if update_file is None:
                    st.warning("Veuillez d'abord sélectionner un fichier Excel à ajouter.")
                else:
                    try:
                        with st.spinner("Fusion des nouvelles données..."):
                            update_file.seek(0)
                            new_df, new_mapping = load_data(
                                update_file,
                                clean_version=DATA_CLEAN_VERSION,
                                sheet_name=sheet_name,
                            )

                            # Harmoniser les noms de colononnes (si le fichier d'update n'a pas exactement les mêmes entêtes)
                            for key, base_col in mapping.items():
                                upd_col = new_mapping.get(key)
                                if upd_col and upd_col in new_df.columns and upd_col != base_col:
                                    new_df = new_df.rename(columns={upd_col: base_col})

                            # Aligner les colonnes (les colonnes dérivées existent dans raw_full_df)
                            for col in raw_full_df.columns:
                                if col not in new_df.columns:
                                    new_df[col] = pd.NA
                            new_df = new_df[raw_full_df.columns]

                            updated_full_df = pd.concat([raw_full_df, new_df], ignore_index=True)

                            # Optionnel: éviter les doublons exacts par ticket si la colonne existe
                            ticket_col = mapping.get("ticket")
                            removed = 0
                            if dedup_by_ticket and ticket_col and ticket_col in updated_full_df.columns:
                                before = len(updated_full_df)
                                ticket = updated_full_df[ticket_col].astype("string").str.strip()
                                dup_mask = ticket.notna() & ticket.ne("") & ticket.duplicated(keep="first")
                                updated_full_df = updated_full_df[~dup_mask]
                                removed = before - len(updated_full_df)

                            if not (isinstance(data_source, str) and os.path.exists(data_source)):
                                st.error(
                                    "Impossible de mettre à jour le fichier source dans ce mode (fichier uploadé). "
                                    "Utilisez le fichier par défaut local ou relancez l'app avec un fichier sur disque."
                                )
                            else:
                                saved_path = None
                                save_note = None
                                try:
                                    updated_full_df.to_excel(data_source, index=False)
                                    saved_path = data_source
                                except PermissionError:
                                    src = Path(data_source)
                                    alt = src.with_name(f"{src.stem} - MAJ {datetime.now():%Y%m%d-%H%M%S}{src.suffix}")
                                    updated_full_df.to_excel(alt, index=False)
                                    st.session_state.local_data_source = str(alt)
                                    saved_path = str(alt)
                                    save_note = (
                                        "Impossible d'écraser le fichier source (souvent parce qu'il est ouvert dans Excel). "
                                        "Une copie mise à jour a été enregistrée."
                                    )

                                msg = f"{len(new_df)} lignes ajoutées."
                                if removed:
                                    msg += f" {removed} doublons supprimés (Ticket)."
                                st.success(msg)
                                if save_note:
                                    st.warning(f"{save_note} Fichier : {saved_path}")

                                # Mettre à jour les stats + chemin dans le dossier actif (multi-compte)
                                try:
                                    if active_project and saved_path:
                                        stats_u = _compute_project_stats(updated_full_df, mapping)
                                        projects.update_project_data(
                                            PROJECT_DB_PATH,
                                            user_id=user_id,
                                            project_id=active_project.id,
                                            data_path=str(saved_path),
                                            date_min=stats_u.get("date_min"),
                                            date_max=stats_u.get("date_max"),
                                            nb_livraisons=stats_u.get("nb_livraisons"),
                                            tonnage_total=stats_u.get("tonnage_total"),
                                            ca_total=stats_u.get("ca_total"),
                                        )
                                        st.session_state.local_data_source = str(saved_path)
                                except Exception:
                                    pass

                                with st.expander("Détails import", expanded=False):
                                    details = {
                                        "Fichier enregistré": saved_path,
                                        "Lignes lues (fichier importé)": int(len(new_df)),
                                        "Lignes totales (après fusion)": int(len(updated_full_df)),
                                        "Doublons supprimés (Ticket)": int(removed),
                                    }
                                    date_col = mapping.get("date")
                                    if date_col and date_col in new_df.columns:
                                        try:
                                            details["Date min (import)"] = str(pd.to_datetime(new_df[date_col], errors="coerce").min())
                                            details["Date max (import)"] = str(pd.to_datetime(new_df[date_col], errors="coerce").max())
                                        except Exception:
                                            pass
                                    mois_col = mapping.get("mois")
                                    if mois_col and mois_col in new_df.columns:
                                        try:
                                            mois_vals = new_df[mois_col].dropna().astype(str).unique().tolist()
                                            details["Mois (import)"] = sorted(mois_vals, key=lambda v: str(v).casefold())
                                        except Exception:
                                            pass
                                    st.write(details)

                                st.cache_data.clear()
                                st.info("Mise à jour enregistrée. Cliquez sur '🔄 Recharger' pour actualiser filtres et graphiques.")
                                if st.button("🔄 Recharger", key="reload_after_update"):
                                    st.rerun()
                    except Exception as e:
                        st.exception(e)

        
        with main_tab_excel:
            # --- Analyses Avancées (Type Excel) ---
            st.subheader(" Analyses Détaillées - Type Excel")
            tab_client, tab_produit, tab_livraison = st.tabs([" Analyse par Produit/Client", " Évolution des Livraisons", " Performance Clients"])
            
            # TAB 1: Analyse par Produit (Correspond à la capture 1 - Tableaux à barres intégrées)
            with tab_client:
                if produit_col and ca_col and poids_col:
                    col_pt1, col_pt2 = st.columns([1.5, 1])
                    
                    with col_pt1:
                        df_prod = df.groupby(produit_col).agg({poids_col: 'sum', ca_col: 'sum'}).reset_index()
                        df_prod = df_prod.sort_values(by=poids_col, ascending=False)
                        
                        st.markdown("**Quantité livrée et CA par Produit**")

                        if df_prod.empty:
                            st.info("Aucune donnée pour cette sélection.")
                        else:
                            max_poids = safe_progress_max(df_prod[poids_col].max())
                            max_ca = safe_progress_max(df_prod[ca_col].max())

                            st.dataframe(
                                df_prod,
                                column_config={
                                    produit_col: st.column_config.TextColumn("Produit"),
                                    poids_col: st.column_config.ProgressColumn(
                                        "QTE LIVREE EN T",
                                        format="%.2f",
                                        min_value=0,
                                        max_value=max_poids,
                                    ),
                                    ca_col: st.column_config.ProgressColumn(
                                        "Chiffre d'affaires",
                                        format="%d €",
                                        min_value=0,
                                        max_value=max_ca,
                                    ),
                                },
                                hide_index=True,
                                use_container_width=True
                            )
                            st.markdown(
                                f"<div class='print-only print-table'>{df_prod.to_html(index=False)}</div>",
                                unsafe_allow_html=True,
                            )
                    
                    with col_pt2:
                        st.markdown("**Répartition (Total)**")
                        if df_prod.empty:
                            st.info("Aucune donnée pour cette sélection.")
                        else:
                            fig_pie = px.pie(df_prod, values=poids_col, names=produit_col, hole=0.3, color_discrete_sequence=theme["pie"])
                            fig_pie.update_traces(textposition='inside', textinfo='none') # Hide text inside to match Excel
                            fig_pie.update_layout(showlegend=True, margin=dict(t=0, b=0, l=0, r=0))
                            st.plotly_chart(apply_plotly_style(fig_pie, kind="pie"), use_container_width=True)

            # TAB 2: Evolution des livraisons (Correspond à la capture 3 - Courbes croisées)
            with tab_produit:
                if date_col and produit_col and poids_col and mapping.get('mois'):
                    st.markdown("**Analyse des quantités livrées par mois et par produit**")
                    
                    # Order months correctly
                    months_order = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']
                    
                    # Get total deliveries per month/product
                    if 'Semaine' in df.columns:
                        # Let user choose by week or month, default to month
                        group_col = mapping['mois']
                        
                        df_evol_prod = df.groupby([group_col, produit_col])[poids_col].sum().reset_index()
                        
                        # Ensure categorical order for months
                        if group_col == mapping['mois']:
                            # Filter to only contain valid formatting
                            df_evol_prod[group_col] = df_evol_prod[group_col].str.lower()
                            # Sort
                            df_evol_prod[group_col] = pd.Categorical(df_evol_prod[group_col], categories=months_order, ordered=True)
                            df_evol_prod = df_evol_prod.sort_values(group_col)
                        
                        fig_lines = px.line(
                            df_evol_prod,
                            x=group_col,
                            y=poids_col,
                            color=produit_col,
                            markers=True,
                            color_discrete_sequence=theme["qualitative"],
                        )
                        fig_lines.update_layout(xaxis_title="Période", yaxis_title="Somme de Net (T)", 
                                                legend_title="Produit", hovermode="x unified")
                    st.plotly_chart(apply_plotly_style(fig_lines), use_container_width=True)

            # TAB 3: Performance Clients (Correspond à la capture 2 - Analyse matricielle)
            with tab_livraison:
                if client_col and ca_col and poids_col and produit_col:
                    st.markdown("**Performance Globale par Client**")
                    
                    df_client_perf = df.groupby(client_col).agg({poids_col: 'sum', ca_col: 'sum'}).reset_index()
                    df_client_perf = df_client_perf.sort_values(by=poids_col, ascending=False)
                    
                    col_c1, col_c2 = st.columns([1, 2])
                    with col_c1:
                        # Show Bar charts for top 20 clients
                        top_clients = df_client_perf.head(20).copy()
                        top_clients = top_clients.sort_values(by=poids_col, ascending=False)
                        
                        st.markdown("**Top 20 Clients**")

                        if top_clients.empty:
                            st.info("Aucune donnée pour cette sélection.")
                        else:
                            max_poids_c = safe_progress_max(top_clients[poids_col].max())
                            max_ca_c = safe_progress_max(top_clients[ca_col].max())

                            st.dataframe(
                                top_clients,
                                column_config={
                                    client_col: st.column_config.TextColumn("Clients"),
                                    poids_col: st.column_config.ProgressColumn(
                                        "Livraison totale",
                                        format="%.2f",
                                        min_value=0,
                                        max_value=max_poids_c,
                                    ),
                                    ca_col: st.column_config.ProgressColumn(
                                        "Chiffre d'affaires",
                                        format="%d €",
                                        min_value=0,
                                        max_value=max_ca_c,
                                    ),
                                },
                                hide_index=True,
                                use_container_width=True
                            )
                            st.markdown(
                                f"<div class='print-only print-table'>{top_clients.to_html(index=False)}</div>",
                                unsafe_allow_html=True,
                            )
                        
                    with col_c2:
                        # Matrix Client vs Product
                        st.markdown("**Livraison par produit (en tonnes)**")
                        pivot_df = pd.pivot_table(df, values=poids_col, index=client_col, columns=produit_col, aggfunc='sum', fill_value=0)
                        pivot_df['Total'] = pivot_df.sum(axis=1)
                        pivot_df = pivot_df.sort_values(by='Total', ascending=False)
                        # Limit to top 20 like the left dataframe to keep them vertically aligned
                        pivot_df_top = pivot_df.head(20) 
                        pivot_df_top = pivot_df_top.drop(columns=['Total'])
                        
                        # Style the dataframe like Excel
                        st.dataframe(
                            pivot_df_top.style.format("{:,.2f}").background_gradient(cmap=theme["table_cmap"], axis=None),
                            height=600,
                            use_container_width=True,
                        )
                        st.markdown(
                            f"<div class='print-only print-table'>{pivot_df_top.to_html()}</div>",
                            unsafe_allow_html=True,
                        )
                        
                        # Add stacked bar chart below the matrix table
                        st.markdown("**Graphique de Livraison par Produit (Top 20 Clients)**")
                        # Prepare data for stacked bars
                        # Melt the pivot table back to long format for Plotly
                        melted_df = pivot_df_top.reset_index().melt(id_vars=client_col, value_vars=pivot_df_top.columns, var_name=produit_col, value_name=poids_col)
                        # Filter out zero values to keep chart clean
                        melted_df = melted_df[melted_df[poids_col] > 0]
                        
                        fig_stacked = px.bar(
                            melted_df,
                            x=client_col,
                            y=poids_col,
                            color=produit_col,
                            title="Répartition des volumes par Produit",
                            color_discrete_sequence=theme["qualitative"],
                        )
                        fig_stacked.update_layout(xaxis_title="Clients (Top 20)", yaxis_title="Tonnage (T)", barmode='stack', hovermode="x unified")
                        st.plotly_chart(apply_plotly_style(fig_stacked), use_container_width=True)

        # --- Base de données Brute ---
        st.markdown("---")
        with st.expander(" Voir les données détaillées brutes"):
            st.dataframe(df, use_container_width=True)
            
    else:
        st.warning("Veuillez charger un fichier Excel (.xls, .xlsx) depuis le panneau de gauche.")
        
except Exception as e:
    st.error(f"Une erreur s'est produite lors de l'analyse du fichier : {e}")
