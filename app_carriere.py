import faulthandler
faulthandler.enable()

import os
import shutil
import io
import math
import re
import secrets
import json
from uuid import uuid4
import html
import sys
import importlib.util
from typing import Optional
from pathlib import Path
from datetime import datetime

import streamlit as st
import pandas as pd
import requests
import plotly.express as px
import plotly.graph_objects as go
import streamlit.components.v1 as components


def _load_local_module(py_filename: str, module_name: str):
    module_path = Path(__file__).resolve().parent / py_filename
    spec = importlib.util.spec_from_file_location(module_name, module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Impossible de charger {py_filename}.")
    mod = importlib.util.module_from_spec(spec)
    # Enregistre le module : identité cohérente + tracebacks plus clairs.
    sys.modules[module_name] = mod
    spec.loader.exec_module(mod)
    return mod


# Force le chargement des modules locaux (évite les collisions avec des paquets installés nommés "auth"/"projects").
auth = _load_local_module("auth.py", "carriboard_auth")
projects = _load_local_module("projects.py", "carriboard_projects")

# Incrémenter cette valeur pour invalider le cache Streamlit quand les règles de nettoyage changent.
DATA_CLEAN_VERSION = "2026-03-02-01"

def normalize_text_series(series: pd.Series, kind: Optional[str] = None) -> pd.Series:
    s = series.astype("string")
    s = s.str.normalize("NFKC")
    s = s.str.replace(r"[\u0000-\u001F\u007F-\u009F]", "", regex=True)  # caractères de contrôle
    s = s.str.replace(
        r"[\u200B-\u200F\u202A-\u202E\u2060-\u2069\u061C\uFEFF]",
        "",
        regex=True,
    )  # caractères invisibles/de formatage (bidi, zéro‑largeur, etc.)
    s = s.str.replace("\u00A0", " ", regex=False)  # espace insécable
    s = s.str.replace("\u202F", " ", regex=False)  # espace insécable étroit
    s = s.str.replace(r"[’‘´`ʹʻ]", "'", regex=True)  # apostrophes
    s = s.str.replace(r"[‐‑‒–—―−﹣－]", "-", regex=True)  # tirets / signe moins

    if kind == "client":
        # Harmonise les séparateurs courants pour éviter les doublons de libellés client.
        # (ex. "RB TRAVAUX" vs "RB-TRAVAUX", "ITB/RDP" vs "ITBRDP").
        s = s.str.replace(r"[-/_]", " ", regex=True)
        s = s.str.replace(r"[.,]", " ", regex=True)
        s = s.str.replace(r"\bINDUTRIES\b", "INDUSTRIES", regex=True)

    if kind == "mois":
        # Standardise les libellés de mois (inclut les erreurs d'encodage fréquentes dans certains Excels)
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
        # Canonicalise les variantes restantes en gardant le libellé le plus fréquent par clé
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
                height=520,
                legend=dict(
                    orientation="v",
                    yanchor="middle",
                    y=0.5,
                    xanchor="right",
                    x=1.0,
                    bgcolor="rgba(255,255,255,0.75)",
                    bordercolor="rgba(0,0,0,0.08)",
                    borderwidth=1,
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
    page_title="Tableau de Bord",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Auth (persistant au rafraîchissement via token en URL) ---
APP_DIR = Path(__file__).resolve().parent


def _pick_data_root(app_dir: Path) -> Path:
    data_dir_raw = (os.environ.get("CARRIBOARD_DATA_DIR") or os.environ.get("DATA_ROOT") or "").strip()
    if not data_dir_raw:
        try:
            data_dir_raw = str(st.secrets.get("CARRIBOARD_DATA_DIR") or st.secrets.get("DATA_ROOT") or "").strip()  # type: ignore[attr-defined]
        except Exception:
            data_dir_raw = ""

    # Si explicitement fourni, on l'utilise tel quel (recommandé sur Streamlit Cloud).
    if data_dir_raw:
        preferred = Path(data_dir_raw).expanduser()
        candidates = [preferred]
    else:
        # Évite d'écrire dans le dossier du dépôt sur Streamlit Cloud :
        # cela peut déclencher le file-watcher et provoquer des relances en boucle.
        candidates = [
            Path.home() / ".carriboard" / "data",
            Path("/tmp") / "carriboard" / "data",
            app_dir / "data",
        ]

    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            probe = candidate / ".write_test"
            probe.write_bytes(b"1")
            probe.unlink(missing_ok=True)
            return candidate
        except Exception:
            continue

    return (app_dir / "data")


@st.cache_resource
def _data_root_cached() -> Path:
    return _pick_data_root(APP_DIR)


DATA_ROOT = _data_root_cached()
AUTH_DB_PATH = DATA_ROOT / "users.sqlite3"


@st.cache_resource
def _auth_db_ready(db_path: Path) -> None:
    auth.init_db(db_path)


_auth_db_ready(AUTH_DB_PATH)


def _qp_read() -> dict:
    if hasattr(st, "query_params"):
        try:
            raw = dict(st.query_params)
        except Exception:
            raw = {}
        out: dict[str, str] = {}
        for k, v in raw.items():
            if isinstance(v, (list, tuple)):
                out[str(k)] = str(v[0]) if v else ""
            else:
                out[str(k)] = str(v)
        return out

    try:
        raw = st.experimental_get_query_params()
    except Exception:
        raw = {}

    out: dict[str, str] = {}
    for k, v in (raw or {}).items():
        if isinstance(v, (list, tuple)):
            out[str(k)] = str(v[0]) if v else ""
        else:
            out[str(k)] = str(v)
    return out


def _qp_write(params: dict[str, str]) -> None:
    clean = {str(k): str(v) for k, v in (params or {}).items() if v not in (None, "")}
    if hasattr(st, "query_params"):
        try:
            st.query_params.clear()
            st.query_params.update(clean)
            return
        except Exception:
            pass
    st.experimental_set_query_params(**clean)


def _qp_set(**updates: Optional[str]) -> None:
    params = _qp_read()
    for k, v in updates.items():
        if v in (None, ""):
            params.pop(str(k), None)
        else:
            params[str(k)] = str(v)
    _qp_write(params)


def _qp_get_first(key: str) -> Optional[str]:
    v = _qp_read().get(str(key))
    v = (v or "").strip()
    return v if v else None


@st.cache_resource
def _auth_token_secret() -> bytes:
    raw = os.environ.get("AUTH_TOKEN_SECRET")
    if not raw:
        try:
            raw = st.secrets.get("AUTH_TOKEN_SECRET")  # type: ignore[attr-defined]
        except Exception:
            raw = None
    if raw:
        return str(raw).encode("utf-8")
    return secrets.token_bytes(32)


def _logout() -> None:
    st.session_state.auth_user = None
    st.session_state.auth_token = None
    st.session_state.pop("active_project_id", None)
    st.session_state.pop("rename_project_id", None)
    st.session_state.pop("local_data_source", None)
    _qp_write({})
    st.rerun()


if "auth_user" not in st.session_state:
    st.session_state.auth_user = None
if "auth_token" not in st.session_state:
    st.session_state.auth_token = None

if _qp_get_first("logout") == "1":
    _logout()

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

    st.markdown(
        """
        <style>
          section[data-testid="stSidebar"] { display: none !important; }
          header[data-testid="stHeader"] { display: none !important; }
          div.block-container { max-width: 980px; padding-top: 48px; }
          div[data-testid="stVerticalBlockBorderWrapper"] {
            border-radius: 16px !important;
            border-color: #e5e7eb !important;
            box-shadow: 0 8px 22px rgba(16,24,40,.10) !important;
            background: #ffffff;
          }
          .login-hero h1 { font-size: 34px; margin: 0 0 6px 0; }
          .login-hero p { color: #475467; margin: 0 0 18px 0; }
        </style>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown(
            """
            <div class="login-hero">
              <h1>CARRIBOARD</h1>
              <p>Connectez-vous pour accéder à vos extractions et au tableau de bord.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.container(border=True):
            tab_login, tab_signup = st.tabs(["Connexion", "Inscription"])

            def _on_auth_success(user_obj: auth.User) -> None:
                st.session_state.auth_user = user_obj
                st.session_state.auth_token = None
                params = _qp_read()
                params.pop("t", None)
                params.pop("logout", None)
                _qp_write(params)
                st.session_state.pop("active_project_id", None)
                st.session_state.pop("rename_project_id", None)
                st.session_state.pop("local_data_source", None)
                st.rerun()

            with tab_login:
                with st.form("login_form_modern", clear_on_submit=False):
                    login_username = st.text_input("Nom d'utilisateur", key="login_username")
                    login_password = st.text_input("Mot de passe", type="password", key="login_password")
                    submit = st.form_submit_button("Se connecter", width="stretch", type="primary")
                if submit:
                    try:
                        user_obj = auth.authenticate(AUTH_DB_PATH, login_username, login_password)
                        _on_auth_success(user_obj)
                    except auth.AuthError as e:
                        st.error(str(e))

            with tab_signup:
                with st.form("signup_form_modern", clear_on_submit=False):
                    signup_username = st.text_input("Nom d'utilisateur", key="signup_username")
                    signup_password = st.text_input("Mot de passe (min 8)", type="password", key="signup_password")
                    signup_password2 = st.text_input("Confirmer", type="password", key="signup_password2")
                    submit2 = st.form_submit_button("Créer le compte", width="stretch", type="primary")
                if submit2:
                    if signup_password != signup_password2:
                        st.error("Les mots de passe ne correspondent pas.")
                    else:
                        try:
                            user_obj = auth.create_user(AUTH_DB_PATH, signup_username, signup_password)
                            _on_auth_success(user_obj)
                        except auth.AuthError as e:
                            st.error(str(e))

    st.stop()

# Bandeau (visible par tous les utilisateurs connectés)
st.markdown(
    """
    <div id="top-banner-wrap" style="
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
        color-scheme: light;
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

      /* Support du mode sombre navigateur/OS (Streamlit peut basculer) */
      @media (prefers-color-scheme: dark) {
        :root {
          color-scheme: dark;
          --bg: #0b1220;
          --surface: #0f172a;
          --text: #f9fafb;
          --muted: #cbd5e1;
          --border: rgba(148,163,184,0.25);
          --shadow1: 0 1px 2px rgba(0,0,0,.35);
          --shadow2: 0 12px 28px rgba(0,0,0,.45);
          --accent: #60a5fa;
          --success: #22c55e;
          --warning: #fbbf24;
        }
      }

      html { font-size: 15px; }
      body, .stApp {
        background: var(--bg);
        color: var(--text);
        font-family: Inter, Roboto, system-ui, -apple-system, "Segoe UI", Arial, sans-serif;
        line-height: 1.45;
      }
      *, *::before, *::after { box-sizing: border-box; }

      /* Force les zones principales à suivre nos variables (évite du texte invisible en thème sombre) */
      div[data-testid="stAppViewContainer"],
      div[data-testid="stHeader"],
      div[data-testid="stToolbar"] {
        background: var(--bg) !important;
        color: var(--text) !important;
      }
      section[data-testid="stSidebar"],
      section[data-testid="stSidebar"] > div {
        background: var(--surface) !important;
        color: var(--text) !important;
      }
      #top-banner-wrap {
        background: var(--surface) !important;
        border-color: var(--border) !important;
        color: var(--text) !important;
      }
      #top-banner-wrap * { color: var(--text) !important; }

      /* Inputs/Select : garde le contraste même en mode sombre */
      div[data-baseweb="input"] input,
      div[data-baseweb="textarea"] textarea,
      div[data-baseweb="select"] > div {
        background-color: var(--surface) !important;
        color: var(--text) !important;
        border-color: var(--border) !important;
      }
      div[data-baseweb="select"] svg { fill: var(--muted) !important; }
      button[role="tab"] { color: var(--muted) !important; }
      button[role="tab"][aria-selected="true"] { color: var(--text) !important; }

      /* Conteneur principal : largeur max + marges confortables */
      div.block-container {
        max-width: 1400px;
        padding-left: 24px;
        padding-right: 24px;
      }
      @media (max-width: 1366px) {
        div.block-container { padding-left: 16px; padding-right: 16px; }
      }

      /* Typographie */
      h1 { font-size: 1.6rem; line-height: 1.2; }
      h2 { font-size: 1.35rem; }
      h3 { font-size: 1.15rem; }
      p, li, label, .stMarkdown, .stText { font-size: 1rem; }
      small, .stCaption, .stCaptionContainer { font-size: 0.875rem; color: var(--muted); }

      /* Zones cliquables */
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

      /* Espacement des sections */
      hr { margin: 18px 0; border-color: var(--border); }

      /* Visuels des indicateurs */
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

      /* Grille des indicateurs (sur-mesure, évite la troncature) */
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

      /* DataFrames : rester dans le conteneur (pas de scroll horizontal global) */
      div[data-testid="stDataFrame"] { max-width: 100%; width: 100%; }
      div[data-testid="stDataFrame"] > div { max-width: 100%; width: 100%; overflow-x: auto; }
      div[data-testid="stDataFrame"] [role="gridcell"] > div {
        white-space: normal !important;
        overflow-wrap: anywhere;
        line-height: 1.2;
      }

      /* Helpers écran vs impression */
      .print-only { display: none; }
      .screen-only { display: inline; }
      @media print {
        .print-only { display: block !important; }
        .screen-only { display: none !important; }
        .kpi-grid { grid-template-columns: repeat(3, minmax(0, 1fr)); }
        div[data-testid="stDataFrame"] > div { overflow: visible !important; }
      }

      /* Tableaux adaptés à l'impression */
      .print-table { margin-top: 10px; overflow: visible; }
      .print-table table { width: 100%; border-collapse: collapse; table-layout: fixed; }
      .print-table th, .print-table td { border: 1px solid #d1d5db; padding: 4px 6px; font-size: 11px; }
      .print-table th { background: #f3f4f6; font-weight: 800; }
      .print-table td { word-break: break-word; }
      @media print {
        .print-table table { page-break-inside: auto; }
        .print-table thead { display: table-header-group; }
        .print-table tr { page-break-inside: avoid; page-break-after: auto; }
      }

      /* Cartes de métriques */
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

      /* Évite le scroll horizontal accidentel */
      .stApp { overflow-x: hidden; }
    </style>
    """,
    unsafe_allow_html=True,
)

components.html(
    """
    <script>
    (function () {
      try {
        const doc = window.parent && window.parent.document;
        if (!doc || !doc.body) return;

        const replacements = [
          // Import de fichier
          ["Drag and drop file here", "Glissez-déposez le fichier ici"],
          ["Drag and drop files here", "Glissez-déposez les fichiers ici"],
          ["Drag and drop file", "Glissez-déposez le fichier"],
          ["Drag and drop files", "Glissez-déposez les fichiers"],
          ["Browse files", "Parcourir les fichiers"],
          ["Browse file", "Parcourir les fichiers"],
          ["No file selected", "Aucun fichier sélectionné"],
          ["Remove file", "Retirer le fichier"],
          ["Remove", "Retirer"],
          ["Clear", "Effacer"],

          // DataFrame / info-bulles UI (Streamlit)
          ["Search", "Rechercher"],
          ["Download data as CSV", "Télécharger en CSV"],
          ["Download as CSV", "Télécharger en CSV"],
          ["Download CSV", "Télécharger en CSV"],
          ["Download", "Télécharger"],
          ["Copy", "Copier"],
          ["Fullscreen", "Plein écran"],
          ["Expand", "Agrandir"],

          // Barre d'outils Plotly (si affichée)
          ["Download plot as a png", "Télécharger le graphique en PNG"],
          ["Zoom in", "Zoom avant"],
          ["Zoom out", "Zoom arrière"],
          ["Pan", "Déplacer"],
          ["Zoom", "Zoom"],
          ["Autoscale", "Ajuster automatiquement"],
          ["Reset axes", "Réinitialiser les axes"],
          ["Box Select", "Sélection rectangle"],
          ["Lasso Select", "Sélection lasso"],
        ];

        function translateValue(v) {
          if (!v) return v;
          let out = String(v);
          for (const [a, b] of replacements) {
            if (out.includes(a)) out = out.split(a).join(b);
          }
          out = out.replace(/\bor\b/gi, "ou");
          out = out.replace(
            /Limit\\s+(\\d+(?:\\.\\d+)?\\s*[KMG]B)\\s+per\\s+file/gi,
            (m, sz) => {
              const fr = String(sz)
                .replace(/GB/gi, "Go")
                .replace(/MB/gi, "Mo")
                .replace(/KB/gi, "Ko");
              return `Limite : ${fr} par fichier`;
            }
          );
          out = out.replace(
            /Limit\\s+(\\d+(?:\\.\\d+)?)\\s*([KMG]B)\\s+per\\s+file/gi,
            (m, n, u) => {
              const unit = String(u)
                .replace(/GB/gi, "Go")
                .replace(/MB/gi, "Mo")
                .replace(/KB/gi, "Ko");
              return `Limite : ${n} ${unit} par fichier`;
            }
          );
          return out;
        }

        const SHOW_TEXT = (doc.defaultView && doc.defaultView.NodeFilter)
          ? doc.defaultView.NodeFilter.SHOW_TEXT
          : 4;

        function translateIn(root) {
          if (!root) return;
          const walker = doc.createTreeWalker(root, SHOW_TEXT);
          let node = null;
          while ((node = walker.nextNode())) {
            const t = node.nodeValue;
            if (!t || !t.trim()) continue;
            if (node.parentElement) {
              try {
                if (node.parentElement.closest("table, svg")) continue;
              } catch (e) {}
            }

            const out = translateValue(t);
            if (out !== t) node.nodeValue = out;
          }

          try {
            const attrEls = root.querySelectorAll("[title],[aria-label],[placeholder],[data-title]");
            for (const el of Array.from(attrEls)) {
              for (const attr of ["title", "aria-label", "placeholder", "data-title"]) {
                const v = el.getAttribute(attr);
                if (!v) continue;
                const out = translateValue(v);
                if (out !== v) el.setAttribute(attr, out);
              }
            }
          } catch (e) {}
        }

        function run() {
          try { translateIn(doc.body); } catch (e) {}
        }

        let scheduled = false;
        function schedule() {
          if (scheduled) return;
          scheduled = true;
          setTimeout(() => {
            scheduled = false;
            try { run(); } catch (e) {}
          }, 50);
        }

        schedule();
        try {
          const obs = new MutationObserver(schedule);
          obs.observe(doc.body, { childList: true, subtree: true, characterData: true });
        } catch (e) {}
      } catch (e) {}
    })();
    </script>
    """,
    height=0,
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
        ws_synth = workbook.add_worksheet("Synthèse - Tableau de bord")
        ws_synth.write(0, 0, "TABLEAU DE BORD - SYNTHÈSE DES LIVRAISONS", title_fmt)
        
        # Layout des indicateurs (2 lignes, 3 colonnes)
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
            
            # Barres de données pour le tonnage
            ws.conditional_format(3, 1, 3 + len(df_prod), 1, {
                'type': 'data_bar', 'bar_color': '#FFA07A'
            })
            # Barres de données pour le chiffre d'affaires
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
            
            # Barres de données
            ws.conditional_format(3, 1, 3 + len(df_client), 1, {'type': 'data_bar', 'bar_color': '#87CEEB'}) # Bleu (style Streamlit)
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
            
            # Échelle de couleurs (carte de chaleur)
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
                # Réordonner l'index
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


def _normalize_ticket_for_dedup(series: pd.Series) -> pd.Series:
    s = series.astype("string").str.strip()
    s = s.str.replace(r"\.0$", "", regex=True)
    s = s.str.casefold()
    return s.where(s.notna() & s.ne("") & s.ne("nan"))


def _drop_exact_duplicates(df: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    before = int(len(df))
    try:
        out = df.drop_duplicates(ignore_index=True)
    except TypeError:
        try:
            hashes = pd.util.hash_pandas_object(df, index=False)
            out = df.loc[~hashes.duplicated(keep="first")].copy()
            out.reset_index(drop=True, inplace=True)
        except Exception:
            out = df.copy()
    return out, before - int(len(out))


def _drop_duplicates_by_ticket(df: pd.DataFrame, ticket_col: str | None) -> tuple[pd.DataFrame, int]:
    if not ticket_col or ticket_col not in df.columns:
        return df, 0
    ticket = _normalize_ticket_for_dedup(df[ticket_col])
    dup_mask = ticket.notna() & ticket.duplicated(keep="first")
    removed = int(dup_mask.sum())
    if not removed:
        return df, 0
    out = df.loc[~dup_mask].copy()
    out.reset_index(drop=True, inplace=True)
    return out, removed


def _hash_rows(df: pd.DataFrame) -> pd.Series | None:
    try:
        return pd.util.hash_pandas_object(df, index=False)
    except Exception:
        return None


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
        # Ajout du numéro de semaine
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

    # Évite de propager les doublons dès l'import (lignes strictement identiques).
    df, _ = _drop_exact_duplicates(df)
        
    return df, col_mapping

dossier_app = os.path.dirname(os.path.abspath(__file__))

DEFAULT_EXTRACTION_SOURCE_NAME = "extraction pont bascule retraité.xlsx"


def _get_default_extraction_url() -> Optional[str]:
    raw = (os.environ.get("CARRIBOARD_DEFAULT_EXTRACTION_URL") or "").strip()
    if raw:
        return raw
    try:
        raw = str(st.secrets.get("CARRIBOARD_DEFAULT_EXTRACTION_URL") or "").strip()  # type: ignore[attr-defined]
    except Exception:
        raw = ""
    return raw or None


@st.cache_resource
def _resolve_default_extraction_path() -> Path:
    candidates = [
        Path(dossier_app) / "extraction pont bascule retraité.xlsx",
        Path(dossier_app) / "extraction pont bsacule retraité.xlsx",
    ]
    for p in candidates:
        try:
            if p.exists():
                return p
        except Exception:
            continue

    # Sur Streamlit Cloud, le fichier n'est pas commité (il est ignoré par .gitignore).
    # On le récupère via une URL (secret/env) et on le stocke dans DATA_ROOT.
    dst = DATA_ROOT / "default_extraction.xlsx"
    try:
        if dst.exists():
            return dst
    except Exception:
        pass

    url = _get_default_extraction_url()
    if url:
        try:
            r = requests.get(url, timeout=35)
            r.raise_for_status()
            content = r.content or b""
            if len(content) >= 1024:
                dst.write_bytes(content)
        except Exception:
            pass

    return dst if dst.exists() else candidates[0]


fichier_defaut = str(_resolve_default_extraction_path())

user = st.session_state.auth_user
st.sidebar.markdown(
    """
    <style>
      .sidebar-user {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 10px;
        padding: 12px 12px 10px 12px;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        background: #ffffff;
        box-shadow: 0 1px 2px rgba(16,24,40,.06);
      }
      .sidebar-user .u-name { font-weight: 800; color: #111827; line-height: 1.2; }
      .sidebar-user .u-role { color: #667085; font-size: 12px; margin-top: 2px; }
      .sidebar-user .u-logout a {
        color: #475467;
        font-size: 12px;
        text-decoration: none;
        padding: 6px 8px;
        border-radius: 10px;
        border: 1px solid transparent;
      }
      .sidebar-user .u-logout a:hover {
        background: #f2f4f7;
        border-color: #e5e7eb;
      }
    </style>
    """,
    unsafe_allow_html=True,
)
st.sidebar.markdown(
    f"""
    <div class="sidebar-user">
      <div>
        <div class="u-name">{html.escape(user.username)}</div>
        <div class="u-role">{html.escape(user.role)}</div>
      </div>
      <div class="u-logout">
        <a
          href="?logout=1"
          target="_self"
          onclick="window.location.href='?logout=1'; return false;"
        >Se déconnecter</a>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.sidebar.markdown("---")

st.sidebar.title(" Paramètres")
st.sidebar.markdown("---")
st.sidebar.caption(f"Données: {DATA_CLEAN_VERSION} • Script: {os.path.abspath(__file__)}")
if (os.environ.get("CARRIBOARD_DEBUG") or "").strip() == "1" or _qp_get_first("debug") == "1":
    with st.sidebar.expander("Debug", expanded=False):
        st.caption(f"DATA_ROOT: {DATA_ROOT}")
        st.caption(f"AUTH_DB_PATH: {AUTH_DB_PATH}")

PROJECT_DB_PATH = DATA_ROOT / "projects.sqlite3"
PROJECT_FILES_DIR = DATA_ROOT / "project_files"


@st.cache_resource
def _projects_db_ready(db_path: Path) -> None:
    projects.init_db(db_path)


_projects_db_ready(PROJECT_DB_PATH)

user_id = int(st.session_state.auth_user.id)

if "active_project_id" not in st.session_state:
    st.session_state.active_project_id = None
if "rename_project_id" not in st.session_state:
    st.session_state.rename_project_id = None
if "delete_project_id" not in st.session_state:
    st.session_state.delete_project_id = None
if "sidebar_upload_counter" not in st.session_state:
    st.session_state.sidebar_upload_counter = 0
if "top_upload_counter" not in st.session_state:
    st.session_state.top_upload_counter = 0
if "show_top_uploader" not in st.session_state:
    st.session_state.show_top_uploader = False
if "project_search" not in st.session_state:
    st.session_state.project_search = ""
if "close_extractions_popover" not in st.session_state:
    st.session_state.close_extractions_popover = False


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
    if not os.path.exists(fichier_defaut):
        return

    # Sur Streamlit Cloud, l'utilisateur peut déjà avoir des extractions :
    # on s'assure simplement que l'extraction "par défaut" existe aussi.
    try:
        default_abs = os.path.abspath(str(fichier_defaut))
        for p in (existing or []):
            try:
                if os.path.abspath(str(p.data_path)) == default_abs:
                    return
            except Exception:
                continue
    except Exception:
        pass

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
        name="Extraction pont bascule retraité",
        data_path=str(fichier_defaut),
        source_filename=DEFAULT_EXTRACTION_SOURCE_NAME,
        date_min=stats.get("date_min"),
        date_max=stats.get("date_max"),
        nb_livraisons=stats.get("nb_livraisons"),
        tonnage_total=stats.get("tonnage_total"),
        ca_total=stats.get("ca_total"),
        theme_idx=0,
    )
    # Ne pas forcer l'ouverture : l'utilisateur choisit d'abord l'extraction.


_ensure_default_project_if_needed()

# Import dans la barre latérale → crée une nouvelle extraction pour ce compte
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
    _qp_set(t=None, p=new_id)
    st.rerun()

# Liste des extractions (pour la barre du haut)
all_projects = projects.list_projects(PROJECT_DB_PATH, user_id)

# Après déploiement / 1ère ouverture, on sélectionne automatiquement l'extraction par défaut
# (pont bascule retraité) si l'utilisateur n'a rien sélectionné.
default_project_id: Optional[str] = None
try:
    default_abs = os.path.abspath(str(fichier_defaut))
    for p in (all_projects or []):
        try:
            if os.path.abspath(str(p.data_path)) == default_abs:
                default_project_id = p.id
                break
        except Exception:
            continue
except Exception:
    default_project_id = None

if not st.session_state.active_project_id:
    requested = _qp_get_first("p")
    if requested and projects.get_project(PROJECT_DB_PATH, user_id, requested) is not None:
        st.session_state.active_project_id = requested
    else:
        if requested:
            _qp_set(p=None)
        if default_project_id:
            st.session_state.active_project_id = default_project_id
active_project = projects.get_project(PROJECT_DB_PATH, user_id, st.session_state.active_project_id) if st.session_state.active_project_id else None

if not st.session_state.active_project_id or active_project is None:
    st.title("Extractions")
    st.caption("Choisissez une extraction pour ouvrir le tableau de bord.")

    if "picker_upload_counter" not in st.session_state:
        st.session_state.picker_upload_counter = 0

    search = st.text_input("Rechercher", value="", key="project_picker_search")
    q = (search or "").strip().casefold()
    filtered = [p for p in all_projects if not q or q in (p.name or "").casefold()]

    if filtered:
        cols = st.columns(2)
        for i, p in enumerate(filtered):
            date_range = ""
            if p.date_min and p.date_max:
                date_range = f"{p.date_min} → {p.date_max}"
            elif p.date_min:
                date_range = f"Depuis {p.date_min}"
            elif p.date_max:
                date_range = f"Jusqu'à {p.date_max}"

            tonnage = _fmt_number_fr(float(p.tonnage_total or 0.0), decimals=0) + " t"
            ca = _human_money(float(p.ca_total or 0.0))
            liv = str(int(p.nb_livraisons or 0))

            with cols[i % 2]:
                with st.container(border=True):
                    st.markdown(f"**{p.name}**")
                    if date_range:
                        st.caption(date_range)
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.metric("Livraisons", liv)
                    with c2:
                        st.metric("Tonnage", tonnage)
                    with c3:
                        st.metric("CA", ca)

                    if st.button("Ouvrir", key=f"open_project_{p.id}", type="primary", width="stretch"):
                        st.session_state.active_project_id = p.id
                        st.session_state.rename_project_id = None
                        _qp_set(t=None, p=p.id)
                        st.rerun()

        st.divider()
    else:
        st.info("Aucune extraction trouvée pour ce compte.")

    st.subheader("Importer une nouvelle extraction")
    picker_file = st.file_uploader(
        "Importer (Excel ou CSV)",
        type=["xls", "xlsx", "xlsm", "csv"],
        key=f"picker_uploader_{st.session_state.picker_upload_counter}",
    )
    if picker_file is not None:
        current = None
        next_theme_idx = 0
        new_id = _create_project_from_upload(picker_file, theme_idx=next_theme_idx)
        st.session_state.active_project_id = new_id
        st.session_state.rename_project_id = new_id
        st.session_state.picker_upload_counter += 1
        _qp_set(t=None, p=new_id)
        st.rerun()

    st.stop()

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
        with st.popover("Voir les extractions", width="stretch"):
            st.markdown(
                """
                <style>
                  /* Rendre les boutons icônes discrets */
                  div[data-testid="stPopoverBody"] button[kind="secondary"] {
                    color: #667085 !important;
                  }
                </style>
                """,
                unsafe_allow_html=True,
            )
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

                        a1, a2, a3 = st.columns([8, 1, 1])
                        with a1:
                            if st.button(
                                (p.name or "(Sans nom)") + (" (actif)" if is_active else ""),
                                key=f"pick_project_{p.id}",
                                type="primary" if is_active else "secondary",
                                width="stretch",
                            ):
                                st.session_state.active_project_id = p.id
                                st.session_state.rename_project_id = None
                                st.session_state.delete_project_id = None
                                st.session_state.close_extractions_popover = True
                                _qp_set(t=None, p=p.id)
                                st.rerun()
                        with a2:
                            if st.button("✎", key=f"rename_project_{p.id}", width="content", help="Renommer"):
                                st.session_state.active_project_id = p.id
                                st.session_state.rename_project_id = p.id
                                st.session_state.delete_project_id = None
                                st.session_state.close_extractions_popover = True
                                _qp_set(t=None, p=p.id)
                                st.rerun()
                        with a3:
                            if st.button("🗑", key=f"delete_project_{p.id}", width="content", help="Supprimer"):
                                st.session_state.delete_project_id = p.id
                                st.session_state.rename_project_id = None
                                st.session_state.close_extractions_popover = True
                                st.rerun()

    with btn_c2:
        if st.button("+ Nouvelle extraction", type="primary", width="stretch"):
            st.session_state.show_top_uploader = not bool(st.session_state.show_top_uploader)

if st.session_state.close_extractions_popover:
    components.html(
        """
        <script>
        (() => {
          const MAX_TRIES = 15;
          let tries = 0;

          function attemptClose() {
            tries += 1;
            try {
              const doc = window.parent?.document;
              if (!doc) return;

              const buttons = Array.from(doc.querySelectorAll("button"));
              const toggle = buttons.find((b) =>
                ((b.textContent || "").trim().includes("Voir les extractions")) &&
                (b.getAttribute("aria-haspopup") || b.getAttribute("aria-expanded") !== null)
              );

              if (toggle && toggle.getAttribute("aria-expanded") === "true") {
                toggle.click();
                return;
              }

              // Repli : tente de fermer le popover ouvert via la touche Échap.
              const esc = new KeyboardEvent("keydown", {
                key: "Escape",
                code: "Escape",
                keyCode: 27,
                which: 27,
                bubbles: true,
              });
              doc.dispatchEvent(esc);
            } catch (e) {}

            if (tries < MAX_TRIES) {
              setTimeout(attemptClose, 100);
            }
          }

          setTimeout(attemptClose, 0);
        })();
        </script>
        """,
        height=0,
    )
    st.session_state.close_extractions_popover = False

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
        _qp_set(t=None, p=new_id)
        st.rerun()

active_project = projects.get_project(PROJECT_DB_PATH, user_id, st.session_state.active_project_id) if st.session_state.active_project_id else None

if st.session_state.delete_project_id:
    target = projects.get_project(PROJECT_DB_PATH, user_id, st.session_state.delete_project_id)
    if target is None:
        st.session_state.delete_project_id = None
    else:
        st.error(f"Suppression d'extraction : **{target.name}**")
        st.caption("⚠️ Cette action est irréversible : l'extraction sera supprimée de la liste et ses fichiers uploadés seront effacés.")
        confirm_text = st.text_input("Tapez SUPPRIMER pour confirmer", value="", key=f"delete_confirm_{target.id}")
        dc1, dc2 = st.columns(2)
        with dc1:
            confirm_delete = st.button(
                "Supprimer définitivement",
                type="primary",
                width="stretch",
                disabled=(confirm_text or "").strip().casefold() != "supprimer",
            )
        with dc2:
            cancel_delete = st.button("Annuler", width="stretch")

        if confirm_delete:
            try:
                projects.delete_project(PROJECT_DB_PATH, user_id=user_id, project_id=target.id)
            except projects.ProjectError as e:
                st.error(str(e))
            else:
                # Nettoyage des fichiers importés quand ils sont stockés sous PROJECT_FILES_DIR
                try:
                    project_dir = PROJECT_FILES_DIR / f"user_{user_id}" / target.id
                    if project_dir.exists():
                        shutil.rmtree(project_dir, ignore_errors=True)
                except Exception:
                    pass

                if st.session_state.active_project_id == target.id:
                    st.session_state.active_project_id = None
                    st.session_state.pop("local_data_source", None)

                st.session_state.rename_project_id = None
                st.session_state.delete_project_id = None
                _qp_set(t=None, p=st.session_state.active_project_id)
                st.toast("Extraction supprimée.")
                st.rerun()

        if cancel_delete:
            st.session_state.delete_project_id = None
            st.rerun()

if active_project and st.session_state.rename_project_id == active_project.id:
    with st.form("rename_project_form", clear_on_submit=False):
        new_name = st.text_input("Nom du dossier", value=active_project.name)
        c1, c2 = st.columns(2)
        with c1:
            ok = st.form_submit_button("Enregistrer", width="stretch")
        with c2:
            cancel = st.form_submit_button("Annuler", width="stretch")
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
      const label = "PDF";
      const actionsId = "top-banner-actions";

       function ensureStyle() {{
         let style = window.parent.document.getElementById(STYLE_ID);
         if (!style) {{
           style = window.parent.document.createElement("style");
           style.id = STYLE_ID;
           window.parent.document.head.appendChild(style);
         }}
         style.textContent = `
           body[data-pdf-capture="1"] section[data-testid="stSidebar"] {{ display: none !important; }}
           body[data-pdf-capture="1"] header[data-testid="stHeader"] {{ display: none !important; }}
           body[data-pdf-capture="1"] div[data-testid="stToolbar"] {{ display: none !important; }}
           body[data-pdf-capture="1"] div[data-testid="stElementToolbar"] {{ display: none !important; }}
           body[data-pdf-capture="1"] div[data-testid="stDecoration"] {{ display: none !important; }}
           body[data-pdf-capture="1"] footer {{ display: none !important; }}
           body[data-pdf-capture="1"] #top-banner-wrap {{ display: none !important; }}
           body[data-pdf-capture="1"] [data-pdf-hide="1"] {{ display: none !important; }}
             body[data-pdf-capture="1"] #${{BTN_ID}} {{ display: none !important; }}
             body[data-pdf-capture="1"] div.block-container {{ padding-top: 0 !important; max-width: none !important; }}
             body[data-pdf-capture="1"] div[data-testid="stElementContainer"] {{ overflow: visible !important; }}
             body[data-pdf-capture="1"] div[data-testid="stHorizontalBlock"] {{ flex-wrap: wrap !important; gap: 18px !important; }}
             body[data-pdf-capture="1"] div[data-testid="column"] {{ width: 100% !important; flex: 1 1 100% !important; min-width: 0 !important; }}
             body[data-pdf-capture="1"] .print-only {{ display: block !important; }}
             body[data-pdf-capture="1"] .screen-only {{ display: none !important; }}
             /* PDF : graphiques + diagrammes (garde les tableaux \"print-only\") */
             body[data-pdf-capture="1"] div[data-testid="stDataFrame"] {{ display: none !important; }}
             body[data-pdf-capture="1"] div[data-testid="stTable"] {{ display: none !important; }}
             body[data-pdf-capture="1"] .print-table {{ display: block !important; }}

            body[data-pdf-capture="1"] div[data-testid="stTabs"] {{ overflow: visible !important; }}
            body[data-pdf-capture="1"] [data-baseweb="tab-panel"] {{ overflow: visible !important; }}
            /* PDF : afficher tous les onglets (pour inclure tous les graphiques) */
            body[data-pdf-capture="1"][data-pdf-all-tabs="1"] [data-pdf-scope="1"] div[data-testid="stTabs"] [role="tablist"] {{ display: none !important; }}
            body[data-pdf-capture="1"][data-pdf-all-tabs="1"] [data-pdf-scope="1"] [data-baseweb="tab-panel"] {{ display: block !important; }}
            body[data-pdf-capture="1"][data-pdf-all-tabs="1"] [data-pdf-scope="1"] [data-baseweb="tab-panel"][hidden] {{ display: block !important; }}
            /* Exception : certains switchers (ex: Tonnage/CA) ne doivent pas être "dépliés" */
            body[data-pdf-capture="1"][data-pdf-all-tabs="1"] [data-pdf-scope="1"] div[data-testid="stTabs"][data-pdf-no-expand="1"] [data-baseweb="tab-panel"][hidden] {{ display: none !important; }}
            body[data-pdf-capture="1"] div[data-testid="stPlotlyChart"], body[data-pdf-capture="1"] .stPlotlyChart {{ overflow: visible !important; }}
            body[data-pdf-capture="1"] div[data-testid="stPlotlyChart"] > div, body[data-pdf-capture="1"] .stPlotlyChart > div {{ overflow: visible !important; height: auto !important; max-height: none !important; }}
            body[data-pdf-capture="1"] .modebar, body[data-pdf-capture="1"] .plotly-notifier, body[data-pdf-capture="1"] .hoverlayer {{ display: none !important; }}
            /* Ne masque Plotly "live" que si les images statiques ont bien été créées */
            body[data-pdf-capture="1"][data-pdf-plots="1"] .print-plotly-wrapper {{ display: block !important; }}
            body[data-pdf-capture="1"][data-pdf-plots="1"] .js-plotly-plot {{ display: none !important; }}
            @media print {{
              @page {{ size: A4 landscape; margin: 10mm; }}
              #${{BTN_ID}} {{ display: none !important; }}
              section[data-testid="stSidebar"] {{ display: none !important; }}
              header[data-testid="stHeader"] {{ display: none !important; }}
             div[data-testid="stToolbar"] {{ display: none !important; }}
              div[data-testid="stElementToolbar"] {{ display: none !important; }}
               div[data-testid="stDecoration"] {{ display: none !important; }}
               footer {{ display: none !important; }}
               #top-banner-wrap {{ display: none !important; }}
              [data-pdf-hide="1"] {{ display: none !important; }}
              div.block-container {{ padding-top: 0 !important; max-width: none !important; }}
              div[data-testid="stElementContainer"] {{ overflow: visible !important; }}
              * {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
              div[data-testid="stHorizontalBlock"] {{ flex-wrap: wrap !important; gap: 18px !important; }}
              div[data-testid="column"] {{ width: 100% !important; flex: 1 1 100% !important; min-width: 0 !important; }}
              /* PDF : graphiques uniquement */
              div[data-testid="stDataFrame"], div[data-testid="stTable"] {{ display: none !important; }}
              div[data-testid="stTabs"], [data-baseweb="tab-panel"] {{ overflow: visible !important; }}
              div[data-testid="stTabs"] [role="tablist"] {{ display: none !important; }}
              /* PDF : afficher tous les onglets (pour inclure tous les graphiques) */
              body[data-pdf-all-tabs="1"] [data-pdf-scope="1"] div[data-testid="stTabs"] [role="tablist"] {{ display: none !important; }}
              body[data-pdf-all-tabs="1"] [data-pdf-scope="1"] [data-baseweb="tab-panel"] {{ display: block !important; }}
              body[data-pdf-all-tabs="1"] [data-pdf-scope="1"] [data-baseweb="tab-panel"][hidden] {{ display: block !important; }}
              /* Exception : certains switchers (ex: Tonnage/CA) ne doivent pas être "dépliés" */
              body[data-pdf-all-tabs="1"] [data-pdf-scope="1"] div[data-testid="stTabs"][data-pdf-no-expand="1"] [data-baseweb="tab-panel"][hidden] {{ display: none !important; }}
              div[data-testid="stPlotlyChart"], .stPlotlyChart, .js-plotly-plot {{
                break-inside: avoid !important;
                page-break-inside: avoid !important;
              }}
              .print-plotly-wrapper img {{
                max-height: 175mm !important;
                object-fit: contain !important;
              }}
             div[data-testid="stPlotlyChart"] > div, .stPlotlyChart > div {{
               overflow: visible !important;
               height: auto !important;
               max-height: none !important;
             }}
             .modebar, .plotly-notifier, .hoverlayer {{ display: none !important; }}
              /* Si des images statiques Plotly ont été préparées, on imprime celles-ci. */
              body[data-pdf-plots="1"] .print-plotly-wrapper {{ display: block !important; }}
              body[data-pdf-plots="1"] .js-plotly-plot {{ display: none !important; }}
           }}
            .print-plotly-wrapper {{ display: none; width: 100%; }}
            .print-plotly-wrapper .print-plot-caption {{
              margin: 0 0 6px 0;
              font-size: 12px;
              font-weight: 700;
              color: #344054;
            }}
            .print-plotly-wrapper img {{
              width: 100% !important;
              max-width: 100% !important;
              height: auto !important;
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
          btn.title = "Imprimer / Enregistrer en PDF (graphiques et diagrammes). Pour supprimer l'URL en bas : désactivez « En-têtes et pieds de page » dans l'impression.";

           const state = {{
             busy: false,
             wrappers: [],
           }};

           function getStreamlitRootWindow() {{
             const wins = [];
             try {{
               let w = window;
               for (let i = 0; i < 10; i += 1) {{
                 wins.push(w);
                 let p = null;
                 try {{ p = w.parent; }} catch (e) {{ p = null; }}
                 if (!p || p === w) break;
                 // Stoppe si cross-origin / inaccessible
                 try {{ void p.document; }} catch (e) {{ break; }}
                 w = p;
               }}
             }} catch (e) {{
               wins.push(window);
             }}

             // Choisit la fenêtre dont le DOM contient vraiment l'app Streamlit
             for (let i = wins.length - 1; i >= 0; i -= 1) {{
               try {{
                 const d = wins[i].document;
                 if (d && d.querySelector && d.querySelector(\"div[data-testid='stMain']\")) return wins[i];
               }} catch (e) {{}}
             }}
             // Fallback raisonnable
             try {{
               if (window.parent && window.parent.document) return window.parent;
             }} catch (e) {{}}
             return window;
           }}

           function cleanupPdf(doc) {{
             try {{ doc.body.removeAttribute("data-pdf-plots"); }} catch (e) {{}}
             try {{ doc.body.removeAttribute("data-pdf-all-tabs"); }} catch (e) {{}}
             try {{ doc.body.removeAttribute("data-pdf-capture"); }} catch (e) {{}}
             try {{
               Array.from(doc.querySelectorAll("[data-pdf-scope='1']")).forEach((el) => {{
                 try {{ el.removeAttribute("data-pdf-scope"); }} catch (e) {{}}
               }});
             }} catch (e) {{}}
             try {{
               Array.from(doc.querySelectorAll("div[data-testid='stTabs'][data-pdf-no-expand='1']")).forEach((el) => {{
                 try {{ el.removeAttribute("data-pdf-no-expand"); }} catch (e) {{}}
               }});
             }} catch (e) {{}}
            try {{
              Array.from(doc.querySelectorAll("[data-pdf-hide='1']")).forEach((el) => {{
                try {{ el.removeAttribute("data-pdf-hide"); }} catch (e) {{}}
              }});
            }} catch (e) {{}}
            try {{
              state.wrappers.forEach((w) => {{ try {{ w.remove(); }} catch (e) {{}} }});
              state.wrappers = [];
            }} catch (e) {{}}
            try {{
              Array.from(doc.querySelectorAll(".js-plotly-plot")).forEach((p) => {{
                try {{
                  if (p.dataset && p.dataset.pdfPrepared) delete p.dataset.pdfPrepared;
                }} catch (e) {{}}
              }});
             }} catch (e) {{}}

             // Cache la section "Données" (brutes/filtrées) : jamais dans le PDF
             try {{
               const norm = (s) => {{
                 try {{
                   return String(s || "")
                     .normalize("NFD")
                     .replace(/[\\u0300-\\u036f]/g, "")
                     .replace(/\\s+/g, " ")
                     .trim()
                     .toLowerCase();
                 }} catch (e) {{
                   return String(s || "").replace(/\\s+/g, " ").trim().toLowerCase();
                 }}
               }};

               // 1) cache l'onglet (st.tabs) des données brutes/filtrées
               const tabs = Array.from(doc.querySelectorAll("div[data-testid='stTabs']"));
               for (const t of tabs) {{
                 const txt = norm(t.textContent || "");
                 const isRawTabs = txt.includes("brutes (complet)") && txt.includes("filtrees");
                 if (isRawTabs) t.setAttribute("data-pdf-hide", "1");
               }}

               // 2) cache le titre "Données" juste au-dessus (si présent)
               const heads = Array.from(doc.querySelectorAll("h1, h2, h3, h4, h5"));
               for (const h of heads) {{
                 if (norm(h.textContent || "") === "donnees") {{
                   const box = h.closest("div[data-testid='stElementContainer']") || h.closest("div");
                   if (box) box.setAttribute("data-pdf-hide", "1");
                 }}
               }}
             }} catch (e) {{}}
           }}

          function sleep(ms) {{
            return new Promise((resolve) => setTimeout(resolve, ms));
          }}

          async function waitForVisiblePlots(doc, rootWin) {{
            try {{
              for (let i = 0; i < 12; i += 1) {{
                const plots = Array.from(doc.querySelectorAll(".js-plotly-plot"));
                let visible = 0;
                let ready = 0;
                for (const p of plots) {{
                  const r = p.getBoundingClientRect ? p.getBoundingClientRect() : null;
                  if (!r || r.width < 50 || r.height < 50) continue;
                  visible += 1;
                  if (r.width >= 240 && r.height >= 240) ready += 1;
                }}
                if (visible === 0 || ready === visible) return;
                try {{ rootWin.dispatchEvent(new Event("resize")); }} catch (e) {{}}
                await sleep(120);
              }}
            }} catch (e) {{}}
          }}

          async function ensureDashboardTab(doc, rootWin) {{
            try {{
              const tabs = Array.from(doc.querySelectorAll("[role='tab']"));
              const dash = tabs.find((t) =>
                String(t.textContent || "").replace(/\\s+/g, " ").trim().toLowerCase().includes("tableau de bord")
              );
              if (dash && dash.getAttribute("aria-selected") !== "true") {{
                dash.click();
                await sleep(220);
                try {{ rootWin.dispatchEvent(new Event("resize")); }} catch (e) {{}}
                await sleep(160);
              }}
            }} catch (e) {{}}
          }}

          async function warmUpAllTabsInScope(doc, rootWin) {{
            // Objectif : cliquer automatiquement sur tous les onglets (st.tabs)
            // afin que Streamlit rende tous les contenus (et donc tous les graphiques) au moins une fois.
            try {{
              const scopes = Array.from(doc.querySelectorAll("[data-pdf-scope='1']"));
              const roots = scopes.length ? scopes : [doc.querySelector("div[data-testid='stMain']") || doc.body];

              const seen = new Set();
              const norm = (s) => {{
                try {{
                  return String(s || "")
                    .normalize("NFD")
                    .replace(/[\\u0300-\\u036f]/g, "")
                    .replace(/[^a-zA-Z0-9]+/g, " ")
                    .replace(/\\s+/g, " ")
                    .trim()
                    .toLowerCase();
                }} catch (e) {{
                  return String(s || "").replace(/\\s+/g, " ").trim().toLowerCase();
                }}
              }};

              // Réveille les onglets principaux (Tableau de bord / Analyses détaillées) si Streamlit rend les panels à la demande.
              const clickTopTab = async (needle) => {{
                try {{
                  const allTabs = Array.from(doc.querySelectorAll("[role='tab']"));
                  const t = allTabs.find((el) => norm(el.textContent).includes(needle));
                  if (t && t.getAttribute("aria-selected") !== "true") {{
                    t.click();
                    await sleep(320);
                    try {{ rootWin.dispatchEvent(new Event("resize")); }} catch (e) {{}}
                    await sleep(180);
                  }}
                }} catch (e) {{}}
              }};
              await clickTopTab("analyses detaillees");
              await clickTopTab("tableau de bord");

              for (const root of roots) {{
                const tablists = Array.from(root.querySelectorAll("[role='tablist']"));
                for (const tablist of tablists) {{
                  const tabs = Array.from(tablist.querySelectorAll("[role='tab']"));

                  // Cas spécial : onglet switcher "Tonnage / Chiffre d'affaires" (Graphique 3)
                  // -> on ne force pas le rendu des 2 pour éviter l'impression en double.
                  try {{
                    if (tabs.length === 2) {{
                      const a = norm(tabs[0]?.textContent);
                      const b = norm(tabs[1]?.textContent);
                      const hasTonnage = a.includes("tonnage") || b.includes("tonnage");
                      const hasCa = a.includes("chiffre d affaires") || b.includes("chiffre d affaires");
                      if (hasTonnage && hasCa) continue;
                    }}
                  }} catch (e) {{}}

                  for (const t of tabs) {{
                    const key = t.getAttribute("aria-controls") || t.id || (t.textContent || "");
                    const sk = String(key || "");
                    if (!sk) continue;
                    // Évite de recliquer indéfiniment la même tab (plusieurs scopes peuvent la contenir).
                    if (seen.has(sk)) continue;
                    seen.add(sk);

                    try {{
                      if (t.getAttribute("aria-selected") !== "true") t.click();
                    }} catch (e) {{}}

                    await sleep(260);
                    try {{ rootWin.dispatchEvent(new Event("resize")); }} catch (e) {{}}
                    await sleep(160);
                  }}
                }}
              }}
            }} catch (e) {{}}
          }}

          function getMainNode(doc) {{
            return (
              doc.querySelector("div.block-container") ||
              doc.querySelector("div[data-testid='stMain']") ||
               doc.body
             );
           }}

           function getAvoidRanges(main) {{
             const mainRect = main.getBoundingClientRect();
             const selectors = [
               ".print-plotly-wrapper",
               "div[data-testid='stDataFrame']",
               "div[data-testid='stPlotlyChart']",
               "div[data-testid='stTabs']",
             ].join(",");
             const ranges = [];
             try {{
               for (const el of Array.from(main.querySelectorAll(selectors))) {{
                 const r = el.getBoundingClientRect();
                 const top = Math.max(0, Math.floor(r.top - mainRect.top));
                 const bottom = Math.max(top + 1, Math.ceil(r.bottom - mainRect.top));
                 if (bottom - top < 60) continue;
                 ranges.push({{ top, bottom }});
               }}
             }} catch (e) {{
               try {{
                 const rw = window.top || window.parent || window;
                 const msg = String((e && e.message) || e || "");
                 if (msg.includes("popup_blocked")) {{
                   try {{ rw.alert("Fenetre bloquee par le navigateur. Autorisez les popups pour generer le PDF."); }} catch (e2) {{}}
                 }} else {{
                   try {{ rw.alert("Erreur lors de la generation du PDF. Actualisez la page et reessayez."); }} catch (e2) {{}}
                 }}
               }} catch (e2) {{}}
             }}
             ranges.sort((a, b) => a.top - b.top);
             return ranges;
           }}

           function adjustedSliceHeight(y, baseH, totalH, ranges) {{
             let h = Math.min(baseH, totalH - y);
             const cut = y + h;
             const minH = 220;
             const pad = 14;
             try {{
               for (const r of ranges) {{
                 if (r.top + pad < cut && r.bottom - pad > cut) {{
                   const before = Math.max(0, r.top - y - pad);
                   if (before >= minH) {{
                     h = Math.min(h, before);
                   }} else {{
                     const after = Math.max(0, r.bottom - y + pad);
                     const maxGrow = Math.max(minH, Math.floor(baseH * 1.35));
                     if (after <= maxGrow) h = Math.max(h, after);
                   }}
                   break;
                 }}
                 try {{ pw.close(); }} catch (e) {{}}
                 try {{ rootWin.alert("Aucun graphique n'a pu etre exporte. Actualisez la page et reessayez."); }} catch (e) {{}}
               }}
             }} catch (e) {{}}
             h = Math.min(h, totalH - y);
             return Math.max(minH, Math.floor(h));
           }}

           function loadScriptOnce(doc, src, checkFn) {{
             return new Promise((resolve, reject) => {{
               try {{
                 if (checkFn && checkFn()) return resolve(true);
                 const existing = Array.from(doc.querySelectorAll("script")).find((s) => s.src === src);
                if (existing) {{
                  existing.addEventListener("load", () => resolve(true), {{ once: true }});
                  setTimeout(() => resolve(true), 1500);
                  return;
                }}
                const s = doc.createElement("script");
                s.src = src;
                s.async = true;
                s.onload = () => resolve(true);
                s.onerror = () => reject(new Error("Failed to load " + src));
                doc.head.appendChild(s);
              }} catch (e) {{
                reject(e);
              }}
            }});
          }}

          async function ensurePdfLibs(doc, rootWin) {{
            const hasHtml2Canvas = () => !!rootWin.html2canvas;
            const hasJsPdf = () => !!(rootWin.jspdf && rootWin.jspdf.jsPDF);

            await loadScriptOnce(
              doc,
              "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js",
              hasHtml2Canvas
            );
            await loadScriptOnce(
              doc,
              "https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js",
              hasJsPdf
            );

            return {{
              html2canvas: rootWin.html2canvas,
              jsPDF: rootWin.jspdf.jsPDF,
            }};
          }}

            function _getPdfScopeRoots(doc) {{
              try {{
                const roots = Array.from(doc.querySelectorAll("[data-pdf-scope='1']"));
                return roots.length ? roots : [doc.body];
              }} catch (e) {{
                return [doc.body];
              }}
            }}

            function getPdfVisualInfo(doc) {{
              const roots = _getPdfScopeRoots(doc);
              const view = doc?.defaultView || window;
              const isVisible = (el) => {{
                try {{
                  const cs = view.getComputedStyle ? view.getComputedStyle(el) : null;
                  if (cs && (cs.display === "none" || cs.visibility === "hidden")) return false;
                  const r = el.getBoundingClientRect ? el.getBoundingClientRect() : null;
                  if (!r) return true;
                  if (r.width < 50 || r.height < 50) return false;
                  return true;
                }} catch (e) {{
                  return true;
                }}
              }};

              let plotly = 0;
              try {{
                const els = roots.flatMap((r) => Array.from(r.querySelectorAll(".js-plotly-plot")));
                plotly = els.filter(isVisible).length;
              }} catch (e) {{}}

              let other = 0;
             const otherSelectors = [
               "div[data-testid='stVegaLiteChart']",
               "div[data-testid='stPyplot']",
               "div[data-testid='stGraphvizChart']",
               "div[data-testid='stDeckGlJsonChart']",
                "div[data-testid='stImage'] img",
               ".print-table",
             ];
              for (const sel of otherSelectors) {{
                try {{
                  const els = roots.flatMap((r) => Array.from(r.querySelectorAll(sel)));
                  if (els.some(isVisible)) other += 1;
                }} catch (e) {{}}
              }}

              return {{ plotly, other, total: plotly + other }};
            }}

            async function preparePlotlyImages(doc, rootWin) {{
              let Plotly = rootWin?.Plotly || doc?.defaultView?.Plotly || window?.Plotly;
              if (!Plotly) {{
                // Streamlit charge parfois Plotly après le rendu des éléments : on attend un peu.
                for (let k = 0; k < 18; k += 1) {{
                  try {{ await sleep(120); }} catch (e) {{}}
                  Plotly = rootWin?.Plotly || doc?.defaultView?.Plotly || window?.Plotly;
                  if (Plotly) break;
                }}
              }}
              if (!Plotly || !Plotly.toImage) return false;

              const scopes = Array.from(doc.querySelectorAll("[data-pdf-scope='1']"));
              const plots = scopes.length
                ? scopes.flatMap((el) => Array.from(el.querySelectorAll(".js-plotly-plot")))
                : Array.from(doc.querySelectorAll(".js-plotly-plot"));
              let created = 0;
              const createdImgs = [];
              const seenUrls = new Set();

              const escapeSel = (s) => {{
                try {{
                  if (window.CSS && CSS.escape) return CSS.escape(String(s));
                }} catch (e) {{}}
                return String(s).replace(/[^a-zA-Z0-9_\\-]/g, "\\\\$&");
              }};

              const getTabContextLabel = (plot) => {{
                try {{
                  const panel = plot.closest("[data-baseweb='tab-panel']");
                  if (!panel || !panel.id) return "";
                  const q = `[role='tab'][aria-controls='${{escapeSel(panel.id)}}']`;
                  const tab = doc.querySelector(q);
                  const txt = String(tab ? tab.textContent : "").replace(/\\s+/g, " ").trim();
                  const low = txt.toLowerCase();
                  if (!txt) return "";
                  // Ignore les grands onglets principaux (inutile de répéter sur chaque graphique)
                  if (low.includes("tableau de bord")) return "";
                  if (low.includes("analyses détaillées") || low.includes("analyses detaillees")) return "";
                  if (low.includes("mise à jour") || low.includes("mise a jour")) return "";
                  return txt;
                }} catch (e) {{
                  return "";
                }}
              }};

              for (const plot of plots) {{
                // Évite les doublons pendant la même exécution.
                if (plot.dataset && plot.dataset.pdfPrepared === "1") continue;

                try {{
                  // N'imprime que les graphiques visibles sur la page courante.
                  try {{
                    const view = doc?.defaultView || rootWin || window;
                    const cs = view.getComputedStyle ? view.getComputedStyle(plot) : null;
                    if (cs && (cs.display === "none" || cs.visibility === "hidden")) continue;
                    const rect0 = plot.getBoundingClientRect ? plot.getBoundingClientRect() : null;
                    if (rect0 && (rect0.width < 50 || rect0.height < 50)) continue;
                  }} catch (e) {{}}

                  try {{
                    if (Plotly.Plots && Plotly.Plots.resize) await Plotly.Plots.resize(plot);
                  }} catch (e) {{}}

                  let w = null;
                  let h = null;
                  try {{
                    const fl = plot._fullLayout;
                    if (fl && fl.width && fl.height) {{
                      w = Math.max(200, Math.round(fl.width));
                      h = Math.max(200, Math.round(fl.height));
                    }}
                  }} catch (e) {{}}

                  if (!w || !h) {{
                    const rect = plot.getBoundingClientRect ? plot.getBoundingClientRect() : null;
                    if (rect && (rect.width < 50 || rect.height < 50)) continue;
                    w = rect && rect.width ? Math.max(200, Math.round(rect.width)) : null;
                    h = rect && rect.height ? Math.max(200, Math.round(rect.height)) : null;
                  }}
                 const opts = {{ format: "png", scale: 2 }};
                 if (w && h) {{
                   opts.width = w;
                   opts.height = h;
                 }}

                  const url = await Plotly.toImage(plot, opts);
                  if (!url) continue;
                  // Évite les doublons (certains onglets peuvent produire la même image).
                  if (seenUrls.has(url)) {{
                    try {{ if (plot.dataset) plot.dataset.pdfPrepared = "1"; }} catch (e) {{}}
                    continue;
                  }}
                  seenUrls.add(url);

                  const wrapper = doc.createElement("div");
                  wrapper.className = "print-plotly-wrapper";
                  wrapper.style.breakInside = "avoid";
                  wrapper.style.pageBreakInside = "avoid";

                  const ctx = getTabContextLabel(plot);
                  if (ctx) {{
                    const cap = doc.createElement("div");
                    cap.className = "print-plot-caption";
                    cap.textContent = ctx;
                    wrapper.appendChild(cap);
                  }}

                  const img = doc.createElement("img");
                  img.src = url;
                  img.alt = "Graphique";
                  wrapper.appendChild(img);
                 createdImgs.push(img);

                 const parent = plot.parentElement;
                 if (parent) {{
                   parent.insertBefore(wrapper, plot);
                   state.wrappers.push(wrapper);
                   if (plot.dataset) plot.dataset.pdfPrepared = "1";
                   created += 1;
                 }}
               }} catch (e) {{}}
             }}

             if (created > 0) {{
               try {{
                 await Promise.all(
                   createdImgs.map(async (img) => {{
                     try {{
                       if (img.decode) await img.decode();
                       else if (!img.complete) await new Promise((r) => {{ img.onload = () => r(true); setTimeout(() => r(true), 500); }});
                     }} catch (e) {{}}
                   }})
                 );
               }} catch (e) {{}}
               try {{ doc.body.setAttribute("data-pdf-plots", "1"); }} catch (e) {{}}
               return true;
             }}
             return false;
           }}

           async function collectPlotlyImages(doc, rootWin, state, btn) {{
             let Plotly = rootWin?.Plotly || doc?.defaultView?.Plotly || window?.Plotly;
             if (!Plotly) {{
               // Streamlit charge parfois Plotly après le rendu des éléments : on attend un peu.
               for (let k = 0; k < 18; k += 1) {{
                 try {{ await sleep(120); }} catch (e) {{}}
                 Plotly = rootWin?.Plotly || doc?.defaultView?.Plotly || window?.Plotly;
                 if (Plotly) break;
               }}
             }}
             if (!Plotly || !Plotly.toImage) return [];

             const view = doc?.defaultView || rootWin || window;
             const isVisible = (el) => {{
               try {{
                 const r = el.getBoundingClientRect ? el.getBoundingClientRect() : null;
                 if (!r || r.width < 50 || r.height < 50) return false;
                 const cs = view.getComputedStyle ? view.getComputedStyle(el) : null;
                 if (!cs) return true;
                 if (cs.display === "none" || cs.visibility === "hidden") return false;
                 return true;
               }} catch (e) {{
                 return true;
               }}
             }};

             const plots = Array.from(doc.querySelectorAll("div[data-testid='stMain'] .js-plotly-plot")).filter(isVisible);
             const out = [];

             for (let i = 0; i < plots.length; i += 1) {{
               const plot = plots[i];
               if (!plot) continue;

               try {{
                 if (btn) btn.innerText = `Export (${{
                   Math.min(i + 1, plots.length)
                 }}/${{
                   plots.length
                 }})…`;

                 try {{
                   if (Plotly.Plots && Plotly.Plots.resize) await Plotly.Plots.resize(plot);
                 }} catch (e) {{}}

                 let w = 1600;
                 let h = 1000;
                 try {{
                   const r = plot.getBoundingClientRect();
                   if (r && r.width && r.height) {{
                     w = Math.max(w, Math.round(r.width));
                     h = Math.max(h, Math.round(r.height));
                   }}
                 }} catch (e) {{}}
                 try {{
                   const fl = plot._fullLayout;
                   if (fl && fl.width && fl.height) {{
                     w = Math.max(w, Math.round(fl.width));
                     h = Math.max(h, Math.round(fl.height));
                   }}
                 }} catch (e) {{}}
                 w = Math.min(2400, w);
                 h = Math.min(1600, h);

                 let url = null;
                 try {{
                   url = await Plotly.toImage(plot, {{ format: "png", scale: 2, width: w, height: h }});
                 }} catch (e) {{}}
                 if (!url) {{
                   try {{ await sleep(160); }} catch (e) {{}}
                   try {{
                     url = await Plotly.toImage(plot, {{ format: "png", scale: 2, width: w, height: h }});
                   }} catch (e) {{}}
                 }}
                 if (!url) {{
                   try {{ await sleep(220); }} catch (e) {{}}
                   try {{
                     url = await Plotly.toImage(plot, {{ format: "png", scale: 2, width: w, height: h }});
                   }} catch (e) {{}}
                 }}
                 if (!url) continue;

                let title = "";
                try {{
                  const fl = plot._fullLayout;
                  const t = fl && fl.title && fl.title.text ? String(fl.title.text) : "";
                  title = (t || "").replace(/<[^>]*>/g, "").trim();
                }} catch (e) {{}}

                out.push({{ title, url }});
                try {{ await sleep(40); }} catch (e) {{}}
              }} catch (e) {{}}
            }}

            if (btn) btn.innerText = "PDF";
            return out;
          }}

           function markPdfHiddenNodes(doc) {{
            // Cache la barre "Extraction active / Voir les extractions / + Nouvelle extraction"
            try {{
              const blocks = Array.from(doc.querySelectorAll("div[data-testid='stHorizontalBlock']"));
              for (const block of blocks) {{
                const t = (block.textContent || "").replace(/\\s+/g, " ").trim().toLowerCase();
                const isProjectBar =
                  t.includes("extraction active") &&
                  (t.includes("voir les extractions") || t.includes("nouvelle extraction"));
                if (isProjectBar) block.setAttribute("data-pdf-hide", "1");
              }}
            }} catch (e) {{}}

            // Repli : repère les boutons et cache leur ligne.
            try {{
              const btns = Array.from(doc.querySelectorAll("button"));
              for (const b of btns) {{
                const txt = ((b.textContent || "").replace(/\\s+/g, " ").trim().toLowerCase());
                if (!txt) continue;
                if (txt.includes("voir les extractions") || txt.includes("nouvelle extraction")) {{
                  const row = b.closest("div[data-testid='stHorizontalBlock']");
                  if (row) row.setAttribute("data-pdf-hide", "1");
                }}
              }}
            }} catch (e) {{}}
          }}

          function setPdfScopeToDashboardPanel(doc) {{
            try {{
              // Nettoie l'ancien scope (si relance).
              Array.from(doc.querySelectorAll("[data-pdf-scope='1']")).forEach((el) => {{
                try {{ el.removeAttribute("data-pdf-scope"); }} catch (e) {{}}
              }});
            }} catch (e) {{}}

            try {{
              const main = doc.querySelector("div[data-testid='stMain']") || doc.body;
              const norm = (s) => {{
                try {{
                  return String(s || "")
                    .normalize("NFD")
                    .replace(/[\\u0300-\\u036f]/g, "")
                    .replace(/\\s+/g, " ")
                    .trim()
                    .toLowerCase();
                }} catch (e) {{
                  return String(s || "").replace(/\\s+/g, " ").trim().toLowerCase();
                }}
              }};

              // Repère le "tablist" principal (Tableau de bord / Analyses détaillées / Mise à jour…)
              const tablists = Array.from(main.querySelectorAll("[role='tablist']"));
              let best = null;
              let bestScore = -1;
              for (const tl of tablists) {{
                const tabs = Array.from(tl.querySelectorAll("[role='tab']"));
                if (tabs.length < 2) continue;
                const texts = tabs.map((t) => norm(t.textContent));
                const hasDash = texts.some((t) => t.includes("tableau de bord"));
                const hasAnalyses = texts.some((t) => t.includes("analyses detaillees") || t.includes("analyses détaillées"));
                const hasUpdate = texts.some((t) => t.includes("mise a jour") || t.includes("mise à jour") || t.includes("tableur"));
                const score = (hasDash ? 1 : 0) + (hasAnalyses ? 1 : 0) + (hasUpdate ? 1 : 0);
                if (score > bestScore) {{
                  bestScore = score;
                  best = tl;
                }}
              }}
              const fallback = () => {{
                try {{
                  const root =
                    doc.querySelector("div[data-testid='stMain'] div.block-container") ||
                    doc.querySelector("div[data-testid='stMain']") ||
                    doc.body;
                  if (root) root.setAttribute("data-pdf-scope", "1");
                  return true;
                }} catch (e) {{}}
                return false;
              }};
              if (!best || bestScore <= 0) return fallback();

              const tabs = Array.from(best.querySelectorAll("[role='tab']"));
              const selected = tabs.find((t) => t.getAttribute("aria-selected") === "true") || tabs[0];
              const panelId = selected ? selected.getAttribute("aria-controls") : null;
              const panel = panelId ? doc.getElementById(panelId) : null;
              if (!panel) return fallback();

              panel.setAttribute("data-pdf-scope", "1");
              return true;
            }} catch (e) {{}}

            return false;
          }}

          function markPdfNoExpandTabGroups(doc) {{
            const norm = (s) => {{
              try {{
                return String(s || "")
                  .normalize("NFD")
                  .replace(/[\\u0300-\\u036f]/g, "")
                  .replace(/[^a-zA-Z0-9]+/g, " ")
                  .replace(/\\s+/g, " ")
                  .trim()
                  .toLowerCase();
              }} catch (e) {{
                return String(s || "").replace(/\\s+/g, " ").trim().toLowerCase();
              }}
            }};

            try {{
              // Nettoyage (si relance)
              Array.from(doc.querySelectorAll("div[data-testid='stTabs'][data-pdf-no-expand='1']")).forEach((el) => {{
                try {{ el.removeAttribute("data-pdf-no-expand"); }} catch (e) {{}}
              }});
            }} catch (e) {{}}

            try {{
              const scopes = Array.from(doc.querySelectorAll("[data-pdf-scope='1']"));
              const roots = scopes.length ? scopes : [doc.querySelector("div[data-testid='stMain']") || doc.body];
              for (const root of roots) {{
                const tablists = Array.from(root.querySelectorAll("[role='tablist']"));
                for (const tablist of tablists) {{
                  const tabs = Array.from(tablist.querySelectorAll("[role='tab']"));
                  if (tabs.length !== 2) continue;
                  const a = norm(tabs[0]?.textContent);
                  const b = norm(tabs[1]?.textContent);
                  const hasTonnage = a.includes("tonnage") || b.includes("tonnage");
                  const hasCa = a.includes("chiffre d affaires") || b.includes("chiffre d affaires");
                  if (!(hasTonnage && hasCa)) continue;
                  const stTabs = tablist.closest("div[data-testid='stTabs']");
                  if (stTabs) stTabs.setAttribute("data-pdf-no-expand", "1");
                }}
              }}
            }} catch (e) {{}}
          }}

          function _stripSensitiveParamsFromUrl(rootWin) {{
            try {{
              const original = String(rootWin.location.href || "");
              const url = new URL(original);
              const keys = ["t", "p", "token", "auth", "session"];
              let changed = false;
              for (const k of keys) {{
                if (url.searchParams.has(k)) {{
                  url.searchParams.delete(k);
                  changed = true;
                }}
              }}
              if (!changed) return null;
              rootWin.history.replaceState(null, "", url.toString());
              return original;
            }} catch (e) {{
              return null;
            }}
          }}

          function _restoreUrl(rootWin, original) {{
            try {{
              if (!original) return;
              rootWin.history.replaceState(null, "", String(original));
            }} catch (e) {{}}
          }}

           btn.onclick = async function() {{
             if (state.busy) return;
             state.busy = true;

             const prevLabel = btn.innerText;
             btn.innerText = "Préparation…";
             btn.disabled = true;

              try {{
                const rootWin = getStreamlitRootWindow();
                const doc = rootWin.document || window.parent.document;
                 const main = getMainNode(doc);

                // Mode "capture" (style tableau de bord) :
                // - masque les éléments UI superflus
                // - remplace les graphiques Plotly par des images statiques
                // - ouvre la fenêtre d'impression du navigateur
               cleanupPdf(doc);
               markPdfHiddenNodes(doc);
                try {{ doc.body.setAttribute("data-pdf-capture", "1"); }} catch (e) {{}}
                const inDash = setPdfScopeToDashboardPanel(doc);
                if (!inDash) {{
                  try {{ cleanupPdf(doc); }} catch (e) {{}}
                  try {{ rootWin.alert("Impossible de repérer les zones à imprimer. Actualisez la page et réessayez."); }} catch (e) {{}}
                  return;
                }}
                try {{ rootWin.dispatchEvent(new Event("resize")); }} catch (e) {{}}
                await sleep(220);
                await waitForVisiblePlots(doc, rootWin);

                btn.innerText = "Préparation de l'impression…";
                const vis = getPdfVisualInfo(doc);
                if (!vis || !vis.total) {{
                  try {{ cleanupPdf(doc); }} catch (e) {{}}
                  try {{ rootWin.alert("Aucun graphique ou diagramme détecté sur la page actuelle."); }} catch (e) {{}}
                  return;
                }}

                // Si la page contient des graphiques Plotly, on doit les convertir en images
                // sinon ils risquent de ne pas apparaître à l'impression.
                if (vis.plotly > 0) {{
                  const ok = await preparePlotlyImages(doc, rootWin);
                  if (!ok) {{
                    try {{ cleanupPdf(doc); }} catch (e) {{}}
                    try {{ rootWin.alert("Impossible d'exporter les graphiques. Actualisez la page et réessayez."); }} catch (e) {{}}
                    return;
                  }}
                }}

                // Le navigateur peut ajouter l'URL en pied de page à l'impression :
                // on retire les paramètres sensibles (ex: t=...) juste le temps de l'impression.
                const originalUrl = _stripSensitiveParamsFromUrl(rootWin);
                const after = () => {{
                  try {{ _restoreUrl(rootWin, originalUrl); }} catch (e) {{}}
                  try {{ cleanupPdf(doc); }} catch (e) {{}}
                }};

                try {{ rootWin.addEventListener("afterprint", after, {{ once: true }}); }} catch (e) {{}}
                try {{ rootWin.focus(); }} catch (e) {{}}
                try {{ rootWin.print(); }} catch (e) {{ after(); }}
                setTimeout(after, 15000);
                return;
              }} catch (e) {{
                const rw = window.top || window.parent || window;
                const msg = String((e && e.message) || e || "");
                if (msg.includes("popup_blocked")) {{
                  try {{ rw.alert("Fenêtre bloquée par le navigateur. Autorisez les popups pour générer le PDF."); }} catch (e2) {{}}
                }} else {{
                  try {{ rw.alert("Erreur lors de la génération du PDF. Actualisez la page et réessayez."); }} catch (e2) {{}}
                }}
              }} finally {{
                btn.disabled = false;
                btn.innerText = prevLabel;
                state.busy = false;
              }}
           }};
        }}
        btn.style.position = "relative";
        btn.style.background = "transparent";
        btn.style.border = "1px solid #d0d5dd";
        btn.style.color = "#667085";
        btn.style.padding = "6px 10px";
        btn.style.fontSize = "12px";
        btn.style.fontWeight = "800";
        btn.style.borderRadius = "10px";
        btn.style.boxShadow = "none";
        btn.style.cursor = "pointer";
        btn.style.whiteSpace = "nowrap";
        btn.style.zIndex = "9999";
        btn.onmouseenter = function() {{ btn.style.background = "#f2f4f7"; }};
        btn.onmouseleave = function() {{ btn.style.background = "transparent"; }};

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
        
        # Ajoute un filtre "Semaine" pour coller au visuel Excel
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

        # En-tête
        st.title(" Tableau de Bord - Suivi des Livraisons")

        ca_col = mapping.get('ca')
        poids_col = mapping.get('poids')
        client_col = mapping.get('client')
        produit_col = mapping.get('produit')
        date_col = mapping.get('date')

        # --- Indicateurs clés (global) ---
        st.markdown("<div class='kpi-title'>Indicateurs clés (global)</div>", unsafe_allow_html=True)
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
        
        st.markdown("---")
        
        # --- VIEWS (Excel vs Dash) ---
        main_tab_dash, main_tab_excel, main_tab_update = st.tabs(
            ["Tableau de bord", "Analyses détaillées", "✏️ Mise à jour du tableur"]
        )
        
        # --- VUE GRAPHIQUES ORIGINALE ---
        with main_tab_dash:
            st.subheader(" Indicateurs & graphiques")
            
            # Reprise des analyses avancées par indicateurs
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
                         st.plotly_chart(apply_plotly_style(fig1), width="stretch")

                # Graphique 2 : Évolution mensuelle du Tonnage (Histogramme)
                with row1_col2:
                    if not df_evol_m.empty:
                         fig2 = px.bar(df_evol_m, x=date_col, y=poids_col, 
                                       title="Graphique 2: Évolution Mensuelle du Tonnage (T)",
                                       color_discrete_sequence=[theme["secondary"]])
                         fig2.update_layout(xaxis_title="Mois", yaxis_title="Tonnage (T)")
                         st.plotly_chart(apply_plotly_style(fig2), width="stretch")

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
                            fig3_t.update_traces(textposition="inside", textinfo="percent")
                            fig3_t.update_layout(margin=dict(t=8, b=24, l=8, r=8), uniformtext_minsize=10, uniformtext_mode="hide")
                            st.plotly_chart(apply_plotly_style(fig3_t, kind="pie"), width="stretch")
                        with tab_p2:
                            fig3_ca = px.pie(
                                df_prod_pie,
                                values=ca_col,
                                names=produit_col,
                                hole=0.4,
                                color_discrete_sequence=theme["pie"],
                            )
                            fig3_ca.update_traces(textposition="inside", textinfo="percent")
                            fig3_ca.update_layout(margin=dict(t=8, b=24, l=8, r=8), uniformtext_minsize=10, uniformtext_mode="hide")
                            st.plotly_chart(apply_plotly_style(fig3_ca, kind="pie"), width="stretch")

                # Graphique 4 : Top 10 Clients (Barre horizontale)
                with row2_col2:
                    if client_col and ca_col and poids_col:
                        df_client_top = df.groupby(client_col).agg({ca_col: 'sum', poids_col: 'sum'}).reset_index()
                        df_client_top = df_client_top.sort_values(by=ca_col, ascending=False).head(10)
                        df_client_top = df_client_top.sort_values(by=ca_col, ascending=True) # Inversion pour barres horizontales
                        
                        fig4 = px.bar(df_client_top, x=ca_col, y=client_col, orientation='h',
                                      title="Graphique 4: Top 10 Clients (CA)", text=ca_col,
                                      color_discrete_sequence=[theme["accent"]],
                                      hover_data=[poids_col])
                        fig4.update_traces(texttemplate='%{text:.2s} €', textposition='outside')
                        fig4.update_layout(xaxis_title="CA (€)", yaxis_title="", margin=dict(l=150))
                        st.plotly_chart(apply_plotly_style(fig4), width="stretch")

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
                        st.plotly_chart(apply_plotly_style(fig5), width="stretch")

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
                        st.plotly_chart(apply_plotly_style(fig6), width="stretch")
            except Exception as e:
                st.write(f"Erreur d'affichage des graphiques : {e}")

        # --- VUE MISE À JOUR DU TABLEUR ---
        with main_tab_update:
            st.subheader(" Mise à jour du tableur (ajout de nouvelles lignes)")
            st.info(
                "Importez un Excel contenant de nouveaux enregistrements. "
                "En mode fichier local (par défaut), le fichier source sera mis à jour et les graphiques se recalculeront automatiquement."
            )

            st.markdown('<div id="update-anchor"></div>', unsafe_allow_html=True)
            if st.session_state.pop("after_update_focus", False):
                components.html(
                    """
                    <script>
                    (function () {
                      try {
                        const doc = window.parent && window.parent.document;
                        if (!doc) return;

                        const tryClickTab = () => {
                          const tabs = Array.from(doc.querySelectorAll('[role="tab"]'));
                          const target = tabs.find((t) => (t.innerText || "").includes("Mise à jour du tableur"));
                          if (target) target.click();
                        };

                        const tryScroll = () => {
                          const el = doc.getElementById("update-anchor");
                          if (el && el.scrollIntoView) el.scrollIntoView({ behavior: "smooth", block: "start" });
                        };

                        tryClickTab();
                        setTimeout(tryScroll, 60);
                        setTimeout(tryClickTab, 120);
                        setTimeout(tryScroll, 220);
                      } catch (e) {}
                    })();
                    </script>
                    """,
                    height=0,
                )

            notice = st.session_state.pop("after_update_notice", None)
            if notice:
                st.success(notice)

            warn = st.session_state.pop("after_update_warning", None)
            if warn:
                st.warning(warn)

            details_after = st.session_state.pop("after_update_details", None)
            if details_after:
                with st.expander("Détails de la mise à jour", expanded=False):
                    st.write(details_after)

            if isinstance(data_source, (str, Path)):
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
                    "Supprimer les doublons par Ticket (si le ticket est unique)",
                    value=False,
                    help="Recommandé si le ticket identifie une livraison unique. Désactivez si vous pensez que le ticket peut se répéter légitimement.",
                    key="update_dedup_ticket",
                )

                submitted = st.form_submit_button("🚀 Mettre à jour", type="primary", width="stretch")

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

                            rows_read = int(len(new_df))
                            ticket_col = mapping.get("ticket")
                            date_col = mapping.get("date")
                            mois_col = mapping.get("mois")

                            import_date_min = None
                            import_date_max = None
                            if date_col and date_col in new_df.columns:
                                try:
                                    dt_import = pd.to_datetime(new_df[date_col], errors="coerce")
                                    import_date_min = dt_import.min()
                                    import_date_max = dt_import.max()
                                except Exception:
                                    pass

                            import_months = None
                            if mois_col and mois_col in new_df.columns:
                                try:
                                    mois_vals = new_df[mois_col].dropna().astype(str).unique().tolist()
                                    import_months = sorted(mois_vals, key=lambda v: str(v).casefold())
                                except Exception:
                                    pass

                            removed_exact_import = 0
                            removed_ticket_import = 0
                            skipped_exact_existing = 0
                            skipped_ticket_existing = 0

                            # 1) Doublons internes au fichier importé
                            new_df, removed_exact_import = _drop_exact_duplicates(new_df)
                            if dedup_by_ticket:
                                new_df, removed_ticket_import = _drop_duplicates_by_ticket(new_df, ticket_col)

                            # 2) Évite d'ajouter ce qui existe déjà dans l'extraction
                            if len(new_df):
                                if dedup_by_ticket and ticket_col and ticket_col in raw_full_df.columns and ticket_col in new_df.columns:
                                    base_ticket = _normalize_ticket_for_dedup(raw_full_df[ticket_col])
                                    base_ticket_set = set(base_ticket.dropna().unique().tolist())
                                    new_ticket = _normalize_ticket_for_dedup(new_df[ticket_col])
                                    mask_ticket_exists = new_ticket.isin(base_ticket_set)
                                    skipped_ticket_existing = int(mask_ticket_exists.sum())
                                    if skipped_ticket_existing:
                                        new_df = new_df.loc[~mask_ticket_exists].copy()
                                        new_df.reset_index(drop=True, inplace=True)

                                base_hash = _hash_rows(raw_full_df)
                                new_hash = _hash_rows(new_df)
                                if base_hash is not None and new_hash is not None:
                                    mask_exact_exists = new_hash.isin(base_hash)
                                    skipped_exact_existing = int(mask_exact_exists.sum())
                                    if skipped_exact_existing:
                                        new_df = new_df.loc[~mask_exact_exists].copy()
                                        new_df.reset_index(drop=True, inplace=True)

                            rows_added = int(len(new_df))

                            updated_full_df = pd.concat([raw_full_df, new_df], ignore_index=True)

                            # 3) Nettoyage final après fusion
                            updated_full_df, removed_exact_after = _drop_exact_duplicates(updated_full_df)
                            removed_ticket_after = 0
                            if dedup_by_ticket:
                                updated_full_df, removed_ticket_after = _drop_duplicates_by_ticket(updated_full_df, ticket_col)

                            data_source_path: Path | None = None
                            if isinstance(data_source, (str, Path)):
                                try:
                                    data_source_path = Path(data_source)
                                except Exception:
                                    data_source_path = None

                            if not (data_source_path and data_source_path.exists()):
                                st.error(
                                    "Impossible de mettre à jour le fichier source dans ce mode (fichier uploadé). "
                                    "Utilisez le fichier par défaut local ou relancez l'app avec un fichier sur disque."
                                )
                            else:
                                saved_path = None
                                save_note = None
                                try:
                                    updated_full_df.to_excel(data_source_path, index=False)
                                    saved_path = str(data_source_path)
                                except PermissionError:
                                    src = data_source_path
                                    alt = src.with_name(f"{src.stem} - MAJ {datetime.now():%Y%m%d-%H%M%S}{src.suffix}")
                                    updated_full_df.to_excel(alt, index=False)
                                    st.session_state.local_data_source = str(alt)
                                    saved_path = str(alt)
                                    save_note = (
                                        "Impossible d'écraser le fichier source (souvent parce qu'il est ouvert dans Excel). "
                                        "Une copie mise à jour a été enregistrée."
                                    )

                                msg = f"{rows_added} lignes ajoutées."

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

                                details = {
                                    "Fichier enregistré": saved_path,
                                    "Lignes lues (fichier importé)": int(rows_read),
                                    "Lignes ajoutées": int(rows_added),
                                    "Lignes totales (après fusion)": int(len(updated_full_df)),
                                    "Doublons supprimés (import, lignes identiques)": int(removed_exact_import),
                                    "Doublons ignorés (déjà présents, lignes identiques)": int(skipped_exact_existing),
                                    "Doublons supprimés (après fusion, lignes identiques)": int(removed_exact_after),
                                }
                                if dedup_by_ticket:
                                    details["Doublons supprimés (import, Ticket)"] = int(removed_ticket_import)
                                    details["Doublons ignorés (Ticket déjà présent)"] = int(skipped_ticket_existing)
                                    details["Doublons supprimés (Ticket, après fusion)"] = int(removed_ticket_after)
                                if import_date_min is not None:
                                    details["Date min (import)"] = str(import_date_min)
                                if import_date_max is not None:
                                    details["Date max (import)"] = str(import_date_max)
                                if import_months is not None:
                                    details["Mois (import)"] = import_months

                                st.cache_data.clear()
                                st.session_state.after_update_notice = msg
                                if save_note:
                                    st.session_state.after_update_warning = f"{save_note} Fichier : {saved_path}"
                                st.session_state.after_update_details = details
                                st.session_state.after_update_focus = True
                                st.rerun()
                    except Exception as e:
                        st.exception(e)

        
        with main_tab_excel:
            # --- Analyses avancées ---
            st.subheader(" Analyses détaillées")
            tab_client, tab_produit, tab_livraison = st.tabs([" Analyse par Produit/Client", " Évolution des Livraisons", " Performance Clients"])
            
            # Onglet 1 : Analyse par produit (tableaux à barres intégrées)
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
                                width="stretch"
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
                            # Aligné sur le Graphique 3 (camembert avec %)
                            fig_pie = px.pie(
                                df_prod,
                                values=poids_col,
                                names=produit_col,
                                hole=0.4,
                                color_discrete_sequence=theme["pie"],
                            )
                            fig_pie.update_traces(textposition="inside", textinfo="percent")
                            fig_pie.update_layout(showlegend=True, margin=dict(t=8, b=24, l=8, r=8), uniformtext_minsize=10, uniformtext_mode="hide")
                            st.plotly_chart(apply_plotly_style(fig_pie, kind="pie"), width="stretch")

            # Onglet 2 : Évolution des livraisons (courbes croisées)
            with tab_produit:
                if date_col and produit_col and poids_col and mapping.get('mois'):
                    st.markdown("**Analyse des quantités livrées par mois et par produit**")
                    
                    # Ordonner correctement les mois
                    months_order = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre']
                    
                    # Calcul des totaux par mois et par produit
                    if 'Semaine' in df.columns:
                        # L'utilisateur peut agréger par semaine ou par mois (par défaut : mois)
                        group_col = mapping['mois']
                        
                        df_evol_prod = df.groupby([group_col, produit_col])[poids_col].sum().reset_index()
                        
                        # Assure l'ordre catégoriel des mois
                        if group_col == mapping['mois']:
                            # Filtrer pour ne conserver que des libellés valides
                            df_evol_prod[group_col] = df_evol_prod[group_col].str.lower()
                            # Trier
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
                        st.plotly_chart(apply_plotly_style(fig_lines), width="stretch")

            # Onglet 3 : Performance clients (analyse matricielle)
            with tab_livraison:
                if client_col and ca_col and poids_col and produit_col:
                    st.markdown("**Performance Globale par Client**")
                    
                    df_client_perf = df.groupby(client_col).agg({poids_col: 'sum', ca_col: 'sum'}).reset_index()
                    df_client_perf = df_client_perf.sort_values(by=poids_col, ascending=False)
                    
                    col_c1, col_c2 = st.columns([1, 2])
                    with col_c1:
                        # Graphiques en barres pour les 20 meilleurs clients
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
                                width="stretch"
                            )
                            st.markdown(
                                f"<div class='print-only print-table'>{top_clients.to_html(index=False)}</div>",
                                unsafe_allow_html=True,
                            )
                        
                    with col_c2:
                        # Matrice client × produit
                        st.markdown("**Livraison par produit (en tonnes)**")
                        pivot_df = pd.pivot_table(df, values=poids_col, index=client_col, columns=produit_col, aggfunc='sum', fill_value=0)
                        pivot_df['Total'] = pivot_df.sum(axis=1)
                        pivot_df = pivot_df.sort_values(by='Total', ascending=False)
                        # Limite aux 20 premiers (comme le tableau de gauche) pour garder l'alignement
                        pivot_df_top = pivot_df.head(20) 
                        pivot_df_top = pivot_df_top.drop(columns=['Total'])
                        
                        # Style du tableau façon Excel
                        st.dataframe(
                            pivot_df_top.style.format("{:,.2f}").background_gradient(cmap=theme["table_cmap"], axis=None),
                            height=600,
                            width="stretch",
                        )
                        st.markdown(
                            f"<div class='print-only print-table'>{pivot_df_top.to_html()}</div>",
                            unsafe_allow_html=True,
                        )
                        
                        # Ajoute un histogramme empilé sous la matrice
                        st.markdown("**Graphique de Livraison par Produit (Top 20 Clients)**")
                        # Prépare les données pour l'histogramme empilé
                        # Repasse la table pivot en format long pour Plotly
                        melted_df = pivot_df_top.reset_index().melt(id_vars=client_col, value_vars=pivot_df_top.columns, var_name=produit_col, value_name=poids_col)
                        # Retire les zéros pour garder un graphique lisible
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
                        st.plotly_chart(apply_plotly_style(fig_stacked), width="stretch")

        # --- Base de données Brute ---
        st.markdown("---")
        st.subheader("Données")
        tab_raw, tab_filtered = st.tabs(["Brutes (complet)", "Filtrées (selon filtres)"])
        with tab_raw:
            st.dataframe(raw_full_df, width="stretch")
        with tab_filtered:
            st.dataframe(df, width="stretch")
            
    else:
        st.warning("Veuillez charger un fichier Excel (.xls, .xlsx) depuis le panneau de gauche.")
        
except Exception as e:
    st.error(f"Une erreur s'est produite lors de l'analyse du fichier : {e}")
