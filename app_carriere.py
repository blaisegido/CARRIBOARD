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
from streamlit_cookies_controller import CookieController


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


def _pie_percent_text(values: "pd.Series", *, decimals: int = 0) -> list[str]:
    try:
        v = pd.to_numeric(values, errors="coerce").fillna(0.0)
        total = float(v.sum())
        if not math.isfinite(total) or total <= 0:
            pcts = [0.0] * int(len(v))
        else:
            pcts = (v / total).astype(float).tolist()
        fmt = "{:." + str(int(decimals)) + "%}"
        return [fmt.format(float(p)) for p in pcts]
    except Exception:
        return [""] * int(len(values))


def _clean_pie_labels(labels: "pd.Series", *, fallback: str = "Non renseigné") -> "pd.Series":
    try:
        s = labels.fillna("").astype(str).str.strip()
        bad = s.eq("") | s.str.casefold().isin({"undefined", "nan", "none", "<na>", "null"})
        return s.where(~bad, fallback)
    except Exception:
        return labels


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




cookie_controller = CookieController()
AUTH_TOKEN_STORAGE_KEY = "carriboard_session_token"

def _logout() -> None:
    tok = st.session_state.get("auth_token")
    if tok and hasattr(auth, "revoke_session_token"):
        try:
            auth.revoke_session_token(AUTH_DB_PATH, str(tok))
        except Exception:
            pass
    st.session_state.auth_user = None
    st.session_state.auth_token = None
    st.session_state._clear_auth_local_storage = True
    st.session_state.pop("active_project_id", None)
    st.session_state.pop("rename_project_id", None)
    st.session_state.pop("local_data_source", None)
    _qp_write({})
    st.rerun()


if "auth_user" not in st.session_state:
    st.session_state.auth_user = None
if "auth_token" not in st.session_state:
    st.session_state.auth_token = None
if "_clear_auth_local_storage" not in st.session_state:
    st.session_state._clear_auth_local_storage = False

if _qp_get_first("logout") == "1":
    _logout()

if st.session_state.auth_user is None:
    token = _qp_get_first("t") or cookie_controller.get(AUTH_TOKEN_STORAGE_KEY)
    if token and hasattr(auth, "verify_session_token") and hasattr(auth, "get_user"):
        try:
            uid = auth.verify_session_token(token, _auth_token_secret(), db_path=AUTH_DB_PATH)  # type: ignore[call-arg]
        except TypeError:
            uid = auth.verify_session_token(token, _auth_token_secret())

        if uid is not None:
            restored = auth.get_user(AUTH_DB_PATH, uid)
            if restored is not None:
                st.session_state.auth_user = restored
                st.session_state.auth_token = token
                _qp_set(t=None, autologin=None, autologin_failed=None, logout=None)
                st.rerun()
        else:
            # Token invalide : évite les boucles d'auto-login.
            if _qp_get_first("autologin") == "1":
                st.session_state._clear_auth_local_storage = True
                _qp_set(t=None, autologin=None, autologin_failed="1")
                st.rerun()

