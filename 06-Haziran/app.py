"""CK Enerji EDAŞ fatura veri çıkarma ve analiz portalı."""

from __future__ import annotations

import contextlib
import base64
import html
import hashlib
import io
import json
import tempfile
import zipfile
import math
import os
import re
import sys
import warnings
from pathlib import Path
from typing import Any, Iterable

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

SYSTEM_USAGE_FILE_DEFAULT = BASE_DIR / "_PortalAdmin_Uploads_Content_FastAccess_0bd2da5220923.xlsx"
SYSTEM_USAGE_FILE_PREFERENCE = BASE_DIR / "config" / "system_usage_file.json"
ENO_FILE_PREFERENCE = BASE_DIR / "config" / "eno_file.json"





# Veri çıkarma için doğrulanmış ana çekirdek kullanılır.
from extract_and_compare import (  # noqa: E402
    clean_turkish_number,
    enrich_from_eno_file,
    normalize_region_name,
    process_all_regions,
    save_to_excel,
)
from core.upload_inputs import UploadInputError, UploadWorkspace  # noqa: E402


def uses_browser_file_inputs() -> bool:
    """Use browser uploads on hosted/Linux servers and native dialogs locally."""
    configured_mode = os.environ.get("SKF_FILE_INPUT_MODE", "auto").strip().casefold()
    if configured_mode in {"browser", "upload", "web"}:
        return True
    if configured_mode in {"native", "local"}:
        return False
    return os.name != "nt"


st.set_page_config(
    page_title="CK Enerji - EDAŞ Fatura Veri Portalı 1.2",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── LAN Güvenliği: IP Whitelist ──────────────────────────────────────────
# Yalnız özel ağ (private network) IP'lerinden erişime izin verir.
# Bu kontrol, sunucu 0.0.0.0 ile LAN'a açıldığında aktif olur.
# KVKK uyumluluğu için ETSO verilerine dış ağlardan erişim engellenir.
_ALLOWED_PREFIXES = ('127.', '10.', '192.168.', '172.16.', '172.17.',
                     '172.18.', '172.19.', '172.20.', '172.21.',
                     '172.22.', '172.23.', '172.24.', '172.25.',
                     '172.26.', '172.27.', '172.28.', '172.29.',
                     '172.30.', '172.31.', '::1', 'localhost')

def _check_ip_whitelist():
    """Streamlit'in query params üzerinden IP kontrolü (bilgilendirme amaçlı)."""
    if uses_browser_file_inputs():
        return
    try:
        from streamlit.web.server.websocket_headers import _get_websocket_headers
        headers = _get_websocket_headers()
        if headers:
            client_ip = headers.get('X-Forwarded-For', headers.get('Host', ''))
            if client_ip and not any(client_ip.startswith(p) for p in _ALLOWED_PREFIXES):
                st.error(f"⛔ Erişim engellendi: {client_ip} — Yalnız yerel ağ erişimi desteklenir.")
                st.stop()
    except Exception:
        pass  # Streamlit sürümü header'ları desteklemiyorsa sessizce geç

_check_ip_whitelist()
# ─────────────────────────────────────────────────────────────────────────


EXPECTED_REGION_COUNT = 21
ANOMALY_MINIMUM_TL = 10_000_000.0
ABSURD_CORRUPTION_THRESHOLD_TL = 50_000_000_000.0
OUR_INVOICES_LABEL = "Bizim Kestiğimiz Faturalar"



MONEY_FIELDS = {
    "Dağıtım Bedeli(TL)",
    "Dağıtım Bedeli (TL)",
    "Güç Bedeli(TL)",
    "Güç Bedeli (TL)",
    "Güç Aşım Bedeli (TL)",
    "Reaktif Bedel (TL)",
    "İlk Reaktif",
    "İlk Reaktif Bedeli (TL)",
    "Standart Dışı Tutar (TL)",
    "Tazminat Bedeli",
    "KDV",
    "Toplam (TL)",
}

QUANTITY_FIELDS = {
    "Güç kW",
    "Güç (kW)",
    "KURULU GÜÇ",
    "Kurulu Güç",
    "Aktif Enerji Tüketim (kWh)",
}

RAW_DATA_COLUMNS = [
    "Dağıtım Bölgesi",
    "Etso Kodu",
    "Müşteri",
    "Tarife Grubu",
    "AG OG",
    "TERİM",
    "Güç kW",
    "KURULU GÜÇ",
    "Aktif Enerji Tüketim (kWh)",
    "Dağıtım Bedeli(TL)",
    "Güç Bedeli(TL)",
    "Güç Aşım Bedeli (TL)",
    "Reaktif Bedel (TL)",
    "İlk Reaktif",
    "Sayax'a Atılacak Tarife",
]

OPTIONAL_RAW_DATA_COLUMNS = [
    "Standart Dışı Tutar (TL)",
    "Tazminat Bedeli",
]


def inject_copy_shortcut_guard() -> None:
    """Streamlit C kısayolunu durdururken tarayıcının kopyalama işlevini koru."""
    components.html(
        """
        <script>
        (() => {
            const parentDocument = window.parent.document;
            if (parentDocument.__ckEnerjiCopyGuardInstalled) {
                return;
            }
            const copyGuard = (event) => {
                const isCopy = (event.ctrlKey || event.metaKey) &&
                    (event.key === 'c' || event.key === 'C');
                if (isCopy) {
                    event.stopImmediatePropagation();
                    event.stopPropagation();
                }
            };
            parentDocument.addEventListener('keydown', copyGuard, true);
            parentDocument.__ckEnerjiCopyGuardInstalled = true;
        })();
        </script>
        """,
        height=0,
    )


def inject_corporate_css() -> None:
    """CK Enerji renkleriyle sade ve kurumsal Streamlit teması uygula."""
    st.markdown(
        """
        <style>
        :root {
            --ck-blue: #305496;
            --ck-orange: #ED7D31;
            --ck-green: #70AD47;
            --ck-navy-gray: #44536A;
            --ck-surface: #F5F7FA;
        }
        [data-testid="stAppViewContainer"] {
            background: linear-gradient(180deg, #FFFFFF 0%, #F7F9FC 100%);
        }
        [data-testid="stMainBlockContainer"] {
            max-width: none;
            padding-top: 3rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }
        [data-testid="stSidebar"] {
            background-color: #F5F7FA;
            border-right: 1px solid #D8DEE9;
        }
        [data-testid="stSidebar"] [data-testid="stImage"] {
            margin-top: -0.65rem !important;
            margin-bottom: -0.2rem !important;
        }
        .ck-selection-check {
            display: inline-flex;
            width: 1.45rem;
            height: 1.45rem;
            align-items: center;
            justify-content: center;
            border-radius: 50%;
            background: #E8F3E3;
            color: #3E7D23;
            font-size: 0.95rem;
            font-weight: 800;
            line-height: 1;
        }
        div[class*="st-key-select_system_usage_file"] button {
            width: 2.8rem !important;
            min-width: 2.8rem !important;
            height: 2.8rem !important;
            padding: 0 !important;
            justify-content: center !important;
        }
        .ck-title {
            color: var(--ck-blue);
            font-size: 2.05rem;
            font-weight: 750;
            line-height: 1.15;
            margin: 0.15rem 0 0.2rem 0;
        }
        .ck-subtitle {
            color: var(--ck-navy-gray);
            font-size: 1rem;
            margin-bottom: 0.25rem;
        }
        .ck-section-title {
            color: var(--ck-navy-gray);
            font-size: 1.5rem;
            font-weight: 650;
            line-height: 1.3;
            margin: 0;
        }
        .stMainBlockContainer div[data-testid="stElementContainer"]:has(.ck-section-title) {
            margin-top: 2.2rem;
            position: relative;
            top: -0.5cm;
        }
        .ck-card {
            min-height: 116px;
            background: #FFFFFF;
            border: 1px solid #E2E7EF;
            border-top: 4px solid var(--ck-orange);
            border-radius: 10px;
            padding: 1rem 0.9rem;
            box-shadow: 0 2px 8px rgba(68, 83, 106, 0.08);
        }
        .ck-card-label {
            color: #657185;
            font-size: 0.83rem;
            font-weight: 650;
        }
        .ck-card-value {
            color: var(--ck-navy-gray);
            font-size: 1.47rem;
            font-weight: 760;
            margin-top: 0.45rem;
            overflow-wrap: anywhere;
        }
        div.stButton > button[kind="primary"],
        div.stDownloadButton > button {
            background-color: var(--ck-blue);
            color: #FFFFFF;
            border-color: var(--ck-blue);
            font-weight: 650;
        }
        div.stButton > button[kind="primary"]:hover,
        div.stDownloadButton > button:hover {
            background-color: var(--ck-orange);
            border-color: var(--ck-orange);
            color: #FFFFFF;
        }
        [data-testid="stChatMessage"] {
            border: 1px solid #E2E7EF;
            border-radius: 9px;
            padding: 0.25rem 0.35rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    # Karanlık mod CSS (session_state'e göre eklenir)
    if st.session_state.get("dark_mode", False):
        st.markdown(
            """
            <style>
            [data-testid="stAppViewContainer"],
            [data-testid="stAppViewContainer"] > div,
            [data-testid="stMain"],
            [data-testid="stMainBlockContainer"],
            section[data-testid="stAppViewContainer"] {
                background: #1E1E2E !important;
                color: #E0E0E0 !important;
            }
            [data-testid="stHeader"],
            header[data-testid="stHeader"] {
                background: #1E1E2E !important;
                color: #E0E0E0 !important;
            }
            [data-testid="stSidebar"],
            [data-testid="stSidebar"] > div,
            section[data-testid="stSidebar"] > div {
                background-color: #252535 !important;
                border-right: 1px solid #3A3A4A !important;
                color: #D0D8E8 !important;
            }
            .ck-title { color: #7BA7E8 !important; }
            .ck-subtitle { color: #A0AAB8 !important; }
            .ck-section-title { color: #D0D8E8 !important; }
            .ck-card {
                background: #252535 !important;
                border-color: #3A3A4A !important;
            }
            .ck-card-label { color: #9AA3B0 !important; }
            .ck-card-value { color: #D0D8E8 !important; }
            /* Tum metin elemanlari */
            [data-testid="stMarkdown"], [data-testid="stText"],
            [data-testid="stMarkdown"] p, [data-testid="stMarkdown"] span,
            [data-testid="stMarkdown"] label, [data-testid="stMarkdown"] div,
            p, span, label, h1, h2, h3, h4, h5, h6, li, td, th,
            .stTextInput label, .stSelectbox label,
            [data-testid="stWidgetLabel"], [data-testid="stWidgetLabel"] p {
                color: #E0E0E0 !important;
            }
            /* Input alanlari */
            [data-testid="stTextInput"] input,
            .stTextInput input,
            [data-baseweb="input"] input,
            [data-baseweb="input"] {
                background-color: #2A2A3C !important;
                color: #E0E0E0 !important;
                border-color: #3A3A4A !important;
            }
            /* Sidebar input ve label */
            [data-testid="stSidebar"] label,
            [data-testid="stSidebar"] p,
            [data-testid="stSidebar"] span,
            [data-testid="stSidebar"] div,
            [data-testid="stSidebar"] .stTextInput label,
            [data-testid="stSidebar"] input {
                color: #D0D8E8 !important;
            }
            [data-testid="stSidebar"] [data-baseweb="input"] input,
            [data-testid="stSidebar"] [data-baseweb="input"] {
                background-color: #2A2A3C !important;
                color: #D0D8E8 !important;
            }
            /* Info/Success/Warning/Error kutuları */
            [data-testid="stAlert"],
            div[data-testid="stAlert"] {
                background-color: #2A2A3C !important;
                color: #D0D8E8 !important;
            }
            [data-testid="stAlert"] p {
                color: #D0D8E8 !important;
            }
            /* Expander (Islem Loglari) */
            [data-testid="stExpander"],
            [data-testid="stExpander"] summary,
            [data-testid="stExpander"] div {
                background-color: #252535 !important;
                color: #D0D8E8 !important;
            }
            /* Code bloklari */
            [data-testid="stCode"], code, pre {
                background-color: #1A1A2A !important;
                color: #C8D0E0 !important;
            }
            /* DataFrame ve veri onizleme */
            [data-testid="stDataFrame"],
            [data-testid="stDataFrame"] div,
            [data-testid="stDataFrame"] table,
            [data-testid="stDataFrame"] th,
            [data-testid="stDataFrame"] td,
            iframe[title="streamlit_dataframe"] {
                background: #252535 !important;
                color: #D0D8E8 !important;
            }
            /* Butonlar */
            div.stButton > button,
            [data-testid="stSidebar"] div.stButton > button {
                background-color: #305496 !important;
                color: #FFFFFF !important;
                border-color: #305496 !important;
            }
            div.stButton > button:hover,
            [data-testid="stSidebar"] div.stButton > button:hover {
                background-color: #ED7D31 !important;
                border-color: #ED7D31 !important;
            }
            /* Download butonu */
            div.stDownloadButton > button {
                background-color: #305496 !important;
                color: #FFFFFF !important;
            }
            /* Selectbox, text_input container */
            [data-baseweb="select"],
            [data-baseweb="select"] div,
            [data-baseweb="popover"] li {
                background-color: #2A2A3C !important;
                color: #D0D8E8 !important;
            }
            /* Tab yapilarini da karanlik yap */
            .stTabs [data-baseweb="tab-list"],
            .stTabs [data-baseweb="tab"] {
                background-color: #252535 !important;
                color: #D0D8E8 !important;
            }
            /* Metric card iceride de karanlik */
            [data-testid="stMetric"],
            [data-testid="stMetricValue"],
            [data-testid="stMetricLabel"] {
                color: #D0D8E8 !important;
            }
            /* File uploader */
            [data-testid="stFileUploader"],
            [data-testid="stFileUploader"] div {
                background-color: #2A2A3C !important;
                color: #D0D8E8 !important;
            }
            /* Genel container kenarliklari */
            [data-testid="stVerticalBlock"],
            [data-testid="stHorizontalBlock"] {
                color: #E0E0E0 !important;
            }
            /* Sayfa header arka plan */
            [data-testid="stToolbar"] {
                background: #1E1E2E !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )

    dark_mode = bool(st.session_state.get("dark_mode", False))
    theme_tokens = (
        "--ck-blue: #78A7E8; --ck-orange: #F09A5B; --ck-green: #86C36A; "
        "--ck-text: #E6EDF6; --ck-muted: #A6B2C2; --ck-canvas: #151C26; "
        "--ck-surface: #1D2734; --ck-sidebar: #19222E; --ck-control: #222E3D; "
        "--ck-hover: #2A394B; --ck-border: #344255; --ck-shadow: 0 10px 28px rgba(0,0,0,.18);"
        if dark_mode else
        "--ck-blue: #305496; --ck-orange: #ED7D31; --ck-green: #5F963E; "
        "--ck-text: #263447; --ck-muted: #66758A; --ck-canvas: #F4F6FA; "
        "--ck-surface: #FFFFFF; --ck-sidebar: #F7F9FC; --ck-control: #FFFFFF; "
        "--ck-hover: #EFF4FA; --ck-border: #E1E7EF; --ck-shadow: 0 10px 28px rgba(38,52,71,.07);"
    )
    st.markdown(
        """
        <style>
        :root { THEME_TOKENS }
        [data-testid="stAppViewContainer"], [data-testid="stMain"] {
            background: var(--ck-canvas);
            color: var(--ck-text);
        }
        [data-testid="stMainBlockContainer"] {
            max-width: 1560px;
            padding: 2.15rem clamp(1rem, 3vw, 2.6rem) 3rem;
        }
        [data-testid="stSidebar"] {
            background: var(--ck-sidebar);
            border-right: 1px solid var(--ck-border);
            box-shadow: 5px 0 24px rgba(38, 52, 71, .035);
        }
        [data-testid="stSidebar"] > div:first-child { padding: .35rem .9rem 1.35rem; }
        [data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: .72rem; }
        [data-testid="stSidebar"] [data-testid="stImage"] {
            margin: .1rem 0 .7rem;
            padding: .85rem 1rem;
            background: var(--ck-surface);
            border: 1px solid var(--ck-border);
            border-radius: 15px;
            box-shadow: var(--ck-shadow);
        }
        [data-testid="stSidebar"] h3 {
            position: relative;
            color: var(--ck-text);
            font-size: 1.08rem;
            font-weight: 740;
            letter-spacing: -.02em;
            text-align: center;
            margin: .15rem 0 .6rem;
            padding: 0 0 .65rem;
            border-bottom: 1px solid var(--ck-border);
        }
        [data-testid="stSidebar"] h3::after {
            position: absolute;
            bottom: -1px;
            left: calc(50% - 18px);
            width: 36px;
            height: 3px;
            border-radius: 3px;
            background: var(--ck-orange);
            content: "";
        }
        [data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {
            color: var(--ck-muted);
            font-size: .78rem;
            font-weight: 590;
            line-height: 1.35;
            margin-bottom: .3rem;
        }
        [data-testid="stSidebar"] [data-testid="stTextInput"] input {
            min-height: 2.55rem;
            padding: .55rem .75rem;
            box-shadow: inset 0 1px 2px rgba(38, 52, 71, .035);
        }
        [data-testid="stSidebar"] [data-testid="stTextInput"] input:focus {
            border-color: var(--ck-blue);
            box-shadow: 0 0 0 3px color-mix(in srgb, var(--ck-blue) 16%, transparent);
        }
        [data-testid="stSidebar"] div.stButton > button {
            min-height: 2.55rem;
            border-radius: 11px;
            padding: .45rem .8rem;
            text-align: left;
        }
        [data-testid="stSidebar"] div.stButton > button[kind="primary"] {
            min-height: 2.75rem;
            text-align: center;
            box-shadow: 0 5px 13px rgba(48, 84, 150, .2);
        }
        [data-testid="stSidebar"] div.stButton > button:disabled {
            box-shadow: none;
            opacity: .58;
            transform: none;
        }
        [data-testid="stSidebar"] [data-testid="stAlert"] {
            padding: .55rem .7rem;
            border-radius: 11px;
        }
        [data-testid="stSidebar"] [data-testid="stAlert"] p {
            font-size: .79rem;
            line-height: 1.45;
        }
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
            font-size: .77rem;
            line-height: 1.45;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] summary {
            color: var(--ck-muted);
            font-size: .78rem;
            font-weight: 650;
        }
        .ck-title {
            position: relative;
            color: var(--ck-blue);
            font-size: clamp(1.65rem, 2.2vw, 2.15rem);
            font-weight: 760;
            letter-spacing: -.035em;
            line-height: 1.14;
            margin: .15rem 0 .4rem;
            padding-left: .9rem;
        }
        .ck-title::before {
            position: absolute;
            inset: .08em auto .08em 0;
            width: 4px;
            border-radius: 4px;
            background: var(--ck-orange);
            content: "";
        }
        .ck-subtitle {
            max-width: 58rem;
            color: var(--ck-muted);
            font-size: .96rem;
            line-height: 1.55;
            margin: .35rem 0 .5rem .9rem;
        }
        .ck-section-title {
            color: var(--ck-text);
            font-size: 1.32rem;
            font-weight: 720;
            letter-spacing: -.025em;
            line-height: 1.35;
        }
        .stMainBlockContainer div[data-testid="stElementContainer"]:has(.ck-section-title) {
            position: static;
            top: auto;
            margin: 1.7rem 0 .8rem;
            padding-bottom: .7rem;
            border-bottom: 1px solid var(--ck-border);
        }
        .ck-card {
            min-height: 112px;
            background: var(--ck-surface);
            border: 1px solid var(--ck-border);
            border-left: 4px solid var(--ck-orange);
            border-radius: 12px;
            padding: 1.05rem 1.1rem;
            box-shadow: var(--ck-shadow);
        }
        .ck-card-label {
            color: var(--ck-muted);
            font-size: .8rem;
            font-weight: 680;
            letter-spacing: .045em;
            text-transform: uppercase;
        }
        .ck-card-value {
            color: var(--ck-text);
            font-size: 1.62rem;
            font-weight: 760;
            letter-spacing: -.03em;
            line-height: 1.25;
            margin-top: .55rem;
        }
        [data-testid="stWidgetLabel"] p,
        [data-testid="stCaptionContainer"] { color: var(--ck-muted); }
        [data-testid="stTextInput"] input,
        [data-testid="stNumberInput"] input,
        [data-baseweb="input"] input,
        [data-baseweb="select"] > div {
            background: var(--ck-control);
            color: var(--ck-text);
            border-color: var(--ck-border);
            border-radius: 9px;
        }
        div.stButton > button, div.stDownloadButton > button {
            min-height: 2.55rem;
            border-radius: 9px;
            font-weight: 650;
            transition: background-color 140ms ease, border-color 140ms ease, transform 140ms ease;
        }
        div.stButton > button:not([kind="primary"]) {
            background: var(--ck-surface);
            color: var(--ck-text);
            border-color: var(--ck-border);
        }
        div.stButton > button:not([kind="primary"]):hover {
            background: var(--ck-hover);
            border-color: var(--ck-blue);
            color: var(--ck-blue);
            transform: translateY(-1px);
        }
        div.stButton > button[kind="primary"], div.stDownloadButton > button {
            background: #305496;
            color: #FFFFFF;
            border-color: #305496;
        }
        div.stButton > button[kind="primary"]:hover,
        div.stDownloadButton > button:hover {
            background: #ED7D31;
            border-color: #ED7D31;
            color: #FFFFFF;
        }
        [data-testid="stAlert"], [data-testid="stDataFrame"] {
            border-radius: 10px;
        }
        [data-testid="stExpander"] {
            overflow: hidden;
            background: var(--ck-surface);
            border: 1px solid var(--ck-border);
            border-radius: 10px;
        }
        @media (max-width: 760px) {
            [data-testid="stMainBlockContainer"] { padding: 1.35rem 1rem 2rem; }
            .ck-title { font-size: 1.6rem; }
            .ck-subtitle { font-size: .9rem; }
            .ck-card { min-height: 96px; padding: .85rem; }
            .ck-card-value { font-size: 1.35rem; }
        }
        </style>
        """.replace("THEME_TOKENS", theme_tokens),
        unsafe_allow_html=True,
    )


def select_folder_via_dialog(dialog_title: str) -> str | None:
    """Yerel Windows klasör seçicisini aç; başarısızsa metin girişi kullanılır."""
    dialog_root = None
    try:
        import tkinter as tk
        from tkinter import filedialog

        dialog_root = tk.Tk()
        dialog_root.withdraw()
        dialog_root.wm_attributes("-topmost", 1)
        selected = filedialog.askdirectory(
            master=dialog_root,
            title=dialog_title,
        )
        return selected or None
    except Exception:
        return None
    finally:
        if dialog_root is not None:
            try:
                dialog_root.destroy()
            except Exception:
                pass


def select_excel_file_via_dialog(
    dialog_title: str,
    initial_path: str | None = None,
) -> str | None:
    """Open a local Excel picker, starting at the currently selected workbook."""
    dialog_root = None
    try:
        import tkinter as tk
        from tkinter import filedialog

        dialog_root = tk.Tk()
        dialog_root.withdraw()
        dialog_root.wm_attributes("-topmost", 1)
        picker_options: dict[str, str] = {}
        if initial_path:
            current_path = Path(initial_path)
            if current_path.parent.is_dir():
                picker_options["initialdir"] = str(current_path.parent)
            if current_path.is_file():
                picker_options["initialfile"] = current_path.name
        selected = filedialog.askopenfilename(
            master=dialog_root,
            title=dialog_title,
            filetypes=[
                ("Excel dosyaları", "*.xlsx *.xlsm *.xls"),
                ("Tüm dosyalar", "*.*"),
            ],
            **picker_options,
        )
        return selected or None
    except Exception:
        return None
    finally:
        if dialog_root is not None:
            try:
                dialog_root.destroy()
            except Exception:
                pass


def safe_float(value: Any) -> float:
    """Sayısal değeri çekirdeğin Türkçe sayı temizleyicisiyle güvenle dönüştür."""
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        number = float(clean_turkish_number(value))
    return number if math.isfinite(number) else 0.0


def format_turkish_number(value: Any, decimals: int = 2, suffix: str = "") -> str:
    """1,000,000.34 değerini 1.000.000,34 biçimine çevir."""
    number = safe_float(value)
    formatted = f"{number:,.{decimals}f}"
    formatted = formatted.replace(",", "\u0000").replace(".", ",").replace("\u0000", ".")
    return f"{formatted}{suffix}"


def format_turkish_integer(value: Any) -> str:
    return format_turkish_number(value, decimals=0)


def format_turkish_currency(value: Any) -> str:
    return format_turkish_number(value, decimals=2, suffix=" TL")


def mask_etso(etso_code: Any) -> str:
    """ETSO kodunu yalnız UI için 12***678 biçiminde maskele."""
    if etso_code is None or str(etso_code).strip().lower() in {"", "none", "nan"}:
        return "***"
    value = str(etso_code).strip()
    if value.endswith(".0") and value[:-2].isdigit():
        value = value[:-2]
    if len(value) <= 4:
        return f"{value[:1]}***"
    if len(value) <= 6:
        return f"{value[:2]}***{value[-1:]}"
    return f"{value[:2]}***{value[-3:]}"


def normalize_column_key(value: Any) -> str:
    text = str(value).strip().lower()
    replacements = {
        "ı": "i",
        "ğ": "g",
        "ü": "u",
        "ö": "o",
        "ç": "c",
        "ş": "s",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    return re.sub(r"[^a-z0-9]", "", text)


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return str(value).strip().lower() in {"", "none", "nan"}


def is_numeric_text(value: Any) -> bool:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return True
    if is_blank(value):
        return False
    return bool(re.fullmatch(r"[+\-()\d\s.,]+", str(value).strip()))


def format_dataframe_for_ui(source_df: pd.DataFrame) -> pd.DataFrame:
    """Ham DataFrame'i değiştirmeden ETSO maskeleme ve Türkçe sayı görünümü uygula."""
    display_df = source_df.copy(deep=True)
    money_keys = {normalize_column_key(field) for field in MONEY_FIELDS}
    quantity_keys = {normalize_column_key(field) for field in QUANTITY_FIELDS}

    for column in display_df.columns:
        normalized_column = normalize_column_key(column)
        if normalized_column == "etsokodu":
            display_df[column] = display_df[column].map(mask_etso)
            continue

        if normalized_column in money_keys:
            display_df[column] = display_df[column].map(
                lambda value: "" if is_blank(value) else (
                    str(value) if str(value).strip().upper() == "X" else format_turkish_currency(value)
                )
            )
            continue

        if normalized_column in quantity_keys:
            display_df[column] = display_df[column].map(
                lambda value: "" if is_blank(value) else (
                    format_turkish_number(value, 2) if is_numeric_text(value) else str(value)
                )
            )
            continue

        if display_df[column].dtype == object:
            display_df[column] = display_df[column].map(
                lambda value: "" if is_blank(value) else str(value)
            )

    return display_df


def percentile_90(values: Iterable[float]) -> float:
    numeric_values = list(values)
    if not numeric_values:
        return 0.0
    return float(pd.Series(numeric_values, dtype="float64").quantile(0.90))


def audit_extracted_dataset(dataset: list[dict[str, Any]]) -> dict[str, Any]:
    """10 milyon TL mutlak alt sınırla bölgesel anomali ve veri kalitesi denetimi."""
    warnings_list: list[dict[str, str]] = []
    absurd_records: list[dict[str, Any]] = []
    high_value_records: list[dict[str, Any]] = []
    negative_cost_records: list[dict[str, Any]] = []
    missing_customer_count = 0
    missing_etso_count = 0
    region_counts: dict[str, int] = {}
    region_values: dict[str, list[float]] = {}

    # Kalite sayaçları anomali eşiğinden bağımsız hesaplanır.
    for row in dataset:
        region = str(row.get("Dağıtım Bölgesi") or "Bilinmeyen Bölge")
        bedel = safe_float(row.get("Dağıtım Bedeli(TL)", 0))
        etso = row.get("Etso Kodu")
        customer = row.get("Müşteri")

        region_counts[region] = region_counts.get(region, 0) + 1
        if 0 < bedel < ABSURD_CORRUPTION_THRESHOLD_TL:
            region_values.setdefault(region, []).append(bedel)
        if bedel < 0:
            negative_cost_records.append({"region": region, "val": bedel})
        if is_blank(customer):
            missing_customer_count += 1
        if is_blank(etso):
            missing_etso_count += 1

    region_baselines: dict[str, dict[str, float]] = {}
    for region, values in region_values.items():
        median_value = float(pd.Series(values, dtype="float64").median()) if values else 0.0
        p90_value = percentile_90(values)
        spike_threshold = max(
            median_value * 7.0,
            p90_value * 3.5,
            ANOMALY_MINIMUM_TL,
        )
        region_baselines[region] = {
            "median": median_value,
            "p90": p90_value,
            "spike_threshold": spike_threshold,
        }

    for row_index, row in enumerate(dataset, start=2):
        bedel = safe_float(row.get("Dağıtım Bedeli(TL)", 0))

        # Kesin kural: 10 milyon TL altındaki hiçbir fatura uyarı paneline girmez.
        if bedel < 10000000.0:
            continue

        region = str(row.get("Dağıtım Bölgesi") or "Bilinmeyen Bölge")
        masked_etso = mask_etso(row.get("Etso Kodu"))
        baseline = region_baselines.get(
            region,
            {"median": 100_000.0, "spike_threshold": ANOMALY_MINIMUM_TL},
        )
        region_median = baseline["median"]
        spike_threshold = baseline["spike_threshold"]

        if bedel > ABSURD_CORRUPTION_THRESHOLD_TL:
            absurd_records.append({"row": row_index, "region": region, "etso": masked_etso, "val": bedel})
            warnings_list.append(
                {
                    "type": "DANGER",
                    "category": "Absürt Tutar",
                    "message": (
                        f"⚠️ KRİTİK: {region} bölgesinde ETSO {masked_etso} için "
                        f"{format_turkish_currency(bedel)} tutarında bozuk veri tespit edildi."
                    ),
                }
            )
        elif bedel >= spike_threshold:
            ratio = bedel / max(1.0, region_median)
            high_value_records.append(
                {"row": row_index, "region": region, "etso": masked_etso, "val": bedel, "ratio": ratio}
            )
            warnings_list.append(
                {
                    "type": "WARNING",
                    "category": "Yüksek Tutar",
                    "message": (
                        f"🔔 DİKKAT: {region} bölgesinde ETSO {masked_etso} için "
                        f"{format_turkish_currency(bedel)} tutarında yüksek fatura tespit edildi; "
                        f"bölge medyanının {format_turkish_number(ratio, 1)} katı."
                    ),
                }
            )

    if negative_cost_records:
        warnings_list.append(
            {
                "type": "INFO",
                "category": "Negatif Tutar",
                "message": (
                    f"ℹ️ Toplam {format_turkish_integer(len(negative_cost_records))} negatif bedelli kayıt "
                    "veri kalitesi sayacına alındı."
                ),
            }
        )
    if missing_customer_count:
        warnings_list.append(
            {
                "type": "INFO",
                "category": "Eksik Müşteri",
                "message": (
                    f"ℹ️ Toplam {format_turkish_integer(missing_customer_count)} kayıtta "
                    "müşteri unvanı bulunmuyor."
                ),
            }
        )
    if missing_etso_count:
        warnings_list.append(
            {
                "type": "INFO",
                "category": "Eksik ETSO",
                "message": (
                    f"ℹ️ Toplam {format_turkish_integer(missing_etso_count)} kayıtta ETSO kodu bulunmuyor."
                ),
            }
        )

    health_score = max(0, 100 - (len(absurd_records) * 20) - (missing_etso_count * 5))
    return {
        "total_records": len(dataset),
        "region_counts": region_counts,
        "absurd_count": len(absurd_records),
        "high_value_count": len(high_value_records),
        "negative_cost_count": len(negative_cost_records),
        "missing_customer_count": missing_customer_count,
        "missing_etso_count": missing_etso_count,
        "health_score": health_score,
        "region_baselines": region_baselines,
        "warnings": warnings_list,
    }


def initialize_session_state() -> None:
    defaults: dict[str, Any] = {
        "selected_root_folder": "",
        "selected_invoice_file": "",
        "dark_mode": False,
        "processing_done": False,
        "extracted_df": None,
        "stats": {},
        "audit_results": {},
        "output_bytes": None,
        "invoice_sources": [],
        "company_comparison_reports": {},
        "subscription_group_mismatches": [],
        "comparison_report_warning": None,
        "company_report_selection": None,
        "company_report_generated": False,
        "processing_logs": [],
        "analysis_signature": None,

    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    if "selected_system_usage_file" not in st.session_state:
        st.session_state["selected_system_usage_file"] = (
            "" if uses_browser_file_inputs() else load_system_usage_file_selection()
        )
    if "selected_reconciliation_consumption_file" not in st.session_state:
        st.session_state["selected_reconciliation_consumption_file"] = ""
    if "selected_eno_file" not in st.session_state:
        st.session_state["selected_eno_file"] = (
            "" if uses_browser_file_inputs() else load_eno_file_selection()
        )


def load_system_usage_file_selection() -> str:
    try:
        settings = json.loads(SYSTEM_USAGE_FILE_PREFERENCE.read_text(encoding="utf-8"))
        selected_file = settings.get("path") if isinstance(settings, dict) else None
        if isinstance(selected_file, str) and selected_file.strip():
            return selected_file.strip()
    except (OSError, json.JSONDecodeError):
        pass
    return str(SYSTEM_USAGE_FILE_DEFAULT)


def save_system_usage_file_selection(file_path: str) -> str | None:
    settings: dict[str, Any] = {}
    try:
        existing = json.loads(SYSTEM_USAGE_FILE_PREFERENCE.read_text(encoding="utf-8"))
        if isinstance(existing, dict):
            settings = existing
    except (OSError, json.JSONDecodeError):
        pass

    settings["path"] = file_path
    temporary_path = SYSTEM_USAGE_FILE_PREFERENCE.with_suffix(".tmp")
    try:
        SYSTEM_USAGE_FILE_PREFERENCE.parent.mkdir(parents=True, exist_ok=True)
        temporary_path.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        temporary_path.replace(SYSTEM_USAGE_FILE_PREFERENCE)
    except OSError as error:
        return str(error)
    return None


def load_eno_file_selection() -> str:
    try:
        settings = json.loads(ENO_FILE_PREFERENCE.read_text(encoding="utf-8"))
        selected_file = settings.get("path") if isinstance(settings, dict) else None
        if isinstance(selected_file, str) and selected_file.strip():
            return selected_file.strip()
    except (OSError, json.JSONDecodeError):
        pass
    return ""


def save_eno_file_selection(file_path: str) -> str | None:
    settings: dict[str, Any] = {}
    try:
        existing = json.loads(ENO_FILE_PREFERENCE.read_text(encoding="utf-8"))
        if isinstance(existing, dict):
            settings = existing
    except (OSError, json.JSONDecodeError):
        pass

    settings["path"] = file_path
    temporary_path = ENO_FILE_PREFERENCE.with_suffix(".tmp")
    try:
        temporary_path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        temporary_path.replace(ENO_FILE_PREFERENCE)
    except OSError as error:
        return str(error)
    return None


def clear_analysis_state() -> None:
    st.session_state["processing_done"] = False
    st.session_state["extracted_df"] = None
    st.session_state["stats"] = {}
    st.session_state["audit_results"] = {}
    st.session_state["output_bytes"] = None
    st.session_state["invoice_sources"] = []
    st.session_state["company_comparison_reports"] = {}
    st.session_state["subscription_group_mismatches"] = []
    st.session_state["comparison_report_warning"] = None
    st.session_state["company_report_selection"] = None
    st.session_state["company_report_generated"] = False
    st.session_state["processing_logs"] = []
    st.session_state["analysis_signature"] = None


def find_mapping_files(root_folder: str) -> list[str]:
    if not root_folder or not os.path.isdir(root_folder):
        return []
    mapping_files = [
        name
        for name in os.listdir(root_folder)
        if name.lower().endswith(".xlsx")
        and "skf" in name.casefold()
        and not name.startswith("~$")
        and os.path.isfile(os.path.join(root_folder, name))
    ]

    def mapping_version_sort_key(name: str) -> tuple[int, str]:
        version_match = re.search(r"güncel\s*(\d+)", name.casefold())
        version = int(version_match.group(1)) if version_match else -1
        return (-version, name.casefold())

    return sorted(mapping_files, key=mapping_version_sort_key)


def count_candidate_region_folders(root_folder: str, excluded_folder: str | None = None) -> int:
    if not root_folder or not os.path.isdir(root_folder):
        return 0
    ignored = {"core", "config", "analytics", "__pycache__", ".git", ".agents"}
    excluded_path = (
        os.path.normcase(os.path.abspath(excluded_folder))
        if excluded_folder and os.path.isdir(excluded_folder)
        else None
    )
    return sum(
        1
        for name in os.listdir(root_folder)
        if os.path.isdir(os.path.join(root_folder, name))
        and (
            excluded_path is None
            or os.path.normcase(os.path.abspath(os.path.join(root_folder, name))) != excluded_path
        )
        and name not in ignored
        and not name.startswith(".")
    )


def normalize_extracted_regions(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized_records = []
    for record in records:
        normalized_record = record.copy()
        normalized_record["Dağıtım Bölgesi"] = normalize_region_name(
            record.get("Dağıtım Bölgesi"),
            to_format="reference",
        )
        normalized_records.append(normalized_record)
    return normalized_records


def load_system_usage_distribution_rates(
    mapping_path: str | None,
    system_usage_file: str | None,
) -> tuple[dict[str, float], str | None]:
    """Resolve tariff group rates from the mapping sheet and selected rate workbook."""
    if not mapping_path or not os.path.isfile(mapping_path):
        return {}, "Tarifeler eşleştirme dosyası bulunamadı."
    if not system_usage_file or not os.path.isfile(system_usage_file):
        return {}, "Sistem Kullanım Bedelleri dosyası bulunamadı."

    from openpyxl import load_workbook
    from openpyxl.utils.cell import column_index_from_string

    mapping_workbook = None
    usage_workbook = None
    usage_xls_workbook = None
    try:
        mapping_workbook = load_workbook(mapping_path, read_only=True, data_only=True)
        if "Tarifeler" not in mapping_workbook.sheetnames:
            return {}, "SKF Başlıkları dosyasında Tarifeler sayfası bulunamadı."

        tariff_sheet = mapping_workbook["Tarifeler"]
        header = next(tariff_sheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
        header_columns = {
            str(value).strip().casefold(): index
            for index, value in enumerate(header)
            if value is not None
        }
        required_headers = ("tarife grubu", "sütun", "satır")
        if any(name not in header_columns for name in required_headers):
            return {}, "Tarifeler sayfasında Tarife Grubu, Sütun ve Satır başlıkları bulunamadı."

        usage_extension = Path(system_usage_file).suffix.casefold()
        if usage_extension == ".xls":
            import xlrd

            usage_xls_workbook = xlrd.open_workbook(system_usage_file, on_demand=True)
            usage_sheet = usage_xls_workbook.sheet_by_index(0)
        else:
            usage_workbook = load_workbook(
                system_usage_file,
                read_only=True,
                data_only=True,
            )
            usage_sheet = usage_workbook.worksheets[0]

        rates: dict[str, float] = {}
        for row in tariff_sheet.iter_rows(min_row=2, values_only=True):
            group_value = row[header_columns["tarife grubu"]]
            column_value = row[header_columns["sütun"]]
            row_value = row[header_columns["satır"]]
            if group_value is None or column_value is None or row_value is None:
                continue

            group_key = str(group_value).strip().casefold()
            column_name = str(column_value).strip().upper()
            try:
                rate_row = int(float(row_value))
                rate_column = column_index_from_string(column_name)
            except (TypeError, ValueError):
                continue

            if usage_extension == ".xls":
                if rate_row < 1 or rate_row > usage_sheet.nrows or rate_column < 1 or rate_column > usage_sheet.ncols:
                    continue
                rate_value = usage_sheet.cell_value(rate_row - 1, rate_column - 1)
            else:
                rate_value = usage_sheet.cell(rate_row, rate_column).value
            # Blank tariff cells represent a zero charge in the tariff table
            # (for example, İSK Sanayi at I6).
            rates[group_key] = safe_float(rate_value)
        return rates, None
    except Exception as error:
        return {}, f"Tarife birim fiyatları okunamadı: {error}"
    finally:
        if mapping_workbook is not None:
            mapping_workbook.close()
        if usage_workbook is not None:
            usage_workbook.close()
        if usage_xls_workbook is not None:
            usage_xls_workbook.release_resources()


def _normalize_tariff_lookup_key(value: Any) -> str:
    import unicodedata

    normalized = unicodedata.normalize("NFKD", str(value or "")).casefold()
    normalized = "".join(char for char in normalized if not unicodedata.combining(char))
    normalized = normalized.replace("ı", "i")
    return "".join(char for char in normalized if char.isalnum())


def normalize_invoice_subscription_group(value: Any) -> Any:
    """Collapse connection-specific invoice group labels to their tariff group."""
    normalized = _normalize_tariff_lookup_key(value)
    if normalized in {"ticarethaneag", "ticarethaneog", "ticarethanebtv1"}:
        return "Ticarethane"
    if normalized in {"sanayiag", "sanayiog"}:
        return "Sanayi"
    return value


def load_system_usage_power_rates(
    mapping_path: str | None,
    system_usage_file: str | None,
) -> tuple[dict[str, dict[str, Any]], str | None]:
    """Resolve tariff group, term, connection and power rates for the invoice comparison sheet."""
    from openpyxl import load_workbook

    if not mapping_path or not os.path.isfile(mapping_path):
        return {}, "SKF Başlıkları dosyası bulunamadı."

    mapping_workbook = None
    try:
        mapping_workbook = load_workbook(mapping_path, read_only=True, data_only=True)
        tariff_sheet = next(
            (ws for ws in mapping_workbook.worksheets if _normalize_tariff_lookup_key(ws.title) == "tarifeler"),
            None,
        )
        layout_sheet = next(
            (
                ws for ws in mapping_workbook.worksheets
                if _normalize_tariff_lookup_key(ws.title) == "bizimkestigimiz2"
            ),
            None,
        )
        if tariff_sheet is None or layout_sheet is None:
            return {}, "SKF Başlıkları dosyasında tarife eşlemesi bulunamadı."

        tariff_classes = {}
        for row in tariff_sheet.iter_rows(min_row=2, values_only=True):
            if len(row) < 3 or row[0] is None:
                continue
            tariff_classes[_normalize_tariff_lookup_key(row[0])] = {
                "rate_category": row[1],
                "subscriber_group": row[2],
            }

        header_row = None
        for row_index, row in enumerate(layout_sheet.iter_rows(values_only=True), start=1):
            if len(row) >= 4 and (
                _normalize_tariff_lookup_key(row[1]) == "tarifegrubu"
                and _normalize_tariff_lookup_key(row[2]) == "terimtipi"
                and _normalize_tariff_lookup_key(row[3]) == "baglantitipi"
            ):
                header_row = row_index
                break
        if header_row is None:
            return {}, "Bizim Kestiğimiz 2 tarife eşlemesi bulunamadı."

        tariff_layout: dict[str, dict[str, Any]] = {}
        found_group = False
        blank_after_groups = False
        for row in layout_sheet.iter_rows(min_row=header_row + 1, min_col=2, max_col=4, values_only=True):
            group_value = row[0] if row else None
            if group_value is None or not str(group_value).strip():
                if found_group:
                    blank_after_groups = True
                continue
            if blank_after_groups:
                break
            group_key = _normalize_tariff_lookup_key(group_value)
            if group_key == "tarifegrubu":
                break
            found_group = True
            tariff_class = tariff_classes.get(group_key, {})
            tariff_layout[group_key] = {
                "tariff_group": str(group_value).strip(),
                "term_type": row[1] if len(row) > 1 else "",
                "connection_type": row[2] if len(row) > 2 else "",
                "rate_category": tariff_class.get("rate_category"),
                "subscriber_group": tariff_class.get("subscriber_group"),
                "power_rate": 0.0,
            }
    except Exception as error:
        return {}, f"Tarife eşlemesi okunamadı: {error}"
    finally:
        if mapping_workbook is not None:
            mapping_workbook.close()

    if not tariff_layout:
        return {}, "SKF Başlıkları dosyasında kullanılabilir tarife grubu bulunamadı."
    if not system_usage_file or not os.path.isfile(system_usage_file):
        return tariff_layout, "Güç tarifesi dosyası bulunamadı."

    rate_rows: list[tuple[str, str, float]] = []
    usage_workbook = None
    usage_xls_workbook = None
    try:
        usage_extension = Path(system_usage_file).suffix.casefold()
        if usage_extension == ".xls":
            import xlrd

            usage_xls_workbook = xlrd.open_workbook(system_usage_file, on_demand=True)
            usage_sheet = usage_xls_workbook.sheet_by_index(0)
            current_category = ""
            for row_index in range(5, min(usage_sheet.nrows, 29)):
                category_value = usage_sheet.cell_value(row_index, 2)
                subscriber_value = usage_sheet.cell_value(row_index, 3)
                if category_value is not None and str(category_value).strip():
                    current_category = str(category_value).strip()
                if subscriber_value is None or not str(subscriber_value).strip():
                    continue
                rate_value = usage_sheet.cell_value(row_index, 9) if usage_sheet.ncols > 9 else 0
                rate_rows.append((
                    _normalize_tariff_lookup_key(current_category),
                    _normalize_tariff_lookup_key(subscriber_value),
                    safe_float(rate_value),
                ))
        else:
            usage_workbook = load_workbook(system_usage_file, read_only=True, data_only=True)
            usage_sheet = usage_workbook.worksheets[0]
            current_category = ""
            for row_index in range(6, min(usage_sheet.max_row, 29) + 1):
                category_value = usage_sheet.cell(row_index, 3).value
                subscriber_value = usage_sheet.cell(row_index, 4).value
                if category_value is not None and str(category_value).strip():
                    current_category = str(category_value).strip()
                if subscriber_value is None or not str(subscriber_value).strip():
                    continue
                rate_rows.append((
                    _normalize_tariff_lookup_key(current_category),
                    _normalize_tariff_lookup_key(subscriber_value),
                    safe_float(usage_sheet.cell(row_index, 10).value),
                ))

        for tariff in tariff_layout.values():
            category_key = _normalize_tariff_lookup_key(tariff.get("rate_category"))
            subscriber_key = _normalize_tariff_lookup_key(tariff.get("subscriber_group"))
            if not category_key:
                tariff["power_rate"] = 0.0
                continue
            candidates = [
                rate for category, subscriber, rate in rate_rows
                if category == category_key
                and subscriber_key
                and (
                    subscriber == subscriber_key
                    or subscriber.startswith(subscriber_key)
                    or subscriber_key.startswith(subscriber)
                )
            ]
            tariff["power_rate"] = candidates[0] if candidates else 0.0
        return tariff_layout, None
    except Exception as error:
        return tariff_layout, f"Güç tarife birim fiyatları okunamadı: {error}"
    finally:
        if usage_workbook is not None:
            usage_workbook.close()
        if usage_xls_workbook is not None:
            usage_xls_workbook.release_resources()


def add_expected_distribution_charge_column(
    worksheet: Any,
    tariff_rates: dict[str, float],
) -> dict[str, Any]:
    """Add the calculated distribution charge beside its source column."""
    from copy import copy
    from openpyxl.styles import PatternFill
    from openpyxl.utils import get_column_letter

    def normalize_header(value: Any) -> str:
        return str(value).strip().casefold() if value is not None else ""

    headers = {
        normalize_header(cell.value): cell.column
        for cell in worksheet[1]
        if cell.value is not None
    }
    distribution_column = headers.get("dağıtım bedeli (tl)")
    if distribution_column is None:
        return {
            "calculated_distribution_amount_count": 0,
            "missing_distribution_rate_count": 0,
            "distribution_amount_warning": "Dağıtım Bedeli (TL) başlığı bulunamadı.",
        }

    expected_header = "Olması Gereken Dağıtım Bedeli"
    existing_expected_column = headers.get(normalize_header(expected_header))
    if existing_expected_column is not None and existing_expected_column != distribution_column + 1:
        return {
            "calculated_distribution_amount_count": 0,
            "missing_distribution_rate_count": 0,
            "distribution_amount_warning": (
                "Olması Gereken Dağıtım Bedeli sütunu Dağıtım Bedeli (TL) yanında değil."
            ),
        }

    if existing_expected_column is None:
        expected_column = distribution_column + 1
        worksheet.insert_cols(expected_column)
        expected_cell = worksheet.cell(1, expected_column)
        source_header_cell = worksheet.cell(1, distribution_column)
        if source_header_cell.has_style:
            expected_cell._style = copy(source_header_cell._style)
    else:
        expected_column = existing_expected_column
        expected_cell = worksheet.cell(1, expected_column)

    expected_cell.value = expected_header
    expected_cell.fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    expected_font = copy(expected_cell.font)
    expected_font.bold = True
    expected_font.color = "000000"
    expected_cell.font = expected_font
    expected_cell.alignment = copy(worksheet.cell(1, distribution_column).alignment)
    worksheet.column_dimensions[expected_cell.column_letter].width = max(
        worksheet.column_dimensions[expected_cell.column_letter].width or 0,
        len(expected_header) + 2,
    )

    def shifted_column(column: int | None) -> int | None:
        if column is None:
            return None
        return column + 1 if existing_expected_column is None and column >= expected_column else column

    active_column = shifted_column(headers.get("aktif enerji tüketim (kwh)"))
    tariff_column = shifted_column(headers.get("tarife grubu"))
    etso_column = next(
        (
            shifted_column(headers[name])
            for name in ("etso", "etso kodu")
            if name in headers
        ),
        None,
    )
    if active_column is None or tariff_column is None or etso_column is None:
        return {
            "calculated_distribution_amount_count": 0,
            "missing_distribution_rate_count": 0,
            "distribution_amount_warning": (
                "ETSO, Tarife Grubu veya Aktif Enerji Tüketim (kWh) başlığı bulunamadı."
            ),
        }

    calculated_count = 0
    missing_rate_count = 0
    expected_last_row = worksheet.max_row
    active_letter = get_column_letter(active_column)
    style_cache: dict[int, Any] = {}
    for row_index in range(2, expected_last_row + 1):
        etso_value = worksheet.cell(row_index, etso_column).value
        tariff_value = worksheet.cell(row_index, tariff_column).value
        energy_value = worksheet.cell(row_index, active_column).value
        if etso_value is None or not str(etso_value).strip() or energy_value is None or energy_value == "":
            continue

        tariff_key = str(tariff_value).strip().casefold() if tariff_value is not None else ""
        rate = tariff_rates.get(tariff_key)
        output_cell = worksheet.cell(row_index, expected_column)
        source_amount_cell = worksheet.cell(row_index, distribution_column)
        if source_amount_cell.has_style:
            style_id = source_amount_cell.style_id
            cached_style = style_cache.get(style_id)
            if cached_style is None:
                cached_style = copy(source_amount_cell._style)
                style_cache[style_id] = cached_style
            output_cell._style = cached_style
        output_cell.number_format = source_amount_cell.number_format or "#,##0.00"
        if rate is None:
            output_cell.value = None
            missing_rate_count += 1
            continue

        output_cell.value = f"=IFERROR({active_letter}{row_index}*{rate!r}/100,0)"
        calculated_count += 1

    return {
        "calculated_distribution_amount_count": calculated_count,
        "missing_distribution_rate_count": missing_rate_count,
        "distribution_amount_warning": None,
    }


def detect_our_invoice_format(header_values: Iterable[Any]) -> str | None:
    """Identify supported invoice exports from their headers, independent of file name."""
    headers = {
        _normalize_tariff_lookup_key(value)
        for value in header_values
        if value is not None
    }
    region_summary_headers = {
        _normalize_tariff_lookup_key("Hizmet Noktas\u0131 No"),
        _normalize_tariff_lookup_key("Da\u011f\u0131t\u0131m Hizmet Noktas\u0131 No"),
        _normalize_tariff_lookup_key("Fatura Abone Grubu"),
        _normalize_tariff_lookup_key("Toplam Da\u011f\u0131t\u0131m Bedeli"),
        _normalize_tariff_lookup_key("Fatura Tutar\u0131"),
        _normalize_tariff_lookup_key("ETSO Kodu"),
    }
    if len(headers.intersection(region_summary_headers)) >= 5:
        return "region_summary"

    invoice_management_header_groups = (
        {
            _normalize_tariff_lookup_key("ETSO"),
            _normalize_tariff_lookup_key("ETSO Kodu"),
        },
        {_normalize_tariff_lookup_key("Da\u011f\u0131t\u0131m B\u00f6lgesi")},
        {_normalize_tariff_lookup_key("M\u00fc\u015fteri")},
        {_normalize_tariff_lookup_key("Tarife Grubu")},
        {_normalize_tariff_lookup_key("Abone Grubu")},
        {_normalize_tariff_lookup_key("Aktif Enerji T\u00fcketim (kWh)")},
        {_normalize_tariff_lookup_key("Da\u011f\u0131t\u0131m Bedeli (TL)")},
    )
    matching_management_headers = sum(
        bool(headers.intersection(header_group))
        for header_group in invoice_management_header_groups
    )
    if matching_management_headers >= 5:
        return "invoice_management"
    return None


def detect_our_invoice_worksheet_format(worksheet: Any) -> str | None:
    header_row = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), ())
    return detect_our_invoice_format(header_row)


def load_region_our_invoice_mapping(
    mapping_path: str | None,
) -> tuple[dict[str, dict[str, Any]], dict[str, dict[str, Any]], str | None]:
    if not mapping_path or not os.path.isfile(mapping_path):
        return {}, {}, "SKF ba\u015fl\u0131k e\u015fle\u015ftirme dosyas\u0131 bulunamad\u0131."

    from openpyxl import load_workbook

    workbook = None
    try:
        workbook = load_workbook(mapping_path, read_only=True, data_only=True)
        source_mapping_sheet = next(
            (
                sheet for sheet in workbook.worksheets
                if "bolge" in _normalize_tariff_lookup_key(sheet.title)
                and "bizim" in _normalize_tariff_lookup_key(sheet.title)
                and "kesti" in _normalize_tariff_lookup_key(sheet.title)
            ),
            None,
        )
        if source_mapping_sheet is None:
            return {}, {}, (
                "SKF Ba\u015fl\u0131klar\u0131 dosyas\u0131nda B\u00f6lge Bizim Kesti\u011fimiz e\u015fle\u015ftirme sayfas\u0131 bulunamad\u0131."
            )

        rows = list(source_mapping_sheet.iter_rows(values_only=True))
        first_row = rows[0] if rows else ()
        second_row = rows[1] if len(rows) > 1 else ()
        configured_fields: dict[str, dict[str, Any]] = {}
        for column_index in range(0, min(len(first_row), len(second_row)), 2):
            field_name = first_row[column_index]
            source_header = second_row[column_index]
            source_column_letter = (
                second_row[column_index + 1]
                if column_index + 1 < len(second_row)
                else None
            )
            if field_name is None or source_header is None:
                continue
            configured_fields[_normalize_tariff_lookup_key(field_name)] = {
                "field_name": str(field_name).strip(),
                "source_header": str(source_header).strip(),
                "source_column_letter": (
                    str(source_column_letter).strip()
                    if source_column_letter is not None
                    else ""
                ),
            }

        required_fields = (
            "ETSO",
            "M\u00fc\u015fteri",
            "Tarife Grubu",
            "AG-OG",
            "Terim",
            "Kurulu G\u00fc\u00e7",
            "S\u00f6zle\u015fme G\u00fcc\u00fc",
            "Aktif Enerji (kwh)",
            "Aktif Enerji TL",
            "Da\u011f\u0131t\u0131m Bedeli",
            "G\u00fc\u00e7 Bedeli",
            "G\u00fc\u00e7 a\u015f\u0131m Bedeli",
            "Reaktif Bedel",
            "BTV",
            "Fiyat Fark\u0131",
            "Kdv",
            "KDV Matrah\u0131",
            "TOPLAM",
        )
        missing_fields = [
            label for label in required_fields
            if _normalize_tariff_lookup_key(label) not in configured_fields
        ]
        if missing_fields:
            return {}, {}, (
                "B\u00f6lge Bizim Kesti\u011fimiz e\u015fle\u015ftirme sayfas\u0131nda eksik alanlar: "
                + ", ".join(missing_fields)
            )

        corrections: dict[str, dict[str, Any]] = {}
        correction_header = (
            _normalize_tariff_lookup_key("Ba\u015fl\u0131k"),
            {
                _normalize_tariff_lookup_key("Y\u00fcklenen"),
                _normalize_tariff_lookup_key("Y\u00fckelenen"),
            },
            _normalize_tariff_lookup_key("D\u00fczeltme"),
        )
        for row_index, row in enumerate(rows):
            normalized_correction_header = tuple(
                _normalize_tariff_lookup_key(row[index]) if index < len(row) else ""
                for index in range(3)
            )
            if (
                normalized_correction_header[0] != correction_header[0]
                or normalized_correction_header[1] not in correction_header[1]
                or normalized_correction_header[2] != correction_header[2]
            ):
                continue
            for correction_row in rows[row_index + 1:]:
                values = list(correction_row[:3])
                values.extend([None] * (3 - len(values)))
                field_name, uploaded_value, corrected_value = values
                if field_name is None and uploaded_value is None and corrected_value is None:
                    break
                if field_name is None or uploaded_value is None or corrected_value is None:
                    continue
                corrections.setdefault(
                    _normalize_tariff_lookup_key(field_name),
                    {},
                )[_normalize_tariff_lookup_key(uploaded_value)] = corrected_value
        return configured_fields, corrections, None
    except Exception as error:
        return {}, {}, f"SKF ba\u015fl\u0131k e\u015fle\u015ftirme dosyas\u0131 okunamad\u0131: {error}"
    finally:
        if workbook is not None:
            workbook.close()


def add_region_invoice_summary_worksheet(
    report_workbook: Any,
    source_worksheet: Any,
    comparison_sheet_base: str,
    mapping_path: str | None,
) -> dict[str, Any]:
    """Aggregate the regional invoice summary by ETSO and create the own-invoice view."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils.cell import column_index_from_string
    from openpyxl.utils import get_column_letter

    configured_fields, corrections, mapping_error = load_region_our_invoice_mapping(mapping_path)
    if mapping_error:
        return {"comparison_sheet_warning": mapping_error}

    source_header_columns: dict[str, list[int]] = {}
    for cell in source_worksheet[1]:
        if cell.value is not None:
            source_header_columns.setdefault(
                _normalize_tariff_lookup_key(cell.value),
                [],
            ).append(cell.column)

    source_columns: dict[str, int] = {}
    missing_headers: list[str] = []
    for field_key, definition in configured_fields.items():
        source_header_key = _normalize_tariff_lookup_key(definition["source_header"])
        matches = source_header_columns.get(source_header_key, [])
        preferred_column = None
        try:
            preferred_column = column_index_from_string(
                definition["source_column_letter"].upper()
            )
        except (TypeError, ValueError):
            pass
        if preferred_column in matches:
            source_columns[field_key] = preferred_column
        elif matches:
            source_columns[field_key] = matches[0]
        else:
            missing_headers.append(definition["source_header"])
    if missing_headers:
        return {
            "comparison_sheet_warning": (
                "B\u00f6lge Bizim Kesti\u011fimiz kayna\u011f\u0131nda e\u015fle\u015ftirilen ba\u015fl\u0131klar bulunamad\u0131: "
                + ", ".join(missing_headers)
            )
        }

    def optional_source_column(*aliases: str) -> int | None:
        for alias in aliases:
            matching_columns = source_header_columns.get(_normalize_tariff_lookup_key(alias), [])
            if matching_columns:
                return matching_columns[0]
        return None

    company_column = optional_source_column("Da\u011f\u0131t\u0131m B\u00f6lgesi", "Da\u011f\u0131t\u0131m \u015eirketi")
    cancellation_columns = [
        column
        for header_key, matching_columns in source_header_columns.items()
        if "muhasebe" in header_key and "iptal" in header_key and "tarih" in header_key
        for column in matching_columns
    ]

    field_names = {
        key: _normalize_tariff_lookup_key(label)
        for key, label in {
            "etso": "ETSO",
            "customer": "M\u00fc\u015fteri",
            "tariff": "Tarife Grubu",
            "connection": "AG-OG",
            "term": "Terim",
            "installed_power": "Kurulu G\u00fc\u00e7",
            "contract_power": "S\u00f6zle\u015fme G\u00fcc\u00fc",
            "energy": "Aktif Enerji (kwh)",
            "active_energy_fee": "Aktif Enerji TL",
            "distribution_fee": "Da\u011f\u0131t\u0131m Bedeli",
            "power_fee": "G\u00fc\u00e7 Bedeli",
            "power_overrun_fee": "G\u00fc\u00e7 a\u015f\u0131m Bedeli",
            "reactive_fee": "Reaktif Bedel",
            "btv_fee": "BTV",
            "price_difference": "Fiyat Fark\u0131",
            "vat": "Kdv",
            "vat_base": "KDV Matrah\u0131",
            "total": "TOPLAM",
        }.items()
    }
    maximum_fields = {"installed_power", "contract_power"}
    sum_fields = {
        "energy", "active_energy_fee", "distribution_fee", "power_fee", "power_overrun_fee", "reactive_fee",
        "btv_fee", "price_difference",
        "vat", "vat_base", "total",
    }
    text_fields = {"customer", "tariff", "connection", "term"}
    aggregates: dict[str, dict[str, Any]] = {}
    included_source_row_count = 0

    def etso_key(value: Any) -> str:
        if value is None:
            return ""
        return "".join(character for character in str(value).strip().upper() if character.isalnum())

    for row_index in range(2, source_worksheet.max_row + 1):
        if any(
            (value := source_worksheet.cell(row_index, column).value) is not None
            and str(value).strip()
            for column in cancellation_columns
        ):
            continue
        etso_value = source_worksheet.cell(row_index, source_columns[field_names["etso"]]).value
        key = etso_key(etso_value)
        if not key:
            continue
        included_source_row_count += 1
        aggregate = aggregates.setdefault(
            key,
            {
                "values": {"etso": etso_value},
                "sums": {},
                "sum_counts": {},
                "maxima": {},
                "rows": 0,
            },
        )
        aggregate["rows"] += 1
        for field in text_fields:
            value = source_worksheet.cell(row_index, source_columns[field_names[field]]).value
            if value is None or not str(value).strip():
                continue
            correction_key = field_names["tariff"] if field == "tariff" else field_names[field]
            corrected_value = corrections.get(correction_key, {}).get(
                _normalize_tariff_lookup_key(value),
                value,
            )
            if not aggregate["values"].get(field):
                aggregate["values"][field] = corrected_value
        if company_column is not None:
            value = source_worksheet.cell(row_index, company_column).value
            if value is not None and str(value).strip() and not aggregate["values"].get("company"):
                aggregate["values"]["company"] = value

        for field in maximum_fields:
            value = source_worksheet.cell(row_index, source_columns[field_names[field]]).value
            if value is None or not str(value).strip():
                continue
            number = safe_float(value)
            current = aggregate["maxima"].get(field)
            aggregate["maxima"][field] = number if current is None else max(current, number)
        for field in sum_fields:
            value = source_worksheet.cell(row_index, source_columns[field_names[field]]).value
            if value is None or not str(value).strip():
                continue
            aggregate["sums"][field] = aggregate["sums"].get(field, 0.0) + safe_float(value)
            aggregate["sum_counts"][field] = aggregate["sum_counts"].get(field, 0) + 1

    if not aggregates:
        return {"comparison_sheet_warning": "B\u00f6lge Bizim Kesti\u011fimiz kayna\u011f\u0131nda ETSO kodlu veri sat\u0131r\u0131 bulunamad\u0131."}

    title = comparison_sheet_base[:31]
    suffix = 2
    while title in report_workbook.sheetnames:
        suffix_text = f" {suffix}"
        title = f"{comparison_sheet_base[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    worksheet = report_workbook.create_sheet(title)
    output_headers = [
        "Da\u011f\u0131t\u0131m \u015eirketi", "Etso", "M\u00fc\u015fteri", "Tarife Grubu",
        "Terim Tipi", "Ba\u011flant\u0131 Tipi", "S\u00f6zle\u015fme G\u00fcc\u00fc (kW)",
        "Aktif Enerji T\u00fcketim (kWh)", "Aktif Enerji TL",
        "Da\u011f\u0131t\u0131m Bedeli (TL)", "G\u00fc\u00e7 Bedeli (TL)",
        "G\u00fc\u00e7 A\u015f\u0131m Bedeli (TL)", "Reaktif Bedel (TL)", "BTV Bedeli (TL)",
        "Fiyat Fark\u0131 (TL)", "KDV", "KDV Matrah\u0131 (TL)", "Toplam (TL)",
        "Kurulu G\u00fc\u00e7 (kW)",
    ]
    for column, value in enumerate(output_headers, start=1):
        cell = worksheet.cell(1, column, value)
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="305496")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def compact_number(value: float | None) -> Any:
        if value is None:
            return None
        return int(value) if value.is_integer() else value

    for row_index, key in enumerate(sorted(aggregates), start=2):
        aggregate = aggregates[key]
        values = aggregate["values"]
        worksheet.cell(row_index, 1, values.get("company"))
        worksheet.cell(row_index, 2, values.get("etso"))
        worksheet.cell(row_index, 3, values.get("customer"))
        worksheet.cell(row_index, 4, values.get("tariff"))
        worksheet.cell(row_index, 5, values.get("term"))
        worksheet.cell(row_index, 6, values.get("connection"))
        worksheet.cell(row_index, 7, compact_number(aggregate["maxima"].get("contract_power")))
        worksheet.cell(row_index, 8, compact_number(
            aggregate["sums"].get("energy") if aggregate["sum_counts"].get("energy") else None
        ))
        worksheet.cell(row_index, 9, compact_number(
            aggregate["sums"].get("active_energy_fee")
            if aggregate["sum_counts"].get("active_energy_fee") else None
        ))
        worksheet.cell(row_index, 10, compact_number(
            aggregate["sums"].get("distribution_fee") if aggregate["sum_counts"].get("distribution_fee") else None
        ))
        worksheet.cell(row_index, 11, compact_number(
            aggregate["sums"].get("power_fee") if aggregate["sum_counts"].get("power_fee") else None
        ))
        worksheet.cell(row_index, 12, compact_number(
            aggregate["sums"].get("power_overrun_fee") if aggregate["sum_counts"].get("power_overrun_fee") else None
        ))
        worksheet.cell(row_index, 13, compact_number(
            aggregate["sums"].get("reactive_fee") if aggregate["sum_counts"].get("reactive_fee") else None
        ))
        worksheet.cell(row_index, 14, compact_number(
            aggregate["sums"].get("btv_fee") if aggregate["sum_counts"].get("btv_fee") else None
        ))
        worksheet.cell(row_index, 15, compact_number(
            aggregate["sums"].get("price_difference")
            if aggregate["sum_counts"].get("price_difference") else None
        ))
        worksheet.cell(row_index, 16, compact_number(
            aggregate["sums"].get("vat") if aggregate["sum_counts"].get("vat") else None
        ))
        worksheet.cell(row_index, 17, compact_number(
            aggregate["sums"].get("vat_base") if aggregate["sum_counts"].get("vat_base") else None
        ))
        worksheet.cell(row_index, 18, compact_number(
            aggregate["sums"].get("total") if aggregate["sum_counts"].get("total") else None
        ))
        worksheet.cell(row_index, 19, compact_number(aggregate["maxima"].get("installed_power")))
        for column in range(1, 20):
            worksheet.cell(row_index, column).alignment = Alignment(vertical="center")
        for column in range(7, 20):
            worksheet.cell(row_index, column).number_format = '#,##0.00;[Red]-#,##0.00;0.00'

    last_row = len(aggregates) + 1
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:S{last_row}"
    worksheet.sheet_view.showGridLines = True
    worksheet.row_dimensions[1].height = 34
    widths = [22, 23, 42, 27, 17, 17, 23, 24, 24, 23, 20, 24, 20, 20, 20, 15, 22, 20, 20]
    for column, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width

    return {
        "comparison_sheet_name": title,
        "region_invoice_source_row_count": included_source_row_count,
        "region_invoice_etso_count": len(aggregates),
        "region_invoice_duplicate_row_count": included_source_row_count - len(aggregates),
    }


def append_selected_invoice_file(
    output_bytes: bytes,
    file_path: str | None,
    mapping_path: str | None = None,
    system_usage_file: str | None = None,
    reconciliation_consumption_file: str | None = None,
) -> tuple[bytes, list[dict[str, Any]], list[str], int]:
    """Copy the selected invoice workbook and add tariff-based distribution amounts."""
    if not file_path:
        return output_bytes, [], [], 0
    if not os.path.isfile(file_path):
        return output_bytes, [], [f"Se\u00e7ilen Excel dosyas\u0131 bulunamad\u0131: {file_path}"], 0

    extension = Path(file_path).suffix.casefold()
    if extension not in {".xlsx", ".xlsm", ".xls"}:
        return output_bytes, [], [f"Desteklenmeyen Excel dosya t\u00fcr\u00fc: {Path(file_path).name}"], 0

    from copy import copy
    from openpyxl import load_workbook

    tariff_rates, tariff_rate_error = load_system_usage_distribution_rates(
        mapping_path,
        system_usage_file or str(SYSTEM_USAGE_FILE_DEFAULT),
    )
    tariff_layout, power_rate_error = load_system_usage_power_rates(
        mapping_path,
        system_usage_file or str(SYSTEM_USAGE_FILE_DEFAULT),
    )
    report_workbook = load_workbook(io.BytesIO(output_bytes), data_only=False)
    source_workbook = None
    source_values_workbook = None
    imported_sheets: list[dict[str, Any]] = []
    try:
        if extension in {".xlsx", ".xlsm"}:
            source_workbook = load_workbook(
                file_path,
                data_only=False,
                read_only=True,
                keep_vba=extension == ".xlsm",
            )
            source_values_workbook = load_workbook(
                file_path,
                data_only=True,
                read_only=True,
                keep_vba=extension == ".xlsm",
            )
            source_sheets = [
                (
                    sheet.title,
                    sheet.iter_rows(values_only=False),
                    source_values_workbook[sheet.title],
                )
                for sheet in source_workbook.worksheets
            ]
            for sheet_number, (source_sheet_name, source_rows, source_values_worksheet) in enumerate(source_sheets, start=1):
                invoice_format = detect_our_invoice_worksheet_format(
                    source_workbook[source_sheet_name]
                )
                if invoice_format is None:
                    return output_bytes, [], [
                        f"{source_sheet_name}: Excel ba\u015fl\u0131klar\u0131 desteklenen fatura bi\u00e7imleriyle e\u015fle\u015fmedi."
                    ], 0
                is_region_summary = invoice_format == "region_summary"
                base_title = OUR_INVOICES_LABEL if sheet_number == 1 else f"{OUR_INVOICES_LABEL} {sheet_number}"
                title = base_title[:31]
                unique_title = title
                suffix = 2
                while unique_title in report_workbook.sheetnames:
                    suffix_text = f" {suffix}"
                    unique_title = f"{base_title[:31 - len(suffix_text)]}{suffix_text}"
                    suffix += 1
                target_sheet = report_workbook.create_sheet(unique_title)
                style_cache: dict[int, Any] = {}

                for row in source_rows:
                    for source_cell in row:
                        if source_cell.value is None:
                            continue
                        target_cell = target_sheet.cell(
                            source_cell.row,
                            source_cell.column,
                            source_cell.value,
                        )
                        target_cell.data_type = source_cell.data_type
                        if source_cell.has_style:
                            style_id = source_cell._style_id
                            cached_style = style_cache.get(style_id)
                            if cached_style is None:
                                target_cell.font = copy(source_cell.font)
                                target_cell.fill = copy(source_cell.fill)
                                target_cell.border = copy(source_cell.border)
                                target_cell.alignment = copy(source_cell.alignment)
                                target_cell.protection = copy(source_cell.protection)
                                target_cell.number_format = source_cell.number_format
                                style_cache[style_id] = copy(target_cell._style)
                            else:
                                target_cell._style = cached_style
                if is_region_summary:
                    calculation_stats = {
                        "calculated_distribution_amount_count": 0,
                        "missing_distribution_rate_count": 0,
                        "distribution_amount_warning": None,
                    }
                    comparison_stats = add_region_invoice_summary_worksheet(
                        report_workbook,
                        target_sheet,
                        "Bizim Kesti\u011fimiz 2" if sheet_number == 1 else f"Bizim Kesti\u011fimiz 2 {sheet_number}",
                        mapping_path,
                    )
                else:
                    calculation_stats = add_expected_distribution_charge_column(
                        target_sheet,
                        tariff_rates,
                    )
                    comparison_stats = add_our_invoice_comparison_worksheet(
                        report_workbook,
                        target_sheet,
                        "Bizim Kestiğimiz 2" if sheet_number == 1 else f"Bizim Kestiğimiz 2 {sheet_number}",
                        tariff_rates,
                        tariff_layout,
                        power_rate_error,
                        source_values_worksheet=source_values_worksheet,
                    )
                imported_sheets.append({
                    "file_path": file_path,
                    "source_sheet_name": source_sheet_name,
                    "output_sheet_name": unique_title,
                    "invoice_format": invoice_format,
                    **calculation_stats,
                    **comparison_stats,
                    "tariff_rate_error": tariff_rate_error if not is_region_summary else None,
                })
        else:
            import xlrd

            source_workbook = xlrd.open_workbook(file_path, on_demand=True)
            for sheet_number, source_sheet in enumerate(source_workbook.sheets(), start=1):
                invoice_format = detect_our_invoice_format(
                    source_sheet.cell_value(0, column_index)
                    for column_index in range(source_sheet.ncols)
                )
                if invoice_format is None:
                    return output_bytes, [], [
                        f"{source_sheet.name}: Excel ba\u015fl\u0131klar\u0131 desteklenen fatura bi\u00e7imleriyle e\u015fle\u015fmedi."
                    ], 0
                is_region_summary = invoice_format == "region_summary"
                base_title = OUR_INVOICES_LABEL if sheet_number == 1 else f"{OUR_INVOICES_LABEL} {sheet_number}"
                title = base_title[:31]
                unique_title = title
                suffix = 2
                while unique_title in report_workbook.sheetnames:
                    suffix_text = f" {suffix}"
                    unique_title = f"{base_title[:31 - len(suffix_text)]}{suffix_text}"
                    suffix += 1
                target_sheet = report_workbook.create_sheet(unique_title)
                for row_index in range(source_sheet.nrows):
                    for column_index in range(source_sheet.ncols):
                        source_cell = source_sheet.cell(row_index, column_index)
                        if source_cell.ctype == xlrd.XL_CELL_EMPTY:
                            continue
                        target_cell = target_sheet.cell(row_index + 1, column_index + 1, source_cell.value)
                        if source_cell.ctype == xlrd.XL_CELL_TEXT:
                            target_cell.data_type = "s"
                        elif source_cell.ctype == xlrd.XL_CELL_ERROR:
                            target_cell.data_type = "e"
                        elif source_cell.ctype == xlrd.XL_CELL_BOOLEAN:
                            target_cell.value = bool(source_cell.value)
                if is_region_summary:
                    calculation_stats = {
                        "calculated_distribution_amount_count": 0,
                        "missing_distribution_rate_count": 0,
                        "distribution_amount_warning": None,
                    }
                    comparison_stats = add_region_invoice_summary_worksheet(
                        report_workbook,
                        target_sheet,
                        "Bizim Kesti\u011fimiz 2" if sheet_number == 1 else f"Bizim Kesti\u011fimiz 2 {sheet_number}",
                        mapping_path,
                    )
                else:
                    calculation_stats = add_expected_distribution_charge_column(
                        target_sheet,
                        tariff_rates,
                    )
                    comparison_stats = add_our_invoice_comparison_worksheet(
                        report_workbook,
                        target_sheet,
                        "Bizim Kestiğimiz 2" if sheet_number == 1 else f"Bizim Kestiğimiz 2 {sheet_number}",
                        tariff_rates,
                        tariff_layout,
                        power_rate_error,
                    )
                imported_sheets.append({
                    "file_path": file_path,
                    "source_sheet_name": source_sheet.name,
                    "output_sheet_name": unique_title,
                    "invoice_format": invoice_format,
                    **calculation_stats,
                    **comparison_stats,
                    "tariff_rate_error": tariff_rate_error if not is_region_summary else None,
                })

        comparison_error = next(
            (source.get("comparison_sheet_warning") for source in imported_sheets if source.get("comparison_sheet_warning")),
            None,
        )
        if comparison_error:
            return output_bytes, [], [comparison_error], 0

        comparison_sheet_names = [
            source.get("comparison_sheet_name")
            for source in imported_sheets
            if source.get("comparison_sheet_name")
        ]
        has_distribution_detail = any(
            _normalize_tariff_lookup_key(sheet.title)
            == _normalize_tariff_lookup_key("Da\u011f\u0131t\u0131m\u0131n Kesti\u011fi")
            for sheet in report_workbook.worksheets
        )
        if comparison_sheet_names and has_distribution_detail:
            reconciliation_stats = add_invoice_distribution_comparison_worksheet(
                report_workbook,
                comparison_sheet_names,
            )
            if reconciliation_stats.get("reconciliation_sheet_warning"):
                return output_bytes, [], [reconciliation_stats["reconciliation_sheet_warning"]], 0
            if imported_sheets:
                imported_sheets[0].update(reconciliation_stats)

        if reconciliation_consumption_file:
            try:
                if not os.path.isfile(reconciliation_consumption_file):
                    raise FileNotFoundError(
                        f"Uzlaştırma Tüketim dosyası bulunamadı: {reconciliation_consumption_file}"
                    )
                consumption_by_etso, consumption_source_stats = load_reconciliation_consumption_by_etso(
                    reconciliation_consumption_file
                )
                comparison_worksheet = next(
                    (
                        sheet for sheet in report_workbook.worksheets
                        if _normalize_tariff_lookup_key(sheet.title).startswith(
                            _normalize_tariff_lookup_key("Karşılaştırma")
                        )
                    ),
                    None,
                )
                if comparison_worksheet is not None:
                    consumption_comparison_stats = add_reconciliation_consumption_columns(
                        report_workbook,
                        comparison_worksheet,
                        consumption_by_etso,
                    )
                elif comparison_sheet_names:
                    consumption_comparison_stats = add_reconciliation_only_comparison_worksheet(
                        report_workbook,
                        comparison_sheet_names,
                        consumption_by_etso,
                    )
                else:
                    consumption_comparison_stats = {
                        "reconciliation_consumption_warning": (
                            "Karşılaştırma sayfası oluşturulamadı; eşleştirilecek fatura sayfası bulunamadı."
                        )
                    }
                if imported_sheets:
                    imported_sheets[0].update(consumption_source_stats)
                    imported_sheets[0].update(consumption_comparison_stats)
            except Exception as error:
                if imported_sheets:
                    imported_sheets[0]["reconciliation_consumption_warning"] = str(error)

        report_workbook.calculation.calcMode = "auto"
        report_workbook.calculation.fullCalcOnLoad = True
        report_workbook.calculation.forceFullCalc = True
        report_workbook.calculation.calcOnSave = True
        result_buffer = io.BytesIO()
        report_workbook.save(result_buffer)
        return result_buffer.getvalue(), imported_sheets, [], 1
    except Exception as error:
        return output_bytes, [], [f"{Path(file_path).name}: {error}"], 0
    finally:
        if source_workbook is not None:
            if extension == ".xls":
                source_workbook.release_resources()
            else:
                source_workbook.close()
        if source_values_workbook is not None:
            source_values_workbook.close()
        report_workbook.close()

def add_our_invoice_comparison_worksheet(
    report_workbook: Any,
    source_worksheet: Any,
    comparison_sheet_base: str,
    distribution_rates: dict[str, float],
    tariff_layout: dict[str, dict[str, Any]],
    power_rate_error: str | None = None,
    source_values_worksheet: Any | None = None,
) -> dict[str, Any]:
    """Append the formula-driven comparison view beside each imported invoice sheet."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    headers_by_key = {
        _normalize_tariff_lookup_key(cell.value): cell.column
        for cell in source_worksheet[1]
        if cell.value is not None
    }

    def find_column(*aliases: str) -> int | None:
        return next(
            (
                headers_by_key[_normalize_tariff_lookup_key(alias)]
                for alias in aliases
                if _normalize_tariff_lookup_key(alias) in headers_by_key
            ),
            None,
        )

    source_aliases = {
        1: ("Dağıtım Bölgesi", "Dağıtım Şirketi"),
        2: ("ETSO", "ETSO Kodu"),
        3: ("Müşteri", "Ünvan"),
        4: ("Tarife Grubu",),
        5: ("Abone Grubu",),
        9: ("Aktif Enerji Tüketim (kWh)",),
        10: ("Olması Gereken Dağıtım Bedeli",),
        11: ("TL Güç", "Güç Bedeli (TL)", "Güç Bedeli(TL)"),
        12: ("Güç Aşım Bedeli (TL)",),
        13: ("Reaktif Bedel (TL)",),
        17: ("Tük Güç", "TÜK GÜÇ"),
        19: ("Durum", "Status"),
    }
    source_column_map = {
        destination_column: find_column(*aliases)
        for destination_column, aliases in source_aliases.items()
    }
    expected_fields = {
        1: "Dağıtım Şirketi",
        2: "ETSO",
        3: "Müşteri",
        4: "Tarife Grubu",
        5: "Abone Grubu",
        9: "Aktif Enerji Tüketim (kWh)",
        10: "Olması Gereken Dağıtım Bedeli",
        11: "Güç Bedeli (TL)",
        12: "Güç Aşım Bedeli (TL)",
        13: "Reaktif Bedel (TL)",
        17: "Tük Güç",
        19: "Durum",
    }
    missing_fields = [
        label for column, label in expected_fields.items()
        if source_column_map.get(column) is None
    ]
    if missing_fields:
        return {
            "comparison_sheet_warning": (
                "Bizim Kestiğimiz 2 oluşturulamadı; kaynakta eksik başlıklar: "
                + ", ".join(missing_fields)
            )
        }
    if not tariff_layout:
        return {
            "comparison_sheet_warning": (
                power_rate_error or "Tarife eşlemesi bulunamadığı için karşılaştırma sayfası oluşturulamadı."
            )
        }

    value_source_columns = {}
    if source_values_worksheet is not None:
        value_headers_by_key = {
            _normalize_tariff_lookup_key(cell.value): cell.column
            for cell in source_values_worksheet[1]
            if cell.value is not None
        }
        value_source_columns = {
            destination_column: next(
                (
                    value_headers_by_key[_normalize_tariff_lookup_key(alias)]
                    for alias in aliases
                    if _normalize_tariff_lookup_key(alias) in value_headers_by_key
                ),
                None,
            )
            for destination_column, aliases in source_aliases.items()
            if destination_column != 10
        }

    title = comparison_sheet_base[:31]
    suffix = 2
    while title in report_workbook.sheetnames:
        suffix_text = f" {suffix}"
        title = f"{comparison_sheet_base[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    worksheet = report_workbook.create_sheet(title)

    output_headers = [
        "Dağıtım Şirketi", "Etso", "Müşteri", "Tarife Grubu", "Abone Grubu",
        "Terim Tipi", "Bağlantı Tipi", "Güç kWh", "Aktif Enerji Tüketim (kWh)",
        "Dağıtım Bedeli (TL)", "Güç Bedeli(TL)", "Güç Aşım Bedeli (TL)",
        "Reaktif Bedel (TL)", "KDV", "KDV Matrahı (TL)", "Toplam (TL)",
        "TÜK GÜÇ", "Güç Fark", "Durum",
    ]
    for column, value in enumerate(output_headers, start=1):
        cell = worksheet.cell(1, column, value)
        cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(fill_type="solid", fgColor="305496")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    helper_headers = [
        "Tarife Grubu", "Terim Tipi", "Bağlantı Tipi",
        "Güç Bedeli (kr/kW/ay)", "Dağıtım Bedeli (kr/kWh)",
    ]
    for column, value in enumerate(helper_headers, start=21):
        worksheet.cell(1, column, value)

    ordered_tariffs = list(tariff_layout.values())
    lookup_end = len(ordered_tariffs) + 1
    for helper_row, tariff in enumerate(ordered_tariffs, start=2):
        tariff_name = tariff["tariff_group"]
        worksheet.cell(helper_row, 21, tariff_name)
        worksheet.cell(helper_row, 22, tariff.get("term_type") or "")
        worksheet.cell(helper_row, 23, tariff.get("connection_type") or "")
        worksheet.cell(helper_row, 24, tariff.get("power_rate") or 0)
        worksheet.cell(
            helper_row,
            25,
            distribution_rates.get(str(tariff_name).strip().casefold()),
        )
    for column in range(21, 26):
        worksheet.column_dimensions[get_column_letter(column)].hidden = True

    worksheet_ref = "'" + title.replace("'", "''") + "'"
    source_expected_col = find_column("Olması Gereken Dağıtım Bedeli")
    expected_letter = get_column_letter(source_expected_col) if source_expected_col else None
    energy_letter = get_column_letter(source_column_map[9])
    tariff_letter = get_column_letter(source_column_map[4])
    last_row = source_worksheet.max_row
    source_value_rows = None
    max_value_column = max(
        [
            source_column_map[destination_column]
            for destination_column in source_column_map
            if destination_column != 10
        ]
        + [column for column in value_source_columns.values() if column is not None],
        default=1,
    )
    if source_values_worksheet is not None:
        source_value_rows = iter(
            source_values_worksheet.iter_rows(
                min_row=2,
                max_row=last_row,
                max_col=max_value_column,
                values_only=True,
            )
        )

    if expected_letter is not None:
        for row_index in range(2, last_row + 1):
            source_worksheet.cell(
                row_index,
                source_expected_col,
                f'=IF(OR({energy_letter}{row_index}="",{tariff_letter}{row_index}=""),"",'
                f'IFERROR({energy_letter}{row_index}*VLOOKUP({tariff_letter}{row_index},'
                f'{worksheet_ref}!$U$2:$Y${lookup_end},5,FALSE),""))',
            )

    for row_index in range(2, last_row + 1):
        value_row = next(source_value_rows, None) if source_value_rows is not None else None
        for destination_column, source_column in source_column_map.items():
            if destination_column == 10:
                tariff_value = source_worksheet.cell(
                    row_index,
                    source_column_map[4],
                ).value
                energy_source_column = value_source_columns.get(9) or source_column_map[9]
                if source_values_worksheet is not None:
                    energy_value = (
                        value_row[energy_source_column - 1]
                        if value_row is not None and energy_source_column <= len(value_row)
                        else None
                    )
                else:
                    energy_value = source_worksheet.cell(row_index, energy_source_column).value
                rate = distribution_rates.get(
                    str(tariff_value).strip().casefold() if tariff_value is not None else ""
                )
                value = (
                    safe_float(energy_value) * rate / 100
                    if rate is not None and energy_value not in (None, "")
                    else None
                )
            else:
                value_source_column = value_source_columns.get(destination_column) or source_column
                if source_values_worksheet is not None:
                    value = (
                        value_row[value_source_column - 1]
                        if value_row is not None and value_source_column <= len(value_row)
                        else None
                    )
                else:
                    value = source_worksheet.cell(row_index, value_source_column).value
            if destination_column == 5:
                value = normalize_invoice_subscription_group(value)
            worksheet.cell(row_index, destination_column, value)

        tariff_value = worksheet.cell(row_index, 4).value
        tariff_info = tariff_layout.get(_normalize_tariff_lookup_key(tariff_value), {})
        worksheet.cell(row_index, 6, tariff_info.get("term_type") or "")
        worksheet.cell(row_index, 7, tariff_info.get("connection_type") or "")
        worksheet.cell(
            row_index,
            8,
            f'=IFERROR(IF(OR($D{row_index}="",K{row_index}=""),0,'
            f'K{row_index}/(VLOOKUP($D{row_index},$U$2:$Y${lookup_end},4,FALSE)/100)),0)',
        )
        worksheet.cell(row_index, 14, f'=IF(O{row_index}="","",O{row_index}*0.2)')
        worksheet.cell(
            row_index,
            15,
            f'=IF(COUNTA(J{row_index}:M{row_index})=0,"",SUM(J{row_index}:M{row_index}))',
        )
        worksheet.cell(row_index, 16, f'=IF(O{row_index}="","",O{row_index}+N{row_index})')
        worksheet.cell(
            row_index,
            18,
            f'=IF(OR(H{row_index}="",Q{row_index}=""),"",H{row_index}-Q{row_index})',
        )

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:S{last_row}"
    worksheet.sheet_view.showGridLines = True
    worksheet.row_dimensions[1].height = 34
    widths = [22, 23, 42, 27, 22, 17, 17, 15, 24, 23, 20, 24, 20, 15, 22, 20, 15, 15, 18]
    for column, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width
    numeric_format = '#,##0.00;[Red]-#,##0.00;0.00'
    body_alignment = Alignment(vertical="center")
    for row_index in range(2, last_row + 1):
        for column in range(1, 20):
            worksheet.cell(row_index, column).alignment = body_alignment
        for column in range(8, 19):
            worksheet.cell(row_index, column).number_format = numeric_format

    report_workbook.calculation.calcMode = "auto"
    report_workbook.calculation.fullCalcOnLoad = True
    report_workbook.calculation.forceFullCalc = True
    report_workbook.calculation.calcOnSave = True
    result = {"comparison_sheet_name": title}
    if power_rate_error:
        result["power_rate_error"] = power_rate_error
    return result


def reconciliation_etso_key(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        text = str(value)
    elif isinstance(value, (int, float)) and math.isfinite(float(value)):
        numeric = float(value)
        text = str(int(numeric)) if numeric.is_integer() else str(value)
    else:
        text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return "".join(character for character in text.strip().upper() if character.isalnum())


def load_reconciliation_consumption_by_etso(
    file_path: str,
) -> tuple[dict[str, float], dict[str, Any]]:
    """Aggregate eligible meter consumption by ETSO, converting MWh to kWh."""
    extension = Path(file_path).suffix.casefold()
    if extension not in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError("Uzlaştırma dosyası desteklenmeyen bir Excel biçiminde.")

    headers_needed = {
        "etso": {"sayacetso"},
        "draw": {"cekisorg"},
        "consumption": {"kayiplicekismwh"},
    }
    aggregates_mwh: dict[str, float] = {}
    qualifying_row_count = 0
    workbook = None
    if extension in {".xlsx", ".xlsm"}:
        from openpyxl import load_workbook

        workbook = load_workbook(file_path, read_only=True, data_only=True, keep_vba=extension == ".xlsm")
        try:
            found_headers = False
            for worksheet in workbook.worksheets:
                header_row_index = None
                columns: dict[str, int] = {}
                for row_index, row in enumerate(worksheet.iter_rows(max_row=12, values_only=True), start=1):
                    normalized = {
                        _normalize_tariff_lookup_key(value): column_index
                        for column_index, value in enumerate(row, start=1)
                        if value is not None
                    }
                    etso_column = next((normalized[name] for name in headers_needed["etso"] if name in normalized), None)
                    draw_column = next((normalized[name] for name in headers_needed["draw"] if name in normalized), None)
                    consumption_column = next((normalized[name] for name in headers_needed["consumption"] if name in normalized), None)
                    if etso_column and draw_column and consumption_column:
                        header_row_index = row_index
                        columns = {"etso": etso_column, "draw": draw_column, "consumption": consumption_column}
                        break
                if header_row_index is None:
                    continue
                found_headers = True
                for row in worksheet.iter_rows(min_row=header_row_index + 1, values_only=True):
                    draw_value = row[columns["draw"] - 1] if len(row) >= columns["draw"] else None
                    if _normalize_tariff_lookup_key(draw_value) != "evet":
                        continue
                    code_value = row[columns["etso"] - 1] if len(row) >= columns["etso"] else None
                    code = reconciliation_etso_key(code_value)
                    if not code:
                        continue
                    consumption_value = row[columns["consumption"] - 1] if len(row) >= columns["consumption"] else None
                    aggregates_mwh[code] = aggregates_mwh.get(code, 0.0) + safe_float(consumption_value)
                    qualifying_row_count += 1
            if not found_headers:
                raise ValueError("Dosyada Sayaç ETSO, Çekiş Org ve Kayıplı Çekiş (MWh) başlıkları bulunamadı.")
        finally:
            workbook.close()
    else:
        import xlrd

        workbook = xlrd.open_workbook(file_path, on_demand=True)
        try:
            found_headers = False
            for worksheet in workbook.sheets():
                header_row_index = None
                columns: dict[str, int] = {}
                for row_index in range(min(worksheet.nrows, 12)):
                    normalized = {
                        _normalize_tariff_lookup_key(worksheet.cell_value(row_index, column_index)): column_index
                        for column_index in range(worksheet.ncols)
                        if worksheet.cell_value(row_index, column_index) not in (None, "")
                    }
                    etso_column = next((normalized[name] for name in headers_needed["etso"] if name in normalized), None)
                    draw_column = next((normalized[name] for name in headers_needed["draw"] if name in normalized), None)
                    consumption_column = next((normalized[name] for name in headers_needed["consumption"] if name in normalized), None)
                    if etso_column is not None and draw_column is not None and consumption_column is not None:
                        header_row_index = row_index
                        columns = {"etso": etso_column, "draw": draw_column, "consumption": consumption_column}
                        break
                if header_row_index is None:
                    continue
                found_headers = True
                for row_index in range(header_row_index + 1, worksheet.nrows):
                    draw_value = worksheet.cell_value(row_index, columns["draw"])
                    if _normalize_tariff_lookup_key(draw_value) != "evet":
                        continue
                    code = reconciliation_etso_key(worksheet.cell_value(row_index, columns["etso"]))
                    if not code:
                        continue
                    consumption_value = worksheet.cell_value(row_index, columns["consumption"])
                    aggregates_mwh[code] = aggregates_mwh.get(code, 0.0) + safe_float(consumption_value)
                    qualifying_row_count += 1
            if not found_headers:
                raise ValueError("Dosyada Sayaç ETSO, Çekiş Org ve Kayıplı Çekiş (MWh) başlıkları bulunamadı.")
        finally:
            workbook.release_resources()

    return (
        {code: value * 1000 for code, value in aggregates_mwh.items()},
        {
            "reconciliation_consumption_source_row_count": qualifying_row_count,
            "reconciliation_consumption_etso_count": len(aggregates_mwh),
        },
    )


def add_reconciliation_consumption_columns(
    report_workbook: Any,
    comparison_sheet: Any,
    consumption_by_etso: dict[str, float],
) -> dict[str, Any]:
    """Append reconciliation consumption and signed differences to a comparison sheet."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    def resolve_cell_value(
        worksheet: Any,
        row_index: int,
        column_index: int,
        visited: set[tuple[str, int, int]] | None = None,
        depth: int = 0,
    ) -> Any:
        if depth > 12:
            return None
        visited = visited or set()
        key = (worksheet.title, row_index, column_index)
        if key in visited:
            return None
        visited.add(key)
        value = worksheet.cell(row_index, column_index).value
        if not (isinstance(value, str) and value.startswith("=")):
            return value

        formula = value.strip()
        references = re.findall(
            r"'((?:[^']|'')+)'!\$?([A-Z]{1,3})\$?(\d+)",
            formula,
        )
        unique_references = set(references)
        if len(unique_references) == 1:
            sheet_name, column_letters, source_row = next(iter(unique_references))
            source_sheet = report_workbook[sheet_name.replace("''", "'")]
            from openpyxl.utils import column_index_from_string

            return resolve_cell_value(
                source_sheet,
                int(source_row),
                column_index_from_string(column_letters),
                visited.copy(),
                depth + 1,
            )

        local_arithmetic = re.fullmatch(
            r"=\$?([A-Z]{1,3})\$?(\d+)([+\-*/])\$?([A-Z]{1,3})\$?(\d+)",
            formula,
        )
        if local_arithmetic:
            first_column, first_row, operator, second_column, second_row = local_arithmetic.groups()
            from openpyxl.utils import column_index_from_string

            first_value = resolve_cell_value(
                worksheet, int(first_row), column_index_from_string(first_column), visited.copy(), depth + 1
            )
            second_value = resolve_cell_value(
                worksheet, int(second_row), column_index_from_string(second_column), visited.copy(), depth + 1
            )
            try:
                first_number = safe_float(first_value)
                second_number = safe_float(second_value)
                if operator == "+":
                    return first_number + second_number
                if operator == "-":
                    return first_number - second_number
                if operator == "*":
                    return first_number * second_number
                if second_number != 0:
                    return first_number / second_number
            except (TypeError, ValueError, ZeroDivisionError):
                return None
        return None

    def numeric_value(value: Any) -> float | None:
        if value is None or isinstance(value, bool):
            return None
        if isinstance(value, str) and value.strip() in {"", "-", "–", "—"}:
            return None
        try:
            return safe_float(value)
        except (TypeError, ValueError):
            return None

    headers = {
        _normalize_tariff_lookup_key(cell.value): cell.column
        for cell in comparison_sheet[1]
        if cell.value is not None
    }

    def find_column(*aliases: str) -> int | None:
        return next(
            (
                headers[_normalize_tariff_lookup_key(alias)]
                for alias in aliases
                if _normalize_tariff_lookup_key(alias) in headers
            ),
            None,
        )

    etso_column = find_column("ETSO", "ETSO Kodu")
    own_energy_column = find_column("Aktif Enerji Tüketim (kWh)")
    distribution_energy_column = find_column("Aktif Enerji Tüketim (kWh) Dağıtım")
    company_column = find_column("Dağıtım Şirketi", "Dağıtım Bölgesi")
    customer_column = find_column("Müşteri", "Ünvan")
    tariff_column = find_column("Tarife Grubu")
    if etso_column is None or own_energy_column is None:
        raise ValueError("Karşılaştırma sayfasında ETSO veya bizim tüketim başlığı bulunamadı.")

    new_headers = [
        "Uzlaştırma Tüketim (kWh)",
        "Bizim Kestiğimiz - Uzlaştırma Tüketim Farkı (kWh)",
        "Dağıtım - Uzlaştırma Tüketim Farkı (kWh)",
    ]
    start_column = comparison_sheet.max_column + 1
    for offset, header in enumerate(new_headers):
        cell = comparison_sheet.cell(1, start_column + offset, header)
        cell.font = Font(name="Calibri", size=10, bold=True, color="1F1F1F")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="F4B183" if offset == 0 else "FFD966",
        )

    thin_gray = Side(style="thin", color="D9E1F2")
    row_border = Border(bottom=thin_gray)
    row_alignment = Alignment(vertical="center")
    matched_rows = 0
    unmatched_rows = 0
    unique_matched_codes: set[str] = set()
    own_difference_rows: list[dict[str, Any]] = []
    distribution_difference_rows: list[dict[str, Any]] = []
    for row_index in range(2, comparison_sheet.max_row + 1):
        code = reconciliation_etso_key(comparison_sheet.cell(row_index, etso_column).value)
        value = consumption_by_etso.get(code) if code else None
        consumption_cell = comparison_sheet.cell(row_index, start_column)
        if value is None:
            consumption_cell.value = "-"
            unmatched_rows += 1
        else:
            consumption_cell.value = value
            matched_rows += 1
            unique_matched_codes.add(code)

        own_letter = get_column_letter(own_energy_column)
        reconciliation_letter = get_column_letter(start_column)
        own_difference_cell = comparison_sheet.cell(row_index, start_column + 1)
        own_difference_cell.value = (
            f'=IF(AND(ISNUMBER({own_letter}{row_index}),ISNUMBER({reconciliation_letter}{row_index})),'
            f'{own_letter}{row_index}-{reconciliation_letter}{row_index},"-")'
        )
        distribution_difference_cell = comparison_sheet.cell(row_index, start_column + 2)
        if distribution_energy_column is None:
            distribution_difference_cell.value = "-"
        else:
            distribution_letter = get_column_letter(distribution_energy_column)
            distribution_difference_cell.value = (
                f'=IF(AND(ISNUMBER({distribution_letter}{row_index}),ISNUMBER({reconciliation_letter}{row_index})),'
                f'{distribution_letter}{row_index}-{reconciliation_letter}{row_index},"-")'
            )

        if value is not None:
            common = {
                "ETSO": resolve_cell_value(comparison_sheet, row_index, etso_column),
                "Dağıtım Şirketi": resolve_cell_value(comparison_sheet, row_index, company_column)
                if company_column is not None else "-",
                "Müşteri": resolve_cell_value(comparison_sheet, row_index, customer_column)
                if customer_column is not None else "-",
                "Tarife Grubu": resolve_cell_value(comparison_sheet, row_index, tariff_column)
                if tariff_column is not None else "-",
                "Uzlaştırma Tüketim (kWh)": value,
            }
            own_energy_value = resolve_cell_value(comparison_sheet, row_index, own_energy_column)
            own_number = numeric_value(own_energy_value)
            reconciliation_number = numeric_value(value)
            if own_number is not None and reconciliation_number is not None:
                own_difference = own_number - reconciliation_number
                if abs(own_difference) > 1e-9:
                    own_difference_rows.append({
                        **common,
                        "Bizim Kestiğimiz Tüketim (kWh)": own_number,
                        "Fark (kWh)": own_difference,
                    })
            if distribution_energy_column is not None:
                distribution_value = resolve_cell_value(
                    comparison_sheet, row_index, distribution_energy_column
                )
                distribution_number = numeric_value(distribution_value)
                if distribution_number is not None and reconciliation_number is not None:
                    distribution_difference = distribution_number - reconciliation_number
                    if abs(distribution_difference) > 1e-9:
                        distribution_difference_rows.append({
                            **common,
                            "Dağıtım Tüketim (kWh)": distribution_number,
                            "Fark (kWh)": distribution_difference,
                        })

        for column in range(start_column, start_column + len(new_headers)):
            cell = comparison_sheet.cell(row_index, column)
            cell.alignment = row_alignment
            cell.border = row_border
            cell.number_format = '#,##0.00;[Red]-#,##0.00;0.00'

    comparison_sheet.column_dimensions[get_column_letter(start_column)].width = 24
    comparison_sheet.column_dimensions[get_column_letter(start_column + 1)].width = 39
    comparison_sheet.column_dimensions[get_column_letter(start_column + 2)].width = 39
    comparison_sheet.row_dimensions[1].height = max(comparison_sheet.row_dimensions[1].height or 0, 48)
    comparison_sheet.auto_filter.ref = (
        f"A1:{get_column_letter(comparison_sheet.max_column)}{max(1, comparison_sheet.max_row)}"
    )
    report_workbook.calculation.calcMode = "auto"
    report_workbook.calculation.fullCalcOnLoad = True
    report_workbook.calculation.forceFullCalc = True
    report_workbook.calculation.calcOnSave = True
    return {
        "reconciliation_consumption_sheet_name": comparison_sheet.title,
        "reconciliation_consumption_matched_count": matched_rows,
        "reconciliation_consumption_unmatched_count": unmatched_rows,
        "reconciliation_consumption_matched_etso_count": len(unique_matched_codes),
        "our_vs_reconciliation_difference_rows": own_difference_rows,
        "distribution_vs_reconciliation_difference_rows": distribution_difference_rows,
    }


def add_reconciliation_only_comparison_worksheet(
    report_workbook: Any,
    invoice_sheet_names: list[str],
    consumption_by_etso: dict[str, float],
) -> dict[str, Any]:
    """Build a useful ETSO comparison when distribution invoice data is not loaded."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    title = "Karşılaştırma"
    suffix = 2
    while title in report_workbook.sheetnames:
        suffix_text = f" {suffix}"
        title = f"Karşılaştırma"[:31 - len(suffix_text)] + suffix_text
        suffix += 1
    worksheet = report_workbook.create_sheet(title)
    headers = [
        "Dağıtım Şirketi", "ETSO", "Müşteri", "Tarife Grubu",
        "Aktif Enerji Tüketim (kWh)", "Aktif Enerji Tüketim (kWh) Dağıtım",
        "Uzlaştırma Tüketim (kWh)",
        "Bizim Kestiğimiz - Uzlaştırma Tüketim Farkı (kWh)",
        "Dağıtım - Uzlaştırma Tüketim Farkı (kWh)", "ETSO Eşleşmesi",
    ]
    header_colors = ["D9E2F3", "D9E2F3", "D9E2F3", "D9B38C", "D9B38C", "F4B183", "F4B183", "FFD966", "FFD966", "FFF2CC"]
    for column, (header, color) in enumerate(zip(headers, header_colors), start=1):
        cell = worksheet.cell(1, column, header)
        cell.font = Font(name="Calibri", size=10, bold=True, color="1F1F1F")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = PatternFill(fill_type="solid", fgColor=color)

    def find_source_column(source_headers: dict[str, int], *aliases: str) -> int | None:
        return next(
            (
                source_headers[_normalize_tariff_lookup_key(alias)]
                for alias in aliases
                if _normalize_tariff_lookup_key(alias) in source_headers
            ),
            None,
        )

    rows: list[tuple[Any, int, dict[str, int]]] = []
    for sheet_name in invoice_sheet_names:
        if sheet_name not in report_workbook.sheetnames:
            continue
        source = report_workbook[sheet_name]
        source_headers = {
            _normalize_tariff_lookup_key(cell.value): cell.column
            for cell in source[1]
            if cell.value is not None
        }
        etso_column = find_source_column(source_headers, "ETSO", "ETSO Kodu", "Etso Kodu")
        energy_column = find_source_column(source_headers, "Aktif Enerji Tüketim (kWh)")
        if etso_column is None or energy_column is None:
            continue
        status_column = find_source_column(source_headers, "Durum", "Status")
        for row_index in range(2, source.max_row + 1):
            code = reconciliation_etso_key(source.cell(row_index, etso_column).value)
            if not code:
                continue
            if status_column is not None and str(source.cell(row_index, status_column).value or "").strip().casefold() != "kesin":
                continue
            rows.append((source, row_index, source_headers))

    matched_rows = 0
    unmatched_rows = 0
    unique_matched_codes: set[str] = set()
    own_difference_rows: list[dict[str, Any]] = []
    thin_gray = Side(style="thin", color="D9E1F2")
    for output_row, (source, source_row, source_headers) in enumerate(rows, start=2):
        company_column = find_source_column(source_headers, "Dağıtım Şirketi", "Dağıtım Bölgesi")
        etso_column = find_source_column(source_headers, "ETSO", "ETSO Kodu", "Etso Kodu")
        customer_column = find_source_column(source_headers, "Müşteri", "Ünvan")
        tariff_column = find_source_column(source_headers, "Tarife Grubu")
        energy_column = find_source_column(source_headers, "Aktif Enerji Tüketim (kWh)")
        code = reconciliation_etso_key(source.cell(source_row, etso_column).value) if etso_column else ""
        for target_column, source_column in ((1, company_column), (2, etso_column), (3, customer_column), (4, tariff_column), (5, energy_column)):
            if source_column is None:
                worksheet.cell(output_row, target_column, "-")
                continue
            source_value = source.cell(source_row, source_column).value
            if isinstance(source_value, str) and source_value.startswith("="):
                reference = f"'{source.title.replace(chr(39), chr(39) * 2)}'!{get_column_letter(source_column)}{source_row}"
                worksheet.cell(output_row, target_column, f'=IF(OR({reference}="",{reference}="-"),"-",{reference})')
            else:
                worksheet.cell(output_row, target_column, source_value if source_value not in (None, "") else "-")
        worksheet.cell(output_row, 6, "Dağıtım verisi yok")
        reconciliation_value = consumption_by_etso.get(code) if code else None
        if reconciliation_value is None:
            worksheet.cell(output_row, 7, "-")
            worksheet.cell(output_row, 8, "-")
            worksheet.cell(output_row, 10, "Uzlaştırmada bulunamadı")
            unmatched_rows += 1
        else:
            worksheet.cell(output_row, 7, reconciliation_value)
            worksheet.cell(output_row, 8, f'=IF(ISNUMBER(E{output_row}),E{output_row}-G{output_row},"-")')
            worksheet.cell(output_row, 10, "Uzlaştırmada eşleşti")
            matched_rows += 1
            unique_matched_codes.add(code)
            own_energy_value = source.cell(source_row, energy_column).value if energy_column else None
            if (
                own_energy_value is not None
                and not (isinstance(own_energy_value, str) and own_energy_value.strip() in {"", "-", "–", "—"})
                and not (isinstance(own_energy_value, str) and own_energy_value.startswith("="))
            ):
                try:
                    own_number = safe_float(own_energy_value)
                    difference = own_number - reconciliation_value
                    if abs(difference) > 1e-9:
                        own_difference_rows.append({
                            "ETSO": source.cell(source_row, etso_column).value if etso_column else "-",
                            "Dağıtım Şirketi": source.cell(source_row, company_column).value if company_column else "-",
                            "Müşteri": source.cell(source_row, customer_column).value if customer_column else "-",
                            "Tarife Grubu": source.cell(source_row, tariff_column).value if tariff_column else "-",
                            "Bizim Kestiğimiz Tüketim (kWh)": own_number,
                            "Uzlaştırma Tüketim (kWh)": reconciliation_value,
                            "Fark (kWh)": difference,
                        })
                except (TypeError, ValueError):
                    pass
        worksheet.cell(output_row, 9, "-")
        for column in range(1, len(headers) + 1):
            cell = worksheet.cell(output_row, column)
            cell.alignment = Alignment(vertical="center")
            cell.border = Border(bottom=thin_gray)
            if column in {5, 7, 8, 9}:
                cell.number_format = '#,##0.00;[Red]-#,##0.00;0.00'

    worksheet.freeze_panes = "E2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, len(rows) + 1)}"
    worksheet.sheet_view.showGridLines = True
    worksheet.row_dimensions[1].height = 48
    for column, width in enumerate([28, 23, 42, 28, 25, 30, 24, 39, 39, 24], start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width
    return {
        "reconciliation_consumption_comparison_sheet_name": title,
        "reconciliation_consumption_matched_count": matched_rows,
        "reconciliation_consumption_unmatched_count": unmatched_rows,
        "reconciliation_consumption_matched_etso_count": len(unique_matched_codes),
        "our_vs_reconciliation_difference_rows": own_difference_rows,
        "distribution_vs_reconciliation_difference_rows": [],
    }


def build_reconciliation_consumption_report(
    rows: list[dict[str, Any]],
    report_kind: str,
) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    is_ours = report_kind == "ours"
    source_consumption_header = (
        "Bizim Kestiğimiz Tüketim (kWh)" if is_ours else "Dağıtım Tüketim (kWh)"
    )
    difference_header = (
        "Bizim Kestiğimiz - Uzlaştırma Farkı (kWh)"
        if is_ours
        else "Dağıtım - Uzlaştırma Farkı (kWh)"
    )
    headers = [
        "ETSO", "Dağıtım Şirketi", "Müşteri", "Tarife Grubu",
        source_consumption_header, "Uzlaştırma Tüketim (kWh)", difference_header,
    ]
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Fark Raporu"
    for column, header in enumerate(headers, start=1):
        cell = worksheet.cell(1, column, header)
        cell.font = Font(bold=True, color="1F1F1F")
        cell.fill = PatternFill(fill_type="solid", fgColor="FFD966")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_index, row in enumerate(rows, start=2):
        values = [
            row.get("ETSO", "-"), row.get("Dağıtım Şirketi", "-"),
            row.get("Müşteri", "-"), row.get("Tarife Grubu", "-"),
            row.get(source_consumption_header, "-"),
            row.get("Uzlaştırma Tüketim (kWh)", "-"),
            row.get("Fark (kWh)", "-"),
        ]
        for column, value in enumerate(values, start=1):
            cell = worksheet.cell(row_index, column, value if value not in (None, "") else "-")
            cell.alignment = Alignment(vertical="center")
            if column >= 5:
                cell.number_format = '#,##0.00;[Red]-#,##0.00;0.00'
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, len(rows) + 1)}"
    worksheet.row_dimensions[1].height = 36
    for column, width in enumerate([23, 28, 42, 28, 28, 25, 34], start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width
    buffer = io.BytesIO()
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


def render_reconciliation_difference_report_area(
    invoice_sources: list[dict[str, Any]],
) -> None:
    source = next(
        (
            item for item in invoice_sources
            if "our_vs_reconciliation_difference_rows" in item
            or "distribution_vs_reconciliation_difference_rows" in item
        ),
        None,
    )
    if source is None:
        return

    ours_rows = source.get("our_vs_reconciliation_difference_rows", [])
    distribution_rows = source.get("distribution_vs_reconciliation_difference_rows", [])
    st.markdown("#### Uzlaştırma Tüketim Fark Raporları")
    st.caption("Raporlar yalnızca tüketim farkı bulunan ETSO kayıtlarını içerir.")
    ours_column, distribution_column = st.columns(2, gap="medium")
    with ours_column:
        st.download_button(
            f"Bizim Kestiğimiz - Uzlaştırma ({format_turkish_integer(len(ours_rows))} ETSO)",
            data=build_reconciliation_consumption_report(ours_rows, "ours") if ours_rows else b"",
            file_name="Bizim Kestigimiz - Uzlastirma Tuketim Farklari.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_ours_reconciliation_consumption_report",
            width="stretch",
            disabled=not ours_rows,
            on_click="ignore",
        )
    with distribution_column:
        st.download_button(
            f"Dağıtım - Uzlaştırma ({format_turkish_integer(len(distribution_rows))} ETSO)",
            data=build_reconciliation_consumption_report(distribution_rows, "distribution")
            if distribution_rows else b"",
            file_name="Dagitim - Uzlastirma Tuketim Farklari.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_distribution_reconciliation_consumption_report",
            width="stretch",
            disabled=not distribution_rows,
            on_click="ignore",
        )
    if not ours_rows and not distribution_rows:
        st.info("Uzlaştırma tüketimiyle farkı bulunan ETSO kaydı yok.")


def add_invoice_distribution_comparison_worksheet(
    report_workbook: Any,
    comparison_sheet_names: list[str],
) -> dict[str, Any]:
    """Compare confirmed invoices with the distribution detail by ETSO code."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    detail_sheet = next(
        (
            sheet for sheet in report_workbook.worksheets
            if _normalize_tariff_lookup_key(sheet.title) == _normalize_tariff_lookup_key("Dağıtımın Kestiği")
        ),
        None,
    )
    if detail_sheet is None:
        return {"reconciliation_sheet_warning": "Dağıtımın Kestiği sayfası bulunamadığı için Karşılaştırma sayfası oluşturulamadı."}

    def get_header_columns(worksheet: Any) -> dict[str, int]:
        headers: dict[str, int] = {}
        for cell in worksheet[1]:
            if cell.value is None:
                continue
            # Prefer visible columns when hidden tariff lookup columns repeat a header.
            headers.setdefault(_normalize_tariff_lookup_key(cell.value), cell.column)
        return headers

    def find_column(headers: dict[str, int], *aliases: str) -> int | None:
        return next(
            (
                headers[_normalize_tariff_lookup_key(alias)]
                for alias in aliases
                if _normalize_tariff_lookup_key(alias) in headers
            ),
            None,
        )

    detail_headers = get_header_columns(detail_sheet)
    detail_aliases = {
        "company": ("Dağıtım Bölgesi", "Dağıtım Şirketi"),
        "etso": ("ETSO Kodu", "Etso Kodu", "ETSO", "Etso"),
        "customer": ("Müşteri", "Ünvan"),
        "abone": ("Abone Grubu", "Tarife Grubu"),
        "term": ("Terim Tipi", "TERİM", "Terim"),
        "connection": ("Bağlantı Tipi", "AG OG"),
        "power": ("Güç kW", "Güç (kW)", "Güç kWh"),
        "energy": ("Aktif Enerji Tüketim (kWh)",),
        "distribution_fee": ("Dağıtım Bedeli (TL)", "Dağıtım Bedeli(TL)"),
        "power_fee": ("Güç Bedeli (TL)", "Güç Bedeli(TL)"),
        "power_overrun_fee": ("Güç Aşım Bedeli (TL)",),
        "reactive_fee": ("Reaktif Bedel (TL)",),
    }
    detail_column_map = {
        key: find_column(detail_headers, *aliases)
        for key, aliases in detail_aliases.items()
    }

    invoice_aliases = {
        "company": ("Dağıtım Şirketi", "Dağıtım Bölgesi"),
        "etso": ("ETSO", "ETSO Kodu", "Etso"),
        "customer": ("Müşteri",),
        "tariff": ("Tarife Grubu",),
        "status": ("Durum", "Status"),
        "abone": ("Abone Grubu",),
        "term": ("Terim Tipi", "TERİM", "Terim"),
        "connection": ("Bağlantı Tipi", "AG OG"),
        "power": ("Güç kWh",),
        "energy": ("Aktif Enerji Tüketim (kWh)",),
        "distribution_fee": ("Dağıtım Bedeli (TL)",),
        "power_fee": ("Güç Bedeli(TL)", "Güç Bedeli (TL)"),
        "power_overrun_fee": ("Güç Aşım Bedeli (TL)",),
        "reactive_fee": ("Reaktif Bedel (TL)",),
    }

    detail_optional = {"abone", "power"}
    detail_missing = [
        key for key, column in detail_column_map.items()
        if column is None and key not in detail_optional
    ]
    if detail_missing:
        return {
            "reconciliation_sheet_warning": (
                "Karşılaştırma sayfası oluşturulamadı; Dağıtımın Kestiği sayfasında eksik başlıklar: "
                + ", ".join(detail_missing)
            )
        }

    def etso_key(value: Any) -> str:
        return reconciliation_etso_key(value)

    def customer_key(value: Any) -> str:
        return _normalize_tariff_lookup_key(value) if value is not None else ""

    distribution_rows: dict[str, list[tuple[int, str]]] = {}
    for row_index in range(2, detail_sheet.max_row + 1):
        code = etso_key(detail_sheet.cell(row_index, detail_column_map["etso"]).value)
        if not code:
            continue
        customer = detail_sheet.cell(row_index, detail_column_map["customer"]).value
        distribution_rows.setdefault(code, []).append((row_index, customer_key(customer)))

    comparisons: list[tuple[Any, int, dict[str, int], int, int | None]] = []
    definite_invoice_count = 0
    for sheet_name in comparison_sheet_names:
        if sheet_name not in report_workbook.sheetnames:
            continue
        invoice_sheet = report_workbook[sheet_name]
        invoice_headers = get_header_columns(invoice_sheet)
        invoice_column_map = {
            key: find_column(invoice_headers, *aliases)
            for key, aliases in invoice_aliases.items()
        }
        invoice_optional = {"status", "abone", "power"}
        missing = [
            key for key, column in invoice_column_map.items()
            if column is None and key not in invoice_optional
        ]
        if missing:
            return {
                "reconciliation_sheet_warning": (
                    f"Karşılaştırma sayfası oluşturulamadı; {sheet_name} sayfasında eksik başlıklar: "
                    + ", ".join(missing)
                )
            }

        for row_index in range(2, invoice_sheet.max_row + 1):
            code = etso_key(invoice_sheet.cell(row_index, invoice_column_map["etso"]).value)
            if not code:
                continue
            status_column = invoice_column_map.get("status")
            if status_column is not None:
                status = invoice_sheet.cell(row_index, status_column).value
                if str(status or "").strip().casefold() != "kesin":
                    continue
            definite_invoice_count += 1
            customer = customer_key(invoice_sheet.cell(row_index, invoice_column_map["customer"]).value)
            matches = distribution_rows.get(code, []) if code else []
            distribution_row: int | None = None
            match_status = 0
            if len(matches) == 1:
                distribution_row = matches[0][0]
                match_status = 1
            elif len(matches) > 1:
                customer_matches = [row for row, name in matches if customer and name == customer]
                if len(customer_matches) == 1:
                    distribution_row = customer_matches[0]
                    match_status = 1
                else:
                    match_status = 2
            comparisons.append((invoice_sheet, row_index, invoice_column_map, match_status, distribution_row))

    isk_sanayi_key = _normalize_tariff_lookup_key("\u0130SK Sanayi")
    comparisons.sort(
        key=lambda comparison: _normalize_tariff_lookup_key(
            comparison[0].cell(
                comparison[1],
                comparison[2]["tariff"],
            ).value
        ) == isk_sanayi_key
    )

    title = "Karşılaştırma"
    suffix = 2
    while title in report_workbook.sheetnames:
        suffix_text = f" {suffix}"
        title = f"Karşılaştırma"[:31 - len(suffix_text)] + suffix_text
        suffix += 1
    worksheet = report_workbook.create_sheet(title)

    headers = ["Dağıtım Şirketi", "ETSO", "Müşteri", "Tarife Grubu"]
    header_roles = ["identity", "identity", "identity", "ours"]
    specs = [
        ("Terim Tipi", "Terim Tipi Dağıtım", None, "term", "term"),
        ("Bağlantı Tipi", "Bağlantı Tipi Dağıtım", None, "connection", "connection"),
        ("Aktif Enerji Tüketim (kWh)", "Aktif Enerji Tüketim (kWh) Dağıtım", None, "energy", "energy"),
        ("Dağıtım Bedeli (TL)", "Dağıtım Bedeli (TL) Dağıtım", "Dağıtım Bedeli Fark", "distribution_fee", "distribution_fee"),
        ("Güç Bedeli (TL)", "Güç Bedeli (TL) Dağıtım", "Güç Bedeli Fark", "power_fee", "power_fee"),
        ("Güç Aşım Bedeli (TL)", "Güç Aşım Bedeli (TL) Dağıtım", "Güç Aşım Bedeli Fark", "power_overrun_fee", "power_overrun_fee"),
        ("Reaktif Bedel (TL)", "Reaktif Bedel (TL) Dağıtım", "Reaktif Bedel Fark", "reactive_fee", "reactive_fee"),
    ]
    if all(
        invoice_column_map.get("abone") is not None
        and detail_column_map.get("abone") is not None
        for _, _, invoice_column_map, _, _ in comparisons
    ):
        specs.insert(0, ("Abone Grubu", "Abone Grubu Dağıtım", None, "abone", "abone"))
    if all(
        invoice_column_map.get("power") is not None
        and detail_column_map.get("power") is not None
        for _, _, invoice_column_map, _, _ in comparisons
    ):
        specs.insert(3 if specs and specs[0][3] == "abone" else 2, ("Güç kWh", "Güç kWh Dağıtım", None, "power", "power"))
    value_columns: dict[str, tuple[int, int, int | None]] = {}
    numeric_columns: set[int] = set()
    for own_header, distribution_header, difference_header, own_key, distribution_key in specs:
        own_column = len(headers) + 1
        headers.append(own_header)
        header_roles.append("ours")
        distribution_column = len(headers) + 1
        headers.append(distribution_header)
        header_roles.append("distribution")
        difference_column = None
        if difference_header:
            difference_column = len(headers) + 1
            headers.append(difference_header)
            header_roles.append("difference")
            numeric_columns.add(difference_column)
        if own_key in {"power", "energy", "distribution_fee", "power_fee", "power_overrun_fee", "reactive_fee"}:
            numeric_columns.update((own_column, distribution_column))
        value_columns[own_key] = (own_column, distribution_column, difference_column)
    headers.append("ETSO Eşleşmesi")
    header_roles.append("match")

    for column, value in enumerate(headers, start=1):
        cell = worksheet.cell(1, column, value)
        cell.font = Font(name="Calibri", size=10, bold=True, color="1F1F1F")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        role = header_roles[column - 1]
        fill_color = {
            "identity": "D9E2F3",
            "ours": "D9B38C",
            "distribution": "F4B183",
            "difference": "FFD966",
            "match": "FFF2CC",
        }[role]
        cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)

    thin_gray = Side(style="thin", color="D9E1F2")
    body_border = Border(bottom=thin_gray)
    body_alignment = Alignment(vertical="center")
    body_fills = {
        "ours": PatternFill(fill_type="solid", fgColor="F2E6D9"),
        "difference": PatternFill(fill_type="solid", fgColor="FFF2CC"),
    }
    mismatch_fill = PatternFill(fill_type="solid", fgColor="FFFF00")
    subscription_group_mismatches: list[dict[str, Any]] = []

    def compare_source_values(left: Any, right: Any, key: str) -> bool:
        if (
            isinstance(left, str) and left.startswith("=")
        ) or (
            isinstance(right, str) and right.startswith("=")
        ):
            return False
        blank_values = {"", "-", "–", "—"}
        left_blank = left is None or (isinstance(left, str) and left.strip() in blank_values)
        right_blank = right is None or (isinstance(right, str) and right.strip() in blank_values)
        if left_blank or right_blank:
            return left_blank != right_blank
        if key == "power":
            try:
                return abs(safe_float(left) - safe_float(right)) > 1e-9
            except (TypeError, ValueError):
                pass
        if key == "abone":
            left = normalize_invoice_subscription_group(left)
            right = normalize_invoice_subscription_group(right)
        if key == "term":
            term_aliases = {
                "tekterimli": "tekterim",
                "tekterim": "tekterim",
                "ciftterimli": "ciftterim",
                "ciftterim": "ciftterim",
            }
            left_key = _normalize_tariff_lookup_key(left)
            right_key = _normalize_tariff_lookup_key(right)
            return term_aliases.get(left_key, left_key) != term_aliases.get(right_key, right_key)
        return _normalize_tariff_lookup_key(left) != _normalize_tariff_lookup_key(right)

    def write_source_value(
        target_cell: Any,
        source_sheet: Any,
        source_row: int,
        source_column: int,
        value_transform: Any | None = None,
        blank_value: Any = "-",
    ) -> None:
        value = source_sheet.cell(source_row, source_column).value
        if value_transform is not None:
            value = value_transform(value)
        if isinstance(value, str) and value.startswith("="):
            quoted_sheet = "'" + source_sheet.title.replace("'", "''") + "'"
            source_ref = f"{quoted_sheet}!{get_column_letter(source_column)}{source_row}"
            formula_blank_value = (
                f'"{blank_value.replace(chr(34), chr(34) * 2)}"'
                if isinstance(blank_value, str)
                else str(blank_value)
            )
            target_cell.value = f'=IF(OR({source_ref}="",{source_ref}="-"),{formula_blank_value},{source_ref})'
        elif (
            value is None
            or (isinstance(value, str) and not value.strip())
            or (blank_value == 0 and isinstance(value, str) and value.strip() in {"-", "–", "—"})
        ):
            target_cell.value = blank_value
        else:
            target_cell.value = value

    for output_row, (invoice_sheet, invoice_row, invoice_columns, match_status, distribution_row) in enumerate(comparisons, start=2):
        mismatched_columns: set[int] = set()
        invoice_company = invoice_sheet.cell(invoice_row, invoice_columns["company"]).value
        invoice_company_present = (
            invoice_company is not None
            and (not isinstance(invoice_company, str) or invoice_company.strip() not in {"", "-", "–", "—"})
        )
        detail_company = (
            detail_sheet.cell(distribution_row, detail_column_map["company"]).value
            if distribution_row is not None
            else None
        )
        detail_company_present = (
            detail_company is not None
            and (not isinstance(detail_company, str) or detail_company.strip() not in {"", "-", "–", "—"})
        )
        if invoice_company_present:
            write_source_value(
                worksheet.cell(output_row, 1),
                invoice_sheet,
                invoice_row,
                invoice_columns["company"],
            )
        elif detail_company_present and distribution_row is not None:
            write_source_value(
                worksheet.cell(output_row, 1),
                detail_sheet,
                distribution_row,
                detail_column_map["company"],
            )
        else:
            worksheet.cell(output_row, 1, "Bilinmeyen Dağıtım Şirketi")
        write_source_value(worksheet.cell(output_row, 2), invoice_sheet, invoice_row, invoice_columns["etso"])
        write_source_value(worksheet.cell(output_row, 3), invoice_sheet, invoice_row, invoice_columns["customer"])
        write_source_value(worksheet.cell(output_row, 4), invoice_sheet, invoice_row, invoice_columns["tariff"])

        for _, _, _, own_key, distribution_key in specs:
            own_column, distribution_column, difference_column = value_columns[own_key]
            write_source_value(
                worksheet.cell(output_row, own_column),
                invoice_sheet,
                invoice_row,
                invoice_columns[own_key],
                normalize_invoice_subscription_group if own_key == "abone" else None,
                0 if difference_column is not None else "-",
            )
            if distribution_row is not None:
                write_source_value(
                    worksheet.cell(output_row, distribution_column),
                    detail_sheet,
                    distribution_row,
                    detail_column_map[distribution_key],
                    blank_value=0 if difference_column is not None else "-",
                )
            elif difference_column is not None:
                worksheet.cell(output_row, distribution_column, 0)
            if difference_column is not None:
                own_letter = get_column_letter(own_column)
                distribution_letter = get_column_letter(distribution_column)
                worksheet.cell(
                    output_row,
                    difference_column,
                    f'={distribution_letter}{output_row}-{own_letter}{output_row}',
                )

            if (
                distribution_row is not None
                and own_key in {"abone", "term", "connection", "power"}
                and compare_source_values(
                    invoice_sheet.cell(invoice_row, invoice_columns[own_key]).value,
                    detail_sheet.cell(distribution_row, detail_column_map[distribution_key]).value,
                    own_key,
                )
            ):
                mismatched_columns.update((own_column, distribution_column))

        if (
            distribution_row is not None
            and invoice_columns.get("abone") is not None
            and detail_column_map.get("abone") is not None
        ):
            ours_group = normalize_invoice_subscription_group(
                invoice_sheet.cell(invoice_row, invoice_columns["abone"]).value
            )
            distribution_group = normalize_invoice_subscription_group(
                detail_sheet.cell(distribution_row, detail_column_map["abone"]).value
            )
            ours_group_key = _normalize_tariff_lookup_key(ours_group)
            distribution_group_key = _normalize_tariff_lookup_key(distribution_group)
            if ours_group_key and distribution_group_key and ours_group_key != distribution_group_key:
                subscription_group_mismatches.append(
                    {
                        "ETSO": invoice_sheet.cell(invoice_row, invoice_columns["etso"]).value or "-",
                        "Dağıtım Şirketi": (
                            invoice_company if invoice_company_present
                            else detail_company if detail_company_present
                            else "Bilinmeyen Dağıtım Şirketi"
                        ),
                        "Bizim Kestiğimiz Abone Grubu": ours_group,
                        "Dağıtımın Kestiği Abone Grubu": distribution_group,
                    }
                )

        worksheet.cell(
            output_row,
            len(headers),
            "Eşleşti" if match_status == 1 else "Birden fazla kayıt" if match_status == 2 else "Bulunamadı",
        )
        for column in range(1, len(headers) + 1):
            cell = worksheet.cell(output_row, column)
            cell.alignment = body_alignment
            cell.border = body_border
            body_fill = None if column == 4 else body_fills.get(header_roles[column - 1])
            if body_fill is not None:
                cell.fill = body_fill
            if cell.value is None or (isinstance(cell.value, str) and not cell.value.strip()):
                cell.value = "-"
            if column in numeric_columns:
                cell.number_format = '#,##0.00;[Red]-#,##0.00;0.00'
        for column in mismatched_columns:
            worksheet.cell(output_row, column).fill = mismatch_fill

    worksheet.freeze_panes = "E2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, len(comparisons) + 1)}"
    worksheet.sheet_view.showGridLines = True
    worksheet.row_dimensions[1].height = 38
    widths = [28, 23, 42, 30]
    for _, _, difference_header, _, _ in specs:
        widths.extend([23, 25])
        if difference_header:
            widths.append(21)
    widths.append(22)
    for column, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width

    report_workbook.calculation.calcMode = "auto"
    report_workbook.calculation.fullCalcOnLoad = True
    report_workbook.calculation.forceFullCalc = True
    report_workbook.calculation.calcOnSave = True
    matched_count = sum(1 for comparison in comparisons if comparison[3] == 1)
    ambiguous_count = sum(1 for comparison in comparisons if comparison[3] == 2)
    return {
        "reconciliation_sheet_name": title,
        "reconciliation_invoice_count": definite_invoice_count,
        "reconciliation_matched_count": matched_count,
        "reconciliation_ambiguous_count": ambiguous_count,
        "reconciliation_unmatched_count": definite_invoice_count - matched_count - ambiguous_count,
        "subscription_group_mismatches": subscription_group_mismatches,
    }


def build_company_comparison_reports(
    output_bytes: bytes | None,
    selected_companies: Iterable[str] | None = None,
    include_group_mismatches: bool = True,
) -> tuple[dict[str, bytes], list[dict[str, Any]], str | None]:
    """Create one standalone, value-based comparison workbook per distribution company."""
    if not output_bytes:
        return {}, [], None

    from copy import copy
    from openpyxl import Workbook, load_workbook
    from openpyxl.utils import get_column_letter, column_index_from_string

    source_workbook = load_workbook(io.BytesIO(output_bytes), data_only=False)
    try:
        selected_company_keys = (
            {_normalize_tariff_lookup_key(company) for company in selected_companies}
            if selected_companies is not None
            else set()
        )
        comparison_sheet = next(
            (
                sheet for sheet in source_workbook.worksheets
                if _normalize_tariff_lookup_key(sheet.title).startswith(
                    _normalize_tariff_lookup_key("Karşılaştırma")
                )
            ),
            None,
        )
        if comparison_sheet is None:
            return {}, [], None

        header_columns: dict[str, int] = {}
        for cell in comparison_sheet[1]:
            if cell.value is not None:
                header_columns.setdefault(_normalize_tariff_lookup_key(cell.value), cell.column)

        def find_header(*names: str) -> int | None:
            return next(
                (
                    header_columns[_normalize_tariff_lookup_key(name)]
                    for name in names
                    if _normalize_tariff_lookup_key(name) in header_columns
                ),
                None,
            )

        company_column = find_header("Dağıtım Şirketi", "Dağıtım Bölgesi")
        etso_column = find_header("ETSO", "Etso")
        ours_group_column = find_header("Abone Grubu")
        distribution_group_column = find_header("Abone Grubu Dağıtım")
        if not all((company_column, etso_column, ours_group_column, distribution_group_column)):
            return {}, [], "Karşılaştırma sayfasında şirket, ETSO veya abone grubu sütunu bulunamadı."

        resolved_values: dict[tuple[str, int, int], Any] = {}
        resolving: set[tuple[str, int, int]] = set()

        def numeric_value(value: Any) -> float | None:
            if value is None or (isinstance(value, str) and value.strip() in {"", "-", "–", "—"}):
                return None
            try:
                return safe_float(value)
            except (TypeError, ValueError):
                return None

        def resolve_reference(reference: str, current_sheet: Any, depth: int) -> Any:
            reference_match = re.fullmatch(
                r"(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?(\d+)",
                reference.strip(),
            )
            if reference_match:
                quoted_name, plain_name, column_letters, row_text = reference_match.groups()
                sheet_name = quoted_name.replace("''", "'") if quoted_name else plain_name.strip()
                target_sheet = source_workbook[sheet_name]
                target_cell = target_sheet.cell(int(row_text), column_index_from_string(column_letters))
                return resolve_cell(target_sheet, target_cell.row, target_cell.column, depth + 1)

            local_match = re.fullmatch(r"\$?([A-Z]{1,3})\$?(\d+)", reference.strip())
            if local_match:
                column_letters, row_text = local_match.groups()
                target_cell = current_sheet.cell(int(row_text), column_index_from_string(column_letters))
                return resolve_cell(current_sheet, target_cell.row, target_cell.column, depth + 1)
            return None

        power_rate_lookups: dict[str, dict[str, float | None]] = {}

        def evaluate_power_formula(worksheet: Any, row: int) -> float:
            tariff_key = _normalize_tariff_lookup_key(resolve_cell(worksheet, row, 4, 1))
            fee = numeric_value(resolve_cell(worksheet, row, 11, 1))
            rate_lookup = power_rate_lookups.get(worksheet.title)
            if rate_lookup is None:
                rate_lookup = {}
                empty_helper_rows = 0
                for helper_row in range(2, worksheet.max_row + 1):
                    helper_tariff = worksheet.cell(helper_row, 21).value
                    if helper_tariff is None or not str(helper_tariff).strip():
                        empty_helper_rows += 1
                        if empty_helper_rows >= 2:
                            break
                        continue
                    empty_helper_rows = 0
                    rate_lookup[_normalize_tariff_lookup_key(helper_tariff)] = numeric_value(
                        worksheet.cell(helper_row, 24).value
                    )
                power_rate_lookups[worksheet.title] = rate_lookup
            rate = rate_lookup.get(tariff_key)
            if fee is None or rate in (None, 0):
                return 0.0
            return fee / (rate / 100)

        def resolve_cell(worksheet: Any, row: int, column: int, depth: int = 0) -> Any:
            key = (worksheet.title, row, column)
            if key in resolved_values:
                return resolved_values[key]
            if depth > 20 or key in resolving:
                return None
            resolving.add(key)
            value = worksheet.cell(row, column).value
            if not (isinstance(value, str) and value.startswith("=")):
                resolved = value
            else:
                resolved = None
                formula = value.strip()
                wrapped_reference = re.fullmatch(
                    r'=IF\(OR\((.+)="",\1="-"\),(.+),\1\)',
                    formula,
                )
                difference = re.fullmatch(
                    r"=\$?([A-Z]{1,3})\$?(\d+)-\$?([A-Z]{1,3})\$?(\d+)",
                    formula,
                )
                if wrapped_reference:
                    source_value = resolve_reference(wrapped_reference.group(1), worksheet, depth)
                    fallback = wrapped_reference.group(2).strip()
                    if fallback.startswith('"') and fallback.endswith('"'):
                        fallback_value: Any = fallback[1:-1].replace('""', '"')
                    else:
                        try:
                            fallback_value = float(fallback)
                        except ValueError:
                            fallback_value = fallback
                    resolved = (
                        fallback_value
                        if source_value is None
                        or (isinstance(source_value, str) and source_value.strip() in {"", "-", "–", "—"})
                        else source_value
                    )
                elif difference:
                    left_column, left_row, right_column, right_row = difference.groups()
                    left_value = numeric_value(
                        resolve_cell(worksheet, int(left_row), column_index_from_string(left_column), depth + 1)
                    ) or 0.0
                    right_value = numeric_value(
                        resolve_cell(worksheet, int(right_row), column_index_from_string(right_column), depth + 1)
                    ) or 0.0
                    resolved = left_value - right_value
                elif (
                    column == 8
                    and "VLOOKUP" in formula.upper()
                    and _normalize_tariff_lookup_key(worksheet.title).startswith(
                        _normalize_tariff_lookup_key("Bizim Kestiğimiz 2")
                    )
                ):
                    resolved = evaluate_power_formula(worksheet, row)
                else:
                    direct_reference = re.fullmatch(r"=(.+)", formula)
                    if direct_reference:
                        resolved = resolve_reference(direct_reference.group(1), worksheet, depth)
                if resolved is None:
                    resolved = "-"
            resolving.discard(key)
            resolved_values[key] = resolved
            return resolved

        headers = [comparison_sheet.cell(1, column).value for column in range(1, comparison_sheet.max_column + 1)]
        source_difference_pattern = re.compile(r"=\$?([A-Z]{1,3})\$?\d+-\$?([A-Z]{1,3})\$?\d+")
        company_rows: dict[str, list[tuple[int, list[Any]]]] = {}
        group_mismatches: list[dict[str, Any]] = []
        build_company_reports = selected_companies is None or bool(selected_company_keys)
        for row_index in range(2, comparison_sheet.max_row + 1):
            company_value = resolve_cell(comparison_sheet, row_index, company_column)
            company_name = str(company_value).strip() if company_value is not None else ""
            row_is_selected = (
                not selected_company_keys
                or _normalize_tariff_lookup_key(company_name) in selected_company_keys
            )
            if include_group_mismatches:
                ours_group = resolve_cell(comparison_sheet, row_index, ours_group_column)
                distribution_group = resolve_cell(comparison_sheet, row_index, distribution_group_column)
                ours_key = _normalize_tariff_lookup_key(normalize_invoice_subscription_group(ours_group))
                distribution_key = _normalize_tariff_lookup_key(
                    normalize_invoice_subscription_group(distribution_group)
                )
                if ours_key and distribution_key and ours_key != distribution_key:
                    group_mismatches.append(
                        {
                            "ETSO": resolve_cell(comparison_sheet, row_index, etso_column),
                            "Dağıtım Şirketi": company_name or "-",
                            "Bizim Kestiğimiz Abone Grubu": ours_group,
                            "Dağıtımın Kestiği Abone Grubu": distribution_group,
                        }
                    )
            if build_company_reports and row_is_selected and company_name and company_name != "-":
                row_values = [
                    resolve_cell(comparison_sheet, row_index, column)
                    for column in range(1, comparison_sheet.max_column + 1)
                ]
                company_rows.setdefault(company_name, []).append((row_index, row_values))

        reports: dict[str, bytes] = {}
        for company_name, company_data in sorted(company_rows.items(), key=lambda item: item[0].casefold()):
            company_workbook = Workbook()
            company_sheet = company_workbook.active
            company_sheet.title = "Karşılaştırma"
            for target_row, row_values in enumerate(
                [
                    (1, headers),
                    *company_data,
                ],
                start=1,
            ):
                source_row, values = row_values
                for column, cell_value in enumerate(values, start=1):
                    target_cell = company_sheet.cell(target_row, column, cell_value)
                    source_cell = comparison_sheet.cell(source_row, column)
                    if isinstance(source_cell.value, str):
                        difference_match = source_difference_pattern.fullmatch(source_cell.value)
                        if difference_match:
                            left_column, right_column = difference_match.groups()
                            target_cell.value = f"={left_column}{target_row}-{right_column}{target_row}"
                    if source_cell.has_style:
                        target_cell.number_format = source_cell.number_format
                        target_cell.font = copy(source_cell.font)
                        target_cell.fill = copy(source_cell.fill)
                        target_cell.border = copy(source_cell.border)
                        target_cell.alignment = copy(source_cell.alignment)
                        target_cell.protection = copy(source_cell.protection)
                source_height = comparison_sheet.row_dimensions[source_row].height
                if source_height is not None:
                    company_sheet.row_dimensions[target_row].height = source_height

            for column in range(1, comparison_sheet.max_column + 1):
                column_letter = get_column_letter(column)
                source_dimension = comparison_sheet.column_dimensions[column_letter]
                if source_dimension.width is not None:
                    company_sheet.column_dimensions[column_letter].width = source_dimension.width
            company_sheet.freeze_panes = comparison_sheet.freeze_panes
            company_sheet.sheet_view.showGridLines = comparison_sheet.sheet_view.showGridLines
            company_sheet.auto_filter.ref = (
                f"A1:{get_column_letter(comparison_sheet.max_column)}{company_sheet.max_row}"
            )
            report_buffer = io.BytesIO()
            company_workbook.calculation.fullCalcOnLoad = True
            company_workbook.calculation.forceFullCalc = True
            company_workbook.calculation.calcMode = "auto"
            company_workbook.save(report_buffer)
            company_workbook.close()
            reports[company_name] = report_buffer.getvalue()

        return reports, group_mismatches, None
    finally:
        source_workbook.close()


def total_distribution_amount(records: Iterable[dict[str, Any]]) -> float:
    return sum(safe_float(record.get("Dağıtım Bedeli(TL)", 0)) for record in records)


def build_processing_logs(
    records: list[dict[str, Any]],
    stats: dict[str, Any],
    audit_results: dict[str, Any],
) -> list[str]:
    read_count = int(stats.get("read_region_count", 0))
    expected_count = int(stats.get("expected_region_count", EXPECTED_REGION_COUNT))
    missing_regions = stats.get("missing_regions", [])

    if not missing_regions and read_count == expected_count:
        region_log = (
            f"Bölge doğrulaması: {format_turkish_integer(expected_count)} bölgenin "
            f"{format_turkish_integer(read_count)}'i başarıyla okundu."
        )
    else:
        region_log = "Şu bölgeler okunamadı: " + ", ".join(missing_regions)

    logs = [
        region_log,
        f"Taranan dosya sayısı: {format_turkish_integer(stats.get('discovered_files', 0))}",
        f"İşlenen dosya sayısı: {format_turkish_integer(stats.get('processed_files', 0))}",
        f"Ham çıkarılan satır: {format_turkish_integer(stats.get('total_extracted', 0))}",
        f"Birleştirme sonrası kayıt: {format_turkish_integer(len(records))}",
        f"Toplam dağıtım bedeli: {format_turkish_currency(total_distribution_amount(records))}",
        f"10.000.000,00 TL ve üzeri anomali: {format_turkish_integer(audit_results.get('high_value_count', 0))}",
    ]
    if stats.get("file_errors"):
        logs.append(f"Dosya okuma hatası: {format_turkish_integer(len(stats['file_errors']))}")
    mapping_warning_count = (
        len(stats.get("mapping_warnings", []))
        + len(stats.get("source_header_warnings", []))
    )
    if mapping_warning_count:
        logs.append(
            f"Mapping kalite uyarısı: {format_turkish_integer(mapping_warning_count)}"
        )
    if stats.get("mapping_features", {}).get("extended_financial_fields"):
        logs.extend([
            f"Standar dışı tutar toplamı: {format_turkish_currency(stats.get('standart_disi_total', 0))}",
            f"Tazminat bedeli toplamı: {format_turkish_currency(stats.get('tazminat_total', 0))}",
            f"Düzeltme eşleşmesi: {format_turkish_integer(stats.get('correction_total', 0))}",
        ])
        correction_examples = stats.get("correction_examples", [])
        selected_examples = []
        for correction_field in ["Tarife Grubu", "AG OG", "TERİM"]:
            field_example = next(
                (
                    example for example in correction_examples
                    if example.get("field") == correction_field
                ),
                None,
            )
            if field_example:
                selected_examples.append(field_example)
        for example in correction_examples:
            if len(selected_examples) >= 6:
                break
            if example not in selected_examples:
                selected_examples.append(example)

        for example in selected_examples:
            logs.append(
                "Düzeltme örneği "
                f"({example.get('region', '-')}, {example.get('field', '-')}): "
                f"{example.get('source', '')} → {example.get('target', '')}"
            )
    return logs


def run_analysis_pipeline(
    root_folder: str,
    mapping_path: str,
    status_callback: Any,
    invoice_file: str | None = None,
    system_usage_file: str | None = None,
    eno_file_path: str | None = None,
    reconciliation_consumption_file: str | None = None,
) -> tuple[pd.DataFrame, dict[str, Any], dict[str, Any], bytes, list[str], list[dict[str, Any]]]:
    """Doğrulanmış çekirdeği çalıştır ve maskesiz Excel'i bellekte oluştur."""
    if not root_folder and invoice_file:
        if status_callback:
            status_callback("Yaln\u0131zca Bizim Kesti\u011fimiz Faturalar verisi haz\u0131rlan\u0131yor.", 10)
        from openpyxl import Workbook, load_workbook

        base_workbook = Workbook()
        info_sheet = base_workbook.active
        info_sheet.title = "Bilgi"
        info_sheet["A1"] = (
            "Bu dosya, da\u011f\u0131t\u0131m faturas\u0131 dosyalar\u0131 y\u00fcklenmeden "
            "yaln\u0131zca Bizim Kesti\u011fimiz Faturalar verisiyle olu\u015fturuldu."
        )
        base_buffer = io.BytesIO()
        base_workbook.save(base_buffer)
        base_workbook.close()
        output_bytes, invoice_sources, invoice_file_errors, invoice_file_count = append_selected_invoice_file(
            base_buffer.getvalue(),
            invoice_file,
            mapping_path=mapping_path,
            system_usage_file=system_usage_file,
            reconciliation_consumption_file=reconciliation_consumption_file,
        )
        if invoice_file_count != 1 or not invoice_sources:
            if not invoice_file_errors:
                invoice_file_errors.append(
                    "Se\u00e7ilen fatura dosyas\u0131ndan Bizim Kesti\u011fimiz 2 sayfas\u0131 olu\u015fturulamad\u0131."
                )

        preview_records: list[dict[str, Any]] = []
        if invoice_sources and not invoice_file_errors:
            preview_workbook = load_workbook(io.BytesIO(output_bytes), read_only=True, data_only=True)
            try:
                for source in invoice_sources:
                    sheet_name = source.get("comparison_sheet_name")
                    if not sheet_name or sheet_name not in preview_workbook.sheetnames:
                        continue
                    worksheet = preview_workbook[sheet_name]
                    rows = worksheet.iter_rows(values_only=True)
                    headers = next(rows, ())
                    preview_columns: list[tuple[str, int]] = []
                    for column_index, header in enumerate(headers):
                        # The workbook's first blank heading separates its visible
                        # comparison fields from the tariff lookup helper table.
                        if is_blank(header):
                            break
                        preview_columns.append((str(header).strip(), column_index))
                    for values in rows:
                        if any(value is not None for value in values):
                            preview_records.append({
                                header: values[column_index]
                                for header, column_index in preview_columns
                                if column_index < len(values)
                            })
            finally:
                preview_workbook.close()
        preview_df = pd.DataFrame(preview_records)
        stats = {
            "expected_region_count": 0,
            "read_region_count": 0,
            "missing_regions": [],
            "invoice_file_count": invoice_file_count,
            "invoice_sheet_count": len(invoice_sources),
            "invoice_file_selected": True,
            "invoice_file_errors": invoice_file_errors,
            "file_errors": list(invoice_file_errors),
            "invoice_only_mode": True,
        }
        region_source_rows = sum(
            int(source.get("region_invoice_source_row_count", 0))
            for source in invoice_sources
        )
        region_etso_count = sum(
            int(source.get("region_invoice_etso_count", 0))
            for source in invoice_sources
        )
        duplicate_etso_rows = sum(
            int(source.get("region_invoice_duplicate_row_count", 0))
            for source in invoice_sources
        )
        logs = [
            "Da\u011f\u0131t\u0131m faturas\u0131 dosyalar\u0131 olmadan yaln\u0131zca Bizim Kesti\u011fimiz Faturalar i\u015flendi."
        ]
        if region_source_rows:
            logs.append(
                f"ETSO bazl\u0131 {format_turkish_integer(region_etso_count)} kay\u0131t haz\u0131rland\u0131; "
                f"{format_turkish_integer(duplicate_etso_rows)} m\u00fckerrer kaynak sat\u0131r\u0131 birle\u015ftirildi. "
                "S\u00f6zle\u015fme ve kurulu g\u00fc\u00e7 i\u00e7in maksimum, t\u00fcketim ve bedeller i\u00e7in toplam al\u0131nd\u0131."
            )
        if invoice_file_errors:
            logs.extend(f"{OUR_INVOICES_LABEL}: {error}" for error in invoice_file_errors)
        for source in invoice_sources:
            if source.get("reconciliation_consumption_sheet_name") or source.get("reconciliation_consumption_comparison_sheet_name"):
                logs.append(
                    "Uzlaştırma tüketimi: "
                    f"{format_turkish_integer(source.get('reconciliation_consumption_source_row_count', 0))} uygun sayaç satırı, "
                    f"{format_turkish_integer(source.get('reconciliation_consumption_etso_count', 0))} benzersiz ETSO, "
                    f"{format_turkish_integer(source.get('reconciliation_consumption_matched_count', 0))} eşleşen karşılaştırma satırı, "
                    f"{format_turkish_integer(source.get('reconciliation_consumption_unmatched_count', 0))} eşleşmeyen karşılaştırma satırı."
                )
            if source.get("reconciliation_consumption_warning"):
                logs.append(f"Uzlaştırma Tüketim: {source['reconciliation_consumption_warning']}")
        if status_callback:
            status_callback("Bizim Kesti\u011fimiz 2 sayfas\u0131 haz\u0131r.", 100)
        return (
            preview_df,
            stats,
            {"health_score": 100},
            output_bytes,
            logs,
            invoice_sources,
        )

    captured_core_output = io.StringIO()
    with contextlib.redirect_stdout(captured_core_output):
        extracted_records, stats = process_all_regions(
            base_path=root_folder,
            mapping_file_path=mapping_path,
            status_callback=status_callback,
            return_stats=True,
        )
        normalized_records = normalize_extracted_regions(extracted_records)
        # ENO zenginleştirme: Sayaç ID → EIC ETSO dönüşümü ve boş Müşteri doldurma
        eno_stats = enrich_from_eno_file(
            normalized_records,
            root_folder,
            eno_file_path=eno_file_path,
        )
        output_buffer = io.BytesIO()
        save_to_excel(normalized_records, output_buffer)
        output_bytes = add_summary_worksheet(output_buffer.getvalue(), normalized_records)
        output_bytes, invoice_sources, invoice_file_errors, invoice_file_count = append_selected_invoice_file(
            output_bytes,
            invoice_file,
            mapping_path=mapping_path,
            system_usage_file=system_usage_file,
            reconciliation_consumption_file=reconciliation_consumption_file,
        )
        if invoice_file and (invoice_file_count != 1 or not invoice_sources) and not invoice_file_errors:
            invoice_file_errors.append(
                "Seçilen fatura dosyasından çıktı sayfaları oluşturulamadı. Dosyanın çalışma sayfalarını kontrol edin."
            )
        stats["invoice_file_count"] = invoice_file_count
        stats["invoice_sheet_count"] = len(invoice_sources)
        stats["invoice_file_selected"] = bool(invoice_file)
        stats["invoice_file_errors"] = invoice_file_errors
        if invoice_file_errors:
            stats.setdefault("file_errors", []).extend(invoice_file_errors)

    audit_results = audit_extracted_dataset(normalized_records)
    raw_columns = list(RAW_DATA_COLUMNS)
    for optional_column in OPTIONAL_RAW_DATA_COLUMNS:
        if any(optional_column in record for record in normalized_records):
            raw_columns.append(optional_column)
    raw_df = pd.DataFrame(normalized_records).reindex(columns=raw_columns)
    logs = build_processing_logs(normalized_records, stats, audit_results)
    if eno_stats.get("eno_file"):
        logs.append(
            "Kesin Liste dosyası kullanıldı | "
            f"Sayaç/EIC eşleşmesi: {format_turkish_integer(eno_stats.get('eno_meter_lookup_count', 0))} | "
            f"EIC/müşteri eşleşmesi: {format_turkish_integer(eno_stats.get('eno_customer_lookup_count', 0))} | "
            f"ETSO dönüşümü: {format_turkish_integer(eno_stats.get('etso_converted', 0))} | "
            f"Müşteri doldurma: {format_turkish_integer(eno_stats.get('musteri_filled', 0))}."
        )
    else:
        logs.append("Kesin Liste dosyası bulunamadı; listeyle zenginleştirme uygulanmadı.")
    if stats.get("invoice_file_selected"):
        calculated_amount_count = sum(
            int(source.get("calculated_distribution_amount_count", 0))
            for source in invoice_sources
        )
        comparison_sheet_count = sum(
            bool(source.get("comparison_sheet_name"))
            for source in invoice_sources
        )
        missing_rate_count = sum(
            int(source.get("missing_distribution_rate_count", 0))
            for source in invoice_sources
        )
        logs.append(
            f"{OUR_INVOICES_LABEL}: "
            f"{format_turkish_integer(stats.get('invoice_file_count', 0))} dosya, "
            f"{format_turkish_integer(stats.get('invoice_sheet_count', 0))} kaynak sayfa, "
            f"{format_turkish_integer(calculated_amount_count)} olması gereken dağıtım bedeli hesaplandı."
        )
        if comparison_sheet_count:
            logs.append(
                f"{format_turkish_integer(comparison_sheet_count)} adet Bizim Kestiğimiz 2 karşılaştırma sayfası oluşturuldu."
            )
        reconciliation = next(
            (
                source for source in invoice_sources
                if source.get("reconciliation_sheet_name")
                or source.get("reconciliation_sheet_warning")
            ),
            None,
        )
        if reconciliation and reconciliation.get("reconciliation_sheet_name"):
            logs.append(
                "Karşılaştırma sayfası: "
                f"{format_turkish_integer(reconciliation.get('reconciliation_invoice_count', 0))} kesin fatura, "
                f"{format_turkish_integer(reconciliation.get('reconciliation_matched_count', 0))} ETSO eşleşmesi, "
                f"{format_turkish_integer(reconciliation.get('reconciliation_unmatched_count', 0))} eşleşmeyen, "
                f"{format_turkish_integer(reconciliation.get('reconciliation_ambiguous_count', 0))} çoklu kayıt."
            )
        consumption_reconciliation = next(
            (
                source for source in invoice_sources
                if source.get("reconciliation_consumption_sheet_name")
                or source.get("reconciliation_consumption_comparison_sheet_name")
                or source.get("reconciliation_consumption_warning")
            ),
            None,
        )
        if consumption_reconciliation and (
            consumption_reconciliation.get("reconciliation_consumption_sheet_name")
            or consumption_reconciliation.get("reconciliation_consumption_comparison_sheet_name")
        ):
            logs.append(
                "Uzlaştırma tüketimi: "
                f"{format_turkish_integer(consumption_reconciliation.get('reconciliation_consumption_source_row_count', 0))} uygun sayaç satırı, "
                f"{format_turkish_integer(consumption_reconciliation.get('reconciliation_consumption_etso_count', 0))} benzersiz ETSO, "
                f"{format_turkish_integer(consumption_reconciliation.get('reconciliation_consumption_matched_count', 0))} eşleşen karşılaştırma satırı, "
                f"{format_turkish_integer(consumption_reconciliation.get('reconciliation_consumption_unmatched_count', 0))} eşleşmeyen karşılaştırma satırı."
            )
        if missing_rate_count:
            logs.append(
                "Tarife Grubu için birim fiyat bulunamadığından hesaplanamayan satır: "
                f"{format_turkish_integer(missing_rate_count)}."
            )
        for source in invoice_sources:
            warning = (
                source.get("comparison_sheet_warning")
                or source.get("reconciliation_sheet_warning")
                or source.get("reconciliation_consumption_warning")
                or source.get("distribution_amount_warning")
                or source.get("tariff_rate_error")
                or source.get("power_rate_error")
            )
            if warning:
                logs.append(f"{source.get('output_sheet_name', OUR_INVOICES_LABEL)}: {warning}")
    return raw_df, stats, audit_results, output_bytes, logs, invoice_sources





def add_summary_worksheet(output_bytes: bytes, records: list[dict[str, Any]]) -> bytes:
    """Rename the detail sheet and append formula-driven company totals."""
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    summary_fields = [
        ("Aktif Enerji Tüketim (kWh)", "Aktif Enerji Tüketim (kWh)"),
        ("Dağıtım Bedeli (TL)", "Dağıtım Bedeli (TL)"),
        ("Güç Bedeli (TL)", "Güç Bedeli (TL)"),
        ("Güç Aşım Bedeli (TL)", "Güç Aşım Bedeli (TL)"),
        ("Reaktif Bedel (TL)", "Reaktif Bedel (TL)"),
        ("KDV Matrahı (TL)", "KDV Matrahı (TL)"),
        ("KDV (TL)", "KDV"),
        ("Toplam (TL)", "Toplam (TL)"),
    ]
    company_first_rows: dict[str, tuple[int, bool]] = {}
    for row_index, record in enumerate(records, start=2):
        company_value = record.get("Dağıtım Bölgesi")
        company = str(company_value).strip() if company_value is not None else ""
        is_unknown_company = not company or company.casefold() == "nan"
        if not company or company.casefold() == "nan":
            company = "Bilinmeyen dağıtım şirketi"
        company_first_rows.setdefault(company, (row_index, is_unknown_company))

    workbook = load_workbook(io.BytesIO(output_bytes))
    detail_sheet = workbook["Çıkarılan Veriler"]
    detail_sheet.title = "Dağıtımın Kestiği"
    summary_sheet = workbook.create_sheet("Özet")

    source_columns = {
        str(cell.value).strip(): cell.column
        for cell in detail_sheet[1]
        if cell.value is not None
    }
    company_column = source_columns["Dağıtım Bölgesi"]
    company_letter = detail_sheet.cell(1, company_column).column_letter
    last_detail_row = max(detail_sheet.max_row, 2)
    detail_name = "'Dağıtımın Kestiği'"
    company_range = f"{detail_name}!${company_letter}$2:${company_letter}${last_detail_row}"

    headers = ["Dağıtım Şirketi", *(header for header, _ in summary_fields)]
    summary_sheet.append(headers)
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(fill_type="solid", fgColor="305496")
    for cell in summary_sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    number_format = '#,##0.00;[Red]-#,##0.00;"-"'
    for company, (source_row, is_unknown_company) in company_first_rows.items():
        row_index = summary_sheet.max_row + 1
        company_source_cell = f"{detail_name}!${company_letter}${source_row}"
        summary_sheet.cell(
            row_index,
            1,
            f'=IF(OR(TRIM({company_source_cell})="",LOWER(TRIM({company_source_cell}))="nan"),"Bilinmeyen dağıtım şirketi",{company_source_cell})',
        )
        summary_sheet.cell(row_index, 1).alignment = Alignment(vertical="center")
        for column_index, (_, source_field) in enumerate(summary_fields, start=2):
            # Reaktif toplamı yalnızca Dağıtımın Kestiği sayfasındaki M sütunundan alınır.
            source_column = 13 if source_field == "Reaktif Bedel (TL)" else source_columns[source_field]
            source_letter = detail_sheet.cell(1, source_column).column_letter
            sum_range = f"{detail_name}!${source_letter}$2:${source_letter}${last_detail_row}"
            sum_formula = f"SUMIF({company_range},$A{row_index},{sum_range})"
            cell = summary_sheet.cell(row_index, column_index)
            if source_field == "Toplam (TL)":
                cell.value = f"=C{row_index}+D{row_index}+E{row_index}+F{row_index}+H{row_index}"
            elif is_unknown_company:
                cell.value = f'=SUMIF({company_range},"",{sum_range})+SUMIF({company_range},"nan",{sum_range})'
            else:
                cell.value = f"={sum_formula}"
            cell.number_format = number_format
            cell.alignment = Alignment(horizontal="right", vertical="center")

    summary_sheet.freeze_panes = "A2"
    summary_sheet.auto_filter.ref = f"A1:I{summary_sheet.max_row}"
    summary_sheet.row_dimensions[1].height = 32
    for column_cells in summary_sheet.columns:
        column_letter = column_cells[0].column_letter
        max_length = max((len(str(cell.value or "")) for cell in column_cells), default=0)
        summary_sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 16), 42)

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"

    result_buffer = io.BytesIO()
    workbook.save(result_buffer)
    return result_buffer.getvalue()


def render_metric_card(container: Any, label: str, value: str) -> None:
    container.markdown(
        f"""
        <div class="ck-card">
            <div class="ck-card-label">{label}</div>
            <div class="ck-card-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_selection_check(container: Any, is_selected: bool) -> None:
    if is_selected:
        container.markdown(
            '<span class="ck-selection-check" title="Seçildi" aria-label="Seçildi">✓</span>',
            unsafe_allow_html=True,
        )
    else:
        container.empty()


def get_upload_workspace() -> UploadWorkspace:
    """Keep uploaded inputs private to this browser's Streamlit session."""
    if "_upload_workspace" not in st.session_state:
        from new_mapping_parser import parse_mapping_file

        region_names: set[str] = set()
        mapping_file = BASE_DIR / "SKF Başlıkları.xlsx"
        if mapping_file.is_file():
            try:
                with contextlib.redirect_stdout(io.StringIO()):
                    mapping = parse_mapping_file(str(mapping_file))
                region_names.update(mapping.get("regions", {}))
                aliases = mapping.get("distribution_names", {})
                region_names.update(aliases)
                region_names.update(aliases.values())
            except (OSError, ValueError, KeyError, zipfile.BadZipFile):
                # The analysis panel reports an unavailable/invalid mapping separately.
                pass
        st.session_state["_upload_workspace"] = UploadWorkspace(region_names=region_names)
    return st.session_state["_upload_workspace"]


def synchronize_browser_upload(
    uploads: list[Any],
    *,
    category: str,
    selected_key: str,
    kind: str = "file",
) -> str:
    """Materialize changed uploads once and invalidate results when inputs change."""
    identity_key = f"_upload_identity_{category}"
    error_key = f"_upload_error_{category}"
    identity = (
        (kind, tuple(
            (
                getattr(upload, "file_id", None)
                or hashlib.sha256(upload.getvalue()).hexdigest(),
                upload.name,
                upload.size,
            )
            for upload in uploads
        ))
        if uploads else None
    )
    selected_path = st.session_state.get(selected_key, "")
    changed = identity != st.session_state.get(identity_key)
    missing_path = bool(uploads and selected_path and not Path(selected_path).exists())
    if changed or missing_path:
        new_path = ""
        error_message = ""
        workspace = get_upload_workspace()
        try:
            if not uploads:
                workspace.clear(category)
            elif kind == "directory":
                new_path = workspace.stage_directory(uploads)
            elif kind == "zip":
                new_path = workspace.stage_zip(uploads[0])
            else:
                new_path = workspace.stage_file(uploads[0], category)
        except (UploadInputError, OSError, zipfile.BadZipFile) as error:
            workspace.clear(category)
            error_message = str(error) or "Dosya hazırlanamadı. Lütfen yeniden seçin."
        if new_path != selected_path or changed:
            clear_analysis_state()
        st.session_state[selected_key] = new_path
        st.session_state[identity_key] = identity
        st.session_state[error_key] = error_message
        selected_path = new_path
    return selected_path


def synchronize_browser_excel_upload(
    widget_key: str, category: str, selected_key: str
) -> None:
    """Apply widget changes before controls such as the sidebar run button render."""
    uploaded = st.session_state.get(widget_key)
    synchronize_browser_upload(
        [uploaded] if uploaded is not None else [],
        category=category,
        selected_key=selected_key,
    )


def render_browser_excel_picker(
    label: str,
    *,
    category: str,
    selected_key: str,
    container: Any,
    help: str | None = None,
) -> str:
    generation = st.session_state.get(f"_upload_generation_{category}", 0)
    widget_key = f"browser_upload_{category}_{generation}"
    uploaded = container.file_uploader(
        label,
        type=["xlsx", "xlsm", "xls"],
        key=widget_key,
        help=help,
        max_upload_size=200,
        on_change=synchronize_browser_excel_upload,
        args=(widget_key, category, selected_key),
    )
    path = synchronize_browser_upload(
        [uploaded] if uploaded is not None else [],
        category=category,
        selected_key=selected_key,
    )
    if error_message := st.session_state.get(f"_upload_error_{category}"):
        container.error(error_message)
    return path


def render_browser_sidebar_inputs() -> tuple[str, str, str]:
    st.sidebar.markdown("**📁 Dağıtım Şirketi SKF**")
    upload_mode = st.sidebar.radio(
        "Dağıtım dosyalarını seçme biçimi",
        ("Klasör", "ZIP"),
        horizontal=True,
        key="distribution_upload_mode",
        label_visibility="collapsed",
    )
    if upload_mode == "Klasör":
        uploaded_files = st.sidebar.file_uploader(
            "Dağıtım klasörünü seçin",
            type=["xlsx", "xls", "xml", "html"],
            accept_multiple_files="directory",
            key="browser_upload_distribution_directory",
            max_upload_size=50,
            help="Bölge klasörlerinin bulunduğu ana klasörü seçin; alt klasörler de alınır.",
        )
        kind = "directory"
    else:
        uploaded_zip = st.sidebar.file_uploader(
            "Dağıtım klasörünü ZIP olarak seçin",
            type=["zip"],
            key="browser_upload_distribution_zip",
            max_upload_size=200,
            help="ZIP içinde bölge klasörlerinin yapısını koruyun.",
        )
        uploaded_files = [uploaded_zip] if uploaded_zip is not None else []
        kind = "zip"
    root_folder = synchronize_browser_upload(
        uploaded_files,
        category="distribution",
        selected_key="selected_root_folder",
        kind=kind,
    )
    if error_message := st.session_state.get("_upload_error_distribution"):
        st.sidebar.error(error_message)
    if root_folder and uploaded_files:
        first_name = uploaded_files[0].name.replace("\\", "/")
        st.session_state["browser_distribution_label"] = (
            first_name.split("/")[0] if kind == "directory" else Path(first_name).stem
        )
    else:
        st.session_state["browser_distribution_label"] = "SKF"

    invoice_file = render_browser_excel_picker(
        "📄 Bizim Kestiğimiz Faturalar",
        category="invoice",
        selected_key="selected_invoice_file",
        container=st.sidebar,
    )
    render_reconciliation_consumption_file_picker()
    selected_eno_file = render_browser_excel_picker(
        "📄 Kesin Liste (isteğe bağlı)",
        category="eno",
        selected_key="selected_eno_file",
        container=st.sidebar,
        help="Sayaç ID → EIC ve müşteri eşleştirmesi için Excel dosyasını seçin.",
    )
    st.sidebar.caption("Kesin Liste seçilmezse yüklenen dağıtım klasöründe otomatik aranır.")
    return root_folder, invoice_file, selected_eno_file


def render_native_sidebar_inputs() -> tuple[str, str, str]:
    root_folder = st.session_state.get("selected_root_folder", "").strip()
    folder_button, folder_status = st.sidebar.columns([9, 1], gap="small", vertical_alignment="center")
    with folder_button:
        select_folder = st.button(
            "📁 Dağıtım Şirketi SKF",
            width="stretch",
            key="select_distribution_folder",
        )
    with folder_status:
        render_selection_check(folder_status, bool(root_folder and os.path.isdir(root_folder)))

    if select_folder:
        selected_folder = select_folder_via_dialog(
            "Dağıtım şirketi SKF klasörlerinin bulunduğu ana dizini seçin"
        )
        if selected_folder:
            clear_analysis_state()
            st.session_state["selected_root_folder"] = selected_folder
            st.session_state["manual_selected_root_folder"] = selected_folder
            st.rerun()
        else:
            st.sidebar.info("Klasör seçilmedi veya pencere açılamadı. Dosya yollarını elle girebilirsiniz.")

    invoice_button, invoice_status = st.sidebar.columns([9, 1], gap="small", vertical_alignment="center")
    with invoice_button:
        select_invoice = st.button(
            "📄 Bizim Kestiğimiz Faturalar",
            width="stretch",
            key="select_our_invoice_file",
        )
    invoice_file = st.session_state.get("selected_invoice_file", "").strip()
    with invoice_status:
        render_selection_check(invoice_status, bool(invoice_file and os.path.isfile(invoice_file)))

    if select_invoice:
        selected_file = select_excel_file_via_dialog(
            "Bizim Kestiğimiz Faturalar için Excel dosyasını seçin"
        )
        if selected_file:
            clear_analysis_state()
            st.session_state["selected_invoice_file"] = selected_file
            st.session_state["manual_selected_invoice_file"] = selected_file
            st.rerun()
        else:
            st.sidebar.info("Excel dosyası seçilmedi veya dosya penceresi açılamadı. Dosya yolunu aşağıya yazabilirsiniz.")

    render_reconciliation_consumption_file_picker()

    selected_eno_file = st.session_state.get("selected_eno_file", "")
    eno_button, eno_status = st.sidebar.columns([9, 1], gap="small", vertical_alignment="center")
    with eno_button:
        select_eno = st.button(
            "📄 Kesin Liste (isteğe bağlı)",
            width="stretch",
            key="select_eno_file",
            help="Sayaç ID → EIC ve müşteri eşleştirmesi için Kesin Liste Excel dosyasını seçin.",
        )
    with eno_status:
        render_selection_check(
            eno_status,
            bool(selected_eno_file and os.path.isfile(selected_eno_file)),
        )
    st.sidebar.caption("Kesin Liste seçilmezse SKF klasörünün içinde otomatik aranır.")

    if select_eno:
        new_eno_file = select_excel_file_via_dialog(
            "Kesin Liste Excel dosyasını seçin",
            initial_path=selected_eno_file,
        )
        if new_eno_file:
            if Path(new_eno_file).suffix.casefold() not in {".xlsx", ".xlsm", ".xls"}:
                st.sidebar.error("Lütfen bir Excel dosyası seçin.")
            else:
                st.session_state["selected_eno_file"] = new_eno_file
                st.session_state["manual_selected_eno_file"] = new_eno_file
                if new_eno_file != selected_eno_file:
                    clear_analysis_state()
                save_error = save_eno_file_selection(new_eno_file)
                if save_error:
                    st.sidebar.warning(
                        "Dosya bu oturum için seçildi ancak sonraki açılışlarda hatırlanamadı."
                    )
                else:
                    st.rerun()
        else:
            st.sidebar.info("Kesin Liste dosyası seçilmedi veya dosya penceresi açılamadı.")

    for selected_key, manual_key in (
        ("selected_root_folder", "manual_selected_root_folder"),
        ("selected_invoice_file", "manual_selected_invoice_file"),
        ("selected_eno_file", "manual_selected_eno_file"),
    ):
        if manual_key not in st.session_state:
            st.session_state[manual_key] = st.session_state.get(selected_key, "")

    with st.sidebar.expander("Dosya yollarını elle gir (isteğe bağlı)", expanded=False):
        st.caption("Dosya seçme penceresi çalışmıyorsa yolları buradan girebilirsiniz.")
        root_folder = st.text_input(
            "Dağıtım Şirketi SKF klasör yolu",
            key="manual_selected_root_folder",
            type="password",
            placeholder="",
        ).strip()
        invoice_file = st.text_input(
            "Bizim Kestiğimiz Faturalar Excel dosyası",
            key="manual_selected_invoice_file",
            type="password",
            placeholder="",
        ).strip()
        selected_eno_file = st.text_input(
            "Kesin Liste Excel dosyası (isteğe bağlı)",
            key="manual_selected_eno_file",
            type="password",
            placeholder="",
            help="Boş bırakılırsa seçilen SKF ana klasöründe ve alt klasörlerinde aranır.",
        ).strip()
    if (
        root_folder != st.session_state.get("selected_root_folder", "")
        or invoice_file != st.session_state.get("selected_invoice_file", "")
        or selected_eno_file != st.session_state.get("selected_eno_file", "")
    ):
        clear_analysis_state()
    st.session_state["selected_root_folder"] = root_folder
    st.session_state["selected_invoice_file"] = invoice_file
    st.session_state["selected_eno_file"] = selected_eno_file
    if load_eno_file_selection() != selected_eno_file:
        save_error = save_eno_file_selection(selected_eno_file)
        if save_error:
            st.sidebar.warning("Kesin Liste dosyası seçimi kalıcı olarak kaydedilemedi.")
    return root_folder, invoice_file, selected_eno_file


def render_sidebar() -> tuple[str, str, str | None, bool]:
    logo_path = BASE_DIR / "CK_Enerji_-_Yatay.png"
    if logo_path.exists():
        st.sidebar.image(str(logo_path), width="stretch")
    st.sidebar.subheader("📂 Veri Kaynağı")
    if uses_browser_file_inputs():
        root_folder, invoice_file, selected_eno_file = render_browser_sidebar_inputs()
    else:
        root_folder, invoice_file, selected_eno_file = render_native_sidebar_inputs()
    if selected_eno_file and os.path.isfile(selected_eno_file):
        valid_eno_file = True
    elif selected_eno_file:
        st.sidebar.error("Seçilen Kesin Liste dosyası bulunamadı.")
        valid_eno_file = False
    else:
        valid_eno_file = True

    # Mapping dosyası sabittir: kodlarla aynı dizindeki 'SKF Başlıkları.xlsx'
    # Kullanıcının veri klasöründe aranmaz; sistem gömülü (embedded) olarak kullanır.
    embedded_mapping_path = str(BASE_DIR / "SKF Başlıkları.xlsx")
    if not os.path.isfile(embedded_mapping_path):
        st.sidebar.error("⚠️ Sistem mapping dosyası bulunamadı: SKF Başlıkları.xlsx")
    mapping_path = embedded_mapping_path if os.path.isfile(embedded_mapping_path) else None
    valid_invoice_file = not invoice_file or os.path.isfile(invoice_file)
    if invoice_file and not valid_invoice_file:
        st.sidebar.error("Seçilen Bizim Kestiğimiz Faturalar Excel dosyası bulunamadı.")

    only_our_invoice_data = st.sidebar.checkbox(
        "Yaln\u0131zca Bizim Kesti\u011fimiz Faturalar",
        key="invoice_only_mode",
        help="Se\u00e7ili da\u011f\u0131t\u0131m klas\u00f6r\u00fc bu i\u015flemde kullan\u0131lmaz; yaln\u0131zca fatura verisi haz\u0131rlan\u0131r.",
    )
    invoice_only_mode = (
        (only_our_invoice_data or not root_folder)
        and bool(invoice_file and valid_invoice_file)
    )
    upload_categories = ["invoice", "reconciliation", "system_usage"]
    if not only_our_invoice_data:
        upload_categories.append("distribution")
    if not invoice_only_mode:
        upload_categories.append("eno")
    valid_browser_uploads = not uses_browser_file_inputs() or not any(
        st.session_state.get(f"_upload_error_{category}")
        for category in upload_categories
    )
    run_button = st.sidebar.button(
        "Yaln\u0131zca Fatura Verisini Haz\u0131rla" if invoice_only_mode else "🚀 Verileri Birleştir ve Karşılaştır",
        type="primary",
        width="stretch",
        disabled=not (
            ((root_folder and os.path.isdir(root_folder)) or invoice_only_mode)
            and mapping_path
            and valid_invoice_file
            and (valid_eno_file or invoice_only_mode)
            and valid_browser_uploads
        ),
    )

    if only_our_invoice_data:
        st.sidebar.info("Da\u011f\u0131t\u0131m klas\u00f6r\u00fc bu i\u015flemde kullan\u0131lmayacak.")
    else:
        candidate_count = count_candidate_region_folders(root_folder)
        st.sidebar.info(
            f"Bulunan aday bölge klasörü: {format_turkish_integer(candidate_count)} / "
            f"{format_turkish_integer(EXPECTED_REGION_COUNT)}"
        )
    if uses_browser_file_inputs():
        st.sidebar.caption(
            "Seçtiğiniz dosyalar bu oturum için sunucuya yüklenip işlenir. "
            "Sonucu Excel olarak indirebilirsiniz."
        )
    else:
        st.sidebar.success("Verileriniz %100 yerelde işlenmektedir, ETSO kodları gizlidir.")



    # Karanlık / Aydınlık mod
    st.sidebar.divider()
    st.sidebar.toggle("🌙 Karanlık Mod", key="dark_mode")

    return ("" if only_our_invoice_data else root_folder), invoice_file, mapping_path, run_button


def render_system_usage_file_picker() -> None:
    if uses_browser_file_inputs():
        render_browser_excel_picker(
            "Sistem Kullanım Bedelleri",
            category="system_usage",
            selected_key="selected_system_usage_file",
            container=st,
        )
        return
    selected_file = st.session_state.get("selected_system_usage_file", "")
    spacer, label_column, button_column, status_column = st.columns(
        [4, 2.8, 1.1, 0.8],
        gap="small",
        vertical_alignment="center",
    )
    with label_column:
        st.markdown("Sistem Kullanım Bedelleri")
    with button_column:
        select_system_usage = st.button(
            "₺",
            key="select_system_usage_file",
            help="Sistem Kullanım Bedelleri Excel dosyasını seçin.",
            width="content",
        )
    with status_column:
        render_selection_check(
            status_column,
            bool(selected_file and os.path.isfile(selected_file)),
        )

    if select_system_usage:
        new_file = select_excel_file_via_dialog(
            "Sistem Kullanım Bedelleri için Excel dosyasını seçin",
            initial_path=selected_file,
        )
        if new_file:
            if Path(new_file).suffix.casefold() not in {".xlsx", ".xlsm", ".xls"}:
                st.error("Lütfen bir Excel dosyası seçin.")
            else:
                st.session_state["selected_system_usage_file"] = new_file
                if new_file != selected_file:
                    clear_analysis_state()
                save_error = save_system_usage_file_selection(new_file)
                if save_error:
                    st.warning(
                        "Dosya bu oturum için seçildi ancak sonraki açılışlarda hatırlanamadı."
                    )
                else:
                    st.rerun()
        else:
            st.info("Dosya seçilmedi veya dosya penceresi açılamadı.")

    selected_file = st.session_state.get("selected_system_usage_file", "")
    if not selected_file or not os.path.isfile(selected_file):
        st.warning(
            "Sistem Kullanım Bedelleri dosyası bulunamadı. Butona basıp geçerli dosyayı seçin.",
            icon="⚠️",
        )


def render_reconciliation_consumption_file_picker() -> None:
    if uses_browser_file_inputs():
        render_browser_excel_picker(
            "⚖️ Uzlaştırma Tüketim",
            category="reconciliation",
            selected_key="selected_reconciliation_consumption_file",
            container=st.sidebar,
            help="Bu dosya bir analizde kullanılır; sonraki analiz için yeniden seçin.",
        )
        return
    selected_file = st.session_state.get("selected_reconciliation_consumption_file", "")
    button_column, status_column = st.sidebar.columns(
        [6, 0.8],
        gap="small",
        vertical_alignment="center",
    )
    with button_column:
        select_reconciliation_file = st.button(
            "⚖️ Uzlaştırma Tüketim",
            key="select_reconciliation_consumption_file",
            help="Bu dosya bir analizde kullanılır; sonraki analiz için yeniden seçin.",
            width="stretch",
        )
    with status_column:
        render_selection_check(
            status_column,
            bool(selected_file and os.path.isfile(selected_file)),
        )

    if select_reconciliation_file:
        new_file = select_excel_file_via_dialog(
            "Uzlaştırma Tüketim için Excel dosyasını seçin",
            initial_path=selected_file,
        )
        if new_file:
            if Path(new_file).suffix.casefold() not in {".xlsx", ".xlsm", ".xls"}:
                st.sidebar.error("Lütfen bir Excel dosyası seçin.")
            else:
                st.session_state["selected_reconciliation_consumption_file"] = new_file
                if new_file != selected_file:
                    clear_analysis_state()
                st.rerun()
        else:
            st.sidebar.info("Dosya seçilmedi veya dosya penceresi açılamadı.")

    selected_file = st.session_state.get("selected_reconciliation_consumption_file", "")
    if selected_file and not os.path.isfile(selected_file):
        st.sidebar.warning(
            "Seçilen Uzlaştırma Tüketim dosyası bulunamadı. Geçerli dosyayı yeniden seçin.",
            icon="⚠️",
        )


_SOURCE_WARNING_OUTPUT_FIELDS = {
    "etso": ("Etso Kodu",),
    "musteri": ("Müşteri",),
    "tarife": ("Tarife Grubu",),
    "ag_og": ("AG OG",),
    "terim": ("TERİM",),
    "güç_kw": ("Güç kW",),
    "kurulu_güç": ("KURULU GÜÇ",),
    "aktif_enerji": ("Aktif Enerji Tüketim (kWh)",),
    "trafo_kaybı": ("Aktif Enerji Tüketim (kWh)",),
    "dagitim_bedeli": ("Dağıtım Bedeli(TL)",),
    "güç_bedeli": ("Güç Bedeli(TL)",),
    "güç_aşım": ("Güç Aşım Bedeli (TL)",),
    "reaktif": ("Reaktif Bedel (TL)",),
    "reaktif2": ("Reaktif Bedel (TL)",),
    "reaktif_tenzil": ("Reaktif Bedel (TL)", "İlk Reaktif"),
    "reaktif_tenzil2": ("Reaktif Bedel (TL)", "İlk Reaktif"),
}

def _merged_warning_field_found(
    merged_df: pd.DataFrame,
    distribution_company: str,
    source_field: str,
) -> bool:
    """Check whether a warned source field has a populated value after merging."""
    if merged_df.empty or "Dağıtım Bölgesi" not in merged_df.columns:
        return False

    company_rows = merged_df.loc[
        merged_df["Dağıtım Bölgesi"].fillna("").astype(str).str.casefold()
        == str(distribution_company).casefold()
    ]
    if company_rows.empty:
        return False

    output_fields = _SOURCE_WARNING_OUTPUT_FIELDS.get(source_field)
    if output_fields is None:
        # Mapping-specific financial fields are consolidated into these totals.
        output_fields = ("Tazminat Bedeli", "Standart Dışı Tutar (TL)")

    for output_field in output_fields:
        if output_field not in company_rows.columns:
            continue
        for value in company_rows[output_field].tolist():
            if value is None or pd.isna(value):
                continue
            if isinstance(value, str):
                value = value.strip()
                if not value or value.casefold() in {"none", "null", "nan", "boş"}:
                    continue
            if is_numeric_text(value) and abs(safe_float(value)) <= 1e-12:
                continue
            return True
    return False


def build_format_warning_rows(
    warnings: list[dict[str, Any]],
    merged_df: pd.DataFrame,
) -> list[dict[str, str]]:
    """Collapse repeated sheet warnings into unique company/header rows."""
    grouped: dict[str, dict[tuple[str, str, str], list[str]]] = {}
    for warning in warnings:
        company = str(warning.get("Dağıtım Bölgesi") or "Bilinmeyen dağıtım şirketi")
        source_field = str(warning.get("Alan") or "")
        expected_header = warning.get("Beklenen başlık")
        expected_column = warning.get("Eşlenen sütun")
        current_header = warning.get("Sütundaki başlık")
        if expected_header is not None:
            expected = f"{expected_column} - {expected_header}" if expected_column else str(expected_header)
            current = f"{expected_column} - {current_header or '(boş)'}" if expected_column else str(current_header or "(boş)")
            is_found = (
                warning.get("Başlık eşleşti") is True
                or _merged_warning_field_found(merged_df, company, source_field)
            )
            found_location = str(warning.get("Bulunduğu Başlık/Sütun") or "-")
        else:
            # The saved layout is represented by a signature, so this fallback
            # cannot name an individual expected or resolved column.
            expected = f"Kayıtlı düzen - {warning.get('Beklenen başlık', 'başlık düzeni')}"
            current = "Başlık düzeni - değişmiş"
            header_location = warning.get("Başlığın bulunduğu sütun")
            if header_location:
                current += f" ({header_location})"
            is_found = not merged_df.empty and bool(
                (merged_df.get("Dağıtım Bölgesi", pd.Series(dtype=str)).fillna("").astype(str).str.casefold()
                 == company.casefold()).any()
            )
            found_location = "-"

        status = "Bulundu" if is_found else "Bulunamadı"
        key = (expected, current, status)
        company_rows = grouped.setdefault(company, {})
        locations = company_rows.setdefault(key, [])
        location = found_location if is_found else "-"
        if location not in locations:
            locations.append(location)

    rows = []
    for company, company_rows in grouped.items():
        for (expected, current, status), locations in company_rows.items():
            rows.append({
                "Dağıtım Şirketi": company,
                "Olması Gereken Başlık/Sütun": expected,
                "Mevcut Başlık/Sütun": current,
                "Bulundu/Bulunamadı": status,
                "Bulunduğu Başlık/Sütun": ", ".join(location for location in locations if location != "-") or "-",
            })
    return rows


def render_format_warning_table(rows: list[dict[str, str]]) -> None:
    """Render a wrapping HTML table so full source header details stay visible."""
    columns = [
        "Dağıtım Şirketi",
        "Olması Gereken Başlık/Sütun",
        "Mevcut Başlık/Sütun",
        "Bulundu/Bulunamadı",
        "Bulunduğu Başlık/Sütun",
    ]
    widths = (18, 21, 21, 15, 25)
    header_cells = "".join(
        f'<th style="width:{width}%;text-align:left;padding:8px;border:1px solid #808080;'
        f'white-space:normal;overflow-wrap:anywhere">{html.escape(column)}</th>'
        for column, width in zip(columns, widths)
    )
    body_rows = []
    company_groups: dict[str, list[dict[str, str]]] = {}
    for row in rows:
        company_groups.setdefault(row["Dağıtım Şirketi"], []).append(row)

    for company, company_rows in company_groups.items():
        for row_index, row in enumerate(company_rows):
            cells = ""
            if row_index == 0:
                cells += (
                    f'<td rowspan="{len(company_rows)}" style="padding:8px;border:1px solid #808080;'
                    f'vertical-align:middle;white-space:normal;overflow-wrap:anywhere">'
                    f'{html.escape(company)}</td>'
                )
            cells += "".join(
                '<td style="padding:8px;border:1px solid #808080;vertical-align:top;'
                f'white-space:normal;overflow-wrap:anywhere">{html.escape(str(row.get(column, "-")))}</td>'
                for column in columns[1:]
            )
            body_rows.append(f"<tr>{cells}</tr>")

    table_html = (
        '<div style="width:100%;overflow-x:auto">'
        '<table style="width:100%;table-layout:fixed;border-collapse:collapse">'
        f"<thead><tr>{header_cells}</tr></thead><tbody>{''.join(body_rows)}</tbody></table></div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


def render_analysis_panel(
    root_folder: str,
    invoice_file: str,
    mapping_path: str | None,
    run_button: bool = False,
) -> None:
    st.markdown(
        '<div class="ck-section-title">⚙️ Veri İşleme ve Kurumsal Rapor</div>',
        unsafe_allow_html=True,
    )

    valid_root = bool(root_folder and os.path.isdir(root_folder))
    valid_invoice_file = not invoice_file or os.path.isfile(invoice_file)
    valid_mapping = bool(mapping_path and os.path.isfile(mapping_path))
    selected_eno_file = str(st.session_state.get("selected_eno_file", "")).strip()
    valid_eno_file = not selected_eno_file or os.path.isfile(selected_eno_file)
    selected_reconciliation_consumption_file = str(
        st.session_state.get("selected_reconciliation_consumption_file", "")
    ).strip()
    if not root_folder and invoice_file and valid_invoice_file:
        st.info("Da\u011f\u0131t\u0131m klas\u00f6r\u00fc y\u00fcklenmeden yaln\u0131zca Bizim Kesti\u011fimiz Faturalar verisi haz\u0131rlanabilir.")
    elif not root_folder:
        st.info("Analize başlamak için soldan dosyaları yükleyin.")
    elif not valid_root:
        st.error("Seçilen Dağıtım Şirketi SKF klasörü bulunamadı.")
    if invoice_file and not valid_invoice_file:
        st.error("Seçilen Bizim Kestiğimiz Faturalar Excel dosyası bulunamadı.")
    if selected_eno_file and not valid_eno_file:
        st.error("Seçilen Kesin Liste Excel dosyası bulunamadı; dosya yolunu kontrol edin.")
    if not valid_mapping:
        st.error("Seçilen klasörde uygun SKF mapping dosyası bulunamadı.")

    current_signature = (
        os.path.normcase(os.path.abspath(root_folder)) if valid_root else "",
        os.path.normcase(os.path.abspath(invoice_file)) if invoice_file and valid_invoice_file else "",
        os.path.normcase(os.path.abspath(mapping_path)) if valid_mapping and mapping_path else "",
        os.path.normcase(os.path.abspath(st.session_state.get("selected_system_usage_file", "")))
        if st.session_state.get("selected_system_usage_file")
        and os.path.isfile(st.session_state.get("selected_system_usage_file", ""))
        else "",
        os.path.normcase(os.path.abspath(selected_reconciliation_consumption_file))
        if selected_reconciliation_consumption_file
        else "",
        os.path.normcase(os.path.abspath(selected_eno_file)) if selected_eno_file else "",
    )
    if (
        st.session_state.get("processing_done")
        and st.session_state.get("analysis_signature") != current_signature
    ):
        clear_analysis_state()

    if (
        run_button
        and mapping_path
        and valid_invoice_file
        and (valid_eno_file or (not root_folder and bool(invoice_file)))
        and (valid_root or (not root_folder and bool(invoice_file)))
    ):
        clear_analysis_state()
        progress_bar = st.progress(0)
        status_placeholder = st.empty()

        def update_progress(message: str, percent: int) -> None:
            safe_percent = min(100, max(0, int(percent)))
            progress_bar.progress(safe_percent)
            status_placeholder.info(message)

        try:
            with st.spinner("Fatura dosyaları okunuyor ve ETSO kayıtları birleştiriliyor."):
                raw_df, stats, audit_results, output_bytes, logs, invoice_sources = run_analysis_pipeline(
                    root_folder,
                    mapping_path,
                    update_progress,
                    invoice_file=invoice_file or None,
                    system_usage_file=st.session_state.get("selected_system_usage_file") or None,
                    eno_file_path=selected_eno_file if valid_eno_file and selected_eno_file else None,
                    reconciliation_consumption_file=selected_reconciliation_consumption_file or None,
                )
            st.session_state["extracted_df"] = raw_df
            st.session_state["stats"] = stats
            st.session_state["audit_results"] = audit_results
            st.session_state["output_bytes"] = output_bytes
            st.session_state["invoice_sources"] = invoice_sources
            st.session_state["subscription_group_mismatches"] = next(
                (
                    source.get("subscription_group_mismatches", [])
                    for source in invoice_sources
                    if "subscription_group_mismatches" in source
                ),
                [],
            )
            st.session_state["processing_logs"] = logs
            if selected_reconciliation_consumption_file:
                st.session_state["selected_reconciliation_consumption_file"] = ""
                if uses_browser_file_inputs():
                    get_upload_workspace().clear("reconciliation")
                    generation_key = "_upload_generation_reconciliation"
                    st.session_state[generation_key] = st.session_state.get(generation_key, 0) + 1
                    st.session_state.pop("_upload_identity_reconciliation", None)
                    st.session_state.pop("_upload_error_reconciliation", None)
                current_signature = (*current_signature[:4], "", *current_signature[5:])
            st.session_state["analysis_signature"] = current_signature
            st.session_state["processing_done"] = True
            progress_bar.progress(100)
            status_placeholder.success("Analiz tamamlandı.")
            if selected_reconciliation_consumption_file:
                st.rerun()
        except Exception:
            clear_analysis_state()
            status_placeholder.empty()
            progress_bar.empty()
            st.error("İşlem tamamlanamadı. Seçilen dosyaları ve dosya içeriklerini kontrol edin.")

    if not st.session_state.get("processing_done"):
        return

    raw_df = st.session_state.get("extracted_df")
    if raw_df is None:
        raw_df = pd.DataFrame()
    stats = st.session_state.get("stats") or {}
    audit_results = st.session_state.get("audit_results") or {}
    output_bytes = st.session_state.get("output_bytes")
    invoice_sources = st.session_state.get("invoice_sources", [])
    invoice_output_blocked = bool(stats.get("invoice_file_selected") and stats.get("invoice_file_errors"))
    render_reconciliation_difference_report_area(invoice_sources)
    if stats.get("invoice_only_mode"):
        for invoice_error in stats.get("invoice_file_errors", []):
            st.error(f"{OUR_INVOICES_LABEL}: {invoice_error}")
        if invoice_output_blocked:
            st.info("Se\u00e7ilen fatura dosyas\u0131ndan Bizim Kesti\u011fimiz 2 sayfas\u0131 olu\u015fturulamad\u0131.")
        has_reconciliation_consumption = any(
            source.get("reconciliation_consumption_comparison_sheet_name")
            or source.get("reconciliation_consumption_warning")
            for source in st.session_state.get("invoice_sources", [])
        )
        if has_reconciliation_consumption:
            st.info(
                "Çıktıda yüklenen fatura verisi ve ETSO bazlı uzlaştırma tüketim karşılaştırması bulunur; "
                "dağıtım faturası yüklenmediği için dağıtım karşılaştırması yapılamadı."
            )
        else:
            st.info("\u00c7\u0131kt\u0131da y\u00fcklenen fatura verisi ve ETSO bazl\u0131 Bizim Kesti\u011fimiz 2 sayfas\u0131 bulunur; da\u011f\u0131t\u0131m kar\u015f\u0131la\u015ft\u0131rmas\u0131 yap\u0131lmad\u0131.")
        if output_bytes and not invoice_output_blocked:
            invoice_stem = Path(invoice_file).stem if invoice_file else "Bizim Kesti\u011fimiz"
            st.download_button(
                "Bizim Kesti\u011fimiz dosyas\u0131n\u0131 indir",
                data=output_bytes,
                file_name=f"{invoice_stem} - Bizim Kesti\u011fimiz.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="invoice_only_output_download",
                width="stretch",
                on_click="ignore",
            )
        st.markdown("#### Bizim Kesti\u011fimiz 2 \u00d6nizleme")
        st.dataframe(
            format_dataframe_for_ui(raw_df),
            width="stretch",
            hide_index=True,
            height=620,
        )
        with st.expander("\u0130\u015flem Loglar\u0131", expanded=False):
            for log_line in st.session_state.get("processing_logs", []):
                st.code(log_line, language=None)
        return
    for invoice_error in stats.get("invoice_file_errors", []):
        st.error(f"{OUR_INVOICES_LABEL}: {invoice_error}")
    if invoice_output_blocked:
        st.info("Seçilen fatura dosyası SKF çıktısına eklenemediği için indirme kapatıldı.")

    read_count = int(stats.get("read_region_count", 0))
    expected_count = int(stats.get("expected_region_count", EXPECTED_REGION_COUNT))
    missing_regions = list(stats.get("missing_regions", []))
    if read_count != expected_count or missing_regions:
        st.warning("Şu bölgeler okunamadı: " + ", ".join(missing_regions))

    source_header_warnings = stats.get("source_header_warnings", [])
    format_warning_rows = build_format_warning_rows(source_header_warnings, raw_df)
    if format_warning_rows:
        affected_company_count = len({row["Dağıtım Şirketi"] for row in format_warning_rows})
        with st.expander(
            f"⚠️ Format Farkı ({format_turkish_integer(affected_company_count)} dağıtım şirketi)",
            expanded=False,
        ):
            render_format_warning_table(format_warning_rows)

    category_validation_warnings = stats.get("category_validation_warnings", [])
    if category_validation_warnings:
        with st.expander(
            f"⚠️ Tarife / AG-OG / Terim değer kontrolü ({format_turkish_integer(len(category_validation_warnings))} bulgu)",
            expanded=True,
        ):
            st.warning(
                "Tarife Grubu yalnız Sanayi, Ticarethane veya Mesken; AG OG yalnız AG veya OG olabilir. "
                "Tanımlanamayan kaynak değerleri çıktıdan boş bırakıldı."
            )
            st.dataframe(
                pd.DataFrame(category_validation_warnings),
                hide_index=True,
                use_container_width=True,
            )

    metric_columns = st.columns(2)
    render_metric_card(metric_columns[0], "Toplam Kayıt", format_turkish_integer(len(raw_df)))
    render_metric_card(
        metric_columns[1],
        "Okunan Bölge",
        f"{format_turkish_integer(read_count)} / {format_turkish_integer(expected_count)}",
    )

    st.caption(
        f"Veri sağlık skoru: %{format_turkish_integer(audit_results.get('health_score', 100))} | "
        f"Birleştirilen mükerrer kayıt: {format_turkish_integer(stats.get('merged_duplicates', 0))} | "
        f"Filtrelenen bozuk kayıt: {format_turkish_integer(stats.get('absurd_filtered', 0))}"
    )

    st.markdown("#### 📊 Şirket Bazlı Karşılaştırma Raporları")
    company_column = next(
        (column for column in ("Dağıtım Bölgesi", "Dağıtım Şirketi") if column in raw_df.columns),
        None,
    )
    company_options_by_key: dict[str, str] = {}
    if company_column:
        for value in raw_df[company_column].dropna().tolist():
            company_name = str(value).strip()
            if company_name and company_name.casefold() != "nan":
                company_options_by_key.setdefault(_normalize_tariff_lookup_key(company_name), company_name)
    company_options = sorted(company_options_by_key.values(), key=str.casefold)
    if "company_report_selected_companies" in st.session_state:
        valid_company_options = set(company_options)
        st.session_state["company_report_selected_companies"] = [
            company
            for company in (st.session_state.get("company_report_selected_companies") or [])
            if company in valid_company_options
        ]
    selected_company_names = st.multiselect(
        "Raporlanacak dağıtım şirketlerini seçin",
        options=company_options,
        key="company_report_selected_companies",
        placeholder="Şirket seçin",
    )
    selected_company_signature = tuple(
        sorted(_normalize_tariff_lookup_key(company) for company in selected_company_names)
    )
    if st.session_state.get("company_report_selection") != selected_company_signature:
        st.session_state["company_report_selection"] = selected_company_signature
        st.session_state["company_comparison_reports"] = {}
        st.session_state["comparison_report_warning"] = None
        st.session_state["company_report_generated"] = False

    generate_reports = st.button(
        "Seçilen Şirketler İçin Rapor Oluştur",
        key="build_company_comparison_reports",
        disabled=not selected_company_names or not output_bytes or invoice_output_blocked,
    )
    if generate_reports:
        try:
            with st.spinner("Seçilen şirketlerin karşılaştırma raporları hazırlanıyor."):
                company_reports, _, comparison_report_warning = (
                    build_company_comparison_reports(
                        output_bytes,
                        selected_company_names,
                        include_group_mismatches=False,
                    )
                )
            st.session_state["company_comparison_reports"] = company_reports
            st.session_state["comparison_report_warning"] = comparison_report_warning
            st.session_state["company_report_generated"] = True
        except Exception as error:
            st.session_state["company_comparison_reports"] = {}
            st.session_state["comparison_report_warning"] = (
                "Şirket bazlı karşılaştırma raporları hazırlanamadı: "
                f"{type(error).__name__}: {error}"
            )
            st.session_state["company_report_generated"] = True

    generated_for_selection = (
        st.session_state.get("company_report_generated")
        and st.session_state.get("company_report_selection") == selected_company_signature
    )
    company_reports = st.session_state.get("company_comparison_reports") or {}
    if st.session_state.get("comparison_report_warning"):
        st.warning(st.session_state["comparison_report_warning"])
    elif generated_for_selection and company_reports:
        st.caption("Seçtiğiniz şirketlerin her birini ayrı Excel raporu olarak indirebilirsiniz.")
        report_items = sorted(company_reports.items(), key=lambda item: item[0].casefold())
        for start_index in range(0, len(report_items), 3):
            report_row = report_items[start_index : start_index + 3]
            report_columns = st.columns(len(report_row), gap="small")
            for offset, (column, (company_name, report_bytes)) in enumerate(zip(report_columns, report_row)):
                safe_company_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", company_name).strip(" .")
                safe_company_name = safe_company_name or "Dağıtım Şirketi"
                with column:
                    st.download_button(
                        company_name,
                        data=report_bytes,
                        file_name=f"{safe_company_name} - Karşılaştırma.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"company_comparison_download_{start_index + offset}",
                        width="stretch",
                        on_click="ignore",
                    )
    elif generated_for_selection:
        st.info("Seçtiğiniz şirketler için karşılaştırma kaydı bulunamadı.")
    else:
        st.caption("İstediğiniz şirketleri seçip rapor düğmesine basın; diğer şirketler bu işlemde çalıştırılmaz.")

    st.markdown("#### Abone Grubu Farkları")
    group_mismatches = st.session_state.get("subscription_group_mismatches") or []
    group_comparison_available = any(
        source.get("reconciliation_sheet_name") for source in invoice_sources
    )
    if group_comparison_available and group_mismatches:
        st.dataframe(
            pd.DataFrame(group_mismatches),
            hide_index=True,
            use_container_width=True,
        )
    elif group_comparison_available:
        st.success("Abone grubu farklı olan ETSO kaydı bulunamadı.")
    else:
        st.caption("Abone grubu farklarını görmek için kesin fatura ve dağıtım verisi içeren bir analiz çalıştırın.")

    preview_title, download_column = st.columns([5, 2], gap="small", vertical_alignment="center")
    with preview_title:
        st.markdown(
            f"#### 🔍 Veri Önizleme - {format_turkish_integer(len(raw_df))} Kayıt"
        )
    with download_column:
        if output_bytes and not invoice_output_blocked:
            selected_folder_name = (
                st.session_state.get("browser_distribution_label", "SKF")
                if uses_browser_file_inputs()
                else Path(os.path.normpath(root_folder)).name or "SKF"
            )
            st.download_button(
                "📥 SKF Dosyasını İndir",
                data=output_bytes,
                file_name=f"{selected_folder_name} SKF.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
                on_click="ignore",
            )
    st.caption(
        "Ekrandaki ETSO kodları maskelidir. İndirilen Excel, kaynak ETSO kodlarını "
        "değiştirmeden korur ve sütun genişliklerini içeriğe göre otomatik ayarlar."
    )
    display_df = format_dataframe_for_ui(raw_df)
    st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,
        height=620,
    )

    if invoice_sources:
        st.markdown(f"#### {OUR_INVOICES_LABEL}")
        st.caption(
            f"{len(invoice_sources)} sayfa çıktı dosyasına eklendi; her sayfa için "
            "Bizim Kestiğimiz 2 karşılaştırma sekmesi oluşturuldu."
        )

    with st.expander("📋 İşlem Logları", expanded=False):
        for log_line in st.session_state.get("processing_logs", []):
            st.code(log_line, language=None)



def main() -> None:
    inject_copy_shortcut_guard()
    inject_corporate_css()
    initialize_session_state()

    root_folder, invoice_file, mapping_path, run_button = render_sidebar()

    header_left, header_right = st.columns([4, 2], gap="small", vertical_alignment="top")
    with header_left:
        st.markdown('<div class="ck-title">⚡ CK Enerji EDAŞ Fatura Veri Portalı 1.2</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="ck-subtitle">21 dağıtım bölgesi için dinamik veri çıkarma, '
            'ETSO birleştirme ve kurumsal Excel raporu</div>',
            unsafe_allow_html=True,
        )
    with header_right:
        render_system_usage_file_picker()

    render_analysis_panel(root_folder, invoice_file, mapping_path, run_button)


if __name__ == "__main__":
    main()
