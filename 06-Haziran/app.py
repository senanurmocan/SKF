"""CK Enerji EDAŞ fatura veri çıkarma ve analiz portalı."""

from __future__ import annotations

import contextlib
import base64
import io
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





# Veri çıkarma için doğrulanmış ana çekirdek kullanılır.
from extract_and_compare import (  # noqa: E402
    clean_turkish_number,
    enrich_from_eno_file,
    normalize_region_name,
    process_all_regions,
    save_to_excel,
)


st.set_page_config(
    page_title="CK Enerji - EDAŞ Fatura Veri Portalı",
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
        [data-testid="stSidebar"] {
            background-color: #F5F7FA;
            border-right: 1px solid #D8DEE9;
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
            margin-bottom: 1.1rem;
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


def select_folder_via_dialog() -> str | None:
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
            title="21 EDAŞ klasörünün bulunduğu ana dizini seçin",
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
        "dark_mode": False,
        "processing_done": False,
        "extracted_df": None,
        "stats": {},
        "audit_results": {},
        "output_bytes": None,
        "processing_logs": [],
        "analysis_signature": None,

    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def clear_analysis_state() -> None:
    st.session_state["processing_done"] = False
    st.session_state["extracted_df"] = None
    st.session_state["stats"] = {}
    st.session_state["audit_results"] = {}
    st.session_state["output_bytes"] = None
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


def count_candidate_region_folders(root_folder: str) -> int:
    if not root_folder or not os.path.isdir(root_folder):
        return 0
    ignored = {"core", "config", "analytics", "__pycache__", ".git", ".agents"}
    return sum(
        1
        for name in os.listdir(root_folder)
        if os.path.isdir(os.path.join(root_folder, name))
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
) -> tuple[pd.DataFrame, dict[str, Any], dict[str, Any], bytes, list[str]]:
    """Doğrulanmış çekirdeği çalıştır ve maskesiz Excel'i bellekte oluştur."""
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
        enrich_from_eno_file(normalized_records, root_folder)
        output_buffer = io.BytesIO()
        save_to_excel(normalized_records, output_buffer)

    audit_results = audit_extracted_dataset(normalized_records)
    raw_columns = list(RAW_DATA_COLUMNS)
    for optional_column in OPTIONAL_RAW_DATA_COLUMNS:
        if any(optional_column in record for record in normalized_records):
            raw_columns.append(optional_column)
    raw_df = pd.DataFrame(normalized_records).reindex(columns=raw_columns)
    logs = build_processing_logs(normalized_records, stats, audit_results)
    return raw_df, stats, audit_results, output_buffer.getvalue(), logs





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


def render_sidebar() -> tuple[str, str | None]:
    logo_path = BASE_DIR / "CK_Enerji_-_Yatay.png"
    if logo_path.exists():
        st.sidebar.image(str(logo_path), width="stretch")

    st.sidebar.subheader("📂 Veri Kaynağı")
    if st.sidebar.button("📁 Klasör Seç", width="stretch"):
        selected_folder = select_folder_via_dialog()
        if selected_folder:
            clear_analysis_state()
            st.session_state["selected_root_folder"] = selected_folder
            st.rerun()
        else:
            st.sidebar.info("Klasör seçilmedi veya pencere açılamadı. Dizin yolunu aşağıdaki alana yazabilirsiniz.")

    root_folder = st.sidebar.text_input(
        "21 EDAŞ klasörünün bulunduğu ana dizin",
        key="selected_root_folder",
        placeholder=r"Örn: D:\Users\...\Desktop\SKF\06-Haziran",
    ).strip()

    # Mapping dosyası sabittir: kodlarla aynı dizindeki 'SKF Başlıkları.xlsx'
    # Kullanıcının veri klasöründe aranmaz; sistem gömülü (embedded) olarak kullanır.
    embedded_mapping_path = str(BASE_DIR / "SKF Başlıkları.xlsx")
    if not os.path.isfile(embedded_mapping_path):
        st.sidebar.error("⚠️ Sistem mapping dosyası bulunamadı: SKF Başlıkları.xlsx")

    candidate_count = count_candidate_region_folders(root_folder)
    st.sidebar.info(
        f"Bulunan aday bölge klasörü: {format_turkish_integer(candidate_count)} / "
        f"{format_turkish_integer(EXPECTED_REGION_COUNT)}"
    )
    st.sidebar.success("Verileriniz %100 yerelde işlenmektedir, ETSO kodları gizlidir.")



    # Karanlık / Aydınlık mod
    st.sidebar.divider()
    st.sidebar.toggle("🌙 Karanlık Mod", key="dark_mode")

    mapping_path = embedded_mapping_path if os.path.isfile(embedded_mapping_path) else None
    return root_folder, mapping_path


def render_analysis_panel(root_folder: str, mapping_path: str | None) -> None:
    st.subheader("⚙️ Veri İşleme ve Kurumsal Rapor")

    valid_root = bool(root_folder and os.path.isdir(root_folder))
    valid_mapping = bool(mapping_path and os.path.isfile(mapping_path))
    if not root_folder:
        st.info("Analize başlamak için soldaki Klasör Seç düğmesini kullanın.")
    elif not valid_root:
        st.error("Seçilen ana klasör bulunamadı.")
    elif not valid_mapping:
        st.error("Seçilen klasörde uygun SKF mapping dosyası bulunamadı.")

    current_signature = (
        os.path.normcase(os.path.abspath(root_folder)) if valid_root else "",
        os.path.normcase(os.path.abspath(mapping_path)) if valid_mapping and mapping_path else "",
    )
    if (
        st.session_state.get("processing_done")
        and st.session_state.get("analysis_signature") != current_signature
    ):
        clear_analysis_state()

    run_button = st.button(
        "🚀 Verileri Birleştir ve Analiz Et",
        type="primary",
        width="stretch",
        disabled=not (valid_root and valid_mapping),
    )

    if run_button and mapping_path:
        clear_analysis_state()
        progress_bar = st.progress(0)
        status_placeholder = st.empty()

        def update_progress(message: str, percent: int) -> None:
            safe_percent = min(100, max(0, int(percent)))
            progress_bar.progress(safe_percent)
            status_placeholder.info(message)

        try:
            with st.spinner("Fatura dosyaları okunuyor ve ETSO kayıtları birleştiriliyor."):
                raw_df, stats, audit_results, output_bytes, logs = run_analysis_pipeline(
                    root_folder,
                    mapping_path,
                    update_progress,
                )
            st.session_state["extracted_df"] = raw_df
            st.session_state["stats"] = stats
            st.session_state["audit_results"] = audit_results
            st.session_state["output_bytes"] = output_bytes
            st.session_state["processing_logs"] = logs
            st.session_state["analysis_signature"] = current_signature
            st.session_state["processing_done"] = True
            progress_bar.progress(100)
            status_placeholder.success("Analiz tamamlandı.")
        except Exception as error:
            clear_analysis_state()
            status_placeholder.empty()
            progress_bar.empty()
            st.error(f"İşlem tamamlanamadı: {safe_error_text(error)}")

    if not st.session_state.get("processing_done"):
        return

    raw_df = st.session_state.get("extracted_df")
    if raw_df is None:
        raw_df = pd.DataFrame()
    stats = st.session_state.get("stats") or {}
    audit_results = st.session_state.get("audit_results") or {}
    output_bytes = st.session_state.get("output_bytes")

    read_count = int(stats.get("read_region_count", 0))
    expected_count = int(stats.get("expected_region_count", EXPECTED_REGION_COUNT))
    missing_regions = list(stats.get("missing_regions", []))
    if read_count == expected_count and not missing_regions:
        st.success(
            f"{format_turkish_integer(expected_count)} bölgenin "
            f"{format_turkish_integer(read_count)}'i başarıyla okundu."
        )
    else:
        st.warning("Şu bölgeler okunamadı: " + ", ".join(missing_regions))

    records = raw_df.to_dict(orient="records")
    total_amount = total_distribution_amount(records)
    metric_columns = st.columns(4)
    render_metric_card(metric_columns[0], "Toplam Kayıt", format_turkish_integer(len(raw_df)))
    render_metric_card(metric_columns[1], "Toplam Dağıtım Bedeli", format_turkish_currency(total_amount))
    render_metric_card(
        metric_columns[2],
        "Okunan Bölge",
        f"{format_turkish_integer(read_count)} / {format_turkish_integer(expected_count)}",
    )
    render_metric_card(
        metric_columns[3],
        "10 Milyon TL Üzeri Anomali",
        format_turkish_integer(audit_results.get("high_value_count", 0)),
    )

    st.caption(
        f"Veri sağlık skoru: %{format_turkish_integer(audit_results.get('health_score', 100))} | "
        f"Birleştirilen mükerrer kayıt: {format_turkish_integer(stats.get('merged_duplicates', 0))} | "
        f"Filtrelenen bozuk kayıt: {format_turkish_integer(stats.get('absurd_filtered', 0))}"
    )

    source_header_warnings = stats.get("source_header_warnings", [])
    if source_header_warnings:
        st.warning(
            f"{format_turkish_integer(len(source_header_warnings))} kaynak başlık uyuşmazlığı "
            "veya format değişikliği bulundu. "
            "Ayrıntıları kontrol edin."
        )
        with st.expander("⚠️ Kaynak başlık / format uyarıları", expanded=True):
            st.dataframe(
                pd.DataFrame(source_header_warnings),
                width="stretch",
                hide_index=True,
                height=min(620, max(180, 38 * min(len(source_header_warnings), 16))),
            )

    with st.expander("📋 İşlem Logları", expanded=False):
        for log_line in st.session_state.get("processing_logs", []):
            st.code(log_line, language=None)

    st.markdown("#### 📥 Maskesiz Orijinal Excel Çıktısı")
    st.caption(
        "Ekrandaki ETSO kodları maskelidir. İndirilen Excel, kaynak ETSO kodlarını "
        "değiştirmeden korur ve sütun genişliklerini içeriğe göre otomatik ayarlar."
    )
    if output_bytes:
        st.download_button(
            "📥 Çıkarılan_Veriler.xlsx Dosyasını İndir",
            data=output_bytes,
            file_name="Çıkarılan_Veriler.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
            on_click="ignore",
        )

    st.markdown("#### 🚨 Anomali ve Veri Kalitesi Paneli")
    st.caption("10.000.000,00 TL altındaki hiçbir fatura tutar anomalisi olarak gösterilmez.")
    warnings_list = audit_results.get("warnings", [])
    monetary_warnings = [
        warning for warning in warnings_list if warning.get("category") in {"Absürt Tutar", "Yüksek Tutar"}
    ]
    quality_warnings = [warning for warning in warnings_list if warning not in monetary_warnings]

    if monetary_warnings:
        for warning in monetary_warnings:
            if warning.get("type") == "DANGER":
                st.error(warning.get("message", ""))
            else:
                st.warning(warning.get("message", ""))
    else:
        st.success("10.000.000,00 TL ve üzeri istatistiksel tutar anomalisi tespit edilmedi.")

    for warning in quality_warnings:
        st.info(warning.get("message", ""))

    st.markdown(
        f"#### 🔍 Veri Önizleme - {format_turkish_integer(len(raw_df))} Kayıt"
    )
    st.success("Verileriniz %100 yerelde işlenmektedir, ETSO kodları gizlidir.")
    display_df = format_dataframe_for_ui(raw_df)
    st.dataframe(
        display_df,
        width="stretch",
        hide_index=True,
        height=620,
    )



def main() -> None:
    inject_copy_shortcut_guard()
    inject_corporate_css()
    initialize_session_state()

    root_folder, mapping_path = render_sidebar()

    header_left, header_right = st.columns([4, 1])
    with header_left:
        st.markdown('<div class="ck-title">⚡ CK Enerji EDAŞ Fatura Veri Portalı</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="ck-subtitle">21 dağıtım bölgesi için dinamik veri çıkarma, '
            'ETSO birleştirme, anomali denetimi ve kurumsal Excel raporu</div>',
            unsafe_allow_html=True,
        )
    with header_right:
        logo_path = BASE_DIR / "CK_Enerji_-_Yatay.png"
        if logo_path.exists():
            st.image(str(logo_path), width="stretch")

    st.success("Verileriniz %100 yerelde işlenmektedir, ETSO kodları gizlidir.")

    render_analysis_panel(root_folder, mapping_path)


if __name__ == "__main__":
    main()
