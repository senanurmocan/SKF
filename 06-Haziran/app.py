"""CK Enerji EDAŞ fatura veri çıkarma ve analiz portalı."""

from __future__ import annotations

import contextlib
import base64
import html
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
            margin-top: -0.6rem;
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
        "processing_logs": [],
        "analysis_signature": None,

    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    if "selected_system_usage_file" not in st.session_state:
        st.session_state["selected_system_usage_file"] = load_system_usage_file_selection()


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


def clear_analysis_state() -> None:
    st.session_state["processing_done"] = False
    st.session_state["extracted_df"] = None
    st.session_state["stats"] = {}
    st.session_state["audit_results"] = {}
    st.session_state["output_bytes"] = None
    st.session_state["invoice_sources"] = []
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
    for row_index in range(2, worksheet.max_row + 1):
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
            output_cell._style = copy(source_amount_cell._style)
        output_cell.number_format = source_amount_cell.number_format or "#,##0.00"
        if rate is None:
            output_cell.value = None
            missing_rate_count += 1
            continue

        active_letter = get_column_letter(active_column)
        output_cell.value = f"=IFERROR({active_letter}{row_index}*{rate!r}/100,0)"
        calculated_count += 1

    return {
        "calculated_distribution_amount_count": calculated_count,
        "missing_distribution_rate_count": missing_rate_count,
        "distribution_amount_warning": None,
    }


def append_selected_invoice_file(
    output_bytes: bytes,
    file_path: str | None,
    mapping_path: str | None = None,
    system_usage_file: str | None = None,
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
                    **calculation_stats,
                    **comparison_stats,
                    "tariff_rate_error": tariff_rate_error,
                })
        else:
            import xlrd

            source_workbook = xlrd.open_workbook(file_path, on_demand=True)
            for sheet_number, source_sheet in enumerate(source_workbook.sheets(), start=1):
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
                    **calculation_stats,
                    **comparison_stats,
                    "tariff_rate_error": tariff_rate_error,
                })

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
        for destination_column, source_column in source_column_map.items():
            if destination_column == 10:
                tariff_value = source_worksheet.cell(
                    row_index,
                    source_column_map[4],
                ).value
                energy_source_column = value_source_columns.get(9) or source_column_map[9]
                energy_source = source_values_worksheet or source_worksheet
                energy_value = energy_source.cell(row_index, energy_source_column).value
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
                value_source = source_values_worksheet or source_worksheet
                value = value_source.cell(row_index, value_source_column).value
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
    for row_index in range(2, last_row + 1):
        for column in range(1, 20):
            worksheet.cell(row_index, column).alignment = Alignment(vertical="center")
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
) -> tuple[pd.DataFrame, dict[str, Any], dict[str, Any], bytes, list[str], list[dict[str, Any]]]:
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
        output_bytes = add_summary_worksheet(output_buffer.getvalue(), normalized_records)
        output_bytes, invoice_sources, invoice_file_errors, invoice_file_count = append_selected_invoice_file(
            output_bytes,
            invoice_file,
            mapping_path=mapping_path,
            system_usage_file=system_usage_file,
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
        if missing_rate_count:
            logs.append(
                "Tarife Grubu için birim fiyat bulunamadığından hesaplanamayan satır: "
                f"{format_turkish_integer(missing_rate_count)}."
            )
        for source in invoice_sources:
            warning = (
                source.get("comparison_sheet_warning")
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
            source_column = source_columns[source_field]
            source_letter = detail_sheet.cell(1, source_column).column_letter
            sum_range = f"{detail_name}!${source_letter}$2:${source_letter}${last_detail_row}"
            cell = summary_sheet.cell(row_index, column_index)
            if is_unknown_company:
                cell.value = (
                    f'=SUMIF({company_range},"",{sum_range})'
                    f'+SUMIF({company_range},"nan",{sum_range})'
                )
            else:
                cell.value = f"=SUMIF({company_range},$A{row_index},{sum_range})"
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


def render_sidebar() -> tuple[str, str, str | None, bool]:
    logo_path = BASE_DIR / "CK_Enerji_-_Yatay.png"
    if logo_path.exists():
        st.sidebar.image(str(logo_path), width="stretch")

    st.sidebar.subheader("📂 Veri Kaynağı")
    if st.sidebar.button("📁 Dağıtım Şirketi SKF", width="stretch", key="select_distribution_folder"):
        selected_folder = select_folder_via_dialog(
            "Dağıtım şirketi SKF klasörlerinin bulunduğu ana dizini seçin"
        )
        if selected_folder:
            clear_analysis_state()
            st.session_state["selected_root_folder"] = selected_folder
            st.rerun()
        else:
            st.sidebar.info("Klasör seçilmedi veya pencere açılamadı. SKF dizin yolunu aşağıya yazabilirsiniz.")

    root_folder = st.sidebar.text_input(
        "Dağıtım Şirketi SKF klasör yolu",
        key="selected_root_folder",
        placeholder=r"Örn: D:\Users\...\Desktop\SKF\06-Haziran",
    ).strip()

    if st.sidebar.button(
        "📄 Bizim Kestiğimiz Faturalar",
        width="stretch",
        key="select_our_invoice_file",
    ):
        selected_file = select_excel_file_via_dialog(
            "Bizim Kestiğimiz Faturalar için Excel dosyasını seçin"
        )
        if selected_file:
            clear_analysis_state()
            st.session_state["selected_invoice_file"] = selected_file
            st.rerun()
        else:
            st.sidebar.info("Excel dosyası seçilmedi veya dosya penceresi açılamadı. Dosya yolunu aşağıya yazabilirsiniz.")

    invoice_file = st.sidebar.text_input(
        "Bizim Kestiğimiz Faturalar Excel dosyası",
        key="selected_invoice_file",
        placeholder=r"Örn: D:\Users\...\Desktop\Faturalar.xlsx",
    ).strip()

    # Mapping dosyası sabittir: kodlarla aynı dizindeki 'SKF Başlıkları.xlsx'
    # Kullanıcının veri klasöründe aranmaz; sistem gömülü (embedded) olarak kullanır.
    embedded_mapping_path = str(BASE_DIR / "SKF Başlıkları.xlsx")
    if not os.path.isfile(embedded_mapping_path):
        st.sidebar.error("⚠️ Sistem mapping dosyası bulunamadı: SKF Başlıkları.xlsx")
    mapping_path = embedded_mapping_path if os.path.isfile(embedded_mapping_path) else None
    valid_invoice_file = not invoice_file or os.path.isfile(invoice_file)
    if invoice_file and not valid_invoice_file:
        st.sidebar.error("Seçilen Bizim Kestiğimiz Faturalar Excel dosyası bulunamadı.")

    run_button = st.sidebar.button(
        "🚀 Verileri Birleştir ve Karşılaştır",
        type="primary",
        width="stretch",
        disabled=not (
            root_folder
            and os.path.isdir(root_folder)
            and mapping_path
            and valid_invoice_file
        ),
    )

    candidate_count = count_candidate_region_folders(root_folder)
    st.sidebar.info(
        f"Bulunan aday bölge klasörü: {format_turkish_integer(candidate_count)} / "
        f"{format_turkish_integer(EXPECTED_REGION_COUNT)}"
    )
    st.sidebar.success("Verileriniz %100 yerelde işlenmektedir, ETSO kodları gizlidir.")



    # Karanlık / Aydınlık mod
    st.sidebar.divider()
    st.sidebar.toggle("🌙 Karanlık Mod", key="dark_mode")

    return root_folder, invoice_file, mapping_path, run_button


def render_system_usage_file_picker() -> None:
    selected_file = st.session_state.get("selected_system_usage_file", "")
    if st.button(
        "📄 Sistem Kullanım Bedelleri",
        width="stretch",
        key="select_system_usage_file",
        help="Sistem kullanım bedelleri Excel dosyasını seçin.",
    ):
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
                        "Dosya bu oturum için seçildi ancak sonraki açılışlarda hatırlanamadı: "
                        f"{save_error}"
                    )
                else:
                    st.rerun()
        else:
            st.info("Dosya seçilmedi veya dosya penceresi açılamadı.")

    selected_file = st.session_state.get("selected_system_usage_file", "")
    if selected_file and os.path.isfile(selected_file):
        st.caption(f"Seçili: {Path(selected_file).name}")
    else:
        st.warning(
            "Sistem Kullanım Bedelleri dosyası bulunamadı. Butona basıp geçerli dosyayı seçin.",
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
    if not root_folder:
        st.info("Analize başlamak için soldan Dağıtım Şirketi SKF klasörünü seçin.")
    elif not valid_root:
        st.error("Seçilen Dağıtım Şirketi SKF klasörü bulunamadı.")
    if invoice_file and not valid_invoice_file:
        st.error("Seçilen Bizim Kestiğimiz Faturalar Excel dosyası bulunamadı.")
    elif not valid_mapping:
        st.error("Seçilen klasörde uygun SKF mapping dosyası bulunamadı.")

    current_signature = (
        os.path.normcase(os.path.abspath(root_folder)) if valid_root else "",
        os.path.normcase(os.path.abspath(invoice_file)) if invoice_file and valid_invoice_file else "",
        os.path.normcase(os.path.abspath(mapping_path)) if valid_mapping and mapping_path else "",
        os.path.normcase(os.path.abspath(st.session_state.get("selected_system_usage_file", "")))
        if st.session_state.get("selected_system_usage_file")
        and os.path.isfile(st.session_state.get("selected_system_usage_file", ""))
        else "",
    )
    if (
        st.session_state.get("processing_done")
        and st.session_state.get("analysis_signature") != current_signature
    ):
        clear_analysis_state()

    if run_button and mapping_path and valid_invoice_file:
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
                )
            st.session_state["extracted_df"] = raw_df
            st.session_state["stats"] = stats
            st.session_state["audit_results"] = audit_results
            st.session_state["output_bytes"] = output_bytes
            st.session_state["invoice_sources"] = invoice_sources
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
    for invoice_error in stats.get("invoice_file_errors", []):
        st.warning(f"{OUR_INVOICES_LABEL}: {invoice_error}")

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

    preview_title, download_column = st.columns([5, 2], gap="small", vertical_alignment="center")
    with preview_title:
        st.markdown(
            f"#### 🔍 Veri Önizleme - {format_turkish_integer(len(raw_df))} Kayıt"
        )
    with download_column:
        if output_bytes:
            selected_folder_name = Path(os.path.normpath(root_folder)).name or "SKF"
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

    invoice_sources = st.session_state.get("invoice_sources", [])
    if invoice_sources:
        st.markdown(f"#### {OUR_INVOICES_LABEL}")
        st.caption("Seçilen Excel sayfaları çıktı dosyasına eklendi; her kaynak sayfa için Bizim Kestiğimiz 2 karşılaştırma sekmesi oluşturulur.")
        for source in invoice_sources:
            st.caption(
                f"{Path(source['file_path']).name} / {source['source_sheet_name']} "
                f"→ {source['output_sheet_name']}"
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
        st.markdown('<div class="ck-title">⚡ CK Enerji EDAŞ Fatura Veri Portalı</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="ck-subtitle">21 dağıtım bölgesi için dinamik veri çıkarma, '
            'ETSO birleştirme, anomali denetimi ve kurumsal Excel raporu</div>',
            unsafe_allow_html=True,
        )
    with header_right:
        render_system_usage_file_picker()
        logo_path = BASE_DIR / "CK_Enerji_-_Yatay.png"
        if logo_path.exists():
            st.image(str(logo_path), width="stretch")

    render_analysis_panel(root_folder, invoice_file, mapping_path, run_button)


if __name__ == "__main__":
    main()
