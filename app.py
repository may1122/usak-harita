from __future__ import annotations

import html
import re
import unicodedata
from pathlib import Path

import folium
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from folium import FeatureGroup
from folium.plugins import Fullscreen, MeasureControl

# =========================================================
# AYÇA UŞAK ECZANE HARİTASI
# VERSION : V1.0
# DATE    : 10.08.2026
# =========================================================

st.set_page_config(
    page_title="Uşak Eczane Haritası",
    page_icon="💊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
PREFERRED_FILE_NAME = "nöbet_ecz_liste_koordinatli.xlsx"


def normalize_text(value: object) -> str:
    """Dosya ve sütun adlarını toleranslı eşleştirmek için normalize eder."""
    if pd.isna(value):
        return ""

    text = unicodedata.normalize("NFKC", str(value)).strip().upper()
    text = text.translate(
        str.maketrans(
            {
                "Ç": "C",
                "Ğ": "G",
                "İ": "I",
                "Ö": "O",
                "Ş": "S",
                "Ü": "U",
            }
        )
    )
    return re.sub(r"[^0-9A-Z]", "", text)


def locate_pharmacy_file() -> Path | None:
    """Önce tercih edilen adı, sonra uygun koordinat Excel'ini otomatik bulur."""
    preferred = BASE_DIR / PREFERRED_FILE_NAME
    if preferred.exists():
        return preferred

    candidates = sorted(BASE_DIR.glob("*.xlsx"))
    for candidate in candidates:
        key = normalize_text(candidate.name)
        if "NOBET" in key and "KOORDINAT" in key:
            return candidate

    for candidate in candidates:
        key = normalize_text(candidate.name)
        if "ECZ" in key or "ECZANE" in key:
            return candidate

    return candidates[0] if candidates else None


def find_column(columns: list[object], aliases: list[str]) -> object | None:
    """Sütun adını farklı yazım ihtimallerine rağmen bulur."""
    normalized = {normalize_text(col): col for col in columns}

    for alias in aliases:
        alias_key = normalize_text(alias)
        if alias_key in normalized:
            return normalized[alias_key]

    for alias in aliases:
        alias_key = normalize_text(alias)
        for key, original in normalized.items():
            if alias_key and alias_key in key:
                return original

    return None


@st.cache_data(show_spinner=False)
def read_pharmacies(path: str, file_version: int) -> pd.DataFrame:
    """Uşak koordinat Excel'ini okuyup standart sütun yapısına çevirir."""
    del file_version

    excel = pd.ExcelFile(path, engine="openpyxl")

    selected_df: pd.DataFrame | None = None
    selected_sheet = ""

    for sheet_name in excel.sheet_names:
        candidate = pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")
        columns = list(candidate.columns)

        name_col = find_column(columns, ["Eczane Adı", "Eczane", "Ad"])
        lat_col = find_column(columns, ["Enlem (Lat)", "Enlem", "Latitude", "Lat"])
        lng_col = find_column(columns, ["Boylam (Lng)", "Boylam", "Longitude", "Lng", "Lon"])

        if name_col is not None and lat_col is not None and lng_col is not None:
            selected_df = candidate.copy()
            selected_sheet = sheet_name
            break

    if selected_df is None:
        raise ValueError(
            "Excel içinde Eczane Adı, Enlem ve Boylam sütunlarını içeren uygun sayfa bulunamadı."
        )

    columns = list(selected_df.columns)
    name_col = find_column(columns, ["Eczane Adı", "Eczane", "Ad"])
    phone_col = find_column(columns, ["Telefon", "Tel", "Phone"])
    address_col = find_column(columns, ["Adres", "Address"])
    lat_col = find_column(columns, ["Enlem (Lat)", "Enlem", "Latitude", "Lat"])
    lng_col = find_column(columns, ["Boylam (Lng)", "Boylam", "Longitude", "Lng", "Lon"])

    result = pd.DataFrame()
    result["Eczane"] = selected_df[name_col].astype(str).str.strip()
    result["Latitude"] = pd.to_numeric(selected_df[lat_col], errors="coerce")
    result["Longitude"] = pd.to_numeric(selected_df[lng_col], errors="coerce")

    if phone_col is not None:
        result["Telefon"] = selected_df[phone_col].fillna("").astype(str).str.strip()
    else:
        result["Telefon"] = ""

    if address_col is not None:
        result["Adres"] = selected_df[address_col].fillna("").astype(str).str.strip()
    else:
        result["Adres"] = ""

    result = result[
        result["Eczane"].ne("")
        & result["Latitude"].between(-90, 90)
        & result["Longitude"].between(-180, 180)
    ].copy()

    result["Anahtar"] = result["Eczane"].map(normalize_text)
    result = result.drop_duplicates(subset=["Anahtar"], keep="first").reset_index(drop=True)
    result.attrs["sheet_name"] = selected_sheet

    if result.empty:
        raise ValueError("Excel'de haritada gösterilebilecek geçerli eczane koordinatı bulunamadı.")

    return result


def add_pharmacy_markers(map_obj: folium.Map, df: pd.DataFrame) -> None:
    pharmacy_layer = FeatureGroup(name="Eczaneler", show=True)

    for _, row in df.iterrows():
        pharmacy_name = html.escape(str(row["Eczane"]))
        phone = html.escape(str(row.get("Telefon", "")))
        address = html.escape(str(row.get("Adres", "")))

        tooltip_html = (
            '<div style="font-size:14px; line-height:1.35;">'
            f"<b>{pharmacy_name}</b>"
            "</div>"
        )

        popup_parts = [
            '<div style="font-family:Arial,sans-serif; font-size:13px; line-height:1.5; min-width:220px;">',
            f'<div style="font-size:15px; font-weight:700; margin-bottom:5px;">{pharmacy_name}</div>',
        ]

        if address:
            popup_parts.append(f"<div><b>Adres:</b> {address}</div>")
        if phone:
            popup_parts.append(f"<div><b>Telefon:</b> {phone}</div>")

        popup_parts.append(
            f'<div style="margin-top:5px; color:#666;">'
            f'{float(row["Latitude"]):.6f}, {float(row["Longitude"]):.6f}'
            "</div>"
        )
        popup_parts.append("</div>")

        folium.CircleMarker(
            location=[float(row["Latitude"]), float(row["Longitude"])],
            radius=6.0,
            color="#FFFFFF",
            weight=1.6,
            fill=True,
            fill_color="#1976D2",
            fill_opacity=0.96,
            tooltip=folium.Tooltip(
                tooltip_html,
                sticky=True,
                direction="top",
            ),
            popup=folium.Popup("".join(popup_parts), max_width=360),
        ).add_to(pharmacy_layer)

    pharmacy_layer.add_to(map_obj)


def build_map(df: pd.DataFrame) -> folium.Map:
    center = [
        float(df["Latitude"].median()),
        float(df["Longitude"].median()),
    ]

    map_obj = folium.Map(
        location=center,
        zoom_start=14,
        tiles=None,
        control_scale=True,
        prefer_canvas=True,
    )

    folium.TileLayer(
        tiles="CartoDB positron",
        name="Sade harita",
        control=True,
        show=True,
    ).add_to(map_obj)

    folium.TileLayer(
        tiles="OpenStreetMap",
        name="Detaylı harita",
        control=True,
        show=False,
    ).add_to(map_obj)

    add_pharmacy_markers(map_obj, df)

    Fullscreen(
        position="topright",
        title="Tam ekran",
        title_cancel="Tam ekrandan çık",
    ).add_to(map_obj)

    MeasureControl(
        position="topright",
        primary_length_unit="meters",
    ).add_to(map_obj)

    folium.LayerControl(
        collapsed=False,
        position="topright",
    ).add_to(map_obj)

    map_obj.fit_bounds(
        [
            [float(df["Latitude"].min()), float(df["Longitude"].min())],
            [float(df["Latitude"].max()), float(df["Longitude"].max())],
        ],
        padding=(25, 25),
    )

    return map_obj


# =========================================================
# UYGULAMA
# =========================================================
st.title("Uşak Eczane Haritası")
st.caption(
    "AYÇA — Eczane adları fareyle üzerine gelince görünür. "
    "Eczaneye tıklayınca adres, telefon ve koordinat bilgisi açılır."
)

pharmacy_path = locate_pharmacy_file()

if pharmacy_path is None:
    st.error(
        "Koordinat Excel dosyası GitHub reposunda bulunamadı. "
        "app.py ile aynı klasöre Uşak koordinat Excel dosyasını yükleyin."
    )
    st.stop()

try:
    pharmacies = read_pharmacies(
        str(pharmacy_path),
        pharmacy_path.stat().st_mtime_ns,
    )

    col1, col2, col3 = st.columns(3)
    col1.metric("Toplam eczane", len(pharmacies))
    col2.metric("Koordinatlı", len(pharmacies))
    col3.metric("Eksik koordinat", 0)

    with st.expander("Kullanılan veri dosyası"):
        st.write(f"Dosya: `{pharmacy_path.name}`")
        st.write(f"Sayfa: `{pharmacies.attrs.get('sheet_name', '-')}`")

    pharmacy_map = build_map(pharmacies)
    components.html(
        pharmacy_map.get_root().render(),
        height=900,
        scrolling=False,
    )

except Exception as exc:
    st.exception(exc)
