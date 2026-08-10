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
# VERSION : V1.1
# DATE    : 10.08.2026
# =========================================================

st.set_page_config(
    page_title="Uşak Eczane Haritası",
    page_icon="💊",
    layout="wide",
)

BASE_DIR = Path(__file__).resolve().parent
PREFERRED_FILE_NAME = "nöbet_ecz_liste_koordinatli.xlsx"

# =========================================================
# UŞAK 3 ANA GRUP
# =========================================================

GROUPS: dict[str, list[str]] = {
    "Grup 1": [
        "ACAR", "AHSEN", "AVGAN", "AYKANAT", "AŞİYAN",
        "BALKAN", "DEMET", "DEMİR", "DÖNMEZ", "DİDEM",
        "EBRU", "FERAH", "GÖKSEL", "GÜL", "MENDEPAZARI",
        "MERT", "SAĞLIK", "SU", "TAN", "ZAFER",
        "ÇEKİÇ", "ÖRNEK", "ÖZLEM", "İREM", "ŞAN",
    ],
    "Grup 2": [
        "AKTAY", "AYAN", "DOKUR", "EYMEN", "FARUK",
        "GÜRAN", "HAZAL", "HİLAL", "IHLAMUR", "MAYA",
        "MURAT", "MUTAFOĞLU", "NUR", "PERDAHCI", "POYRAZ",
        "SERAP", "SEVİM", "YAVUZ", "YÜKSEL", "ZÜMRÜT",
        "ÇAVUSOĞLU", "ÖZSEZER", "ÖZYAVUZ", "ÖZÇELİK",
        "İLKE", "ŞEYMA",
    ],
    "Grup 3": [
        "AKKAYA", "AYDOĞDU", "BAŞER", "BİZİM", "DURAN",
        "EGE HAYAT", "ERDEM", "EYLÜL", "FATİH", "GÜLŞİFA",
        "GÜNEŞ", "GÜVEN", "HUZUR", "KİRAZ", "MASAL DİYARI",
        "SAMANCI", "SELCEN", "SERKAN", "SÜMER", "UMUT",
        "VİTAMİN", "YAĞIZ", "YAŞAM", "YENİ ŞİFA", "YEŞİM",
        "ÖMÜR",
    ],
}

GROUP_COLORS = {
    "Grup 1": "#1976D2",   # mavi
    "Grup 2": "#2E7D32",   # yeşil
    "Grup 3": "#E65100",   # turuncu
    "Grupsuz": "#757575",  # gri
}


def normalize_text(value: object) -> str:
    """Dosya, sütun ve eczane adlarını toleranslı eşleştirir."""
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
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^0-9A-Z ]", "", text).strip()

    # "ECZANESİ" vb. ekler eşleşmeyi bozmasın.
    for suffix in [" ECZANESI", " ECZANE"]:
        if text.endswith(suffix):
            text = text[: -len(suffix)].strip()

    return text


def build_group_map() -> dict[str, str]:
    """Eczane adını ilgili 3 ana gruptan biriyle eşleştirir."""
    group_map: dict[str, str] = {}

    for group_name, pharmacy_names in GROUPS.items():
        for pharmacy_name in pharmacy_names:
            key = normalize_text(pharmacy_name)

            if key in group_map:
                raise ValueError(
                    f"'{pharmacy_name}' birden fazla grupta tanımlanmış: "
                    f"{group_map[key]} ve {group_name}"
                )

            group_map[key] = group_name

    return group_map


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
        lng_col = find_column(
            columns,
            ["Boylam (Lng)", "Boylam", "Longitude", "Lng", "Lon"],
        )

        if name_col is not None and lat_col is not None and lng_col is not None:
            selected_df = candidate.copy()
            selected_sheet = sheet_name
            break

    if selected_df is None:
        raise ValueError(
            "Excel içinde Eczane Adı, Enlem ve Boylam sütunlarını "
            "içeren uygun sayfa bulunamadı."
        )

    columns = list(selected_df.columns)
    name_col = find_column(columns, ["Eczane Adı", "Eczane", "Ad"])
    phone_col = find_column(columns, ["Telefon", "Tel", "Phone"])
    address_col = find_column(columns, ["Adres", "Address"])
    lat_col = find_column(columns, ["Enlem (Lat)", "Enlem", "Latitude", "Lat"])
    lng_col = find_column(
        columns,
        ["Boylam (Lng)", "Boylam", "Longitude", "Lng", "Lon"],
    )

    result = pd.DataFrame()
    result["Eczane"] = selected_df[name_col].astype(str).str.strip()
    result["Latitude"] = pd.to_numeric(selected_df[lat_col], errors="coerce")
    result["Longitude"] = pd.to_numeric(selected_df[lng_col], errors="coerce")

    if phone_col is not None:
        result["Telefon"] = (
            selected_df[phone_col].fillna("").astype(str).str.strip()
        )
    else:
        result["Telefon"] = ""

    if address_col is not None:
        result["Adres"] = (
            selected_df[address_col].fillna("").astype(str).str.strip()
        )
    else:
        result["Adres"] = ""

    result = result[
        result["Eczane"].ne("")
        & result["Latitude"].between(-90, 90)
        & result["Longitude"].between(-180, 180)
    ].copy()

    result["Anahtar"] = result["Eczane"].map(normalize_text)

    group_map = build_group_map()
    result["Grup"] = result["Anahtar"].map(group_map).fillna("Grupsuz")

    result = (
        result.drop_duplicates(subset=["Anahtar"], keep="first")
        .reset_index(drop=True)
    )

    result.attrs["sheet_name"] = selected_sheet

    if result.empty:
        raise ValueError(
            "Excel'de haritada gösterilebilecek geçerli eczane koordinatı bulunamadı."
        )

    return result


def add_legend(map_obj: folium.Map) -> None:
    legend_html = f"""
    <div style="
        position: fixed;
        bottom: 30px;
        left: 30px;
        z-index: 9999;
        background: white;
        border: 1px solid #bbb;
        border-radius: 8px;
        padding: 10px 14px;
        font-size: 13px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.2);
    ">
        <div style="font-weight:700; margin-bottom:6px;">Uşak Grupları</div>
        <div><span style="color:{GROUP_COLORS['Grup 1']};font-size:20px;">●</span> Grup 1</div>
        <div><span style="color:{GROUP_COLORS['Grup 2']};font-size:20px;">●</span> Grup 2</div>
        <div><span style="color:{GROUP_COLORS['Grup 3']};font-size:20px;">●</span> Grup 3</div>
        <div><span style="color:{GROUP_COLORS['Grupsuz']};font-size:20px;">●</span> Grupsuz</div>
    </div>
    """

    map_obj.get_root().html.add_child(folium.Element(legend_html))


def add_pharmacy_markers(map_obj: folium.Map, df: pd.DataFrame) -> None:
    group_layers = {
        "Grup 1": FeatureGroup(name="Grup 1", show=True),
        "Grup 2": FeatureGroup(name="Grup 2", show=True),
        "Grup 3": FeatureGroup(name="Grup 3", show=True),
        "Grupsuz": FeatureGroup(name="Grupsuz", show=True),
    }

    for _, row in df.iterrows():
        pharmacy_name = html.escape(str(row["Eczane"]))
        phone = html.escape(str(row.get("Telefon", "")))
        address = html.escape(str(row.get("Adres", "")))
        group_name = str(row.get("Grup", "Grupsuz"))
        color = GROUP_COLORS.get(group_name, GROUP_COLORS["Grupsuz"])

        tooltip_html = (
            '<div style="font-size:14px; line-height:1.35;">'
            f"<b>{pharmacy_name}</b><br>"
            f"Grup: <b>{html.escape(group_name)}</b>"
            "</div>"
        )

        popup_parts = [
            '<div style="font-family:Arial,sans-serif; '
            'font-size:13px; line-height:1.5; min-width:220px;">',
            f'<div style="font-size:15px; font-weight:700; '
            f'margin-bottom:5px;">{pharmacy_name}</div>',
            f"<div><b>Grup:</b> {html.escape(group_name)}</div>",
        ]

        if address:
            popup_parts.append(f"<div><b>Adres:</b> {address}</div>")

        if phone:
            popup_parts.append(f"<div><b>Telefon:</b> {phone}</div>")

        popup_parts.append(
            f'<div style="margin-top:5px; color:#666;">'
            f'{float(row["Latitude"]):.6f}, '
            f'{float(row["Longitude"]):.6f}'
            "</div>"
        )
        popup_parts.append("</div>")

        folium.CircleMarker(
            location=[
                float(row["Latitude"]),
                float(row["Longitude"]),
            ],
            radius=7.0,
            color="#FFFFFF",
            weight=1.7,
            fill=True,
            fill_color=color,
            fill_opacity=0.96,
            tooltip=folium.Tooltip(
                tooltip_html,
                sticky=True,
                direction="top",
            ),
            popup=folium.Popup(
                "".join(popup_parts),
                max_width=360,
            ),
        ).add_to(group_layers[group_name])

    for layer in group_layers.values():
        layer.add_to(map_obj)


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
    add_legend(map_obj)

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
            [
                float(df["Latitude"].min()),
                float(df["Longitude"].min()),
            ],
            [
                float(df["Latitude"].max()),
                float(df["Longitude"].max()),
            ],
        ],
        padding=(25, 25),
    )

    return map_obj


# =========================================================
# UYGULAMA
# =========================================================

st.title("Uşak Eczane Grup Haritası")
st.caption(
    "AYÇA — Uşak eczaneleri 3 ana gruba ayrılmıştır. "
    "Eczane adları fareyle üzerine gelince, detayları tıklayınca görünür."
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

    group_1_count = int((pharmacies["Grup"] == "Grup 1").sum())
    group_2_count = int((pharmacies["Grup"] == "Grup 2").sum())
    group_3_count = int((pharmacies["Grup"] == "Grup 3").sum())
    ungrouped_count = int((pharmacies["Grup"] == "Grupsuz").sum())

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("Toplam eczane", len(pharmacies))
    col2.metric("Grup 1", group_1_count)
    col3.metric("Grup 2", group_2_count)
    col4.metric("Grup 3", group_3_count)
    col5.metric("Grupsuz", ungrouped_count)

    with st.expander("Grup listelerini göster"):
        for group_name in ["Grup 1", "Grup 2", "Grup 3"]:
            names = (
                pharmacies.loc[
                    pharmacies["Grup"] == group_name,
                    "Eczane",
                ]
                .astype(str)
                .sort_values()
                .tolist()
            )
            st.markdown(f"**{group_name} — {len(names)} eczane**")
            st.write(", ".join(names))

        if ungrouped_count:
            ungrouped_names = (
                pharmacies.loc[
                    pharmacies["Grup"] == "Grupsuz",
                    "Eczane",
                ]
                .astype(str)
                .sort_values()
                .tolist()
            )
            st.warning(
                "Grupsuz eczaneler: "
                + ", ".join(ungrouped_names)
            )

    with st.expander("Kullanılan veri dosyası"):
        st.write(f"Dosya: `{pharmacy_path.name}`")
        st.write(
            f"Sayfa: `{pharmacies.attrs.get('sheet_name', '-')}`"
        )

    pharmacy_map = build_map(pharmacies)

    components.html(
        pharmacy_map.get_root().render(),
        height=900,
        scrolling=False,
    )

except Exception as exc:
    st.exception(exc)
