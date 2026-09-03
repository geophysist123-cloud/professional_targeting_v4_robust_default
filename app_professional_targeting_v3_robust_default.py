# ============================================================
# EGYPT IMW / SGRID GEOPHYSICAL EXPLORATION DASHBOARD - app_m4.py
# ============================================================
# Complete Streamlit application
#
# Main capabilities:
#   1) Egypt IMW / SGrid calculation
#   2) Verified control point: 25°N, 34°E -> NG-36 SE G3
#   3) Egypt 1907 / TM projected coordinates
#   4) SQLite database with automatic schema repair
#   5) Excel / CSV import - processes ALL rows
#   6) Manual DMS input
#   7) Geographic / Satellite / Terrain basemaps
#   8) IMW / SGrid grid on map
#   9) Magnetic / Gravity analysis
#  10) Professional magnetic + gravity target scoring
#  11) Robust Z-score (recommended/default) / Percentile / Winsorized Min-Max normalization
#  12) Magnetic / Gravity weights + data confidence
#  13) Magnetic / Gravity contributions + concordance
#  14) Target score + priority ranking
#  11) Heatmap visualization
#  12) IMW filtering
#  13) Multi-format export:
#        CSV, Excel, JSON, GeoJSON, KML
#  14) Optional Shapefile export when geopandas is installed
#  15) Optional PDF report when reportlab is installed
#
# Install:
#   pip install streamlit pandas openpyxl folium streamlit-folium pyproj
#
# Optional:
#   pip install geopandas reportlab
#
# Run:
#   streamlit run app_m4.py
# ============================================================

import io
import json
import math
import sqlite3
from datetime import datetime
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape

import folium  # pyright: ignore[reportMissingImports]
import pandas as pd
import streamlit as st
from folium.plugins import HeatMap  # pyright: ignore[reportMissingImports]
from pyproj import Transformer  # pyright: ignore[reportMissingImports]
from streamlit_folium import st_folium  # pyright: ignore[reportMissingImports]

# ============================================================
# 1. STREAMLIT PAGE
# ============================================================

st.set_page_config(
    page_title="Egypt Geophysical Exploration Dashboard",
    page_icon="🧲",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
.hero {padding: 1.25rem 1.5rem; border-radius: 18px; background: linear-gradient(135deg,#0b1f33,#164e63); color: white; margin-bottom: 1rem;}
.hero h1 {margin: 0; font-size: 2.1rem;}
.hero p {margin: .35rem 0 0; opacity: .9;}
.section-card {padding: 1rem 1.1rem; border: 1px solid rgba(128,128,128,.25); border-radius: 14px; margin-bottom: 1rem;}
.small-note {font-size: .85rem; opacity: .75;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>🧲 Egypt Geophysical Exploration & Targeting Dashboard</h1>
  <p>IMW / SGrid • Egypt 1907 / TM • Magnetic • Gravity • GIS • Exploration Targeting</p>
</div>
""", unsafe_allow_html=True)
st.subheader("Egypt IMW / SGrid Geophysical Exploration Dashboard")
st.caption(
    "IMW / SGrid • Egypt 1907 / TM • Magnetic • Gravity • GIS • Exploration Targets"
)


# ============================================================
# 2. PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DB_NAME = BASE_DIR / "mining_geophysics.db"


# ============================================================
# 3. DATABASE SCHEMA
# ============================================================

DATABASE_COLUMNS = {
    "Well_Name": "TEXT",
    "Lat": "REAL",
    "Lon": "REAL",
    "Easting": "REAL",
    "Northing": "REAL",
    "Projection_EPSG": "INTEGER",
    "IMW_1M": "TEXT",
    "IMW_QUADRANT": "TEXT",
    "IMW_250K": "TEXT",
    "IMW_100K": "TEXT",
    "Correct_IMW_Code": "TEXT",
    "Mag_Anomaly": "REAL",
    "Grav_Anomaly": "REAL",
    "Mag_Normalized": "REAL",
    "Grav_Normalized": "REAL",
    "Mag_Score": "REAL",
    "Grav_Score": "REAL",
    "Mag_Contribution": "REAL",
    "Grav_Contribution": "REAL",
    "Data_Confidence": "REAL",
    "Concordance": "REAL",
    "Target_Score": "REAL",
    "Target_Priority": "TEXT",
    "Normalization_Method": "TEXT",
    "Created_At": "TEXT",
}


def get_connection():
    conn = sqlite3.connect(str(DB_NAME))
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def initialize_database():
    """
    Creates the database if necessary and automatically adds any
    missing columns to old versions of geophysical_wells.

    This specifically fixes:
        no such column: Projection_EPSG
    """
    conn = get_connection()

    try:
        cursor = conn.cursor()

        cursor.execute(
            """
            CREATE TABLE IF NOT EXISTS geophysical_wells (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                Well_Name TEXT,
                Lat REAL,
                Lon REAL,
                Easting REAL,
                Northing REAL,
                Projection_EPSG INTEGER,
                IMW_1M TEXT,
                IMW_QUADRANT TEXT,
                IMW_250K TEXT,
                IMW_100K TEXT,
                Correct_IMW_Code TEXT,
                Mag_Anomaly REAL,
                Grav_Anomaly REAL,
                Mag_Normalized REAL,
                Grav_Normalized REAL,
                Mag_Score REAL,
                Grav_Score REAL,
                Mag_Contribution REAL,
                Grav_Contribution REAL,
                Data_Confidence REAL,
                Concordance REAL,
                Target_Score REAL,
                Target_Priority TEXT,
                Normalization_Method TEXT,
                Created_At TEXT
            )
            """
        )

        conn.commit()

        cursor.execute("PRAGMA table_info(geophysical_wells)")
        existing = {str(row[1]).lower() for row in cursor.fetchall()}

        for column_name, column_type in DATABASE_COLUMNS.items():
            if column_name.lower() not in existing:
                cursor.execute(
                    f'ALTER TABLE geophysical_wells '
                    f'ADD COLUMN "{column_name}" {column_type}'
                )

        conn.commit()

    finally:
        conn.close()


initialize_database()


# ============================================================
# 4. EGYPT 1907 / TRANSVERSE MERCATOR
# ============================================================

def get_projection_epsg(lat, lon):
    """
    Selects the Egypt 1907 TM belt used by the application.

    EPSG:
        22991 = Egypt 1907 / Blue Belt
        22992 = Egypt 1907 / Red Belt
        22993 = Egypt 1907 / Purple Belt
        22994 = Egypt 1907 / Extended Purple Belt
    """
    lat = float(lat)
    lon = float(lon)

    if lon >= 32.0:
        return 22991

    if 29.0 <= lon < 33.0:
        return 22992

    if lon < 29.0:
        if lat >= 28.183333:
            return 22993
        return 22994

    return 22992


@st.cache_resource
def get_transformer(epsg):
    return Transformer.from_crs(
        4326,
        int(epsg),
        always_xy=True,
    )


def latlon_to_projected(lat, lon):
    epsg = get_projection_epsg(lat, lon)
    transformer = get_transformer(epsg)

    easting, northing = transformer.transform(
        float(lon),
        float(lat),
    )

    return float(easting), float(northing), int(epsg)


# ============================================================
# 5. DMS
# ============================================================

def dms_to_dd(degrees, minutes, seconds):
    degrees = float(degrees)
    minutes = float(minutes)
    seconds = float(seconds)

    sign = -1 if degrees < 0 else 1

    return sign * (
        abs(degrees)
        + minutes / 60.0
        + seconds / 3600.0
    )


# ============================================================
# 6. IMW 1M SHEETS
# ============================================================

IMW_SHEETS = [
    {
        "prefix": "NF",
        "lat_min": 20.0,
        "lat_max": 24.0,
    },
    {
        "prefix": "NG",
        "lat_min": 24.0,
        "lat_max": 28.0,
    },
    {
        "prefix": "NH",
        "lat_min": 28.0,
        "lat_max": 32.0,
    },
]

IMW_LONGITUDE_SHEETS = [
    {
        "number": 35,
        "lon_min": 24.0,
        "lon_max": 30.0,
    },
    {
        "number": 36,
        "lon_min": 30.0,
        "lon_max": 36.0,
    },
]


# ============================================================
# 7. LETTER GRID
# ============================================================

# SOUTH -> NORTH
#
# M N O P
# I J K L
# E F G H
# A B C D
#
# Each letter = 1.5° longitude x 1° latitude.

LETTER_GRID = [
    ["A", "B", "C", "D"],
    ["E", "F", "G", "H"],
    ["I", "J", "K", "L"],
    ["M", "N", "O", "P"],
]


# ============================================================
# 8. FIND 1M SHEET
# ============================================================

def get_1m_sheet(lat, lon):
    lat = float(lat)
    lon = float(lon)

    selected_lat_band = None
    selected_lon_band = None

    for band in IMW_SHEETS:
        if band["lat_min"] <= lat < band["lat_max"]:
            selected_lat_band = band
            break

    for band in IMW_LONGITUDE_SHEETS:
        if band["lon_min"] <= lon < band["lon_max"]:
            selected_lon_band = band
            break

    if selected_lat_band is None or selected_lon_band is None:
        return None

    sheet_code = (
        f"{selected_lat_band['prefix']}-"
        f"{selected_lon_band['number']}"
    )

    return {
        "code": sheet_code,
        "prefix": selected_lat_band["prefix"],
        "number": selected_lon_band["number"],
        "lat_min": selected_lat_band["lat_min"],
        "lat_max": selected_lat_band["lat_max"],
        "lon_min": selected_lon_band["lon_min"],
        "lon_max": selected_lon_band["lon_max"],
    }


# ============================================================
# 9. QUADRANT
# ============================================================

def get_quadrant(lat, lon, sheet):
    mid_lat = (sheet["lat_min"] + sheet["lat_max"]) / 2.0
    mid_lon = (sheet["lon_min"] + sheet["lon_max"]) / 2.0

    north = float(lat) >= mid_lat
    east = float(lon) >= mid_lon

    if north and east:
        return "NE"
    if north and not east:
        return "NW"
    if not north and east:
        return "SE"

    return "SW"


# ============================================================
# 10. LETTER A-P
# ============================================================

def get_letter_position(lat, lon, sheet):
    letter_width = 1.5
    letter_height = 1.0

    col = int(
        (float(lon) - sheet["lon_min"]) / letter_width
    )

    row = int(
        (float(lat) - sheet["lat_min"]) / letter_height
    )

    col = max(0, min(3, col))
    row = max(0, min(3, row))

    return row, col


def get_letter(lat, lon, sheet):
    row, col = get_letter_position(
        lat,
        lon,
        sheet,
    )

    return LETTER_GRID[row][col]


# ============================================================
# 11. NUMBER 1-6
# ============================================================

# Inside each letter:
#
# 4 | 5 | 6
# ---------
# 1 | 2 | 3


def get_number(lat, lon, sheet):
    letter_width = 1.5
    letter_height = 1.0

    row, col = get_letter_position(
        lat,
        lon,
        sheet,
    )

    letter_lon_min = (
        sheet["lon_min"]
        + col * letter_width
    )

    letter_lat_min = (
        sheet["lat_min"]
        + row * letter_height
    )

    local_lon = float(lon) - letter_lon_min
    local_lat = float(lat) - letter_lat_min

    number_col = int(local_lon / 0.5)
    number_col = max(0, min(2, number_col))

    number_row = int(local_lat / 0.5)
    number_row = max(0, min(1, number_row))

    if number_row == 0:
        return number_col + 1

    return number_col + 4


# ============================================================
# 12. COMPLETE IMW CALCULATION
# ============================================================

def calculate_imw(lat, lon):
    lat = float(lat)
    lon = float(lon)

    sheet = get_1m_sheet(lat, lon)

    if sheet is None:
        return {
            "IMW_1M": None,
            "IMW_QUADRANT": None,
            "IMW_250K": None,
            "IMW_100K": None,
            "Correct_IMW_Code": "OUTSIDE_DEFINED_IMW",
        }

    quadrant = get_quadrant(
        lat,
        lon,
        sheet,
    )

    letter = get_letter(
        lat,
        lon,
        sheet,
    )

    number = get_number(
        lat,
        lon,
        sheet,
    )

    code = (
        f"{sheet['code']} "
        f"{quadrant} "
        f"{letter}{number}"
    )

    return {
        "IMW_1M": sheet["code"],
        "IMW_QUADRANT": quadrant,
        "IMW_250K": letter,
        "IMW_100K": f"{letter}{number}",
        "Correct_IMW_Code": code,
    }


# ============================================================
# 13. VERIFIED CONTROL POINT
# ============================================================

CONTROL_LAT = 25.0
CONTROL_LON = 34.0
CONTROL_EXPECTED = "NG-36 SE G3"


def verify_control_point():
    result = calculate_imw(
        CONTROL_LAT,
        CONTROL_LON,
    )

    return result["Correct_IMW_Code"] == CONTROL_EXPECTED


# ============================================================
# 14. GEOPHYSICAL SCORING
# ============================================================

def _numeric_series(series):
    """Return a numeric Series while preserving the original index."""
    return pd.to_numeric(series, errors="coerce")


def _empty_score(series):
    return pd.Series([float("nan")] * len(series), index=series.index, dtype="float64")


def _single_value_score(series):
    """A single observation cannot establish relative anomaly strength."""
    numeric = _numeric_series(series)
    return pd.Series(
        [50.0 if pd.notna(v) else float("nan") for v in numeric],
        index=series.index,
        dtype="float64",
    )


def robust_z_score(series):
    """
    Robust normalization:
        robust z = (|x| - median) / (1.4826 * MAD)
    The normal CDF maps the robust z to 0..100.
    """
    numeric = _numeric_series(series).abs()
    valid = numeric.dropna()

    if valid.empty:
        return _empty_score(series)

    if len(valid) < 2:
        return _single_value_score(series)

    median = float(valid.median())
    mad = float((valid - median).abs().median())

    if math.isclose(mad, 0.0):
        return percentile_score(series)

    robust_z = (numeric - median) / (1.4826 * mad)

    def cdf_to_score(value):
        if pd.isna(value):
            return float("nan")
        cdf = 0.5 * (1.0 + math.erf(float(value) / math.sqrt(2.0)))
        return max(0.0, min(100.0, cdf * 100.0))

    return robust_z.apply(cdf_to_score)


def percentile_score(series):
    """Percentile-rank normalization of anomaly magnitude to 0..100."""
    numeric = _numeric_series(series).abs()
    valid = numeric.dropna()

    if valid.empty:
        return _empty_score(series)

    if len(valid) < 2 or math.isclose(float(valid.min()), float(valid.max())):
        return _single_value_score(series)

    ranks = numeric.rank(method="average", pct=True) * 100.0
    return ranks.where(numeric.notna())


def winsorized_minmax_score(series, lower=5.0, upper=95.0):
    """
    Winsorized Min-Max normalization.
    Extreme values are clipped to the requested percentiles before scaling.
    """
    numeric = _numeric_series(series).abs()
    valid = numeric.dropna()

    if valid.empty:
        return _empty_score(series)

    if len(valid) < 2 or math.isclose(float(valid.min()), float(valid.max())):
        return _single_value_score(series)

    low = float(valid.quantile(lower / 100.0))
    high = float(valid.quantile(upper / 100.0))

    if math.isclose(low, high):
        return percentile_score(series)

    clipped = numeric.clip(lower=low, upper=high)
    score = (clipped - low) / (high - low) * 100.0
    return score.where(numeric.notna())


def minmax_score(series):
    """Backward-compatible alias for the professional default."""
    return robust_z_score(series)


NORMALIZATION_METHODS = {
    "Robust Z-score": robust_z_score,
    "Percentile": percentile_score,
    "Winsorized Min-Max": winsorized_minmax_score,
}


def calculate_data_confidence(df):
    """
    Data confidence follows the requested geophysical-data completeness rules:
      100 = magnetic + gravity available
       60 = gravity missing, magnetic available
       40 = magnetic available, gravity missing
       NaN = both magnetic and gravity missing (no score)

    The result is deliberately NaN for rows with no geophysical inputs so that
    "both missing" is represented as "No Score", not as 0% confidence.
    """
    mag_present = _numeric_series(
        df.get("Mag_Anomaly", pd.Series(index=df.index))
    ).notna()
    grav_present = _numeric_series(
        df.get("Grav_Anomaly", pd.Series(index=df.index))
    ).notna()

    confidence = pd.Series(float("nan"), index=df.index, dtype="float64")
    confidence.loc[mag_present & grav_present] = 100.0
    confidence.loc[mag_present & ~grav_present] = 60.0
    confidence.loc[~mag_present & grav_present] = 40.0

    return confidence


def calculate_target_scores(
    df,
    magnetic_weight=60.0,
    gravity_weight=40.0,
    normalization_method="Robust Z-score",
    concordance_weight=15.0,
):
    """
    Professional geophysical targeting pipeline.

    1. Normalize magnetic/gravity anomaly magnitude.
    2. Apply magnetic/gravity weights.
    3. Calculate data confidence from input completeness.
    4. Calculate per-method contributions.
    5. Calculate concordance between normalized magnetic and gravity.
    6. Blend weighted evidence and concordance using the requested
       Concordance Weight % (default 15%, therefore evidence weight is 85%).
    """
    result = df.copy()

    method = NORMALIZATION_METHODS.get(
        normalization_method,
        robust_z_score,
    )

    result["Mag_Normalized"] = method(result["Mag_Anomaly"])
    result["Grav_Normalized"] = method(result["Grav_Anomaly"])

    # Preserve the historical score column names for compatibility.
    result["Mag_Score"] = result["Mag_Normalized"]
    result["Grav_Score"] = result["Grav_Normalized"]

    try:
        magnetic_weight = float(magnetic_weight)
        gravity_weight = float(gravity_weight)
    except (TypeError, ValueError):
        magnetic_weight, gravity_weight = 60.0, 40.0

    magnetic_weight = max(0.0, magnetic_weight)
    gravity_weight = max(0.0, gravity_weight)

    total_weight = magnetic_weight + gravity_weight
    if total_weight <= 0:
        magnetic_weight, gravity_weight = 60.0, 40.0
        total_weight = 100.0

    result["Data_Confidence"] = calculate_data_confidence(result)
    confidence_factor = result["Data_Confidence"].fillna(0.0) / 100.0

    # Contributions are expressed directly on the 0..100 target-score scale.
    # The requested confidence is applied to the available evidence.
    result["Mag_Contribution"] = (
        result["Mag_Normalized"].fillna(0.0)
        * magnetic_weight
        / total_weight
        * confidence_factor
    )
    result["Grav_Contribution"] = (
        result["Grav_Normalized"].fillna(0.0)
        * gravity_weight
        / total_weight
        * confidence_factor
    )

    # Concordance rewards agreement between the two independent normalized
    # signals. With only one signal available, 50 is neutral. With neither
    # signal available, concordance remains NaN ("No Score").
    both = result["Mag_Normalized"].notna() & result["Grav_Normalized"].notna()
    mag_only = result["Mag_Normalized"].notna() & result["Grav_Normalized"].isna()
    grav_only = result["Mag_Normalized"].isna() & result["Grav_Normalized"].notna()

    result["Concordance"] = float("nan")
    result.loc[both, "Concordance"] = (
        100.0
        - (
            result.loc[both, "Mag_Normalized"]
            - result.loc[both, "Grav_Normalized"]
        ).abs()
    ).clip(lower=0.0, upper=100.0)

    result.loc[mag_only | grav_only, "Concordance"] = 50.0

    # Validate and normalize the explicit Concordance Weight %.
    try:
        concordance_weight = float(concordance_weight)
    except (TypeError, ValueError):
        concordance_weight = 15.0

    concordance_weight = max(0.0, min(100.0, concordance_weight))
    evidence_weight = 100.0 - concordance_weight

    core_score = (
        result["Mag_Contribution"].fillna(0.0)
        + result["Grav_Contribution"].fillna(0.0)
    )
    result["Target_Score"] = (
        core_score * (evidence_weight / 100.0)
        + result["Concordance"].fillna(0.0) * (concordance_weight / 100.0)
    )

    # Both inputs missing means there is no score at all.
    no_data = result["Data_Confidence"].isna()
    result.loc[no_data, "Target_Score"] = float("nan")

    def priority(value):
        if pd.isna(value):
            return "No Data"
        if value >= 80:
            return "VERY HIGH"
        if value >= 65:
            return "HIGH"
        if value >= 50:
            return "MEDIUM"
        return "LOW"

    result["Target_Priority"] = result["Target_Score"].apply(priority)
    result["Normalization_Method"] = normalization_method

    return result


# ============================================================
# 15. PROCESS DATAFRAME
# ============================================================

def process_dataframe(df_input):
    df = df_input.copy()

    df = df.loc[
        :,
        ~df.columns.astype(str).str.startswith("Unnamed"),
    ]

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
    )

    required_columns = [
        "Well_Name",
        "Lat",
        "Lon",
        "Mag_Anomaly",
        "Grav_Anomaly",
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing columns: "
            + ", ".join(missing_columns)
        )

    for column in [
        "Lat",
        "Lon",
        "Mag_Anomaly",
        "Grav_Anomaly",
    ]:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        )

    results = []

    for _, row in df.iterrows():
        lat = row["Lat"]
        lon = row["Lon"]

        if pd.isna(lat) or pd.isna(lon):
            results.append(
                {
                    "Easting": None,
                    "Northing": None,
                    "Projection_EPSG": None,
                    "IMW_1M": None,
                    "IMW_QUADRANT": None,
                    "IMW_250K": None,
                    "IMW_100K": None,
                    "Correct_IMW_Code": "INVALID_COORDINATES",
                }
            )
            continue

        try:
            easting, northing, epsg = latlon_to_projected(
                lat,
                lon,
            )

            imw = calculate_imw(
                lat,
                lon,
            )

            results.append(
                {
                    "Easting": easting,
                    "Northing": northing,
                    "Projection_EPSG": epsg,
                    "IMW_1M": imw["IMW_1M"],
                    "IMW_QUADRANT": imw["IMW_QUADRANT"],
                    "IMW_250K": imw["IMW_250K"],
                    "IMW_100K": imw["IMW_100K"],
                    "Correct_IMW_Code": imw["Correct_IMW_Code"],
                }
            )

        except Exception:
            results.append(
                {
                    "Easting": None,
                    "Northing": None,
                    "Projection_EPSG": None,
                    "IMW_1M": None,
                    "IMW_QUADRANT": None,
                    "IMW_250K": None,
                    "IMW_100K": None,
                    "Correct_IMW_Code": "PROJECTION_ERROR",
                }
            )

    calculated = pd.DataFrame(results)

    for column in calculated.columns:
        df[column] = calculated[column].values

    df = calculate_target_scores(df)

    return df


# ============================================================
# 16. SAVE DATABASE
# ============================================================

def save_to_database(df):
    initialize_database()

    columns_to_save = [
        "Well_Name",
        "Lat",
        "Lon",
        "Easting",
        "Northing",
        "Projection_EPSG",
        "IMW_1M",
        "IMW_QUADRANT",
        "IMW_250K",
        "IMW_100K",
        "Correct_IMW_Code",
        "Mag_Anomaly",
        "Grav_Anomaly",
        "Mag_Normalized",
        "Grav_Normalized",
        "Mag_Score",
        "Grav_Score",
        "Mag_Contribution",
        "Grav_Contribution",
        "Data_Confidence",
        "Concordance",
        "Target_Score",
        "Target_Priority",
        "Normalization_Method",
    ]

    save_df = df.copy()

    for column in columns_to_save:
        if column not in save_df.columns:
            save_df[column] = None

    save_df = save_df[
        columns_to_save
    ].copy()

    save_df["Created_At"] = datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    conn = get_connection()

    try:
        save_df.to_sql(
            "geophysical_wells",
            conn,
            if_exists="append",
            index=False,
        )

        conn.commit()

    finally:
        conn.close()


# ============================================================
# 17. LOAD DATABASE
# ============================================================

def create_embedded_demo_data(n=36):
    """Create deterministic demo observations when no external data exists."""
    rows = []
    for i in range(n):
        lat = 24.25 + (i % 9) * 0.42
        lon = 32.15 + (i // 9) * 0.72 + (i % 3) * 0.08

        # Deterministic synthetic anomalies with several deliberately
        # concordant high-anomaly locations.
        anomaly_cluster = 220.0 * math.exp(
            -(
                ((lat - 26.1) / 0.9) ** 2
                + ((lon - 33.7) / 0.9) ** 2
            )
        )
        mag = (
            180.0
            + 18.0 * math.sin(i * 0.9)
            + anomaly_cluster
            + (i % 5) * 14.0
        )
        grav = (
            6.0
            + 1.2 * math.cos(i * 0.7)
            + anomaly_cluster / 35.0
            + (i % 4) * 0.6
        )

        # Add a few negative anomalies to ensure magnitude-based normalization
        # remains useful for either anomaly polarity.
        if i in (7, 19, 31):
            mag *= -1.0
        if i in (11, 27):
            grav *= -1.0

        rows.append(
            {
                "Well_Name": f"DEMO-{i + 1:03d}",
                "Lat": round(lat, 6),
                "Lon": round(lon, 6),
                "Mag_Anomaly": round(mag, 3),
                "Grav_Anomaly": round(grav, 3),
            }
        )

    return pd.DataFrame(rows)


def load_database():
    initialize_database()

    conn = get_connection()

    try:
        df = pd.read_sql_query(
            """
            SELECT
                id,
                Well_Name,
                Lat,
                Lon,
                Easting,
                Northing,
                Projection_EPSG,
                IMW_1M,
                IMW_QUADRANT,
                IMW_250K,
                IMW_100K,
                Correct_IMW_Code,
                Mag_Anomaly,
                Grav_Anomaly,
                Mag_Normalized,
                Grav_Normalized,
                Mag_Score,
                Grav_Score,
                Mag_Contribution,
                Grav_Contribution,
                Data_Confidence,
                Concordance,
                Target_Score,
                Target_Priority,
                Normalization_Method,
                Created_At
            FROM geophysical_wells
            ORDER BY id DESC
            """,
            conn,
        )

    finally:
        conn.close()

    if df.empty:
        # Prefer a local demo_data.csv when supplied, but never require it.
        demo_path = BASE_DIR / "demo_data.csv"
        if demo_path.exists():
            try:
                demo = pd.read_csv(demo_path)
            except Exception:
                demo = create_embedded_demo_data()
        else:
            demo = create_embedded_demo_data()

        demo = process_dataframe(demo)
        demo.insert(0, "id", range(1, len(demo) + 1))
        demo["Created_At"] = "DEMO"
        return demo

    return df


# ============================================================
# 18. RECALCULATE OLD DATABASE
# ============================================================

def recalculate_all_imw():
    initialize_database()

    conn = get_connection()

    try:
        cursor = conn.cursor()

        cursor.execute(
            """
            SELECT
                id,
                Lat,
                Lon,
                Mag_Anomaly,
                Grav_Anomaly
            FROM geophysical_wells
            WHERE
                Lat IS NOT NULL
                AND Lon IS NOT NULL
            """
        )

        rows = cursor.fetchall()

        count = 0

        all_values = []

        for row_id, lat, lon, mag, grav in rows:
            try:
                easting, northing, epsg = latlon_to_projected(
                    lat,
                    lon,
                )

                imw = calculate_imw(
                    lat,
                    lon,
                )

                all_values.append(
                    {
                        "id": row_id,
                        "Mag_Anomaly": mag,
                        "Grav_Anomaly": grav,
                        "Easting": easting,
                        "Northing": northing,
                        "Projection_EPSG": epsg,
                        **imw,
                    }
                )

            except Exception:
                continue

        temp = pd.DataFrame(all_values)

        if not temp.empty:
            scored = calculate_target_scores(temp)

            for _, row in scored.iterrows():
                cursor.execute(
                    """
                    UPDATE geophysical_wells
                    SET
                        Easting = ?,
                        Northing = ?,
                        Projection_EPSG = ?,
                        IMW_1M = ?,
                        IMW_QUADRANT = ?,
                        IMW_250K = ?,
                        IMW_100K = ?,
                        Correct_IMW_Code = ?,
                        Mag_Normalized = ?,
                        Grav_Normalized = ?,
                        Mag_Score = ?,
                        Grav_Score = ?,
                        Mag_Contribution = ?,
                        Grav_Contribution = ?,
                        Data_Confidence = ?,
                        Concordance = ?,
                        Target_Score = ?,
                        Target_Priority = ?,
                        Normalization_Method = ?
                    WHERE id = ?
                    """,
                    (
                        row["Easting"],
                        row["Northing"],
                        row["Projection_EPSG"],
                        row["IMW_1M"],
                        row["IMW_QUADRANT"],
                        row["IMW_250K"],
                        row["IMW_100K"],
                        row["Correct_IMW_Code"],
                        row["Mag_Normalized"],
                        row["Grav_Normalized"],
                        row["Mag_Score"],
                        row["Grav_Score"],
                        row["Mag_Contribution"],
                        row["Grav_Contribution"],
                        row["Data_Confidence"],
                        row["Concordance"],
                        row["Target_Score"],
                        row["Target_Priority"],
                        row["Normalization_Method"],
                        int(row["id"]),
                    ),
                )

                count += 1

        conn.commit()

    finally:
        conn.close()

    return count


# ============================================================
# 19. CLEAR DATABASE
# ============================================================

def clear_database():
    conn = get_connection()

    try:
        conn.execute(
            "DELETE FROM geophysical_wells"
        )
        conn.commit()

    finally:
        conn.close()


# ============================================================
# 20. MAP GRID
# ============================================================

def add_imw_grid(fmap, show_detail=True):
    """
    Draws the geographic IMW boundaries and 1-degree letter grid.
    The grid is geographic, so Folium displays it correctly on the
    geographic map. Projected Easting/Northing remain in the database.
    """

    # --------------------------------------------------------
    # Major 1M boundaries
    # --------------------------------------------------------

    for lat in [20.0, 24.0, 28.0, 32.0]:
        folium.PolyLine(
            [
                [lat, 24.0],
                [lat, 36.0],
            ],
            weight=3,
            opacity=0.8,
            tooltip=f"IMW Latitude {lat}°",
        ).add_to(fmap)

    for lon in [24.0, 30.0, 36.0]:
        folium.PolyLine(
            [
                [20.0, lon],
                [32.0, lon],
            ],
            weight=3,
            opacity=0.8,
            tooltip=f"IMW Longitude {lon}°",
        ).add_to(fmap)

    if not show_detail:
        return

    # --------------------------------------------------------
    # 1-degree letter grid
    # --------------------------------------------------------

    for lat in range(20, 33):
        folium.PolyLine(
            [
                [float(lat), 24.0],
                [float(lat), 36.0],
            ],
            weight=1,
            opacity=0.25,
        ).add_to(fmap)

    for lon in range(24, 37):
        folium.PolyLine(
            [
                [20.0, float(lon)],
                [32.0, float(lon)],
            ],
            weight=1,
            opacity=0.25,
        ).add_to(fmap)

    # --------------------------------------------------------
    # Add sheet / letter labels at cell centers
    # --------------------------------------------------------

    for sheet in IMW_SHEETS:
        for lon_sheet in IMW_LONGITUDE_SHEETS:
            for row in range(4):
                for col in range(4):
                    lat_min = sheet["lat_min"] + row
                    lat_center = lat_min + 0.5

                    lon_min = (
                        lon_sheet["lon_min"]
                        + col * 1.5
                    )

                    lon_center = lon_min + 0.75

                    letter = LETTER_GRID[row][col]

                    folium.Marker(
                        [lat_center, lon_center],
                        icon=folium.DivIcon(
                            html=(
                                '<div style="'
                                'font-size:11px;'
                                'font-weight:bold;'
                                'color:#222;'
                                'text-align:center;'
                                '">'
                                f"{sheet['prefix']}-"
                                f"{lon_sheet['number']} "
                                f"{letter}"
                                "</div>"
                            )
                        ),
                        tooltip=(
                            f"{sheet['prefix']}-"
                            f"{lon_sheet['number']} "
                            f"{letter}"
                        ),
                    ).add_to(fmap)


# ============================================================
# 21. MAP BASEMAPS
# ============================================================

def add_basemaps(fmap):
    folium.TileLayer(
        tiles="OpenStreetMap",
        name="🗺️ Geographic / OpenStreetMap",
        control=True,
    ).add_to(fmap)

    folium.TileLayer(
        tiles=(
            "https://server.arcgisonline.com/"
            "ArcGIS/rest/services/World_Imagery/"
            "MapServer/tile/{z}/{y}/{x}"
        ),
        attr="Esri World Imagery",
        name="🛰️ Satellite / Esri",
        overlay=False,
        control=True,
    ).add_to(fmap)

    folium.TileLayer(
        tiles=(
            "https://server.arcgisonline.com/"
            "ArcGIS/rest/services/"
            "World_Topo_Map/"
            "MapServer/tile/{z}/{y}/{x}"
        ),
        attr="Esri World Topographic Map",
        name="⛰️ Terrain / Topographic",
        overlay=False,
        control=True,
    ).add_to(fmap)


# ============================================================
# 22. MAP POINTS
# ============================================================

def add_data_points(fmap, map_df):
    for _, row in map_df.iterrows():
        imw_code = str(
            row.get(
                "Correct_IMW_Code",
                "",
            )
        )

        target_score = row.get(
            "Target_Score",
            float("nan"),
        )

        if pd.notna(target_score):
            if target_score >= 80:
                marker_color = "red"
            elif target_score >= 65:
                marker_color = "orange"
            elif target_score >= 50:
                marker_color = "green"
            else:
                marker_color = "blue"
        else:
            marker_color = "blue"

        well_name = str(
            row.get(
                "Well_Name",
                "Station",
            )
        )

        popup_html = (
            f"<b>{xml_escape(well_name)}</b><br>"
            f"Latitude: {row['Lat']:.6f}<br>"
            f"Longitude: {row['Lon']:.6f}<br>"
            f"<b>IMW: {xml_escape(imw_code)}</b><br>"
            f"Magnetic: {row.get('Mag_Anomaly', '')} nT<br>"
            f"Gravity: {row.get('Grav_Anomaly', '')} mGal<br>"
            f"Target Score: "
            f"{row.get('Target_Score', '')}<br>"
            f"Priority: "
            f"{xml_escape(str(row.get('Target_Priority', '')))}"
        )

        folium.Marker(
            location=[
                row["Lat"],
                row["Lon"],
            ],
            popup=folium.Popup(
                popup_html,
                max_width=420,
            ),
            tooltip=imw_code,
            icon=folium.Icon(
                color=marker_color,
                icon="info-sign",
            ),
        ).add_to(fmap)


# ============================================================
# 23. MAP WITH HEATMAP
# ============================================================

def create_map(
    map_df,
    basemap="Satellite",
    show_grid=True,
    show_points=True,
    show_magnetic_heatmap=False,
    show_gravity_heatmap=False,
    show_target_heatmap=False,
):
    if map_df.empty:
        return None

    center_lat = float(
        map_df["Lat"].mean()
    )
    center_lon = float(
        map_df["Lon"].mean()
    )

    fmap = folium.Map(
        location=[
            center_lat,
            center_lon,
        ],
        zoom_start=7,
        control_scale=True,
        tiles=None,
    )

    add_basemaps(fmap)

    # Put requested basemap first in visible order.
    # Folium will still provide the layer control.

    if show_grid:
        add_imw_grid(
            fmap,
            show_detail=True,
        )

    if show_points:
        add_data_points(
            fmap,
            map_df,
        )

    # --------------------------------------------------------
    # Magnetic heatmap
    # --------------------------------------------------------

    if show_magnetic_heatmap:
        temp = map_df.dropna(
            subset=[
                "Lat",
                "Lon",
                "Mag_Anomaly",
            ]
        ).copy()

        if not temp.empty:
            values = temp["Mag_Anomaly"].astype(float)

            min_value = values.min()
            max_value = values.max()

            if max_value > min_value:
                weights = (
                    (values - min_value)
                    / (max_value - min_value)
                )
            else:
                weights = pd.Series(
                    [1.0] * len(temp),
                    index=temp.index,
                )

            heat_data = [
                [
                    float(row["Lat"]),
                    float(row["Lon"]),
                    float(weights.loc[index]),
                ]
                for index, row in temp.iterrows()
            ]

            HeatMap(
                heat_data,
                name="🧲 Magnetic Heatmap",
                radius=25,
                blur=20,
                min_opacity=0.35,
            ).add_to(fmap)

    # --------------------------------------------------------
    # Gravity heatmap
    # --------------------------------------------------------

    if show_gravity_heatmap:
        temp = map_df.dropna(
            subset=[
                "Lat",
                "Lon",
                "Grav_Anomaly",
            ]
        ).copy()

        if not temp.empty:
            values = temp["Grav_Anomaly"].astype(float)

            min_value = values.min()
            max_value = values.max()

            if max_value > min_value:
                weights = (
                    (values - min_value)
                    / (max_value - min_value)
                )
            else:
                weights = pd.Series(
                    [1.0] * len(temp),
                    index=temp.index,
                )

            heat_data = [
                [
                    float(row["Lat"]),
                    float(row["Lon"]),
                    float(weights.loc[index]),
                ]
                for index, row in temp.iterrows()
            ]

            HeatMap(
                heat_data,
                name="🌋 Gravity Heatmap",
                radius=25,
                blur=20,
                min_opacity=0.35,
            ).add_to(fmap)

    # --------------------------------------------------------
    # Target heatmap
    # --------------------------------------------------------

    if show_target_heatmap:
        temp = map_df.dropna(
            subset=[
                "Lat",
                "Lon",
                "Target_Score",
            ]
        ).copy()

        if not temp.empty:
            values = temp["Target_Score"].astype(float)

            min_value = values.min()
            max_value = values.max()

            if max_value > min_value:
                weights = (
                    (values - min_value)
                    / (max_value - min_value)
                )
            else:
                weights = pd.Series(
                    [1.0] * len(temp),
                    index=temp.index,
                )

            heat_data = [
                [
                    float(row["Lat"]),
                    float(row["Lon"]),
                    float(weights.loc[index]),
                ]
                for index, row in temp.iterrows()
            ]

            HeatMap(
                heat_data,
                name="🎯 Exploration Target Heatmap",
                radius=30,
                blur=25,
                min_opacity=0.4,
            ).add_to(fmap)

    folium.LayerControl(
        collapsed=False
    ).add_to(fmap)

    return fmap


# ============================================================
# 24. EXPORT HELPERS
# ============================================================

def dataframe_to_excel_bytes(df):
    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Geophysical_Data",
        )

    output.seek(0)

    return output.getvalue()


def dataframe_to_json_bytes(df):
    text = df.to_json(
        orient="records",
        force_ascii=False,
        indent=2,
        date_format="iso",
    )

    return text.encode("utf-8")


def dataframe_to_geojson(df):
    features = []

    for _, row in df.iterrows():
        if pd.isna(row.get("Lat")) or pd.isna(row.get("Lon")):
            continue

        properties = {}

        for column in df.columns:
            if column in ["Lat", "Lon"]:
                continue

            value = row[column]

            if pd.isna(value):
                properties[column] = None
            elif isinstance(
                value,
                (pd.Timestamp, datetime),
            ):
                properties[column] = str(value)
            elif hasattr(value, "item"):
                try:
                    properties[column] = value.item()
                except Exception:
                    properties[column] = str(value)
            else:
                properties[column] = value

        features.append(
            {
                "type": "Feature",
                "properties": properties,
                "geometry": {
                    "type": "Point",
                    "coordinates": [
                        float(row["Lon"]),
                        float(row["Lat"]),
                    ],
                },
            }
        )

    collection = {
        "type": "FeatureCollection",
        "name": "Egypt_Geophysical_Wells",
        "features": features,
    }

    return json.dumps(
        collection,
        ensure_ascii=False,
        indent=2,
    ).encode("utf-8")


def dataframe_to_kml(df):
    parts = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<kml xmlns="http://www.opengis.net/kml/2.2">',
        "<Document>",
        "<name>Egypt Geophysical Wells</name>",
    ]

    for _, row in df.iterrows():
        if pd.isna(row.get("Lat")) or pd.isna(row.get("Lon")):
            continue

        name = xml_escape(
            str(row.get("Well_Name", "Station"))
        )

        description = (
            f"IMW: {xml_escape(str(row.get('Correct_IMW_Code', '')))}\n"
            f"Magnetic: {xml_escape(str(row.get('Mag_Anomaly', '')))} nT\n"
            f"Gravity: {xml_escape(str(row.get('Grav_Anomaly', '')))} mGal\n"
            f"Target Score: {xml_escape(str(row.get('Target_Score', '')))}"
        )

        lon = float(row["Lon"])
        lat = float(row["Lat"])

        parts.extend(
            [
                "<Placemark>",
                f"<name>{name}</name>",
                f"<description>{xml_escape(description)}</description>",
                "<Point>",
                f"<coordinates>{lon},{lat},0</coordinates>",
                "</Point>",
                "</Placemark>",
            ]
        )

    parts.extend(
        [
            "</Document>",
            "</kml>",
        ]
    )

    return "\n".join(parts).encode("utf-8")


def dataframe_to_shapefile_zip(df):
    """
    Optional Shapefile export.

    Requires:
        geopandas
        pyogrio OR fiona

    Returns:
        ZIP bytes.
    """

    from zipfile import ZIP_DEFLATED, ZipFile

    import geopandas as gpd

    temp_dir = BASE_DIR / "_shapefile_export_tmp"
    temp_dir.mkdir(
        exist_ok=True
    )

    # Remove old files from previous export.
    for item in temp_dir.iterdir():
        if item.is_file():
            item.unlink()

    geo_df = df.dropna(
        subset=[
            "Lat",
            "Lon",
        ]
    ).copy()

    if geo_df.empty:
        raise ValueError(
            "No valid coordinates available for Shapefile export."
        )

    geometry = gpd.points_from_xy(
        geo_df["Lon"],
        geo_df["Lat"],
    )

    geo_df = gpd.GeoDataFrame(
        geo_df,
        geometry=geometry,
        crs="EPSG:4326",
    )

    shp_path = temp_dir / "Egypt_Geophysical_Wells.shp"

    geo_df.to_file(
        shp_path,
        driver="ESRI Shapefile",
    )

    zip_buffer = io.BytesIO()

    with ZipFile(
        zip_buffer,
        "w",
        ZIP_DEFLATED,
    ) as zip_file:
        for item in temp_dir.iterdir():
            if item.is_file():
                zip_file.write(
                    item,
                    arcname=item.name,
                )

    zip_buffer.seek(0)

    return zip_buffer.getvalue()


def dataframe_to_pdf(df):
    """
    Optional PDF export.

    Requires reportlab.
    """

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    output = io.BytesIO()

    document = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=10 * mm,
        leftMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )

    styles = getSampleStyleSheet()

    story = []

    story.append(
        Paragraph(
            "Egypt IMW / SGrid Geophysical Report",
            styles["Title"],
        )
    )

    story.append(
        Paragraph(
            (
                "Egypt 1907 / TM • "
                f"Generated: {datetime.now():%Y-%m-%d %H:%M:%S}"
            ),
            styles["Normal"],
        )
    )

    story.append(Spacer(1, 8))

    summary = [
        ["Records", len(df)],
        [
            "Verified control point",
            "PASS" if verify_control_point() else "FAIL",
        ],
    ]

    if "Mag_Anomaly" in df.columns:
        summary.extend(
            [
                [
                    "Magnetic minimum",
                    round(
                        pd.to_numeric(
                            df["Mag_Anomaly"],
                            errors="coerce",
                        ).min(),
                        3,
                    ),
                ],
                [
                    "Magnetic maximum",
                    round(
                        pd.to_numeric(
                            df["Mag_Anomaly"],
                            errors="coerce",
                        ).max(),
                        3,
                    ),
                ],
            ]
        )

    if "Grav_Anomaly" in df.columns:
        summary.extend(
            [
                [
                    "Gravity minimum",
                    round(
                        pd.to_numeric(
                            df["Grav_Anomaly"],
                            errors="coerce",
                        ).min(),
                        3,
                    ),
                ],
                [
                    "Gravity maximum",
                    round(
                        pd.to_numeric(
                            df["Grav_Anomaly"],
                            errors="coerce",
                        ).max(),
                        3,
                    ),
                ],
            ]
        )

    summary_table = Table(
        summary,
        colWidths=[60 * mm, 45 * mm],
    )

    summary_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (0, -1), colors.lightgrey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )

    story.append(summary_table)
    story.append(Spacer(1, 10))

    display_columns = [
        column
        for column in [
            "Well_Name",
            "Lat",
            "Lon",
            "IMW_1M",
            "Correct_IMW_Code",
            "Mag_Anomaly",
            "Grav_Anomaly",
            "Mag_Normalized",
            "Grav_Normalized",
            "Mag_Contribution",
            "Grav_Contribution",
            "Data_Confidence",
            "Concordance",
            "Target_Score",
            "Target_Priority",
            "Normalization_Method",
        ]
        if column in df.columns
    ]

    report_df = df[display_columns].head(40).copy()

    table_data = [
        [
            Paragraph(
                str(column),
                styles["BodyText"],
            )
            for column in display_columns
        ]
    ]

    for _, row in report_df.iterrows():
        table_data.append(
            [
                Paragraph(
                    str(row[column]),
                    styles["BodyText"],
                )
                for column in display_columns
            ]
        )

    data_table = Table(
        table_data,
        repeatRows=1,
    )

    data_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.lightgrey,
                ),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )

    story.append(data_table)

    document.build(story)

    output.seek(0)

    return output.getvalue()


# ============================================================
# 25. IMPORT SECTION
# ============================================================

def read_uploaded_file(uploaded_file):
    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded_file)

    return pd.read_excel(uploaded_file)


# ============================================================
# 26. SIDEBAR NAVIGATION
# ============================================================

st.sidebar.markdown("## 🧭 Exploration Workspace")
st.sidebar.caption("Professional Geophysical Targeting • Egypt IMW / SGrid")

section = st.sidebar.radio(
    "Select Section",
    [
        "🏠 Dashboard",
        "📍 IMW / Coordinates",
        "📥 Import Data",
        "⚡ Geophysical Analysis",
        "🎯 Exploration Targets",
        "🗺️ GIS Map",
        "🗄️ Database",
        "📤 Export",
        "🔧 Database Repair",
    ],
    key="main_navigation_m4",
)


# ============================================================
# 27. DASHBOARD
# ============================================================

if section == "🏠 Dashboard":
    df_dashboard = load_database()

    st.header("🏠 Dashboard")
    st.caption("Demo data is loaded automatically when the local database is empty. Import your own data to replace it.")

    if df_dashboard.empty:
        st.info(
            "قاعدة البيانات فارغة. ابدأ من Import Data أو IMW / Coordinates."
        )
    else:
        total_records = len(df_dashboard)

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric(
                "عدد السجلات",
                total_records,
            )

        with col2:
            st.metric(
                "IMW Sheets",
                df_dashboard["IMW_1M"]
                .dropna()
                .nunique(),
            )

        with col3:
            high_count = int(
                (
                    df_dashboard["Target_Priority"]
                    .isin(
                        [
                            "HIGH",
                            "VERY HIGH",
                        ]
                    )
                ).sum()
            )

            st.metric(
                "High Targets",
                high_count,
            )

        with col4:
            st.metric(
                "Control Point",
                "PASS"
                if verify_control_point()
                else "FAIL",
            )

        st.write("---")

        st.subheader("📊 آخر السجلات")

        st.dataframe(
            df_dashboard.head(20),
            use_container_width=True,
            height=450,
        )


# ============================================================
# 28. IMW / COORDINATES
# ============================================================

elif section == "📍 IMW / Coordinates":
    st.header("📍 IMW / SGrid / Egypt 1907")

    tab1, tab2 = st.tabs(
        [
            "🎯 Control Point",
            "📝 Manual DMS",
        ]
    )

    with tab1:
        st.subheader(
            "Verified Control Point"
        )

        st.write(
            "25°00′00″ N, 34°00′00″ E"
        )

        result = calculate_imw(
            CONTROL_LAT,
            CONTROL_LON,
        )

        easting, northing, epsg = latlon_to_projected(
            CONTROL_LAT,
            CONTROL_LON,
        )

        result_display = {
            **result,
            "Lat": CONTROL_LAT,
            "Lon": CONTROL_LON,
            "Easting": easting,
            "Northing": northing,
            "Projection_EPSG": epsg,
        }

        st.dataframe(
            pd.DataFrame(
                [result_display]
            ),
            use_container_width=True,
        )

        if verify_control_point():
            st.success(
                "✅ VERIFIED: "
                "25°00′00″ N, 34°00′00″ E "
                "→ NG-36 SE G3"
            )
        else:
            st.error(
                "❌ Control point calculation failed."
            )

    with tab2:
        st.subheader(
            "📝 إدخال نقطة يدوية DMS"
        )

        well_name = st.text_input(
            "اسم البئر / المحطة",
            value="Station_01",
            key="manual_well_name_m4",
        )

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("### Latitude")

            lat_d = st.number_input(
                "Latitude Degrees",
                min_value=18,
                max_value=32,
                value=25,
                key="m4_lat_degree",
            )

            lat_m = st.number_input(
                "Latitude Minutes",
                min_value=0,
                max_value=59,
                value=0,
                key="m4_lat_minutes",
            )

            lat_s = st.number_input(
                "Latitude Seconds",
                min_value=0.0,
                max_value=59.999,
                value=0.0,
                key="m4_lat_seconds",
            )

        with col2:
            st.markdown("### Longitude")

            lon_d = st.number_input(
                "Longitude Degrees",
                min_value=24,
                max_value=36,
                value=34,
                key="m4_lon_degree",
            )

            lon_m = st.number_input(
                "Longitude Minutes",
                min_value=0,
                max_value=59,
                value=0,
                key="m4_lon_minutes",
            )

            lon_s = st.number_input(
                "Longitude Seconds",
                min_value=0.0,
                max_value=59.999,
                value=0.0,
                key="m4_lon_seconds",
            )

        st.markdown("### 🧲 Geophysical Data")

        mag = st.number_input(
            "Magnetic Anomaly (nT)",
            min_value=-100000.0,
            max_value=100000.0,
            value=550.0,
            key="m4_manual_mag",
        )

        grav = st.number_input(
            "Gravity Anomaly (mGal)",
            min_value=-100000.0,
            max_value=100000.0,
            value=12.0,
            key="m4_manual_grav",
        )

        if st.button(
            "🔍 Calculate IMW",
            type="primary",
            key="m4_calculate_manual",
        ):
            lat = dms_to_dd(
                lat_d,
                lat_m,
                lat_s,
            )

            lon = dms_to_dd(
                lon_d,
                lon_m,
                lon_s,
            )

            try:
                easting, northing, epsg = latlon_to_projected(
                    lat,
                    lon,
                )

                imw = calculate_imw(
                    lat,
                    lon,
                )

                result = {
                    "Well_Name": well_name,
                    "Lat": lat,
                    "Lon": lon,
                    "Easting": easting,
                    "Northing": northing,
                    "Projection_EPSG": epsg,
                    **imw,
                    "Mag_Anomaly": mag,
                    "Grav_Anomaly": grav,
                }

                result_df = calculate_target_scores(
                    pd.DataFrame([result])
                )

                st.session_state[
                    "manual_result_m4"
                ] = result_df

            except Exception as error:
                st.error(
                    f"Calculation error: {error}"
                )

        if "manual_result_m4" in st.session_state:
            result_df = st.session_state[
                "manual_result_m4"
            ]

            st.subheader(
                "📍 نتيجة الحساب"
            )

            st.dataframe(
                result_df,
                use_container_width=True,
            )

            result_code = str(
                result_df.iloc[0]["Correct_IMW_Code"]
            )

            result_lat = float(
                result_df.iloc[0]["Lat"]
            )

            result_lon = float(
                result_df.iloc[0]["Lon"]
            )

            if (
                result_code == CONTROL_EXPECTED
                and abs(result_lat - CONTROL_LAT) < 0.000001
                and abs(result_lon - CONTROL_LON) < 0.000001
            ):
                st.success(
                    "✅ Verified control point → NG-36 SE G3"
                )

            if st.button(
                "💾 حفظ النقطة",
                key="m4_save_manual",
            ):
                save_to_database(
                    result_df
                )

                st.success(
                    "تم حفظ النقطة في قاعدة البيانات."
                )

                st.session_state.pop(
                    "manual_result_m4",
                    None,
                )

                st.rerun()


# ============================================================
# 29. IMPORT DATA
# ============================================================

elif section == "📥 Import Data":
    st.header("📥 Import Excel / CSV")

    st.info(
        """
البرنامج يعالج جميع الصفوف الموجودة في الملف.
الأعمدة المطلوبة:
Well_Name, Lat, Lon, Mag_Anomaly, Grav_Anomaly
"""
    )

    uploaded_file = st.file_uploader(
        "اختر Excel أو CSV",
        type=[
            "xlsx",
            "xls",
            "csv",
        ],
        key="m4_file_uploader",
    )

    if uploaded_file is not None:
        try:
            df_input = read_uploaded_file(
                uploaded_file
            )

            st.success(
                f"تمت قراءة {len(df_input)} صف."
            )

            st.dataframe(
                df_input.head(20),
                use_container_width=True,
            )

            if st.button(
                "⚙️ معالجة جميع الصفوف وحساب IMW",
                type="primary",
                key="m4_process_upload",
            ):
                df_processed = process_dataframe(
                    df_input
                )

                st.session_state[
                    "processed_upload_m4"
                ] = df_processed

                st.success(
                    f"تمت معالجة جميع الصفوف: "
                    f"{len(df_processed)}"
                )

        except Exception as error:
            st.error(
                f"خطأ في الملف: {error}"
            )

    if "processed_upload_m4" in st.session_state:
        df_processed = st.session_state[
            "processed_upload_m4"
        ]

        st.write("---")

        st.subheader(
            "📊 البيانات بعد الحساب"
        )

        st.dataframe(
            df_processed,
            use_container_width=True,
            height=500,
        )

        c1, c2 = st.columns(2)

        with c1:
            if st.button(
                "💾 حفظ جميع الصفوف في SQLite",
                type="primary",
                key="m4_save_upload",
            ):
                try:
                    save_to_database(
                        df_processed
                    )

                    st.success(
                        f"تم حفظ {len(df_processed)} صف."
                    )

                    st.session_state.pop(
                        "processed_upload_m4",
                        None,
                    )

                except Exception as error:
                    st.error(
                        f"Database error: {error}"
                    )

        with c2:
            csv_data = (
                df_processed
                .to_csv(index=False)
                .encode("utf-8-sig")
            )

            st.download_button(
                "📥 تحميل CSV",
                data=csv_data,
                file_name="Egypt_IMW_Processed.csv",
                mime="text/csv",
                key="m4_download_processed_csv",
            )


# ============================================================
# 30. GEOPHYSICAL ANALYSIS
# ============================================================

elif section == "⚡ Geophysical Analysis":
    st.header("⚡ التحليلات الجيوفيزيائية")

    df_analysis = load_database()

    if df_analysis.empty:
        st.warning(
            "لا توجد بيانات. قم باستيراد بيانات أولاً."
        )
    else:
        st.subheader("🎯 الأهداف المغناطيسية")

        mag_numeric = pd.to_numeric(
            df_analysis["Mag_Anomaly"],
            errors="coerce",
        )

        grav_numeric = pd.to_numeric(
            df_analysis["Grav_Anomaly"],
            errors="coerce",
        )

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.metric(
                "Mag Min",
                f"{mag_numeric.min():.2f}",
            )

        with c2:
            st.metric(
                "Mag Max",
                f"{mag_numeric.max():.2f}",
            )

        with c3:
            st.metric(
                "Gravity Min",
                f"{grav_numeric.min():.2f}",
            )

        with c4:
            st.metric(
                "Gravity Max",
                f"{grav_numeric.max():.2f}",
            )

        st.write("---")

        st.subheader("🧲 Magnetic Analysis")

        magnetic_chart = (
            df_analysis[
                ["Well_Name", "Mag_Anomaly"]
            ]
            .dropna()
            .set_index("Well_Name")
        )

        st.bar_chart(
            magnetic_chart,
            y="Mag_Anomaly",
        )

        st.subheader("🌋 Gravity Analysis")

        gravity_chart = (
            df_analysis[
                ["Well_Name", "Grav_Anomaly"]
            ]
            .dropna()
            .set_index("Well_Name")
        )

        st.bar_chart(
            gravity_chart,
            y="Grav_Anomaly",
        )

        st.subheader("🧲 + 🌋 Magnetic vs Gravity")

        scatter_df = df_analysis[
            [
                "Mag_Anomaly",
                "Grav_Anomaly",
            ]
        ].dropna()

        if not scatter_df.empty:
            st.scatter_chart(
                scatter_df,
                x="Mag_Anomaly",
                y="Grav_Anomaly",
            )

        st.subheader("📊 تحليل حسب IMW")

        imw_summary = (
            df_analysis.groupby(
                "IMW_1M",
                dropna=False,
            )
            .agg(
                Records=("id", "count"),
                Magnetic_Mean=(
                    "Mag_Anomaly",
                    "mean",
                ),
                Magnetic_Max=(
                    "Mag_Anomaly",
                    "max",
                ),
                Gravity_Mean=(
                    "Grav_Anomaly",
                    "mean",
                ),
                Gravity_Max=(
                    "Grav_Anomaly",
                    "max",
                ),
                Target_Mean=(
                    "Target_Score",
                    "mean",
                ),
            )
            .reset_index()
        )

        st.dataframe(
            imw_summary,
            use_container_width=True,
        )

        st.subheader(
            "🎯 أعلى الأهداف الجيوفيزيائية"
        )

        target_columns = [
            column
            for column in [
                "Well_Name",
                "Lat",
                "Lon",
                "Correct_IMW_Code",
                "Mag_Anomaly",
                "Grav_Anomaly",
                "Mag_Score",
                "Grav_Score",
                "Target_Score",
                "Target_Priority",
            ]
            if column in df_analysis.columns
        ]

        top_targets = (
            df_analysis[
                target_columns
            ]
            .sort_values(
                "Target_Score",
                ascending=False,
            )
            .head(20)
        )

        st.dataframe(
            top_targets,
            use_container_width=True,
        )


# ============================================================
# 31. EXPLORATION TARGETS
# ============================================================

elif section == "🎯 Exploration Targets":
    st.header("🎯 Professional Geophysical Targeting")

    df_targets = load_database()

    if df_targets.empty:
        st.warning("لا توجد بيانات.")
    else:
        st.markdown(
            """
            **Professional Targeting Pipeline**

            Demo data is embedded and loaded automatically when the database is
            empty. The targeting model normalizes magnetic and gravity anomaly
            magnitudes, applies user-defined weights, accounts for data
            confidence, measures magnetic–gravity concordance, and produces a
            ranked Target Score.
            """
        )

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            normalization_method = st.selectbox(
                "Data Normalization",
                list(NORMALIZATION_METHODS.keys()),
                index=0,
                key="m4_normalization_method",
                help="Robust Z-score is the recommended and default normalization method.",
            )

        with c2:
            magnetic_weight = st.slider(
                "Magnetic Weight %",
                min_value=0,
                max_value=100,
                value=60,
                key="m4_mag_weight",
            )

        with c3:
            gravity_weight = 100 - magnetic_weight
            st.metric("Gravity Weight %", gravity_weight)

        with c4:
            concordance_weight = st.slider(
                "Concordance Weight %",
                min_value=0,
                max_value=100,
                value=15,
                key="m4_concordance_weight",
                help="Default: 15%. The remaining 85% is the weighted magnetic + gravity evidence.",
            )

        evidence_weight = 100 - concordance_weight

        recalculated = calculate_target_scores(
            df_targets,
            magnetic_weight=magnetic_weight,
            gravity_weight=gravity_weight,
            normalization_method=normalization_method,
            concordance_weight=concordance_weight,
        )

        st.subheader("⚖️ Weighting, Concordance & Data Confidence")
        w1, w2, w3, w4, w5 = st.columns(5)
        with w1:
            st.metric("Magnetic Weight", f"{magnetic_weight:.0f}%")
        with w2:
            st.metric("Gravity Weight", f"{gravity_weight:.0f}%")
        with w3:
            st.metric("Concordance Weight", f"{concordance_weight:.0f}%")
        with w4:
            st.metric("Evidence Weight", f"{evidence_weight:.0f}%")
        with w5:
            mean_conf = pd.to_numeric(
                recalculated["Data_Confidence"], errors="coerce"
            ).mean()
            st.metric(
                "Mean Data Confidence",
                f"{mean_conf:.1f}%",
            )

        st.caption(
            "Data Confidence rules: 100% = magnetic + gravity; "
            "60% = magnetic only; 40% = gravity only; "
            "No Score = both missing."
        )

        st.subheader("🧮 Contributions & Concordance")
        contribution_columns = [
            "Well_Name",
            "Mag_Normalized",
            "Grav_Normalized",
            "Mag_Contribution",
            "Grav_Contribution",
            "Data_Confidence",
            "Concordance",
            "Target_Score",
            "Target_Priority",
        ]
        st.dataframe(
            recalculated[
                [c for c in contribution_columns if c in recalculated.columns]
            ]
            .sort_values("Target_Score", ascending=False),
            use_container_width=True,
            height=350,
        )

        st.subheader("🏆 Ranked Targets")

        target_df = (
            recalculated
            .dropna(subset=["Lat", "Lon", "Target_Score"])
            .sort_values("Target_Score", ascending=False)
            .copy()
        )

        display_columns = [
            "Well_Name",
            "Lat",
            "Lon",
            "Correct_IMW_Code",
            "Mag_Anomaly",
            "Grav_Anomaly",
            "Mag_Normalized",
            "Grav_Normalized",
            "Mag_Contribution",
            "Grav_Contribution",
            "Data_Confidence",
            "Concordance",
            "Target_Score",
            "Target_Priority",
        ]

        st.dataframe(
            target_df[
                [c for c in display_columns if c in target_df.columns]
            ].head(50),
            use_container_width=True,
            height=550,
        )

        st.subheader("📊 Target Score Distribution")
        score_chart = (
            target_df[["Well_Name", "Target_Score"]]
            .head(30)
            .set_index("Well_Name")
        )

        if not score_chart.empty:
            st.bar_chart(score_chart, y="Target_Score")

        st.subheader("⚠️ Important Interpretation Note")
        st.warning(
            """
            Target Score هو مؤشر ترتيب استكشافي مبني على البيانات المتاحة،
            وليس إثباتاً لوجود خام أو تركيب جيولوجي محدد.
            التطبيع هنا يستخدم **مقدار الشذوذ |anomaly magnitude|**؛ لذلك
            الإشارة الموجبة/السالبة لا تعني تلقائياً أفضلية استكشافية.
            يجب دمج النتيجة مع الجيولوجيا، التراكيب، العينات، الاستشعار عن بعد
            والبيانات الحقلية قبل اتخاذ قرار استكشافي.
            """
        )


# ============================================================
# 32. GIS MAP
# ============================================================

elif section == "🗺️ GIS Map":
    st.header("🗺️ GIS / Satellite / IMW / Geophysical Map")

    df_map = load_database()

    if df_map.empty:
        st.warning(
            "لا توجد بيانات لرسمها على الخريطة."
        )
    else:
        valid_map = df_map.dropna(
            subset=[
                "Lat",
                "Lon",
            ]
        ).copy()

        if valid_map.empty:
            st.warning(
                "لا توجد إحداثيات صحيحة."
            )
        else:
            st.sidebar.subheader(
                "🗺️ Map Controls"
            )

            map_basemap = st.sidebar.selectbox(
                "Basemap",
                [
                    "Satellite",
                    "Geographic",
                    "Terrain",
                ],
                key="m4_basemap",
            )

            show_grid = st.sidebar.checkbox(
                "Show IMW / SGrid",
                value=True,
                key="m4_show_grid",
            )

            show_points = st.sidebar.checkbox(
                "Show Wells / Stations",
                value=True,
                key="m4_show_points",
            )

            show_mag_heatmap = st.sidebar.checkbox(
                "🧲 Magnetic Heatmap",
                value=False,
                key="m4_mag_heat",
            )

            show_grav_heatmap = st.sidebar.checkbox(
                "🌋 Gravity Heatmap",
                value=False,
                key="m4_grav_heat",
            )

            show_target_heatmap = st.sidebar.checkbox(
                "🎯 Target Heatmap",
                value=False,
                key="m4_target_heat",
            )

            selected_imw = st.sidebar.selectbox(
                "Filter IMW",
                [
                    "ALL"
                ]
                + sorted(
                    valid_map["IMW_1M"]
                    .dropna()
                    .astype(str)
                    .unique()
                    .tolist()
                ),
                key="m4_map_imw_filter",
            )

            if selected_imw != "ALL":
                valid_map = valid_map[
                    valid_map["IMW_1M"].astype(str)
                    == selected_imw
                ].copy()

            if valid_map.empty:
                st.info(
                    "لا توجد نقاط بعد تطبيق الفلتر."
                )
            else:
                fmap = create_map(
                    valid_map,
                    basemap=map_basemap,
                    show_grid=show_grid,
                    show_points=show_points,
                    show_magnetic_heatmap=show_mag_heatmap,
                    show_gravity_heatmap=show_grav_heatmap,
                    show_target_heatmap=show_target_heatmap,
                )

                st_folium(
                    fmap,
                    width=1200,
                    height=700,
                    key="m4_main_map",
                )


# ============================================================
# 33. DATABASE
# ============================================================

elif section == "🗄️ Database":
    st.header("🗄️ SQLite Database")

    df_database = load_database()

    if df_database.empty:
        st.info(
            "قاعدة البيانات فارغة."
        )
    else:
        st.success(
            f"عدد السجلات: {len(df_database)}"
        )

        codes = (
            df_database[
                "Correct_IMW_Code"
            ]
            .dropna()
            .astype(str)
            .unique()
            .tolist()
        )

        codes = sorted(codes)

        selected_code = st.selectbox(
            "اختيار كود IMW",
            ["ALL"] + codes,
            key="m4_database_filter",
        )

        if selected_code == "ALL":
            display_df = df_database.copy()
        else:
            display_df = df_database[
                df_database[
                    "Correct_IMW_Code"
                ].astype(str)
                == selected_code
            ].copy()

        st.dataframe(
            display_df,
            use_container_width=True,
            height=550,
        )

        if st.button(
            "🗑️ حذف جميع السجلات",
            key="m4_clear_database",
        ):
            clear_database()

            st.success(
                "تم حذف جميع السجلات."
            )

            st.rerun()


# ============================================================
# 34. EXPORT
# ============================================================

elif section == "📤 Export":
    st.header("📤 Multi-format Export")

    df_export = load_database()

    if df_export.empty:
        st.warning(
            "لا توجد بيانات للتصدير."
        )
    else:
        st.success(
            f"عدد السجلات المتاحة للتصدير: {len(df_export)}"
        )

        export_imw = st.selectbox(
            "اختيار IMW للتصدير",
            [
                "ALL"
            ]
            + sorted(
                df_export["IMW_1M"]
                .dropna()
                .astype(str)
                .unique()
                .tolist()
            ),
            key="m4_export_imw",
        )

        if export_imw != "ALL":
            export_df = df_export[
                df_export["IMW_1M"].astype(str)
                == export_imw
            ].copy()
        else:
            export_df = df_export.copy()

        st.write(
            f"عدد الصفوف: {len(export_df)}"
        )

        st.dataframe(
            export_df.head(30),
            use_container_width=True,
        )

        st.write("---")

        col1, col2, col3 = st.columns(3)

        with col1:
            csv_bytes = (
                export_df
                .to_csv(index=False)
                .encode("utf-8-sig")
            )

            st.download_button(
                "📄 CSV",
                data=csv_bytes,
                file_name="Egypt_Geophysical_Wells.csv",
                mime="text/csv",
                key="m4_export_csv",
            )

        with col2:
            xlsx_bytes = dataframe_to_excel_bytes(
                export_df
            )

            st.download_button(
                "📊 Excel XLSX",
                data=xlsx_bytes,
                file_name="Egypt_Geophysical_Wells.xlsx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                ),
                key="m4_export_xlsx",
            )

        with col3:
            json_bytes = dataframe_to_json_bytes(
                export_df
            )

            st.download_button(
                "🧾 JSON",
                data=json_bytes,
                file_name="Egypt_Geophysical_Wells.json",
                mime="application/json",
                key="m4_export_json",
            )

        col4, col5, col6 = st.columns(3)

        with col4:
            geojson_bytes = dataframe_to_geojson(
                export_df
            )

            st.download_button(
                "🌍 GeoJSON",
                data=geojson_bytes,
                file_name="Egypt_Geophysical_Wells.geojson",
                mime="application/geo+json",
                key="m4_export_geojson",
            )

        with col5:
            kml_bytes = dataframe_to_kml(
                export_df
            )

            st.download_button(
                "🌐 KML",
                data=kml_bytes,
                file_name="Egypt_Geophysical_Wells.kml",
                mime="application/vnd.google-earth.kml+xml",
                key="m4_export_kml",
            )

        with col6:
            if st.button(
                "🧭 Shapefile ZIP",
                key="m4_prepare_shapefile",
            ):
                try:
                    shp_bytes = dataframe_to_shapefile_zip(
                        export_df
                    )

                    st.session_state[
                        "m4_shp_bytes"
                    ] = shp_bytes

                    st.success(
                        "تم إنشاء Shapefile."
                    )

                except ImportError:
                    st.error(
                        "Shapefile يحتاج geopandas. "
                        "ثبّت: pip install geopandas"
                    )

                except Exception as error:
                    st.error(
                        f"Shapefile error: {error}"
                    )

        if "m4_shp_bytes" in st.session_state:
            st.download_button(
                "📦 تحميل Shapefile ZIP",
                data=st.session_state[
                    "m4_shp_bytes"
                ],
                file_name="Egypt_Geophysical_Wells_Shapefile.zip",
                mime="application/zip",
                key="m4_download_shapefile",
            )

        st.write("---")

        if st.button(
            "📄 Generate PDF Report",
            key="m4_generate_pdf",
        ):
            try:
                pdf_bytes = dataframe_to_pdf(
                    export_df
                )

                st.session_state[
                    "m4_pdf_bytes"
                ] = pdf_bytes

                st.success(
                    "تم إنشاء التقرير PDF."
                )

            except ImportError:
                st.error(
                    "PDF يحتاج reportlab. "
                    "ثبّت: pip install reportlab"
                )

            except Exception as error:
                st.error(
                    f"PDF error: {error}"
                )

        if "m4_pdf_bytes" in st.session_state:
            st.download_button(
                "📥 تحميل PDF Report",
                data=st.session_state[
                    "m4_pdf_bytes"
                ],
                file_name="Egypt_Geophysical_Report.pdf",
                mime="application/pdf",
                key="m4_download_pdf",
            )


# ============================================================
# 35. DATABASE REPAIR
# ============================================================

elif section == "🔧 Database Repair":
    st.header("🔧 Database Repair / Recalculation")

    st.info(
        """
        هذا القسم مخصص للنسخة القديمة من قاعدة البيانات.

        يقوم البرنامج أولاً بإضافة أي أعمدة مفقودة تلقائياً،
        ومنها Projection_EPSG، ثم يعيد حساب:
        Easting
        Northing
        Projection_EPSG
        IMW_1M
        IMW_QUADRANT
        IMW_250K
        IMW_100K
        Correct_IMW_Code
        Magnetic Score
        Gravity Score
        Target Score
        Target Priority
        """
    )

    if st.button(
        "🔧 إصلاح قاعدة البيانات وإعادة حساب جميع الأكواد",
        type="primary",
        key="m4_repair_database",
    ):
        try:
            initialize_database()

            count = recalculate_all_imw()

            st.success(
                f"تم إصلاح قاعدة البيانات وإعادة حساب {count} سجل."
            )

            st.rerun()

        except Exception as error:
            st.error(
                f"Database repair error: {error}"
            )

    st.write("---")

    st.subheader(
        "🔎 فحص أعمدة قاعدة البيانات"
    )

    conn = get_connection()

    try:
        schema_df = pd.read_sql_query(
            "PRAGMA table_info(geophysical_wells)",
            conn,
        )
    finally:
        conn.close()

    st.dataframe(
        schema_df,
        use_container_width=True,
    )


# ============================================================
# 36. GLOBAL VERIFIED CONTROL POINT
# ============================================================

st.write("---")

if verify_control_point():
    st.success(
        "✅ Verified control point: "
        "25°00′00″ N, 34°00′00″ E "
        "→ NG-36 SE G3"
    )
else:
    st.error(
        "❌ Control point calculation failed."
    )


# ============================================================
# 37. FOOTER
# ============================================================

st.markdown("---")
st.caption("Egypt Geophysical Exploration Dashboard • Streamlit • IMW / SGrid • Demo / Research Tool")
