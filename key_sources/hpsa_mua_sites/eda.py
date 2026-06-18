"""
Exploratory analysis of HRSA HPSA and MUA/P (Medically Underserved Area/Population)
data across occupational domains (Primary Care, Dental Health, Mental Health).

HPSA detail rows are geographic *components* of a designation. MUA/P detail rows follow
a similar component structure (census tract, county, county subdivision). Boundary
shapefiles support point-in-polygon address checks.
"""

from __future__ import annotations

import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

try:
    import geopandas as gpd
    from shapely.geometry import Point

    HAS_GEOPANDAS = True
except ImportError:
    HAS_GEOPANDAS = False

# ---------------------------------------------------------------------------
# Paths and domain configuration
# ---------------------------------------------------------------------------
HPSA_BASE = Path(
    "/Users/karlhonerlaw/Library/CloudStorage/GoogleDrive-honerlaw@gmail.com"
    "/My Drive/CesarK_CPT_project/data/key_data_sources"
    "/HPSA MUAP Sites"
)
OUTPUT_DIR = HPSA_BASE / "output"
PILOT_SITES_PATH = Path(
    "/Users/karlhonerlaw/Library/CloudStorage/GoogleDrive-honerlaw@gmail.com"
    "/My Drive/CesarK_CPT_project/data/community_health_center_pilot_sites.xlsx"
)
FACILITY_MATCH_MAX_DISTANCE_M = 150

# HRSA layer prefixes: CMP = component polygons, PNT = facility points, PLY = designation polygons.
HPSA_DOMAINS: dict[str, dict[str, str]] = {
    "primary_care": {
        "label": "Primary Care",
        "folder": "HPSA – Primary Care",
        "detail_file": "BCD_HPSA_FCT_DET_PC.xlsx",
        "discipline_class": "Primary Care",
    },
    "dental": {
        "label": "Dental Health",
        "folder": "HPSA – Dental Health",
        "detail_file": "BCD_HPSA_FCT_DET_DH.xlsx",
        "discipline_class": "Dental Health",
    },
    "mental_health": {
        "label": "Mental Health",
        "folder": "HPSA – Mental Health",
        "detail_file": "BCD_HPSA_FCT_DET_MH.xlsx",
        "discipline_class": "Mental Health",
    },
}


@dataclass
class HPSADomainPaths:
    key: str
    label: str
    discipline_class: str
    detail_xlsx: Path
    boundaries_dir: Path
    area_component_shp: Path | None
    facility_shp: Path | None


@dataclass
class HPSADomainAssets:
    paths: HPSADomainPaths
    designated: pd.DataFrame
    sites: pd.DataFrame
    geography: pd.DataFrame
    tract_lookup: pd.DataFrame
    county_lookup: pd.DataFrame
    county_subdivision_lookup: pd.DataFrame
    facility_sites: pd.DataFrame
    area_boundaries: "gpd.GeoDataFrame | None" = None
    facility_boundaries: "gpd.GeoDataFrame | None" = None


def find_boundary_shp(boundaries_dir: Path, layer_prefix: str) -> Path | None:
    """
    Locate an HRSA boundary shapefile by layer prefix.

    Matches HRSA naming such as HPSA_CMPPC_SHP_DET_CUR_VX.shp (component),
    HPSA_PNTDH_SHP_DET_CUR_VX.shp (facility), HPSA_PLYMH_SHP_DET_CUR_VX.shp (designation).
    """
    if not boundaries_dir.exists():
        return None
    matches = sorted(boundaries_dir.rglob(f"HPSA_{layer_prefix}*_DET_CUR_VX.shp"))
    return matches[0] if matches else None


def get_domain_paths(domain_key: str) -> HPSADomainPaths:
    """Resolve spreadsheet and boundary paths for one HPSA occupational domain."""
    if domain_key not in HPSA_DOMAINS:
        raise KeyError(
            f"Unknown domain {domain_key!r}. Expected one of {list(HPSA_DOMAINS)}."
        )

    cfg = HPSA_DOMAINS[domain_key]
    domain_dir = HPSA_BASE / cfg["folder"]
    boundaries_dir = domain_dir / "boundaries"
    return HPSADomainPaths(
        key=domain_key,
        label=cfg["label"],
        discipline_class=cfg["discipline_class"],
        detail_xlsx=domain_dir / cfg["detail_file"],
        boundaries_dir=boundaries_dir,
        area_component_shp=find_boundary_shp(boundaries_dir, "CMP"),
        facility_shp=find_boundary_shp(boundaries_dir, "PNT"),
    )


MUA_DIR = HPSA_BASE / "Medically Underserved Areas Populations (MUA P)"
MUA_DET_XLSX = MUA_DIR / "MUA_DET.xlsx"
MUA_BOUNDARIES_SHP = MUA_DIR / "MUA_SHP/MUA_SHP_DET_CUR_VX.shp"
MUA_SITE_KEY = "mua_id"

MUA_SITE_COLUMNS = [
    "mua_id",
    "mua_area_code",
    "mua_service_area_name",
    "designation_type",
    "designation_type_code",
    "mua_status_code",
    "mua_status_description",
    "designation_date",
    "mua_update_date",
    "imu_score",
    "population_type",
    "mua_population_type_code",
    "mua_metropolitan_description",
    "rural_status_description",
    "primary_state_abbreviation",
    "primary_state_name",
    "primary_state_fips_code",
    "designation_population_in_a_mua",
    "mua_total_resident_civilian_population",
    "providers_per_1000_population",
    "ratio_of_providers_per_1000_population",
]

MUA_GEO_COLUMNS = [
    "mua_id",
    "mua_service_area_name",
    "designation_type",
    "mua_component_geographic_type_code",
    "mua_component_geographic_type_description",
    "mua_component_geographic_name",
    "census_tract",
    "tract_fips",
    "county_fips",
    "county_subdiv_fips",
    "state_and_county_federal_information_processing_standard_code",
    "common_state_fips_code",
    "common_state_name",
    "common_county_name",
    "common_state_county_fips_code",
    "primary_state_abbreviation",
]


@dataclass
class MUAPAssets:
    designated: pd.DataFrame
    sites: pd.DataFrame
    geography: pd.DataFrame
    tract_lookup: pd.DataFrame
    county_lookup: pd.DataFrame
    county_subdivision_lookup: pd.DataFrame
    boundaries: "gpd.GeoDataFrame | None" = None


def standardize_mua_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize MUA/P spreadsheet column names to snake_case aliases."""
    out = df.copy()
    out.columns = out.columns.str.lower().str.replace(" ", "_")
    renamed: dict[str, str] = {}
    for col in out.columns:
        new_col = col.replace("mua/p_", "mua_")
        new_col = new_col.replace("medically_underserved_area/population_(mua/p)_", "mua_")
        new_col = new_col.replace(
            "designation_population_in_a_medically_underserved_area/population_(mua/p)",
            "designation_population_in_a_mua",
        )
        new_col = new_col.replace(
            "medically_underserved_area/population_(mua/p)_withdrawal_date",
            "mua_withdrawal_date",
        )
        new_col = new_col.replace(
            "medically_underserved_area/population_(mua/p)_withdrawal_date_in_text_format",
            "mua_withdrawal_date_text",
        )
        renamed[col] = new_col
    return out.rename(columns=renamed)


def normalize_county_fips(value: Any) -> str | None:
    """Return a 5-digit county FIPS string, or None for invalid placeholders."""
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not re.fullmatch(r"\d+(\.0+)?", text):
        return None
    return str(int(float(text))).zfill(5)


def parse_census_tract_suffix(tract_raw: Any) -> str | None:
    """Convert HRSA tract notation (e.g. 4057.00) to a 6-digit tract suffix."""
    if pd.isna(tract_raw):
        return None
    text = str(tract_raw).strip()
    if text.lower() in {"not applicable", "na", "nan"}:
        return None
    if not re.match(r"^[\d.]+$", text):
        return None
    if "." in text:
        major, minor = text.split(".", 1)
        return major.zfill(4) + minor.ljust(2, "0")[:2].zfill(2)
    return text.replace(".", "").zfill(6)


def build_mua_tract_fips(row: pd.Series) -> str | None:
    """Build an 11-digit census tract FIPS from county + tract fields."""
    county5 = normalize_county_fips(
        row.get("state_and_county_federal_information_processing_standard_code")
    )
    suffix = parse_census_tract_suffix(row.get("census_tract"))
    if county5 and suffix:
        return county5 + suffix
    return None


def load_designated_mua(path: Path = MUA_DET_XLSX) -> pd.DataFrame:
    """Load MUA/P detail file and keep only Designated records."""
    df = standardize_mua_columns(pd.read_excel(path))
    designated = df[df["mua_status_description"] == "Designated"].copy()
    designated["tract_fips"] = designated.apply(build_mua_tract_fips, axis=1)
    designated["county_fips"] = designated[
        "state_and_county_federal_information_processing_standard_code"
    ].map(normalize_county_fips)
    designated["county_subdiv_fips"] = designated["county_subdivision_fips_code"].apply(
        lambda x: str(int(float(x)))
        if pd.notna(x) and re.fullmatch(r"\d+(\.0+)?", str(x).strip())
        else None
    )
    return designated


def profile_mua_dataset(df: pd.DataFrame) -> dict:
    """Return a structured profile of MUA/P row counts, keys, dates, and nulls."""
    date_cols = [
        c for c in df.columns if "date" in c and df[c].dtype != "datetime64[ns]"
    ]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    completeness = pd.DataFrame(
        {
            "column": df.columns,
            "dtype": df.dtypes.astype(str).values,
            "non_null": df.notna().sum().values,
            "null_pct": (100 * df.isna().mean()).round(1).values,
            "n_unique": df.nunique().values,
        }
    ).sort_values("null_pct", ascending=False)

    constant_cols = completeness.loc[completeness["n_unique"] == 1, "column"].tolist()
    rows_per_site = df.groupby(MUA_SITE_KEY).size()
    designation_counts = (
        df.groupby("designation_type")
        .agg(
            rows=(MUA_SITE_KEY, "size"),
            unique_sites=(MUA_SITE_KEY, "nunique"),
            unique_tracts=("tract_fips", "nunique"),
        )
        .sort_values("rows", ascending=False)
    )
    component_counts = df["mua_component_geographic_type_description"].value_counts()

    return {
        "raw_rows": len(df),
        "unique_sites": df[MUA_SITE_KEY].nunique(),
        "unique_service_area_names": df["mua_service_area_name"].nunique(),
        "unique_tract_fips": df["tract_fips"].nunique(),
        "rows_per_site": rows_per_site.describe().to_dict(),
        "designation_type_breakdown": designation_counts,
        "component_type_breakdown": component_counts,
        "date_ranges": {
            col: {"min": df[col].min(), "max": df[col].max(), "nulls": int(df[col].isna().sum())}
            for col in df.columns
            if "date" in col
        },
        "completeness": completeness,
        "constant_columns": constant_cols,
        "identifier_notes": {
            MUA_SITE_KEY: "Primary key for an MUA/P designation (unique within MUA/P data).",
            "mua_area_code": "Legacy/service area code; links to some shapefile MuaSrcID values.",
            "tract_fips": (
                "Derived 11-digit tract FIPS from county + census tract fields; "
                "may not match current Census geocoder county codes in all states."
            ),
            "MuaSrcID": "Shapefile source id (often zero-padded area code, not mua_id).",
        },
    }


def build_mua_site_table(df: pd.DataFrame) -> pd.DataFrame:
    """One row per unique MUA/P designation (mua_id)."""
    site_cols = [c for c in MUA_SITE_COLUMNS if c in df.columns]
    sites = (
        df.sort_values(MUA_SITE_KEY)
        .groupby(MUA_SITE_KEY, as_index=False)[site_cols]
        .first()
    )
    sites["n_geography_rows"] = df.groupby(MUA_SITE_KEY).size().values
    sites["n_census_tracts"] = (
        df.loc[df["mua_component_geographic_type_description"] == "Census Tract"]
        .groupby(MUA_SITE_KEY)
        .size()
        .reindex(sites[MUA_SITE_KEY], fill_value=0)
        .values
    )
    return sites.sort_values(["primary_state_abbreviation", "mua_service_area_name"])


def build_mua_geography_table(df: pd.DataFrame) -> pd.DataFrame:
    """One row per MUA/P geographic component."""
    geo_cols = [c for c in MUA_GEO_COLUMNS if c in df.columns]
    geo = df[geo_cols].drop_duplicates().copy()
    component = geo["mua_component_geographic_type_description"]
    geo["geo_id"] = None
    geo.loc[component == "Census Tract", "geo_id"] = geo.loc[
        component == "Census Tract", "tract_fips"
    ]
    geo.loc[component == "Single County", "geo_id"] = geo.loc[
        component == "Single County", "county_fips"
    ]
    geo.loc[component == "County Subdivision", "geo_id"] = geo.loc[
        component == "County Subdivision", "county_subdiv_fips"
    ]
    return geo.sort_values([MUA_SITE_KEY, "mua_component_geographic_type_description", "geo_id"])


def build_mua_tract_lookup(geography: pd.DataFrame) -> pd.DataFrame:
    lookup = _aggregate_hpsa_lookup(
        geography.dropna(subset=["tract_fips"]),
        "tract_fips",
        site_key=MUA_SITE_KEY,
        name_col="mua_service_area_name",
    )
    return lookup.rename(columns={"tract_fips": "census_tract_fips"})


def build_mua_county_lookup(geography: pd.DataFrame) -> pd.DataFrame:
    county_rows = geography.loc[
        geography["mua_component_geographic_type_description"] == "Single County"
    ]
    return _aggregate_hpsa_lookup(
        county_rows,
        "county_fips",
        site_key=MUA_SITE_KEY,
        name_col="mua_service_area_name",
    )


def build_mua_county_subdivision_lookup(geography: pd.DataFrame) -> pd.DataFrame:
    subdiv_rows = geography.loc[
        geography["mua_component_geographic_type_description"] == "County Subdivision"
    ]
    return _aggregate_hpsa_lookup(
        subdiv_rows,
        "county_subdiv_fips",
        site_key=MUA_SITE_KEY,
        name_col="mua_service_area_name",
    )


def load_mua_boundaries(path: Path = MUA_BOUNDARIES_SHP) -> "gpd.GeoDataFrame":
    """Load designated MUA/P designation polygons (MUA_SHP layer)."""
    if not HAS_GEOPANDAS:
        raise ImportError("geopandas is required for boundary-based MUA/P checks")
    gdf = gpd.read_file(path)
    return gdf.loc[gdf["MuaStatCD"] == "D"].copy()


def build_mua_assets(*, load_boundaries: bool = True) -> MUAPAssets:
    """Load MUA/P spreadsheets, lookup tables, and optional boundary polygons."""
    if not MUA_DET_XLSX.exists():
        raise FileNotFoundError(f"MUA/P detail file not found: {MUA_DET_XLSX}")

    designated = load_designated_mua(MUA_DET_XLSX)
    geography = build_mua_geography_table(designated)
    boundaries = None
    if load_boundaries and HAS_GEOPANDAS and MUA_BOUNDARIES_SHP.exists():
        boundaries = load_mua_boundaries(MUA_BOUNDARIES_SHP)

    return MUAPAssets(
        designated=designated,
        sites=build_mua_site_table(designated),
        geography=geography,
        tract_lookup=build_mua_tract_lookup(geography),
        county_lookup=build_mua_county_lookup(geography),
        county_subdivision_lookup=build_mua_county_subdivision_lookup(geography),
        boundaries=boundaries,
    )


def write_mua_outputs(assets: MUAPAssets, profile: dict) -> Path:
    """Write MUA/P CSV outputs and return the output directory."""
    mua_out = OUTPUT_DIR / "mua"
    mua_out.mkdir(parents=True, exist_ok=True)
    assets.sites.to_csv(mua_out / "mua_sites.csv", index=False)
    assets.geography.to_csv(mua_out / "mua_geography.csv", index=False)
    assets.tract_lookup.to_csv(mua_out / "mua_tract_lookup.csv", index=False)
    assets.county_lookup.to_csv(mua_out / "mua_county_lookup.csv", index=False)
    assets.county_subdivision_lookup.to_csv(
        mua_out / "mua_county_subdivision_lookup.csv", index=False
    )
    profile["completeness"].to_csv(mua_out / "column_completeness.csv", index=False)
    profile["designation_type_breakdown"].to_csv(
        mua_out / "designation_type_breakdown.csv"
    )
    return mua_out


def print_mua_profile_summary(profile: dict) -> None:
    """Print a human-readable MUA/P profile summary."""
    print("=" * 72)
    print("MUA/P — Medically Underserved Areas/Populations (Designated records profile)")
    print("=" * 72)
    print(f"Detail rows (geography components): {profile['raw_rows']:,}")
    print(f"Unique MUA/P designations (mua_id): {profile['unique_sites']:,}")
    print(f"Unique service area names:          {profile['unique_service_area_names']:,}")
    print(f"Unique tract FIPS (derived):        {profile['unique_tract_fips']:,}")
    print()
    print("Rows per mua_id:")
    for stat, val in profile["rows_per_site"].items():
        print(f"  {stat:>6}: {val:,.1f}" if isinstance(val, float) else f"  {stat:>6}: {val}")
    print()
    print("By designation type:")
    print(profile["designation_type_breakdown"].to_string())
    print()
    print("By geography component type:")
    print(profile["component_type_breakdown"].to_string())
    print()
    print("Date ranges:")
    for col, info in profile["date_ranges"].items():
        print(f"  {col}: {info['min']} → {info['max']}  (nulls: {info['nulls']})")
    print()
    print(f"Constant columns ({len(profile['constant_columns'])}): "
          f"{', '.join(profile['constant_columns'])}")
    print()
    high_null = profile["completeness"].loc[profile["completeness"]["null_pct"] > 50]
    print(f"Columns with >50% nulls ({len(high_null)}):")
    print(high_null[["column", "null_pct", "n_unique"]].head(15).to_string(index=False))
    print()
    print("Identifier notes:")
    for key, note in profile["identifier_notes"].items():
        print(f"  {key}: {note}")


def _lookup_mua_matches(
    lookup: pd.DataFrame,
    geo_col: str,
    geo_id: str | None,
) -> dict[str, Any]:
    """Return MUA/P matches for one geography id from a lookup table."""
    if not geo_id or lookup is None:
        return _empty_hpsa_match()

    if geo_col not in lookup.columns:
        return _empty_hpsa_match()

    hit = lookup.loc[lookup[geo_col] == geo_id]
    if hit.empty:
        return _empty_hpsa_match()

    row = hit.iloc[0]
    return {
        "hpsa_ids": row["hpsa_ids"],
        "hpsa_names": row["hpsa_names"],
        "designation_types": row["designation_types"],
    }


def _mua_matches_from_boundary_hits(hits: "gpd.GeoDataFrame") -> dict[str, Any]:
    if hits.empty:
        return _empty_hpsa_match()
    deduped = hits.drop_duplicates(subset=["MuaSrcID"])
    return {
        "hpsa_ids": deduped["MuaSrcID"].tolist(),
        "hpsa_names": deduped["MuaSvcArNM"].tolist(),
        "designation_types": deduped["MuaDgnTypD"].tolist(),
    }


def match_point_to_mua_boundaries(
    longitude: float | None,
    latitude: float | None,
    boundaries: "gpd.GeoDataFrame",
    *,
    state: str | None = None,
) -> dict[str, Any]:
    """Return MUA/P matches for a geocoded point using HRSA MUA_SHP polygons."""
    if longitude is None or latitude is None:
        return _empty_hpsa_match()

    polys = boundaries
    if state:
        polys = boundaries.loc[boundaries["PriStAbbr"].str.upper() == state.upper()]

    point = _point_geodataframe(longitude, latitude)
    hits = gpd.sjoin(point, polys, predicate="within")
    return _mua_matches_from_boundary_hits(hits)


def check_address_in_mua(
    address: str,
    city: str,
    state: str,
    zip_code: str,
    *,
    assets: MUAPAssets,
    geocode: bool = True,
    geocode_result: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Determine whether an address falls within a designated MUA/P.

    Uses MUA_SHP polygons when available (preferred). Falls back to geocoded
    tract/county lookup tables only when the shapefile is not loaded.
    """
    geo = geocode_result or (
        geocode_address(address, city, state, zip_code) if geocode else {}
    )

    match_methods: list[str] = []

    if assets.boundaries is not None:
        final_match = match_point_to_mua_boundaries(
            geo.get("longitude"),
            geo.get("latitude"),
            assets.boundaries,
            state=state,
        )
        if final_match["hpsa_ids"]:
            match_methods.append("mua_polygon")
    else:
        tract_match = _lookup_mua_matches(
            assets.tract_lookup, "census_tract_fips", geo.get("census_tract_fips")
        )
        county_match = _lookup_mua_matches(
            assets.county_lookup, "county_fips", geo.get("county_fips")
        )
        subdiv_match = _lookup_mua_matches(
            assets.county_subdivision_lookup,
            "county_subdiv_fips",
            geo.get("county_subdivision_fips"),
        )
        final_match = _merge_hpsa_matches(tract_match, county_match, subdiv_match)
        if tract_match["hpsa_ids"]:
            match_methods.append("census_tract")
        if county_match["hpsa_ids"]:
            match_methods.append("county")
        if subdiv_match["hpsa_ids"]:
            match_methods.append("county_subdivision")

    return {
        "in_mua": bool(final_match["hpsa_ids"]),
        "match_methods": match_methods,
        "matched_mua_ids": final_match["hpsa_ids"],
        "matched_mua_names": final_match["hpsa_names"],
        "matched_designation_types": final_match["designation_types"],
    }


CENSUS_GEOCODER_URL = (
    "https://geocoding.geo.census.gov/geocoder/geographies/onelineaddress"
)
GEOCODER_BENCHMARK = "Public_AR_Current"
GEOCODER_VINTAGE = "Current_Current"
GEOCODER_SLEEP_SECONDS = 0.25

SITE_KEY = "hpsa_id"

# Site-level attributes (constant within each hpsa_id).
SITE_COLUMNS = [
    "hpsa_id",
    "hpsa_name",
    "designation_type",
    "hpsa_discipline_class",
    "hpsa_score",
    "pc_mcta_score",
    "hpsa_status",
    "hpsa_designation_date",
    "hpsa_designation_last_update_date",
    "metropolitan_indicator",
    "hpsa_degree_of_shortage",
    "hpsa_fte",
    "hpsa_designation_population",
    "%_of_population_below_100%_poverty",
    "hpsa_formal_ratio",
    "hpsa_population_type",
    "rural_status",
    "primary_state_abbreviation",
    "primary_state_name",
    "primary_state_fips_code",
    "provider_type",
    "hpsa_shortage",
    "hpsa_estimated_served_population",
    "hpsa_estimated_underserved_population",
    # Facility-only fields (populated when geo id == "POINT")
    "hpsa_address",
    "hpsa_city",
    "hpsa_postal_code",
    "longitude",
    "latitude",
    "bhcmis_organization_identification_number",
    "hpsa_component_source_identification_number",
]

GEO_COLUMNS = [
    "hpsa_id",
    "hpsa_name",
    "designation_type",
    "hpsa_component_type_code",
    "hpsa_component_type_description",
    "hpsa_component_name",
    "hpsa_geography_identification_number",
    "state_and_county_federal_information_processing_standard_code",
    "common_state_fips_code",
    "common_state_name",
    "common_county_name",
    "common_state_county_fips_code",
    "primary_state_abbreviation",
]


def load_designated_hpsa(path: Path) -> pd.DataFrame:
    """Load a domain detail file and keep only Designated HPSA records."""
    df = pd.read_excel(path)
    df.columns = df.columns.str.lower().str.replace(" ", "_")
    designated = df[df["hpsa_status"] == "Designated"].copy()
    return designated


def build_domain_assets(
    domain_key: str,
    *,
    load_boundaries: bool = True,
) -> HPSADomainAssets:
    """Load spreadsheets, lookup tables, and optional boundary layers for one domain."""
    paths = get_domain_paths(domain_key)
    if not paths.detail_xlsx.exists():
        raise FileNotFoundError(f"Detail file not found: {paths.detail_xlsx}")

    designated = load_designated_hpsa(paths.detail_xlsx)
    sites = build_site_table(designated)
    geography = build_geography_table(designated)

    area_boundaries = None
    facility_boundaries = None
    if load_boundaries and HAS_GEOPANDAS:
        if paths.area_component_shp and paths.area_component_shp.exists():
            area_boundaries = load_area_hpsa_boundaries(
                paths.area_component_shp, paths.discipline_class
            )
        if paths.facility_shp and paths.facility_shp.exists():
            facility_boundaries = load_facility_hpsa_boundaries(
                paths.facility_shp, paths.discipline_class
            )

    return HPSADomainAssets(
        paths=paths,
        designated=designated,
        sites=sites,
        geography=geography,
        tract_lookup=build_tract_lookup(geography),
        county_lookup=build_county_lookup(geography),
        county_subdivision_lookup=build_county_subdivision_lookup(geography),
        facility_sites=sites.loc[sites["is_facility_designation"]].copy(),
        area_boundaries=area_boundaries,
        facility_boundaries=facility_boundaries,
    )


def write_domain_outputs(assets: HPSADomainAssets, profile: dict) -> Path:
    """Write per-domain CSV outputs and return the output directory."""
    domain_out = OUTPUT_DIR / assets.paths.key
    domain_out.mkdir(parents=True, exist_ok=True)

    assets.sites.to_csv(domain_out / "hpsa_sites.csv", index=False)
    assets.geography.to_csv(domain_out / "hpsa_geography.csv", index=False)
    assets.tract_lookup.to_csv(domain_out / "hpsa_tract_lookup.csv", index=False)
    assets.county_lookup.to_csv(domain_out / "hpsa_county_lookup.csv", index=False)
    assets.county_subdivision_lookup.to_csv(
        domain_out / "hpsa_county_subdivision_lookup.csv", index=False
    )
    profile["completeness"].to_csv(domain_out / "column_completeness.csv", index=False)
    profile["designation_type_breakdown"].to_csv(
        domain_out / "designation_type_breakdown.csv"
    )
    return domain_out


def profile_dataset(df: pd.DataFrame) -> dict:
    """Return a structured profile of row counts, keys, dates, and nulls."""
    date_cols = [
        c for c in df.columns if "date" in c and df[c].dtype != "datetime64[ns]"
    ]
    for col in date_cols:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    completeness = pd.DataFrame(
        {
            "column": df.columns,
            "dtype": df.dtypes.astype(str).values,
            "non_null": df.notna().sum().values,
            "null_pct": (100 * df.isna().mean()).round(1).values,
            "n_unique": df.nunique().values,
        }
    ).sort_values("null_pct", ascending=False)

    constant_cols = completeness.loc[completeness["n_unique"] == 1, "column"].tolist()

    rows_per_site = df.groupby(SITE_KEY).size()
    designation_counts = (
        df.groupby("designation_type")
        .agg(
            rows=(SITE_KEY, "size"),
            unique_sites=(SITE_KEY, "nunique"),
            unique_geo_ids=("hpsa_geography_identification_number", "nunique"),
        )
        .sort_values("rows", ascending=False)
    )

    component_counts = df["hpsa_component_type_description"].value_counts()

    return {
        "raw_rows": len(df),
        "unique_sites": df[SITE_KEY].nunique(),
        "unique_hpsa_names": df["hpsa_name"].nunique(),
        "unique_geo_components": df["hpsa_geography_identification_number"].nunique(),
        "rows_per_site": rows_per_site.describe().to_dict(),
        "designation_type_breakdown": designation_counts,
        "component_type_breakdown": component_counts,
        "date_ranges": {
            col: {"min": df[col].min(), "max": df[col].max(), "nulls": int(df[col].isna().sum())}
            for col in df.columns
            if "date" in col
        },
        "completeness": completeness,
        "constant_columns": constant_cols,
        "identifier_notes": {
            SITE_KEY: "Primary key for an HPSA designation (site or area). One name per id.",
            "hpsa_geography_identification_number": (
                "Component geography: 11-digit census tract FIPS, 5-digit county FIPS, "
                "10-digit county subdivision FIPS, or 'POINT' for facility designations."
            ),
            "hpsa_component_source_identification_number": (
                "HRSA component source id; 1:1 with hpsa_id for facility (POINT) rows."
            ),
            "bhcmis_organization_identification_number": (
                "Bureau of Primary Health Care org id; present for some FQHC rows only."
            ),
        },
    }


def build_site_table(df: pd.DataFrame) -> pd.DataFrame:
    """One row per unique HPSA designation (hpsa_id)."""
    site_cols = [c for c in SITE_COLUMNS if c in df.columns]
    sites = (
        df.sort_values(SITE_KEY)
        .groupby(SITE_KEY, as_index=False)[site_cols]
        .first()
    )
    sites["n_geography_rows"] = df.groupby(SITE_KEY).size().values
    sites["n_census_tracts"] = (
        df.loc[df["hpsa_component_type_description"] == "Census Tract"]
        .groupby(SITE_KEY)
        .size()
        .reindex(sites[SITE_KEY], fill_value=0)
        .values
    )
    sites["is_facility_designation"] = sites["designation_type"].isin(
        [
            "Rural Health Clinic",
            "Federally Qualified Health Center",
            "Federally Qualified Health Center Look Alike",
            "Indian Health Service, Tribal Health, and Urban Indian Health Organizations",
            "Correctional Facility",
            "Other Facility",
        ]
    )
    return sites.sort_values(["primary_state_abbreviation", "hpsa_name"])


def build_geography_table(df: pd.DataFrame) -> pd.DataFrame:
    """One row per geographic component (tract, county, subdivision, or POINT)."""
    geo_cols = [c for c in GEO_COLUMNS if c in df.columns]
    geo = df[geo_cols].drop_duplicates().copy()
    geo = geo.rename(
        columns={"hpsa_geography_identification_number": "geo_id_raw"}
    )
    geo["geo_id"] = geo["geo_id_raw"].astype(str)

    component = geo["hpsa_component_type_description"]
    geo["census_tract_fips"] = geo["geo_id"].where(component == "Census Tract")
    geo["county_fips"] = geo["geo_id"].where(component == "Single County")
    geo["county_subdivision_fips"] = geo["geo_id"].where(
        component == "County Subdivision"
    )
    geo["is_point_facility"] = geo["geo_id"] == "POINT"

    return geo.sort_values([SITE_KEY, "hpsa_component_type_description", "geo_id"])


def _aggregate_hpsa_lookup(
    geography: pd.DataFrame,
    geo_col: str,
    *,
    site_key: str = SITE_KEY,
    name_col: str = "hpsa_name",
) -> pd.DataFrame:
    """Build a geo-id → designation lookup for one geography level."""
    rows = geography.dropna(subset=[geo_col]).copy()
    return (
        rows.groupby(geo_col, as_index=False)
        .agg(
            n_hpsa_sites=(site_key, "nunique"),
            hpsa_ids=(site_key, lambda s: sorted(s.unique().tolist())),
            hpsa_names=(name_col, lambda s: sorted(s.unique().tolist())),
            designation_types=("designation_type", lambda s: sorted(s.unique().tolist())),
        )
        .sort_values(geo_col)
    )


def build_tract_lookup(geography: pd.DataFrame) -> pd.DataFrame:
    """Tract FIPS → HPSA designations for joining external address/tract data."""
    return _aggregate_hpsa_lookup(geography, "census_tract_fips")


def build_county_lookup(geography: pd.DataFrame) -> pd.DataFrame:
    """5-digit county FIPS → HPSA designations."""
    return _aggregate_hpsa_lookup(geography, "county_fips")


def build_county_subdivision_lookup(geography: pd.DataFrame) -> pd.DataFrame:
    """10-digit county subdivision FIPS → HPSA designations."""
    return _aggregate_hpsa_lookup(geography, "county_subdivision_fips")


def load_pilot_sites(path: Path = PILOT_SITES_PATH) -> pd.DataFrame:
    """Load community health center pilot site addresses."""
    df = pd.read_excel(path)
    df.columns = df.columns.str.lower().str.replace(" ", "_")
    required = {"address", "city", "state", "zip"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Pilot sites file missing columns: {sorted(missing)}")
    for col in required:
        df[col] = df[col].astype(str).str.strip()
    df["zip"] = df["zip"].str.replace(r"\.0$", "", regex=True).str.zfill(5)
    return df.reset_index(drop=True)


def format_oneline_address(
    address: str, city: str, state: str, zip_code: str
) -> str:
    """Build a single-line address string for the Census geocoder."""
    zip_code = re.sub(r"\.0$", "", str(zip_code).strip()).zfill(5)
    return f"{address}, {city}, {state} {zip_code}"


def geocode_address(
    address: str,
    city: str,
    state: str,
    zip_code: str,
    *,
    sleep_seconds: float = GEOCODER_SLEEP_SECONDS,
) -> dict[str, Any]:
    """
    Geocode a US address via the Census Bureau geographies API.

    Returns census tract, county, and county-subdivision FIPS when matched.
    """
    oneline = format_oneline_address(address, city, state, zip_code)
    params = urllib.parse.urlencode(
        {
            "address": oneline,
            "benchmark": GEOCODER_BENCHMARK,
            "vintage": GEOCODER_VINTAGE,
            "format": "json",
        }
    )
    url = f"{CENSUS_GEOCODER_URL}?{params}"

    try:
        with urllib.request.urlopen(url, timeout=30) as response:
            payload = json.load(response)
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError) as exc:
        return {
            "geocode_success": False,
            "geocode_error": str(exc),
            "geocode_input": oneline,
            "geocode_matched_address": None,
            "latitude": None,
            "longitude": None,
            "census_tract_fips": None,
            "county_fips": None,
            "county_subdivision_fips": None,
        }

    if sleep_seconds:
        time.sleep(sleep_seconds)

    matches = payload.get("result", {}).get("addressMatches", [])
    if not matches:
        return {
            "geocode_success": False,
            "geocode_error": "No address match returned by Census geocoder",
            "geocode_input": oneline,
            "geocode_matched_address": None,
            "latitude": None,
            "longitude": None,
            "census_tract_fips": None,
            "county_fips": None,
            "county_subdivision_fips": None,
        }

    match = matches[0]
    geographies = match.get("geographies", {})
    coordinates = match.get("coordinates", {})

    def first_geoid(geo_type: str) -> str | None:
        items = geographies.get(geo_type, [])
        return items[0]["GEOID"] if items else None

    return {
        "geocode_success": True,
        "geocode_error": None,
        "geocode_input": oneline,
        "geocode_matched_address": match.get("matchedAddress"),
        "latitude": coordinates.get("y"),
        "longitude": coordinates.get("x"),
        "census_tract_fips": first_geoid("Census Tracts"),
        "county_fips": first_geoid("Counties"),
        "county_subdivision_fips": first_geoid("County Subdivisions"),
    }


def _lookup_hpsa_matches(
    lookup: pd.DataFrame,
    geo_col: str,
    geo_id: str | None,
) -> dict[str, Any]:
    """Return HPSA ids/names/types for a single geography id, or empty lists."""
    if not geo_id or lookup is None:
        return {"hpsa_ids": [], "hpsa_names": [], "designation_types": []}

    hit = lookup.loc[lookup[geo_col] == geo_id]
    if hit.empty:
        return {"hpsa_ids": [], "hpsa_names": [], "designation_types": []}

    row = hit.iloc[0]
    return {
        "hpsa_ids": row["hpsa_ids"],
        "hpsa_names": row["hpsa_names"],
        "designation_types": row["designation_types"],
    }


def _merge_hpsa_matches(*match_dicts: dict[str, Any]) -> dict[str, Any]:
    """Combine HPSA matches from multiple geography levels, deduplicated by id."""
    ids: list[Any] = []
    names_by_id: dict[Any, str] = {}
    types_by_id: dict[Any, str] = {}

    for match in match_dicts:
        for hpsa_id, hpsa_name, designation_type in zip(
            match["hpsa_ids"], match["hpsa_names"], match["designation_types"]
        ):
            if hpsa_id not in names_by_id:
                ids.append(hpsa_id)
            names_by_id[hpsa_id] = hpsa_name
            types_by_id[hpsa_id] = designation_type

    return {
        "hpsa_ids": ids,
        "hpsa_names": [names_by_id[i] for i in ids],
        "designation_types": [types_by_id[i] for i in ids],
    }


def _normalize_address(value: str) -> str:
    """Normalize an address string for fuzzy facility matching."""
    value = value.lower()
    value = re.sub(r"[^\w\s]", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _match_hpsa_facility(
    address: str,
    city: str,
    state: str,
    zip_code: str,
    facility_sites: pd.DataFrame,
) -> dict[str, Any]:
    """Match against directly designated HPSA facility rows (POINT designations)."""
    target = _normalize_address(
        f"{address} {city} {state} {str(zip_code).zfill(5)}"
    )
    target_zip = str(zip_code).zfill(5)

    candidates = facility_sites.copy()
    candidates["zip5"] = (
        candidates["hpsa_postal_code"]
        .astype(str)
        .str.extract(r"(\d{5})", expand=False)
    )
    candidates = candidates.loc[candidates["zip5"] == target_zip]
    if candidates.empty:
        return {"hpsa_ids": [], "hpsa_names": [], "designation_types": []}

    candidates["norm_address"] = candidates.apply(
        lambda row: _normalize_address(
            f"{row.get('hpsa_address', '')} {row.get('hpsa_city', '')} "
            f"{row.get('primary_state_abbreviation', '')} {row.get('zip5', '')}"
        ),
        axis=1,
    )

    exact = candidates.loc[candidates["norm_address"] == target]
    if exact.empty:
        # Allow partial match when the normalized street/city appears in facility text.
        partial = candidates.loc[
            candidates["norm_address"].str.contains(
                _normalize_address(f"{address} {city}"), regex=False
            )
        ]
        hit = partial
    else:
        hit = exact

    if hit.empty:
        return {"hpsa_ids": [], "hpsa_names": [], "designation_types": []}

    return {
        "hpsa_ids": hit[SITE_KEY].tolist(),
        "hpsa_names": hit["hpsa_name"].tolist(),
        "designation_types": hit["designation_type"].tolist(),
    }


def _empty_hpsa_match() -> dict[str, list[Any]]:
    return {"hpsa_ids": [], "hpsa_names": [], "designation_types": []}


def load_area_hpsa_boundaries(
    path: Path,
    discipline_class: str,
) -> "gpd.GeoDataFrame":
    """Load designated area HPSA component polygons for one occupational domain."""
    if not HAS_GEOPANDAS:
        raise ImportError("geopandas is required for boundary-based HPSA checks")
    gdf = gpd.read_file(path)
    return gdf.loc[
        (gdf["HpsStatCD"] == "D") & (gdf["DscpClsDes"] == discipline_class)
    ].copy()


def load_facility_hpsa_boundaries(
    path: Path,
    discipline_class: str,
) -> "gpd.GeoDataFrame":
    """Load designated facility HPSA point locations for one occupational domain."""
    if not HAS_GEOPANDAS:
        raise ImportError("geopandas is required for boundary-based HPSA checks")
    gdf = gpd.read_file(path)
    return gdf.loc[
        (gdf["HpsStatCD"] == "D") & (gdf["DscpClsDes"] == discipline_class)
    ].copy()


def _point_geodataframe(longitude: float, latitude: float) -> "gpd.GeoDataFrame":
    return gpd.GeoDataFrame(
        {"id": [1]},
        geometry=[Point(longitude, latitude)],
        crs="EPSG:4326",
    )


def _boundary_matches_from_hits(hits: "gpd.GeoDataFrame") -> dict[str, Any]:
    """Collapse spatial join hits to unique HPSA designations."""
    if hits.empty:
        return _empty_hpsa_match()

    deduped = hits.drop_duplicates(subset=["HpsSrcID"])
    return {
        "hpsa_ids": deduped["HpsSrcID"].tolist(),
        "hpsa_names": deduped["HpsNM"].tolist(),
        "designation_types": deduped["HpsTypDes"].tolist(),
    }


def match_point_to_area_boundaries(
    longitude: float | None,
    latitude: float | None,
    area_boundaries: "gpd.GeoDataFrame",
    *,
    state: str | None = None,
) -> dict[str, Any]:
    """Return area HPSA matches for a geocoded point using HRSA polygons."""
    if longitude is None or latitude is None:
        return _empty_hpsa_match()

    boundaries = area_boundaries
    if state:
        boundaries = area_boundaries.loc[
            area_boundaries["StAbbr"].str.upper() == state.upper()
        ]

    point = _point_geodataframe(longitude, latitude)
    hits = gpd.sjoin(point, boundaries, predicate="within")
    return _boundary_matches_from_hits(hits)


def match_point_to_facility_boundaries(
    longitude: float | None,
    latitude: float | None,
    facility_boundaries: "gpd.GeoDataFrame",
    *,
    address: str,
    city: str,
    state: str,
    zip_code: str,
    max_distance_m: float = FACILITY_MATCH_MAX_DISTANCE_M,
) -> dict[str, Any]:
    """
    Match a geocoded point to facility HPSA locations.

    Uses normalized address text first, then nearest facility point within
    ``max_distance_m`` meters.
    """
    if facility_boundaries.empty:
        return _empty_hpsa_match()

    target = _normalize_address(
        f"{address} {city} {state} {str(zip_code).zfill(5)}"
    )
    target_zip = str(zip_code).zfill(5)

    candidates = facility_boundaries.copy()
    candidates["zip5"] = (
        candidates["HpsZipCD"].astype(str).str.extract(r"(\d{5})", expand=False)
    )
    candidates = candidates.loc[candidates["zip5"] == target_zip]
    if not candidates.empty:
        candidates["norm_address"] = candidates.apply(
            lambda row: _normalize_address(
                f"{row.get('HpsAddr', '')} {row.get('HpsCity', '')} "
                f"{row.get('HpsStAbbr', '')} {row.get('zip5', '')}"
            ),
            axis=1,
        )
        address_hit = candidates.loc[
            (candidates["norm_address"] == target)
            | candidates["norm_address"].str.contains(
                _normalize_address(f"{address} {city}"), regex=False
            )
        ]
        if not address_hit.empty:
            return {
                "hpsa_ids": address_hit["HpsSrcID"].tolist(),
                "hpsa_names": address_hit["HpsNM"].tolist(),
                "designation_types": address_hit["HpsTypDes"].tolist(),
            }

    if longitude is None or latitude is None:
        return _empty_hpsa_match()

    point = _point_geodataframe(longitude, latitude)
    projected_point = point.to_crs("EPSG:5070")
    projected_facilities = facility_boundaries.to_crs("EPSG:5070")
    if state:
        projected_facilities = projected_facilities.loc[
            projected_facilities["HpsStAbbr"].str.upper() == state.upper()
        ]

    nearest = gpd.sjoin_nearest(
        projected_point,
        projected_facilities,
        max_distance=max_distance_m,
        distance_col="distance_m",
    )
    return _boundary_matches_from_hits(nearest)


def check_address_in_hpsa(
    address: str,
    city: str,
    state: str,
    zip_code: str,
    *,
    tract_lookup: pd.DataFrame | None = None,
    county_lookup: pd.DataFrame | None = None,
    county_subdivision_lookup: pd.DataFrame | None = None,
    facility_sites: pd.DataFrame | None = None,
    area_boundaries: "gpd.GeoDataFrame | None" = None,
    facility_boundaries: "gpd.GeoDataFrame | None" = None,
    geocode: bool = True,
    geocode_result: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Determine whether an address falls within a designated HPSA for one domain.

    Uses HRSA boundary shapefiles when available (preferred). Falls back to
    geocoded tract/county lookup tables only when shapefiles are not loaded.
    """
    geo = geocode_result or (
        geocode_address(address, city, state, zip_code) if geocode else {}
    )

    match_methods: list[str] = []

    if area_boundaries is not None:
        area_match = match_point_to_area_boundaries(
            geo.get("longitude"),
            geo.get("latitude"),
            area_boundaries,
            state=state,
        )
        if area_match["hpsa_ids"]:
            match_methods.append("area_polygon")
    else:
        tract_match = _lookup_hpsa_matches(
            tract_lookup, "census_tract_fips", geo.get("census_tract_fips")
        )
        county_match = _lookup_hpsa_matches(
            county_lookup, "county_fips", geo.get("county_fips")
        )
        subdiv_match = _lookup_hpsa_matches(
            county_subdivision_lookup,
            "county_subdivision_fips",
            geo.get("county_subdivision_fips"),
        )
        area_match = _merge_hpsa_matches(tract_match, county_match, subdiv_match)
        if tract_match["hpsa_ids"]:
            match_methods.append("census_tract")
        if county_match["hpsa_ids"]:
            match_methods.append("county")
        if subdiv_match["hpsa_ids"]:
            match_methods.append("county_subdivision")

    if facility_boundaries is not None:
        facility_match = match_point_to_facility_boundaries(
            geo.get("longitude"),
            geo.get("latitude"),
            facility_boundaries,
            address=address,
            city=city,
            state=state,
            zip_code=zip_code,
        )
        if facility_match["hpsa_ids"]:
            match_methods.append("facility_point")
    elif facility_sites is not None:
        facility_match = _match_hpsa_facility(
            address, city, state, zip_code, facility_sites
        )
        if facility_match["hpsa_ids"]:
            match_methods.append("facility_spreadsheet")
    else:
        facility_match = _empty_hpsa_match()

    all_match = _merge_hpsa_matches(area_match, facility_match)

    return {
        **geo,
        "in_hpsa": bool(all_match["hpsa_ids"]),
        "in_hpsa_area": bool(area_match["hpsa_ids"]),
        "is_hpsa_facility": bool(facility_match["hpsa_ids"]),
        "match_methods": match_methods,
        "matched_hpsa_ids": all_match["hpsa_ids"],
        "matched_hpsa_names": all_match["hpsa_names"],
        "matched_designation_types": all_match["designation_types"],
    }


def check_addresses_in_hpsa(
    addresses: pd.DataFrame,
    *,
    tract_lookup: pd.DataFrame,
    county_lookup: pd.DataFrame,
    county_subdivision_lookup: pd.DataFrame,
    facility_sites: pd.DataFrame,
    area_boundaries: "gpd.GeoDataFrame | None" = None,
    facility_boundaries: "gpd.GeoDataFrame | None" = None,
) -> pd.DataFrame:
    """Run HPSA checks for every row in an address dataframe."""
    results = []
    for _, row in addresses.iterrows():
        results.append(
            check_address_in_hpsa(
                row["address"],
                row["city"],
                row["state"],
                row["zip"],
                tract_lookup=tract_lookup,
                county_lookup=county_lookup,
                county_subdivision_lookup=county_subdivision_lookup,
                facility_sites=facility_sites,
                area_boundaries=area_boundaries,
                facility_boundaries=facility_boundaries,
            )
        )
    result_df = pd.DataFrame(results)
    return pd.concat([addresses.reset_index(drop=True), result_df], axis=1)


def check_addresses_all_domains(
    addresses: pd.DataFrame,
    domain_assets: dict[str, HPSADomainAssets],
    mua_assets: MUAPAssets | None = None,
) -> pd.DataFrame:
    """
    Geocode each address once, then check HPSA and MUA/P membership.

    Returns a wide dataframe with per-domain ``in_hpsa_{domain_key}`` columns and,
    when ``mua_assets`` is provided, ``in_mua`` / ``matched_mua_names`` columns.
    """
    rows: list[dict[str, Any]] = []

    for _, row in addresses.iterrows():
        geo = geocode_address(row["address"], row["city"], row["state"], row["zip"])
        record: dict[str, Any] = {
            "address": row["address"],
            "city": row["city"],
            "state": row["state"],
            "zip": row["zip"],
            "geocode_success": geo.get("geocode_success"),
            "geocode_matched_address": geo.get("geocode_matched_address"),
            "census_tract_fips": geo.get("census_tract_fips"),
            "county_fips": geo.get("county_fips"),
            "latitude": geo.get("latitude"),
            "longitude": geo.get("longitude"),
        }

        for domain_key, assets in domain_assets.items():
            check = check_address_in_hpsa(
                row["address"],
                row["city"],
                row["state"],
                row["zip"],
                tract_lookup=assets.tract_lookup,
                county_lookup=assets.county_lookup,
                county_subdivision_lookup=assets.county_subdivision_lookup,
                facility_sites=assets.facility_sites,
                area_boundaries=assets.area_boundaries,
                facility_boundaries=assets.facility_boundaries,
                geocode=False,
                geocode_result=geo,
            )
            record[f"in_hpsa_{domain_key}"] = check["in_hpsa"]
            record[f"in_hpsa_area_{domain_key}"] = check["in_hpsa_area"]
            record[f"is_hpsa_facility_{domain_key}"] = check["is_hpsa_facility"]
            record[f"matched_hpsa_names_{domain_key}"] = check["matched_hpsa_names"]
            record[f"match_methods_{domain_key}"] = check["match_methods"]

        record["in_hpsa_any"] = any(
            record[f"in_hpsa_{domain_key}"] for domain_key in domain_assets
        )

        if mua_assets is not None:
            mua_check = check_address_in_mua(
                row["address"],
                row["city"],
                row["state"],
                row["zip"],
                assets=mua_assets,
                geocode=False,
                geocode_result=geo,
            )
            record["in_mua"] = mua_check["in_mua"]
            record["matched_mua_names"] = mua_check["matched_mua_names"]
            record["match_methods_mua"] = mua_check["match_methods"]

        record["in_hpsa_or_mua"] = record["in_hpsa_any"] or (
            mua_assets is not None and record.get("in_mua", False)
        )
        rows.append(record)

    return pd.DataFrame(rows)


def print_profile_summary(profile: dict, *, domain_label: str) -> None:
    """Print a human-readable summary to stdout."""
    print("=" * 72)
    print(f"HPSA — {domain_label} (Designated records profile)")
    print("=" * 72)
    print(f"Detail rows (geography components): {profile['raw_rows']:,}")
    print(f"Unique HPSA designations (hpsa_id):  {profile['unique_sites']:,}")
    print(f"Unique HPSA names:                   {profile['unique_hpsa_names']:,}")
    print(f"Unique geography components:         {profile['unique_geo_components']:,}")
    print()
    print("Rows per hpsa_id:")
    for stat, val in profile["rows_per_site"].items():
        print(f"  {stat:>6}: {val:,.1f}" if isinstance(val, float) else f"  {stat:>6}: {val}")
    print()
    print("By designation type:")
    print(profile["designation_type_breakdown"].to_string())
    print()
    print("By geography component type:")
    print(profile["component_type_breakdown"].to_string())
    print()
    print("Date ranges:")
    for col, info in profile["date_ranges"].items():
        print(f"  {col}: {info['min']} → {info['max']}  (nulls: {info['nulls']})")
    print()
    print(f"Constant columns ({len(profile['constant_columns'])}): "
          f"{', '.join(profile['constant_columns'])}")
    print()
    high_null = profile["completeness"].loc[profile["completeness"]["null_pct"] > 50]
    print(f"Columns with >50% nulls ({len(high_null)}):")
    print(high_null[["column", "null_pct", "n_unique"]].to_string(index=False))
    print()
    print("Identifier notes:")
    for key, note in profile["identifier_notes"].items():
        print(f"  {key}: {note}")


def print_linking_guidance() -> None:
    """Print strategies for joining this data to job/site locations."""
    print()
    print("=" * 72)
    print("Linking jobs / healthcare sites to HPSA status")
    print("=" * 72)
    print(
        """
This file mixes two designation models:

1. FACILITY designations (4,844 sites, 1 row each, geo_id = "POINT")
   Types: FQHC, RHC, IHS/Tribal/Urban, Correctional, FQHC Look-Alike, Other.
   Link via: address + lat/lon, ZIP, or bhcmis_organization_identification_number
   (FQHC only). Match your job/site record to hpsa_sites.csv → is_hpsa = True.

2. AREA designations (Population, Geographic, High Needs Geographic)
   One hpsa_id spans many rows — census tracts (11-digit FIPS), whole counties
   (5-digit FIPS), or county subdivisions (10-digit FIPS).
   Link via: geocode the job/site to a census tract (or county), then join to
   hpsa_tract_lookup.csv or hpsa_geography.csv.

Recommended join workflow for arbitrary job locations:
  a) Geocode job address → lat/lon (Census geocoder).
  b) Point-in-polygon against HRSA boundary shapefiles (HPSA_CMP* for area HPSAs,
     HPSA_PNT* for facility HPSAs, MUA_SHP for MUA/P).
  c) Tract/county lookup tables (exported as CSV) are a fallback only when
     shapefiles or geopandas are unavailable — not used when SHP layers are loaded.

Caveats:
  - 149 census tracts appear in more than one HPSA designation; retain all matches.
  - hpsa_name is not unique (7120 names vs 7666 ids); always key on hpsa_id.
  - hpsa_id is not unique across domains; always include discipline/domain.
  - This analysis supports Primary Care, Dental Health, and Mental Health separately.

MUA/P (Medically Underserved Areas/Populations):
  - Detail file: MUA_DET.xlsx (component rows); shapefile: MUA_SHP (designation polygons).
  - Address checks use MUA_SHP point-in-polygon when the shapefile is loaded.
"""
    )


def print_pilot_site_results(
    results: pd.DataFrame,
    domain_assets: dict[str, HPSADomainAssets],
    *,
    mua_assets: MUAPAssets | None = None,
) -> None:
    """Print a concise summary of pilot-site HPSA and MUA/P checks."""
    print()
    print("=" * 72)
    print("Community health center pilot sites — HPSA & MUA/P checks")
    print("=" * 72)
    print(f"Sites checked: {len(results)}")
    print(f"In any designated HPSA: {results['in_hpsa_any'].sum()}")
    if "in_mua" in results.columns:
        print(f"In designated MUA/P: {results['in_mua'].sum()}")
        print(f"In HPSA or MUA/P: {results['in_hpsa_or_mua'].sum()}")
    print(f"Geocode failures: {(~results['geocode_success']).sum()}")
    print()

    for domain_key, assets in domain_assets.items():
        label = assets.paths.label
        in_col = f"in_hpsa_{domain_key}"
        area_col = f"in_hpsa_area_{domain_key}"
        facility_col = f"is_hpsa_facility_{domain_key}"
        method = "shapefile" if assets.area_boundaries is not None else "tract lookup"
        print(
            f"{label}: {results[in_col].sum()} in HPSA "
            f"({results[area_col].sum()} area, {results[facility_col].sum()} facility) "
            f"[{method}]"
        )

    if mua_assets is not None and "in_mua" in results.columns:
        method = "shapefile" if mua_assets.boundaries is not None else "tract lookup"
        print(f"MUA/P: {results['in_mua'].sum()} in MUA/P [{method}]")

    print()
    display_cols = ["address", "city", "zip", "geocode_success", "in_hpsa_any"]
    if "in_mua" in results.columns:
        display_cols.extend(["in_mua", "in_hpsa_or_mua"])
    for domain_key in domain_assets:
        display_cols.append(f"in_hpsa_{domain_key}")
    if "matched_mua_names" in results.columns:
        display_cols.append("matched_mua_names")
    for domain_key in domain_assets:
        display_cols.append(f"matched_hpsa_names_{domain_key}")
    print(results[display_cols].to_string(index=False))


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)

    if not HAS_GEOPANDAS:
        print("Warning: geopandas not installed; boundary-based checks will be skipped.")

    domain_assets: dict[str, HPSADomainAssets] = {}

    for domain_key in HPSA_DOMAINS:
        print()
        assets = build_domain_assets(domain_key, load_boundaries=HAS_GEOPANDAS)
        domain_assets[domain_key] = assets

        profile = profile_dataset(assets.designated)
        print_profile_summary(profile, domain_label=assets.paths.label)

        domain_out = write_domain_outputs(assets, profile)
        print()
        print(f"Output tables written to: {domain_out}")
        print(f"  hpsa_sites.csv          — {len(assets.sites):,} unique HPSA designations")
        print(f"  hpsa_geography.csv      — {len(assets.geography):,} geography components")
        print(f"  hpsa_tract_lookup.csv   — {len(assets.tract_lookup):,} census tracts")

        if assets.area_boundaries is not None:
            facility_count = (
                len(assets.facility_boundaries)
                if assets.facility_boundaries is not None
                else 0
            )
            print(
                f"  Boundary layers loaded: {len(assets.area_boundaries):,} area polygons, "
                f"{facility_count:,} facility points"
            )
        elif assets.paths.area_component_shp:
            print(f"  Boundary shapefile found but not loaded: {assets.paths.area_component_shp.name}")
        else:
            print("  No component boundary shapefile (HPSA_CMP*_DET_CUR_VX.shp) found")

        assert assets.sites[SITE_KEY].is_unique
        assert len(assets.sites) == assets.designated[SITE_KEY].nunique()

    print()
    print("=" * 72)
    print("MUA/P analysis")
    print("=" * 72)

    mua_assets: MUAPAssets | None = None
    if MUA_DET_XLSX.exists():
        mua_assets = build_mua_assets(load_boundaries=HAS_GEOPANDAS)
        mua_profile = profile_mua_dataset(mua_assets.designated)
        print_mua_profile_summary(mua_profile)
        mua_out = write_mua_outputs(mua_assets, mua_profile)
        print()
        print(f"Output tables written to: {mua_out}")
        print(f"  mua_sites.csv     — {len(mua_assets.sites):,} unique MUA/P designations")
        print(f"  mua_geography.csv — {len(mua_assets.geography):,} geography components")
        print(f"  mua_tract_lookup.csv — {len(mua_assets.tract_lookup):,} census tracts")
        if mua_assets.boundaries is not None:
            print(f"  MUA_SHP polygons loaded: {len(mua_assets.boundaries):,}")
        assert mua_assets.sites[MUA_SITE_KEY].is_unique
    else:
        print(f"MUA/P detail file not found: {MUA_DET_XLSX}")

    print_linking_guidance()

    if PILOT_SITES_PATH.exists():
        pilot_sites = load_pilot_sites(PILOT_SITES_PATH)
        pilot_results = check_addresses_all_domains(
            pilot_sites, domain_assets, mua_assets=mua_assets
        )
        pilot_results.to_csv(OUTPUT_DIR / "pilot_sites_hpsa_check.csv", index=False)
        print_pilot_site_results(
            pilot_results, domain_assets, mua_assets=mua_assets
        )
        print()
        print("Pilot site results written to:", OUTPUT_DIR / "pilot_sites_hpsa_check.csv")
    else:
        print()
        print(f"Pilot sites file not found: {PILOT_SITES_PATH}")
        print("Skipping address-level HPSA checks.")


if __name__ == "__main__":
    main()
