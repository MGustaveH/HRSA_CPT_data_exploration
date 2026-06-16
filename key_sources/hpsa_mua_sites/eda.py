"""
Exploratory analysis of HRSA Primary Care HPSA detail data
(BCD_HPSA_FCT_DET_PC.xlsx).

Each row is a geographic *component* of an HPSA designation. Area-based
designations (Population, Geographic, High Needs Geographic) span many rows —
one per census tract, county, or county subdivision. Facility-based designations
(FQHC, RHC, IHS, etc.) have exactly one row each with geo id "POINT".
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
DATA_PATH = Path(
    "/Users/karlhonerlaw/Library/CloudStorage/GoogleDrive-honerlaw@gmail.com"
    "/My Drive/CesarK_CPT_project/data/key_data_sources/data_samples"
    "/HPSA MUAP Sites/HPSA – Primary Care/BCD_HPSA_FCT_DET_PC.xlsx"
)
OUTPUT_DIR = Path(__file__).parent / "output"

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


def load_designated_hpsa(path: Path = DATA_PATH) -> pd.DataFrame:
    """Load the detail file and keep only Designated primary-care HPSAs."""
    df = pd.read_excel(path)
    df.columns = df.columns.str.lower().str.replace(" ", "_")
    designated = df[df["hpsa_status"] == "Designated"].copy()
    return designated


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


def build_tract_lookup(geography: pd.DataFrame) -> pd.DataFrame:
    """Tract FIPS → HPSA designations for joining external address/tract data."""
    tract_rows = geography.dropna(subset=["census_tract_fips"]).copy()
    lookup = (
        tract_rows.groupby("census_tract_fips", as_index=False)
        .agg(
            n_hpsa_sites=(SITE_KEY, "nunique"),
            hpsa_ids=(SITE_KEY, lambda s: sorted(s.unique().tolist())),
            hpsa_names=("hpsa_name", lambda s: sorted(s.unique().tolist())),
            designation_types=("designation_type", lambda s: sorted(s.unique().tolist())),
        )
        .sort_values("census_tract_fips")
    )
    return lookup


def print_profile_summary(profile: dict) -> None:
    """Print a human-readable summary to stdout."""
    print("=" * 72)
    print("HPSA Primary Care — Designated records profile")
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
  a) Geocode job address → census tract FIPS (11 digits: SSCCCTTTTTT).
  b) Left-join to hpsa_tract_lookup on census_tract_fips.
  c) If n_hpsa_sites >= 1, the location is in at least one designated HPSA.
  d) Optionally join hpsa_sites for scores, designation type, dates.

Caveats:
  - 149 census tracts appear in more than one HPSA designation; retain all matches.
  - Area-based rows have no lat/lon; only facility rows do.
  - hpsa_name is not unique (7120 names vs 7666 ids); always key on hpsa_id.
  - This file is Primary Care only; other discipline HPSAs are separate datasets.
"""
    )


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)

    df = load_designated_hpsa()
    profile = profile_dataset(df)
    print_profile_summary(profile)

    sites = build_site_table(df)
    geography = build_geography_table(df)
    tract_lookup = build_tract_lookup(geography)

    sites.to_csv(OUTPUT_DIR / "hpsa_sites.csv", index=False)
    geography.to_csv(OUTPUT_DIR / "hpsa_geography.csv", index=False)
    tract_lookup.to_csv(OUTPUT_DIR / "hpsa_tract_lookup.csv", index=False)
    profile["completeness"].to_csv(OUTPUT_DIR / "column_completeness.csv", index=False)
    profile["designation_type_breakdown"].to_csv(
        OUTPUT_DIR / "designation_type_breakdown.csv"
    )

    print()
    print("Output tables written to:", OUTPUT_DIR)
    print(f"  hpsa_sites.csv          — {len(sites):,} unique HPSA designations")
    print(f"  hpsa_geography.csv      — {len(geography):,} geography components")
    print(f"  hpsa_tract_lookup.csv   — {len(tract_lookup):,} census tracts with HPSA coverage")
    print(f"  column_completeness.csv — null/unique stats for all columns")

    # Sanity checks
    assert sites[SITE_KEY].is_unique, "Site table must have one row per hpsa_id"
    assert len(sites) == df[SITE_KEY].nunique()
    assert geography.drop_duplicates(subset=[SITE_KEY, "geo_id"]).shape[0] == len(geography)

    print_linking_guidance()


if __name__ == "__main__":
    main()
