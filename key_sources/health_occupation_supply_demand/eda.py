"""
Exploratory analysis of HRSA Healthcare Occupations Supply & Demand projections.

Each row is a projected supply/demand snapshot for one occupation within a
profession group, year, state, and rurality level. Values are in FTE (full-time
equivalent). Scenario columns (fewer graduates, urban effect, etc.) are stored
in wide format alongside baseline Supply and Demand.

Source: Workforce_Projections_FullData.xlsx, sheet "Stacked FY2025".
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
DATA_PATH = Path(
    "/Users/karlhonerlaw/Library/CloudStorage/GoogleDrive-honerlaw@gmail.com"
    "/My Drive/CesarK_CPT_project/data/key_data_sources"
    "/Healthcare Occupations Supply Demand/Workforce_Projections_FullData.xlsx"
)
DATA_SHEET = "Stacked FY2025"
OUTPUT_DIR = DATA_PATH.parent / "output"

ROW_KEY = ["year", "profession_group", "profession", "state", "rurality"]

DIMENSION_COLUMNS = [
    "year",
    "profession_group",
    "profession_group_definition",
    "profession",
    "profession_definition",
    "state",
    "rurality",
    "region",
]

SUPPLY_COLUMNS = [
    "fte_supply_projections__fewer_graduates",
    "fte_supply_projections__more_graduates",
    "fte_supply_projections__retire_early",
    "fte_supply_projections__retire_late",
    "fte_supply_projections__supply",
]

DEMAND_COLUMNS = [
    "fte_demand_projections__demand",
    "fte_demand_projections__urban_effect",
    "fte_demand_projections__insurance_effect",
    "fte_demand_projections__race_effect",
    "fte_demand_projections__improved_access_combination_scenario",
    "fte_demand_projections__income_effect",
    "fte_demand_projections__unmet_need_1",
    "fte_demand_projections__unmet_need_2",
    "fte_demand_projections__elevated_need",
]

VALUE_COLUMNS = SUPPLY_COLUMNS + DEMAND_COLUMNS + ["percent_adequacy"]

DEMAND_COL = "fte_demand_projections__demand"
NATIONAL_GEO = {"state": "Total", "rurality": "Total"}

SCENARIO_LABELS = {
    "fte_supply_projections__fewer_graduates": "Supply — fewer graduates",
    "fte_supply_projections__more_graduates": "Supply — more graduates",
    "fte_supply_projections__retire_early": "Supply — retire early",
    "fte_supply_projections__retire_late": "Supply — retire late",
    "fte_supply_projections__supply": "Supply — baseline",
    "fte_demand_projections__demand": "Demand — baseline",
    "fte_demand_projections__urban_effect": "Demand — urbanization effect",
    "fte_demand_projections__insurance_effect": "Demand — insurance coverage effect",
    "fte_demand_projections__race_effect": "Demand — race/ethnicity effect",
    "fte_demand_projections__improved_access_combination_scenario": (
        "Demand — improved access (combination scenario)"
    ),
    "fte_demand_projections__income_effect": "Demand — income effect",
    "fte_demand_projections__unmet_need_1": "Demand — unmet need scenario 1",
    "fte_demand_projections__unmet_need_2": "Demand — unmet need scenario 2",
    "fte_demand_projections__elevated_need": "Demand — elevated need",
}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Lowercase column names and replace punctuation with underscores."""
    out = df.copy()
    out.columns = (
        out.columns.str.lower()
        .str.replace(" ", "_")
        .str.replace("-", "_")
        .str.replace("(", "", regex=False)
        .str.replace(")", "", regex=False)
        .str.replace("__", "_", regex=False)
    )
    return out


def load_workforce_projections(path: Path = DATA_PATH) -> pd.DataFrame:
    """Load and clean the stacked FY2025 workforce projections sheet."""
    if not path.exists():
        raise FileNotFoundError(f"Workforce projections file not found: {path}")

    df = normalize_columns(pd.read_excel(path, sheet_name=DATA_SHEET, header=0))
    df = df.loc[df["year"] != "Year"].copy()
    df["year"] = pd.to_numeric(df["year"], errors="coerce")
    df = df.loc[df["year"].notna()].copy()
    df["year"] = df["year"].astype(int)

    for col in VALUE_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df.reset_index(drop=True)


def build_dimension_summary(df: pd.DataFrame) -> dict[str, Any]:
    """Summarize categorical dimensions and their cardinality."""
    profession_by_group = (
        df.groupby("profession_group")["profession"]
        .apply(lambda values: sorted(values.unique()))
        .to_dict()
    )

    state_rurality_patterns = (
        df.groupby("state")["rurality"]
        .apply(lambda values: tuple(sorted(values.unique())))
        .value_counts()
        .rename("n_states")
        .reset_index()
        .rename(columns={"rurality": "rurality_pattern"})
    )

    rows_per_year = df.groupby("year").size().sort_index()

    return {
        "years": sorted(df["year"].unique().tolist()),
        "year_range": {"min": int(df["year"].min()), "max": int(df["year"].max())},
        "profession_groups": sorted(df["profession_group"].unique().tolist()),
        "n_profession_groups": df["profession_group"].nunique(),
        "n_professions": df["profession"].nunique(),
        "profession_by_group": profession_by_group,
        "n_states": df["state"].nunique(),
        "states": sorted(df["state"].unique().tolist()),
        "rurality_levels": sorted(df["rurality"].unique().tolist()),
        "regions": sorted(df["region"].unique().tolist()),
        "rows_per_year": rows_per_year,
        "state_rurality_patterns": state_rurality_patterns,
    }


def build_scenario_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Report non-null counts and geographic scope for each scenario column."""
    rows: list[dict[str, Any]] = []
    for col in VALUE_COLUMNS:
        if col not in df.columns:
            continue
        populated = df.loc[df[col].notna()]
        rows.append(
            {
                "column": col,
                "label": SCENARIO_LABELS.get(col, col),
                "non_null": int(populated.shape[0]),
                "null_pct": round(100 * df[col].isna().mean(), 1),
                "national_only": bool(
                    populated.empty or populated["state"].eq("Total").all()
                ),
                "states_with_values": int(populated["state"].nunique()),
                "rurality_levels": ", ".join(sorted(populated["rurality"].unique()))
                if not populated.empty
                else "",
            }
        )
    return pd.DataFrame(rows).sort_values(["null_pct", "column"])


def _count_demand_decline_years(group: pd.DataFrame, demand_col: str) -> pd.Series:
    """Count year-over-year demand decreases within one occupation series."""
    ordered = group.sort_values("year")
    declines = ordered[demand_col].diff() < 0
    return pd.Series(
        {
            "n_decline_years": int(declines.sum()),
            "any_yoy_decline": bool(declines.any()),
        }
    )


def build_occupation_demand_trends(
    df: pd.DataFrame,
    group_cols: list[str],
    *,
    start_year: int,
    end_year: int,
    demand_col: str = DEMAND_COL,
) -> pd.DataFrame:
    """Summarize baseline demand change and year-over-year declines for one geography slice."""
    start = (
        df.loc[df["year"] == start_year, [*group_cols, demand_col]]
        .drop_duplicates(subset=group_cols)
        .rename(columns={demand_col: "demand_start"})
    )
    end = (
        df.loc[df["year"] == end_year, [*group_cols, demand_col]]
        .drop_duplicates(subset=group_cols)
        .rename(columns={demand_col: "demand_end"})
    )

    trends = start.merge(end, on=group_cols, how="inner")
    trends["start_year"] = start_year
    trends["end_year"] = end_year
    trends["change_abs"] = trends["demand_end"] - trends["demand_start"]
    trends["change_pct"] = (
        100 * (trends["demand_end"] / trends["demand_start"] - 1)
    ).round(2)

    yoy = (
        df.groupby(group_cols, as_index=False)
        .apply(_count_demand_decline_years, demand_col=demand_col, include_groups=False)
    )
    trends = trends.merge(yoy, on=group_cols, how="left")
    return trends.sort_values(group_cols).reset_index(drop=True)


def build_demand_trend_summary(df: pd.DataFrame) -> dict[str, Any]:
    """Analyze baseline demand trends nationally and by state."""
    start_year = int(df["year"].min())
    end_year = int(df["year"].max())

    national = df.loc[
        (df["state"] == NATIONAL_GEO["state"])
        & (df["rurality"] == NATIONAL_GEO["rurality"])
    ]
    national_trends = build_occupation_demand_trends(
        national,
        ["profession_group", "profession"],
        start_year=start_year,
        end_year=end_year,
    )

    state_rows = df.loc[(df["state"] != "Total") & (df["rurality"] == "Total")]
    state_trends = build_occupation_demand_trends(
        state_rows,
        ["state", "profession_group", "profession"],
        start_year=start_year,
        end_year=end_year,
    )
    state_net_declining = state_trends.loc[state_trends["change_abs"] < 0].copy()
    state_decline_by_profession = (
        state_net_declining.groupby(["profession_group", "profession"], as_index=False)
        .size()
        .rename(columns={"size": "n_states_declining"})
        .sort_values("n_states_declining", ascending=False)
    )

    return {
        "start_year": start_year,
        "end_year": end_year,
        "national_trends": national_trends,
        "national_net_declining": national_trends.loc[
            national_trends["change_abs"] < 0
        ].sort_values("change_pct"),
        "national_yoy_declining": national_trends.loc[
            national_trends["any_yoy_decline"]
        ].sort_values("n_decline_years", ascending=False),
        "state_trends": state_trends,
        "state_net_declining": state_net_declining.sort_values("change_pct"),
        "state_decline_by_profession": state_decline_by_profession,
    }


def build_supply_gap_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Identify occupations that lack state-level supply but still have demand."""
    supply_col = "fte_supply_projections__supply"
    demand_col = "fte_demand_projections__demand"

    state_rows = df.loc[df["state"] != "Total"].copy()
    summary = (
        state_rows.groupby(["profession_group", "profession"], as_index=False)
        .agg(
            rows=("year", "size"),
            supply_non_null=(supply_col, lambda s: int(s.notna().sum())),
            demand_non_null=(demand_col, lambda s: int(s.notna().sum())),
        )
        .assign(
            supply_missing=lambda frame: frame["rows"] - frame["supply_non_null"],
            demand_missing=lambda frame: frame["rows"] - frame["demand_non_null"],
        )
        .sort_values(["supply_missing", "profession_group", "profession"], ascending=False)
    )
    return summary


def profile_dataset(df: pd.DataFrame) -> dict[str, Any]:
    """Return a structured profile of row counts, keys, dates, and nulls."""
    dimensions = build_dimension_summary(df)

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

    duplicate_keys = int(df.duplicated(subset=ROW_KEY).sum())
    adequacy_check = df.loc[
        df["percent_adequacy"].notna()
        & df["fte_supply_projections__supply"].notna()
        & df["fte_demand_projections__demand"].notna()
    ].copy()
    if not adequacy_check.empty:
        adequacy_check["calc_adequacy"] = (
            adequacy_check["fte_supply_projections__supply"]
            / adequacy_check["fte_demand_projections__demand"]
        )
        adequacy_max_diff = float(
            (adequacy_check["calc_adequacy"] - adequacy_check["percent_adequacy"])
            .abs()
            .max()
        )
    else:
        adequacy_max_diff = None

    return {
        "raw_rows": len(df),
        "unique_keys": df[ROW_KEY].drop_duplicates().shape[0],
        "duplicate_key_rows": duplicate_keys,
        "dimensions": dimensions,
        "scenario_summary": build_scenario_summary(df),
        "supply_gap_summary": build_supply_gap_summary(df),
        "demand_trends": build_demand_trend_summary(df),
        "completeness": completeness,
        "constant_columns": constant_cols,
        "adequacy_is_supply_over_demand": adequacy_max_diff == 0.0
        if adequacy_max_diff is not None
        else None,
        "adequacy_max_abs_diff": adequacy_max_diff,
        "identifier_notes": {
            "row_key": (
                "Natural key: year + profession_group + profession + state + rurality. "
                "Unique across all 102,528 rows."
            ),
            "state_total": (
                "State = 'Total' is the US national aggregate, not a missing value. "
                "Only this row set includes Metro / NonMetro rurality splits."
            ),
            "profession_group_vs_profession": (
                "Profession Group is a HRSA analytical bucket (7 values). Profession is "
                "the specific occupation within that bucket (121 distinct labels)."
            ),
            "percent_adequacy": (
                "Equals baseline Supply / Demand when both are present. Null wherever "
                "Supply is null."
            ),
            "scenario_columns": (
                "Sensitivity/scenario FTE projections are stored as separate columns "
                "(wide format), not as a Scenario dimension. Most scenarios are "
                "populated only for State = 'Total'."
            ),
        },
    }


def write_outputs(df: pd.DataFrame, profile: dict[str, Any]) -> Path:
    """Write profile tables and a cleaned long-form-friendly CSV."""
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    df.to_csv(OUTPUT_DIR / "workforce_projections_clean.csv", index=False)
    profile["completeness"].to_csv(OUTPUT_DIR / "column_completeness.csv", index=False)
    profile["scenario_summary"].to_csv(OUTPUT_DIR / "scenario_summary.csv", index=False)
    profile["supply_gap_summary"].to_csv(
        OUTPUT_DIR / "supply_gap_by_profession.csv", index=False
    )

    profession_catalog = []
    for group, professions in profile["dimensions"]["profession_by_group"].items():
        for profession in professions:
            profession_catalog.append(
                {"profession_group": group, "profession": profession}
            )
    pd.DataFrame(profession_catalog).to_csv(
        OUTPUT_DIR / "profession_catalog.csv", index=False
    )

    demand_trends = profile["demand_trends"]
    demand_trends["national_trends"].to_csv(
        OUTPUT_DIR / "demand_trends_national.csv", index=False
    )
    demand_trends["national_net_declining"].to_csv(
        OUTPUT_DIR / "demand_trends_national_declining.csv", index=False
    )
    demand_trends["national_yoy_declining"].to_csv(
        OUTPUT_DIR / "demand_trends_national_yoy_declines.csv", index=False
    )
    demand_trends["state_net_declining"].to_csv(
        OUTPUT_DIR / "demand_trends_state_declining.csv", index=False
    )
    demand_trends["state_decline_by_profession"].to_csv(
        OUTPUT_DIR / "demand_trends_state_decline_by_profession.csv", index=False
    )

    return OUTPUT_DIR


def print_structure_summary(profile: dict[str, Any]) -> None:
    """Print a concise explanation of how the dataset is organized."""
    dims = profile["dimensions"]
    print("=" * 72)
    print("Healthcare Occupations Supply & Demand — structure")
    print("=" * 72)
    print(
        """
Your understanding is largely correct: the dataset projects healthcare workforce
supply and demand over time, broken out by profession group, specific occupation,
state, and (where available) rurality.

Key structural details beyond that summary:

1. GRAIN
   One row = one occupation × year × state × rurality level.
   There are 7 Profession Groups and 121 distinct Profession labels. Groups such as
   "Primary Care" and "All Health Workforce" contain overlapping occupation types
   with different labels (e.g. "Nurse Practitioners (PC)" vs "Nurse Practitioners").

2. GEOGRAPHY
   - 51 states + DC + a US aggregate row where State = "Total".
   - State-level rows use Rurality = "Total" only.
   - Metro / NonMetro splits exist only for the US aggregate (State = "Total").
   - Region is an HRSA census region derived from state (Northeast, Midwest, South,
     West, plus "US Total" for the national aggregate rows).

3. TIME
   Projections span 2023–2038 (16 years). Each year contains 6,408 rows.

4. SCENARIOS (wide format)
   HRSA includes multiple supply and demand scenarios as separate numeric columns,
   not as a "Scenario" field. Baseline values are Supply and Demand. Additional
   supply scenarios include fewer/more graduates and early/late retirement. Demand
   scenarios include urbanization, insurance, race, income, access, unmet need, and
   elevated need variants. Most scenario columns are populated only at the national
   level (State = "Total").

5. COMPLETENESS
   Demand is present for every row. Baseline Supply is missing for 31.5% of rows —
   primarily 39 occupations at the state level (many Allied Health and all Long-Term
   Care occupations). Percent Adequacy is null wherever Supply is null and otherwise
   equals Supply / Demand.
"""
    )
    print("Profession groups and occupation counts:")
    for group in dims["profession_groups"]:
        count = len(dims["profession_by_group"][group])
        print(f"  {group}: {count}")
    print()
    print("Rurality patterns by state:")
    print(dims["state_rurality_patterns"].to_string(index=False))


def print_demand_trends_summary(profile: dict[str, Any]) -> None:
    """Print baseline demand trend findings for national and state geographies."""
    trends = profile["demand_trends"]
    start_year = trends["start_year"]
    end_year = trends["end_year"]
    national = trends["national_trends"]
    net_declining = trends["national_net_declining"]
    yoy_declining = trends["national_yoy_declining"]
    state_declining = trends["state_net_declining"]
    state_by_prof = trends["state_decline_by_profession"]

    print()
    print("=" * 72)
    print("Healthcare Occupations Supply & Demand — demand trends")
    print("=" * 72)
    print(
        f"Baseline demand column: {DEMAND_COL}\n"
        f"National reference geography: State = 'Total', Rurality = 'Total'\n"
        f"Comparison window: {start_year} → {end_year}"
    )
    print()
    print(
        f"National occupations with lower {end_year} vs {start_year} demand: "
        f"{len(net_declining)} of {len(national)}"
    )
    if net_declining.empty:
        print("  (none)")
    else:
        display_cols = [
            "profession_group",
            "profession",
            "demand_start",
            "demand_end",
            "change_abs",
            "change_pct",
            "n_decline_years",
        ]
        print(net_declining[display_cols].to_string(index=False))

    print()
    print(
        f"National occupations with any year-over-year demand decline: "
        f"{len(yoy_declining)} of {len(national)}"
    )
    if not yoy_declining.empty:
        print(
            yoy_declining[
                ["profession_group", "profession", "change_pct", "n_decline_years"]
            ]
            .head(10)
            .to_string(index=False)
        )
        if len(yoy_declining) > 10:
            print(
                "  ... see demand_trends_national_yoy_declines.csv for full list"
            )

    print()
    print(
        f"State-level profession combinations with lower {end_year} vs {start_year} "
        f"demand: {len(state_declining)} of {len(trends['state_trends'])}"
    )
    print(
        "Note: state FTE values are often rounded to the nearest 10, so large "
        "percentage drops can reflect rounding as well as modeled decline."
    )
    if not state_by_prof.empty:
        print()
        print("States with net demand decline, by occupation (top 10):")
        print(state_by_prof.head(10).to_string(index=False))
        if len(state_by_prof) > 10:
            print(
                "  ... see demand_trends_state_decline_by_profession.csv for full list"
            )


def print_profile_summary(profile: dict[str, Any]) -> None:
    """Print a human-readable profiling summary to stdout."""
    dims = profile["dimensions"]
    print()
    print("=" * 72)
    print("Healthcare Occupations Supply & Demand — dataset profile")
    print("=" * 72)
    print(f"Rows:                         {profile['raw_rows']:,}")
    print(f"Unique row keys:              {profile['unique_keys']:,}")
    print(f"Duplicate key rows:           {profile['duplicate_key_rows']:,}")
    print()
    print(
        f"Years:                        {dims['year_range']['min']}–{dims['year_range']['max']} "
        f"({len(dims['years'])} years)"
    )
    print(f"Profession groups:            {dims['n_profession_groups']}")
    print(f"Distinct profession labels:   {dims['n_professions']}")
    print(f"States / areas:               {dims['n_states']} (includes State = 'Total')")
    print(f"Rurality levels:              {', '.join(dims['rurality_levels'])}")
    print(f"Regions:                      {', '.join(dims['regions'])}")
    print()
    print("Rows per year:")
    for year, count in dims["rows_per_year"].items():
        print(f"  {year}: {count:,}")
    print()
    print("Scenario / value column coverage:")
    print(
        profile["scenario_summary"][
            ["label", "non_null", "null_pct", "national_only", "rurality_levels"]
        ].to_string(index=False)
    )
    print()
    if profile["adequacy_is_supply_over_demand"]:
        print("Percent Adequacy matches Supply / Demand exactly where both are present.")
    print()
    missing_supply = profile["supply_gap_summary"].loc[
        profile["supply_gap_summary"]["supply_missing"] > 0
    ]
    print(
        f"Occupations with missing state-level Supply: {len(missing_supply)} "
        f"(of {dims['n_professions']} profession labels)"
    )
    print(missing_supply.head(10).to_string(index=False))
    if len(missing_supply) > 10:
        print(f"  ... and {len(missing_supply) - 10} more (see supply_gap_by_profession.csv)")
    print()
    print(f"Constant columns ({len(profile['constant_columns'])}): "
          f"{', '.join(profile['constant_columns']) or '(none)'}")
    print()
    high_null = profile["completeness"].loc[profile["completeness"]["null_pct"] > 50]
    print(f"Columns with >50% nulls ({len(high_null)}):")
    print(high_null[["column", "null_pct", "n_unique"]].to_string(index=False))
    print()
    print("Identifier notes:")
    for key, note in profile["identifier_notes"].items():
        print(f"  {key}: {note}")


def main() -> None:
    df = load_workforce_projections()
    profile = profile_dataset(df)

    print_structure_summary(profile)
    print_profile_summary(profile)
    print_demand_trends_summary(profile)

    out_dir = write_outputs(df, profile)
    print()
    print(f"Output tables written to: {out_dir}")
    print(f"  workforce_projections_clean.csv — {len(df):,} cleaned projection rows")
    print("  column_completeness.csv       — null rates and cardinalities")
    print("  scenario_summary.csv          — scenario column coverage")
    print("  supply_gap_by_profession.csv  — state-level supply gaps by occupation")
    print("  profession_catalog.csv        — profession group → profession mapping")
    print("  demand_trends_national.csv    — national demand change by occupation")
    print("  demand_trends_national_declining.csv — occupations with net national decline")
    print("  demand_trends_national_yoy_declines.csv — occupations with any YoY decline")
    print("  demand_trends_state_declining.csv — state-level net demand declines")
    print("  demand_trends_state_decline_by_profession.csv — decline counts by occupation")

    assert profile["duplicate_key_rows"] == 0
    assert profile["unique_keys"] == len(df)


if __name__ == "__main__":
    main()
