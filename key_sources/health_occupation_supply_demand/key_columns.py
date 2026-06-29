import pandas as pd
from pathlib import Path
import plotly.express as px
from plotly.offline import plot

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
DATA_PATH = Path(
    "/Users/karlhonerlaw/Library/CloudStorage/GoogleDrive-honerlaw@gmail.com"
    "/My Drive/CesarK_CPT_project/data/key_data_sources"
    "/Healthcare Occupations Supply Demand/Workforce_Projections_FullData.xlsx"
)

# load supply and demand data 
df = pd.read_excel(DATA_PATH)
len(df)


########################################################
# Reproducing the HRSA Supply/Demand plots 
########################################################

# Key learning: 
# HRSA supply/demand data is stored in a stacked format where data is shown both at an aggregate level (e.g. "Total") as well as at more granular levelss (e.g. by State, Rurality, and Profession Group)

# To reproduce the HRSA Supply/Demand plots, we need to:
# filter to the correct Profession Group (e.g. "All Health Workforce")
# filter to the correct Profession (e.g. "Nurse Practitioner")
# filter to the correct "State" and "Rurality" levels (note: "Total" is the US national aggregate)

# Examples: 
# Supply/Demand plot for "All Health Workforce" and "Registered Nurses" at the US national aggregate level

# filter data 
len(df)
df_filtered = df[df["Profession Group"] == "All Health Workforce"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["Profession"] == "Registered Nurses"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["State"] == "Total"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["Rurality"] == "Total"].copy()
len(df_filtered)

# subset to the correct columns and reshape to long format for plotting
df_filtered.columns
df_filtered_sd = df_filtered[["Year", "FTE Supply Projections - Supply", "FTE Demand Projections - Demand"]].copy()

df_filtered_sd = df_filtered_sd.rename(columns={"Year": "year", "FTE Supply Projections - Supply": "supply", "FTE Demand Projections - Demand": "demand"})
len(df_filtered_sd)

# reshape to long format for plotting
df_filtered_sd = df_filtered_sd.melt(id_vars=["year"], var_name="type", value_name="value")
len(df_filtered_sd)
df_filtered_sd.head()

fig = px.line(df_filtered_sd, x="year", y="value", color="type")
plot(fig)

# percent adequacy
fig = px.line(df_filtered, x="Year", y="Percent Adequacy")
plot(fig)


########################################################
# Supply and Demand Across All Professions
########################################################

# supply and demand across all professions
df_filtered = df[df["Profession Group"] == "All Health Workforce"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["State"] == "Total"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["Rurality"] == "Total"].copy()
len(df_filtered)

# reshape to long format for plotting
df_filtered_sd = df_filtered[["Year", "FTE Supply Projections - Supply", "FTE Demand Projections - Demand", "Profession"]].copy()
len(df_filtered_sd)

df_filtered_sd = df_filtered_sd.rename(columns={"Year": "year", "FTE Supply Projections - Supply": "supply", "FTE Demand Projections - Demand": "demand"})
len(df_filtered_sd)

df_filtered_sd = df_filtered_sd.melt(id_vars=["year", "Profession"], var_name="type", value_name="value")
len(df_filtered_sd)

fig = px.line(df_filtered_sd, x="year", y="value", color="Profession", line_dash = "type")
plot(fig, filename="supply_demand_all_professions.html")

########################################################
# Percent Adequacy Across All Professions
########################################################

# percent adequacy across all professions
df_filtered = df[df["Profession Group"] == "All Health Workforce"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["State"] == "Total"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["Rurality"] == "Total"].copy()
len(df_filtered)

fig = px.line(df_filtered, x="Year", y="Percent Adequacy", color="Profession")
plot(fig, filename="percent_adequacy_all_professions.html")