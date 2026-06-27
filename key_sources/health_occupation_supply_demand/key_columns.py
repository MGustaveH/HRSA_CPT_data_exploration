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
df_filtered = df_filtered[["Year", "FTE Supply Projections - Supply", "FTE Demand Projections - Demand"]].copy()

df_filtered = df_filtered.rename(columns={"Year": "year", "FTE Supply Projections - Supply": "supply", "FTE Demand Projections - Demand": "demand"})
len(df_filtered)

# reshape to long format for plotting
df_filtered = df_filtered.melt(id_vars=["year"], var_name="type", value_name="value")
len(df_filtered)
df_filtered.head()

fig = px.line(df_filtered, x="year", y="value", color="type")
plot(fig)










# print a list of columns that have the word "nurse" anywhere in their row values
nurse_columns = []
for col in df.columns:
    print("checking column: ", col)
    col_unique_vals = list(df[col].unique())
    for val in col_unique_vals:
        val = str(val).lower()
        if "nurse" in val:
            print("found nurse in column: ", col)
            nurse_columns.append(col)
            break
    else:
        print("no nurse found in column: ", col)

print(nurse_columns)


# check unique values of "Profession" column
profession_unique_vals = list(df["Profession"].unique())
nurse_professions = []
for val in profession_unique_vals:
    print(val)
    if "nurse" in val.lower():
        nurse_professions.append(val)
print(nurse_professions)


########################################################
# Reproducing the HRSA Supply/Demand plots 
########################################################

# Key learning: 
# HRSA supply/demand data is stored in a stacked format where data is shown both at an aggregate level as well as at more granular levelss (e.g. by State, Rurality, and Profession Group)

# To reproduce the HRSA Supply/Demand plots, we need to:
# filter to the correct Profession Group 
# filter to the correct "State" and "Rurality" levels (note: "Total" is the US national aggregate)

# Examples: 

966430 / 1025080