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



# percent adequacy across all professions
df_filtered = df[df["Profession Group"] == "All Health Workforce"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["State"] == "Total"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["Rurality"] == "Total"].copy()
len(df_filtered)

fig = px.line(df_filtered, x="Year", y="Percent Adequacy", color="Profession")
plot(fig)


########################################################
# Explore Profession Group and Professions
########################################################

len(df['Profession Group Definition'].unique())
len(df['Profession Group'].unique())

# print unique Profession Group Definition values 
for pg in df['Profession Group'].unique():
    print(pg)
    print(df[df['Profession Group'] == pg]['Profession Group Definition'].values[0])
    print("-" * 100)


# Do any Professions appear in multiple Profession Groups?
df_dedup = df[['Profession', 'Profession Group']].drop_duplicates().sort_values('Profession')
len(df_dedup)

# count the number of times a Profession appears in the df
df_dedup['Profession'].value_counts(dropna=False)
df_dedup['Profession'].value_counts(dropna=False).value_counts(dropna=False)

# Takeaway: every profession only appears once 


# Relationship between Profession and Profession Definition 
len(df['Profession'].unique())
len(df['Profession Definition'].unique())

# Takeaway: Some Professions must share definitions 

# count the number of Profession Definitions for each Profession
df_dedup_definition = df[['Profession', 'Profession Definition']].drop_duplicates()
len(df_dedup_definition)

df_dedup_definition['Profession Definition'].value_counts(dropna=False)
df_dedup_definition['Profession Definition'].value_counts(dropna=False).value_counts(dropna=False)

# subset to definitions which appear multiple times 
df_dedup_definition_def_cts = df_dedup_definition.groupby('Profession Definition').count().reset_index().rename(columns = {'Profession':'def_ct'})
df_dedup_definition_def_cts.head()
df_dedup_definition_def_cts = df_dedup_definition_def_cts.sort_values(by = 'def_ct', ascending = False)

# filter to the 'Profession Definition' that have a def_ct > 1 
df_sing_def_mult_prof = df_dedup_definition_def_cts[df_dedup_definition_def_cts['def_ct'] > 1].copy()
len(df_sing_def_mult_prof)

# subset to professions that have definitions in df_sing_def_mult_prof
df_prof_sharing_def = df_dedup_definition[df_dedup_definition['Profession Definition'].isin(df_sing_def_mult_prof['Profession Definition'])]
len(df_prof_sharing_def)
df_prof_sharing_def.head()

# sort by 'Profession Definition'
df_prof_sharing_def = df_prof_sharing_def.sort_values(by = 'Profession Definition')

# print these Professions 
for prof in df_prof_sharing_def['Profession']: 
    print(prof)


#############################################
# Determine How to Handle Professions which appear in multiple Profession Groups
# e.g.
# Nurse Practitioners
# Nurse Practitioners (PC)
# Nurse Practitioners (WH)
#############################################

# Let's start by exploring their supply/demand 

# filter data 
len(df)
df_filtered = df[df["Profession"].isin(["Nurse Practitioners", "Nurse Practitioners (PC)", "Nurse Practitioners (WH)"])].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["State"] == "Total"].copy()
len(df_filtered)

df_filtered = df_filtered[df_filtered["Rurality"] == "Total"].copy()
len(df_filtered)

# subset to the correct columns and reshape to long format for plotting
df_filtered.columns
df_filtered_sd = df_filtered[["Year", 'Profession', "FTE Supply Projections - Supply", "FTE Demand Projections - Demand"]].copy()

df_filtered_sd = df_filtered_sd.rename(columns={"Year": "year", "FTE Supply Projections - Supply": "supply", "FTE Demand Projections - Demand": "demand"})
len(df_filtered_sd)

# reshape to long format for plotting
df_filtered_sd = df_filtered_sd.melt(id_vars=["year", 'Profession'], var_name="type", value_name="value")
len(df_filtered_sd)
df_filtered_sd.head()

fig = px.line(df_filtered_sd, x="year", y="value", color="type", line_dash = 'Profession')
plot(fig, filename='NP_dis_agg.html')

# Takeaway: these look to be distinct, perhaps we combine their values into a "Total" group


# High Level Thoughts on How to Account for some Professions Sharing a Profession Definition
# And therefore being so similar as to likely map to the same roles in the CPT 

# I don't think I want to create records in the table... 
# Instead I want to create a function which can flexibly handle a user's needs 

# filter to Total Population 
df_filtered = df[df["State"] == "Total"].copy()
len(df_filtered)
df_filtered = df_filtered[df_filtered["Rurality"] == "Total"].copy()
len(df_filtered)

# Create a Profession column which does not include designations in parenthases (i.e. "(LTC)", "(PC)", "(WH)")
# problem... what if this changes or isn't the right pattern? 
# working from the definitions is probably safer 

# subset to the correct columns and reshape to long format for plotting
df_filtered.columns
df_filtered_sd = df_filtered[["Year", 'Profession Definition', "FTE Supply Projections - Supply", "FTE Demand Projections - Demand"]].copy()

# rename columns 
df_filtered_sd = df_filtered_sd.rename(columns={"Year": "year", "Profession Definition":'pd', "FTE Supply Projections - Supply": "supply", "FTE Demand Projections - Demand": "demand"})
len(df_filtered_sd)

# groupby year and 'Profession Definition' and calculate the sum of supply and demand
df_filtered_sd_grp = df_filtered_sd.groupby(['year', 'pd']).agg('sum').reset_index()
df_filtered_sd_grp.head()
len(df_filtered_sd_grp)
df_filtered_sd_grp.columns


# For each Profession Definition, make a list of associated professions 
df_pd_profs = pd.DataFrame()
for p in df_filtered['Profession Definition'].unique(): 
    df_temp = df_filtered[df_filtered['Profession Definition'] == p].copy()
    df_temp_profs = list(df_temp['Profession'].unique())

    # make a dataframe and concat onto df_pd_profs
    df_pd_prof = pd.DataFrame()
    df_pd_prof['pd'] = [p]
    df_pd_prof['Profession'] = [df_temp_profs]
    df_pd_prof['prof_ct'] = [len(df_temp_profs)]


    df_pd_profs = pd.concat([df_pd_profs, df_pd_prof])

len(df_pd_profs)
df_pd_profs['prof_ct'].value_counts(dropna=False)

df_pd_profs = df_pd_profs.sort_values(by = 'prof_ct', ascending = False)
df_pd_profs['Profession'].head()

# save 'Profession' lists as string 
df_pd_profs['Profession'] = df_pd_profs['Profession'].astype('str')
df_pd_profs['Profession'].head()

# join df_pd_profs onto df_filtered_sd_grp on df 
len(df_filtered_sd_grp)
df_filtered_sd_grp = df_filtered_sd_grp.merge(df_pd_profs, on = 'pd', how = 'left')
len(df_filtered_sd_grp)
df_filtered_sd_grp.columns

# subset to key columns 
df_filtered_sd_grp_sub = df_filtered_sd_grp[['year', 'Profession', 'supply', 'demand']]

# reshape to long format for plotting
df_filtered_sd_grp_mlt = df_filtered_sd_grp_sub.melt(id_vars=["year", 'Profession'], var_name="type", value_name="value")
len(df_filtered_sd_grp_mlt)
df_filtered_sd_grp_mlt.head()

fig = px.line(df_filtered_sd_grp_mlt, x="year", y="value", color="Profession", line_dash = 'type')
plot(fig, filename='agg.html')

# Takeaway: 
# Combining by definition works well
# When checking the sum of supply/demand values for Nurse Practicioner, these tie out 