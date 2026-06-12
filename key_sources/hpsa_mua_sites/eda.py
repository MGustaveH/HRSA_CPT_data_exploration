import pandas as pd

data = pd.read_excel("/Users/karlhonerlaw/Library/CloudStorage/GoogleDrive-honerlaw@gmail.com/My Drive/CesarK_CPT_project/data/key_data_sources/data_samples/HPSA MUAP Sites/HPSA – Primary Care/BCD_HPSA_FCT_DET_PC.xlsx")
len(data)

data.columns
len(data.columns)

# convert column names to lowercase and replace spaces with underscores
data.columns = data.columns.str.lower().str.replace(" ", "_")
data.columns

# check hpsa_status
data["hpsa_status"].value_counts(dropna=False)

# filter for hpsa_status == "Designated"
data = data[data["hpsa_status"] == "Designated"].copy()
len(data)

# look for low value columns to remove by checking the count of unique values by column
df_unique_counts = pd.DataFrame()
for col in data.columns:
    df_unique_counts = pd.concat([df_unique_counts, pd.DataFrame({
        "column": [col],
        "unique_values": [len(data[col].unique())]
    })], ignore_index=True)
df_unique_counts.sort_values(by="unique_values", ascending=False)

df_unique_counts['unique_values'].value_counts(dropna=False)

# let's look into the columns that have these values 
# unique_values
# 1        10
# 60       10
# 4         5
# 9         3
# 5         3

# print the columns that have 1 unique value
l1_unique_cols = df_unique_counts[df_unique_counts["unique_values"] == 1]["column"]
for col in l1_unique_cols:
    print(col)
    print(data[col].unique())
    print(data[col].value_counts(dropna=False))
    print("-"*100)

# remove these columns from the data 
len(l1_unique_cols)
data = data.drop(columns=l1_unique_cols)
len(data.columns)

# print the columns that have 60 unique values
l60_unique_cols = df_unique_counts[df_unique_counts["unique_values"] == 60]["column"]
for col in l60_unique_cols:
    print(col)
    print(data[col].unique())
    print(data[col].value_counts(dropna=False))
    print("-"*100)

# subset to the data in the l60_unique_cols and then drop duplicates
data_l60 = data[l60_unique_cols]
data_l60 = data_l60.drop_duplicates()
len(data_l60)

# save this subset to a csv file locally
data_l60 = data_l60.sort_values(by="primary_state_abbreviation")
data_l60.to_csv("data_l60.csv", index=False)

# keep common_state_fips_code,common_state_name, but drop the other columns in l60_unique_cols
len(l60_unique_cols)
l60_unique_cols = [col for col in l60_unique_cols if col not in ["common_state_fips_code", "common_state_name"]]
len(l60_unique_cols)

# remove these columns from the data 
data = data.drop(columns=l60_unique_cols)
len(data)
len(data.columns)

# print the columns that have 4 unique values
l4_unique_cols = df_unique_counts[df_unique_counts["unique_values"] == 4]["column"]
for col in l4_unique_cols:
    print(col)
    print(data[col].unique())
    print(data[col].value_counts(dropna=False))
    print("-"*100)

# these all look good, not majorly duplicative columns 

# print the columns that have 9 unique values
l9_unique_cols = df_unique_counts[df_unique_counts["unique_values"] == 9]["column"]
for col in l9_unique_cols:
    print(col)
    print(data[col].unique())
    print(data[col].value_counts(dropna=False))
    print("-"*100)

# crosstab of designation_type vs hpsa_type_code
pd.crosstab(data['designation_type'], data['hpsa_type_code'])

# looks to be a one to one mapping, we can remove hpsa_type_code
data = data.drop(columns = 'hpsa_type_code')


# print the columns that have 5 unique values
l5_unique_cols = df_unique_counts[df_unique_counts["unique_values"] == 5]["column"]
for col in l5_unique_cols:
    print(col)
    print(data[col].unique())
    print(data[col].value_counts(dropna=False))
    print("-"*100)

# differences, we can keep these 


# save the data to a csv file locally
data = data.sort_values(by="hpsa_name")
data.to_csv("data_columns_dropped.csv", index=False)

data['hpsa_name'].value_counts(dropna=False)