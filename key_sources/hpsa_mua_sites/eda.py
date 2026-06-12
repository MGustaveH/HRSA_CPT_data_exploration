import pandas as pd

data = pd.read_excel("/Users/karlhonerlaw/Library/CloudStorage/GoogleDrive-honerlaw@gmail.com/My Drive/CesarK_CPT_project/data/key_data_sources/data_samples/HPSA MUAP Sites/HPSA – Primary Care/BCD_HPSA_FCT_DET_PC.xlsx")
len(data)

data.columns

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
unique_values
# 1        10
# 60       10
# 4         5
# 9         3
# 5         3

# print the columns that have 1 unique value
df_unique_counts[df_unique_counts["unique_values"] == 1]["column"]
