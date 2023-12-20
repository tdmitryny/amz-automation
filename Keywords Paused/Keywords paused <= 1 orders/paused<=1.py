import pandas as pd
import openpyxl
import warnings
from decimal import Decimal

warnings.simplefilter("ignore")

df1 = pd.read_excel("bulk.xlsx", sheet_name="Sponsored Products Campaigns")
df1.columns = df1.columns.str.replace(" ", "_")

#Convert to decimal
columns_to_number_df1 = ['Ad_Group_ID', 'Campaign_ID', 'Keyword_ID', 'Portfolio_ID', 'Ad_ID', 'Product_Targeting_ID']

def scientific_to_decimal(scientific_str):
    return Decimal(scientific_str)

for col in columns_to_number_df1:
    df1[col] = df1[col].apply(scientific_to_decimal)


df1.query("Entity == 'Keyword'", inplace=True)
df1["Operation"] = "update"
df1["State"] = "paused"


filt_df1 = df1.loc[(df1["Clicks"] >= 20) & (df1["Orders"] <= 1) & (df1["ACOS"] >= 0.31)].copy()


filt_df1.columns = df1.columns.str.replace("_", " ")

filt_df1.to_excel("keypaused.xlsx", sheet_name="keyword-stop", index=False)