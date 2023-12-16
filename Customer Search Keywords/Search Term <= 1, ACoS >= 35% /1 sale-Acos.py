import numpy as np
import pandas as pd
import openpyxl
import warnings
from decimal import Decimal

warnings.simplefilter("ignore")

# Import data from Excel
df2 = pd.read_excel("bulk.xlsx", sheet_name="SP Search Term Report")



#Data frame
data = ['Product',
        'Entity',
        'Operation',
        'Campaign ID',
        'Ad Group ID',
        'Portfolio ID',
        'Ad ID',
        'Keyword ID',
        'Product Targeting ID',
        'Campaign Name',
        'Ad Group Name',
        'Campaign Name (Informational only)',
        'Ad Group Name (Informational only)',
        'Portfolio Name (Informational only)',
        'Start Date',
        'End Date',
        'Targeting Type',
        'State',
        'Campaign State (Informational only)',
        'Ad Group State (Informational only)',
        'Daily Budget',
        'SKU',
        'ASIN (Informational only)',
        'Eligibility Status (Informational only)',
        'Reason for Ineligibility (Informational only)',
        'Ad Group Default Bid',
        'Ad Group Default Bid (Informational only)',
        'Bid',
        'Keyword Text',
        'Match Type',
        'Bidding Strategy',
        'Placement',
        'Percentage',
        'Product Targeting Expression',
        'Resolved Product Targeting Expression (Informational only)',
        'Impressions',
        'Clicks',
        'Click-through Rate',
        'Spend',
        'Sales',
        'Orders',
        'Units',
        'Conversion Rate',
        'ACOS',
        'CPC',
        'ROAS'
        ]

df1 = pd.DataFrame(columns=data)
df1.columns = df1.columns.str.replace(" ", "_")


# Access columns I need to replace space with space
df2.columns = df2.columns.str.replace(" ", "_")


# Optimization for Sponsored Products Campaigns and remove ASINS
filt_df2 = df2.loc[(df2["Clicks"] >= 20) & (df2["Orders"] <= 1) & (~df2["Customer_Search_Term"].str.contains("b0"))].copy()


#Define the column mappings and values to assign

column_mappings = {
    "Product": filt_df2["Product"],
    "Entity": "Negative Keyword",
    "Operation": "Create",
    "Campaign_ID": filt_df2["Campaign_ID"],
    "Ad_Group_ID": filt_df2["Ad_Group_ID"],
    "Keyword_ID": filt_df2["Keyword_ID"],
    "Product_Targeting_ID": filt_df2["Product_Targeting_ID"],
    "Campaign_Name_(Informational_only)": filt_df2["Campaign_Name_(Informational_only)"],
    "State": "enabled",
    "Keyword_Text": filt_df2["Customer_Search_Term"],#remove dublication
    "Match_Type": "Negative Exact",
    "Product_Targeting_Expression": filt_df2["Product_Targeting_Expression"]
}


# Use a loop to assign values to df1
for column, value in column_mappings.items():
    df1[column] = value



#Scientific to decimal
columns_to_number_df1 = ['Ad_Group_ID', 'Campaign_ID', 'Keyword_ID', 'Portfolio_ID', 'Product_Targeting_ID']

def scientific_to_decimal(scientific_str):
    return Decimal(scientific_str)

for col in columns_to_number_df1:
    df1[col] = df1[col].apply(scientific_to_decimal)


# Reset index
filt_df2.reset_index(inplace=True, drop=True)
df1.columns = df1.columns.str.replace("_", " ")

# Saving data to Excel with sheets index=False
with (pd.ExcelWriter('new_data.xlsx') as writer):
    df1.to_excel(writer, sheet_name='upload', index=False)

