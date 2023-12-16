import numpy as np
import pandas as pd
import openpyxl
import warnings
from decimal import Decimal
import time

warnings.simplefilter("ignore")
start = time.perf_counter()
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




# Optimization for Sponsored Products Campaigns
filt_df2 = df2.loc[(df2["Clicks"] >= 20) & (df2["Orders"] <= 1) \
                   & (df2["Customer_Search_Term"].str.contains("b0"))].copy()


#Working with ASINS
column_mappings_asins = {
            "Product": filt_df2["Product"],
            "Entity": "Negative Product Targeting",
            "Operation": "create",
            "Campaign_ID": filt_df2["Campaign_ID"],
            "Ad_Group_ID": filt_df2["Ad_Group_ID"],
            "Product_Targeting_ID": filt_df2["Product_Targeting_ID"],
            "Campaign_Name_(Informational_only)": filt_df2["Campaign_Name_(Informational_only)"],
            "Ad_Group_Name_(Informational_only)": filt_df2["Ad_Group_Name_(Informational_only)"],
            "State": "enabled",
            "Product_Targeting_Expression": filt_df2["Customer_Search_Term"].apply(lambda x: f'asin="{x.upper()}"'),
            "Clicks": filt_df2["Clicks"],
            "Orders": filt_df2["Orders"],
            "ACOS": filt_df2["ACOS"]
        }




for column, value in column_mappings_asins.items():
    df1[column] = value


#Scientific to decimal
columns_to_number_df1 = ['Ad_Group_ID', 'Campaign_ID', 'Keyword_ID', 'Portfolio_ID', \
                         'Product_Targeting_ID', 'Ad_Group_State_(Informational_only)']

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

end = time.perf_counter()
print(f"It took{end - start: .2f}s")