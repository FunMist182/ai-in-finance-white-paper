# -*- coding: utf-8 -*-
"""
Created on Tue Feb 17 14:20:40 2026

@author: david
"""

import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

# 1) File locations
# Code is inside the "Tesco code" folder, so this points one folder back
folder = Path(__file__).resolve().parent.parent

df = pd.read_excel(folder / "Tesco_Master_Dataset_2010_2015.xlsx")

# 2) Rename columns
df = df.rename(columns={
    "Revenue (£m)": "Revenue",
    "Profit before tax (£m)": "PBT",
    "Net cash from operating activities (£m)": "CFO",
    "Trade and other receivables (£m)": "Receivables",
    "Inventories (£m)": "Inventory",
    "Profit after tax (£m)": "PAT",
    "Total assets (£m)": "Total_Assets"
})

# Keep only the columns we need
df = df[[
    "Year", "Revenue", "PBT", "CFO",
    "Receivables", "Inventory", "PAT", "Total_Assets"
]].copy()

df = df.sort_values("Year")

# Check for missing values
print("Missing values:")
print(df.isna().sum())

# 3) Red-flag metrics

# Better accrual measure
df["Total_Accruals"] = df["PAT"] - df["CFO"]

# Average total assets
df["Average_Total_Assets"] = (
    df["Total_Assets"] + df["Total_Assets"].shift(1)
) / 2

# Accruals relative to assets
df["Accruals_to_Assets"] = (
    df["Total_Accruals"] / df["Average_Total_Assets"]
)

# Cash conversion
df["Cash_Conversion"] = df["CFO"] / df["PAT"]

# Working-capital ratios
df["Receivables_to_Sales"] = df["Receivables"] / df["Revenue"]
df["Inventory_to_Sales"] = df["Inventory"] / df["Revenue"]

# 4) Year-on-year changes
for c in [
    "Revenue",
    "PAT",
    "CFO",
    "Receivables",
    "Inventory"
]:
    df[c + "_YoY_%"] = df[c].pct_change(fill_method=None)

# 5) Directional warning flags

# Compare current accrual ratio with the previous 3-year median
previous_accrual_median = (
    df["Accruals_to_Assets"]
    .shift(1)
    .rolling(3)
    .median()
)

df["Flag_Positive_Accrual_Build"] = (
    df["Accruals_to_Assets"] > previous_accrual_median
)

# Receivables growing more than 5 percentage points faster than revenue
df["Flag_Receivables_Growing_Faster"] = (
    df["Receivables_YoY_%"] >
    df["Revenue_YoY_%"] + 0.05
)

# Inventory growing more than 5 percentage points faster than revenue
df["Flag_Inventory_Growing_Faster"] = (
    df["Inventory_YoY_%"] >
    df["Revenue_YoY_%"] + 0.05
)

# First year has no previous-year comparison
flag_columns = [
    "Flag_Receivables_Growing_Faster",
    "Flag_Inventory_Growing_Faster"
]

for col in flag_columns:
    df[col] = df[col].astype("boolean")

df.loc[df["Year"] == df["Year"].min(), flag_columns] = pd.NA

print(df[[
    "Year",
    "Revenue",
    "PAT",
    "CFO",
    "Total_Accruals",
    "Accruals_to_Assets",
    "Flag_Positive_Accrual_Build",
    "Flag_Receivables_Growing_Faster",
    "Flag_Inventory_Growing_Faster"
]])

# 6) Plot trends

plt.figure()
plt.plot(df["Year"], df["Accruals_to_Assets"] * 100, marker="o")
plt.title("Accruals to Average Total Assets")
plt.xlabel("Year")
plt.ylabel("%")
plt.grid(True)

plt.savefig(
    folder / "accruals_to_assets.png",
    dpi=300,
    bbox_inches="tight"
)

plt.show()

plt.figure()
plt.plot(df["Year"], df["Cash_Conversion"], marker="o")
plt.title("Cash Conversion (CFO / Profit After Tax)")
plt.xlabel("Year")
plt.ylabel("Ratio")
plt.grid(True)

plt.savefig(
    folder / "cash_conversion.png",
    dpi=300,
    bbox_inches="tight"
)

plt.show()

# 7) Save output
df.to_csv(
    folder / "tesco_redflags_updated.csv",
    index=False
)

print("Saved: tesco_redflags_updated.csv")