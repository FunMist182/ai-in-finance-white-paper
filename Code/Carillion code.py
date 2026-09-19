# -*- coding: utf-8 -*-
"""
Carillion, 2010-2016: adapted from David's Tesco code.py.
Requires pandas and matplotlib. Run this file to display the two charts and
save the PNGs and CSV alongside it. The data are embedded so no Excel file
is needed. All financial amounts are GBP millions; years end 31 December.

Data sources (Carillion annual reports; printed page numbers):
2010: 2010 report, pp. 54, 57, 58.
2011: 2012 report, comparative columns, pp. 64, 67, 68.
2012-2013: 2013 report, pp. 74, 77, 78.
2014-2015: 2015 report, pp. 80, 83, 84.
2016: 2016 report, pp. 90, 93, 94 (original published accounts).
https://www.annualreports.co.uk/HostedData/AnnualReportArchive/c/LSE_CLLN_2010.pdf
https://www.annualreports.co.uk/HostedData/AnnualReportArchive/c/LSE_CLLN_2012.pdf
https://www.annualreports.co.uk/HostedData/AnnualReportArchive/c/LSE_CLLN_2013.pdf
https://www.annualreports.co.uk/HostedData/AnnualReportArchive/c/LSE_CLLN_2015.pdf
https://www.petercrow.com/storage/12406/a1926d75-5407-4a05-872b-9300cd8ebd5e/0930aq-carillion-annual-report-2016-original.pdf

Definitions: revenue is statutory Group revenue excluding joint ventures'
revenue; PAT is total Group profit for the year, including non-controlling
interests; CFO is net cash flows from operating activities after the items
classified there by Carillion; receivables are the balance-sheet current
trade and other receivables. PBT and PAT are reported, not underlying profits.

Source-vintage limitation: 2012 uses the 2013 restated comparative (IAS 19
and acquisition accounting); 2014 uses the 2015 comparative balance sheet.
2010-2011 use the earlier reports above, before the 2013 IAS 19 restatement.
This is not a fully retrospectively harmonised or strictly as-first-reported
series. Accounting changes can affect year-on-year comparisons.
The existing 2015 data are retained: the 2016 report restates 2015
receivables to 1270.8 and total assets to 3869.5; this file retains 1270.5
and 3870.1 from the 2015 report to preserve the previous analysis.

The Tesco calculations are retained, including missing 2010 average assets.
Accrual flags require three previous annual accrual ratios: 2010-2013 are
unavailable and 2014 is the first eligible year. Flags use nullable booleans
internally; missing flags display/export as "Unavailable", never False.
The *_YoY_% columns retain pandas' fractional changes (0.05 means 5%).
These fixed rules are prompts for investigation, not an insolvency forecast.
This retrospective case study does not demonstrate prediction of collapse.
"""

import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

# 1) File locations and Carillion data
folder = Path(__file__).resolve().parent

df = pd.DataFrame({
    "Year": [2010, 2011, 2012, 2013, 2014, 2015, 2016],
    "Revenue (£m)": [4236.5, 4153.2, 3666.2, 3332.6, 3493.9, 3950.7, 4394.9],
    "Profit before tax (£m)": [167.9, 142.8, 164.8, 110.6, 142.6, 155.1, 146.7],
    "Net cash from operating activities (£m)": [126.8, 103.2, -25.6, -78.4, 123.8, 73.3, 73.3],
    "Trade and other receivables (£m)": [1052.4, 1094.6, 1111.5, 1212.3, 1328.0, 1270.5, 1664.0],
    "Inventories (£m)": [40.6, 71.6, 55.3, 48.6, 50.1, 64.3, 78.8],
    "Profit after tax (£m)": [152.8, 138.0, 154.9, 106.3, 127.5, 139.4, 129.5],
    "Total assets (£m)": [3151.3, 3699.2, 3862.4, 3639.9, 3896.4, 3870.1, 4433.1]
})

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
    .rolling(3, min_periods=3)
    .median()
)

df["Flag_Positive_Accrual_Build"] = (
    df["Accruals_to_Assets"] > previous_accrual_median
).astype("boolean")
accrual_available = (
    df["Accruals_to_Assets"].notna() & previous_accrual_median.notna()
)
df.loc[~accrual_available, "Flag_Positive_Accrual_Build"] = pd.NA

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

# Label unavailable flags explicitly in the display and CSV while keeping
# numeric missing values blank in the CSV and boolean flags in df.
output_df = df.copy()
all_flag_columns = ["Flag_Positive_Accrual_Build"] + flag_columns
for col in all_flag_columns:
    output_df[col] = df[col].astype(object).where(df[col].notna(), "Unavailable")

print(output_df[[
    "Year",
    "Revenue",
    "PAT",
    "CFO",
    "Total_Accruals",
    "Accruals_to_Assets",
    "Flag_Positive_Accrual_Build",
    "Flag_Receivables_Growing_Faster",
    "Flag_Inventory_Growing_Faster"
]].to_string(index=False))

# 6) Plot trends

plt.figure()
plt.plot(df["Year"], df["Accruals_to_Assets"] * 100, marker="o")
plt.title("Accruals to Average Total Assets")
plt.xlabel("Year")
plt.ylabel("%")
plt.grid(True)

plt.savefig(
    folder / "carillion_accruals_to_assets.png",
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
    folder / "carillion_cash_conversion.png",
    dpi=300,
    bbox_inches="tight"
)

plt.show()

# 7) Save output
output_df.to_csv(
    folder / "carillion_redflags_updated.csv",
    index=False
)

print("Saved: carillion_redflags_updated.csv")
