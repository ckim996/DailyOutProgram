import pandas as pd

TOOLS_FILE = "data/DailyOutTools.xlsx"

db = pd.read_excel(TOOLS_FILE, sheet_name="DB")
nonmounts = pd.read_excel(TOOLS_FILE, sheet_name="Nonmounts")

db["SKU"] = db["SKU"].astype(str).str.strip()
nonmounts["SKU"] = nonmounts["SKU"].astype(str).str.strip()

def lookup_sku(sku):

    """
        Returns the row of SKU in DailyOutTools.xlsx on Sheet names: DB or Nonmounts
    """

    match = db[db["SKU"] == sku]
    if not match.empty:
        return match.iloc[0]

    match = nonmounts[nonmounts["SKU"] == sku]
    if not match.empty:
        row = match.iloc[0].copy()
        row["Interchange (not in order)"] = row.get("Category",None)
        return row

    return None
