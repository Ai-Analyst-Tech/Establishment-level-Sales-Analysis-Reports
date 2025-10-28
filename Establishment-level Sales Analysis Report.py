#AGGREGATING QUANTITIES BY ESTABLISHMENT AND ITEM CODE
from pathlib import Path
import pandas as pd

#Define file paths
base = Path(r"c:\Users\Hp\Desktop\20th Oct to 26th Oct 2025 Weekly Sales Report")
input_file = base / "20th Oct 25 - 26thOct 25 Weekly Sales Analysis Report.xlsx"   # <-- Correct extension
output_file = base / "Weekly_Sales_Report1.xlsx"

#Read Excel file
data = pd.read_excel(input_file)

#Normalize Shop code 
def normalize_shop_code(code):
    code = str(code).strip()  # Convert to string and remove spaces
    code = code.replace("08S", "008")  # Special case mapping
    code = code.zfill(3) if code.isdigit() else code  # Pad numbers to 3 digits
    return code

data["Shop code"] = data["Shop code"].apply(normalize_shop_code)

#First aggregation (by Item code)
agg_df = data.groupby(["Item code"]).agg({
    "Quantity": "sum",
    "Total amount": "sum",
    "Buying Price": "first"
}).reset_index()

#Second aggregation (by Shop code, Item code, etc.)
agg_data = (
    data.groupby(["Shop code", "Item code", "Description", "Category"], as_index=False)
    .agg({
        "Quantity": "sum",
        "Unit selling price": "first",
        "Total amount": "sum",
        "Buying Price": "first",
        "Mark Up %": "first"
    })
)

#Export to Excel with each shop in its own sheet
with pd.ExcelWriter(output_file, engine="openpyxl", mode="w") as writer:
    agg_df.to_excel(writer, sheet_name="By_Item_Code", index=False)
    agg_data.to_excel(writer, sheet_name="By_Shop_Item", index=False)

    #Separate sheet for each shop
    for shop, shop_df in agg_data.groupby("Shop code"):
        sheet_name = f"Shop_{shop}"
        sheet_name = sheet_name[:31]  #Excel sheet names max 31 chars
        shop_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Aggregated data exported successfully to: {output_file}")
