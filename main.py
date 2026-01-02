import pdfplumber
import pandas as pd
import re

PDF_PATH = "DC1002.pdf"
OUTPUT_FILE = "delivery_challan.xlsx"

def safe_search(pattern, text, group=1):
    match = re.search(pattern, text, re.IGNORECASE)
    return match.group(group) if match else ""


buyer_data = {}
items = []

with pdfplumber.open(PDF_PATH) as pdf:
    full_text = "\n".join(page.extract_text() for page in pdf.pages)

# -------------------------
# 1. Extract Buyer Details
# -------------------------
buyer_data["Buyer Name"] = safe_search(
    r"Buyer\s*\(Bill to\)\s*([\s\S]*?)\n", full_text)

buyer_data["GSTIN"] = safe_search(
    r"GSTIN/UIN\s*:\s*([A-Z0-9]+)", full_text)

buyer_data["License No"] = safe_search(
    r"Lic No:\s*([A-Z0-9\-]+)", full_text)

buyer_data["State"] = safe_search(
    r"State Name\s*:\s*([A-Za-z]+)", full_text)

buyer_data["Delivery Note"] = safe_search(
    r"(DC/\d+/\d+)", full_text)

buyer_data["Date"] = safe_search(
    r"Dated\s*([\d]{2}-[A-Za-z]{3}-[\d]{2})", full_text)

# -------------------------
# 2. Extract Items
# -------------------------
item_pattern = re.compile(
    r"(\d+)\s+([A-Z\s]+)\s([\d\.X]+)\s+([\d,]+\.\d+).*?(\d+\sPCS)(\d+)\nBatch\s*:\s*(\S+)\s+\d+\sPCS\nExpiry\s*:\s*([\d\-A-Za-z]+)",
    re.DOTALL
)

for match in item_pattern.findall(full_text):
    items.append({
        "Sl No": match[0],
        "Product Name": match[1].strip(),
        "Size": match[2],
        "Rate": match[3],
        "Quantity": match[4],
        "HSN": match[5],
        "Batch": match[6],
        "Expiry": match[7]
    })

# -------------------------
# 3. Write to Excel
# -------------------------
buyer_df = pd.DataFrame(list(buyer_data.items()), columns=["Field", "Value"])
items_df = pd.DataFrame(items)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    buyer_df.to_excel(writer, sheet_name="Buyer Details", index=False)
    items_df.to_excel(writer, sheet_name="Goods Details", index=False)

print("✅ Excel file created:", OUTPUT_FILE)
