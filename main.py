import pandas as pd
from datetime import datetime
import re

def extract_delivery_challan_data(pdf_text):
    """
    Extract buyer details and product information from delivery challan text
    """
    
    # Extract Buyer Details
    buyer_info = {
        'Buyer Name': 'Amrit Pharmacy- AIIMS, Deoghar',
        'Address Line 1': 'Ground Floor, Night Shelter Building',
        'Address Line 2': 'Campus Of AIIMS, Deoghar',
        'Address Line 3': 'PO & PS- Devipur',
        'City': 'Deoghar',
        'State': 'Jharkhand',
        'State Code': '20',
        'License No': 'JH-DEO-132417',
        'GSTIN': '20AAACH5598K1ZF'
    }
    
    # Extract Supplier Details
    supplier_info = {
        'Supplier Name': 'OPTIMA HEALTHCARE',
        'Supplier Address': '125, Vardhaman Crown Mall, Plot No-2, Sector-19 Dwarka New Delhi',
        'Supplier GSTIN': '07DYGPG3515M1ZA',
        'Supplier License': 'RMD/DCD/HO-2638/3641',
        'Supplier Email': 'healthcareoptimaa@gmail.com'
    }
    
    # Product Data
    products = [
        # TRACKFLEX Products
        {'Sl_No': 1, 'Description': 'TRACKFLEX 2.50X20', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251725020012', 'Expiry': '19-Aug-27'},
        {'Sl_No': 2, 'Description': 'TRACKFLEX 2.50X24', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251125024017', 'Expiry': '10-Jun-27'},
        {'Sl_No': 3, 'Description': 'TRACKFLEX 2.50X28', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250625028012', 'Expiry': '29-Apr-27'},
        {'Sl_No': 4, 'Description': 'TRACKFLEX 2.50X32', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250725032002', 'Expiry': '6-May-27'},
        {'Sl_No': 5, 'Description': 'TRACKFLEX 2.50X36', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250825036011', 'Expiry': '15-May-27'},
        {'Sl_No': 6, 'Description': 'TRACKFLEX 2.50X40', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250825040005', 'Expiry': '15-May-27'},
        {'Sl_No': 7, 'Description': 'TRACKFLEX 2.50X43', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250925043004', 'Expiry': '27-May-27'},
        {'Sl_No': 8, 'Description': 'TRACKFLEX 2.75X13', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250727513005', 'Expiry': '6-May-27'},
        {'Sl_No': 9, 'Description': 'TRACKFLEX 2.75X16', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250727516014', 'Expiry': '6-May-27'},
        {'Sl_No': 10, 'Description': 'TRACKFLEX 2.75X20', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250727520007', 'Expiry': '6-May-27'},
        {'Sl_No': 11, 'Description': 'TRACKFLEX 2.75X24', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250927524007', 'Expiry': '27-May-27'},
        {'Sl_No': 12, 'Description': 'TRACKFLEX 2.75X28', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250727528019', 'Expiry': '6-May-27'},
        {'Sl_No': 13, 'Description': 'TRACKFLEX 2.75X36', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251327536018', 'Expiry': '17-Jun-27'},
        {'Sl_No': 14, 'Description': 'TRACKFLEX 2.75X40', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250927540001', 'Expiry': '27-May-27'},
        {'Sl_No': 15, 'Description': 'TRACKFLEX 2.75X43', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250727543001', 'Expiry': '6-May-27'},
        {'Sl_No': 16, 'Description': 'TRACKFLEX 3.00X13', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250730013007', 'Expiry': '6-May-27'},
        {'Sl_No': 17, 'Description': 'TRACKFLEX 3.00X16', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250930016002', 'Expiry': '27-May-27'},
        {'Sl_No': 18, 'Description': 'TRACKFLEX 3.00X20', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250930020011', 'Expiry': '27-May-27'},
        {'Sl_No': 19, 'Description': 'TRACKFLEX 3.00X28', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250830028023', 'Expiry': '15-May-27'},
        {'Sl_No': 20, 'Description': 'TRACKFLEX 3.00X32', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251330032022', 'Expiry': '17-Jun-27'},
        {'Sl_No': 21, 'Description': 'TRACKFLEX 3.00X36', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250830036003', 'Expiry': '15-May-27'},
        {'Sl_No': 22, 'Description': 'TRACKFLEX 3.00X40', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250730040001', 'Expiry': '6-May-27'},
        {'Sl_No': 23, 'Description': 'TRACKFLEX 3.00X43', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250830043002', 'Expiry': '15-May-27'},
        {'Sl_No': 24, 'Description': 'TRACKFLEX 3.50X13', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250735013004', 'Expiry': '6-May-27'},
        {'Sl_No': 25, 'Description': 'TRACKFLEX 3.50X16', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250635016021', 'Expiry': '29-Apr-27'},
        {'Sl_No': 26, 'Description': 'TRACKFLEX 3.50X20', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250735020029', 'Expiry': '6-May-27'},
        {'Sl_No': 27, 'Description': 'TRACKFLEX 3.50X24', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250735024025', 'Expiry': '6-May-27'},
        {'Sl_No': 28, 'Description': 'TRACKFLEX 3.50X28', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250835028008', 'Expiry': '15-May-27'},
        {'Sl_No': 29, 'Description': 'TRACKFLEX 3.50X32', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250835032022', 'Expiry': '15-May-27'},
        {'Sl_No': 30, 'Description': 'TRACKFLEX 3.50X36', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S242335036011', 'Expiry': '2-Dec-26'},
        {'Sl_No': 31, 'Description': 'TRACKFLEX 3.50X40', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251035040008', 'Expiry': '30-May-27'},
        {'Sl_No': 32, 'Description': 'TRACKFLEX 3.50X43', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250935043002', 'Expiry': '27-May-27'},
        {'Sl_No': 33, 'Description': 'TRACKFLEX 4.00X16', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251540016019', 'Expiry': '6-Aug-27'},
        {'Sl_No': 34, 'Description': 'TRACKFLEX 4.00X20', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S242340020004', 'Expiry': '2-Dec-26'},
        {'Sl_No': 35, 'Description': 'TRACKFLEX 4.00X24', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251740024005', 'Expiry': '19-Aug-27'},
        {'Sl_No': 36, 'Description': 'TRACKFLEX 4.00X28', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251640028001', 'Expiry': '12-Aug-27'},
        {'Sl_No': 37, 'Description': 'TRACKFLEX 4.00X32', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251340032008', 'Expiry': '17-Jun-27'},
        {'Sl_No': 38, 'Description': 'TRACKFLEX 4.00X36', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S242340036002', 'Expiry': '2-Dec-26'},
        {'Sl_No': 39, 'Description': 'TRACKFLEX 4.50X13', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251645013004', 'Expiry': '12-Aug-27'},
        {'Sl_No': 40, 'Description': 'TRACKFLEX 4.50X20', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251745020001', 'Expiry': '19-Aug-27'},
        {'Sl_No': 41, 'Description': 'TRACKFLEX 4.50X24', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S251345024002', 'Expiry': '17-Jun-27'},
        {'Sl_No': 42, 'Description': 'TRACKFLEX 4.50X32', 'HSN_SAC': '90219090', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 20500.00, 'Amount': 20500.00, 'Batch': 'S250845032002', 'Expiry': '15-May-27'},
        # POLARIS NC BALLOON Products
        {'Sl_No': 43, 'Description': 'POLARIS NC BALLOON 2.50X10', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC25010C2505008', 'Expiry': '28-Oct-28'},
        {'Sl_No': 44, 'Description': 'POLARIS NC BALLOON 2.50X20', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC25020C2404003', 'Expiry': '17-Jun-27'},
        {'Sl_No': 45, 'Description': 'POLARIS NC BALLOON 2.75X08', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC27508C2404014', 'Expiry': '17-Jun-27'},
        {'Sl_No': 46, 'Description': 'POLARIS NC BALLOON 2.75X10', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC27510C2505016', 'Expiry': '28-Oct-28'},
        {'Sl_No': 47, 'Description': 'POLARIS NC BALLOON 2.75X15', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC27515C2502066', 'Expiry': '12-Feb-28'},
        {'Sl_No': 48, 'Description': 'POLARIS NC BALLOON 2.75X20', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC27520C2404013', 'Expiry': '17-Jun-27'},
        {'Sl_No': 49, 'Description': 'POLARIS NC BALLOON 3.00X10', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC30010C2504065', 'Expiry': '28-May-28'},
        {'Sl_No': 50, 'Description': 'POLARIS NC BALLOON 3.00X12', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC30012C2504043', 'Expiry': '28-May-28'},
        {'Sl_No': 51, 'Description': 'POLARIS NC BALLOON 3.50X20', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC35020C2505036', 'Expiry': '28-Oct-28'},
        {'Sl_No': 52, 'Description': 'POLARIS NC BALLOON 4.00X10', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC40010C2303006', 'Expiry': '30-Jul-26'},
        {'Sl_No': 53, 'Description': 'POLARIS NC BALLOON 4.00X12', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'NC40012C2302025', 'Expiry': '12-Jul-26'},
        # POLARIS SC BALLOON Products
        {'Sl_No': 54, 'Description': 'POLARIS SC BALLOON 2.00X15', 'HSN_SAC': '90183920', 'Quantity': 1, 'Unit': 'PCS', 'Rate': 7200.00, 'Amount': 7200.00, 'Batch': 'SC20015C2501042', 'Expiry': '11-Feb-28'},
    ]
    
    return buyer_info, supplier_info, products


def create_excel_file(buyer_info, supplier_info, products, filename='Delivery_Challan_DC_25_1002.xlsx'):
    """
    Create Excel file with buyer details and product information
    """
    
    # Create a Pandas Excel writer
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        
        # Sheet 1: Buyer & Supplier Details
        buyer_df = pd.DataFrame([buyer_info])
        buyer_df.to_excel(writer, sheet_name='Buyer Details', index=False)
        
        supplier_df = pd.DataFrame([supplier_info])
        supplier_df.to_excel(writer, sheet_name='Supplier Details', index=False)
        
        # Sheet 2: Product Details
        products_df = pd.DataFrame(products)
        products_df.to_excel(writer, sheet_name='Product Details', index=False)
        
        # Sheet 3: Summary
        summary_data = {
            'Delivery Challan No': ['DC/25/1002'],
            'e-Way Bill No': ['7915 8179 3578'],
            'Date': ['23-Nov-25'],
            'Total Items': [len(products)],
            'Total Quantity': [sum(p['Quantity'] for p in products)],
            'Total Amount': [sum(p['Amount'] for p in products)],
            'Destination': ['Jharkhand']
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    print(f"Excel file '{filename}' created successfully!")
    print(f"Total products: {len(products)}")
    print(f"Total amount: ₹{sum(p['Amount'] for p in products):,.2f}")


# Main execution
if __name__ == "__main__":
    # Extract data
    buyer_info, supplier_info, products = extract_delivery_challan_data("")
    
    # Create Excel file
    create_excel_file(buyer_info, supplier_info, products)
    
    # Display summary
    print("\n" + "="*50)
    print("EXTRACTION SUMMARY")
    print("="*50)
    print(f"Buyer: {buyer_info['Buyer Name']}")
    print(f"Location: {buyer_info['City']}, {buyer_info['State']}")
    print(f"Total Products: {len(products)}")
    print(f"TRACKFLEX items: {sum(1 for p in products if 'TRACKFLEX' in p['Description'])}")
    print(f"POLARIS items: {sum(1 for p in products if 'POLARIS' in p['Description'])}")
    print("="*50)