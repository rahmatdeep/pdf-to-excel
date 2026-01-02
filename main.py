import pandas as pd
import PyPDF2
import re
from datetime import datetime

def extract_text_from_pdf(pdf_path):
    """Extract text from PDF file"""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            text = ""
            for page in pdf_reader.pages:
                text += page.extract_text()
            return text
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return None

def extract_buyer_details(text):
    """Extract buyer information from PDF text"""
    buyer_info = {}
    
    # Find the buyer section (after "Buyer (Bill to)")
    buyer_section = re.search(r'Buyer \(Bill to\)(.*?)(?=Delivery Note No\.|e-Way Bill)', text, re.DOTALL)
    
    if buyer_section:
        buyer_text = buyer_section.group(1)
        
        # Extract address lines
        lines = [line.strip() for line in buyer_text.split('\n') if line.strip() and not line.strip().startswith('Lic No') and not line.strip().startswith('GSTIN') and not line.strip().startswith('State Name')]
        
        buyer_info['Buyer Name'] = lines[0] if len(lines) > 0 else ''
        buyer_info['Address Line 1'] = lines[1] if len(lines) > 1 else ''
        buyer_info['Address Line 2'] = lines[2] if len(lines) > 2 else ''
        buyer_info['Address Line 3'] = lines[3] if len(lines) > 3 else ''
        buyer_info['City'] = lines[4] if len(lines) > 4 else ''
        
        # Extract License No from buyer section
        lic_match = re.search(r'Lic No[:\s]+([A-Z0-9\-]+)', buyer_text)
        if lic_match:
            buyer_info['License No'] = lic_match.group(1).strip()
        
        # Extract buyer GSTIN (should be the second GSTIN in the document)
        gstin_match = re.search(r'GSTIN/UIN\s*:\s*([A-Z0-9]+)', buyer_text)
        if gstin_match:
            buyer_info['GSTIN'] = gstin_match.group(1).strip()
        
        # Extract buyer State
        state_match = re.search(r'State Name\s*:\s*([^,]+),\s*Code\s*:\s*(\d+)', buyer_text)
        if state_match:
            buyer_info['State'] = state_match.group(1).strip()
            buyer_info['State Code'] = state_match.group(2).strip()
    
    return buyer_info

def extract_supplier_details(text):
    """Extract supplier information from PDF text"""
    supplier_info = {}
    
    # Extract supplier name
    supplier_match = re.search(r'DELIVERY CHALLAN\s*\n([^\n]+)', text)
    if supplier_match:
        supplier_info['Supplier Name'] = supplier_match.group(1).strip()
    
    # Extract supplier address
    address_match = re.search(r'DELIVERY CHALLAN\s*\n[^\n]+\s*\n([^\n]+)', text)
    if address_match:
        supplier_info['Supplier Address'] = address_match.group(1).strip()
    
    # Extract supplier GSTIN (first occurrence is supplier)
    gstin_matches = re.findall(r'GSTIN/UIN:\s*([A-Z0-9]+)', text)
    if gstin_matches:
        supplier_info['Supplier GSTIN'] = gstin_matches[0].strip()
    
    # Extract License
    license_match = re.search(r'License No-\s*([A-Z0-9/\-]+)', text)
    if license_match:
        supplier_info['Supplier License'] = license_match.group(1).strip()
    
    # Extract Email
    email_match = re.search(r'E-Mail\s*:\s*([^\s]+)', text)
    if email_match:
        supplier_info['Supplier Email'] = email_match.group(1).strip()
    
    return supplier_info

def extract_products(text):
    """Extract product details including batch numbers and expiry dates"""
    products = []
    
    # Pattern to match product lines with all details
    # This pattern looks for: Sl No, Description, HSN, Quantity, Rate, Amount
    product_pattern = r'(\d+)\s+([\w\s]+?)\s+(\d+[\.,]\d+)\s+PCS(\d+[\.,]\d+)\s+(\d+)\s+PCS\s+(\d+)'
    
    # Pattern for batch and expiry
    batch_pattern = r'Batch\s*:\s*([A-Z0-9]+)'
    expiry_pattern = r'Expiry\s*:\s*([0-9\-A-Za-z]+)'
    
    # Split text into lines for processing
    lines = text.split('\n')
    
    current_product = None
    for i, line in enumerate(lines):
        # Look for product description patterns
        if 'TRACKFLEX' in line or 'POLARIS' in line:
            # Extract product description and details
            parts = line.split()
            
            # Try to find serial number before this line
            sl_no = None
            for j in range(max(0, i-2), i):
                if lines[j].strip().isdigit():
                    sl_no = int(lines[j].strip())
                    break
            
            # Extract description
            if 'TRACKFLEX' in line:
                desc_match = re.search(r'(TRACKFLEX\s+[\d\.]+X\d+)', line)
            elif 'POLARIS NC BALLOON' in line:
                desc_match = re.search(r'(POLARIS NC BALLOON\s+[\d\.]+X\d+)', line)
            elif 'POLARIS SC BALLOON' in line:
                desc_match = re.search(r'(POLARIS SC BALLOON\s+[\d\.]+X\d+)', line)
            else:
                desc_match = None
            
            if desc_match:
                description = desc_match.group(1)
                
                # Extract rate and amount
                amount_match = re.search(r'(\d+[\.,]\d+)\s+PCS(\d+[\.,]\d+)', line)
                rate = None
                amount = None
                if amount_match:
                    rate = float(amount_match.group(1).replace(',', ''))
                    amount = float(amount_match.group(2).replace(',', ''))
                
                # Extract HSN code
                hsn_match = re.search(r'(\d{8})', line)
                hsn = hsn_match.group(1) if hsn_match else ''
                
                current_product = {
                    'Sl_No': sl_no,
                    'Description': description,
                    'HSN_SAC': hsn,
                    'Quantity': 1,
                    'Unit': 'PCS',
                    'Rate': rate,
                    'Amount': amount,
                    'Batch': '',
                    'Expiry': ''
                }
        
        # Look for batch number
        if current_product and 'Batch' in line:
            batch_match = re.search(batch_pattern, line)
            if batch_match:
                current_product['Batch'] = batch_match.group(1).strip()
        
        # Look for expiry date
        if current_product and 'Expiry' in line:
            expiry_match = re.search(expiry_pattern, line)
            if expiry_match:
                current_product['Expiry'] = expiry_match.group(1).strip()
                # Product is complete, add to list
                products.append(current_product.copy())
                current_product = None
    
    return products

def extract_document_info(text):
    """Extract document metadata"""
    doc_info = {}
    
    # Extract DC number
    dc_match = re.search(r'DC/\d+/\d+', text)
    if dc_match:
        doc_info['DC Number'] = dc_match.group(0)
    
    # Extract e-Way Bill
    eway_match = re.search(r'e-Way Bill No\.\s*\n([0-9\s]+)', text)
    if eway_match:
        doc_info['e-Way Bill'] = eway_match.group(1).strip()
    
    # Extract date
    date_match = re.search(r'dt\.\s*([0-9\-A-Za-z]+)', text)
    if date_match:
        doc_info['Date'] = date_match.group(1).strip()
    
    # Extract destination
    dest_match = re.search(r'Destination\s*\n([^\n]+)', text)
    if dest_match:
        doc_info['Destination'] = dest_match.group(1).strip()
    
    return doc_info

def create_excel_file(buyer_info, supplier_info, products, doc_info, filename='Delivery_Challan_Extracted.xlsx'):
    """Create Excel file with all extracted data"""
    
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        
        # Sheet 1: Document Info
        doc_df = pd.DataFrame([doc_info])
        doc_df.to_excel(writer, sheet_name='Document Info', index=False)
        
        # Sheet 2: Buyer Details
        buyer_df = pd.DataFrame([buyer_info])
        buyer_df.to_excel(writer, sheet_name='Buyer Details', index=False)
        
        # Sheet 3: Supplier Details
        supplier_df = pd.DataFrame([supplier_info])
        supplier_df.to_excel(writer, sheet_name='Supplier Details', index=False)
        
        # Sheet 4: Product Details
        products_df = pd.DataFrame(products)
        products_df.to_excel(writer, sheet_name='Product Details', index=False)
        
        # Sheet 5: Summary
        summary_data = {
            'Total Items': [len(products)],
            'Total Quantity': [sum(p['Quantity'] for p in products if p.get('Quantity'))],
            'Total Amount': [sum(p['Amount'] for p in products if p.get('Amount'))],
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    print(f"\n✅ Excel file '{filename}' created successfully!")
    return filename

def main(pdf_path):
    """Main function to process PDF and create Excel"""
    
    print("="*60)
    print("PDF TO EXCEL CONVERTER - DELIVERY CHALLAN")
    print("="*60)
    
    # Step 1: Extract text from PDF
    print("\n📄 Extracting text from PDF...")
    text = extract_text_from_pdf(pdf_path)
    
    if not text:
        print("❌ Failed to extract text from PDF")
        return
    
    print("✅ Text extracted successfully")
    
    # Step 2: Extract all information
    print("\n🔍 Extracting buyer details...")
    buyer_info = extract_buyer_details(text)
    
    print("🔍 Extracting supplier details...")
    supplier_info = extract_supplier_details(text)
    
    print("🔍 Extracting product details...")
    products = extract_products(text)
    
    print("🔍 Extracting document information...")
    doc_info = extract_document_info(text)
    
    # Step 3: Create Excel file
    print("\n📊 Creating Excel file...")
    filename = create_excel_file(buyer_info, supplier_info, products, doc_info)
    
    # Step 4: Display summary
    print("\n" + "="*60)
    print("EXTRACTION SUMMARY")
    print("="*60)
    print(f"Buyer: {buyer_info.get('Buyer Name', 'N/A')}")
    print(f"Supplier: {supplier_info.get('Supplier Name', 'N/A')}")
    print(f"Total Products Extracted: {len(products)}")
    if products:
        total_amount = sum(p.get('Amount', 0) for p in products if p.get('Amount'))
        print(f"Total Amount: ₹{total_amount:,.2f}")
    print("="*60)
    print(f"\n✅ All data saved to: {filename}")

if __name__ == "__main__":
    # Specify your PDF file path here
    pdf_file_path = "DC1002.pdf"  # Change this to your PDF file path
    
    # Run the extraction
    main(pdf_file_path)
    
    print("\n💡 Required libraries: pip install pandas openpyxl PyPDF2")