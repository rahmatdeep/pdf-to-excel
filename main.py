import streamlit as st
import pandas as pd
import PyPDF2
import re
from io import BytesIO
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Delivery Challan Extractor",
    page_icon="📄",
    layout="wide"
)

def extract_text_from_pdf(pdf_file):
    """Extract text from uploaded PDF file"""
    try:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page in pdf_reader.pages:
            text += page.extract_text()
        return text
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
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
    
    # Extract supplier section (before Buyer section)
    supplier_section = re.search(r'DELIVERY CHALLAN(.*?)Buyer \(Bill to\)', text, re.DOTALL)
    
    if supplier_section:
        supplier_text = supplier_section.group(1)
        
        # Extract supplier name (usually first line after DELIVERY CHALLAN)
        lines = [line.strip() for line in supplier_text.split('\n') if line.strip()]
        if lines:
            supplier_info['Supplier Name'] = lines[0]
        
        # Extract supplier address
        if len(lines) > 1:
            supplier_info['Supplier Address'] = lines[1]
        
        # Extract supplier GSTIN (first occurrence)
        gstin_match = re.search(r'GSTIN/UIN:\s*([A-Z0-9]+)', supplier_text)
        if gstin_match:
            supplier_info['Supplier GSTIN'] = gstin_match.group(1).strip()
        
        # Extract License
        license_match = re.search(r'License No-\s*([A-Z0-9/\-]+)', supplier_text)
        if license_match:
            supplier_info['Supplier License'] = license_match.group(1).strip()
        
        # Extract Email
        email_match = re.search(r'E-Mail\s*:\s*([^\s]+)', supplier_text)
        if email_match:
            supplier_info['Supplier Email'] = email_match.group(1).strip()
    
    return supplier_info

def extract_products(text):
    """Extract product details including batch numbers and expiry dates"""
    products = []
    
    # Pattern for batch and expiry
    batch_pattern = r'Batch\s*:\s*([A-Z0-9]+)'
    expiry_pattern = r'Expiry\s*:\s*([0-9\-A-Za-z]+)'
    
    # Split text into lines for processing
    lines = text.split('\n')
    
    current_product = None
    sl_no_counter = 1
    
    for i, line in enumerate(lines):
        # Look for product description patterns
        if 'TRACKFLEX' in line or 'POLARIS' in line:
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
                
                # Extract rate and amount - improved pattern
                # Looking for pattern like: 20,500.00 PCS20,500.00 or 7,200.00 PCS7,200.00
                amount_match = re.search(r'([\d,]+\.?\d*)\s*PCS\s*([\d,]+\.?\d*)', line)
                rate = None
                amount = None
                if amount_match:
                    rate_str = amount_match.group(1).replace(',', '')
                    amount_str = amount_match.group(2).replace(',', '')
                    try:
                        rate = float(rate_str)
                        amount = float(amount_str)
                    except ValueError:
                        # If conversion fails, try to find numbers in the line
                        numbers = re.findall(r'[\d,]+\.?\d*', line)
                        if len(numbers) >= 2:
                            try:
                                rate = float(numbers[-2].replace(',', ''))
                                amount = float(numbers[-1].replace(',', ''))
                            except:
                                rate = 0.0
                                amount = 0.0
                else:
                    # Fallback: try to find all numbers in the line
                    numbers = re.findall(r'[\d,]+\.?\d*', line)
                    if len(numbers) >= 2:
                        try:
                            rate = float(numbers[-2].replace(',', ''))
                            amount = float(numbers[-1].replace(',', ''))
                        except:
                            rate = 0.0
                            amount = 0.0
                
                # Extract HSN code
                hsn_match = re.search(r'(\d{8})', line)
                hsn = hsn_match.group(1) if hsn_match else ''
                
                current_product = {
                    'Sl_No': sl_no_counter,
                    'Description': description,
                    'HSN_SAC': hsn,
                    'Quantity': 1,
                    'Unit': 'PCS',
                    'Rate': rate if rate else 0.0,
                    'Amount': amount if amount else 0.0,
                    'Batch': '',
                    'Expiry': ''
                }
                sl_no_counter += 1
        
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
    eway_match = re.search(r'e-Way Bill No\.\s*\n?([0-9\s]+)', text)
    if eway_match:
        doc_info['e-Way Bill'] = eway_match.group(1).strip()
    
    # Extract date
    date_match = re.search(r'dt\.\s*([0-9\-A-Za-z]+)', text)
    if date_match:
        doc_info['Date'] = date_match.group(1).strip()
    
    # Extract destination
    dest_match = re.search(r'Destination\s*\n?([^\n]+)', text)
    if dest_match:
        doc_info['Destination'] = dest_match.group(1).strip()
    
    return doc_info

def create_excel_file(buyer_info, supplier_info, products, doc_info):
    """Create Excel file in memory and return as bytes"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
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
        total_qty = sum(p.get('Quantity', 0) for p in products if isinstance(p.get('Quantity'), (int, float)))
        total_amt = sum(p.get('Amount', 0) for p in products if isinstance(p.get('Amount'), (int, float)))
        
        summary_data = {
            'Total Items': [len(products)],
            'Total Quantity': [total_qty],
            'Total Amount': [total_amt],
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
    
    output.seek(0)
    return output

# Streamlit UI
def main():
    # Header
    st.title("📄 Delivery Challan to Excel Converter")
    st.markdown("Upload your delivery challan PDF and download the extracted data as Excel")
    
    # Create two columns for layout
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### 📤 Upload PDF")
        uploaded_file = st.file_uploader(
            "Choose a PDF file",
            type=['pdf'],
            help="Upload your delivery challan PDF file"
        )
    
    with col2:
        st.markdown("### ℹ️ Information")
        st.info("""
        **This tool extracts:**
        - Buyer Details
        - Supplier Details
        - Product Information
        - Batch Numbers
        - Expiry Dates
        """)
    
    # Process button
    if uploaded_file is not None:
        st.markdown("---")
        
        # Show file details
        file_details = {
            "Filename": uploaded_file.name,
            "File Size": f"{uploaded_file.size / 1024:.2f} KB"
        }
        st.write("**Uploaded File:**", file_details)
        
        # Process button
        if st.button("🔄 Extract Data", type="primary", use_container_width=True):
            with st.spinner("Processing PDF... Please wait"):
                
                # Extract text from PDF
                text = extract_text_from_pdf(uploaded_file)
                
                if text:
                    # Extract all information
                    buyer_info = extract_buyer_details(text)
                    supplier_info = extract_supplier_details(text)
                    products = extract_products(text)
                    doc_info = extract_document_info(text)
                    
                    # Display extraction summary
                    st.success("✅ Data extracted successfully!")
                    
                    # Show summary in columns
                    st.markdown("### 📊 Extraction Summary")
                    sum_col1, sum_col2, sum_col3 = st.columns(3)
                    
                    with sum_col1:
                        st.metric("Total Products", len(products))
                    
                    with sum_col2:
                        total_amount = sum(p.get('Amount', 0) for p in products if isinstance(p.get('Amount'), (int, float)) and p.get('Amount'))
                        st.metric("Total Amount", f"₹{total_amount:,.2f}")
                    
                    with sum_col3:
                        st.metric("DC Number", doc_info.get('DC Number', 'N/A'))
                    
                    # Display details in expanders
                    st.markdown("---")
                    
                    with st.expander("👤 Buyer Details", expanded=False):
                        st.json(buyer_info)
                    
                    with st.expander("🏢 Supplier Details", expanded=False):
                        st.json(supplier_info)
                    
                    with st.expander("📦 Product Details Preview", expanded=False):
                        if products:
                            st.dataframe(pd.DataFrame(products).head(10), use_container_width=True)
                            if len(products) > 10:
                                st.info(f"Showing first 10 of {len(products)} products")
                    
                    # Create Excel file
                    excel_file = create_excel_file(buyer_info, supplier_info, products, doc_info)
                    
                    # Download button
                    st.markdown("---")
                    st.markdown("### 💾 Download Excel File")
                    
                    # Generate filename with timestamp
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    dc_number = doc_info.get('DC Number', 'DC').replace('/', '_')
                    filename = f"Delivery_Challan_{dc_number}_{timestamp}.xlsx"
                    
                    st.download_button(
                        label="⬇️ Download Excel File",
                        data=excel_file,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                    
                    st.success(f"✅ Click the button above to download '{filename}'")
                else:
                    st.error("❌ Failed to extract text from PDF. Please check if the file is valid.")
    
    else:
        # Instructions when no file is uploaded
        st.markdown("---")
        st.markdown("""
        ### 📋 How to use:
        1. Click on **'Browse files'** button above
        2. Select your delivery challan PDF file
        3. Click **'Extract Data'** button
        4. Review the extracted information
        5. Download the Excel file
        
        ### 📝 Supported Format:
        - Delivery challan PDF files
        - Contains buyer, supplier, and product information
        - Includes batch numbers and expiry dates
        """)
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: gray;'>Made with ❤️ using Streamlit | © 2024</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()