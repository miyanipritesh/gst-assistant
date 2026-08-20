import streamlit as st
import pandas as pd

st.set_page_config(page_title="AI GST Filing Assistant", layout="wide")
st.title("📊 Monthly GST Auto-Filing Assistant")

uploaded_file = st.file_uploader("Upload Monthly GST Excel File", type=["xlsx", "xls"])

if uploaded_file:
    excel = pd.ExcelFile(uploaded_file)
    
    # Process HSN Summary
    if 'HSN Summary' in excel.sheet_names:
        df_hsn = pd.read_excel(uploaded_file, sheet_name='HSN Summary', header=None)
        hsn_rows = df_hsn.values[4:]
        
        taxable = sum([float(r[6]) for r in hsn_rows if pd.notna(r[0])])
        igst = sum([float(r[7]) for r in hsn_rows if pd.notna(r[0])])
        cgst = sum([float(r[8]) for r in hsn_rows if pd.notna(r[0])])
        sgst = sum([float(r[9]) for r in hsn_rows if pd.notna(r[0])])
        total_tax = igst + cgst + sgst
        
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Taxable Turnover", f"₹{taxable:,.2f}")
        col2.metric("Total Output GST", f"₹{total_tax:,.2f}")
        col3.metric("IGST", f"₹{igst:,.2f}")
        col4.metric("CGST + SGST", f"₹{(cgst + sgst):,.2f}")
        
        st.subheader("GSTR-3B Tax Liability Breakdown")
        st.write(f"👉 **Cash Payable in 3B (without ITC deduction):** ₹{round(total_tax)}")
