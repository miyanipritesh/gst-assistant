
import streamlit as st
import pandas as pd

st.set_page_config(page_title="AI GST Filing Assistant", layout="wide")
st.title("📊 Monthly GST Auto-Filing Assistant")

uploaded_file = st.file_uploader("Upload Monthly GST Excel File", type=["xlsx", "xls"])

if uploaded_file:
    excel = pd.ExcelFile(uploaded_file)
    
    # --- AAPKA ORIGINAL HSN LOGIC ---
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
        st.write(f"👉 **Gross Output Tax in 3B:** ₹{round(total_tax)}")

        # --- NAYA FEATURE 1: ITC / TCS Cash Deduction ---
        st.divider()
        st.subheader("💳 Input Tax Credit (ITC) Deduction")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            itc_igst = c1.number_input("Purchase IGST", min_value=0.0, value=0.0)
        with c2:
            itc_cgst = c2.number_input("Purchase CGST", min_value=0.0, value=0.0)
        with c3:
            itc_sgst = c3.number_input("Purchase SGST", min_value=0.0, value=0.0)
        with c4:
            tcs_val = c4.number_input("TCS Credit", min_value=0.0, value=0.0)

        net_cash = max(0.0, (max(0.0, igst - itc_igst) + max(0.0, cgst - itc_cgst) + max(0.0, sgst - itc_sgst)) - tcs_val)
        st.success(f"👉 **Net Cash Payable in Bank/Challan:** ₹{round(net_cash)}")

        st.divider()

        # --- NAYA FEATURE 2: Excel Ka Complete Data Breakdown ---
        st.subheader("📋 Sheet Wise Complete Breakdown")

        # 1. HSN Breakdown Table
        st.write("**1. HSN Summary Data (GSTR-1 Table 12)**")
        hsn_list = []
        for r in hsn_rows:
            if pd.notna(r[0]) and str(r[0]).strip() != '':
                hsn_list.append({
                    "HSN Code": str(r[0]), "Qty": float(r[3]), "Taxable (₹)": float(r[6]),
                    "IGST (₹)": float(r[7]), "CGST (₹)": float(r[8]), "SGST (₹)": float(r[9])
                })
        if hsn_list:
            st.dataframe(pd.DataFrame(hsn_list), use_container_width=True)

        # 2. B2B Invoices Table
        st.write("**2. B2B Registered Sales (GSTR-1 Table 4A)**")
        b2b_list = []
        if 'B2B' in excel.sheet_names:
            df_b2b = pd.read_excel(uploaded_file, sheet_name='B2B', header=None)
            for r in df_b2b.values[4:]:
                if pd.notna(r[0]) and str(r[0]).strip() != '':
                    b2b_list.append({
                        "GSTIN": str(r[0]), "Invoice No": str(r[2]), "Date": str(r[3]),
                        "State (POS)": str(r[5]), "Taxable Value (₹)": float(r[11]), "Total (₹)": float(r[4])
                    })
        if b2b_list:
            st.dataframe(pd.DataFrame(b2b_list), use_container_width=True)
        else:
            st.info("B2B sheet mein koi data nahi hai.")

        # 3. B2C State-Wise Sales Table
        st.write("**3. B2C State-Wise Sales (GSTR-1 Table 7)**")
        b2cs_list = []
        if 'B2C Small' in excel.sheet_names:
            df_b2cs = pd.read_excel(uploaded_file, sheet_name='B2C Small', header=None)
            for r in df_b2cs.values[4:]:
                if pd.notna(r[1]) and pd.notna(r[4]) and float(r[4]) > 0:
                    b2cs_list.append({
                        "State (POS)": str(r[1]),
                        "Rate": f"{float(r[3])*100:.0f}%" if pd.notna(r[3]) else "5%",
                        "Taxable Value (₹)": float(r[4])
                    })
        if b2cs_list:
            st.dataframe(pd.DataFrame(b2cs_list), use_container_width=True)
        else:
            st.info("B2C Small sheet mein koi data nahi hai.")
