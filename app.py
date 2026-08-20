import streamlit as st
import pandas as pd

st.set_page_config(page_title="AI GST Filing Assistant", layout="wide")
st.title("📊 Monthly GST Auto-Filing Assistant & Breakdown")

uploaded_file = st.file_uploader("Upload Monthly GST Excel File", type=["xlsx", "xls"])

if uploaded_file:
    excel = pd.ExcelFile(uploaded_file)
    
    # 1. Extract HSN Summary & Tax Computation
    hsn_records = []
    taxable_total = 0.0
    igst_total = 0.0
    cgst_total = 0.0
    sgst_total = 0.0
    gross_total = 0.0
    
    if 'HSN Summary' in excel.sheet_names:
        df_hsn = pd.read_excel(uploaded_file, sheet_name='HSN Summary', header=None)
        hsn_rows = df_hsn.values[4:]
        
        for r in hsn_rows:
            if pd.notna(r[0]) and str(r[0]).strip() != '':
                hsn_code = str(r[0])
                uqc = str(r[2]) if pd.notna(r[2]) else "PCS"
                qty = float(r[3]) if pd.notna(r[3]) else 0
                rate = f"{float(r[4])*100:.0f}%" if pd.notna(r[4]) else "0%"
                gross = float(r[5]) if pd.notna(r[5]) else 0.0
                taxable = float(r[6]) if pd.notna(r[6]) else 0.0
                igst = float(r[7]) if pd.notna(r[7]) else 0.0
                cgst = float(r[8]) if pd.notna(r[8]) else 0.0
                sgst = float(r[9]) if pd.notna(r[9]) else 0.0
                
                taxable_total += taxable
                igst_total += igst
                cgst_total += cgst
                sgst_total += sgst
                gross_total += gross
                
                hsn_records.append({
                    "HSN Code": hsn_code,
                    "UQC": uqc,
                    "Total Quantity": qty,
                    "GST Rate": rate,
                    "Taxable Value (₹)": f"₹{taxable:,.2f}",
                    "IGST (₹)": f"₹{igst:,.2f}",
                    "CGST (₹)": f"₹{cgst:,.2f}",
                    "SGST (₹)": f"₹{sgst:,.2f}",
                    "Gross Total (₹)": f"₹{gross:,.2f}"
                })

    total_output_tax = igst_total + cgst_total + sgst_total

    # 2. Top Summary KPI Cards
    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("Gross Sales", f"₹{gross_total:,.2f}")
    col2.metric("Taxable Turnover", f"₹{taxable_total:,.2f}")
    col3.metric("Total Output GST", f"₹{total_output_tax:,.2f}")
    col4.metric("IGST", f"₹{igst_total:,.2f}")
    col5.metric("CGST + SGST", f"₹{(cgst_total + sgst_total):,.2f}")
    
    st.divider()

    # 3. GSTR-3B Cash Calculator with ITC / TCS
    st.subheader("💳 GSTR-3B Net Cash Liability Calculator")
    c_itc1, c_itc2, c_itc3, c_tcs = st.columns(4)
    with c_itc1:
        itc_igst = st.number_input("Purchase ITC (IGST)", min_value=0.0, value=0.0, step=50.0)
    with c_itc2:
        itc_cgst = st.number_input("Purchase ITC (CGST)", min_value=0.0, value=0.0, step=50.0)
    with c_itc3:
        itc_sgst = st.number_input("Purchase ITC (SGST)", min_value=0.0, value=0.0, step=50.0)
    with c_tcs:
        tcs_credit = st.number_input("E-Commerce TCS Credit", min_value=0.0, value=0.0, step=50.0)

    net_igst_cash = max(0.0, igst_total - itc_igst)
    net_cgst_cash = max(0.0, cgst_total - itc_cgst)
    net_sgst_cash = max(0.0, sgst_total - itc_sgst)
    net_total_cash = max(0.0, (net_igst_cash + net_cgst_cash + net_sgst_cash) - tcs_credit)

    st.success(f"👉 **Total Cash Tax to Pay in GSTR-3B:** ₹{round(net_total_cash):,}  *(IGST: ₹{net_igst_cash:,.2f} | CGST: ₹{net_cgst_cash:,.2f} | SGST: ₹{net_sgst_cash:,.2f})*")

    st.divider()

    # 4. Detailed Data Breakdown Tabs
    tab1, tab2, tab3 = st.tabs(["📦 HSN Product Breakdown", "🏬 B2B Invoices (Table 4A)", "🛒 B2C Small Sales (Table 7)"])

    with tab1:
        st.subheader("HSN Wise Summary (GSTR-1 Table 12)")
        if hsn_records:
            st.dataframe(pd.DataFrame(hsn_records), use_container_width=True)
        else:
            st.info("Koi HSN data nahi mila.")

    with tab2:
        st.subheader("B2B Registered Invoices (GSTR-1 Table 4A)")
        b2b_records = []
        if 'B2B' in excel.sheet_names:
            df_b2b = pd.read_excel(uploaded_file, sheet_name='B2B', header=None)
            for r in df_b2b.values[4:]:
                if pd.notna(r[0]) and str(r[0]).strip() != '':
                    b2b_records.append({
                        "Recipient GSTIN": str(r[0]),
                        "Invoice No": str(r[2]),
                        "Invoice Date": str(r[3]),
                        "Place of Supply": str(r[5]),
                        "GST Rate": f"{float(r[10])*100:.0f}%" if pd.notna(r[10]) else "0%",
                        "Taxable Value (₹)": f"₹{float(r[11]):,.2f}" if pd.notna(r[11]) else "₹0.00",
                        "Invoice Value (₹)": f"₹{float(r[4]):,.2f}" if pd.notna(r[4]) else "₹0.00"
                    })
        if b2b_records:
            st.dataframe(pd.DataFrame(b2b_records), use_container_width=True)
        else:
            st.info("Koi B2B invoices nahi hain.")

    with tab3:
        st.subheader("B2C State-Wise Small Sales (GSTR-1 Table 7)")
        b2cs_records = []
        if 'B2C Small' in excel.sheet_names:
            df_b2cs = pd.read_excel(uploaded_file, sheet_name='B2C Small', header=None)
            for r in df_b2cs.values[4:]:
                if pd.notna(r[1]) and pd.notna(r[4]) and float(r[4]) > 0:
                    pos = str(r[1])
                    rate = float(r[3]) if pd.notna(r[3]) else 0.05
                    taxable = float(r[4])
                    
                    # Inter vs Intra tax calculation
                    is_intra = pos.startswith('24') # Gujarat
                    igst = 0.0 if is_intra else taxable * rate
                    cgst = (taxable * rate / 2) if is_intra else 0.0
                    sgst = (taxable * rate / 2) if is_intra else 0.0

                    b2cs_records.append({
                        "Place of Supply": pos,
                        "Rate": f"{rate*100:.0f}%",
                        "Taxable Value (₹)": f"₹{taxable:,.2f}",
                        "IGST (₹)": f"₹{igst:,.2f}",
                        "CGST (₹)": f"₹{cgst:,.2f}",
                        "SGST (₹)": f"₹{sgst:,.2f}"
                    })
        if b2cs_records:
            st.dataframe(pd.DataFrame(b2cs_records), use_container_width=True)
        else:
            st.info("Koi B2C Small data nahi mila.")
