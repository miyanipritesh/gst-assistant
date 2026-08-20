import streamlit as st
import pandas as pd
import json

st.set_page_config(page_title="Advance GST Pro Assistant", layout="wide", page_icon="📈")
st.title("🚀 Advance GST Filing Assistant & Analytics Pro")

uploaded_file = st.file_uploader("Upload Monthly GST Excel File", type=["xlsx", "xls"])

if uploaded_file:
    excel = pd.ExcelFile(uploaded_file)
    
    # 1. HSN Data Extraction
    hsn_records = []
    taxable_total = 0.0
    igst_total = 0.0
    cgst_total = 0.0
    sgst_total = 0.0
    gross_total = 0.0
    
    if 'HSN Summary' in excel.sheet_names:
        df_hsn = pd.read_excel(uploaded_file, sheet_name='HSN Summary', header=None)
        for r in df_hsn.values[4:]:
            if pd.notna(r[0]) and str(r[0]).strip() != '':
                hsn_code = str(r[0])
                uqc = str(r[2]) if pd.notna(r[2]) else "PCS"
                qty = float(r[3]) if pd.notna(r[3]) else 0
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
                    "HSN Code": hsn_code, "UQC": uqc, "Qty": qty,
                    "Taxable (₹)": taxable, "IGST (₹)": igst,
                    "CGST (₹)": cgst, "SGST (₹)": sgst, "Total (₹)": gross
                })

    total_output_tax = igst_total + cgst_total + sgst_total

    # 2. Top Summary KPI Cards
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("Gross Sales", f"₹{gross_total:,.2f}")
    kpi2.metric("Taxable Sales", f"₹{taxable_total:,.2f}")
    kpi3.metric("Total GST Liability", f"₹{total_output_tax:,.2f}")
    kpi4.metric("IGST", f"₹{igst_total:,.2f}")
    kpi5.metric("CGST + SGST", f"₹{(cgst_total + sgst_total):,.2f}")
    
    st.divider()

    # 3. ITC, TCS & Cash Tax Computation
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

    st.success(f"👉 **Total Cash Challan Amount for GSTR-3B:** ₹{round(net_total_cash):,}  (IGST: ₹{net_igst_cash:,.2f} | CGST: ₹{net_cgst_cash:,.2f} | SGST: ₹{net_sgst_cash:,.2f})")

    st.divider()

    # 4. Feature Tabs
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 State Sales Analytics", 
        "📝 GSTR-3B Form Copy-Paste", 
        "💾 Portal JSON Export",
        "📦 HSN Product Breakdown", 
        "🏬 B2B & B2C Invoices"
    ])

    # B2C Data Preparation
    b2cs_records = []
    if 'B2C Small' in excel.sheet_names:
        df_b2cs = pd.read_excel(uploaded_file, sheet_name='B2C Small', header=None)
        for r in df_b2cs.values[4:]:
            if pd.notna(r[1]) and pd.notna(r[4]) and float(r[4]) > 0:
                pos = str(r[1])
                rate = float(r[3]) if pd.notna(r[3]) else 0.05
                taxable = float(r[4])
                is_intra = pos.startswith('24') # Gujarat
                b2cs_records.append({
                    "State": pos,
                    "Taxable Value": taxable,
                    "IGST": 0.0 if is_intra else taxable * rate,
                    "CGST": (taxable * rate / 2) if is_intra else 0.0,
                    "SGST": (taxable * rate / 2) if is_intra else 0.0
                })
    df_b2c_clean = pd.DataFrame(b2cs_records)

    with tab1:
        st.subheader("📍 State-Wise Sales Breakdown & Trend")
        if not df_b2c_clean.empty:
            chart_data = df_b2c_clean.set_index('State')['Taxable Value']
            st.bar_chart(chart_data)
            st.dataframe(df_b2c_clean, use_container_width=True)
        else:
            st.info("State sales data not available.")

    with tab2:
        st.subheader("📋 Direct GSTR-3B Portal Form Mapping")
        st.caption("GST portal khol kar exact Table 3.1 mein yeh values enter karein:")
        
        gstr3b_mapping = [
            {"Table Number": "3.1 (a) Outward Taxable Supplies", "Taxable (₹)": f"₹{taxable_total:,.2f}", "IGST (₹)": f"₹{igst_total:,.2f}", "CGST (₹)": f"₹{cgst_total:,.2f}", "SGST (₹)": f"₹{sgst_total:,.2f}"},
            {"Table Number": "4 (A)(5) All Other Eligible ITC", "Taxable (₹)": "-", "IGST (₹)": f"₹{itc_igst:,.2f}", "CGST (₹)": f"₹{itc_cgst:,.2f}", "SGST (₹)": f"₹{itc_sgst:,.2f}"},
            {"Table Number": "6.1 Net Payment in Cash", "Taxable (₹)": "-", "IGST (₹)": f"₹{net_igst_cash:,.2f}", "CGST (₹)": f"₹{net_cgst_cash:,.2f}", "SGST (₹)": f"₹{net_sgst_cash:,.2f}"}
        ]
        st.table(pd.DataFrame(gstr3b_mapping))

    with tab3:
        st.subheader("⚡ Download Portal Uploadable JSON")
        st.caption("Yeh JSON file GSTR-1 offline tool mein directly import ki ja sakti hai.")
        
        portal_json = {
            "version": "GSTR1_v2.0",
            "cur_gt": round(gross_total, 2),
            "cur_txval": round(taxable_total, 2),
            "tax_details": {
                "igst": round(igst_total, 2),
                "cgst": round(cgst_total, 2),
                "sgst": round(sgst_total, 2),
                "total_tax": round(total_output_tax, 2)
            },
            "b2cs": b2cs_records,
            "hsn": hsn_records
        }
        json_str = json.dumps(portal_json, indent=4)
        st.download_button(
            label="📥 Download GSTR-1 JSON File",
            data=json_str,
            file_name="GSTR1_Monthly_Data.json",
            mime="application/json"
        )

    with tab4:
        st.subheader("HSN Summary (Table 12)")
        if hsn_records:
            st.dataframe(pd.DataFrame(hsn_records), use_container_width=True)

    with tab5:
        st.subheader("B2B Registered Invoices (Table 4A)")
        b2b_records = []
        if 'B2B' in excel.sheet_names:
            df_b2b = pd.read_excel(uploaded_file, sheet_name='B2B', header=None)
            for r in df_b2b.values[4:]:
                if pd.notna(r[0]) and str(r[0]).strip() != '':
                    b2b_records.append({
                        "GSTIN": str(r[0]), "Invoice No": str(r[2]), "Date": str(r[3]),
                        "Place of Supply": str(r[5]), "Taxable (₹)": float(r[11]), "Invoice Value (₹)": float(r[4])
                    })
        if b2b_records:
            st.dataframe(pd.DataFrame(b2b_records), use_container_width=True)
        else:
            st.info("No B2B invoices found.")
