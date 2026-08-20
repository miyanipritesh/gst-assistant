import streamlit as st
import pandas as pd
import json
from datetime import datetime

st.set_page_config(page_title="GST Auto-Filer Pro", layout="wide", page_icon="🧾")
st.title("🧾 GST Auto-Filer & Reconciliation Pro")

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

    # 2. B2B & B2C Extraction for Reconciliation Check
    b2b_sum = 0.0
    if 'B2B' in excel.sheet_names:
        df_b2b = pd.read_excel(uploaded_file, sheet_name='B2B', header=None)
        for r in df_b2b.values[4:]:
            if pd.notna(r[0]) and pd.notna(r[11]):
                b2b_sum += float(r[11])
                
    b2cs_sum = 0.0
    b2cs_records = []
    if 'B2C Small' in excel.sheet_names:
        df_b2cs = pd.read_excel(uploaded_file, sheet_name='B2C Small', header=None)
        for r in df_b2cs.values[4:]:
            if pd.notna(r[1]) and pd.notna(r[4]) and float(r[4]) > 0:
                pos = str(r[1])
                rate = float(r[3]) if pd.notna(r[3]) else 0.05
                taxable = float(r[4])
                b2cs_sum += taxable
                is_intra = pos.startswith('24') # Gujarat State Code
                b2cs_records.append({
                    "State": pos, "Rate": f"{rate*100:.0f}%", "Taxable": taxable,
                    "IGST": 0.0 if is_intra else taxable * rate,
                    "CGST": (taxable * rate / 2) if is_intra else 0.0,
                    "SGST": (taxable * rate / 2) if is_intra else 0.0
                })

    # 3. Mismatch Reconciliation Alert
    diff = round(abs((b2b_sum + b2cs_sum) - taxable_total), 2)
    if diff == 0:
        st.success("✅ **Data Verified:** B2B + B2C Sales Summary HSN Total se perfectly match ho rahi hai.")
    else:
        st.warning(f"⚠️ **Mismatch Detected:** B2B+B2C Total (₹{b2b_sum + b2cs_sum:,.2f}) aur HSN Total (₹{taxable_total:,.2f}) mein ₹{diff} ka difference hai.")

    # 4. Summary Metrics
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Gross Turnover", f"₹{gross_total:,.2f}")
    kpi2.metric("Total Taxable Sales", f"₹{taxable_total:,.2f}")
    kpi3.metric("Total Output GST", f"₹{total_output_tax:,.2f}")
    kpi4.metric("Amazon 1% TCS Value", f"₹{(taxable_total * 0.01):,.2f}")

    st.divider()

    # 5. Due Date Banner & Cash Calculator
    st.subheader("📅 Filing Due Dates & Net Cash Tax")
    st.info("📌 **GSTR-1 Due Date:** 11th of next month | **GSTR-3B Due Date:** 20th of next month")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        itc_igst = st.number_input("Purchase ITC (IGST)", min_value=0.0, value=0.0, step=50.0)
    with c2:
        itc_cgst = st.number_input("Purchase ITC (CGST)", min_value=0.0, value=0.0, step=50.0)
    with c3:
        itc_sgst = st.number_input("Purchase ITC (SGST)", min_value=0.0, value=0.0, step=50.0)
    with c4:
        tcs_credit = st.number_input("TCS Credit to Claim", min_value=0.0, value=round(taxable_total * 0.01, 2), step=10.0)

    net_igst_cash = max(0.0, igst_total - itc_igst)
    net_cgst_cash = max(0.0, cgst_total - itc_cgst)
    net_sgst_cash = max(0.0, sgst_total - itc_sgst)
    net_total_cash = max(0.0, (net_igst_cash + net_cgst_cash + net_sgst_cash) - tcs_credit)

    st.success(f"💰 **Final Challan Amount to Pay:** ₹{round(net_total_cash):,}  (IGST: ₹{net_igst_cash:,.2f} | CGST: ₹{net_cgst_cash:,.2f} | SGST: ₹{net_sgst_cash:,.2f})")
