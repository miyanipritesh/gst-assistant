

import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="AI GST Filing Assistant & Audit Validator", layout="wide")
st.title("🛡️ Monthly GST Auto-Filing & 100% Audit Validator")

# Helper function for safe numerical conversions
def safe_float(val, default=0.0):
    try:
        if pd.isna(val) or str(val).strip() == '':
            return default
        return float(val)
    except (ValueError, TypeError):
        return default

# GSTIN 15-character structure validation
def is_valid_gstin(gstin):
    pattern = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"
    return bool(re.match(pattern, str(gstin).strip()))

uploaded_file = st.file_uploader("Upload Monthly GST Excel File", type=["xlsx", "xls"])

if uploaded_file:
    try:
        excel = pd.ExcelFile(uploaded_file)
    except Exception as e:
        st.error(f"❌ File read error: Kripya valid Excel format upload karein. ({e})")
        st.stop()
    
    # ---------------- 1. SUPPLIER STATE DETECTION ----------------
    supplier_state_code = "24" # Default Gujarat
    if 'GSTIN' in excel.sheet_names:
        df_gstin = pd.read_excel(uploaded_file, sheet_name='GSTIN', header=None)
        for val in df_gstin.values.flatten():
            val_str = str(val).strip().upper()
            if len(val_str) == 15 and val_str[:2].isdigit():
                supplier_state_code = val_str[:2]
                break

    # ---------------- 2. HSN SUMMARY EXTRACTION ----------------
    hsn_rows = []
    taxable_hsn_sum = 0.0
    igst_hsn_sum = 0.0
    cgst_hsn_sum = 0.0
    sgst_hsn_sum = 0.0
    gross_hsn_sum = 0.0
    
    if 'HSN Summary' in excel.sheet_names:
        df_hsn = pd.read_excel(uploaded_file, sheet_name='HSN Summary', header=None)
        raw_hsn_values = df_hsn.values[4:] if len(df_hsn.values) > 4 else []
        
        for r in raw_hsn_values:
            if len(r) > 9 and pd.notna(r[0]) and str(r[0]).strip() != '':
                hsn_code = str(r[0]).strip()
                uqc = str(r[2]).strip() if pd.notna(r[2]) else "PCS"
                qty = safe_float(r[3])
                rate = safe_float(r[4])
                gross = safe_float(r[5])
                taxable = safe_float(r[6])
                igst = safe_float(r[7])
                cgst = safe_float(r[8])
                sgst = safe_float(r[9])
                
                taxable_hsn_sum += taxable
                igst_hsn_sum += igst
                cgst_hsn_sum += cgst
                sgst_hsn_sum += sgst
                gross_hsn_sum += gross
                
                hsn_rows.append({
                    "HSN Code": hsn_code, "UQC": uqc, "Qty": qty,
                    "GST Rate": f"{rate*100:.0f}%",
                    "Taxable (₹)": taxable, "IGST (₹)": igst,
                    "CGST (₹)": cgst, "SGST (₹)": sgst, "Gross Total (₹)": gross
                })

    # ---------------- 3. B2B EXTRACTION & VALIDATION ----------------
    b2b_list = []
    b2b_taxable_sum = 0.0
    b2b_errors = []
    
    if 'B2B' in excel.sheet_names:
        df_b2b = pd.read_excel(uploaded_file, sheet_name='B2B', header=None)
        raw_b2b_values = df_b2b.values[4:] if len(df_b2b.values) > 4 else []
        
        for r in raw_b2b_values:
            if len(r) > 11 and pd.notna(r[0]) and str(r[0]).strip() != '':
                buyer_gstin = str(r[0]).strip().upper()
                inv_no = str(r[2]).strip()
                inv_date = str(r[3]).strip()
                pos = str(r[5]).strip()
                inv_val = safe_float(r[4])
                taxable_val = safe_float(r[11])
                rate = safe_float(r[10])
                
                # Check GSTIN Format
                if not is_valid_gstin(buyer_gstin):
                    b2b_errors.append(f"Invoice {inv_no}: Invalid GSTIN format '{buyer_gstin}'")
                
                b2b_taxable_sum += taxable_val
                b2b_list.append({
                    "Buyer GSTIN": buyer_gstin, "Invoice No": inv_no, "Date": inv_date,
                    "Place of Supply": pos, "Rate": f"{rate*100:.0f}%",
                    "Taxable Value (₹)": taxable_val, "Invoice Value (₹)": inv_val
                })

    # ---------------- 4. B2C SMALL EXTRACTION & VALIDATION ----------------
    b2cs_list = []
    b2cs_taxable_sum = 0.0
    
    if 'B2C Small' in excel.sheet_names:
        df_b2cs = pd.read_excel(uploaded_file, sheet_name='B2C Small', header=None)
        raw_b2cs_values = df_b2cs.values[4:] if len(df_b2cs.values) > 4 else []
        
        for r in raw_b2cs_values:
            if len(r) > 4 and pd.notna(r[1]) and pd.notna(r[4]):
                pos = str(r[1]).strip()
                rate = safe_float(r[3], 0.05)
                taxable_val = safe_float(r[4])
                
                if taxable_val > 0:
                    is_intra = pos.startswith(supplier_state_code)
                    igst = 0.0 if is_intra else round(taxable_val * rate, 2)
                    cgst = round((taxable_val * rate) / 2, 2) if is_intra else 0.0
                    sgst = round((taxable_val * rate) / 2, 2) if is_intra else 0.0
                    
                    b2cs_taxable_sum += taxable_val
                    b2cs_list.append({
                        "Place of Supply": pos, "Rate": f"{rate*100:.0f}%",
                        "Taxable Value (₹)": taxable_val,
                        "IGST (₹)": igst, "CGST (₹)": cgst, "SGST (₹)": sgst
                    })

    # ---------------- 5. COMPREHENSIVE AUDIT & NOTICE VALIDATION ----------------
    st.subheader("🔍 Automated GST Audit & Notice Prevention Check")
    audit_passed = True
    
    # Check 1: HSN vs Outward Sales Reconciliation
    diff = round(abs((b2b_taxable_sum + b2cs_taxable_sum) - taxable_hsn_sum), 2)
    if diff == 0.0:
        st.success("✅ **100% Match:** B2B + B2C Total Sales HSN Table se perfectly match ho rahi hai. (No Mismatch Risk)")
    else:
        audit_passed = False
        st.error(f"⚠️ **Mismatch Alert (Risk of ASMT-10 Notice):** B2B+B2C Total (₹{b2b_taxable_sum + b2cs_taxable_sum:,.2f}) aur HSN Total (₹{taxable_hsn_sum:,.2f}) mein ₹{diff} ka difference hai!")

    # Check 2: B2B Invalids
    if b2b_errors:
        audit_passed = False
        for err in b2b_errors:
            st.error(f"⚠️ **B2B GSTIN Error:** {err} - GST portal par upload karte waqt error aayega.")

    if audit_passed:
        st.caption("✨ Sabhi core validations pass ho chuke hain. Aap safely yehi data GSTR-1 aur GSTR-3B mein file kar sakte hain.")

    st.divider()

    # ---------------- 6. SUMMARY METRICS ----------------
    total_tax = igst_hsn_sum + cgst_hsn_sum + sgst_hsn_sum
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Taxable Turnover", f"₹{taxable_hsn_sum:,.2f}")
    col2.metric("Total Output GST", f"₹{total_tax:,.2f}")
    col3.metric("IGST", f"₹{igst_hsn_sum:,.2f}")
    col4.metric("CGST + SGST", f"₹{(cgst_hsn_sum + sgst_hsn_sum):,.2f}")
    
    st.subheader("GSTR-3B Tax Liability Breakdown")
    st.write(f"👉 **Gross Output Tax in Table 3.1:** ₹{round(total_tax):,}")

    # ---------------- 7. ITC / TCS DEDUCTION CALCULATOR ----------------
    st.divider()
    st.subheader("💳 Input Tax Credit (ITC) Deduction")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        itc_igst = c1.number_input("Purchase IGST", min_value=0.0, value=0.0, step=50.0)
    with c2:
        itc_cgst = c2.number_input("Purchase CGST", min_value=0.0, value=0.0, step=50.0)
    with c3:
        itc_sgst = c3.number_input("Purchase SGST", min_value=0.0, value=0.0, step=50.0)
    with c4:
        tcs_val = c4.number_input("E-Commerce TCS Credit", min_value=0.0, value=0.0, step=10.0)

    net_igst_cash = max(0.0, igst_hsn_sum - itc_igst)
    net_cgst_cash = max(0.0, cgst_hsn_sum - itc_cgst)
    net_sgst_cash = max(0.0, sgst_hsn_sum - itc_sgst)
    net_cash = max(0.0, (net_igst_cash + net_cgst_cash + net_sgst_cash) - tcs_val)

    st.success(f"👉 **Final Net Cash Payable (Bank Challan):** ₹{round(net_cash):,} *(IGST: ₹{net_igst_cash:,.2f} | CGST: ₹{net_cgst_cash:,.2f} | SGST: ₹{net_sgst_cash:,.2f})*")

    st.divider()

    # ---------------- 8. DATA BREAKDOWN TABLES ----------------
    st.subheader("📋 Sheet Wise Complete Breakdown")

    # 1. HSN Table
    st.write("**1. HSN Summary Data (GSTR-1 Table 12)**")
    if hsn_rows:
        st.dataframe(pd.DataFrame(hsn_rows), use_container_width=True)
    else:
        st.info("HSN summary data nahi mila.")

    # 2. B2B Table
    st.write("**2. B2B Registered Sales (GSTR-1 Table 4A)**")
    if b2b_list:
        st.dataframe(pd.DataFrame(b2b_list), use_container_width=True)
    else:
        st.info("B2B sheet mein koi data nahi mila.")

    # 3. B2C Table
    st.write("**3. B2C State-Wise Sales (GSTR-1 Table 7)**")
    if b2cs_list:
        st.dataframe(pd.DataFrame(b2cs_list), use_container_width=True)
    else:
        st.info("B2C Small sheet mein koi data nahi mila.")
