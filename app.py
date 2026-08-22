import streamlit as st
import pandas as pd
import json
import io
import re
import zipfile
from datetime import datetime

st.set_page_config(page_title="Multi-Platform GST Auto-Filer & Analytics Pro", layout="wide", page_icon="🛍️")
st.title("🛍️ Multi-Platform E-Commerce GST Auto-Filer & Analytics Pro")
st.caption("Amazon, Flipkart aur Meesho ki Excel (.xlsx) ya ZIP (.zip) files direct upload karein.")

def safe_float(val, default=0.0):
    try:
        if pd.isna(val) or str(val).strip() == '':
            return default
        return float(val)
    except (ValueError, TypeError):
        return default

def is_valid_gstin(gstin):
    pattern = r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"
    return bool(re.match(pattern, str(gstin).strip()))

# --- AUTO DETECTOR ---
def detect_ecommerce_platform(file_bytes, filename=""):
    fn_low = filename.lower()
    if 'meesho' in fn_low or 'tcs_sales' in fn_low or 'gst_2479906' in fn_low:
        return "Meesho", "🟪 Meesho"
    if 'gstr1-july' in fn_low or 'gstr-1' in fn_low or 'amazon' in fn_low:
        return "Amazon", "🟧 Amazon"
        
    try:
        excel = pd.ExcelFile(file_bytes)
        sheets = [s.strip().lower() for s in excel.sheet_names]
        
        if any('section 7(a)' in s or 'section 12 in gstr-1' in s or 'section 7(b)' in s or 'section 3 in gstr-8' in s for s in sheets):
            return "Flipkart", "🟦 Flipkart"
        if any('b2c small' in s or 'hsn summary' in s or 'b2cl cn' in s for s in sheets):
            return "Amazon", "🟧 Amazon"
        if any('identifier' in s or 'tcs_sales' in s for s in sheets):
            return "Meesho", "🟪 Meesho"
            
        return "Unknown", "⚪ Unknown Platform"
    except Exception:
        return "Unknown", "⚪ Unknown Platform"

# --- PARSER 1: AMAZON ---
def parse_amazon(file_bytes):
    excel = pd.ExcelFile(file_bytes)
    hsn_records, b2cs_records, b2b_records = [], [], []
    taxable_sum, igst_sum, cgst_sum, sgst_sum, gross_sum = 0.0, 0.0, 0.0, 0.0, 0.0
    supplier_gstin = "N/A"
    supplier_state = "24"

    if 'GSTIN' in excel.sheet_names:
        df_gstin = pd.read_excel(file_bytes, sheet_name='GSTIN', header=None)
        for val in df_gstin.values.flatten():
            val_str = str(val).strip().upper()
            if len(val_str) == 15 and val_str[:2].isdigit():
                supplier_gstin = val_str
                supplier_state = val_str[:2]
                break

    if 'HSN Summary' in excel.sheet_names:
        df_hsn = pd.read_excel(file_bytes, sheet_name='HSN Summary', header=None)
        for r in df_hsn.values[4:]:
            if len(r) > 9 and pd.notna(r[0]) and str(r[0]).strip() != '':
                qty = safe_float(r[3])
                gross = safe_float(r[5])
                taxable = safe_float(r[6])
                igst = safe_float(r[7])
                cgst = safe_float(r[8])
                sgst = safe_float(r[9])
                
                taxable_sum += taxable
                igst_sum += igst
                cgst_sum += cgst
                sgst_sum += sgst
                gross_sum += gross
                
                hsn_records.append({
                    "Platform": "Amazon", "HSN Code": str(r[0]).strip(), "UQC": str(r[2]).strip() if pd.notna(r[2]) else "PCS",
                    "Qty": qty, "GST Rate": f"{safe_float(r[4])*100:.0f}%", "Rate_Num": safe_float(r[4])*100, "Taxable (₹)": taxable,
                    "IGST (₹)": igst, "CGST (₹)": cgst, "SGST (₹)": sgst, "Gross Total (₹)": gross
                })

    if 'B2C Small' in excel.sheet_names:
        df_b2cs = pd.read_excel(file_bytes, sheet_name='B2C Small', header=None)
        for r in df_b2cs.values[4:]:
            if len(r) > 4 and pd.notna(r[1]) and safe_float(r[4]) > 0:
                pos = str(r[1]).strip()
                taxable = safe_float(r[4])
                rate = safe_float(r[3], 0.05)
                is_intra = pos.startswith(supplier_state)
                igst = 0.0 if is_intra else round(taxable * rate, 2)
                cgst = round((taxable * rate) / 2, 2) if is_intra else 0.0
                sgst = round((taxable * rate) / 2, 2) if is_intra else 0.0
                b2cs_records.append({
                    "Platform": "Amazon", "Place of Supply": pos, "Rate": f"{rate*100:.0f}%", "Rate_Num": rate*100,
                    "Taxable Value (₹)": taxable, "IGST (₹)": igst, "CGST (₹)": cgst, "SGST (₹)": sgst,
                    "Gross Value (₹)": round(taxable + igst + cgst + sgst, 2)
                })

    if 'B2B' in excel.sheet_names:
        df_b2b = pd.read_excel(file_bytes, sheet_name='B2B', header=None)
        for r in df_b2b.values[4:]:
            if len(r) > 11 and pd.notna(r[0]) and str(r[0]).strip() != '':
                b2b_records.append({
                    "Platform": "Amazon", "Buyer GSTIN": str(r[0]).strip().upper(), "Invoice No": str(r[2]).strip(),
                    "Date": str(r[3]).strip(), "Place of Supply": str(r[5]).strip(), "Rate": f"{safe_float(r[10])*100:.0f}%", "Rate_Num": safe_float(r[10])*100,
                    "Taxable Value (₹)": safe_float(r[11]), "Gross / Invoice Value (₹)": safe_float(r[4])
                })

    return {
        "platform": "Amazon", "supplier_gstin": supplier_gstin, "supplier_state": supplier_state,
        "gross": gross_sum, "taxable": taxable_sum, "returns_gross": 0.0, "returns_taxable": 0.0,
        "igst": igst_sum, "cgst": cgst_sum, "sgst": sgst_sum,
        "total_tax": igst_sum + cgst_sum + sgst_sum,
        "tcs": round(taxable_sum * 0.005, 2),
        "hsn": hsn_records, "b2cs": b2cs_records, "b2b": b2b_records
    }

# --- PARSER 2: FLIPKART ---
def parse_flipkart(file_bytes):
    excel = pd.ExcelFile(file_bytes)
    hsn_records, b2cs_records, b2b_records = [], [], []
    taxable_sum, igst_sum, cgst_sum, sgst_sum, gross_sum = 0.0, 0.0, 0.0, 0.0, 0.0
    supplier_gstin = "24ECEPM6676L1Z0"
    supplier_state = "24"
    returns_taxable = 0.0

    if 'Section 12 in GSTR-1' in excel.sheet_names:
        df_hsn = pd.read_excel(file_bytes, sheet_name='Section 12 in GSTR-1')
        if not df_hsn.empty and 'GSTIN' in df_hsn.columns and pd.notna(df_hsn['GSTIN'].iloc[0]):
            supplier_gstin = str(df_hsn['GSTIN'].iloc[0]).strip().upper()
            supplier_state = supplier_gstin[:2]

        for _, r in df_hsn.iterrows():
            qty = safe_float(r.get('Total Quantity in Nos.', 0))
            gross = safe_float(r.get('Total\n Value Rs.', 0))
            taxable = safe_float(r.get('Total Taxable Value Rs.', 0))
            igst = safe_float(r.get('IGST Amount Rs.', 0))
            cgst = safe_float(r.get('CGST Amount Rs.', 0))
            sgst = safe_float(r.get('SGST Amount Rs.', 0))
            
            taxable_sum += taxable
            igst_sum += igst
            cgst_sum += cgst
            sgst_sum += sgst
            gross_sum += gross
            
            hsn_records.append({
                "Platform": "Flipkart", "HSN Code": str(r.get('HSN Number', '')).strip(), "UQC": "NOS",
                "Qty": qty, "GST Rate": "5%", "Rate_Num": 5.0, "Taxable (₹)": taxable,
                "IGST (₹)": igst, "CGST (₹)": cgst, "SGST (₹)": sgst, "Gross Total (₹)": gross
            })

    if 'Section 7(A)(2) in GSTR-1' in excel.sheet_names:
        df_7a = pd.read_excel(file_bytes, sheet_name='Section 7(A)(2) in GSTR-1')
        for _, r in df_7a.iterrows():
            taxable = safe_float(r.get('Aggregate Taxable Value Rs.', 0))
            cgst = safe_float(r.get('CGST Amount Rs.', 0))
            sgst = safe_float(r.get('SGST /UT Amount Rs.', 0))
            returns_taxable += safe_float(r.get('Taxable Sales Return Value Rs.', 0))
            if taxable > 0:
                b2cs_records.append({
                    "Platform": "Flipkart", "Place of Supply": f"{supplier_state}-Gujarat", "Rate": "5%", "Rate_Num": 5.0,
                    "Taxable Value (₹)": taxable, "IGST (₹)": 0.0, "CGST (₹)": cgst, "SGST (₹)": sgst,
                    "Gross Value (₹)": round(taxable + cgst + sgst, 2)
                })

    if 'Section 7(B)(2) in GSTR-1' in excel.sheet_names:
        df_7b = pd.read_excel(file_bytes, sheet_name='Section 7(B)(2) in GSTR-1')
        for _, r in df_7b.iterrows():
            taxable = safe_float(r.get('Aggregate Taxable Value Rs.', 0))
            igst = safe_float(r.get('IGST Amount Rs.', 0))
            state = str(r.get('Delivered State (PoS)', '')).strip()
            returns_taxable += safe_float(r.get('Taxable Sales Return Value Rs.', 0))
            if taxable > 0:
                b2cs_records.append({
                    "Platform": "Flipkart", "Place of Supply": state, "Rate": "5%", "Rate_Num": 5.0,
                    "Taxable Value (₹)": taxable, "IGST (₹)": igst, "CGST (₹)": 0.0, "SGST (₹)": 0.0,
                    "Gross Value (₹)": round(taxable + igst, 2)
                })

    tcs_total = 0.0
    if 'Section 3 in GSTR-8' in excel.sheet_names:
        df_tcs = pd.read_excel(file_bytes, sheet_name='Section 3 in GSTR-8')
        tcs_total = safe_float(df_tcs['TCS IGST amount Rs.'].sum()) + safe_float(df_tcs['TCS CGST amount Rs.'].sum()) + safe_float(df_tcs['TCS SGST amount Rs.'].sum())

    return {
        "platform": "Flipkart", "supplier_gstin": supplier_gstin, "supplier_state": supplier_state,
        "gross": gross_sum, "taxable": taxable_sum, "returns_gross": round(returns_taxable * 1.05, 2), "returns_taxable": round(returns_taxable, 2),
        "igst": igst_sum, "cgst": cgst_sum, "sgst": sgst_sum,
        "total_tax": igst_sum + cgst_sum + sgst_sum,
        "tcs": round(tcs_total, 2),
        "hsn": hsn_records, "b2cs": b2cs_records, "b2b": []
    }

# --- PARSER 3: MEESHO (NET SALES) ---
def parse_meesho_frames(df_sales, df_returns):
    df_sales = df_sales.copy()
    df_returns = df_returns.copy()
    df_sales.columns = [c.strip().lower() for c in df_sales.columns]
    df_returns.columns = [c.strip().lower() for c in df_returns.columns]
    
    supplier_gstin = "24ECEPM6676L1Z0"
    supplier_state = "24"
    if 'gstin' in df_sales.columns and not df_sales.empty:
        supplier_gstin = str(df_sales['gstin'].iloc[0]).strip().upper()
        supplier_state = supplier_gstin[:2]
    
    ret_gross = safe_float(df_returns['total_invoice_value'].sum()) if 'total_invoice_value' in df_returns.columns else 0.0
    ret_taxable = safe_float(df_returns['total_taxable_sale_value'].sum()) if 'total_taxable_sale_value' in df_returns.columns else 0.0

    df_sales['sign'] = 1
    df_returns['sign'] = -1
    df_all = pd.concat([df_sales, df_returns], ignore_index=True)
    
    df_all['net_taxable'] = df_all['total_taxable_sale_value'] * df_all['sign']
    df_all['net_gross'] = df_all['total_invoice_value'] * df_all['sign']
    df_all['net_tax'] = df_all['tax_amount'] * df_all['sign']
    df_all['net_qty'] = df_all['quantity'] * df_all['sign']
    
    def is_intra_state(state_val):
        s = str(state_val).strip().upper()
        return s.startswith(supplier_state) or s == 'GUJARAT' or s == 'IN-GJ'
    
    df_all['is_intra'] = df_all['end_customer_state_new'].apply(is_intra_state)
    df_all['igst'] = df_all.apply(lambda r: 0.0 if r['is_intra'] else r['net_tax'], axis=1)
    df_all['cgst'] = df_all.apply(lambda r: (r['net_tax'] / 2.0) if r['is_intra'] else 0.0, axis=1)
    df_all['sgst'] = df_all.apply(lambda r: (r['net_tax'] / 2.0) if r['is_intra'] else 0.0, axis=1)
    
    taxable_sum = df_all['net_taxable'].sum()
    gross_sum = df_all['net_gross'].sum()
    igst_sum = df_all['igst'].sum()
    cgst_sum = df_all['cgst'].sum()
    sgst_sum = df_all['sgst'].sum()
    total_tax_sum = igst_sum + cgst_sum + sgst_sum
    
    state_grp = df_all.groupby('end_customer_state_new').agg({
        'net_taxable': 'sum', 'igst': 'sum', 'cgst': 'sum', 'sgst': 'sum', 'net_gross': 'sum', 'gst_rate': 'first'
    }).reset_index()
    
    b2cs_records = []
    for _, r in state_grp.iterrows():
        if round(r['net_taxable'], 2) != 0:
            b2cs_records.append({
                "Platform": "Meesho", "Place of Supply": str(r['end_customer_state_new']).strip().title(),
                "Rate": f"{r['gst_rate']:.0f}%", "Rate_Num": float(r['gst_rate']), "Taxable Value (₹)": round(r['net_taxable'], 2),
                "IGST (₹)": round(r['igst'], 2), "CGST (₹)": round(r['cgst'], 2), "SGST (₹)": round(r['sgst'], 2),
                "Gross Value (₹)": round(r['net_gross'], 2)
            })
            
    hsn_grp = df_all.groupby(['hsn_code', 'gst_rate']).agg({
        'net_qty': 'sum', 'net_taxable': 'sum', 'igst': 'sum', 'cgst': 'sum', 'sgst': 'sum', 'net_gross': 'sum'
    }).reset_index()
    
    hsn_records = []
    for _, r in hsn_grp.iterrows():
        hsn_records.append({
            "Platform": "Meesho", "HSN Code": str(int(r['hsn_code'])) if pd.notna(r['hsn_code']) else "9999",
            "UQC": "PCS", "Qty": r['net_qty'], "GST Rate": f"{r['gst_rate']:.0f}%", "Rate_Num": float(r['gst_rate']),
            "Taxable (₹)": round(r['net_taxable'], 2), "IGST (₹)": round(r['igst'], 2),
            "CGST (₹)": round(r['cgst'], 2), "SGST (₹)": round(r['sgst'], 2), "Gross Total (₹)": round(r['net_gross'], 2)
        })
        
    return {
        "platform": "Meesho", "supplier_gstin": supplier_gstin, "supplier_state": supplier_state,
        "gross": round(gross_sum, 2), "taxable": round(taxable_sum, 2),
        "returns_gross": round(ret_gross, 2), "returns_taxable": round(ret_taxable, 2),
        "igst": round(igst_sum, 2), "cgst": round(cgst_sum, 2), "sgst": round(sgst_sum, 2),
        "total_tax": round(total_tax_sum, 2), "tcs": round(taxable_sum * 0.005, 2),
        "hsn": hsn_records, "b2cs": b2cs_records, "b2b": []
    }

# --- UPLOADER & ZIP DISPATCHER ---
uploaded_files = st.file_uploader(
    "Upload GST Files (.xlsx, .xls, .zip)", 
    type=["xlsx", "xls", "zip", "csv"], 
    accept_multiple_files=True
)

if uploaded_files:
    platform_results = []
    
    for file_obj in uploaded_files:
        file_name = file_obj.name
        
        # 1. ZIP File Processing
        if file_name.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(file_obj) as z:
                    extracted_names = [n for n in z.namelist() if n.endswith(('.xlsx', '.xls', '.csv')) and not n.startswith('__MACOSX/')]
                    
                    if any('tcs_sales' in n for n in extracted_names):
                        sales_name = next(n for n in extracted_names if 'tcs_sales.' in n or n.endswith('tcs_sales.xlsx'))
                        returns_name = next((n for n in extracted_names if 'tcs_sales_return' in n), None)
                        
                        df_s = pd.read_excel(io.BytesIO(z.read(sales_name)))
                        df_r = pd.read_excel(io.BytesIO(z.read(returns_name))) if returns_name else pd.DataFrame(columns=df_s.columns)
                        
                        st.success(f"📦 **ZIP File:** `{file_name}` ➔ **Identified Platform:** **🟪 Meesho** (Net Sales Processed)")
                        platform_results.append(parse_meesho_frames(df_s, df_r))
                    else:
                        for inner_filename in extracted_names:
                            inner_bytes = io.BytesIO(z.read(inner_filename))
                            p_id, p_badge = detect_ecommerce_platform(inner_bytes, inner_filename)
                            inner_bytes.seek(0)
                            st.success(f"📦 **ZIP File:** `{file_name}` ➔ **Extracted:** `{inner_filename}` ➔ **Platform:** **{p_badge}**")
                            
                            if p_id == "Flipkart":
                                platform_results.append(parse_flipkart(inner_bytes))
                            elif p_id == "Meesho":
                                df_s = pd.read_excel(inner_bytes)
                                platform_results.append(parse_meesho_frames(df_s, pd.DataFrame(columns=df_s.columns)))
                            else:
                                platform_results.append(parse_amazon(inner_bytes))
            except Exception as e:
                st.error(f"Error unzipping {file_name}: {e}")
        else:
            # 2. Standalone Files
            try:
                p_id, p_badge = detect_ecommerce_platform(file_obj, file_name)
                file_obj.seek(0)
                st.success(f"📁 **File:** `{file_name}` ➔ **Platform:** **{p_badge}**")
                
                if p_id == "Flipkart":
                    platform_results.append(parse_flipkart(file_obj))
                elif p_id == "Meesho":
                    df_single = pd.read_excel(file_obj)
                    platform_results.append(parse_meesho_frames(df_single, pd.DataFrame(columns=df_single.columns)))
                else:
                    platform_results.append(parse_amazon(file_obj))
            except Exception as e:
                st.error(f"Error processing {file_name}: {e}")

    # Combine All Extracted Platform Data
    combined_gross = sum(p['gross'] for p in platform_results)
    combined_taxable = sum(p['taxable'] for p in platform_results)
    combined_igst = sum(p['igst'] for p in platform_results)
    combined_cgst = sum(p['cgst'] for p in platform_results)
    combined_sgst = sum(p['sgst'] for p in platform_results)
    combined_total_tax = sum(p['total_tax'] for p in platform_results)
    combined_tcs = sum(p['tcs'] for p in platform_results)
    total_returns_loss = sum(p['returns_gross'] for p in platform_results)
    
    all_hsn = [item for p in platform_results for item in p['hsn']]
    all_b2cs = [item for p in platform_results for item in p['b2cs']]
    all_b2b = [item for p in platform_results for item in p['b2b']]

    st.divider()

    # 1. Platform Comparison Table
    st.subheader("📊 Platform-Wise Sales & Tax Summary")
    comp_data = []
    for p in platform_results:
        comp_data.append({
            "Platform": p['platform'],
            "Gross Sales (₹)": f"₹{p['gross']:,.2f}",
            "Taxable Sales (₹)": f"₹{p['taxable']:,.2f}",
            "Returns / RTO Value (₹)": f"₹{p['returns_gross']:,.2f}",
            "IGST (₹)": f"₹{p['igst']:,.2f}",
            "CGST (₹)": f"₹{p['cgst']:,.2f}",
            "SGST (₹)": f"₹{p['sgst']:,.2f}",
            "Total Output GST (₹)": f"₹{p['total_tax']:,.2f}",
            "TCS Credit (₹)": f"₹{p['tcs']:,.2f}"
        })
    st.table(pd.DataFrame(comp_data))

    st.divider()

    # 2. Consolidated KPI Cards
    st.subheader("🌐 Combined Total Tax Liability (All Platforms)")
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("Total Gross Sales", f"₹{combined_gross:,.2f}")
    kpi2.metric("Total Taxable Sales", f"₹{combined_taxable:,.2f}")
    kpi3.metric("Total Output GST", f"₹{combined_total_tax:,.2f}")
    kpi4.metric("Total IGST", f"₹{combined_igst:,.2f}")
    kpi5.metric("Total CGST + SGST", f"₹{(combined_cgst + combined_sgst):,.2f}")

    # 3. GSTR-3B Cash Calculator & Late Fee Estimator
    st.divider()
    calc_c1, calc_c2 = st.columns([1.5, 1])
    with calc_c1:
        st.subheader("💳 GSTR-3B Final Cash Payment Calculator")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            itc_igst = c1.number_input("Purchase IGST (ITC)", min_value=0.0, value=0.0, step=100.0)
        with c2:
            itc_cgst = c2.number_input("Purchase CGST (ITC)", min_value=0.0, value=0.0, step=100.0)
        with c3:
            itc_sgst = c3.number_input("Purchase SGST (ITC)", min_value=0.0, value=0.0, step=100.0)
        with c4:
            tcs_claim = c4.number_input("Combined TCS Credit", min_value=0.0, value=round(combined_tcs, 2), step=50.0)

        net_igst = max(0.0, combined_igst - itc_igst)
        net_cgst = max(0.0, combined_cgst - itc_cgst)
        net_sgst = max(0.0, combined_sgst - itc_sgst)
        final_cash_tax = max(0.0, (net_igst + net_cgst + net_sgst) - tcs_claim)

        st.success(f"👉 **Consolidated Net Cash to Pay (GSTR-3B Challan):** ₹{round(final_cash_tax):,} *(IGST: ₹{net_igst:,.2f} | CGST: ₹{net_cgst:,.2f} | SGST: ₹{net_sgst:,.2f})*")

    with calc_c2:
        st.subheader("⏰ Section 47 Late Fee & Interest Estimator")
        days_delayed = st.number_input("Days Delayed (After 20th Due Date)", min_value=0, value=0, step=1)
        late_fee = days_delayed * 50.0 # Rs 50/day standard
        interest_18 = (final_cash_tax * 0.18 * (days_delayed / 365.0)) if days_delayed > 0 else 0.0
        
        if days_delayed > 0:
            st.warning(f"⚠️ **Estimated Late Fee:** ₹{late_fee:,.2f} | **Interest (18% p.a.):** ₹{interest_18:,.2f}")
        else:
            st.info("✅ On-Time Filing: Late Fee ₹0.00 | Interest ₹0.00")

    st.divider()

    # 4. WhatsApp & Email 1-Click Shareable Summary
    st.subheader("📲 WhatsApp / Email Ready Filing Summary")
    wa_summary_text = f"""*GST FILING SUMMARY FOR ACCOUNTANT/CA*
----------------------------------------
*Turnover Details:*
• Gross Turnover: ₹{combined_gross:,.2f}
• Net Taxable Sales: ₹{combined_taxable:,.2f}
• Total Output Tax: ₹{combined_total_tax:,.2f}
  - IGST: ₹{combined_igst:,.2f}
  - CGST: ₹{combined_cgst:,.2f}
  - SGST: ₹{combined_sgst:,.2f}

*GSTR-3B Payment Challan:*
• Total Eligible ITC Claimed: ₹{(itc_igst + itc_cgst + itc_sgst):,.2f}
• E-Commerce TCS Credit: ₹{tcs_claim:,.2f}
• *Net Cash Challan Amount: ₹{round(final_cash_tax):,}*

*Platform Split:*
{chr(10).join([f"• {p['platform']}: Taxable ₹{p['taxable']:,.2f} | Tax ₹{p['total_tax']:,.2f}" for p in platform_results])}
----------------------------------------
_Generated via Multi-Platform GST Pro Assistant_"""
    
    st.text_area("Copy this text and paste on WhatsApp / Email:", value=wa_summary_text, height=180)

    st.divider()

    # 5. Return & RTO Impact Analytics
    st.subheader("📦 Return & RTO Impact Analytics")
    rto_c1, rto_c2 = st.columns([1, 1])
    with rto_c1:
        st.metric("Total Returns & RTO Loss (Gross)", f"₹{total_returns_loss:,.2f}")
        st.caption("E-Commerce returns aur cancellations ki wajah se blocked revenue turnover.")
    with rto_c2:
        rto_chart_data = pd.DataFrame([{"Platform": p['platform'], "Returns (₹)": p['returns_gross']} for p in platform_results]).set_index('Platform')
        st.bar_chart(rto_chart_data)

    st.divider()

    # 6. GSTR-3B & Table 14 Mapping
    map_c1, map_c2 = st.columns([1, 1])
    with map_c1:
        st.subheader("📋 Direct GSTR-3B Portal Form Mapping")
        gstr3b_table = [
            {"GST Portal Section": "Table 3.1(a) Outward Taxable Supplies", "Taxable Value (₹)": f"₹{combined_taxable:,.2f}", "IGST (₹)": f"₹{combined_igst:,.2f}", "CGST (₹)": f"₹{combined_cgst:,.2f}", "SGST (₹)": f"₹{combined_sgst:,.2f}"},
            {"GST Portal Section": "Table 4(A)(5) All Other Eligible ITC", "Taxable Value (₹)": "-", "IGST (₹)": f"₹{itc_igst:,.2f}", "CGST (₹)": f"₹{itc_cgst:,.2f}", "SGST (₹)": f"₹{itc_sgst:,.2f}"},
            {"GST Portal Section": "Table 6.1 Payment of Tax (Net Cash)", "Taxable Value (₹)": "-", "IGST (₹)": f"₹{net_igst:,.2f}", "CGST (₹)": f"₹{net_cgst:,.2f}", "SGST (₹)": f"₹{net_sgst:,.2f}"}
        ]
        st.table(pd.DataFrame(gstr3b_table))

    with map_c2:
        st.subheader("🏬 GSTR-1 Table 14 (Supplies via ECO - Sec 52)")
        eco_table = []
        for p in platform_results:
            eco_table.append({
                "E-Commerce Platform": p['platform'],
                "Net Taxable Value (₹)": f"₹{p['taxable']:,.2f}",
                "IGST (₹)": f"₹{p['igst']:,.2f}",
                "CGST (₹)": f"₹{p['cgst']:,.2f}",
                "SGST (₹)": f"₹{p['sgst']:,.2f}"
            })
        st.table(pd.DataFrame(eco_table))

    st.divider()

    # 7. Export Center
    st.subheader("📥 Export Combined Reports & Official Offline Utility Template")
    d1, d2, d3 = st.columns(3)
    
    df_hsn_unified = pd.DataFrame(all_hsn)
    if not df_hsn_unified.empty:
        df_hsn_unified = df_hsn_unified.groupby(['HSN Code', 'GST Rate']).agg({
            'Qty': 'sum', 'Taxable (₹)': 'sum', 'IGST (₹)': 'sum', 'CGST (₹)': 'sum', 'SGST (₹)': 'sum', 'Gross Total (₹)': 'sum'
        }).reset_index()

    df_b2c_unified = pd.DataFrame(all_b2cs)
    if not df_b2c_unified.empty:
        df_b2c_unified = df_b2c_unified.groupby(['Place of Supply', 'Rate']).agg({
            'Taxable Value (₹)': 'sum', 'IGST (₹)': 'sum', 'CGST (₹)': 'sum', 'SGST (₹)': 'sum', 'Gross Value (₹)': 'sum'
        }).reset_index()

    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine='openpyxl') as writer:
        pd.DataFrame(comp_data).to_excel(writer, sheet_name='Platform Summary', index=False)
        if not df_hsn_unified.empty:
            df_hsn_unified.to_excel(writer, sheet_name='Unified HSN Table 12', index=False)
        if not df_b2c_unified.empty:
            df_b2c_unified.to_excel(writer, sheet_name='Unified B2C Table 7', index=False)
        pd.DataFrame(all_hsn).to_excel(writer, sheet_name='Raw Combined HSN', index=False)
        pd.DataFrame(all_b2cs).to_excel(writer, sheet_name='Raw Combined B2C', index=False)
        if all_b2b:
            pd.DataFrame(all_b2b).to_excel(writer, sheet_name='B2B Invoices', index=False)
            
    with d1:
        st.download_button("📊 Download Consolidated Excel (CA Audit)", data=excel_buf.getvalue(), file_name="GST_Combined_Platform_Audit.xlsx", use_container_width=True)

    offline_excel_buf = io.BytesIO()
    with pd.ExcelWriter(offline_excel_buf, engine='openpyxl') as writer:
        b2cs_offline_rows = []
        for x in all_b2cs:
            b2cs_offline_rows.append({
                "Type": "OE", "Place of Supply": x.get('Place of Supply', ''),
                "Applicable % of Tax Rate": "", "Rate": x.get('Rate_Num', 5.0),
                "Taxable Value": x.get('Taxable Value (₹)', 0.0), "Cess Amount": 0.0
            })
        pd.DataFrame(b2cs_offline_rows).to_excel(writer, sheet_name='b2cs', index=False)

        hsn_offline_rows = []
        for x in all_hsn:
            hsn_offline_rows.append({
                "HSN": x.get('HSN Code', ''), "Description": "Goods", "UQC": x.get('UQC', 'PCS'),
                "Total Quantity": x.get('Qty', 1.0), "Total Value": x.get('Gross Total (₹)', 0.0),
                "Taxable Value": x.get('Taxable (₹)', 0.0), "Integrated Tax Amount": x.get('IGST (₹)', 0.0),
                "Central Tax Amount": x.get('CGST (₹)', 0.0), "State/UT Tax Amount": x.get('SGST (₹)', 0.0),
                "Cess Amount": 0.0
            })
        pd.DataFrame(hsn_offline_rows).to_excel(writer, sheet_name='hsn', index=False)

    with d2:
        st.download_button("🏛️ Download Official GSTR-1 Offline Template (.xlsx)", data=offline_excel_buf.getvalue(), file_name="GSTR1_Offline_Utility_Import.xlsx", use_container_width=True)

    json_payload = {
        "gross_sales": round(combined_gross, 2), "taxable_sales": round(combined_taxable, 2),
        "total_tax": round(combined_total_tax, 2), "igst": round(combined_igst, 2),
        "cgst": round(combined_cgst, 2), "sgst": round(combined_sgst, 2),
        "platforms": comp_data, "unified_hsn": df_hsn_unified.to_dict(orient='records') if not df_hsn_unified.empty else [],
        "unified_b2cs": df_b2c_unified.to_dict(orient='records') if not df_b2c_unified.empty else []
    }
    with d3:
        st.download_button("⚡ Download Combined GSTR-1 JSON", data=json.dumps(json_payload, indent=4).encode('utf-8'), file_name="GSTR1_Combined_Offline.json", mime="application/json", use_container_width=True)

    st.divider()

    # 8. Detailed Tables Breakdown
    st.subheader("📋 Platform-Wise & Combined Detailed Data Breakdown")
    main_tab1, main_tab2, main_tab3 = st.tabs(["📦 HSN Summary (Table 12)", "🛒 B2C State-Wise Sales (Table 7)", "🏬 B2B Invoices (Table 4A)"])

    with main_tab1:
        st.write("### HSN Wise Breakdown")
        hsn_sub_tabs = st.tabs(["🌟 Unified Aggregated HSN (Portal Ready)", "🌐 Raw Combined HSN", "🟧 Amazon HSN", "🟦 Flipkart HSN", "🟪 Meesho HSN"])
        with hsn_sub_tabs[0]:
            st.dataframe(df_hsn_unified if not df_hsn_unified.empty else pd.DataFrame(), use_container_width=True)
        with hsn_sub_tabs[1]:
            st.dataframe(pd.DataFrame(all_hsn) if all_hsn else pd.DataFrame(), use_container_width=True)
        with hsn_sub_tabs[2]:
            amz_hsn = [x for x in all_hsn if x.get("Platform") == "Amazon"]
            st.dataframe(pd.DataFrame(amz_hsn) if amz_hsn else pd.DataFrame(), use_container_width=True)
        with hsn_sub_tabs[3]:
            fk_hsn = [x for x in all_hsn if x.get("Platform") == "Flipkart"]
            st.dataframe(pd.DataFrame(fk_hsn) if fk_hsn else pd.DataFrame(), use_container_width=True)
        with hsn_sub_tabs[4]:
            meesho_hsn = [x for x in all_hsn if x.get("Platform") == "Meesho"]
            st.dataframe(pd.DataFrame(meesho_hsn) if meesho_hsn else pd.DataFrame(), use_container_width=True)

    with main_tab2:
        st.write("### B2C State-Wise Sales Breakdown")
        b2c_sub_tabs = st.tabs(["🌟 Unified Aggregated States (Portal Ready)", "🌐 Raw Combined B2C", "🟧 Amazon B2C", "🟦 Flipkart B2C", "🟪 Meesho B2C"])
        with b2c_sub_tabs[0]:
            st.dataframe(df_b2c_unified if not df_b2c_unified.empty else pd.DataFrame(), use_container_width=True)
        with b2c_sub_tabs[1]:
            st.dataframe(pd.DataFrame(all_b2cs) if all_b2cs else pd.DataFrame(), use_container_width=True)
        with b2c_sub_tabs[2]:
            amz_b2cs = [x for x in all_b2cs if x.get("Platform") == "Amazon"]
            st.dataframe(pd.DataFrame(amz_b2cs) if amz_b2cs else pd.DataFrame(), use_container_width=True)
        with b2c_sub_tabs[3]:
            fk_b2cs = [x for x in all_b2cs if x.get("Platform") == "Flipkart"]
            st.dataframe(pd.DataFrame(fk_b2cs) if fk_b2cs else pd.DataFrame(), use_container_width=True)
        with b2c_sub_tabs[4]:
            meesho_b2cs = [x for x in all_b2cs if x.get("Platform") == "Meesho"]
            st.dataframe(pd.DataFrame(meesho_b2cs) if meesho_b2cs else pd.DataFrame(), use_container_width=True)

    with main_tab3:
        st.write("### B2B Registered Invoices")
        b2b_sub_tabs = st.tabs(["🌐 Combined B2B", "🟧 Amazon B2B", "🟦 Flipkart B2B", "🟪 Meesho B2B"])
        with b2b_sub_tabs[0]:
            st.dataframe(pd.DataFrame(all_b2b) if all_b2b else pd.DataFrame(), use_container_width=True)
        with b2b_sub_tabs[1]:
            amz_b2b = [x for x in all_b2b if x.get("Platform") == "Amazon"]
            st.dataframe(pd.DataFrame(amz_b2b) if amz_b2b else pd.DataFrame(), use_container_width=True)
        with b2b_sub_tabs[2]:
            fk_b2b = [x for x in all_b2b if x.get("Platform") == "Flipkart"]
            st.dataframe(pd.DataFrame(fk_b2b) if fk_b2b else pd.DataFrame(), use_container_width=True)
        with b2b_sub_tabs[3]:
            meesho_b2b = [x for x in all_b2b if x.get("Platform") == "Meesho"]
            st.dataframe(pd.DataFrame(meesho_b2b) if meesho_b2b else pd.DataFrame(), use_container_width=True)
