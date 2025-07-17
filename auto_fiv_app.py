import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="🧾 FIV Generator", layout="wide")

def detect_header_row(df_raw):
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains('STT', na=False).any():
            return idx
    raise ValueError("Không tìm thấy dòng header chứa 'STT'")

def load_and_flatten_eas(eas_bytes):
    df_raw = pd.read_excel(io.BytesIO(eas_bytes), header=None)
    df_raw.iloc[:, 0] = df_raw.iloc[:, 0].astype(str)
    df_raw = df_raw[~df_raw.iloc[:, 0].str.contains(r'^\[\d+\]$', na=False)].reset_index(drop=True)
    header_row = detect_header_row(df_raw)
    df = pd.read_excel(io.BytesIO(eas_bytes), header=[header_row, header_row+1])
    # Flatten MultiIndex columns
    flat = []
    for top, sub in df.columns:
        if pd.notna(sub) and not str(sub).startswith("Unnamed"):
            flat.append(sub)
        else:
            flat.append(top)
    df.columns = [str(x).strip() for x in flat]
    return df

def clean_eas(df):
    rename_map = {
        'Tên người mua(Buyer Name)': 'Buyer Name',
        'Ngày, tháng, năm phát hành': 'ISSUE_DATE',
        'Doanh số bán chưa có thuế(Revenue excluding VAT)': 'Revenue_ex_VAT',
        'Thuế GTGT(VAT amount)': 'VAT_Amount',
        'Ký hiệu mẫu hóa đơn': 'InvoiceSerial',
        'Số hóa đơn': 'InvoiceNumber'
    }
    df = df.rename(columns=rename_map)
    mst_col = next((c for c in df.columns if 'Mã số thuế' in c or 'Tax code' in c), None)
    if mst_col:
        df = df.rename(columns={mst_col: 'TaxCode'})
    return df.dropna(subset=['Buyer Name', 'Revenue_ex_VAT']).reset_index(drop=True)

def build_fiv(df_eas, df_kh):
    taxkey = next((c for c in df_kh.columns 
                   if any(x in c for x in ['MST','CMND','PASSPORT','Tax code'])), None)
    recs = []
    for i, row in df_eas.iterrows():
        buyer = row['Buyer Name']
        acc = pd.NA
        if 'TaxCode' in row and pd.notna(row['TaxCode']) and taxkey:
            m = df_kh[df_kh[taxkey] == row['TaxCode']]['Customer account']
            if not m.empty: acc = m.iat[0]
        if pd.isna(acc):
            m = df_kh[df_kh['Name']==buyer]['Customer account']
            if not m.empty: acc = m.iat[0]

        rev = row['Revenue_ex_VAT']
        vat = row.get('VAT_Amount', 0)
        total = rev + vat

        recs.append({
            'IdRef':                         i+1,
            'InvoiceDate':                   row['ISSUE_DATE'],
            'DocumentDate':                  row['ISSUE_DATE'],
            'CurrencyCode':                  'VND',
            'CustAccount':                   acc,
            'InvoiceAccount':                acc,
            'SalesName':                     buyer,
            'APMA_DimA':                     'TX',
            'APMC_DimC':                     '0000',
            'APMD_DimD':                     '00',
            'APMF_DimF':                     '0000',
            'TaxGroupHeader':                'OU',
            'PostingProfile':                '131103',
            'LineNum':                       1,
            'Description':                   'Doanh thu dịch vụ spa',
            'SalesPrice':                    rev,
            'SalesQty':                      1,
            'LineAmount':                    rev,
            'TaxAmount':                     vat,
            'TotalAmount':                   total,
            'TaxGroupLine':                  'OU',
            'TaxItemGroup':                  '10%',
            'Line_MainAccountId':            '511301',
            'Line_APMA_DimA':                'TX',
            'Line_APMC_DimC':                '5301',
            'Line_APMD_DimD':                '00',
            'Line_APMF_DimF':                '0000',
            'BHS_VATInvocieDate_VATInvoice': row['ISSUE_DATE'],
            'BHS_Form_VATInvoice':           '',
            'BHS_Serial_VATInvoice':         row.get('InvoiceSerial',''),
            'BHS_Number_VATInvoice':         row.get('InvoiceNumber',''),
            'BHS_Description_VATInvoice':    'Doanh thu dịch vụ spa'
        })
    cols = recs[0].keys()
    return pd.DataFrame(recs, columns=cols)

# --- Streamlit UI ---
st.title("🧾 FIV Generator")
st.markdown("""
Upload hai file **EAS.xlsx** và **KH.xlsx**, ứng dụng sẽ tự động sinh file **Completed_FIV.xlsx**  
- Lookup ưu tiên theo MST/Tax code  
- Fallback theo Buyer Name  
- Tính TotalAmount = Revenue_ex_VAT + VAT_Amount  
- IdRef xuất dạng TEXT (tam giác xanh)  
- Các cột date format dd/mm/yyyy
""")

eas_file = st.file_uploader("Chọn file EAS.xlsx", type="xlsx")
kh_file  = st.file_uploader("Chọn file KH.xlsx", type="xlsx")

if eas_file and kh_file:
    try:
        df_kh     = pd.read_excel(kh_file)
        eas_bytes = eas_file.read()

        df_raw = load_and_flatten_eas(eas_bytes)
        df_eas = clean_eas(df_raw)
        df_fiv = build_fiv(df_eas, df_kh)

        # --- Ép IdRef thành text string ---
        df_fiv['IdRef'] = df_fiv['IdRef'].astype(str)

        # --- Giữ datetime64 chỉ phần ngày, loại bỏ giờ ---
        for c in ['InvoiceDate','DocumentDate','BHS_VATInvocieDate_VATInvoice']:
            df_fiv[c] = pd.to_datetime(df_fiv[c], errors='raise').dt.normalize()

        # --- Xuất Excel với định dạng cột ---
        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            df_fiv.to_excel(writer, index=False, sheet_name='FIV')
            wb = writer.book
            ws = writer.sheets['FIV']

            # IdRef → Text format để có tam giác xanh
            tf = wb.add_format({'num_format': '@'})
            ws.set_column('A:A', 10, tf)

            # Date format dd/mm/yyyy (dấu slash literal)
            dfmt = wb.add_format({'num_format': 'dd\\/mm\\/yyyy'})
            # cột B,C và AB
            ws.set_column('B:C', 12, dfmt)
            ws.set_column('AB:AB', 12, dfmt)

        out.seek(0)
        st.download_button(
            "📥 Tải Completed_FIV.xlsx",
            data=out.getvalue(),
            file_name="Completed_FIV.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Có lỗi: {e}")
