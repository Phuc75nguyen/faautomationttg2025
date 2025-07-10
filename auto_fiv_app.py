import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="FIV Generator", layout="wide")

def detect_header_row(df_raw):
    """Tìm dòng header (dòng chứa 'STT') trong DataFrame raw."""
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains('STT', na=False).any():
            return idx
    raise ValueError("Không tìm thấy dòng header chứa 'STT'")

def load_and_flatten_eas(eas_bytes):
    """Đọc file EAS.xlsx với 2 dòng header, sau đó flatten tên cột."""
    # Đọc nguyên file vào DataFrame không header để detect header_row
    df_raw = pd.read_excel(io.BytesIO(eas_bytes), header=None)
    header_row = detect_header_row(df_raw)
    # Đọc lại với 2 dòng header
    df = pd.read_excel(io.BytesIO(eas_bytes), header=[header_row, header_row+1])
    # Flatten multi-index columns
    flat_cols = []
    for top, sub in df.columns:
        if pd.notna(sub) and not str(sub).startswith("Unnamed"):
            flat_cols.append(str(sub).strip())
        else:
            flat_cols.append(str(top).strip())
    df.columns = flat_cols
    return df

def clean_eas(df):
    """Đổi tên các cột quan trọng và lọc bỏ dòng thiếu Buyer Name hoặc Revenue."""
    # Đổi tên cố định
    rename_map = {
        'Tên người mua(Buyer Name)': 'Buyer Name',
        'Ngày, tháng, năm phát hành': 'ISSUE_DATE',
        'Doanh số bán chưa có thuế(Revenue excluding VAT)': 'Revenue_ex_VAT',
        'Thuế GTGT(VAT amount)': 'VAT_Amount',
        'Ký hiệu mẫu hóa đơn': 'InvoiceSerial',
        'Số hóa đơn': 'InvoiceNumber'
    }
    df = df.rename(columns=rename_map)

    # Tự động detect cột MST/Tax code nếu có
    mst_col = next((c for c in df.columns 
                    if 'Mã số thuế' in c or 'Tax code' in c), None)
    if mst_col:
        df = df.rename(columns={mst_col: 'TaxCode'})

    # Chỉ giữ các dòng có đủ Buyer Name & Revenue_ex_VAT
    df = df.dropna(subset=['Buyer Name', 'Revenue_ex_VAT']).reset_index(drop=True)
    return df

def build_fiv(df_eas, df_kh):
    """Tạo DataFrame FIV, ưu tiên lookup theo TaxCode rồi fallback Buyer Name."""
    # Detect cột MST/CMND/PASSPORT trên sheet KH
    taxkey_kh = next((c for c in df_kh.columns 
                      if any(x in c for x in ['MST','CMND','PASSPORT','Tax code'])), None)

    records = []
    for idx, row in df_eas.iterrows():
        buyer = row['Buyer Name']
        cust_acc = pd.NA

        # 1) Lookup theo MST trước (nếu có)
        if 'TaxCode' in row and pd.notna(row['TaxCode']) and taxkey_kh:
            m = df_kh[df_kh[taxkey_kh] == row['TaxCode']]['Customer account']
            if not m.empty:
                cust_acc = m.iloc[0]

        # 2) Nếu chưa tìm được, fallback theo Buyer Name
        if pd.isna(cust_acc):
            m = df_kh[df_kh['Name'] == buyer]['Customer account']
            if not m.empty:
                cust_acc = m.iloc[0]

        # Tính amounts
        line_amount = row['Revenue_ex_VAT']
        vat_amount  = row.get('VAT_Amount', 0)
        total_amt   = line_amount + vat_amount

        records.append({
            'IdRef':                         idx + 1,
            'InvoiceDate':                   row['ISSUE_DATE'],
            'DocumentDate':                  row['ISSUE_DATE'],
            'CurrencyCode':                  'VND',
            'CustAccount':                   cust_acc,
            'InvoiceAccount':                cust_acc,
            'SalesName':                     buyer,
            'APMA_DimA':                     'TX',
            'APMC_DimC':                     '',
            'APMD_DimD':                     '',
            'APMF_DimF':                     '',
            'TaxGroupHeader':                '131103',
            'PostingProfile':                1,
            'LineNum':                       1,
            'Description':                   'Doanh thu dịch vụ spa',
            'SalesPrice':                    line_amount,
            'SalesQty':                      1,
            'LineAmount':                    line_amount,
            'TaxAmount':                     vat_amount,
            'TotalAmount':                   total_amt,
            'TaxGroupLine':                  'OU',
            'TaxItemGroup':                  '10%',
            'Line_MainAccountId':            '511301',
            'Line_APMA_DimA':                'TX',
            'Line_APMC_DimC':                '5301',
            'Line_APMD_DimD':                '00',
            'Line_APMF_DimF':                '0000',
            'BHS_VATInvoiceDate_VATInvoice': row['ISSUE_DATE'],
            'BHS_Form_VATInvoice':           row.get('InvoiceSerial', ''),
            'BHS_Serial_VATInvoice':         row.get('InvoiceSerial', ''),
            'BHS_Number_VATInvoice':         row.get('InvoiceNumber', ''),
            'BHS_Description_VATInvoice':    'Doanh thu dịch vụ spa'
        })

    columns_order = [
        'IdRef','InvoiceDate','DocumentDate','CurrencyCode','CustAccount','InvoiceAccount',
        'SalesName','APMA_DimA','APMC_DimC','APMD_DimD','APMF_DimF','TaxGroupHeader',
        'PostingProfile','LineNum','Description','SalesPrice','SalesQty','LineAmount',
        'TaxAmount','TotalAmount','TaxGroupLine','TaxItemGroup','Line_MainAccountId',
        'Line_APMA_DimA','Line_APMC_DimC','Line_APMD_DimD','Line_APMF_DimF',
        'BHS_VATInvoiceDate_VATInvoice','BHS_Form_VATInvoice','BHS_Serial_VATInvoice',
        'BHS_Number_VATInvoice','BHS_Description_VATInvoice'
    ]
    return pd.DataFrame(records, columns=columns_order)

st.title("🧾 FIV Generator")
st.markdown("""
Upload hai file **EAS.xlsx** và **KH.xlsx**, ứng dụng sẽ tự động sinh file **Completed_FIV.xlsx**  
- Ưu tiên lookup theo MST/Tax code  
- Fallback theo Buyer Name nếu MST không tìm thấy  
- Tính `TotalAmount = Revenue_ex_VAT + VAT_Amount`
""")

eas_file = st.file_uploader("Chọn file EAS.xlsx", type="xlsx")
kh_file  = st.file_uploader("Chọn file KH.xlsx", type="xlsx")

if eas_file and kh_file:
    try:
        df_kh  = pd.read_excel(kh_file)
        eas_bytes = eas_file.read()
        df_raw   = load_and_flatten_eas(eas_bytes)
        df_eas   = clean_eas(df_raw)
        df_fiv   = build_fiv(df_eas, df_kh)

        towrite = io.BytesIO()
        with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
            df_fiv.to_excel(writer, index=False, sheet_name="FIV")
        towrite.seek(0)

        st.download_button(
            "📥 Tải Completed_FIV.xlsx",
            data=towrite,
            file_name="Completed_FIV.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Có lỗi xảy ra: {e}")
