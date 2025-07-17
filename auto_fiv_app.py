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
    """Đọc file EAS.xlsx, bỏ qua các dòng không chứa dữ liệu thực tế (dòng [1], [2], [3]...)"""
    df_raw = pd.read_excel(io.BytesIO(eas_bytes), header=None)
    df_raw.iloc[:, 0] = df_raw.iloc[:, 0].astype(str)
    df_raw = df_raw[~df_raw.iloc[:, 0].str.contains(r'^\[\d+\]$', na=False)].reset_index(drop=True)
    header_row = detect_header_row(df_raw)
    df = pd.read_excel(io.BytesIO(eas_bytes), header=[header_row, header_row+1])

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

    df = df.dropna(subset=['Buyer Name', 'Revenue_ex_VAT']).reset_index(drop=True)
    return df

def build_fiv(df_eas, df_kh):
    """Tạo DataFrame FIV, ưu tiên lookup theo TaxCode rồi fallback Buyer Name."""
    taxkey_kh = next((c for c in df_kh.columns if any(x in c for x in ['MST','CMND','PASSPORT','Tax code'])), None)

    records = []
    for idx, row in df_eas.iterrows():
        buyer = row['Buyer Name']
        cust_acc = pd.NA

        if 'TaxCode' in row and pd.notna(row['TaxCode']) and taxkey_kh:
            m = df_kh[df_kh[taxkey_kh] == row['TaxCode']]['Customer account']
            if not m.empty:
                cust_acc = m.iloc[0]

        if pd.isna(cust_acc):
            m = df_kh[df_kh['Name'] == buyer]['Customer account']
            if not m.empty:
                cust_acc = m.iloc[0]

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
            'APMC_DimC':                     '0000',
            'APMD_DimD':                     '00',
            'APMF_DimF':                     '0000',
            'TaxGroupHeader':                'OU',
            'PostingProfile':                '131103',
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
            'BHS_VATInvocieDate_VATInvoice': row['ISSUE_DATE'],
            'BHS_Form_VATInvoice':           '',
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
        'BHS_VATInvocieDate_VATInvoice','BHS_Form_VATInvoice','BHS_Serial_VATInvoice',
        'BHS_Number_VATInvoice','BHS_Description_VATInvoice'
    ]
    return pd.DataFrame(records, columns=columns_order)

st.title("🧾 FIV Generator")
st.markdown("""
Upload hai file **EAS.xlsx** và **KH.xlsx**, ứng dụng sẽ tự động sinh file **Completed_FIV.xlsx**  
- Ưu tiên lookup theo MST/Tax code  
- Fallback theo Buyer Name nếu MST không tìm thấy  
- Tính TotalAmount = Revenue_ex_VAT + VAT_Amount
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

        # 1. Chuyển đổi các cột ngày tháng sang đúng kiểu datetime của pandas
        date_columns = ['InvoiceDate', 'DocumentDate', 'BHS_VATInvocieDate_VATInvoice']
        for col in date_columns:
            # Chuyển đổi và bỏ thông tin giờ/phút/giây không cần thiết
            #df_fiv[col] = pd.to_datetime(df_fiv[col]).dt.normalize()
            #for col in date_columns:
            #df_fiv[col] = pd.to_datetime(df_fiv[col]).dt.date
            df_fiv[col] = pd.to_datetime(df_fiv[col], errors='raise').dt.normalize()
            # 2) Ép thành chuỗi theo định dạng dd-mm-yyyy
            #df_fiv[col] = df_fiv[col].dt.strftime('%d-%m-%Y')
            df_fiv[col] = df_fiv[col].dt.strftime('%d/%m/%Y')

        with pd.ExcelWriter('Completed_FIV.xlsx', engine='openpyxl') as writer:
            df_fiv.to_excel(writer, index=False, sheet_name='FIV')

        towrite = io.BytesIO()
        # Xuất Excel với xlsxwriter để ép định dạng cột
        with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:        
            df_fiv.to_excel(writer, index=False, sheet_name='FIV')
            workbook  = writer.book
            worksheet = writer.sheets['FIV']

            # Format Text cho cột A (IdRef)
            text_fmt = workbook.add_format({'num_format': '@'})
            worksheet.set_column(0, 0, 10, text_fmt)

            # Format Short Date (dd-mm-yyyy) cho cột B,C và AB
            date_fmt = workbook.add_format({'num_format': 'dd-mm-yyyy'})
            worksheet.set_column(1, 2, 12, date_fmt)    # InvoiceDate & DocumentDate
            worksheet.set_column(27, 27, 12, date_fmt)  # BHS_VATInvocieDate_VATInvoice

            towrite.seek(0)

        st.download_button(
            "📥 Tải Completed_FIV.xlsx",
            data=towrite.getvalue(),
            file_name="Completed_FIV.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Có lỗi xảy ra: {e}")
