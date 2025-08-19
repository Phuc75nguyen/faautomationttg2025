import streamlit as st
import pandas as pd
import io
import datetime

# Set a generic page title since the application now supports multiple tools
st.set_page_config(page_title="📑 Automation Tools", layout="wide")

def parse_vietnamese_date(value: str) -> pd.Timestamp:
    """
    Convert a date string like '15 Thg 7 2023' to a pandas Timestamp.

    The input format consists of a day, the literal 'Thg', a month number,
    and a year. For example, '15 Thg 7 2023' means 15 July 2023.

    Parameters
    ----------
    value : str
        A date string in the format 'DD Thg M YYYY'.

    Returns
    -------
    pd.Timestamp or pd.NaT
        A pandas Timestamp representing the date, or NaT if parsing fails.
    """
    if isinstance(value, str):
        parts = value.split()
        # Expected format: [day, 'Thg', month, year]
        if len(parts) == 4 and parts[1].lower() == 'thg':
            day, _, month, year = parts
            try:
                # Compose a standard date string and convert
                date_str = f"{day}-{month}-{year}"
                return pd.to_datetime(date_str, format="%d-%m-%Y", errors="coerce")
            except Exception:
                return pd.NaT
    return pd.NaT

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

    # Flatten MultiIndex
    flat_cols = []
    for top, sub in df.columns:
        if pd.notna(sub) and not str(sub).startswith("Unnamed"):
            flat_cols.append(str(sub).strip())
        else:
            flat_cols.append(str(top).strip())
    df.columns = flat_cols
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
    df = df.dropna(subset=['Buyer Name', 'Revenue_ex_VAT']).reset_index(drop=True)
    return df

def build_fiv(df_eas, df_kh):
    taxkey_kh = next((c for c in df_kh.columns 
                      if any(x in c for x in ['MST','CMND','PASSPORT','Tax code'])), None)
    records = []
    for idx, row in df_eas.iterrows():
        buyer = row['Buyer Name']
        cust_acc = pd.NA
        if 'TaxCode' in row and pd.notna(row['TaxCode']) and taxkey_kh:
            m = df_kh[df_kh[taxkey_kh] == row['TaxCode']]['Customer account']
            if not m.empty:
                cust_acc = m.iat[0]
        if pd.isna(cust_acc):
            m = df_kh[df_kh['Name'] == buyer]['Customer account']
            if not m.empty:
                cust_acc = m.iat[0]

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

    cols_order = [
        'IdRef','InvoiceDate','DocumentDate','CurrencyCode','CustAccount','InvoiceAccount',
        'SalesName','APMA_DimA','APMC_DimC','APMD_DimD','APMF_DimF','TaxGroupHeader',
        'PostingProfile','LineNum','Description','SalesPrice','SalesQty','LineAmount',
        'TaxAmount','TotalAmount','TaxGroupLine','TaxItemGroup','Line_MainAccountId',
        'Line_APMA_DimA','Line_APMC_DimC','Line_APMD_DimD','Line_APMF_DimF',
        'BHS_VATInvocieDate_VATInvoice','BHS_Form_VATInvoice','BHS_Serial_VATInvoice',
        'BHS_Number_VATInvoice','BHS_Description_VATInvoice'
    ]
    return pd.DataFrame(records, columns=cols_order)

# Add a sidebar for selecting which tool to use
st.sidebar.title("Chức năng")
tool_choice = st.sidebar.radio(
    "Chọn chức năng",
    ["Senspa Automation Excel-AX", "Agoda LCB"],
    index=0,
)

if tool_choice == "Senspa Automation Excel-AX":
    # FIV generator UI
    st.title("🧾 FIV Generator")
    st.markdown("""
    Upload hai file **EAS.xlsx** và **KH.xlsx**, ứng dụng sẽ tự động sinh file **Completed_FIV.xlsx**  
    - Lookup ưu tiên theo MST/Tax code  
    - Fallback theo Buyer Name  
    - Tính TotalAmount = Revenue_ex_VAT + VAT_Amount
    """)

    eas_file = st.file_uploader("Chọn file EAS.xlsx", type="xlsx", key="eas")
    kh_file  = st.file_uploader("Chọn file KH.xlsx", type="xlsx", key="kh")

    if eas_file and kh_file:
        try:
            df_kh     = pd.read_excel(kh_file)
            eas_bytes = eas_file.read()

            df_raw = load_and_flatten_eas(eas_bytes)
            df_eas = clean_eas(df_raw)
            df_fiv = build_fiv(df_eas, df_kh)

            df_fiv['IdRef'] = df_fiv['IdRef'].astype(str)
            date_cols = ['InvoiceDate', 'DocumentDate', 'BHS_VATInvocieDate_VATInvoice']
            for c in date_cols:
                df_fiv[c] = pd.to_datetime(df_fiv[c]).dt.date

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter', date_format='dd/mm/yyyy') as writer:
                df_fiv.to_excel(writer, index=False, sheet_name='FIV')

            output.seek(0)
            st.download_button(
                "📥 Tải Completed_FIV.xlsx",
                data=output.getvalue(),
                file_name="Completed_FIV.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Có lỗi: {e}")

elif tool_choice == "Agoda LCB":
    # Agoda LCB processing UI
    st.title("🏨 Agoda LCB Processor")
    st.markdown("""
    Tải lên file Excel đối chiếu từ Agoda, chọn khoảng ngày trả phòng muốn lọc,
    và ứng dụng sẽ xuất ra file đã được xử lý.
    """)

    # Allow the user to pick a date range; default to the last 7 days
    today = datetime.date.today()
    default_start = today - datetime.timedelta(days=7)
    default_end = today
    col1, col2 = st.columns(2)
    start_date = col1.date_input("Ngày bắt đầu", value=default_start, max_value=today)
    end_date = col2.date_input("Ngày kết thúc", value=default_end, max_value=today)

    if start_date > end_date:
        st.error("Ngày bắt đầu phải nhỏ hơn hoặc bằng ngày kết thúc.")

    agoda_file = st.file_uploader("Chọn file Agoda (Excel)", type=["xlsx"], key="agoda")

    if agoda_file and (start_date <= end_date):
        try:
            # Đọc sheet mặc định theo tên trong file Agoda
            df = pd.read_excel(agoda_file, sheet_name="file tải xuống từ Agoda")

            # Chuyển đổi cột ngày trả phòng sang kiểu datetime
            df["Ngày trả phòng"] = df["Ngày trả phòng"].apply(parse_vietnamese_date)

            # Chuyển đổi cột doanh thu và số tiền bị trừ về kiểu số (bỏ dấu phẩy)
            df["Doanh thu thực"] = (
                df["Doanh thu thực"].astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
                .astype(float)
            )
            df["Số tiền bị trừ"] = (
                df["Số tiền bị trừ"].astype(str)
                .str.replace(",", "", regex=False)
                .str.strip()
                .astype(float)
            )

            # Lọc dữ liệu theo khoảng ngày trả phòng
            start_ts = pd.to_datetime(start_date)
            end_ts = pd.to_datetime(end_date)
            mask_date = (df["Ngày trả phòng"] >= start_ts) & (df["Ngày trả phòng"] <= end_ts)
            df_filtered = df.loc[mask_date].copy()

            # Giữ lại các dòng có doanh thu và số tiền bị trừ > 0
            df_filtered = df_filtered[
                (df_filtered["Doanh thu thực"] > 0) & (df_filtered["Số tiền bị trừ"] > 0)
            ]

            # Loại bỏ các cột Unnamed do thừa
            df_filtered = df_filtered.loc[
                :, ~df_filtered.columns.str.contains("^Unnamed")
            ]

            # Hiển thị bảng dữ liệu kết quả
            st.subheader("Bảng dữ liệu sau khi lọc")
            st.dataframe(df_filtered)

            # Xuất file excel cho người dùng tải về
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_filtered.to_excel(writer, index=False, sheet_name="Agoda")
            output.seek(0)
            file_name = f"Agoda_processed_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.xlsx"
            st.download_button(
                "📥 Tải file Agoda đã xử lý",
                data=output.getvalue(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Có lỗi khi xử lý file Agoda: {e}")
