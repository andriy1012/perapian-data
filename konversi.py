import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO


st.title("🔢 Pemisahan Data Customer dengan Invoice by Andriy")
st.write("Aplikasi untuk memisahkan informasi customer dan generate invoice")

# Upload file
uploaded_file = st.file_uploader("Unggah file Excel atau CSV", type=['xlsx', 'csv'])

if uploaded_file is not None:
    try:
        # Baca file berdasarkan ekstensi
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file)
        else:
            df = pd.read_csv(uploaded_file)

        # Tampilkan data asli
        st.subheader("Data Asli")
        st.write(df.head())

        # Dapatkan nama kolom aktual
        actual_columns = list(df.columns)

        # Pemetaan kolom oleh user
        st.subheader("Pemetaan Kolom")
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        with col1:
            tanggal_col = st.selectbox("Pilih kolom Tanggal*", actual_columns, index=0)
        with col2:
            nama_produk_col = st.selectbox("Pilih kolom Nama Produk*", actual_columns, index=1)
        with col3:
            merek_col = st.selectbox("Pilih kolom Merek*", actual_columns, index=2)
        with col4:
            qty_col = st.selectbox("Pilih kolom Quantity*", actual_columns, index=3)
        with col5:
            cust_gabungan_col = st.selectbox("Pilih kolom Customer Gabungan*", actual_columns, index=4)
        with col6:
            cust_asli_col = st.selectbox("Pilih kolom Customer Asli*", actual_columns, index=5)

        if st.button("Proses Data dan Generate Invoice"):
            with st.spinner('Memproses data...'):
                # Rename kolom
                rename_mapping = {
                    tanggal_col: 'Tanggal',
                    nama_produk_col: 'Nama_Produk',
                    merek_col: 'Merek',
                    qty_col: 'Qty',
                    cust_gabungan_col: 'Customer_Gabungan',
                    cust_asli_col: 'Customer_Asli'
                }
                df_renamed = df.rename(columns=rename_mapping)

                # Simpan nilai asli Tanggal sebagai string tanpa ".0"
                df_renamed['Tanggal_Raw'] = df_renamed['Tanggal'].apply(lambda x: str(int(x)) if pd.notnull(x) else '')

                # Format ke datetime dan ke YYYY-MM-DD
                try:
                    df_renamed['Tanggal'] = pd.to_datetime(
                        df_renamed['Tanggal_Raw'], format='%Y%m%d', errors='coerce'
                    )
                    df_renamed['Tanggal_Formatted'] = df_renamed['Tanggal'].dt.strftime('%Y-%m-%d')
                except Exception as e:
                    st.error(f"Error konversi tanggal: {str(e)}")
                    df_renamed['Tanggal_Formatted'] = df_renamed['Tanggal_Raw']

                # Bersihkan data dari nilai kosong
                required_columns = ['Tanggal', 'Nama_Produk', 'Merek', 'Qty', 'Customer_Gabungan', 'Customer_Asli']
                df_cleaned = df_renamed.dropna(subset=required_columns).copy()

                # Proses alokasi customer
                expanded_rows = []
                for index, row in df_cleaned.iterrows():
                    tanggal = row.get('Tanggal_Formatted', '')
                    tanggal_raw = row.get('Tanggal_Raw', '')
                    product_name = row.get('Nama_Produk', '')
                    merek = row.get('Merek', '')
                    total_qty = row.get('Qty', 0)
                    cust_gabungan = str(row.get('Customer_Gabungan', ''))
                    cust_asli = str(row.get('Customer_Asli', ''))

                    # Deteksi pola alokasi jumlah dan customer
                    pattern = r'(\d+)\s*KG\s*[^,]*?([A-Z][A-Z\s]+?)(?=,|$|SISA)'
                    matches = re.findall(pattern, cust_gabungan)

                    if matches:
                        total_allocated = 0
                        customer_allocations = []
                        for qty_str, customer in matches:
                            qty = int(qty_str)
                            total_allocated += qty
                            customer_allocations.append({'customer': customer.strip(), 'qty': qty})

                        remaining_qty = total_qty - total_allocated
                        if remaining_qty > 0:
                            sisa_pattern = r'SISA.*?([A-Z][A-Z\s]+?)(?:\s|$)'
                            sisa_match = re.search(sisa_pattern, cust_gabungan)
                            if sisa_match:
                                sisa_customer = sisa_match.group(1).strip()
                                customer_allocations.append({'customer': sisa_customer, 'qty': remaining_qty})

                        for allocation in customer_allocations:
                            expanded_rows.append({
                                'Tanggal': tanggal,
                                'Tanggal_Raw': tanggal_raw,
                                'Nama_Produk': product_name,
                                'Merek': merek,
                                'Quantity': allocation['qty'],
                                'Customer': allocation['customer'],
                                'Customer_Gabungan': cust_gabungan,
                                'Customer_Asli': cust_asli
                            })
                    else:
                        expanded_rows.append({
                            'Tanggal': tanggal,
                            'Tanggal_Raw': tanggal_raw,
                            'Nama_Produk': product_name,
                            'Merek': merek,
                            'Quantity': total_qty,
                            'Customer': cust_gabungan,
                            'Customer_Gabungan': cust_gabungan,
                            'Customer_Asli': cust_asli
                        })

                result_df = pd.DataFrame(expanded_rows)

                # ✅ Generate invoice dengan format INV-YYYYMMDD-001
                invoice_counter = 1
                invoice_map = {}
                invoice_numbers = []

                for index, row in result_df.iterrows():
                    key = f"{row['Customer_Asli']}_{row['Tanggal_Raw']}"
                    if key not in invoice_map:
                        try:
                            dt = datetime.strptime(row['Tanggal_Raw'], '%Y%m%d')
                            inv_date = dt.strftime('%Y%m%d')
                        except:
                            inv_date = datetime.now().strftime('%Y%m%d')
                        invoice_map[key] = f"INV-{inv_date}-{invoice_counter:03d}"
                        invoice_counter += 1
                    invoice_numbers.append(invoice_map[key])

                result_df['Invoice'] = invoice_numbers

                # Urutkan kolom
                result_df = result_df[['Invoice', 'Tanggal', 'Nama_Produk', 'Merek',
                                       'Quantity', 'Customer', 'Customer_Asli', 'Customer_Gabungan', 'Tanggal_Raw']]

                # Tampilkan hasil
                st.subheader("Hasil Pemrosesan dengan Invoice")
                st.write(result_df)

                # Ringkasan invoice
                st.subheader("Ringkasan per Invoice")
                summary_df = result_df.groupby(['Invoice', 'Tanggal', 'Customer_Asli']).agg({
                    'Quantity': 'sum',
                    'Nama_Produk': lambda x: ', '.join(x.unique()),
                    'Merek': lambda x: ', '.join(x.unique()),
                    'Customer': lambda x: ', '.join(x.unique())
                }).reset_index()
                st.write(summary_df)

                # Fungsi bantu untuk mengonversi DataFrame ke bytes Excel
                def to_excel(df):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Sheet1')
                    return output.getvalue()

                # Konversi DataFrame ke format Excel (bytes)
                xlsx = to_excel(result_df)
                xlsx_summary = to_excel(summary_df)

                # Tombol unduh
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        "Unduh Detail Lengkap (xlsx)",
                        data=xlsx,
                        file_name='detail_transaksi.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                with col2:
                    st.download_button(
                        "Unduh Ringkasan Invoice (xlsx)",
                        data=xlsx_summary,
                        file_name='ringkasan_invoice.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

                st.success("Proses selesai! Invoice telah digenerate berdasarkan Customer Asli.")

    except Exception as e:
        st.error(f"Terjadi kesalahan: {str(e)}")
