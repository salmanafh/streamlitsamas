import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import openpyxl
import streamlit as st
import pandas as pd
import xlsxwriter
from io import BytesIO
import datetime

st.set_page_config(
    page_title="Aplikasi Monitoring Samas",
    page_icon=":chart_with_upwards_trend:"
)

muatan = {
    "Cangkang Sawit": 100,
}

def createInvoice(nomor, terima_dari, npwp, untuk, pekerjaan, jenis_muatan, volume, tanggal, atas_nama):
    # Specify the path to the Excel file
    file_path = 'KW-LTMPLB-2023 - Contoh.xlsx'
    workbook = openpyxl.load_workbook(file_path)
    # Select the active sheet
    sheet = workbook.active

    # Rewrite nomor
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "119/SAMAS-KW/LTM/VIII/2023":
                cell.value = nomor

    # Rewrite terima_dari
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "PT. ALAM MULTI MEGA":
                cell.value = terima_dari
    
    # Rewrite npwp
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "90.422.411.0-307.000":
                cell.value = npwp

    # Rewrite untuk
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "Upah Kontraktor a/n PT.SELALU AMAN MANDIRI ABADI SUKSES":
                cell.value = untuk
                
    # Rewrite pekerjaan
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "Biaya Jasa Angkutan Darat Dari PT SAP ke Jetty LKS ":
                cell.value = pekerjaan
    
    # Rewrite jenis_muatan
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "Cangkang Sawit":
                cell.value = jenis_muatan
    
    # Rewrite volume
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == 207970:
                cell.value = volume
    
    harga_jenis_muatan = muatan[jenis_muatan]
    harga_total = volume * harga_jenis_muatan
    # Rewrite Harga Total
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == 20797000:
                cell.value = harga_total
    
    ppn = (harga_total * 11) / 100
    # Rewrite PPN
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "=N14*11%":
                cell.value = ppn
    
    bersih = harga_total + ppn 
    # Rewrite bersih
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "=N14+N15":
                cell.value = bersih
    
    # Rewrite bersih yang besar
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "=N16":
                cell.value = bersih
    
    # Rewrite tanggal
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "Palembang, 11 Agustus 2023":
                cell.value = tempat + ", " + tanggal
                
    # Rewrite atas_nama
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "Muhamad Sobari":
                cell.value = atas_nama
    

    # Save the workbook
    workbook.save('invoice.xlsx')
 
st.subheader("Buat Laporan Otomatis")
uploaded_file = st.file_uploader('Upload File Excel')
 
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.dataframe(df)

st.button('Buat Laporan')

# Mendapatkan tanggal dan waktu saat ini
sekarang = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=7)))
sekarang = sekarang.strftime("%d-%B-%Y")
st.subheader("Buat Invoice")
npwp = st.text_input(label='NPWP: ', placeholder="NPWP", value='')
nomor = st.text_input(label='Nomor Invoice: ', placeholder="Nomor", value='')
terima_dari = st.text_input(label='Telah Terima Dari: ', placeholder="Nama Perusahaan", value='')
untuk = st.text_input(label='Untuk Pembayaran: ', placeholder="Keterangan Pembayaran", value='')
pekerjaan = st.text_input(label='Pekerjaan: ', placeholder="Keterangan Pekerjaan", value='')
jenis_muatan = st.selectbox('Jenis Muatan: ', ("Cangkang Sawit", "Cocopeat", "Kopra", "Jagung", "Kelapa", "Sekam", "Pupuk", "Bibit"), placeholder="Pilih Jenis Muatan")
volume = st.text_input(label='Volume (Kg): ', placeholder="Kg")
tempat = st.selectbox('Tempat: ', ("Palembang", "Jambi"))
tanggal = st.text_input(label='Tanggal: ', placeholder=f"Sekarang: {sekarang}", value='')
atas_nama = st.text_input(label='Atas Nama: ', placeholder="Nama", value='')

filename = "Invoice {nomor}.xlsx".format(nomor=nomor)
invoice = "invoice.xlsx"
col1, col2 = st.columns(2)
with col1:
    if st.button('Buat Invoice'):
        invoice = createInvoice(nomor, terima_dari, npwp, untuk, pekerjaan, jenis_muatan, int(volume), tanggal, atas_nama)
with col2:
    with open("invoice.xlsx", "rb") as f:
        st.download_button(label = 'Download Invoice',
                            data=f.read(),
                            file_name=filename,
                            mime='application/vnd.ms-excel',)
st.subheader("Monitoring Cash In - Out")
# Load data from Excel file
laporan_cash_in_out = pd.ExcelFile("Contoh Laporan Cash In -Out Januari 23.xlsx")
cash_df = pd.read_excel(laporan_cash_in_out).fillna(0)
min_date = cash_df.Tanggal.min()
max_date = cash_df.Tanggal.max()
with st.sidebar:
    date_range = st.date_input(
        label="Select Date Range",
        min_value=min_date,
        max_value=max_date,
        value=[min_date, max_date],
        format="YYYY.MM.DD",
    )
    if len(date_range) == 1:
        main_df = cash_df
    else:
        start_date, end_date = date_range
        main_df = cash_df[(cash_df.Tanggal >= str(start_date)) & (cash_df.Tanggal <= str(end_date))]

ss_01 = main_df.loc[(main_df["Project"] == "SS-01")]
ss_02 = main_df.loc[(main_df["Project"] == "SS-02")]
ss_03 = main_df.loc[(main_df["Project"] == "SS-03")]
ss_05 = main_df.loc[(main_df["Project"] == "SS-05")]
ss_06 = main_df.loc[(main_df["Project"] == "SS-06")]
ss_09 = main_df.loc[(main_df["Project"] == "SS-09")]
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["SS-09", "SS-02", "SS-03", "SS-05", "SS-06", "SS-01"])

with tab6:
    # Create bar plot
    fig, ax = plt.subplots()
    ax.pie(ss_01[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

    # Set plot title and labels
    ax.set_title("SS - 09 Cash In - Out")

    # Display the plot in Streamlit
    st.pyplot(fig)
    st.write(ss_01.tail(5))

with tab2:
    # Create bar plot
    fig, ax = plt.subplots()
    ax.pie(ss_02[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)
    
    # Set plot title and labels
    ax.set_title("SS - 02 Cash In - Out")
    
    # Display the Plot in Streamlit
    st.pyplot(fig)
    st.write(ss_02.tail(5))
    
with tab3:
    fig, ax = plt.subplots()
    ax.pie(ss_03[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

    # Set plot title and labels
    ax.set_title("SS - 03 Cash In - Out")

    # Display the plot in Streamlit
    st.pyplot(fig)
    st.write(ss_03.tail(5))

with tab4:
    fig, ax = plt.subplots()
    ax.pie(ss_05[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

    # Set plot title and labels
    ax.set_title("SS - 05 Cash In - Out")

    # Display the plot in Streamlit
    st.pyplot(fig)
    st.write(ss_05.tail(5))
    
with tab5:
    fig, ax = plt.subplots()
    ax.pie(ss_06[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

    # Set plot title and labels
    ax.set_title("SS - 06 Cash In - Out")

    # Display the plot in Streamlit
    st.pyplot(fig)
    st.write(ss_06.tail(5))
    
with tab1:
    fig, ax = plt.subplots()
    ax.pie(ss_09[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

    # Set plot title and labels
    ax.set_title("SS - 01 Cash In - Out")

    # Display the plot in Streamlit
    st.pyplot(fig)
    st.write(ss_09.tail(5))
    
