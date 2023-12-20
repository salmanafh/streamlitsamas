import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import streamlit as st
import datetime
import os, datetime, openpyxl, base64
from copy import copy

st.set_page_config(
    page_title="Aplikasi Monitoring Samas",
    page_icon=":chart_with_upwards_trend:",
    layout="wide",
)

def createInvoice(nomor, tanggal, terima_dari = "", pekerjaan = "", jenis_muatan = "", harga= 0, nomor_volume= 0, nomor_faktur = "", **kwargs):
    
    # Specify the path to the Excel file
    if kwargs:
        file_path = "KW-141 PLANTATION-PT.SBA 08 DESEMBER 2023.xlsx"
    elif type(nomor_volume) == str:
        file_path = 'KW-PERMATA BANK.xlsx'
    else:
        file_path = 'KW-LTMPLB-2023 - Contoh.xlsx'
    workbook = openpyxl.load_workbook(file_path)
    # Select the active sheet
    sheet = workbook.active
    
    if kwargs:
        sheet = workbook.active
        for _, values in kwargs.items():
            mjo_input_df = pd.DataFrame(values)
            
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == "=P13*11%":
                    format_rp = cell.number_format
            
        for i in range(0, len(mjo_input_df) - 1):
            sheet.insert_rows(idx = 12 + i*2, amount = 2)
            border_format_right = copy(sheet.cell(row=10, column=17).border)
            border_format_left = copy(sheet.cell(row=10, column=1).border)
            sheet[f"Q{12 + i*2 - 1}"].border = border_format_right
            sheet[f"Q{12 + i*2}"].border = border_format_right
            sheet[f"Q{12 + i*2 + 1}"].border = border_format_right
            sheet[f"A{12 + i*2 - 1}"].border = border_format_left
            sheet[f"A{12 + i*2}"].border = border_format_left
            sheet[f"A{12 + i*2 + 1}"].border = border_format_left
        
        for i in range(0, len(mjo_input_df)):
            sheet[f"H{10 + i*2}"] = mjo_input_df.iloc[i, 0]
            sheet[f"G{10 + i*2}"] = i + 1
            sheet[f"M{10 + i*2}"] = "No. SPK/PO :"
            sheet[f"N{10 + i*2}"] = mjo_input_df.iloc[i, 1]
            sheet[f"P{10 + i*2}"] = mjo_input_df.iloc[i, 2]
            sheet.cell(row = 10 + i*2, column = 16).number_format = format_rp
            sheet[f"I{11 + i*2}"] = mjo_input_df.iloc[i, 3]
            sheet[f"J{11 + i*2}"] = "Tonase :"
            sheet[f"K{11 + i*2}"] = mjo_input_df.iloc[i, 4]
            sheet[f"L{11 + i*2}"] = "Ha"
            koordinate_harga_total = sheet.cell(row = 13 + i*2, column = 16).coordinate
            koordinate_ppn = sheet.cell(row = 14 + i*2, column = 16).coordinate
            koordinate_bersih = sheet.cell(row = 15 + i*2, column = 16).coordinate
            koordinate_bersih_besar = sheet.cell(row = 17 + i*2, column = 2).coordinate
            
        harga_total = mjo_input_df["harga"].sum()
        # write harga total
        sheet[koordinate_harga_total] = harga_total
        
        # write ppn
        ppn = (harga_total * 11) / 100
        sheet[koordinate_ppn] = ppn
            
        # Write harga bersih
        harga_bersih = harga_total + ppn
        sheet[koordinate_bersih] = harga_bersih
            
        # Write nomor faktur
        sheet[koordinate_bersih_besar] = harga_bersih
        
        # Write harga bersih yang besar
        sheet["E6"] = sheet.cell(row=6, column=5).value.replace("B17", f"{koordinate_bersih_besar}")

    # Write nomor
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "119/SAMAS-KW/LTM/VIII/2023":
                cell.value = nomor

    # Rewrite terima_dari
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "PT. ALAM MULTI MEGA":
                cell.value = terima_dari
                
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
    
    # Rewrite nomor spk atau volume
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == 207970:
                cell.value = nomor_volume
            elif cell.value == "Nomor SPK : 065/SPK-PM/CRES/II/2023":
                cell.value = "Nomor SPK : " + nomor_volume
    
    # Rewrite Harga Total
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == 20797000:
                cell.value = harga
    
    # Rewrite PPN
    ppn = harga * 11 / 100
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "=N14*11%":
                cell.value = ppn
    
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "=S1":
                cell.value = jenis_muatan
    
    bersih = harga + ppn 
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
                cell.value = "Palembang, " + tanggal
    
    # Save the workbook
    sheet.sheet_view.showGridLines = False
    workbook.save('invoice.xlsx')
    return True
 
# with open("C:/Users/salma/OneDrive/Documents/KW-PERMATA BANK.pdf", "rb") as f:
#     base64_pdf = base64.b64encode(f.read()).decode('utf-8')

#     # Embedding PDF in HTML
#     pdf_display = F'<iframe src="data:application/pdf;base64,{base64_pdf}" style="width: 100%; height: 100vh;" type="application/pdf"></iframe>'

#     # Displaying File
#     st.markdown(pdf_display, unsafe_allow_html=True)

st.header("Buat Invoice")
# Create a list of unique values in the column 'Project'
jenis = st.selectbox('Jenis: ', ("Samas", "MJU"), placeholder="Pilih Jenis")
nomor = st.text_input(label='Nomor Invoice: ', placeholder="Nomor", value='')
terima_dari = st.text_input(label='Telah Terima Dari: ', placeholder="Nama Perusahaan", value='')
if jenis == "MJO":
    nomor_faktur = st.text_input(label='Nomor Faktur: ', placeholder="Nomor Faktur", value="")
    jumlah_pekerjaan = st.number_input(label='Jumlah Pekerjaan: ', placeholder="Jumlah Pekerjaan", value=0)
    list_pekerjaan = ["pekerjaan", "nomor_spk", "harga", "nomor_pekerjaann", "tonase"]
    pekerjaan = []
    nomor_spk = []
    harga = []
    nomor_pekerjaan = []
    tonase = []
    for i in range(1, jumlah_pekerjaan + 1):
        col1, col2 = st.columns(2)
        col3, col4, col5 = st.columns(3)
        with col1:
            pekerjaan.append(st.text_input(label=f'Pekerjaan {i}: ', placeholder=f"Keterangan Pekerjaan {i}", value=''))
        with col2:
            nomor_spk.append(st.text_input(label=f'Nomor SPK / PO {i}: ', placeholder=f"Nomor SPK / PO {i}", value=''))
        with col3:
            harga.append(st.number_input(label=f"Harga {i} (Rp): ", value=0))
        with col4:
            nomor_pekerjaan.append(st.text_input(label=f'Nomor Pekerjaan {i}: ', placeholder=f"Nomor Pekerjaan {i}", value=''))
        with col5:
            tonase.append(st.number_input(label=f'Tonase {i} (Ha) : ', value=0))
    
    dictionary_pekerjaan = {"Pekerjaan": pekerjaan,
                            "nomor_spk": nomor_spk,
                            "harga": harga,
                            "nomor_pekerjaan": nomor_pekerjaan,
                            "tonase": tonase}
else:
    pekerjaan = st.text_input(label='Pekerjaan: ', placeholder="Keterangan Pekerjaan", value='')
    jenis_muatan = st.selectbox('Jenis Muatan: ', ("-","Cangkang Sawit", "Cocopeat", "Kopra", "Jagung", "Kelapa", "Sekam", "Pupuk", "Bibit"), placeholder="Pilih Jenis Muatan")
    if jenis_muatan == "-":
        harga = st.number_input(label="Harga (Rp): ", placeholder="Rp ", value=0)
        nomor_volume = st.text_input(label="Nomor SPK: ", placeholder="Nomor SPK", value="")
        volume = False
    else:
        hcol1, hcol2 = st.columns(2)
        with hcol1:
            nomor_volume = st.number_input(label='Volume (Kg): ', placeholder="Kg", value=0)
        with hcol2: 
            harga_barang = st.number_input(label='Harga Per Kilo (Rp): ', placeholder="Rp", value=0)
        harga = nomor_volume * harga_barang
tanggal = st.date_input('Tanggal: ', datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=7))))
tanggal = tanggal.strftime("%d %B %Y")

filename = "Invoice {nomor}.xlsx".format(nomor=nomor)
invoice = False
col1, col2 = st.columns(2)
with col1:
    if jenis == "MJO":
        if st.button('Buat Invoice'):
            invoice = createInvoice(nomor, terima_dari, tanggal, kwargs=dictionary_pekerjaan)
    else:
        if st.button('Buat Invoice'):
            invoice = createInvoice(nomor, tanggal, terima_dari, pekerjaan, jenis_muatan, harga, nomor_volume)
with col2:
    with open("invoice.xlsx", "rb") as f:
        st.download_button(label = 'Download Invoice',
                            data=f.read(),
                            file_name=filename,
                            disabled=False if invoice else True,
                            mime='application/vnd.ms-excel',)

st.header("Dashboard Cash In - Out")
uploaded_file = st.file_uploader('Upload File Excel')
 
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.dataframe(df)


# Load data from Excel file
laporan_cash_in_out = pd.ExcelFile("Contoh Laporan Cash In -Out Januari 23.xlsx")
cash_df = pd.read_excel(laporan_cash_in_out).fillna(0)
min_date = cash_df.Tanggal.min()
max_date = cash_df.Tanggal.max()
with st.sidebar:
    st.image("logo_samas-removebg-preview.png", use_column_width=True)    


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
    main_df["Tanggal"] = cash_df["Tanggal"].dt.strftime("%d %B %Y")
    
ss_01 = main_df.loc[(main_df["Project"] == "SS-01")]
ss_02 = main_df.loc[(main_df["Project"] == "SS-02")]
ss_03 = main_df.loc[(main_df["Project"] == "SS-03")]
ss_05 = main_df.loc[(main_df["Project"] == "SS-05")]
ss_06 = main_df.loc[(main_df["Project"] == "SS-06")]
ss_09 = main_df.loc[(main_df["Project"] == "SS-09")]

# fig, ax = plt.subplots()
# ax.pie(ss_01[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

# Set plot title and labels
# ax.set_title("SS - 09 Cash In - Out")
# tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["SS-09", "SS-02", "SS-03", "SS-05", "SS-06", "SS-01"])

# with tab6:
#     # Create bar plot
#     fig, ax = plt.subplots()
#     ax.pie(ss_01[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

#     # Set plot title and labels
#     ax.set_title("SS - 09 Cash In - Out")

#     # Display the plot in Streamlit
#     st.pyplot(fig)
#     st.write(ss_01.tail(5))

# with tab2:
#     # Create bar plot
#     fig, ax = plt.subplots()
#     ax.pie(ss_02[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)
    
#     # Set plot title and labels
#     ax.set_title("SS - 02 Cash In - Out")
    
#     # Display the Plot in Streamlit
#     st.pyplot(fig)
#     st.write(ss_02.tail(5))
    
# with tab3:
#     fig, ax = plt.subplots()
#     ax.pie(ss_03[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

#     # Set plot title and labels
#     ax.set_title("SS - 03 Cash In - Out")

#     # Display the plot in Streamlit
#     st.pyplot(fig)
#     st.write(ss_03.tail(5))

# with tab4:
#     fig, ax = plt.subplots()
#     ax.pie(ss_05[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

#     # Set plot title and labels
#     ax.set_title("SS - 05 Cash In - Out")

#     # Display the plot in Streamlit
#     st.pyplot(fig)
#     st.write(ss_05.tail(5))
    
# with tab5:
#     fig, ax = plt.subplots()
#     ax.pie(ss_06[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

#     # Set plot title and labels
#     ax.set_title("SS - 06 Cash In - Out")

#     # Display the plot in Streamlit
#     st.pyplot(fig)
#     st.write(ss_06.tail(5))
    
# with tab1:
#     fig, ax = plt.subplots()
#     ax.pie(ss_09[["Kredit", "Debit"]].sum(), labels=["Kredit", "Debit"], autopct="%1.1f%%", startangle=90)

#     # Set plot title and labels
#     ax.set_title("SS - 01 Cash In - Out")

#     # Display the plot in Streamlit
#     st.pyplot(fig)
#     st.write(ss_09.tail(5))
    
