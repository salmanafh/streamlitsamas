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

def createInvoice(nomor, tanggal, terima_dari = "", pekerjaan = "", jenis_muatan = "", harga= 0, volume=0, nomor_spk= "", bank = "", direktur = "MUHAMAD SOBARI", nomor_faktur = "", **kwargs):
    
    # Specify the path to the Excel file
    if kwargs:
        file_path = "KW-141 PLANTATION-PT.SBA 08 DESEMBER 2023.xlsx"
    elif volume == False:
        file_path = 'KW-PERMATA BANK.xlsx'
    else:
        file_path = 'KW-LTMPLB-2023 - Contoh.xlsx'
    workbook = openpyxl.load_workbook(file_path)
    # Select the active sheet
    sheet = workbook.active
    
    if kwargs:
        for _, values in kwargs.items():
            input_df = pd.DataFrame(values)
            
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == "=P13*11%":
                    format_rp = cell.number_format
                
        for i in range(0, len(input_df) - 1):
            sheet.insert_rows(idx = 12 + i*2, amount = 2)
            border_format_right = copy(sheet.cell(row=10, column=17).border)
            border_format_left = copy(sheet.cell(row=10, column=1).border)
            sheet[f"Q{12 + i*2 - 1}"].border = border_format_right
            sheet[f"Q{12 + i*2}"].border = border_format_right
            sheet[f"Q{12 + i*2 + 1}"].border = border_format_right
            sheet[f"A{12 + i*2 - 1}"].border = border_format_left
            sheet[f"A{12 + i*2}"].border = border_format_left
            sheet[f"A{12 + i*2 + 1}"].border = border_format_left

        for i in range(0, len(input_df)):
            sheet[f"H{10 + i*2}"] = input_df.iloc[i, 0]
            sheet[f"G{10 + i*2}"] = i + 1
            sheet[f"M{10 + i*2}"] = "No. SPK/PO :"
            sheet[f"N{10 + i*2}"] = input_df.iloc[i, 1]
            sheet[f"P{10 + i*2}"] = input_df.iloc[i, 2]
            sheet.cell(row = 10 + i*2, column = 16).number_format = format_rp
            sheet[f"I{11 + i*2}"] = input_df.iloc[i, 3]
            sheet[f"J{11 + i*2}"] = "Tonase :"
            sheet[f"K{11 + i*2}"] = input_df.iloc[i, 4]
            sheet[f"L{11 + i*2}"] = "Ha"
            koordinate_harga_total = sheet.cell(row = 13 + i*2, column = 16).coordinate
            koordinate_ppn = sheet.cell(row = 14 + i*2, column = 16).coordinate
            koordinate_bersih = sheet.cell(row = 15 + i*2, column = 16).coordinate
            koordinate_bersih_besar = sheet.cell(row = 17 + i*2, column = 2).coordinate

            harga_total = input_df["harga"].sum()
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

    # Write nomor
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "  PT. SELALU AMAN MANDIRI ABADI SUKSES":
                if jenis == "MJU":
                    cell.value = "  CV MAJU JAYA UTAMA"
                elif jenis == "MJU SU":
                    cell.value = "  PT MJU SUKSES UTAMA"
                else:
                    pass
    
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == ": PT.SELALU AMAN MANDIRI ABADI SUKSES":
                if jenis == "MJU":
                    cell.value = ": CV MAJU JAYA UTAMA"
                elif jenis == "MJU SU":
                    cell.value = ": PT MJU SUKSES UTAMA"
                else:
                    pass
                
    # Write nomor faktur
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "010.009.23.35623055":
                cell.value = nomor_faktur

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
    
    # Rewrite volume
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == 207970:
                cell.value = volume
    
    # Rewrite nomor spk
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "Nomor SPK : 065/SPK-PM/CRES/II/2023":
                if nomor_spk == "-":
                    cell.value = ""
                else:
                    cell.value = "Nomor SPK : " + nomor_spk

    # Rewrite nomor spk yang ada volume
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "Nomor SPK":
                if nomor_spk == "-":
                    sheet.delete_rows(13)
                else:
                    cell.value = "Nomor SPK : " + nomor_spk
    
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
    
    # Rewrite Direktur
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == "MUHAMAD SOBARI":
                cell.value = direktur
    
    # Rewrite nama Bank
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == ": MANDIRI":
                if bank == "Permata Bank":
                    cell.value = "PERMATA BANK"
                elif bank == "BRI":
                    cell.value = "BANK RAKYAT INDONESIA"
                else:
                    pass
    
    # Rewrite no rek Bank
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == ": 1120060000101":
                if bank == "Permata Bank":
                    cell.value = ": 971211164"
                elif bank == "BRI":
                    if jenis == "MJU":
                        cell.value = ": 034201001703300"
                    elif jenis == "MJU SU":
                        cell.value = ": 110401000551309"
                    else:
                        pass
    
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
jenis = st.selectbox('Perusahaan : ', ("Samas", "MJU", "MJU SU"), placeholder="Pilih Jenis")
nomor = st.text_input(label='Nomor Invoice: ', placeholder="Nomor", value='')
terima_dari = st.text_input(label='Telah Terima Dari: ', placeholder="Nama Perusahaan", value='')
pengadaan_plantasi_transportasi = st.selectbox('Jenis Pekerjaan: ', ("Pengadaan", "Jasa Transportasi", "Plantation"))
nomor_faktur = st.text_input(label='Nomor Faktur: ', placeholder="Nomor Faktur", value="")
if pengadaan_plantasi_transportasi == "Plantation":
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
            pekerjaan.append(st.text_input(label=f'Jenis Kegiatan {i}: ', placeholder=f"Keterangan Pekerjaan {i}", value=''))
        with col2:
            nomor_spk.append(st.text_input(label=f'Nomor SPK / PO {i}: ', placeholder=f"Nomor SPK / PO {i}", value=''))
        with col3:
            harga.append(st.number_input(label=f"Harga {i} (Rp): ", value=0))
        with col4:
            nomor_pekerjaan.append(st.text_input(label=f'Nomor Pekerjaan {i}: ', placeholder=f"Nomor Pekerjaan {i}", value=''))
        with col5:
            tonase.append(st.number_input(label=f'Tonase {i} (Ha) : ', value=0))
    
    dictionary_pekerjaan = {"pekerjaan": pekerjaan,
                            "nomor_spk": nomor_spk,
                            "harga": harga,
                            "nomor_pekerjaan": nomor_pekerjaan,
                            "tonase": tonase}

elif pengadaan_plantasi_transportasi == "Pengadaan":
    pekerjaan = st.text_input(label='Pekerjaan: ', placeholder="Keterangan Pekerjaan", value='')
    jenis_muatan = st.selectbox(f'Jenis Muatan: ', ("-","Arang Sekam", "Cocopeat"), placeholder="Pilih Jenis Muatan")
    if jenis_muatan == "-":
        vcol1, vcol2 = st.columns(2)
        with vcol1:
            harga = st.number_input(label=f"Harga (Rp): ", placeholder="Rp ", value=0)
        with vcol2:
            nomor_spk = st.text_input(label=f"Nomor SPK: ", placeholder="Nomor SPK", value="")
        volume = False
    else:
        vcol1, vcol2, vcol3 = st.columns(3)
        with vcol1:  
            volume = st.number_input(label=f'Volume Barang (Rp): ', placeholder="Rp", value=0)
        with vcol2:
            harga_barang = st.number_input(label=f'Harga Per Kilo (Rp): ', placeholder="Rp", value=0)
        with vcol3: 
            nomor_spk = st.number_input(label=f'Nomor SPK (Rp): ', placeholder="Rp", value=0)
        harga = volume * harga_barang
elif pengadaan_plantasi_transportasi == "Jasa Transportasi":
    pekerjaan = st.text_input(label='Pekerjaan: ', placeholder="Keterangan Pekerjaan", value='')
    jenis_muatan = st.selectbox(f'Jenis Muatan: ', ("-","Kelapa", "Ekspedisi", "Semen", "Kopra", "Pupuk", "Sawit", "Karnel", "Cangkang", "Batubara"), placeholder="Pilih Jenis Muatan")
    if jenis_muatan == "-":
        vcol1, vcol2 = st.columns(2)
        with vcol1:
            harga = st.number_input(label=f"Harga (Rp): ", placeholder="Rp ", value=0)
        with vcol2:
            nomor_spk = st.text_input(label=f"Nomor SPK: ", placeholder="Nomor SPK", value="")
        volume = False
    else:
        vcol1, vcol2, vcol3 = st.columns(3)
        with vcol1:  
            volume = st.number_input(label=f'Volume Barang (Rp): ', placeholder="Rp", value=0)
        with vcol2:
            harga_barang = st.number_input(label=f'Harga Per Kilo (Rp): ', placeholder="Rp", value=0)
        with vcol3: 
            nomor_spk = st.number_input(label=f'Nomor SPK (Rp): ', placeholder="Rp", value=0)
        harga = volume * harga_barang

tanggal = st.date_input('Tanggal: ', datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=7))))
direktur = st.selectbox('Direktur: ', ("Muhamad Sobari", "Mara Ispana", "Yossi Korpriyanto", "Zulkirom"), placeholder="Direktur")
bank = st.selectbox('Bank: ', ("BRI", "Mandiri", "PERMATA BANK"), placeholder="Bank")
tanggal = tanggal.strftime("%d %B %Y")

filename = "Invoice {nomor}.xlsx".format(nomor=nomor)
invoice = False
col1, col2 = st.columns(2)
with col1:
    if pengadaan_plantasi_transportasi == "Jasa Transportasi" or pengadaan_plantasi_transportasi == "Pengadaan":
        if st.button('Buat Invoice'):
            invoice = createInvoice(nomor, tanggal, terima_dari, pekerjaan, jenis_muatan, harga, volume, nomor_spk, bank = bank, direktur = direktur, nomor_faktur=nomor_faktur)
    else:
        if st.button('Buat Invoice'):
            invoice = createInvoice(nomor, tanggal, terima_dari, direktur = direktur, bank = bank, nomor_faktur=nomor_faktur, kwargs=dictionary_pekerjaan)
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
    if jenis == "Samas":
        st.image("logo_samas.png", use_column_width=True)
    elif jenis == "MJU":
        st.image("logo_mju.png", use_column_width=True)
    elif jenis == "MJU SU":
        st.image("logo_mju_su.jpeg", use_column_width=True)

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
    
