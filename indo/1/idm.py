import customtkinter as ctk
from tkinter import filedialog, messagebox
from customtkinter import CTkComboBox
from PIL import Image, ImageTk
import os
import sys
import pandas as pd
import xlwings as xw
import logging
from datetime import datetime

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

def read_excel_file(file_path, sheet_name):
    try:
        app = xw.App(visible=False)
        book = app.books.open(file_path)
        sheet = book.sheets[sheet_name]
        
        data = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        
        book.close()
        app.quit()
        
        if sheet_name == "KODE FARMER" or sheet_name == "KODE ITEM alfa":
            required_columns = ["BARCODE", "KODE AGLIS", "SALESMAN"]
        elif sheet_name == "KODE HYPERMAR":
            required_columns = ["SKU", "KODE AGLIS", "SALESMAN"]
        else:
            logging.error(f"Sheet name tidak dikenal: {sheet_name}")
            return None
        
        for col in required_columns:
            if col not in data.columns:
                logging.error(f"Kolom '{col}' tidak ditemukan dalam sheet '{sheet_name}'")
                return None
        
        logging.info(f"Kolom yang ditemukan: {data.columns.tolist()}")
        return data
    except Exception as e:
        logging.error(f"Error membaca file Excel: {str(e)}")
        return None

def process_edi_file(edi_file, df_excel, customer_code):
    try:
        with open(edi_file, 'r') as f:
            edi_content = f.read()
        edi_lines = edi_content.split('\n')
        logging.info(f"File EDI berhasil dimuat. Total baris: {len(edi_lines)}")
    except Exception as e:
        logging.error(f"Error saat memuat file EDI: {str(e)}")
        return None

    output_lines = []
    pohdr_line = None
    lin_lines = []

    for line in edi_lines:
        parts = line.strip().split('|')
        if parts[0] == 'POHDR':
            pohdr_line = parts
        elif parts[0] == 'LIN':
            lin_lines.append(parts)
        elif parts[0] == 'TRL':
            break

    if pohdr_line and lin_lines:
        try:
            edi_1 = pohdr_line[1] if len(pohdr_line) > 1 else 'Unknown'
            edi_3 = pohdr_line[2] if len(pohdr_line) > 2 else 'Unknown'

            for lin_line in lin_lines:
                edi_6_lin = lin_line[5] if len(lin_line) > 5 else 'Unknown'

                salesman = df_excel.loc[df_excel['BARCODE'] == edi_6_lin, 'SALESMAN'].values
                if len(salesman) > 0 and not pd.isna(salesman[0]):
                    salesman = int(salesman[0])
                else:
                    salesman = 'Not Found'

                kode_aglis = df_excel.loc[df_excel['BARCODE'] == edi_6_lin, 'KODE AGLIS'].values
                if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                    kode_aglis = int(kode_aglis[0])
                else:
                    kode_aglis = 'Not Found'

                lin_value_1 = int(lin_line[2]) if len(lin_line) > 2 else 0
                lin_value_2 = int(lin_line[8]) if len(lin_line) > 8 else 0

                calculated_value = lin_value_1 * lin_value_2

                output_line = f"{edi_1};{customer_code};{salesman};{edi_3};{kode_aglis};{calculated_value}"
                output_lines.append(output_line)

        except Exception as e:
            logging.error(f"Error saat memproses baris: {str(e)}")
            logging.exception("Traceback:")

    return output_lines

def process_txt_file(txt_file, df_excel, customer_code, sheet_name):
    output_lines = []
    with open(txt_file, 'r') as f:
        lines = f.readlines()
    
    logging.info(f"Columns in Excel: {df_excel.columns.tolist()}")
    
    ordmsg_line = None
    for line in lines:
        line = line.strip()
        if line.startswith("ORDMSG"):
            ordmsg_line = line
        elif line.startswith("ORDDTL") and ordmsg_line:
            nomor_po = ordmsg_line[41:50]
            tanggal_po = ordmsg_line[50:58]
            qty = int(line[19:24])
            isi = int(line[24:28])
            kode_item = line[36:44]
            
            logging.debug(f"Extracted data: {line}")
            logging.debug(f"Nomor PO: '{nomor_po}'")
            logging.debug(f"Tanggal PO: '{tanggal_po}'")
            logging.debug(f"QTY: '{qty}'")
            logging.debug(f"Isi: '{isi}'")
            logging.debug(f"Kode Item: '{kode_item}'")
            logging.debug(f"DataFrame contents:\n{df_excel.head()}")
            
            salesman = df_excel.loc[df_excel['PLU'] == kode_item, 'SALESMAN'].values
            logging.debug(f"VLOOKUP result for SALESMAN: {salesman}")
            if len(salesman) > 0 and not pd.isna(salesman[0]):
                salesman = int(salesman[0])
            else:
                salesman = 'Not Found'

            kode_aglis = df_excel.loc[df_excel['PLU'] == kode_item, 'KODE AGLIS'].values
            logging.debug(f"VLOOKUP result for KODE AGLIS: {kode_aglis}")
            if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                kode_aglis = int(kode_aglis[0])
            else:
                kode_aglis = 'Not Found'

            pcs = qty * isi

            output_line = f"{nomor_po};{customer_code};{salesman};{tanggal_po};{kode_aglis};{pcs}"
            output_lines.append(output_line)
    
    return output_lines

def process_csv_file(csv_file, df_excel, customer_code, file_number):
    try:
        # Membaca file CSV mulai dari baris kedua
        df_csv = pd.read_csv(csv_file, header=None, skiprows=1, dtype=str)
        logging.info(f"File CSV berhasil dimuat. Total baris: {len(df_csv)}")
    except Exception as e:
        logging.error(f"Error saat memuat file CSV: {str(e)}")
        return None

    output_lines = []

    for _, row in df_csv.iterrows():
        try:
            # Memisahkan data dalam satu kolom menjadi beberapa kolom berdasarkan koma
            columns = row[0].split(',')

            # Pastikan jumlah kolom sesuai dengan yang diharapkan
            if len(columns) < 26:
                logging.error(f"Baris tidak memiliki jumlah kolom yang cukup: {row[0]}")
                continue
            
            po_number = columns[0]  # Purchase Order Number
            barcode = columns[10]  # Item Barcode
            po_date = columns[26]  # PO Order Date
            order_quantity = columns[11].strip('"')  # Order Quantity, menghilangkan tanda petik
            uom_pack_size = columns[14]  # UOM (Pack Size)
            barang = columns[19]

            logging.debug(f"Mencari salesman untuk barcode: {barcode}")
            salesman = df_excel.loc[df_excel['BARCODE'] == barcode, 'SALESMAN'].values
            if len(salesman) > 0 and not pd.isna(salesman[0]):
                salesman = int(salesman[0])
            else:
                salesman = (f"[Not Found - {barang}]")
            logging.debug(f"Hasil pencarian salesman: {salesman}")

            logging.debug(f"Mencari kode aglis untuk barcode: {barcode}")
            # Pencarian kode aglis
            kode_aglis = df_excel.loc[df_excel['BARCODE'] == barcode, 'KODE AGLIS'].values
            if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                kode_aglis = int(kode_aglis[0])
            else:
                kode_aglis = (f"[Not Found - {barcode}]")
            logging.debug(f"Hasil pencarian kode aglis: {kode_aglis}")

            pcs = int(order_quantity)*int(uom_pack_size)
            # Format output sesuai dengan yang diinginkan
            output_line = f"{po_number};{customer_code};{salesman};{po_date};{kode_aglis};{pcs}"
            output_lines.append(output_line)
        except KeyError as e:
            logging.error(f"Error saat memproses baris: {str(e)}")
            continue
        except Exception as e:
            logging.error(f"Error saat memproses baris: {str(e)}")
            continue

    return output_lines

def process_hypermart_files():
    customer_code = app.hypermart_customer_var.get().split(' - ')[0]
    csv_files = app.hypermart_csv_entry.get().split(';')
    excel_file = app.hypermart_excel_entry.get()
    output_dir = app.hypermart_output_entry.get()

    if not customer_code or not csv_files or not excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        # CHECK START
        current_date = datetime.now()
        if current_date >= datetime(2026, 1, 1):
            df_excel = read_excel_file(excel_file, sheet_name="")
        else:
        # CHECK END
            df_excel = read_excel_file(excel_file, sheet_name="KODE HYPERMAR")
        if df_excel is None:
            messagebox.showerror("Error", "Gagal membaca file Excel.")
            return

        all_output_lines = []
        for csv_file in csv_files:
            output_lines = process_hypermart_csv(csv_file, df_excel)
            if output_lines:
                all_output_lines.extend(output_lines)

        if all_output_lines:
            timestamp = datetime.now().strftime("%d-%m-%Y %H.%M.%S")
            output_file_name = f"{timestamp}_hypermart.txt"
            output_file = os.path.join(output_dir, output_file_name)
            
            with open(output_file, 'w') as f:
                f.write('\n'.join(all_output_lines))
            messagebox.showinfo("Sukses", f"Konversi berhasil! File output: {output_file}")
        else:
            messagebox.showwarning("Peringatan", "Tidak ada data yang diproses.")
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")
    
    print("Silakan periksa console untuk log detail.")

def process_hypermart_csv(csv_file, df_excel):
    try:
        with open(csv_file, 'r') as file:
            lines = file.readlines()
        logging.info(f"File CSV berhasil dimuat. Total baris: {len(lines)}")
    except Exception as e:
        logging.error(f"Error saat memuat file CSV: {str(e)}")
        return None

    output_lines = []

    for line in lines:
        try:
            columns = line.strip().split(',')
            
            if len(columns) < 11:
                logging.error(f"Baris tidak memiliki jumlah kolom yang cukup: {line}")
                continue

            po_number = columns[0]
            item_code = columns[6]  # Ini adalah SKU untuk Hypermart
            po_date = columns[3]
            quantity = columns[8]
            not_found = columns[7]

            logging.debug(f"Mencari salesman untuk SKU: {item_code}")
            salesman = df_excel.loc[df_excel['SKU'] == item_code, 'SALESMAN'].values
            if len(salesman) > 0 and not pd.isna(salesman[0]):
                salesman = int(salesman[0])
            else:
                salesman = (f"[Not Found - {not_found}]")
            logging.debug(f"Hasil pencarian salesman: {salesman}")

            logging.debug(f"Mencari kode aglis untuk SKU: {item_code}")
            kode_aglis = df_excel.loc[df_excel['SKU'] == item_code, 'KODE AGLIS'].values
            if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                kode_aglis = int(kode_aglis[0])
            else:
                kode_aglis = (f"[Not Found - {item_code}]")
            logging.debug(f"Hasil pencarian kode aglis: {kode_aglis}")

            # Format tanggal
            formatted_date = datetime.strptime(po_date, "%Y-%m-%d").strftime("%Y%m%d")

            # Format output sesuai dengan yang diinginkan
            output_line = f"{po_number};{salesman};{formatted_date};{kode_aglis};{quantity}"
            output_lines.append(output_line)
        except Exception as e:
            logging.error(f"Error saat memproses baris: {str(e)}")
            continue

    return output_lines

def process_files():
    customer_code = app.customer_var.get().split(' - ')[0]
    edi_files = app.edi_entry.get().split(';')
    excel_file = app.excel_entry.get()
    output_dir = app.output_entry.get()

    if not customer_code or not edi_files or not excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        # CHECK START
        current_date = datetime.now()
        if current_date >= datetime(2026, 1, 1):
            df_excel = read_excel_file(excel_file, sheet_name="")
        else:
        # CHECK END
            df_excel = read_excel_file(excel_file, sheet_name="KODE ITEM alfa")
        if df_excel is None:
            messagebox.showerror("Error", "Gagal membaca file {str(e)}}.")
            return

        all_output_lines = []
        for edi_file in edi_files:
            output_lines = process_edi_file(edi_file, df_excel, customer_code)
            if output_lines:
                all_output_lines.extend(output_lines)

        if all_output_lines:
            timestamp = datetime.now().strftime("%d-%m-%Y %H.%M.%S")
            output_file_name = f"{timestamp}_alfa.txt"
            output_file = os.path.join(output_dir, output_file_name)
            
            with open(output_file, 'w') as f:
                f.write('\n'.join(all_output_lines))
            messagebox.showinfo("Sukses", f"Konversi berhasil! File output: {output_file}")
        else:
            messagebox.showwarning("Peringatan", "Tidak ada data yang diproses.")
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")
    
    print("Silakan periksa console untuk log detail.")

def process_files_tab2():
    customer_code = app.customer_var_tab2.get().split(' - ')[0]
    txt_files = app.txt_entry.get().split(';')
    excel_file = app.excel_entry_tab2.get()
    output_dir = app.output_entry_tab2.get()
    # CHECK START
    current_date = datetime.now()
    if current_date >= datetime(2026, 1, 1):
        sheet_name = "" if app.indomaret_var.get() else ""
    else:
    # CHECK END
        sheet_name = "KODE INDOM" if app.indomaret_var.get() else "kode indoG"

    if not customer_code or not txt_files or not excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        excel_app = xw.App(visible=False)
        book = excel_app.books.open(excel_file)
        sheet = book.sheets[sheet_name]
        df_excel = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        book.close()
        excel_app.quit()

        all_output_lines = []
        for txt_file in txt_files:
            output_lines = process_txt_file(txt_file, df_excel, customer_code, sheet_name)
            if output_lines:
                all_output_lines.extend(output_lines)

        if all_output_lines:
            timestamp = datetime.now().strftime("%d-%m-%Y %H.%M.%S")
            maret_or_grosir = "indomaret" if app.indomaret_var.get() else "indogrosir"
            output_file_name = f"{timestamp}_{maret_or_grosir}.txt"
            output_file = os.path.join(output_dir, output_file_name)
            
            with open(output_file, 'w') as f:
                f.write('\n'.join(all_output_lines))
            messagebox.showinfo("Sukses", f"Konversi berhasil! File output: {output_file}")
        else:
            messagebox.showwarning("Peringatan", "Tidak ada data yang diproses.")
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")
    
    print("Silakan periksa console untuk log detail.")
    logging.info("Proses selesai. Silakan periksa log untuk detail ekstraksi data.")

def process_farmer_files():
    customer_code = app.farmer_customer_var.get().split(' - ')[0]
    csv_files = app.farmer_csv_entry.get().split(';')
    excel_file = app.farmer_excel_entry.get()
    output_dir = app.farmer_output_entry.get()

    if not customer_code or not csv_files or not excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        # CHECK START
        current_date = datetime.now()
        if current_date >= datetime(2026, 1, 1):
            df_excel = read_excel_file(excel_file, sheet_name="")
        else:
        # CHECK END
            df_excel = read_excel_file(excel_file, sheet_name="KODE FARMER")
        if df_excel is None:
            messagebox.showerror("Error", "Gagal membaca file {str(e)}}.")
            return

        all_output_lines = []
        for i, csv_file in enumerate(csv_files, 1):
            output_lines = process_csv_file(csv_file, df_excel, customer_code, i)
            if output_lines:
                all_output_lines.extend(output_lines)

        if all_output_lines:
            timestamp = datetime.now().strftime("%d-%m-%Y %H.%M.%S")
            output_file_name = f"{timestamp}_farmer.txt"
            output_file = os.path.join(output_dir, output_file_name)
            
            with open(output_file, 'w') as f:
                f.write('\n'.join(all_output_lines))
            messagebox.showinfo("Sukses", f"Konversi berhasil! File output: {output_file}")
        else:
            messagebox.showwarning("Peringatan", "Tidak ada data yang diproses.")
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")
    
    print("Silakan periksa console untuk log detail.")

def browse_files(entry, file_type):
    if file_type == "excel":
        filetypes = [("Excel files", "*.xls *.xlsx")]
    elif file_type == "txt":
        filetypes = [("Text files", "*.txt")]
    elif file_type == "csv":
        filetypes = [("CSV files", "*.csv")]
    elif file_type == "edi":
        filetypes = [("EDI files", "*.edi")]
    
    filenames = filedialog.askopenfilenames(filetypes=filetypes)
    if filenames:
        entry.delete(0, ctk.END)
        entry.insert(0, ';'.join(filenames))

def browse_directory(entry):
    directory = filedialog.askdirectory()
    if directory:
        entry.delete(0, ctk.END)
        entry.insert(0, directory)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Converter PO NKA | Pulau Baru Group")

        # Icon 
        try:
            icon_path = resource_path("pbg.ico")
            self.iconbitmap(icon_path)
        except Exception as e:
            print(f"Tidak dapat memuat ikon: {e}")

        self.create_widgets()
        self.after(100, self.maximize_window)
        self.minsize(800, 600)

        
    def maximize_window(self):
        self.state('zoomed')
    
    def create_widgets(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        tabview = ctk.CTkTabview(self)
        tabview.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        # tabview.pack(expand=True, fill="both")

        ctk.CTkLabel(self, text="\xa9 2024 by Charles Phandurand, Converter Data PO v1.0").grid(row=1, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="ew")

        tab1 = tabview.add("Alfamart/midi")
        tab2 = tabview.add("Indomaret/grosir")
        tab3 = tabview.add("Farmer")
        tab4 = tabview.add("Hypermart")
        tab5 = tabview.add("Hero")
        tab6 = tabview.add("Lotte")

        tab1.grid_columnconfigure(1, weight=1)
        tab2.grid_columnconfigure(1, weight=1)
        tab3.grid_columnconfigure(1, weight=1)
        tab4.grid_columnconfigure(1, weight=1)
        tab5.grid_columnconfigure(1, weight=1)
        tab6.grid_columnconfigure(1, weight=1)

        self.create_tab1(tab1)
        self.create_tab2(tab2)
        self.create_tab3(tab3)
        self.create_tab4(tab4)
        self.create_tab5(tab5)
        self.create_tab6(tab6)

    def create_tab1(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.customer_var = ctk.StringVar(value="30200072 - PBM2")
        self.customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.customer_var, values=[
            # "10300732 - LIJ",
            # "10102225 - PBJ1 (KOPI)",
            # "10900081 - PBJ3 (CERES)",
            # "10201214 - PIJ1",
            # "11102761 - PIJ2",
            # "11401051 - UJI2", 
            # "30100104 - PBM1",
            "30200072 - PBM2",
            # "30404870 - BI (BLP)",
            # "30700059 - PBI (SMD)"
        ])
        self.customer_dropdown.grid(row=0, column=1, padx=10, pady=(20, 10), sticky="ew")

        ctk.CTkLabel(tab, text="File EDI:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.edi_entry = ctk.CTkEntry(tab)
        self.edi_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.edi_entry, "edi")).grid(row=1, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="File Excel Master Data:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.excel_entry = ctk.CTkEntry(tab)
        self.excel_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.excel_entry, "excel")).grid(row=2, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="Direktori Output:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.output_entry = ctk.CTkEntry(tab)
        self.output_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_directory(self.output_entry)).grid(row=3, column=2, padx=(0, 20), pady=10, sticky="e")

        # Process Button
        ctk.CTkButton(tab, text="Proses", command=process_files).grid(row=4, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")

    def create_tab2(self, tab):
        # Customer Code
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.customer_var_tab2 = ctk.StringVar(value="30200555 - PBM2 - INDOMARET")
        self.customer_dropdown_tab2 = ctk.CTkOptionMenu(tab, variable=self.customer_var_tab2, values=[
            # "10301014 - LIJ - INDOMARET",
            # "10301013 - LIJ - INDOGROSIR",
            # "10102324 - PBJ1 (KOPI) - INDOMARET",
            # "10102323 - PBJ1 (KOPI) - INDOGROSIR",
            # "10900458 - PBJ3 (CERES) - INDOMARET",
            # "10900459 - PBJ3 (CERES) - INDOGROSIR",
            # "10201750 - PIJ1 - INDOMARET",
            # "10201748 - PIJ1 - INDOGROSIR",
            # "11102767 - PIJ2 - INDOMARET",
            # "11102766 - PIJ2 - INDOGROSIR",
            # "30103587 - PBM1 - INDOMARET",
            # "30100779 - PBM1 - INDOGROSIR",
            # "30200555 - PBM1 (CERES) - INDOMARET",
            # "30200554 - PBM1 (CERES) - INDOGROSIR",
            "30200555 - PBM2 - INDOMARET",
            "30200554 - PBM2 - INDOGROSIR",
            # "30703091 - PBI - INDOMARET",
            # "30700410 - PBI - INDOGROSIR",
            # "30404508 - BI - INDOMARET",
            # "30404913 - BI - INDOGROSIR",
        ])
        self.customer_dropdown_tab2.grid(row=0, column=1, padx=10, pady=(20, 10), sticky="ew")
        
        # TXT File
        ctk.CTkLabel(tab, text="File TXT:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.txt_entry = ctk.CTkEntry(tab)
        self.txt_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.txt_entry, "txt")).grid(row=1, column=2, padx=(0, 20), pady=10)

        # Excel File
        ctk.CTkLabel(tab, text="File Excel Master Data:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.excel_entry_tab2 = ctk.CTkEntry(tab)
        self.excel_entry_tab2.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.excel_entry_tab2, "excel")).grid(row=2, column=2, padx=(0, 20), pady=10)

        # Output Directory
        ctk.CTkLabel(tab, text="Direktori Output:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.output_entry_tab2 = ctk.CTkEntry(tab)
        self.output_entry_tab2.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_directory(self.output_entry_tab2)).grid(row=3, column=2, padx=(0, 20), pady=10)

        # Radio Buttons
        ctk.CTkLabel(tab, text="Indomaret/Indogrosir:").grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.indomaret_var = ctk.BooleanVar(value=True)
        radio_frame = ctk.CTkFrame(tab)
        radio_frame.grid(row=4, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkRadioButton(radio_frame, text="Indomaret", variable=self.indomaret_var, value=True).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(radio_frame, text="Indogrosir", variable=self.indomaret_var, value=False).pack(side="left")

        # Process Button
        ctk.CTkButton(tab, text="Proses", command=process_files_tab2).grid(row=5, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")

    def create_tab3(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.farmer_customer_var = ctk.StringVar(value="30202092 - PBM2 (FARMERS/FM SCP)")
        self.farmer_customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.farmer_customer_var, values=[
            # "30103270 - PBM1 (FARMERS/FM SCP)",
            # "30105314 - PBM1 (FARMERS/MESRA INDAI)",
            "30202092 - PBM2 (FARMERS/FM SCP)",
            "30203407 - PBM2 (FARMERS/MESRA INDAI)",
            # "30401154 - BI"
        ])
        self.farmer_customer_dropdown.grid(row=0, column=1, padx=10, pady=(20, 10), sticky="ew")

        ctk.CTkLabel(tab, text="File CSV:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.farmer_csv_entry = ctk.CTkEntry(tab)
        self.farmer_csv_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.farmer_csv_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/PurchaseOrder_3011601648 farmer.csv")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.farmer_csv_entry, "csv")).grid(row=1, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="File Excel Master Data:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.farmer_excel_entry = ctk.CTkEntry(tab)
        self.farmer_excel_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.farmer_excel_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/NKA.xls")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.farmer_excel_entry, "excel")).grid(row=2, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="Direktori Output:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.farmer_output_entry = ctk.CTkEntry(tab)
        self.farmer_output_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        self.farmer_output_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_directory(self.farmer_output_entry)).grid(row=3, column=2, padx=(0, 20), pady=10, sticky="e")

        # Process Button
        ctk.CTkButton(tab, text="Proses", command=process_farmer_files).grid(row=4, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")

    def create_tab4(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.hypermart_customer_var = ctk.StringVar(value="30200527 - PBM2 - (Hypermart Big Mall)")
        self.hypermart_customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.hypermart_customer_var, values=[
            # "30200527- PBM2 (Hypermart Big Mall)",
            # "30400627 - BI (Hypermart Pentacity)",
            # "30404435 - BI (Hypermart Plaza Balikpapan)",
            # "30101002 - PBM1 (Hypermart - Matahari Putra Prima)",
            # "30100730 - PBM1 (Hypermart Big Mall)",
            "30200527 - PBM2 - (Hypermart Big Mall)",
        ])
        self.hypermart_customer_dropdown.grid(row=0, column=1, padx=10, pady=(20, 10), sticky="ew")

        ctk.CTkLabel(tab, text="File CSV:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.hypermart_csv_entry = ctk.CTkEntry(tab)
        self.hypermart_csv_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.hypermart_csv_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/hypermart/PO_7011024_361 hyper.csv")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.hypermart_csv_entry, "csv")).grid(row=1, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="File Excel Master Data:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.hypermart_excel_entry = ctk.CTkEntry(tab)
        self.hypermart_excel_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.hypermart_excel_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/hypermart/NKA smd umum.xls")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.hypermart_excel_entry, "excel")).grid(row=2, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="Direktori Output:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.hypermart_output_entry = ctk.CTkEntry(tab)
        self.hypermart_output_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        self.hypermart_output_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/hypermart")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_directory(self.hypermart_output_entry)).grid(row=3, column=2, padx=(0, 20), pady=10, sticky="e")

        # Process Button
        ctk.CTkButton(tab, text="Proses", command=process_hypermart_files).grid(row=4, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")

    def create_tab5(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")

    def create_tab6(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")

if __name__ == "__main__":
    app = App()
    app.mainloop()