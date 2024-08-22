import customtkinter as ctk
from tkinter import filedialog, messagebox
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

def read_excel_file(file_path):
    try:
        app = xw.App(visible=False)
        book = app.books.open(file_path)
        sheet = book.sheets["KODE ITEM alfa"]
        
        data = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        
        book.close()
        app.quit()
        
        required_columns = ["BARCODE", "KODE AGLIS", "SALESMAN"]
        for col in required_columns:
            if col not in data.columns:
                logging.error(f"Kolom '{col}' tidak ditemukan dalam sheet 'KODE ITEM alfa'")
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
            qty = line[19:24]
            isi = line[24:28]
            kode_item = line[36:44]
            
            logging.debug(f"Extracted data: { line }")
            logging.debug(f"Nomor PO: '{nomor_po}'")
            logging.debug(f"Tanggal PO: '{tanggal_po}'")
            logging.debug(f"QTY: '{qty}'")
            logging.debug(f"Isi: '{isi}'")
            logging.debug(f"Kode Item: '{kode_item}'")
            logging.debug(f"DataFrame contents:\n{df_excel.head()}")

            qty = int(qty)
            isi = int(isi)
            
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

def process_files():
    customer_code = app.customer_var.get().split(' - ')[0]
    edi_files = app.edi_entry.get().split(';')
    excel_file = app.excel_entry.get()
    output_dir = app.output_entry.get()

    if not customer_code or not edi_files or not excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        df_excel = read_excel_file(excel_file)
        if df_excel is None:
            messagebox.showerror("Error", "Gagal membaca file Excel.")
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

def browse_files(entry, file_type):
    if file_type == "excel":
        filetypes = [("Excel files", "*.xls *.xlsx")]
    elif file_type == "txt":
        filetypes = [("Text files", "*.txt")]
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
        self.title("Converter PO | Pulau Baru Group")
        self.geometry(f"{700}x{430}")
        self.resizable(0, 0)

        # Icon 
        try:
            icon_path = resource_path("pbg.ico")
            self.iconbitmap(icon_path)
        except Exception as e:
            print(f"Tidak dapat memuat ikon: {e}")

        # Center the window
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 1.7)
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))

        self.create_widgets()
    
    def create_widgets(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        tabview = ctk.CTkTabview(self)
        tabview.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        # tabview.pack(expand=True, fill="both")

        ctk.CTkLabel(self, text="\xa9 2024 by Charles Phandurand, Converter Data PO v1.0").grid(row=1, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="ew")

        tab1 = tabview.add("Alfamart/midi")
        tab2 = tabview.add("Indomaret/grosir")
        
        tab1.grid_columnconfigure(1, weight=1)
        tab2.grid_columnconfigure(1, weight=1)
        # tab2.grid_columnconfigure(4, weight=0)

        self.create_tab1(tab1)
        self.create_tab2(tab2)

    def create_tab1(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.customer_var = ctk.StringVar(value="10102225 - PBJ1 (KOPI)")
        self.customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.customer_var, values=[
            "10102225 - PBJ1 (KOPI)",
            "10900081 - PBJ3 (CERES)",
            "10201214 - PIJ1",
            "11102761 - PIJ2",
            "10300732 - LIJ                                                                   ",
            "30404870 - BI (BLP)",
            "11401051 - UJI2",
            "30100104 - PBM1",
            "30200072 - PBM2",
            "30700059 - PBI (SMD)"
        ], width=200)
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
        self.customer_var_tab2 = ctk.StringVar(value="10301014 - LIJ")
        self.customer_dropdown_tab2 = ctk.CTkOptionMenu(tab, variable=self.customer_var_tab2, values=[
            "10301014 - LIJ                                                                   ",
            "10102324 - PBJ3 (KOPI)",
            "10900458 - PBJ3 (CERES)",
            "10201750 - PIJ",
            "30103587 - PBM",
            "30200555 - PBM (CERES)",
            "30703091 - PBI",
            "30404508 - BI",
            "10301013 - LIJ",
            "10102323 - PBJ3 (KOPI)",
            "10900459 - PBJ3 (CERES)",
            "10201748 - PIJ",
            "30100779 - PBM",
            "30200554 - PBM (CERES)",
            "30700410 - PBI",
            "30404913 - BI",
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
        ctk.CTkRadioButton(tab, text="Indomaret", variable=self.indomaret_var, value=True).grid(row=4, column=1, padx=10, pady=10, sticky="w")
        ctk.CTkRadioButton(tab, text="Indogrosir", variable=self.indomaret_var, value=False).grid(row=4, column=1, padx=10, pady=10, sticky="ns")

        # Process Button
        ctk.CTkButton(tab, text="Proses", command=process_files_tab2).grid(row=5, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")

if __name__ == "__main__":
    app = App()
    app.mainloop()