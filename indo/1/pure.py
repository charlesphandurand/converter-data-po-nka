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
import csv

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

def read_excel_file(file_path, sheet_name):
    try:
        app = xw.App(visible=False)
        book = app.books.open(file_path)
        sheet = book.sheets[sheet_name]
        
        data = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        
        book.close()
        app.quit()
        
        if sheet_name == "KODE FARMER":
            required_columns = ["BARCODE", "KODE AGLIS", "SALESMAN"]
        elif sheet_name == "KODE HERO":
            required_columns = ["BARCODE", "KODE AGLIS", "SALESMAN"]
        elif sheet_name == "KODE LOTTE":
            required_columns = ["BARCODE", "KODE AGLIS", "SALESMAN"]
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

def process_hero_files():
    customer_code = app.hero_customer_var.get().split(' - ')[0]
    csv_files = app.hero_csv_entry.get().split(';')
    excel_file = app.hero_excel_entry.get()
    output_dir = app.hero_output_entry.get()

    if not customer_code or not csv_files or not excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        df_excel = read_excel_file(excel_file, sheet_name="KODE HERO")
        if df_excel is None:
            messagebox.showerror("Error", "Gagal membaca file Excel.")
            return

        all_output_lines = []
        for csv_file in csv_files:
            output_lines = process_hero_csv(csv_file, df_excel, customer_code)
            if output_lines:
                all_output_lines.extend(output_lines)

        if all_output_lines:
            timestamp = datetime.now().strftime("%d-%m-%Y %H.%M.%S")
            output_file_name = f"{timestamp}_hero.txt"
            output_file = os.path.join(output_dir, output_file_name)
            
            with open(output_file, 'w') as f:
                f.write('\n'.join(all_output_lines))
            messagebox.showinfo("Sukses", f"Konversi berhasil! File output: {output_file}")
        else:
            messagebox.showwarning("Peringatan", "Tidak ada data yang diproses.")
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")
    
    print("Silakan periksa console untuk log detail.")

def process_hero_csv(csv_file, df_excel, customer_code):
    output_lines = []
    try:
        with open(csv_file, 'r') as file:
            content = file.read().replace('"', '')  # Menghapus semua tanda kutip ganda
            rows = csv.reader(content.splitlines(), delimiter=',')
            next(rows, None)  # Skip the header row if it exists
            for row in rows:
                if len(row) < 51:  # Memastikan baris memiliki setidaknya 51 kolom
                    logging.error(f"Baris tidak memiliki jumlah kolom yang cukup: {row}")
                    continue

                po_number = row[0]
                po_date = row[1]
                barcode = row[27]  # Menggunakan indeks 25 untuk barcode (kolom ke-26)
                barang = row[29]
                qty = int(row[32]) * int(row[33])

                salesman = df_excel.loc[df_excel['BARCODE'] == barcode, 'SALESMAN'].values
                if len(salesman) > 0 and not pd.isna(salesman[0]):
                    salesman = int(salesman[0])
                else:
                    salesman = f"[Not Found - {barang}]"

                kode_aglis = df_excel.loc[df_excel['BARCODE'] == barcode, 'KODE AGLIS'].values
                if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                    kode_aglis = int(kode_aglis[0])
                else:
                    kode_aglis = f"[Not Found - {barcode}]"

                output_line = f"{po_number};{customer_code};{salesman};{po_date};{kode_aglis};{qty}"
                output_lines.append(output_line)
                logging.info(f"Baris berhasil diproses: {output_line}")

                # Debug logging
                logging.debug(f"Processing row: {row}")
                logging.debug(f"Barcode: {barcode}, Qty: {qty}, Salesman: {salesman}, Kode Aglis: {kode_aglis}")

    except Exception as e:
        logging.error(f"Error saat memproses file CSV Hero: {str(e)}")

    return output_lines

def process_lotte_files():
    customer_code = app.lotte_customer_var.get().split(' - ')[0]
    excel_files = app.lotte_excel_entry.get().split(';')
    master_excel_file = app.lotte_master_excel_entry.get()
    output_dir = app.lotte_output_entry.get()

    if not customer_code or not excel_files or not master_excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        df_master = read_excel_file(master_excel_file, sheet_name="KODE LOTTE")
        if df_master is None:
            messagebox.showerror("Error", "Gagal membaca file Excel master.")
            return

        all_output_lines = []
        for excel_file in excel_files:
            output_lines = process_lotte_excel(excel_file, df_master, customer_code)
            if output_lines:
                all_output_lines.extend(output_lines)

        if all_output_lines:
            timestamp = datetime.now().strftime("%d-%m-%Y %H.%M.%S")
            output_file_name = f"{timestamp}_lotte.txt"
            output_file = os.path.join(output_dir, output_file_name)
            
            with open(output_file, 'w') as f:
                f.write('\n'.join(all_output_lines))
            messagebox.showinfo("Sukses", f"Konversi berhasil! File output: {output_file}")
        else:
            messagebox.showwarning("Peringatan", "Tidak ada data yang diproses.")
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")
    
    print("Silakan periksa console untuk log detail.")

def process_lotte_excel(excel_file, df_master, customer_code):
    output_lines = []
    try:
        df = pd.read_excel(excel_file, sheet_name=0, header=None)
        logging.info(f"File Excel berhasil dibaca. Ukuran dataframe: {df.shape}")

        # Cek apakah file kosong
        if df.empty:
            logging.warning("File Excel kosong.")
            return output_lines

        # Ambil nomor PO dari sel B2
        po_number = str(df.iloc[1, 1])
        logging.info(f"Nomor PO ditemukan: {po_number}")

        # Ambil tanggal PO dari sel B3 dan ubah formatnya
        po_date = pd.to_datetime(df.iloc[2, 1])
        po_date_formatted = po_date.strftime('%Y%m%d')
        logging.info(f"Tanggal PO ditemukan: {po_date}, diformat menjadi: {po_date_formatted}")

        # Cari header untuk data produk
        header_row = None
        for i, row in df.iterrows():
            if row.astype(str).str.contains('PROD_CD|SCMRK_CD|STORE ORDER QTY|UOM', case=False).any():
                header_row = i
                break

        if header_row is None:
            logging.error("Tidak dapat menemukan header untuk data produk.")
            return output_lines

        # Gunakan baris header yang ditemukan
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row+1:].reset_index(drop=True)

        logging.info(f"Kolom yang ditemukan dalam data produk: {df.columns.tolist()}")

        # Proses setiap baris data produk
        for _, row in df.iterrows():
            try:
                scmrk_cd = str(row.get('SCMRK_CD', row.get('PROD_CD', '')))
                if pd.isna(scmrk_cd) or scmrk_cd == '':
                    continue

                salesman = df_master.loc[df_master['BARCODE'] == scmrk_cd, 'SALESMAN'].values
                if len(salesman) > 0 and not pd.isna(salesman[0]):
                    salesman = int(salesman[0])
                else:
                    salesman = f"[Not Found - {row.get('PROD_DESC', scmrk_cd)}]"

                kode_aglis = df_master.loc[df_master['BARCODE'] == scmrk_cd, 'KODE AGLIS'].values
                if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                    kode_aglis = int(kode_aglis[0])
                else:
                    kode_aglis = f"[Not Found - {scmrk_cd}]"

                qty = int(float(row.get('STORE ORDER QTY', 0))) * int(float(row.get('UOM', 1)))

                output_line = f"{po_number};{customer_code};{salesman};{po_date_formatted};{kode_aglis};{qty}"
                output_lines.append(output_line)
                logging.info(f"Baris berhasil diproses: {output_line}")
            except Exception as row_error:
                logging.warning(f"Error saat memproses baris: {row_error}")

    except Exception as e:
        logging.error(f"Error saat memproses file Excel Lotte: {str(e)}")
        logging.exception("Traceback lengkap:")

    return output_lines

def browse_files(entry, file_type):
    if file_type == "excel":
        filetypes = [("Excel files", "*.xls *.xlsx")]
    elif file_type == "csv":
        filetypes = [("CSV files", "*.csv")]
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

        ctk.CTkLabel(self, text="\xa9 2024 by Charles Phandurand, Converter Data PO v1.0").grid(row=1, column=0, columnspan=2, padx=10, pady=(5, 10), sticky="ew")

        tab1 = tabview.add("Lotte")
        tab2 = tabview.add("Hero")
        
        tab1.grid_columnconfigure(1, weight=1)
        tab2.grid_columnconfigure(1, weight=1)

        self.create_tab1(tab1)
        self.create_tab2(tab2)

    def create_tab1(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.lotte_customer_var = ctk.StringVar(value="30400858 - BI (Lotte)")
        self.lotte_customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.lotte_customer_var, values=[
            "30400858 - BI (Lotte)",
            "30200702 - PBM2 (Lotte Shopping)",
        ])
        self.lotte_customer_dropdown.grid(row=0, column=1, padx=10, pady=(20, 10), sticky="ew")

        ctk.CTkLabel(tab, text="File Excel:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.lotte_excel_entry = ctk.CTkEntry(tab)
        self.lotte_excel_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.lotte_excel_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/lotte/PO2311010603200049_20240830140336.xlsx")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.lotte_excel_entry, "excel")).grid(row=1, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="File Excel Master Data:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.lotte_master_excel_entry = ctk.CTkEntry(tab)
        self.lotte_master_excel_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.lotte_master_excel_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/lotte/NKA smd umum.xls")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.lotte_master_excel_entry, "excel")).grid(row=2, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="Direktori Output:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.lotte_output_entry = ctk.CTkEntry(tab)
        self.lotte_output_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        self.lotte_output_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/lotte")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_directory(self.lotte_output_entry)).grid(row=3, column=2, padx=(0, 20), pady=10, sticky="e")

        # Process Button
        ctk.CTkButton(tab, text="Proses", command=process_lotte_files).grid(row=4, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")
        
    def create_tab2(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.hero_customer_var = ctk.StringVar(value="11102761 - PIJ2")
        self.hero_customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.hero_customer_var, values=[
            "11102761 - PIJ2",
            "10300732 - LIJ",
            "30404870 - BI (BLP)",
        ])
        self.hero_customer_dropdown.grid(row=0, column=1, padx=10, pady=(20, 10), sticky="ew")

        ctk.CTkLabel(tab, text="File CSV:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.hero_csv_entry = ctk.CTkEntry(tab)
        self.hero_csv_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.hero_csv_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/hero/PurchaseOrder_57741173 hero.csv")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.hero_csv_entry, "csv")).grid(row=1, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="File Excel Master Data:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.hero_excel_entry = ctk.CTkEntry(tab)
        self.hero_excel_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.hero_excel_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/hero/NKA smd umum.xls")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.hero_excel_entry, "excel")).grid(row=2, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="Direktori Output:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.hero_output_entry = ctk.CTkEntry(tab)
        self.hero_output_entry.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
        self.hero_output_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/hero")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_directory(self.hero_output_entry)).grid(row=3, column=2, padx=(0, 20), pady=10, sticky="e")

        # Process Button
        ctk.CTkButton(tab, text="Proses", command=process_hero_files).grid(row=4, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")

if __name__ == "__main__":
    app = App()
    app.mainloop()