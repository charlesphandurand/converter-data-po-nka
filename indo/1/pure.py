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

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

def read_excel_file(file_path, sheet_name="KODE ITEM alfa"):
    try:
        app = xw.App(visible=False)
        book = app.books.open(file_path)
        sheet = book.sheets[sheet_name]
        
        data = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        
        book.close()
        app.quit()
        
        required_columns = ["BARCODE", "KODE AGLIS", "SALESMAN"]
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

def process_farmer_files():
    customer_code = app.farmer_customer_var.get().split(' - ')[0]
    csv_files = app.farmer_csv_entry.get().split(';')
    excel_file = app.farmer_excel_entry.get()
    output_dir = app.farmer_output_entry.get()

    if not customer_code or not csv_files or not excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        df_excel = read_excel_file(excel_file, sheet_name="KODE FARMER")
        if df_excel is None:
            messagebox.showerror("Error", "Gagal membaca file Excel.")
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

            logging.debug(f"Mencari salesman untuk barcode: {barcode}")
            salesman = df_excel.loc[df_excel['BARCODE'] == barcode, 'SALESMAN'].values
            if len(salesman) > 0 and not pd.isna(salesman[0]):
                salesman = int(salesman[0])
            else:
                salesman = 'Not Found'
            logging.debug(f"Hasil pencarian salesman: {salesman}")

            logging.debug(f"Mencari kode aglis untuk barcode: {barcode}")
            # Pencarian kode aglis
            kode_aglis = df_excel.loc[df_excel['BARCODE'] == barcode, 'KODE AGLIS'].values
            if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                kode_aglis = int(kode_aglis[0])
            else:
                kode_aglis = 'Not Found'
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

        tab1 = tabview.add("Alfamart/midi")
        tab2 = tabview.add("Farmer")
        
        tab1.grid_columnconfigure(1, weight=1)
        tab2.grid_columnconfigure(1, weight=1)

        self.create_tab1(tab1)
        self.create_tab2(tab2)

    def create_tab1(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.customer_var = ctk.StringVar(value="11102761 - PIJ2")
        self.customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.customer_var, values=[
            "11102761 - PIJ2",
            "10300732 - LIJ",
            "30404870 - BI (BLP)",
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
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.farmer_customer_var = ctk.StringVar(value="11102761 - PIJ2")
        self.farmer_customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.farmer_customer_var, values=[
            "11102761 - PIJ2",
            "10300732 - LIJ",
            "30404870 - BI (BLP)",
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

if __name__ == "__main__":
    app = App()
    app.mainloop()