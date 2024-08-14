import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os
import xlwings as xw
import logging
from datetime import datetime

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def read_excel_file(file_path):
    try:
        app = xw.App(visible=False)
        book = app.books.open(file_path)
        sheet = book.sheets["KODE ITEM alfa"]
        
        # Baca data dari sheet "KODE ITEM alfa"
        data = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        
        book.close()
        app.quit()
        
        # Pastikan kolom yang dibutuhkan ada
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

                # VLOOKUP untuk SALESMAN
                salesman = df_excel.loc[df_excel['BARCODE'] == edi_6_lin, 'SALESMAN'].values
                salesman = int(salesman[0]) if len(salesman) > 0 else 'Not Found'

                # VLOOKUP untuk KODE AGLIS
                kode_aglis = df_excel.loc[df_excel['BARCODE'] == edi_6_lin, 'KODE AGLIS'].values
                kode_aglis = int(kode_aglis[0]) if len(kode_aglis) > 0 else 'Not Found'

                # Calculate the last value
                lin_value_1 = int(lin_line[2]) if len(lin_line) > 2 else 0
                lin_value_2 = int(lin_line[8]) if len(lin_line) > 8 else 0

                calculated_value = lin_value_1 * lin_value_2

                # Format output line
                output_line = f"{edi_1};{customer_code};{salesman};{edi_3};{kode_aglis};{calculated_value}"
                output_lines.append(output_line)

        except Exception as e:
            logging.error(f"Error saat memproses baris: {str(e)}")
            logging.exception("Traceback:")

    return output_lines

def process_files():
    customer_code = customer_var.get().split(' - ')[0]  # Ambil kode customer dari pilihan dropdown
    edi_files = edi_entry.get().split(';')  # Assuming multiple files are separated by semicolons
    excel_file = excel_entry.get()
    output_dir = output_entry.get()

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
            # Dapatkan timestamp saat ini
            timestamp = datetime.now().strftime("%d-%m-%Y %H.%M.%S")
            output_file_name = f"{timestamp}.txt"
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
    elif file_type == "edi":
        filetypes = [("Edi files", "*.edi")]    
    
    filename = filedialog.askopenfilename(filetypes=filetypes)
    if filename:
        entry.delete(0, tk.END)
        entry.insert(0, filename)

def browse_directory(entry):
    directory = filedialog.askdirectory()
    if directory:
        entry.delete(0, tk.END)
        entry.insert(0, directory)

# Buat window utama
root = tk.Tk()
root.title("Convert PO Alfamart | Pulau Baru Group")

# Mengubah warna background jendela
root.configure(bg='#CBE2B5')
root.iconbitmap(r'C:\Users\TOSHIBA PORTEGE Z30C\Desktop\program python\alfamart\3\pbg.ico')

# Mendapatkan ukuran layar
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Mendapatkan ukuran jendela
window_width = 400  # Lebar jendela, bisa disesuaikan
window_height = 200  # Tinggi jendela, bisa disesuaikan

# Menghitung posisi x dan y untuk menempatkan jendela di tengah layar
position_x = int(screen_width / 2 - window_width / 2)
position_y = int(screen_height / 2 - window_height / 1.5)

# Menetapkan ukuran jendela dan menempatkannya di tengah
root.geometry(f'{window_width}x{window_height}+{position_x}+{position_y}')

# Membuat semua kolom dan baris fleksibel
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=2)  # Kolom untuk entri dibuat lebih fleksibel
root.columnconfigure(2, weight=1)
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)
root.rowconfigure(2, weight=1)
root.rowconfigure(3, weight=1)
root.rowconfigure(4, weight=1)

# Daftar customer codes
customer_code_options = [
    "10102225 - PBJ1 (KOPI)",
    "10900081 - PBJ3 (CERES)",
    "10201214 - PIJ1",
    "11102761 - PIJ2",
    "10300732 - LIJ",
    "30404870 - BI (BLP)",
    "11401051 - UJI2",
    "30100104 - PBM1",
    "30200072 - PBM2",
    "30700059 - PBI (SMD)"
]

# Variabel untuk menyimpan pilihan customer code
customer_var = tk.StringVar(root)
customer_var.set(customer_code_options[0])  # set nilai default

# Elemen-elemen GUI
tk.Label(root, text="Customer Code:", bg='#CBE2B5', fg='black').grid(row=0, column=0, sticky="e", padx=5, pady=5)
customer_dropdown = ttk.Combobox(root, textvariable=customer_var, values=customer_code_options, state="readonly")
customer_dropdown.grid(row=0, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)

tk.Label(root, text="File EDI:", bg='#CBE2B5', fg='black').grid(row=1, column=0, sticky="e", padx=5, pady=5)
edi_entry = tk.Entry(root)
edi_entry.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(root, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_files(edi_entry, "edi")).grid(row=1, column=2, sticky="nsew", padx=5, pady=5)

tk.Label(root, text="File Excel Master Data:", bg='#CBE2B5', fg='black').grid(row=2, column=0, sticky="e", padx=5, pady=5)
excel_entry = tk.Entry(root)
excel_entry.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(root, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_files(excel_entry, "excel")).grid(row=2, column=2, sticky="nsew", padx=5, pady=5)

tk.Label(root, text="Direktori Output:", bg='#CBE2B5', fg='black').grid(row=3, column=0, sticky="e", padx=5, pady=5)
output_entry = tk.Entry(root)
output_entry.grid(row=3, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(root, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_directory(output_entry)).grid(row=3, column=2, sticky="nsew", padx=5, pady=5)

# Membuat tombol "Proses" yang fleksibel
proses_button = tk.Button(root, text="Proses", bg='#2C3E50', fg='white', command=process_files)
proses_button.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)

tk.Label(root, text="© 2024 by Charles Phandurand, Converter Data PO Alfamart v1.0", bg='#CBE2B5', fg='black', anchor='w').grid(row=5, column=0, columnspan=3, pady=(0, 6), sticky="nsew")

# Jalankan aplikasi
root.mainloop()