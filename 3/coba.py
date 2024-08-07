import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import xlwings as xw
import logging

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

def process_edi_file(edi_file, df_excel):
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
                output_line = f"{edi_1};10300732;{salesman};{edi_3};{kode_aglis};{calculated_value}"
                output_lines.append(output_line)

        except Exception as e:
            logging.error(f"Error saat memproses baris: {str(e)}")
            logging.exception("Traceback:")

    return output_lines

def process_files():
    edi_files = edi_entry.get().split(';')  # Assuming multiple files are separated by semicolons
    excel_file = excel_entry.get()
    output_dir = output_entry.get()

    if not edi_files or not excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih semua file dan direktori yang diperlukan.")
        return

    try:
        df_excel = read_excel_file(excel_file)
        if df_excel is None:
            messagebox.showerror("Error", "Gagal membaca file Excel.")
            return

        all_output_lines = []
        for edi_file in edi_files:
            output_lines = process_edi_file(edi_file, df_excel)
            if output_lines:
                all_output_lines.extend(output_lines)

        if all_output_lines:
            output_file = os.path.join(output_dir, "combined_output.txt")
            with open(output_file, 'w') as f:
                f.write('\n'.join(all_output_lines))
            messagebox.showinfo("Sukses", f"Konversi berhasil! File output: {output_file}")
        else:
            messagebox.showwarning("Peringatan", "Tidak ada data yang diproses.")
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")
    
    print("Silakan periksa console untuk log detail.")

def browse_files(entry):
    filenames = filedialog.askopenfilenames()
    entry.delete(0, tk.END)
    entry.insert(0, ';'.join(filenames))

def browse_directory(entry):
    directory = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, directory)

# Buat window utama
root = tk.Tk()
root.title("Konverter EDI ke TXT by Charles")

# Buat dan susun elemen-elemen GUI
tk.Label(root, text="File EDI:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
edi_entry = tk.Entry(root, width=50)
edi_entry.grid(row=0, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=lambda: browse_files(edi_entry)).grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="File Excel Master Data:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
excel_entry = tk.Entry(root, width=50)
excel_entry.grid(row=1, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=lambda: browse_files(excel_entry)).grid(row=1, column=2, padx=5, pady=5)

tk.Label(root, text="Direktori Output:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=2, column=1, padx=5, pady=5)
tk.Button(root, text="Browse", command=lambda: browse_directory(output_entry)).grid(row=2, column=2, padx=5, pady=5)

tk.Button(root, text="Proses", command=process_files).grid(row=3, column=1, pady=10)

# Jalankan aplikasi
root.mainloop()