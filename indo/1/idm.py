import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import os
import xlwings as xw
import logging
from datetime import datetime

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

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

def process_txt_file(txt_file, df_excel, customer_code, sheet_name):
    output_lines = []
    with open(txt_file, 'r') as f:
        lines = f.readlines()
    
    logging.info(f"Columns in Excel: {df_excel.columns.tolist()}")
    
    for i in range(0, len(lines), 2):
        if i+1 < len(lines):
            ordmsg_line = lines[i].strip()
            orddtl_line = lines[i+1].strip()
            
            if ordmsg_line.startswith("ORDMSG") and orddtl_line.startswith("ORDDTL"):
                nomor_po = ordmsg_line[49:52]
                tanggal_po = ordmsg_line[54:51]
                qty = orddtl_line[12:22]
                isi = orddtl_line[25:26]
                kode_item = orddtl_line[34:48]
                
                logging.debug(f"Extracted data:")
                logging.debug(f"Nomor PO: '{nomor_po}'")
                logging.debug(f"Tanggal PO: '{tanggal_po}'")
                logging.debug(f"QTY: '{qty}'")
                logging.debug(f"Isi: '{isi}'")
                logging.debug(f"Kode Item: '{kode_item}'")
                
                qty = int(qty)
                isi = int(isi)
                
                # VLOOKUP untuk SALESMAN
                salesman = df_excel.loc[df_excel['PLU'] == kode_item, 'SALESMAN'].values
                logging.debug(f"VLOOKUP result for SALESMAN: {salesman}")
                salesman = int(salesman[0]) if len(salesman) > 0 else 'Not Found'

                # VLOOKUP untuk KODE AGLIS
                kode_aglis = df_excel.loc[df_excel['PLU'] == kode_item, 'KODE AGLIS'].values
                logging.debug(f"VLOOKUP result for KODE AGLIS: {kode_aglis}")
                kode_aglis = int(kode_aglis[0]) if len(kode_aglis) > 0 else 'Not Found'

                pcs = qty * isi

                output_line = f"{nomor_po};{customer_code};{salesman};{tanggal_po};{kode_aglis};{pcs}"
                output_lines.append(output_line)
    
    return output_lines

def process_files_root():
    customer_code = customer_var_root.get().split(' - ')[0]
    txt_files = txt_entry.get().split(';')
    excel_file = excel_entry_root.get()
    output_dir = output_entry_root.get()
    sheet_name = "KODE INDOM" if indomaret_var.get() else "kode indoG"

    if not customer_code or not txt_files or not excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        app = xw.App(visible=False)
        book = app.books.open(excel_file)
        sheet = book.sheets[sheet_name]
        df_excel = sheet.used_range.options(pd.DataFrame, index=False, header=True).value
        book.close()
        app.quit()

        all_output_lines = []
        for txt_file in txt_files:
            output_lines = process_txt_file(txt_file, df_excel, customer_code, sheet_name)
            if output_lines:
                all_output_lines.extend(output_lines)

        if all_output_lines:
            timestamp = datetime.now().strftime("%d-%m-%Y %H.%M.%S")
            maret_or_grosir = "indomaret" if indomaret_var.get() else "indogrosir"
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

# Buat window utama
root = tk.Tk()
root.title("Converter PO | Pulau Baru Group")

# Mengatur ukuran jendela utama dan menempatkannya di tengah
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 400
window_height = 220
position_x = int(screen_width / 2 - window_width / 2)
position_y = int(screen_height / 2 - window_height / 1.5)
root.geometry(f'{window_width}x{window_height}+{position_x}+{position_y}')
root.iconbitmap(r'C:\Users\TOSHIBA PORTEGE Z30C\Desktop\program python\alfamart\3\pbg.ico')

# Mengatur warna background
root.configure(bg='#CBE2B5')

# Elemen-elemen GUI pada Tab 2
root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=2)
root.columnconfigure(2, weight=1)
root.rowconfigure(0, weight=1)
root.rowconfigure(1, weight=1)
root.rowconfigure(2, weight=1)
root.rowconfigure(3, weight=1)
root.rowconfigure(4, weight=1)
root.rowconfigure(5, weight=1)


# Elemen-elemen GUI pada Tab 2
# indomaret_var = tk.BooleanVar(value=True)
# tk.Radiobutton(root, text="Indomaret", variable=indomaret_var, value=True, bg='#CBE2B5').grid(row=4, column=0, sticky="w", padx=5, pady=5)
# tk.Radiobutton(root, text="Indogrosir", variable=indomaret_var, value=False, bg='#CBE2B5').grid(row=4, column=1, sticky="w", padx=5, pady=5)

tk.Label(root, text="Customer Code:", bg='#CBE2B5', fg='black').grid(row=0, column=0, sticky="e", padx=5, pady=5)
customer_var_root = tk.StringVar(root)
customer_var_root.set("10301014 - LIJ")
customer_dropdown_root = ttk.Combobox(root, textvariable=customer_var_root, values=[
    # Indomaret
    "10301014 - LIJ",
    "10102324 - PBJ3 (KOPI)",
    "10900458 - PBJ3 (CERES)",
    "10201750 - PIJ",
    "30103587 - PBM",
    "30200555 - PBM (CERES)",
    "30703091 - PBI",
    "30404508 - BI",
    # Indogrosir
    "10301013 - LIJ",
    "10102323 - PBJ3 (KOPI)",
    "10900459 - PBJ3 (CERES)",
    "10201748 - PIJ",
    "30100779 - PBM",
    "30200554 - PBM (CERES)",
    "30700410 - PBI",
    "30404913 - BI",
], state="readonly")
customer_dropdown_root.grid(row=0, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)

tk.Label(root, text="File TXT:", bg='#CBE2B5', fg='black').grid(row=1, column=0, sticky="e", padx=5, pady=5)
txt_entry = tk.Entry(root)
txt_entry.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(root, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_files(txt_entry, "txt")).grid(row=1, column=2, sticky="nsew", padx=5, pady=5)

tk.Label(root, text="File Excel Master Data:", bg='#CBE2B5', fg='black').grid(row=2, column=0, sticky="e", padx=5, pady=5)
excel_entry_root = tk.Entry(root)
excel_entry_root.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(root, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_files(excel_entry_root, "excel")).grid(row=2, column=2, sticky="nsew", padx=5, pady=5)

tk.Label(root, text="Direktori Output:", bg='#CBE2B5', fg='black').grid(row=3, column=0, sticky="e", padx=5, pady=5)
output_entry_root = tk.Entry(root)
output_entry_root.grid(row=3, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(root, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_directory(output_entry_root)).grid(row=3, column=2, sticky="nsew", padx=5, pady=5)

# Tombol "Proses" pada Tab 2
proses_button_root = tk.Button(root, text="Proses", bg='#2C3E50', fg='white', command=process_files_root)
proses_button_root.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)

tk.Label(root, text="© 2024 by Charles Phandurand, Converter Data PO v1.0", bg='#CBE2B5', fg='black', anchor='w').grid(row=6, column=0, columnspan=3, pady=(0, 6), sticky="nsew")

# Jalankan aplikasi
root.mainloop()
