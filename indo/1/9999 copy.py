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

                # VLOOKUP untuk SALESMAN
                salesman = df_excel.loc[df_excel['BARCODE'] == edi_6_lin, 'SALESMAN'].values
                if len(salesman) > 0 and not pd.isna(salesman[0]):
                    salesman = int(salesman[0])
                else:
                    salesman = 'Not Found'

                # VLOOKUP untuk KODE AGLIS
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
    customer_code = customer_var.get().split(' - ')[0]
    edi_files = edi_entry.get().split(';')
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

def browse_files(entry, file_type):
    if file_type == "excel":
        filetypes = [("Excel files", "*.xls *.xlsx")]
    elif file_type == "txt":
        filetypes = [("Text files", "*.txt")]
    elif file_type == "edi":
        filetypes = [("EDI files", "*.edi")]
    
    filenames = filedialog.askopenfilenames(filetypes=filetypes)
    if filenames:
        entry.delete(0, tk.END)
        # Gabungkan semua path file yang dipilih menjadi satu string yang dipisahkan oleh tanda koma atau lainnya
        entry.insert(0, ';'.join(filenames))


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
                nomor_po = ordmsg_line[41:50]
                tanggal_po = ordmsg_line[50:58]
                qty = orddtl_line[19:24]
                isi = orddtl_line[24:29]
                kode_item = orddtl_line[36:44]
                
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

def process_files_tab2():
    customer_code = customer_var_tab2.get().split(' - ')[0]
    txt_files = txt_entry.get().split(';')
    excel_file = excel_entry_tab2.get()
    output_dir = output_entry_tab2.get()
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
# root.iconbitmap(r'C:\Users\TOSHIBA PORTEGE Z30C\Desktop\program python\alfamart\3\pbg.ico')

# Mengatur ukuran jendela utama dan menempatkannya di tengah
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 400
window_height = 300
position_x = int(screen_width / 2 - window_width / 2)
position_y = int(screen_height / 2 - window_height / 1.5)
root.geometry(f'{window_width}x{window_height}+{position_x}+{position_y}')

# Membuat tab control
tab_control = ttk.Notebook(root)

# Tab 1 untuk konversi EDI
tab1 = tk.Frame(tab_control, bg='#CBE2B5')
tab_control.add(tab1, text='Alfamart/midi')

# Tab 2 kosong untuk sementara
tab2 = tk.Frame(tab_control, bg='#CBE2B5')
tab_control.add(tab2, text='Indomaret/grosir')

# Tambahkan tab_control ke root
tab_control.pack(expand=1, fill='both')

# Elemen-elemen GUI pada Tab 1
tab1.columnconfigure(0, weight=1)
tab1.columnconfigure(1, weight=2)
tab1.columnconfigure(2, weight=1)
tab1.rowconfigure(0, weight=1)
tab1.rowconfigure(1, weight=1)
tab1.rowconfigure(2, weight=1)
tab1.rowconfigure(3, weight=1)
tab1.rowconfigure(4, weight=1)

# Elemen-elemen GUI pada Tab 2
tab2.columnconfigure(0, weight=1)
tab2.columnconfigure(1, weight=2)
tab2.columnconfigure(2, weight=1)
tab2.rowconfigure(0, weight=1)
tab2.rowconfigure(1, weight=1)
tab2.rowconfigure(2, weight=1)
tab2.rowconfigure(3, weight=1)
tab2.rowconfigure(4, weight=1)
tab2.rowconfigure(5, weight=1)

tk.Label(tab1, text="Customer Code:", bg='#CBE2B5', fg='black').grid(row=0, column=0, sticky="e", padx=5, pady=5)
customer_var = tk.StringVar(tab1)
customer_var.set("10102225 - PBJ1 (KOPI)")
customer_dropdown = ttk.Combobox(tab1, textvariable=customer_var, values=[
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
], state="readonly")
customer_dropdown.grid(row=0, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)

tk.Label(tab1, text="File EDI:", bg='#CBE2B5', fg='black').grid(row=1, column=0, sticky="e", padx=5, pady=5)
edi_entry = tk.Entry(tab1)
edi_entry.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(tab1, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_files(edi_entry, "edi")).grid(row=1, column=2, sticky="nsew", padx=5, pady=5)

tk.Label(tab1, text="File Excel Master Data:", bg='#CBE2B5', fg='black').grid(row=2, column=0, sticky="e", padx=5, pady=5)
excel_entry = tk.Entry(tab1)
excel_entry.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(tab1, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_files(excel_entry, "excel")).grid(row=2, column=2, sticky="nsew", padx=5, pady=5)

tk.Label(tab1, text="Direktori Output:", bg='#CBE2B5', fg='black').grid(row=3, column=0, sticky="e", padx=5, pady=5)
output_entry = tk.Entry(tab1)
output_entry.grid(row=3, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(tab1, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_directory(output_entry)).grid(row=3, column=2, sticky="nsew", padx=5, pady=5)

# Tombol "Proses" pada Tab 1
proses_button = tk.Button(tab1, text="Proses", bg='#2C3E50', fg='white', command=process_files)
proses_button.grid(row=4, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)

# Elemen-elemen GUI pada Tab 2
indomaret_var = tk.BooleanVar(value=True)
tk.Radiobutton(tab2, text="Indomaret", variable=indomaret_var, value=True, bg='#CBE2B5').grid(row=4, column=0, sticky="w", padx=5, pady=5)
tk.Radiobutton(tab2, text="Indogrosir", variable=indomaret_var, value=False, bg='#CBE2B5').grid(row=4, column=1, sticky="w", padx=5, pady=5)

tk.Label(tab2, text="Customer Code:", bg='#CBE2B5', fg='black').grid(row=0, column=0, sticky="e", padx=5, pady=5)
customer_var_tab2 = tk.StringVar(tab2)
customer_var_tab2.set("10301014 - LIJ")
customer_dropdown_tab2 = ttk.Combobox(tab2, textvariable=customer_var_tab2, values=[
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
customer_dropdown_tab2.grid(row=0, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)

tk.Label(tab2, text="File TXT:", bg='#CBE2B5', fg='black').grid(row=1, column=0, sticky="e", padx=5, pady=5)
txt_entry = tk.Entry(tab2)
txt_entry.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(tab2, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_files(txt_entry, "txt")).grid(row=1, column=2, sticky="nsew", padx=5, pady=5)

tk.Label(tab2, text="File Excel Master Data:", bg='#CBE2B5', fg='black').grid(row=2, column=0, sticky="e", padx=5, pady=5)
excel_entry_tab2 = tk.Entry(tab2)
excel_entry_tab2.grid(row=2, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(tab2, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_files(excel_entry_tab2, "excel")).grid(row=2, column=2, sticky="nsew", padx=5, pady=5)

tk.Label(tab2, text="Direktori Output:", bg='#CBE2B5', fg='black').grid(row=3, column=0, sticky="e", padx=5, pady=5)
output_entry_tab2 = tk.Entry(tab2)
output_entry_tab2.grid(row=3, column=1, sticky="nsew", padx=5, pady=5)
tk.Button(tab2, text="Browse", bg='#2C3E50', fg='white', command=lambda: browse_directory(output_entry_tab2)).grid(row=3, column=2, sticky="nsew", padx=5, pady=5)

# Tombol "Proses" pada Tab 2
proses_button_tab2 = tk.Button(tab2, text="Proses", bg='#2C3E50', fg='white', command=process_files_tab2)
proses_button_tab2.grid(row=5, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)

tk.Label(tab1, text="© 2024 by Charles Phandurand, Converter Data PO v1.0", bg='#CBE2B5', fg='black', anchor='w').grid(row=5, column=0, columnspan=3, pady=(0, 6), sticky="nsew")
tk.Label(tab2, text="© 2024 by Charles Phandurand, Converter Data PO v1.0", bg='#CBE2B5', fg='black', anchor='w').grid(row=6, column=0, columnspan=3, pady=(0, 6), sticky="nsew")

# Jalankan aplikasi
root.mainloop()
