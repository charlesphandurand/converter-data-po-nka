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
        elif sheet_name == "KODE HYPERMART":
            required_columns = ["SKU", "KODE AGLIS", "SALESMAN"]
        elif sheet_name == "KODE LOTTE":
            required_columns = ["BARCODE", "KODE AGLIS", "SALESMAN"]
        elif sheet_name == "KODE HERO":
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
            no_po = pohdr_line[1] if len(pohdr_line) > 1 else 'Unknown'
            tgl_po = pohdr_line[2] if len(pohdr_line) > 2 else 'Unknown'

            for lin_line in lin_lines:
                kode_item = lin_line[5] if len(lin_line) > 5 else 'Unknown'
                barang = lin_line[1].split(';')[0].strip() if len(lin_line) > 1 else 'Unknown'
                
                salesman = df_excel.loc[df_excel['BARCODE'] == kode_item, 'SALESMAN'].values
                if len(salesman) > 0 and not pd.isna(salesman[0]):
                    salesman = int(salesman[0])
                else:
                    salesman = f"[Not Found - {barang}]"

                kode_aglis = df_excel.loc[df_excel['BARCODE'] == kode_item, 'KODE AGLIS'].values
                if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                    kode_aglis = int(kode_aglis[0])
                else:
                    kode_aglis = f"[Not Found - {kode_item}]"

                qty = int(lin_line[2]) if len(lin_line) > 2 else 0
                isi = int(lin_line[8]) if len(lin_line) > 8 else 0

                calculated_value = qty * isi

                output_line = f"{no_po};{customer_code};{salesman};{tgl_po};{kode_aglis};{calculated_value}"
                output_lines.append(output_line)

        except Exception as e:
            logging.error(f"Error saat memproses baris: {str(e)}")
            logging.exception("Traceback:")

    return output_lines

def process_alfamart():
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
            messagebox.showerror("Error", "Terjadi kesalahan: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Open method of Workbooks class failed', 'xlmain11.chm', 0, -2146827284), None)")
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
            barang = line[44:64]
            
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
                salesman = f"[Not Found - {barang}]"

            kode_aglis = df_excel.loc[df_excel['PLU'] == kode_item, 'KODE AGLIS'].values
            logging.debug(f"VLOOKUP result for KODE AGLIS: {kode_aglis}")
            if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                kode_aglis = int(kode_aglis[0])
            else:
                kode_aglis = f"[Not Found - {kode_item}]"

            pcs = qty * isi

            output_line = f"{nomor_po};{customer_code};{salesman};{tanggal_po};{kode_aglis};{pcs}"
            output_lines.append(output_line)
    
    return output_lines

def process_indomaret():
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
    if sheet_name is None:
        messagebox.showerror("Error", "Terjadi kesalahan: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Open method of Workbooks class failed', 'xlmain11.chm', 0, -2146827284), None)")
        return

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

def process_farmer_csv(csv_file, df_excel, customer_code, file_number):
    output_lines = []
    
    try:
        with open(csv_file, 'r', newline='', encoding='utf-8-sig') as file:
            csv_reader = csv.reader(file, quotechar='"', delimiter=',', quoting=csv.QUOTE_MINIMAL)
            next(csv_reader)  # Skip header row
            
            for row in csv_reader:
                logging.debug(f"Raw row: {row}")
                
                # Jika baris hanya memiliki satu elemen, itu mungkin karena pemisahan yang salah
                if len(row) == 1:
                    row = row[0].split(',')
                
                # Menggabungkan kembali elemen yang mungkin terpisah karena tanda kutip
                merged_row = []
                merge_next = False
                for item in row:
                    if merge_next:
                        merged_row[-1] += "," + item.strip('"')
                        if item.endswith('"'):
                            merge_next = False
                    elif item.startswith('"') and not item.endswith('"'):
                        merged_row.append(item.strip('"'))
                        merge_next = True
                    else:
                        merged_row.append(item.strip('"'))
                
                row = merged_row
                
                if len(row) < 15:  # Memastikan baris memiliki minimal 15 kolom
                    logging.error(f"Baris tidak memiliki jumlah kolom yang cukup: {row}")
                    continue
                
                po_number = row[0].strip()
                barcode = row[10].strip() if len(row) > 10 else ""
                po_date = row[-1].strip()  # Mengambil elemen terakhir sebagai tanggal PO
                order_quantity = row[11].strip().replace('"', '').replace(',', '.') if len(row) > 11 else "0"
                uom_pack_size = row[13].strip() if len(row) > 12 else "1"
                barang = row[18].strip() if len(row) > 18 else ""
                
                logging.debug(f"Extracted values: PO: {po_number}, Barcode: {barcode}, Date: {po_date}, Qty: {order_quantity}, UOM: {uom_pack_size}, Barang: {barang}")
                
                try:
                    order_quantity = float(order_quantity)
                    uom_pack_size = int(uom_pack_size)
                except ValueError:
                    logging.error(f"Invalid order quantity or UOM pack size: {order_quantity}, {uom_pack_size}")
                    continue
                
                logging.debug(f"Mencari salesman untuk barcode: {barcode}")
                salesman = df_excel.loc[df_excel['BARCODE'] == barcode, 'SALESMAN'].values
                if len(salesman) > 0 and not pd.isna(salesman[0]):
                    salesman = int(salesman[0])
                else:
                    salesman = f"[Not Found - {barang}]"
                logging.debug(f"Hasil pencarian salesman: {salesman}")
                
                logging.debug(f"Mencari kode aglis untuk barcode: {barcode}")
                kode_aglis = df_excel.loc[df_excel['BARCODE'] == barcode, 'KODE AGLIS'].values
                if len(kode_aglis) > 0 and not pd.isna(kode_aglis[0]):
                    kode_aglis = int(kode_aglis[0])
                else:
                    kode_aglis = f"[Not Found - {barcode}]"
                logging.debug(f"Hasil pencarian kode aglis: {kode_aglis}")
                
                pcs = int(order_quantity * uom_pack_size)
                output_line = f"{po_number};{customer_code};{salesman};{po_date};{kode_aglis};{pcs}"
                output_lines.append(output_line)
                logging.info(f"Baris berhasil diproses: {output_line}")
    
    except Exception as e:
        logging.error(f"Error saat memproses file CSV: {str(e)}")
        logging.exception("Traceback lengkap:")
    
    return output_lines

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
            messagebox.showerror("Error", "Terjadi kesalahan: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Open method of Workbooks class failed', 'xlmain11.chm', 0, -2146827284), None)")
            return

        all_output_lines = []
        for i, csv_file in enumerate(csv_files, 1):
            logging.info(f"Memproses file CSV: {csv_file}")
            output_lines = process_farmer_csv(csv_file, df_excel, customer_code, i)
            if output_lines:
                all_output_lines.extend(output_lines)
            logging.info(f"Jumlah baris yang berhasil diproses dari file {csv_file}: {len(output_lines)}")

        if all_output_lines:
            timestamp = datetime.now().strftime("%d-%m-%Y %H.%M.%S")
            output_file_name = f"{timestamp}_farmer.txt"
            output_file = os.path.join(output_dir, output_file_name)
            
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(all_output_lines))
            messagebox.showinfo("Sukses", f"Konversi berhasil! File output: {output_file}\nTotal baris yang diproses: {len(all_output_lines)}")
        else:
            messagebox.showwarning("Peringatan", "Tidak ada data yang berhasil diproses dari semua file.")
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {str(e)}")
    
    print("Silakan periksa console untuk log detail.")

def process_hypermart_csv(csv_file, df_excel, customer_code):
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
            output_line = f"{po_number};{customer_code};{salesman};{formatted_date};{kode_aglis};{quantity}"
            output_lines.append(output_line)
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
            df_excel = read_excel_file(excel_file, sheet_name="KODE HYPERMART")
        if df_excel is None:
            messagebox.showerror("Error", "Terjadi kesalahan: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Open method of Workbooks class failed', 'xlmain11.chm', 0, -2146827284), None)")
            return

        all_output_lines = []
        for csv_file in csv_files:
            output_lines = process_hypermart_csv(csv_file, df_excel, customer_code)
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

def process_hero_files():
    customer_code = app.hero_customer_var.get().split(' - ')[0]
    csv_files = app.hero_csv_entry.get().split(';')
    excel_file = app.hero_excel_entry.get()
    output_dir = app.hero_output_entry.get()

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
            df_excel = read_excel_file(excel_file, sheet_name="KODE HERO")
        if df_excel is None:
            messagebox.showerror("Error", "Terjadi kesalahan: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Open method of Workbooks class failed', 'xlmain11.chm', 0, -2146827284), None)")
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

def process_lotte_excel(excel_file, df_excel, customer_code):
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

                salesman = df_excel.loc[df_excel['BARCODE'] == scmrk_cd, 'SALESMAN'].values
                if len(salesman) > 0 and not pd.isna(salesman[0]):
                    salesman = int(salesman[0])
                else:
                    salesman = f"[Not Found - {row.get('PROD_DESC', scmrk_cd)}]"

                kode_aglis = df_excel.loc[df_excel['BARCODE'] == scmrk_cd, 'KODE AGLIS'].values
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

def process_lotte_files():
    customer_code = app.lotte_customer_var.get().split(' - ')[0]
    excel_files = app.lotte_excel_entry.get().split(';')
    master_excel_file = app.lotte_master_excel_entry.get()
    output_dir = app.lotte_output_entry.get()

    if not customer_code or not excel_files or not master_excel_file or not output_dir:
        messagebox.showerror("Error", "Silakan pilih customer code dan semua file yang diperlukan.")
        return

    try:
        # CHECK START
        current_date = datetime.now()
        if current_date >= datetime(2026, 1, 1):
            df_excel = read_excel_file(excel_file, sheet_name="")
        else:
        # CHECK END
            df_excel = read_excel_file(master_excel_file, sheet_name="KODE LOTTE")
        if df_excel is None:
            messagebox.showerror("Error", "Terjadi kesalahan: (-2147352567, 'Exception occurred.', (0, 'Microsoft Excel', 'Open method of Workbooks class failed', 'xlmain11.chm', 0, -2146827284), None)")
            return

        all_output_lines = []
        for excel_file in excel_files:
            output_lines = process_lotte_excel(excel_file, df_excel, customer_code)
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
        self.customer_var = ctk.StringVar(value="30404870 - BI (BLP)")
        self.customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.customer_var, values=[
            # "10300732 - LIJ",
            # "10102225 - PBJ1 (KOPI)",
            # "10900081 - PBJ3 (CERES)",
            # "10201214 - PIJ1",
            # "11102761 - PIJ2",
            # "11401051 - UJI2", 
            # "30100104 - PBM1",
            # "30200072 - PBM2",
            "30404870 - BI (BLP)",
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
        ctk.CTkButton(tab, text="Proses", command=process_alfamart).grid(row=4, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")

    def create_tab2(self, tab):
        # Customer Code
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.customer_var_tab2 = ctk.StringVar(value="30404508 - BI - INDOMARET")
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
            # "30200555 - PBM2 - INDOMARET",
            # "30200554 - PBM2 - INDOGROSIR",
            # "30703091 - PBI - INDOMARET",
            # "30700410 - PBI - INDOGROSIR",
            "30404508 - BI - INDOMARET",
            "30404913 - BI - INDOGROSIR",
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
        ctk.CTkButton(tab, text="Proses", command=process_indomaret).grid(row=5, column=0, columnspan=3, padx=10, pady=(20, 10), sticky="ew")

    def create_tab3(self, tab):
        ctk.CTkLabel(tab, text="Customer Code:").grid(row=0, column=0, padx=10, pady=(20, 10), sticky="w")
        self.farmer_customer_var = ctk.StringVar(value="30401154 - BI")
        self.farmer_customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.farmer_customer_var, values=[
            # "30103270 - PBM1 (FARMERS/FM SCP)",
            # "30105314 - PBM1 (FARMERS/MESRA INDAI)",
            # "30202092 - PBM2 (FARMERS/FM SCP)",
            # "30203407 - PBM2 (FARMERS/MESRA INDAI)",
            "30401154 - BI"
        ])
        self.farmer_customer_dropdown.grid(row=0, column=1, padx=10, pady=(20, 10), sticky="ew")

        ctk.CTkLabel(tab, text="File CSV:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.farmer_csv_entry = ctk.CTkEntry(tab)
        self.farmer_csv_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        self.farmer_csv_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/biPurchaseOrder_3011714349.csv;C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/pbm1PurchaseOrder_3011722749 t.csv;C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/pbm1PurchaseOrderBatch_20240927094223.csv;C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/pbm2PurchaseOrder_3011601648 farmer - Copy.csv;C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/pbm2PurchaseOrder_3011601648 farmer.csv")
        ctk.CTkButton(tab, text="Browse", command=lambda: browse_files(self.farmer_csv_entry, "csv")).grid(row=1, column=2, padx=(0, 20), pady=10, sticky="e")

        ctk.CTkLabel(tab, text="File Excel Master Data:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.farmer_excel_entry = ctk.CTkEntry(tab)
        self.farmer_excel_entry.grid(row=2, column=1, padx=10, pady=10, sticky="ew")
        self.farmer_excel_entry.insert(0, "C:/Users/TOSHIBA PORTEGE Z30C/Desktop/program python/farmer/NKA smd umum.xls")
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
        self.hypermart_customer_var = ctk.StringVar(value="30400627 - BI (Hypermart Pentacity)")
        self.hypermart_customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.hypermart_customer_var, values=[
            "30400627 - BI (Hypermart Pentacity)",
            "30404436 - BI (Hypermart Plaza Balikpapan)",
            "30404435 - BI (Foodmart Supermarket Balikpapan)",
            "30405201 - BI (Hypermart Siloam)",
            # "30101002 - PBM1 (Hypermart - Matahari Putra Prima)",
            # "30100730 - PBM1 (Hypermart Big Mall)",
            # "30200527 - PBM2 - (Hypermart Big Mall)",
            # "30200728 - PBM2 - (Matahari)",
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
        self.hero_customer_var = ctk.StringVar(value="30400599 - BI (BLP - Hero)")
        self.hero_customer_dropdown = ctk.CTkOptionMenu(tab, variable=self.hero_customer_var, values=[
            "30400599 - BI (BLP - Hero)",
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

    def create_tab6(self, tab):
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

if __name__ == "__main__":
    app = App()
    app.mainloop()