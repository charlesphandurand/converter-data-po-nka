import PySimpleGUI as sg
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
        required_columns = ["KODE AGLIS", "SALESMAN"]
        for col in required_columns:
            if col not in data.columns:
                logging.error(f"Kolom '{col}' tidak ditemukan dalam sheet 'KODE ITEM alfa'")
                return None
        
        logging.info(f"Kolom yang ditemukan: {data.columns.tolist()}")
        return data
    except Exception as e:
        logging.error(f"Error membaca file Excel: {str(e)}")
        return None

def process_edi_file(edi_file, excel_file, output_directory):
    # Baca file Excel
    try:
        df_excel = read_excel_file(excel_file)
        if df_excel is None:
            logging.error("Gagal membaca file Excel.")
            return None
        logging.info(f"File Excel berhasil dimuat. Shape: {df_excel.shape}")
    except Exception as e:
        logging.error(f"Error saat memuat file Excel: {str(e)}")
        return None

    # Baca file EDI
    try:
        with open(edi_file, 'r') as f:
            edi_content = f.read()
        edi_lines = edi_content.replace('\n', ' ').split()
        logging.info(f"File EDI berhasil dimuat. Total baris: {len(edi_lines)}")
    except Exception as e:
        logging.error(f"Error saat memuat file EDI: {str(e)}")
        return None

    output_lines = []
    output_filename = None
    for line in edi_lines:
        parts = line.split('|')
        if parts[0] == 'POHDR':
            try:
                edi_1 = parts[1]  # 1GZ1POF24004291
                output_filename = f"{edi_1}.txt"
                edi_3 = parts[2]  # 20240619 (tanggal yang kita inginkan)
                edi_6 = parts[5]  # 1GZ1 (kode untuk VLOOKUP)

                # VLOOKUP untuk SALESMAN dan KODE AGLIS
                kode_sales = df_excel.loc[df_excel['KODE AGLIS'] == edi_6, 'SALESMAN'].values
                kode_sales = kode_sales[0] if len(kode_sales) > 0 else 'Not Found'
                
                kode_aglis = df_excel.loc[df_excel['KODE AGLIS'] == edi_6, 'KODE AGLIS'].values
                kode_aglis = kode_aglis[0] if len(kode_aglis) > 0 else 'Not Found'

                # Format output line
                output_line = f"{edi_1};10300732;{kode_sales};{edi_3};{kode_aglis};20"
                output_lines.append(output_line)
                logging.debug(f"Baris output: {output_line}")
            except Exception as e:
                logging.error(f"Error saat memproses baris POHDR: {str(e)}")
            break  # Keluar dari loop setelah memproses POHDR

    if output_filename:
        output_file = os.path.join(output_directory, output_filename)
        # Tulis ke file output
        try:
            with open(output_file, 'w') as f:
                f.write('\n'.join(output_lines))
            logging.info(f"File output berhasil ditulis. Total baris: {len(output_lines)}")
        except Exception as e:
            logging.error(f"Error saat menulis file output: {str(e)}")
    
    return output_filename

# Kode GUI
layout = [
    [sg.Text("File EDI:"), sg.Input(), sg.FileBrowse(key="-EDI-")],
    [sg.Text("File Excel Master Data:"), sg.Input(), sg.FileBrowse(key="-EXCEL-")],
    [sg.Text("Direktori Output:"), sg.Input(), sg.FolderBrowse(key="-OUTPUT-DIR-")],
    [sg.Button("Proses"), sg.Button("Keluar")]
]

window = sg.Window("Konverter EDI ke TXT", layout)

while True:
    event, values = window.read()
    if event == sg.WINDOW_CLOSED or event == "Keluar":
        break
    elif event == "Proses":
        edi_file = values["-EDI-"]
        excel_file = values["-EXCEL-"]
        output_dir = values["-OUTPUT-DIR-"]
        
        if not edi_file or not excel_file or not output_dir:
            sg.popup("Silakan pilih semua file dan direktori yang diperlukan.")
        else:
            try:
                output_filename = process_edi_file(edi_file, excel_file, output_dir)
                if output_filename:
                    sg.popup(f"Konversi berhasil! File output: {output_filename}")
                else:
                    sg.popup("Konversi gagal. Silakan periksa log untuk detail.")
            except Exception as e:
                sg.popup_error(f"Terjadi kesalahan: {str(e)}")
            finally:
                print("Silakan periksa console untuk log detail.")

window.close()