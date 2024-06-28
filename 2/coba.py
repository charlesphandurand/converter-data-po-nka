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
        edi_lines = edi_content.split('\n')
        logging.info(f"File EDI berhasil dimuat. Total baris: {len(edi_lines)}")
    except Exception as e:
        logging.error(f"Error saat memuat file EDI: {str(e)}")
        return None

    output_lines = []
    output_filename = None
    pohdr_line = None
    lin_line = None

    for line in edi_lines:
        parts = line.strip().split('|')
        if parts[0] == 'POHDR':
            pohdr_line = parts
        elif parts[0] == 'LIN':
            lin_line = parts
            break  # We only need the first LIN line

    if pohdr_line and lin_line:
        try:
            edi_1 = pohdr_line[1] if len(pohdr_line) > 1 else 'Unknown'
            output_filename = f"{edi_1}.txt"
            edi_3 = pohdr_line[2] if len(pohdr_line) > 2 else 'Unknown'
            edi_6_lin = lin_line[5] if len(lin_line) > 5 else 'Unknown'

            logging.debug(f"POHDR: {pohdr_line}")
            logging.debug(f"LIN: {lin_line}")
            logging.debug(f"EDI_1: {edi_1}, EDI_3: {edi_3}, EDI_6_LIN: {edi_6_lin}")

            # VLOOKUP untuk SALESMAN
            salesman = df_excel.loc[df_excel['BARCODE'] == edi_6_lin, 'SALESMAN'].values
            salesman = int(salesman[0]) if len(salesman) > 0 else 'Not Found'

            # VLOOKUP untuk KODE AGLIS
            kode_aglis = df_excel.loc[df_excel['BARCODE'] == edi_6_lin, 'KODE AGLIS'].values
            kode_aglis = int(kode_aglis[0]) if len(kode_aglis) > 0 else 'Not Found'

            # Format output line
            output_line = f"{edi_1};10300732;{salesman};{edi_3};{kode_aglis};20"
            output_lines.append(output_line)
            logging.debug(f"Baris output: {output_line}")
        except Exception as e:
            logging.error(f"Error saat memproses baris: {str(e)}")
            logging.exception("Traceback:")

    if output_filename and output_lines:
        output_file = os.path.join(output_directory, output_filename)
        # Tulis ke file output
        try:
            with open(output_file, 'w') as f:
                f.write('\n'.join(output_lines))
            logging.info(f"File output berhasil ditulis. Total baris: {len(output_lines)}")
        except Exception as e:
            logging.error(f"Error saat menulis file output: {str(e)}")
    else:
        logging.warning("Tidak ada data yang diproses atau nama file output tidak ditentukan.")
    
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