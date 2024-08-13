import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog

def bersihkan_spasi_dan_buat_excel():
    # Membuka dialog file untuk memilih file input
    root = tk.Tk()
    root.withdraw()
    file_input = filedialog.askopenfilename(title="Pilih File Input", filetypes=[("EDI Files", "*.edi"), ("All Files", "*.*")])
    if not file_input:
        return

    # Membuka dialog file untuk menyimpan file output
    file_output = filedialog.asksaveasfilename(title="Simpan File Output", defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
    if not file_output:
        return

    # Membuka dialog untuk memasukkan delimiter
    delimiter = simpledialog.askstring("Masukkan Delimiter", "Masukkan delimiter yang digunakan (misalnya: ','):")
    if not delimiter:
        return

    # Membaca file teks
    with open(file_input, 'r') as file:
        data = file.read().replace('\n', '')

    # Membersihkan spasi
    data_bersih = ' '.join(data.split())

    # Memisahkan data berdasarkan delimiter
    baris = data_bersih.split(delimiter)

    # Membuat DataFrame dari daftar baris
    df = pd.DataFrame([baris])

    # Menyimpan DataFrame ke file Excel
    df.to_excel(file_output, index=False, header=False)
    print(f"File Excel {file_output} berhasil dibuat.")

# Menjalankan program
if __name__ == "__main__":
    bersihkan_spasi_dan_buat_excel()