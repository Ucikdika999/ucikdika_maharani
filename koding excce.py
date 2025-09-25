import json
import pandas as pd

# Baca file JSON
with open("UCIK DIKA MAHARANI_V3925016_TI B_LDTI.json", "r", encoding="utf-8") as f:
    data = f.read()

# Karena format JSON tidak standar (ada beberapa blok array dengan "nama"),
# kita bisa parsing manual.
# Contoh ekstraksi sederhana:
import re
import ast

sections = re.split(r'"nama":', data)[1:]  # pisahkan berdasarkan "nama":
parsed_data = {}

for section in sections:
    # Ambil nama tabel
    nama = section.split("\n", 1)[0].strip().strip('" ,')
    
    # Cari array datanya
    match = re.search(r"\[(.*?)\]", section, re.S)
    if match:
        try:
            arr = ast.literal_eval("[" + match.group(1) + "]")
            parsed_data[nama] = arr
        except Exception as e:
            print(f"Gagal parsing {nama}: {e}")

# Simpan ke Excel dengan sheet per kategori
with pd.ExcelWriter("output_data.xlsx", engine="openpyxl") as writer:
    for nama, records in parsed_data.items():
        df = pd.DataFrame(records)
        df.to_excel(writer, sheet_name=nama[:30], index=False)

print("File Excel berhasil dibuat: output_data.xlsx")
