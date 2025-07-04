import pandas as pd
import os

# --- File path & sheet ---
file_path = 'sample_data\data_excel.xlsb'
sheet_name = 'Global'

# --- Load Excel dari baris ke-4 karena data mulai di I5 ---
df = pd.read_excel(file_path, sheet_name=sheet_name, engine='pyxlsb', header=3)

# --- Bersihkan nama kolom ---
df.columns = df.columns.map(lambda x: str(x).strip().replace('\n', ''))

# --- Cari kolom 'art' (kode artikel) ---
art_col = next((col for col in df.columns if col.lower() == 'art'), None)
if art_col is None:
    raise ValueError("❌ Kolom 'art' tidak ditemukan!")

# --- Ambil semua kolom yang bernama 'Post1' (bisa muncul berkali-kali) ---
post1_cols = [col for col in df.columns if col.strip().startswith('Post1')]

if not post1_cols:
    raise ValueError("❌ Tidak ditemukan kolom yang mengandung 'Post1'!")

# --- Ambil data art + semua Post1 ---
result_df = df[[art_col] + post1_cols].copy()

# --- Drop baris kosong art ---
result_df = result_df.dropna(subset=[art_col])

# --- Ganti NaN jadi 0 ---
result_df.fillna(0, inplace=True)

# --- Rename kolom Post1 jadi Post1_1, Post1_2, dst. ---
post_col_names = [f'Post1_{i+1}' for i in range(len(post1_cols))]
result_df.columns = [art_col] + post_col_names

# --- Simpan ke folder output ---
output_path = os.path.join('output', 'output_sample.xlsx')
result_df.to_excel(output_path, index=False)

print(f"✅ File berhasil disimpan ke: {output_path}")
