import pandas as pd
import os
import re

print("🚀 Script dimulai...")

# --- 1. PENGATURAN PATH ---
folder_path_kuartal = r'C:\Users\ASUS\Documents\Investasi\Rekap Analisa Fundamental\ID'
file_path_sektor = r'C:\Users\ASUS\Documents\Investasi\Klasifikasi Sektor Subindustri.xlsx'
output_path_csv = r'C:\Users\ASUS\Documents\Investasi\data_fundamental_konsolidasi.csv'

# --- 2. PROSES SEMUA FILE EXCEL KUARTALAN ---
list_dataframes = []

print(f"Membaca file dari folder: '{folder_path_kuartal}'")

for filename in os.listdir(folder_path_kuartal):
    # Proses file dengan ekstensi .xlsx dan .xlsm
    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):
        try:
            match = re.search(r'(\d{4})\sKuartal\s(\d)', filename)
            if not match:
                print(f"  - ⚠️  Format nama file '{filename}' tidak sesuai, dilewati.")
                continue

            tahun = int(match.group(1))
            kuartal = int(match.group(2))
            
            print(f"  - Memproses file: {filename} (Tahun: {tahun}, Kuartal: {kuartal})")

            file_path = os.path.join(folder_path_kuartal, filename)

            df_data = pd.read_excel(file_path, sheet_name='Data')
            df_rekap = pd.read_excel(file_path, sheet_name='Rekap')

            # --- PERBAIKAN AKURAT DI SINI ---
            df_data.rename(columns={'Saham': 'Kode Emiten'}, inplace=True)
            df_rekap.rename(columns={'Saham': 'Kode Emiten'}, inplace=True)
            # ---------------------------------

            df_merged = pd.merge(df_data, df_rekap, on='Kode Emiten', how='outer')
            df_merged['Tahun'] = tahun
            df_merged['Kuartal'] = kuartal
            
            list_dataframes.append(df_merged)

        except Exception as e:
            print(f"  - ❌ Gagal memproses file {filename}. Error: {e}")

# --- 3. GABUNGKAN SEMUA DATA & TAMBAHKAN INFO SEKTOR ---
if not list_dataframes:
    print("\nTidak ada data yang berhasil diproses. Script berhenti.")
else:
    print("\nMenggabungkan semua data kuartalan...")
    df_master = pd.concat(list_dataframes, ignore_index=True)
    
    try:
        print(f"Menambahkan data klasifikasi dari: '{file_path_sektor}'")
        df_sektor = pd.read_excel(file_path_sektor)
        
        # Kolom di file sektor sudah benar 'Kode Emiten'
        kolom_sektor = ['Kode Emiten', 'Nama Entitas', 'Sektor', 'Subsektor', 'Industri', 'Subindustri']
        
        df_final = pd.merge(df_master, df_sektor[kolom_sektor], on='Kode Emiten', how='left')

        cols_to_move = ['Tahun', 'Kuartal', 'Kode Emiten', 'Nama Entitas', 'Sektor', 'Subsektor', 'Industri', 'Subindustri']
        df_final = df_final[cols_to_move + [col for col in df_final.columns if col not in cols_to_move]]
        
        # --- 4. SIMPAN HASIL KE CSV ---
        df_final.to_csv(output_path_csv, index=False)
        print(f"\n✅ Sukses! Data telah dikonsolidasi dan disimpan di '{output_path_csv}'")
        print(f"Total {len(df_final)} baris data telah diproses.")
    
    except FileNotFoundError:
        print(f"  - ❌ Gagal! File klasifikasi tidak ditemukan di '{file_path_sektor}'. Pastikan nama dan lokasinya benar.")
    except Exception as e:
        print(f"  - ❌ Gagal menambahkan info sektor. Error: {e}")