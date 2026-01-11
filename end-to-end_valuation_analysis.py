import subprocess
import sys
import time  # tambahan

def _fmt_hms(seconds: float) -> str:
    return time.strftime("%H:%M:%S", time.gmtime(seconds))

def jalankan_semua_script():
    total_start_time = time.perf_counter()  # 1) mulai timer utama SEBELUM input
    try:
        # 1) Input sekali di awal
        tahun = input("Masukkan tahun (misal: 2025): ").strip()
        kuartal = input("Masukkan kuartal (1/2/3/4): ").strip()

        if not (tahun.isdigit() and len(tahun) == 4):
            print("Tahun harus 4 digit angka, contoh 2025.")
            return
        if kuartal not in {"1", "2", "3", "4"}:
            print("Kuartal harus 1, 2, 3, atau 4.")
            return

        # Input untuk dipipe ke stdin masing-masing script (akhiri newline)
        piped_input = f"{tahun}\n{kuartal}\n"

        # 2) Path dan file
        path_script_1_dir = r"C:\Users\ASUS\Documents\Investasi\Automation\Web Scarping"
        file_script_1 = "scarper_lk.py"

        path_script_2_dir = r"C:\Users\ASUS\Documents\Investasi\Automation"
        file_script_2 = "rekap_fundamental.py"
        file_script_3 = "Konsolidasi.py"

        # 3) Script 1
        print(f"\n--- Menjalankan {file_script_1} untuk Tahun {tahun} Kuartal {kuartal} ---")
        start_script_1 = time.perf_counter()
        subprocess.run(
            [sys.executable, file_script_1],
            cwd=path_script_1_dir,
            input=piped_input,
            text=True,
            check=True
        )
        durasi_script_1 = time.perf_counter() - start_script_1
        print(f"--- Selesai {file_script_1} (Durasi: {_fmt_hms(durasi_script_1)} ({durasi_script_1:.2f} detik)) ---")

        # 4) Script 2
        print(f"\n--- Menjalankan {file_script_2} untuk Tahun {tahun} Kuartal {kuartal} ---")
        start_script_2 = time.perf_counter()
        subprocess.run(
            [sys.executable, file_script_2],
            cwd=path_script_2_dir,
            input=piped_input,
            text=True,
            check=True
        )
        durasi_script_2 = time.perf_counter() - start_script_2
        print(f"--- Selesai {file_script_2} (Durasi: {_fmt_hms(durasi_script_2)} ({durasi_script_2:.2f} detik)) ---")

        # 5) Script 3 (tanpa input)
        print(f"\n--- Menjalankan {file_script_3} ---")
        start_script_3 = time.perf_counter()
        subprocess.run(
            [sys.executable, file_script_3],
            cwd=path_script_2_dir,
            check=True
        )
        durasi_script_3 = time.perf_counter() - start_script_3
        print(f"--- Selesai {file_script_3} (Durasi: {_fmt_hms(durasi_script_3)} ({durasi_script_3:.2f} detik)) ---")

        total_end_time = time.perf_counter()  # 3) selesai timer utama
        total_durasi = total_end_time - total_start_time

        print("\n=== SEMUA SCRIPT BERHASIL DIJALANKAN ===")

        # 4) Minta input jumlah file baru
        try:
            jumlah_file_str = input("Masukkan total jumlah file baru yang diproses: ").strip()
            jumlah_file = int(jumlah_file_str)
            if jumlah_file <= 0:
                print("Jumlah file harus > 0. Menggunakan 1 sebagai default.")
                jumlah_file = 1
        except ValueError:
            print("Input tidak valid, menggunakan 1 sebagai default.")
            jumlah_file = 1

        # 5) Rata-rata
        rata_rata_per_file = total_durasi / jumlah_file

        # 6) Ringkasan
        print("\n--- RINGKASAN WAKTU EKSEKUSI ---")
        print(f"Waktu Eksekusi {file_script_1}: {_fmt_hms(durasi_script_1)} ({durasi_script_1:.2f} detik)")
        print(f"Waktu Eksekusi {file_script_2}: {_fmt_hms(durasi_script_2)} ({durasi_script_2:.2f} detik)")
        print(f"Waktu Eksekusi {file_script_3}: {_fmt_hms(durasi_script_3)} ({durasi_script_3:.2f} detik)")
        print("------------------------------------------")
        print(f"Total Waktu Eksekusi: {_fmt_hms(total_durasi)} ({total_durasi:.2f} detik)")
        print(f"Jumlah file baru diproses: {jumlah_file}")
        print(f"Rata-rata waktu per file baru: {rata_rata_per_file:.2f} detik/file")

    except subprocess.CalledProcessError as e:
        print("\n!!! ERROR: Eksekusi script gagal !!!")
        print(f"Command: {e.cmd}")
        print(f"Return code: {e.returncode}")
        print("Lihat output di atas untuk detail error dari script yang gagal.")
    except FileNotFoundError as e:
        print("\n!!! ERROR: File script tidak ditemukan !!!")
        print(f"Error: {e}")
    except Exception as e:
        print("\n!!! ERROR: Terjadi kesalahan yang tidak terduga !!!")
        print(f"Error: {e}")

if __name__ == "__main__":
    jalankan_semua_script()