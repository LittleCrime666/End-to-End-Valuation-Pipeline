import os
import time
import urllib.parse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException, ElementClickInterceptedException, StaleElementReferenceException

# Input dinamis tahun dan kuartal (periode cukup 1/2/3/4)
target_tahun = input("Masukkan tahun (misal: 2025): ").strip()
target_quarter = input("Masukkan kuartal (1/2/3/4): ").strip()
quarter_map = {'1': 'Triwulan 1', '2': 'Triwulan 2', '3': 'Triwulan 3', '4': 'Triwulan 4'}
target_periode = quarter_map.get(target_quarter, f'Triwulan {target_quarter}')

# Setup folder download dinamis
DOWNLOAD_FOLDER = fr"C:\Users\ASUS\Documents\Investasi\Laporan Keuangan\{target_tahun} Q{target_quarter}"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_FOLDER,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

def js_click(el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.2)
    driver.execute_script("arguments[0].click();", el)

def dismiss_overlays():
    for xp in [
        "//button[contains(., 'Terima') or contains(., 'Setuju') or contains(., 'Accept')]",
        "//*[@id='onetrust-accept-btn-handler']",
        "//button[contains(@class,'btn-close') or contains(@class,'close')]",
    ]:
        try:
            btns = driver.find_elements(By.XPATH, xp)
            if btns:
                js_click(btns[0])
                time.sleep(0.3)
        except Exception:
            pass

def wait_download(file_path: str, timeout=90):
    cr = file_path + ".crdownload"
    for _ in range(timeout):
        if os.path.exists(file_path):
            return True
        if os.path.exists(cr):
            time.sleep(0.5)
        else:
            time.sleep(0.5)
    return False

# Map kuartal ke token yang muncul pada href/filename
roman_map = {'1': 'I', '2': 'II', '3': 'III', '4': 'IV'}
roman = roman_map.get(target_quarter, target_quarter)
tw_token = f"/TW{target_quarter}/"
roman_token = f"-{roman}-"  # contoh: -III-

try:
    print("🚀 Scarper LK mulai...")
    print("Membuka halaman IDX...")
    driver.get("https://idx.co.id/id/perusahaan-tercatat/laporan-keuangan-dan-tahunan/")
    wait = WebDriverWait(driver, 25)

    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(1)
    dismiss_overlays()

    # Pilih tahun
    tahun_label = wait.until(EC.element_to_be_clickable((By.XPATH, f"//label[contains(normalize-space(.), '{target_tahun}')]")))
    try:
        tahun_label.click()
    except ElementClickInterceptedException:
        js_click(tahun_label)
    time.sleep(0.4)

    # Pilih kuartal: dukung label 'Triwulan 3' dan 'TW3'
    periode_label = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        f"//label[contains(normalize-space(.), 'Triwulan {target_quarter}')] | //label[contains(normalize-space(.), 'TW{target_quarter}')]"
    )))
    try:
        periode_label.click()
    except ElementClickInterceptedException:
        js_click(periode_label)
    time.sleep(0.4)

    # Terapkan
    terapkan_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Terapkan')]")))
    try:
        terapkan_btn.click()
    except ElementClickInterceptedException:
        js_click(terapkan_btn)

    total_files_found = 0
    total_files_downloaded = 0

    # Tunggu tabel
    wait.until(EC.presence_of_element_located((By.XPATH, "//table")))
    time.sleep(1)

    while True:
        dismiss_overlays()

        # Kumpulkan href lebih dulu untuk menghindari stale
        elems = driver.find_elements(By.XPATH, "//a[contains(@href, '.xlsx')]")
        hrefs = []
        for e in elems:
            try:
                href = e.get_attribute("href")
                if href:
                    hrefs.append(href)
            except StaleElementReferenceException:
                continue

        # Filter hanya kuartal target (TWx atau -ROMAN-)
        filtered = [h for h in hrefs if (tw_token in h) or (roman_token in os.path.basename(urllib.parse.urlsplit(h).path))]
        print(f"Link .xlsx halaman ini (match Q{target_quarter}): {len(filtered)}")
        total_files_found += len(filtered)

        # Download via tab baru (tanpa klik elemen, hindari intercepted/stale)
        main_handle = driver.current_window_handle
        for url in filtered:
            file_name = os.path.basename(urllib.parse.urlsplit(url).path)
            file_path = os.path.join(DOWNLOAD_FOLDER, file_name)
            if os.path.exists(file_path):
                print(f"File sudah ada: {file_name}")
                continue

            print(f"Memulai download: {file_name}")
            driver.execute_script("window.open('about:blank','_blank');")
            driver.switch_to.window(driver.window_handles[-1])
            driver.get(url)

            ok = wait_download(file_path, timeout=120)
            driver.close()
            driver.switch_to.window(main_handle)

            if ok:
                print(f"Berhasil mengunduh: {file_name}")
                total_files_downloaded += 1
            else:
                print(f"Gagal mengunduh: {file_name}")

        # Next page
        try:
            next_btn = driver.find_element(By.XPATH, "//button[contains(@class, 'btn-arrow') and contains(@class, '--next')]")
            old_table = driver.find_element(By.XPATH, "//table")
            if next_btn.is_enabled():
                js_click(next_btn)
                wait.until(EC.staleness_of(old_table))
                wait.until(EC.presence_of_element_located((By.XPATH, "//table")))
                time.sleep(0.8)
                print("Berpindah ke halaman berikutnya...")
            else:
                print("Mencapai halaman terakhir.")
                break
        except NoSuchElementException:
            print("Mencapai halaman terakhir.")
            break

finally:
    driver.quit()
    print("Semua proses download selesai.")
    print(f"Total link match kuartal {target_quarter}: {total_files_found}")
    print(f"Total file berhasil diunduh: {total_files_downloaded}")
