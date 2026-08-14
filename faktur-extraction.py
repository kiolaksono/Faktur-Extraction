import os
import re
import pandas as pd
import pdfplumber

def extract_faktur_data(pdf_path):
    """
    Fungsi untuk mengekstrak Nomor Faktur Pajak, Nomor Invoice, dan Nomor SAP dari satu file PDF.
    """
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
                
    # 1. Ekstraksi Kode dan Nomor Seri Faktur Pajak (16 digit angka)
    faktur_pajak_match = re.search(r"Kode dan Nomor Seri Faktur Pajak:\s*(\d+)", full_text)
    nomor_faktur_pajak = faktur_pajak_match.group(1) if faktur_pajak_match else "Tidak Ditemukan"
    
    # 2 & 3. Ekstraksi Nomor Invoice dan Nomor SAP dari baris Referensi
    # Pola teks: (Referensi: LL-2000.005174/VII/26 ( 4900005327 ))
    referensi_match = re.search(r"Referensi:\s*([^\(\n]+?)\s*\(\s*(\d+)\s*\)", full_text)
    
    if referensi_match:
        nomor_invoice = referensi_match.group(1).strip()
        nomor_sap = referensi_match.group(2).strip()
    else:
        nomor_invoice = "Tidak Ditemukan"
        nomor_sap = "Tidak Ditemukan"
        
    return {
        "Nama File": os.path.basename(pdf_path),
        "Nomor Faktur Pajak": nomor_faktur_pajak,
        "Nomor Invoice / Referensi": nomor_invoice,
        "Nomor SAP": nomor_sap
    }

def process_all_pdfs(folder_path, output_excel):
    """
    Fungsi untuk memproses seluruh file PDF di dalam folder dan menyimpan hasilnya ke file Excel.
    """
    results = []
    
    # Cek semua file dalam folder
    for file_name in os.listdir(folder_path):
        if file_name.lower().endswith('.pdf'):
            pdf_path = os.path.join(folder_path, file_name)
            print(f"Memproses: {file_name}...")
            
            try:
                data = extract_faktur_data(pdf_path)
                results.append(data)
            except Exception as e:
                print(f"Gagal memproses {file_name}: {e}")
                
    if results:
        # Convert ke DataFrame pandas
        df = pd.DataFrame(results)
        
        # Tambahkan kolom nomor urut
        df.insert(0, 'No', range(1, 1 + len(df)))
        
        # Simpan ke Excel
        df.to_excel(output_excel, index=False)
        print(f"\nSelesai! Total {len(results)} file berhasil diekstrak dan disimpan ke: {output_excel}")
    else:
        print("Tidak ada file PDF yang berhasil diproses.")

# ==========================================
# Cara Penggunaan:
# ==========================================
if __name__ == "__main__":
    # Ganti 'folder_pdf' sesuai dengan nama folder tempat Anda menyimpan file-file PDF
    FOLDER_PDF = "./"  
    FILE_OUTPUT = "Rekap_Data_Faktur_Pajak.xlsx"
    
    process_all_pdfs(FOLDER_PDF, FILE_OUTPUT)
