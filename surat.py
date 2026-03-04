import pandas as pd
from docxtpl import DocxTemplate
import os
import subprocess
import time
import sys

# --- KONFIGURASI PATH ---
excel_file = "data/data_debitur_fake.xlsx"
template_path = "templates/Template Surat.docx"
output_dir = "output"
# Path LibreOffice (Default macOS)
libreoffice_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"


# --- FUNGSI HELPER TANGGUH ---
def format_rupiah(v):
    """Mengubah angka menjadi format ribuan titik (1.000.000)."""
    try:
        # Cek apakah v adalah angka dan bukan NaN (kosong di Excel)
        if isinstance(v, (int, float)) and not pd.isna(v):
            return f"{v:,.0f}".replace(",", ".")
        return v
    except:
        return v


# --- PERSIAPAN LOKASI ---
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

if not os.path.exists(libreoffice_path):
    print(f"❌ ERROR: LibreOffice tidak ditemukan di {libreoffice_path}")
    sys.exit()

# --- LOAD DATA ---
try:
    df = pd.read_excel(excel_file)
    print(f"📊 Memproses {len(df)} data nasabah...")
except Exception as e:
    print(f"❌ Gagal membaca Excel: {e}")
    sys.exit()

start_time = time.time()
generated_docx_files = []

# --- TAHAP 1: RENDER WORD ---
print("📝 Tahap 1: Membuat file Word (Docx)...")

for index, row in df.iterrows():
    try:
        doc_tpl = DocxTemplate(template_path)
        raw_data = row.to_dict()
        context = {}

        for k, v in raw_data.items():
            clean_key = str(k).replace(" ", "_")

            # LOGIKA UTAMA: Format semua angka, KECUALI kolom "Nomor"
            if isinstance(v, (int, float)) and not pd.isna(v) and "nomor" not in str(k).lower():
                context[clean_key] = format_rupiah(v)
            else:
                context[clean_key] = v

        doc_tpl.render(context)

        # Sanitasi nama file agar aman dari karakter aneh
        nama_nasabah = "".join([c for c in str(row['Nama Nasabah']) if c.isalnum() or c == ' ']).strip()
        file_path = os.path.join(output_dir, f"Somasi_{nama_nasabah}.docx")

        doc_tpl.save(file_path)
        generated_docx_files.append(file_path)

    except Exception as e:
        print(f"⚠️ Error di baris {index + 1}: {e}")

# --- TAHAP 2: KONVERSI PDF MASSAL ---
if generated_docx_files:
    print(f"🔄 Tahap 2: Konversi {len(generated_docx_files)} file ke PDF...")
    try:
        command = [
                      libreoffice_path, '--headless', '--convert-to', 'pdf',
                      '--outdir', output_dir
                  ] + generated_docx_files

        subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

        # --- TAHAP 3: PEMBERSIHAN ---
        print("🧹 Tahap 3: Menghapus file Word sementara...")
        for docx_path in generated_docx_files:
            if os.path.exists(docx_path):
                os.remove(docx_path)

        durasi = time.time() - start_time
        print(f"\n✨ SELESAI DALAM {durasi:.2f} DETIK!")
        print(f"📂 Hasil PDF tersedia di folder: {output_dir}")

    except Exception as e:
        print(f"❌ Gagal konversi PDF: {e}")