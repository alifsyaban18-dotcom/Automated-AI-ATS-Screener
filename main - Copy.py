import gspread
import os
import requests
import PyPDF2
import io
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google import genai
from google.genai import types

# ==========================================
# 1. SETUP KUNCI RAHASIA & PENGATURAN
# ==========================================
API_KEY_GEMINI = "YOUR_API_KEY"
client = genai.Client(api_key=API_KEY_GEMINI)
LINK_SHEETS = "YOUR_LINK"

EMAIL_PENGIRIM = "YOUR_EMAIL" 
PASSWORD_APLIKASI = "YOUR_PASSWORD" 

# ==========================================
# 2. FUNGSI UNTUK MEMBACA PDF
# ==========================================
def baca_cv_pdf(link_drive):
    if not link_drive or link_drive.strip() == "":
        return "TIDAK ADA LINK CV YANG DIUNGGAH"
    try:
        match = re.search(r'/d/([a-zA-Z0-9_-]+)', link_drive)
        if not match:
            match = re.search(r'id=([a-zA-Z0-9_-]+)', link_drive)
        if match:
            file_id = match.group(1)
        else:
            return f"[Gagal. Link tidak valid: {link_drive}]"

        url_download = f'https://drive.google.com/uc?id={file_id}&export=download'
        response = requests.get(url_download)
        file_memori = io.BytesIO(response.content)

        pembaca = PyPDF2.PdfReader(file_memori)
        teks_hasil = ""
        for halaman in pembaca.pages:
            teks_extrak = halaman.extract_text()
            if teks_extrak:
                teks_hasil += teks_extrak + "\n"
        
        if teks_hasil.strip() == "":
             return "[PDF kosong/berupa gambar]"
        return teks_hasil
    except Exception as e:
        return f"[Gagal membaca PDF. Alasan: {e}]"

# ==========================================
# 3. FUNGSI UNTUK MENGIRIM EMAIL OTOMATIS
# ==========================================
def kirim_email(email_tujuan, nama, skor, alasan):
    # Setup server Gmail
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(EMAIL_PENGIRIM, PASSWORD_APLIKASI)

    pesan = MIMEMultipart()
    pesan['From'] = EMAIL_PENGIRIM
    pesan['To'] = email_tujuan

    # Logika Penentuan Lolos / Tidak
    if skor >= 80:
        pesan['Subject'] = "Selamat! Anda Lolos Seleksi Tahap Awal"
        body = f"""
        Halo {nama},

        Selamat! Berdasarkan hasil screening awal oleh sistem ATS cerdas kami, Anda dinyatakan LOLOS ke tahap selanjutnya dengan skor {skor}/100.

        Berikut adalah catatan positif dari AI Reviewer kami mengenai profil Anda:
        "{alasan}"

        Silakan bergabung ke grup WhatsApp berikut untuk jadwal interview:
        https://chat.whatsapp.com/contoh_link_grup_lu

        Salam Hangat,
        Tim Rekrutmen (AI Bot)
        """
    else:
        pesan['Subject'] = "Update Status Lamaran Kerja Anda"
        body = f"""
        Halo {nama},

        Terima kasih atas ketertarikan Anda melamar di perusahaan kami. 
        Setelah melakukan screening otomatis secara saksama, mohon maaf saat ini kami belum bisa melanjutkan profil Anda ke tahap interview (Skor: {skor}/100).

        Sebagai bentuk transparansi, berikut adalah feedback dari AI Reviewer kami yang mungkin bisa membantu pengembangan karier Anda ke depannya:
        "{alasan}"

        Terima kasih dan semoga sukses di kesempatan berikutnya!

        Salam Hangat,
        Tim Rekrutmen (AI Bot)
        """

    pesan.attach(MIMEText(body, 'plain'))
    server.send_message(pesan)
    server.quit()

# ==========================================
# 4. PROSES UTAMA (NYAMBUNG SHEETS -> AI -> EMAIL)
# ==========================================
folder_sekarang = os.path.dirname(os.path.abspath(__file__))
lokasi_kunci = os.path.join(folder_sekarang, 'credentials.json')

try:
    print("Membuka database Google Sheets...")
    gc = gspread.service_account(filename=lokasi_kunci)
    sh = gc.open_by_url(LINK_SHEETS)
    worksheet = sh.sheet1
    data_pelamar = worksheet.get_all_records()

    print(f"Berhasil terhubung! Memeriksa {len(data_pelamar)} pelamar...\n")
    print("-" * 50)

    for index, pelamar in enumerate(data_pelamar):
        baris_excel = index + 2 
        kunci_kolom = list(pelamar.keys())
        
        nama = pelamar[kunci_kolom[1]]
        email_pelamar = pelamar[kunci_kolom[2]] # Mengambil email pelamar
        posisi = pelamar[kunci_kolom[4]]
        pengalaman = pelamar[kunci_kolom[5]]
        link_cv = pelamar[kunci_kolom[6]]

        status_pelamar = pelamar.get('Status', '')
        if status_pelamar == 'Sudah Diproses':
            print(f"⏩ Melewati {nama} (Sudah diproses)")
            continue

        if not nama:
            continue

        print(f"\n🔍 Menganalisis pelamar BARU: {nama}")
        isi_cv = baca_cv_pdf(link_cv)

        print("🤖 Menunggu hasil dari Otak Gemini...")
        prompt_hrd = f"""
        Kamu adalah HRD Manager Senior yang sangat kritis, analitis, dan objektif. 
        Tugasmu menyaring pelamar bernama {nama} untuk posisi {posisi}.

        Cerita pelamar: {pengalaman}
        Isi CV pelamar: {isi_cv}

        ATURAN KETAT (Patuhi 100%):
        1. Jika "isi CV pelamar" berisi pesan error/gagal baca/kosong, Skor MAKSIMAL 35.
        2. Jika bahasa pelamar bercanda/fiktif, penalti besar.
        3. Klaim di cerita WAJIB terbukti di dalam CV. Jika tidak sinkron, berikan skor rendah.

        Format jawaban:
        SKOR: [angka 0-100]
        ALASAN: [Berikan alasan profesional 3 kalimat, bahas langsung ke pelamarnya sebagai feedback HRD]
        """

        jawaban_ai = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt_hrd,
            config=types.GenerateContentConfig(temperature=0.0)
        )
        
        teks_ai = jawaban_ai.text
        
        cari_skor = re.search(r'SKOR:\s*(\d+)', teks_ai)
        cari_alasan = re.search(r'ALASAN:\s*(.*)', teks_ai, re.DOTALL)

        skor_final = int(cari_skor.group(1)) if cari_skor else 0
        alasan_final = cari_alasan.group(1).strip() if cari_alasan else teks_ai

        print(f"⭐️ SKOR DIDAPAT: {skor_final}")
        
        print("✍️ Menulis laporan ke Google Sheets...")
        worksheet.update_cell(baris_excel, 9, skor_final)
        worksheet.update_cell(baris_excel, 10, alasan_final)
        worksheet.update_cell(baris_excel, 11, "Sudah Diproses")
        
        print(f"✉️ Mengirim email keputusan ke {email_pelamar}...")
        try:
            kirim_email(email_pelamar, nama, skor_final, alasan_final)
            print("✅ Email berhasil terkirim!")
        except Exception as e:
            print(f"❌ Gagal mengirim email: {e}")

    print("\n🎉 MISI SELESAI! ATS BOT BERHASIL DIJALANKAN! 🎉")

except Exception as error:
    print(f"\n❌ Ada masalah:\n{error}")
