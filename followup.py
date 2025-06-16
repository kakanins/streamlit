# CARA JALANKAN: streamlit run followup.py --server.maxUploadSize=1024

import streamlit as st
import pandas as pd
from io import BytesIO
import datetime

st.title("📞 Otomatisasi Follow-Up (File Lama Saja)")
st.write("Upload file Excel lama Anda (mis. hasil harian tele).")

today_date = datetime.date.today()

uploaded_files = st.file_uploader("Upload beberapa file Excel (.xlsx) lama", type=["xlsx"], accept_multiple_files=True)

mapping_fu = {
    "Tanya Pasangan": 1,
    "Tanya-Tanya": 1,
    "Belum Minat": 3,
    "Angsuran Masih Panjang": "Next Month",
    "Plafond Rendah": 2,
    "Tidak Aktif": "Next Month",
    "Tidak Terdaftar": None,
    "Tidak Diangkat": 1,
    "Dialihkan/Sibuk": "Next Month",
    "Janji Telpon Ulang": 1,
    "Bunga Tinggi": 2,
}

def hitung_tgl_fu(row):
    """
    Menghitung tanggal follow-up berdasarkan kolom 'RESULT' dan 'FollowUp(Hari)'.
    """
    if pd.isnull(row["FollowUp(Hari)"]):
        return pd.NaT # Not a Time, untuk nilai yang tidak perlu follow-up
    elif row["FollowUp(Hari)"] == "Next Month":
        # Menambahkan satu bulan ke tanggal TGL
        return (pd.to_datetime(row["TGL"]) + pd.DateOffset(months=1)).date()
    else:
        # Menambahkan jumlah hari sesuai 'FollowUp(Hari)'
        return (pd.to_datetime(row["TGL"]) + pd.Timedelta(days=int(row["FollowUp(Hari)"]))).date()

if uploaded_files:
    # Dalam skenario ini, kita asumsikan semua file adalah 'file_lama'
    # Tidak perlu memisahkan berdasarkan '_baru' di sini.
    
    jumlah_tele = st.number_input("Jumlah Tele Baru", min_value=1, value=2, step=1, 
                                  help="Masukkan berapa banyak tele baru yang akan menerima data follow-up.")
    nama_tele_baru = [st.text_input(f"Nama Tele Baru {i+1}", value=f"Tele_{i+1}") for i in range(jumlah_tele)]

    proses = st.button("🚀 Proses Semua File Follow-Up Lama")

    if proses:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_parts_lama = []
            nama_sheet_tele_lama_list = [] # List untuk menyimpan nama sheet dari file lama

            # Loop melalui setiap file yang diunggah
            for file in uploaded_files:
                xls = pd.ExcelFile(file)
                # Loop melalui setiap sheet dalam file Excel
                for sheet_name in xls.sheet_names:
                    df = xls.parse(sheet_name)
                    
                    # Tulis DataFrame asli ke output tanpa modifikasi
                    # Ini untuk memastikan sheet asli tetap ada dan bersih
                    df.to_excel(writer, sheet_name=sheet_name, index=False) 
                    
                    # Buat salinan DataFrame untuk pemrosesan lebih lanjut
                    df_copy = df.copy() 
                    df_copy["TELE_LAMA"] = sheet_name # Tambahkan kolom TELE_LAMA berdasarkan nama sheet
                    df_copy["Tanggal Upload"] = today_date # Tambahkan kolom Tanggal Upload
                    df_parts_lama.append(df_copy)
                    nama_sheet_tele_lama_list.append(sheet_name) # Simpan nama sheet sebagai 'TELE_LAMA'

            # Jika ada data dari file lama yang diunggah
            if df_parts_lama:
                df_lanjutan = pd.concat(df_parts_lama, ignore_index=True)
                df_lanjutan["RESULT"] = df_lanjutan["RESULT"].astype(str).str.strip().str.title()
                df_lanjutan["FollowUp(Hari)"] = df_lanjutan["RESULT"].map(mapping_fu)
                # Pastikan kolom TGL ada dan dikonversi ke tanggal
                df_lanjutan["TGL"] = pd.to_datetime(df_lanjutan.get("TGL", today_date), errors="coerce").dt.date
                df_lanjutan["Tanggal FollowUp"] = df_lanjutan.apply(hitung_tgl_fu, axis=1)

                # Pisahkan data yang perlu di-follow-up dan yang tidak
                df_fu = df_lanjutan[~df_lanjutan["FollowUp(Hari)"].isnull()].copy()
                df_tidak = df_lanjutan[df_lanjutan["FollowUp(Hari)"].isnull()].copy()

                # Distribusi merata ke tele baru
                if len(nama_tele_baru) > 0 and not df_fu.empty:
                    block_size = len(df_fu) // len(nama_tele_baru)
                    remainder = len(df_fu) % len(nama_tele_baru)
                    tele_baru_labels = []
                    start = 0
                    for i, nama_tele in enumerate(nama_tele_baru):
                        end = start + block_size + (1 if i < remainder else 0)
                        tele_baru_labels.extend([nama_tele] * (end - start))
                        start = end
                    df_fu["TELE_BARU"] = tele_baru_labels
                else:
                    df_fu["TELE_BARU"] = "N/A" # Jika tidak ada tele baru atau df_fu kosong

                df_fu["TELE_LAMA"] = df_fu["TELE_LAMA"].astype(str)

                # Urutan kolom, pastikan TELE_LAMA dan TELE_BARU di akhir
                cols = [c for c in df_fu.columns if c not in ["TELE_LAMA", "TELE_BARU"]] + ["TELE_LAMA", "TELE_BARU"]
                df_fu = df_fu[cols]

                # Tulis hasil follow-up ke sheet-sheet terpisah
                df_fu.to_excel(writer, sheet_name="FU Lanjutan", index=False)
                df_fu[df_fu["FollowUp(Hari)"] == 1].to_excel(writer, sheet_name="FU Besok", index=False)
                df_fu[df_fu["FollowUp(Hari)"] == 2].to_excel(writer, sheet_name="FU Lusa", index=False)
                df_fu[df_fu["FollowUp(Hari)"] == 3].to_excel(writer, sheet_name="FU 3 Hari", index=False)
                df_fu[df_fu["FollowUp(Hari)"] == "Next Month"].to_excel(writer, sheet_name="FU Next Month", index=False)

                # Tulis data untuk setiap tele baru
                for tele in nama_tele_baru:
                    df_chunk = df_fu[df_fu["TELE_BARU"] == tele].copy()
                    df_chunk.to_excel(writer, sheet_name=tele, index=False)

                # Tulis data yang tidak bisa di-follow-up
                if not df_tidak.empty:
                    df_tidak["TELE_BARU"] = None # Karena tidak ada pembagian ke tele baru untuk data ini
                    df_tidak.to_excel(writer, sheet_name="Tidak Bisa FU", index=False)
            else:
                st.warning("⚠️ Tidak ada data follow-up lama yang ditemukan dari file yang diunggah.")

        st.success("✅ Semua file berhasil diproses!")
        st.download_button("📥 Download Excel FU", data=output.getvalue(), file_name="FU_Output_Lama.xlsx")
