import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import itertools # Import untuk fungsi cycle

st.title("📞 Otomatisasi Follow-Up dan Pembagian Tele (Multi-File)")
st.write("Upload file Excel dengan `_baru` di nama file dan file Excel lama.")

today_date = datetime.date.today()

uploaded_files = st.file_uploader("Upload beberapa file Excel (.xlsx)", type=["xlsx"], accept_multiple_files=True)

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
        return pd.NaT
    elif row["FollowUp(Hari)"] == "Next Month":
        # Menambahkan satu bulan ke tanggal TGL
        return (pd.to_datetime(row["TGL"]) + pd.DateOffset(months=1)).date()
    else:
        # Menambahkan jumlah hari sesuai 'FollowUp(Hari)'
        return (pd.to_datetime(row["TGL"]) + pd.Timedelta(days=int(row["FollowUp(Hari)"]))).date()

def assign_tele_baru(df_to_assign, nama_tele_baru):

    if df_to_assign.empty or not nama_tele_baru:
        if "TELE_BARU" not in df_to_assign.columns: # Pastikan kolom ada sebelum mencoba mengisinya
            df_to_assign["TELE_BARU"] = "N/A"
        else:
            df_to_assign["TELE_BARU"] = df_to_assign["TELE_BARU"].fillna("N/A")
        return df_to_assign

    df_to_assign["TELE_BARU"] = None 

    # Tahap pertama: Mencoba menetapkan ke tele yang sama jika TELE_LAMA cocok dengan nama_tele_baru
    if "TELE_LAMA" in df_to_assign.columns:
        for tele in nama_tele_baru:
            direct_match_mask = (df_to_assign["TELE_LAMA"] == tele) & (df_to_assign["TELE_BARU"].isnull())
            df_to_assign.loc[direct_match_mask, "TELE_BARU"] = tele

    # Tahap kedua: Round-robin untuk baris yang masih belum ditetapkan
    unassigned_indices = df_to_assign[df_to_assign["TELE_BARU"].isnull()].index.tolist()
    
    if unassigned_indices:
        tele_iterator = itertools.cycle(nama_tele_baru)
        for idx in unassigned_indices:
            df_to_assign.loc[idx, "TELE_BARU"] = next(tele_iterator)

    df_to_assign["TELE_BARU"].fillna("N/A", inplace=True)
    
    return df_to_assign


if uploaded_files:
    file_baru = [f for f in uploaded_files if "_baru" in f.name.lower()]
    file_lama = [f for f in uploaded_files if "_baru" not in f.name.lower()]

    jumlah_tele = st.number_input("Jumlah Tele Baru", min_value=1, value=2, step=1,
                                  help="Masukkan berapa banyak tele baru yang akan menerima data follow-up.")
    nama_tele_baru = [st.text_input(f"Nama Tele Baru {i+1}", value=f"Tele_{i+1}") for i in range(jumlah_tele)]
    nama_tele_baru.sort()

    proses = st.button("🚀 Proses Semua File")

    if proses:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            nama_sheet_tele_lama = [] 
            df_parts_lama = []

            # Tahap 1: Memproses semua file yang diunggah
            for file in uploaded_files:
                xls = pd.ExcelFile(file)
                for sheet in xls.sheet_names:
                    df = xls.parse(sheet)
                    # Tulis DataFrame asli ke output tanpa modifikasi
                    df.to_excel(writer, sheet_name=sheet, index=False) 
                    
                    # Jika ini adalah file lama (tidak ada '_baru' di namanya)
                    if "_baru" not in file.name.lower():
                        df_copy = df.copy() 
                        df_copy["TELE_LAMA"] = sheet 
                        df_copy["Tanggal Upload"] = today_date 
                        df_parts_lama.append(df_copy)
                        nama_sheet_tele_lama.append(sheet) 

            jumlah_tele_lama = len(nama_sheet_tele_lama)

            # === MASTER BARU ===
            if file_baru:
                df_master = pd.ExcelFile(file_baru[0]).parse(0) 
                df_master["Tanggal Upload"] = today_date
                df_master["RESULT"] = df_master["RESULT"].astype(str).str.strip().str.title()
                df_master["FollowUp(Hari)"] = df_master["RESULT"].map(mapping_fu)
                df_master["TGL"] = pd.to_datetime(df_master.get("TGL", today_date), errors="coerce").dt.date
                df_master["Tanggal FollowUp"] = df_master.apply(hitung_tgl_fu, axis=1)

                # --- PERBAIKAN: Jika TELE_LAMA tidak ada di file master baru, inisialisasi dengan "N/A" ---
                if "TELE_LAMA" not in df_master.columns:
                    df_master["TELE_LAMA"] = "N/A"
                # Jika ada, biarkan nilai aslinya, tidak perlu ditimpa.

                df_fu_only = df_master[~df_master["FollowUp(Hari)"].isnull()].copy()
                df_tidak_fu = df_master[df_master["FollowUp(Hari)"].isnull()].copy()
                
                # *** PERBAIKAN UTAMA: Distribusi TELE_BARU untuk file baru menggunakan round-robin murni ***
                if len(nama_tele_baru) > 0 and not df_fu_only.empty:
                    tele_baru_iterator = itertools.cycle(nama_tele_baru)
                    df_fu_only["TELE_BARU"] = [next(tele_baru_iterator) for _ in range(len(df_fu_only))]
                else:
                    df_fu_only["TELE_BARU"] = "N/A"

                if "TELE_BARU" not in df_master.columns:
                    df_master["TELE_BARU"] = "N/A" # Inisialisasi jika belum ada

                if "TELE_BARU" not in df_tidak_fu.columns:
                    df_tidak_fu["TELE_BARU"] = None # Atau "N/A" jika prefer
                else:
                    df_tidak_fu["TELE_BARU"].fillna(None, inplace=True) # Isi NaN dengan None

                # Gabungkan kembali df_fu_only dan df_tidak_fu ke dalam satu DataFrame untuk "Data_Terproses_Baru"
                df_processed_master_full = pd.concat([df_fu_only, df_tidak_fu], ignore_index=True)
                
                # Sorting the combined processed master data for consistent output
                if "TELE_BARU" in df_processed_master_full.columns:
                    df_processed_master_full = df_processed_master_full.sort_values(by=["TELE_BARU", "TELE_LAMA"], na_position='last').reset_index(drop=True)
                
                # Mengubah nama sheet "Master_Data7k" menjadi "Data_Terproses_Baru"
                df_processed_master_full.to_excel(writer, sheet_name="Data_Terproses_Baru", index=False)

                # Distribusi data FU dari master baru ke tele baru (dari df_fu_only)
                for tele in nama_tele_baru: # nama_tele_baru sudah diurutkan
                    df_chunk = df_fu_only[df_fu_only["TELE_BARU"] == tele].copy()
                    if not df_chunk.empty:
                        df_chunk.to_excel(writer, sheet_name=tele, index=False)

                # Tulis data FU dari master baru ke sheet FU berdasarkan hari (menggunakan df_fu_only asli)
                if not df_fu_only.empty:
                    df_fu_only[df_fu_only["FollowUp(Hari)"] == 1].to_excel(writer, sheet_name="FU Besok", index=False)
                    df_fu_only[df_fu_only["FollowUp(Hari)"] == 2].to_excel(writer, sheet_name="FU Lusa", index=False)
                    df_fu_only[df_fu_only["FollowUp(Hari)"] == 3].to_excel(writer, sheet_name="FU 3 Hari", index=False)
                    df_fu_only[df_fu_only["FollowUp(Hari)"] == "Next Month"].to_excel(writer, sheet_name="FU Next Month", index=False)

                # Tulis data yang tidak bisa di-follow-up dari master baru (menggunakan df_tidak_fu asli)
                if not df_tidak_fu.empty:
                    # TELE_BARU sudah diset None saat pembentukan df_tidak_fu di atas jika belum ada
                    df_tidak_fu.to_excel(writer, sheet_name="Tidak Bisa FU", index=False)

            elif df_parts_lama: 
                df_lanjutan = pd.concat(df_parts_lama, ignore_index=True)
                df_lanjutan["RESULT"] = df_lanjutan["RESULT"].astype(str).str.strip().str.title()
                df_lanjutan["FollowUp(Hari)"] = df_lanjutan["RESULT"].map(mapping_fu)
                df_lanjutan["TGL"] = pd.to_datetime(df_lanjutan.get("TGL", today_date), errors="coerce").dt.date
                df_lanjutan["Tanggal FollowUp"] = df_lanjutan.apply(hitung_tgl_fu, axis=1)

                df_fu = df_lanjutan[~df_lanjutan["FollowUp(Hari)"].isnull()].copy()
                df_tidak = df_lanjutan[df_lanjutan["FollowUp(Hari)"].isnull()].copy()

                # Panggil fungsi penugasan TELE_BARU yang baru (ini akan mencoba mencocokkan TELE_LAMA)
                df_fu = assign_tele_baru(df_fu, nama_tele_baru)
                
                df_fu["TELE_LAMA"] = df_fu["TELE_LAMA"].astype(str)

                # Mengurutkan df_fu berdasarkan TELE_BARU setelah penugasan
                if "TELE_BARU" in df_fu.columns:
                    df_fu = df_fu.sort_values(by="TELE_BARU").reset_index(drop=True)

                # Urutan kolom
                cols = [c for c in df_fu.columns if c not in ["TELE_LAMA", "TELE_BARU"]] + ["TELE_LAMA", "TELE_BARU"]
                df_fu = df_fu[cols]

                df_fu.to_excel(writer, sheet_name="FU Lanjutan", index=False)
                df_fu[df_fu["FollowUp(Hari)"] == 1].to_excel(writer, sheet_name="FU Besok", index=False)
                df_fu[df_fu["FollowUp(Hari)"] == 2].to_excel(writer, sheet_name="FU Lusa", index=False)
                df_fu[df_fu["FollowUp(Hari)"] == 3].to_excel(writer, sheet_name="FU 3 Hari", index=False)
                df_fu[df_fu["FollowUp(Hari)"] == "Next Month"].to_excel(writer, sheet_name="FU Next Month", index=False)

                for tele in nama_tele_baru: # nama_tele_baru sudah diurutkan
                    df_chunk = df_fu[df_fu["TELE_BARU"] == tele].copy()
                    if not df_chunk.empty:
                        df_chunk.to_excel(writer, sheet_name=tele, index=False)

                if not df_tidak.empty:
                    df_tidak["TELE_BARU"] = None
                    df_tidak.to_excel(writer, sheet_name="Tidak Bisa FU", index=False)
            else:
                st.info("ℹ️ Silakan upload file Excel untuk diproses.")

        st.success("✅ Semua file berhasil diproses!")
        st.download_button("📥 Download Excel FU", data=output.getvalue(), file_name="FU_Output_Final.xlsx")

