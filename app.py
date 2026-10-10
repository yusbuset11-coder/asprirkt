from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import requests
import streamlit as st

# Konfigurasi halaman web
st.set_page_config(
    page_title="ASPRI RKT - Automasi Prioritas Rencana Kerja Tahunan",
    layout="wide",
)

# 🌐 URL WEB APP GOOGLE APPS SCRIPT
GOOGLE_SHEETS_WEBHOOK_URL = "https://script.google.com/macros/s/AKfycbxhjUEmpBmFXhqvUVXDY0s4uve2Wp2J8mWz-x3rqAS2uUUlADl9MI_0WiUsaJcWEPL2/exec"

# 🎨 Custom CSS Profesional (Adaptif Sempurna untuk Light & Dark Mode)
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

    /* Terapkan font ke seluruh aplikasi */
    .stApp {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
    }

    /* Sembunyikan footer bawaan Streamlit */
    footer {visibility: hidden;}

    /* Hero Banner Utama Sejajar Sempurna */
    .hero-banner {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%) !important;
        padding: 30px 40px;
        border-radius: 14px;
        color: #ffffff !important;
        box-shadow: 0 8px 20px -4px rgba(37, 99, 235, 0.25);
        margin-bottom: 22px;
        display: flex;
        align-items: center;
        gap: 25px;
    }
    .hero-banner * {
        color: #ffffff !important;
    }
    .hero-logo-container {
        background: #ffffff !important;
        padding: 6px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        flex-shrink: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        width: 95px;
        height: 95px;
    }
    .hero-logo-container img {
        width: 75px !important;
        height: 75px !important;
        object-fit: contain !important;
    }
    .hero-content {
        flex-grow: 1;
    }
    .hero-title {
        font-size: 2.4rem !important;
        font-weight: 800 !important;
        color: #ffffff !important;
        margin: 0 0 4px 0 !important;
        letter-spacing: -0.5px;
        line-height: 1.15;
    }
    .hero-subtitle {
        font-size: 1.2rem !important;
        font-weight: 700 !important;
        color: #93c5fd !important;
        margin: 0 0 4px 0 !important;
        line-height: 1.3;
    }
    .hero-desc {
        font-size: 0.98rem !important;
        font-weight: 400 !important;
        color: #e2e8f0 !important;
        margin: 0 !important;
    }

    /* Styling Kotak File Uploader dengan Background Berwarna */
    [data-testid="stFileUploader"] {
        border-radius: 0 0 10px 10px !important;
        padding: 14px !important;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02) !important;
        margin-bottom: 18px !important;
    }
    [data-testid="stFileUploader"] section, 
    [data-testid="stFileUploader"] div[data-baseweb="file-uploader"] {
        border: 2px dashed #3b82f6 !important;
        border-radius: 8px !important;
        padding: 12px !important;
    }
    
    /* Membungkus teks ukuran file dengan badge kontras tinggi */
    [data-testid="stFileUploader"] small {
        background-color: #1e293b !important;
        color: #60a5fa !important;
        padding: 4px 10px !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        border: 1px solid #3b82f6 !important;
        display: inline-block !important;
        margin-top: 6px !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
    }

    /* Styling Sidebar Profesional Bergaya Dark Minimalist */
    [data-testid="stSidebar"] {
        background-color: #0f172a !important;
        border-right: 1px solid #1e293b !important;
    }
    [data-testid="stSidebar"] p, [data-testid="stSidebar"] span, [data-testid="stSidebar"] label, [data-testid="stSidebar"] h3 {
        color: #f8fafc !important;
    }
    [data-testid="stSidebar"] .stSelectbox label {
        font-size: 14px !important;
        font-weight: 600 !important;
        color: #cbd5e1 !important;
    }

    /* Perbaikan Tombol Sidebar (Ganti Sekolah / Keluar) Agar Selalu Terlihat Jelas */
    [data-testid="stSidebar"] button {
        background-color: #1e293b !important;
        color: #ffffff !important;
        border: 1px solid #475569 !important;
        font-weight: 700 !important;
    }
    [data-testid="stSidebar"] button:hover {
        background-color: #334155 !important;
        color: #60a5fa !important;
        border-color: #60a5fa !important;
    }

    /* Tombol Utama (Primary Button di Dalam Aplikasi) */
    .stButton button[kind="primary"] {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        border-radius: 8px !important;
        padding: 0.6rem 1.2rem !important;
        border: none !important;
        box-shadow: 0 4px 6px rgba(37, 99, 235, 0.3);
        transition: all 0.2s ease;
    }
    .stButton button[kind="primary"]:hover {
        background: linear-gradient(135deg, #1d4ed8 100%, #1e40af 100%) !important;
        box-shadow: 0 6px 8px rgba(37, 99, 235, 0.4);
    }

    /* KONTROL KELAS & LABEL FORM LOGIN */
    [data-testid="stForm"] {
        background-color: #ffffff !important;
        padding: 16px 25px !important;
        border-radius: 0 0 16px 16px !important;
        border: 2px solid #cbd5e1 !important;
        border-top: none !important;
        box-shadow: 0 15px 35px -10px rgba(30,58,138,0.15) !important;
        max-width: 560px !important;
        margin: 0 auto !important;
    }
    [data-testid="stForm"] label p, div[data-testid="stTextInput"] label p {
        color: #0f172a !important;
        font-weight: 800 !important;
        font-size: 1.02rem !important;
    }
    [data-testid="stTextInput"] input {
        background-color: #f8fafc !important;
        color: #0f172a !important;
        border: 1px solid #cbd5e1 !important;
        font-weight: 600 !important;
    }
    [data-testid="stTextInput"] input::placeholder {
        color: #64748b !important;
    }

    /* TOMBOL FORM SUBMIT LOGIN */
    [data-testid="stFormSubmitButton"] button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        font-weight: 800 !important;
        font-size: 1.05rem !important;
        border-radius: 10px !important;
        padding: 0.65rem 1.5rem !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.35) !important;
        width: 100% !important;
        transition: all 0.2s ease;
    }
    [data-testid="stFormSubmitButton"] button:hover {
        background: linear-gradient(135deg, #1d4ed8 100%, #1e40af 100%) !important;
        color: #ffffff !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.45) !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# --- FUNGSI HELPER GAMBAR BASE64 ---
import base64


def get_image_base64(path):
  if os.path.exists(path):
    with open(path, "rb") as img_file:
      return base64.b64encode(img_file.read()).decode("utf-8")
  return None


# --- INISIALISASI SESSION STATE UNTUK LOGIN ---
if "logged_in" not in st.session_state:
  st.session_state.logged_in = False
if "npsn" not in st.session_state:
  st.session_state.npsn = ""
if "nama_sekolah" not in st.session_state:
  st.session_state.nama_sekolah = ""

# --- HALAMAN LOGIN SEKOLAH ---
if not st.session_state.logged_in:
  logo_b64 = get_image_base64("logo.png")
  if logo_b64:
    login_logo_html = (
        f'<div style="text-align: center; margin-bottom: 4px;"><img'
        f' src="data:image/png;base64,{logo_b64}" alt="Logo ASPRI RKT"'
        f' style="width: 55px; height: 55px; object-fit: contain;'
        f' border-radius: 10px; box-shadow: 0 4px 12px rgba(37,99,235,0.25);'
        f' padding: 4px; background: white; border: 1px solid'
        f' #cbd5e1;"></div>'
    )
  else:
    login_logo_html = (
        '<div style="text-align: center; font-size: 30px; margin-bottom:'
        ' 4px;">📊</div>'
    )

  # Ditarik lebih ke atas dengan margin-top negatif tipis (-18px)
  st.markdown(
      f"""
        <div style="max-width: 560px; margin: -18px auto 0 auto; background: #ffffff !important; padding: 18px 25px 10px 25px; border-radius: 16px 16px 0 0; box-shadow: 0 10px 25px -10px rgba(30,58,138,0.15); border: 2px solid #cbd5e1; border-bottom: none;">
            {login_logo_html}
            <div style="text-align: center;">
                <h1 style="color: #2563eb !important; font-weight: 800 !important; font-size: 1.85rem !important; margin-bottom: 2px !important; letter-spacing: -0.5px !important;">ASPRI RKT</h1>
                <p style="color: #1e3a8a !important; font-size: 1.22rem !important; font-weight: 800 !important; margin: 4px 0 6px 0 !important; letter-spacing: -0.2px !important;">Automasi Prioritas Rencana Kerja Tahunan</p>
                <p style="color: #0f172a !important; font-weight: 700 !important; font-size: 0.88rem !important; line-height: 1.3 !important; margin-top: 2px !important;">Masukkan identitas sekolah Anda untuk menggunakan Aplikasi ASPRI RKT secara gratis.</p>
            </div>
        </div>
        """,
      unsafe_allow_html=True,
  )

  with st.form("form_login_sekolah"):
    input_nama = st.text_input(
        "Nama Satuan Pendidikan",
        placeholder="Contoh: SMK Negeri 2 Bangkalan",
    )
    input_npsn_login = st.text_input(
        "NPSN (Nomor Pokok Sekolah Nasional)",
        placeholder="Contoh: 2053xxxx",
    )

    submitted = st.form_submit_button(
        "🚀 Masuk ke Aplikasi ASPRI RKT", use_container_width=True
    )

    if submitted:
      if input_nama.strip() and input_npsn_login.strip():
        st.session_state.logged_in = True
        st.session_state.nama_sekolah = input_nama.strip().upper()
        st.session_state.npsn = input_npsn_login.strip()

        waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # --- 1. KIRIM KE GOOGLE SHEETS SECARA ONLINE ---
        try:
          payload = {
              "timestamp": waktu_sekarang,
              "nama_sekolah": st.session_state.nama_sekolah,
              "npsn": st.session_state.npsn,
          }
          requests.post(GOOGLE_SHEETS_WEBHOOK_URL, json=payload, timeout=5)
        except Exception:
          pass  # Jika offline/gagal koneksi, aplikasi tetap berjalan lancar

        # --- 2. CADANGAN: CATAT KE CSV LOKAL ---
        log_file = "school_users_log.csv"
        log_data = pd.DataFrame([{
            "Timestamp": waktu_sekarang,
            "Nama Sekolah": st.session_state.nama_sekolah,
            "NPSN": st.session_state.npsn,
        }])

        if os.path.exists(log_file):
          log_data.to_csv(
              log_file, mode="a", header=False, index=False, encoding="utf-8"
          )
        else:
          log_data.to_csv(log_file, index=False, encoding="utf-8")

        st.rerun()
      else:
        st.error(
            "⚠️ Harap isi Nama Satuan Pendidikan dan NPSN dengan lengkap dan"
            " benar!"
        )

  # --- FOOTER / NAMA PENGEMBANG DI BAWAH HALAMAN LOGIN ---
  st.markdown(
      """
        <div style="max-width: 560px; margin: 8px auto 0 auto; text-align: center;">
            <p style="color: #94a3b8 !important; font-size: 0.85rem !important; font-weight: 600 !important; letter-spacing: 0.2px; margin: 0;">
                Developed by <span style="color: #60a5fa !important; font-weight: 700;">Yustinus Budi Setyanta</span> - PS Cabdin Bangkalan
            </p>
        </div>
        """,
      unsafe_allow_html=True,
  )

  st.stop()


# --- FUNGSI SMART GENERATOR IMPLEMENTASI OPERASIONAL & UNIK ---
def generate_penjelasan_operasional_cerdas(kegiatan, akar_masalah):
  k_clean = kegiatan.strip()
  if k_clean.endswith("."):
    k_clean = k_clean[:-1]

  k_lower = k_clean.lower()

  if any(
      kw in k_lower
      for kw in [
          "komunitas belajar",
          "kombel",
          "diskusi",
          "refleksi",
          "belajar mandiri",
          "kolaborasi",
      ]
  ):
    return (
        f"1. Membentuk tim pengelola Komunitas Belajar (Kombel) untuk merancang"
        f" program: {k_clean}.\n2. Menyusun jadwal pertemuan rutin mingguan"
        f" guna membedah materi rapor pendidikan dan berbagi praktik baik antar"
        f" guru.\n3. Melaksanakan sesi diskusi reflektif dan lokakarya internal"
        f" secara kolaboratif.\n4. Melakukan penilaian sejawat (peer review),"
        f" evaluasi berkala, dan pelaporan progres ke kepala sekolah."
    )
  elif any(
      kw in k_lower
      for kw in [
          "peningkatan kompetensi",
          "pelatihan",
          "bimbingan",
          "bimtek",
          "workshop",
          "narasumber",
          "aktivitas belajar guru",
      ]
  ):
    return (
        f"1. Menugaskan guru dan kepala sekolah untuk mengikuti program:"
        f" {k_clean}.\n2. Menyiapkan KAK (Kerangka Acuan Kerja), instrumen"
        f" pemantauan, serta jadwal diseminasi di sekolah.\n3. Menyelenggarakan"
        f" lokakarya pengimbasan materi kepada seluruh dewan guru secara"
        f" terstruktur.\n4. Menerapkan hasil peningkatan kompetensi ke dalam"
        f" modul ajar/RPP dan memonitor penerapannya di kelas."
    )
  elif any(
      kw in k_lower
      for kw in [
          "alat",
          "bahan",
          "buku",
          "media",
          "pengadaan",
          "sarana",
          "prasarana",
          "tik",
          "literasi",
          "geometri",
          "sastra",
          "numerasi",
      ]
  ):
    return (
        f"1. Melakukan inventarisasi dan analisis kebutuhan spesifik sarana/media"
        f" untuk mendukung: {k_clean}.\n2. Menyusun RAB dan melaksanakan"
        f" pengadaan barang/jasa sesuai regulasi ARKAS/BOS.\n3. Mendistribusikan"
        f" sarana pendukung kepada guru dan siswa serta memastikan kesiapan"
        f" operasionalnya.\n4. Memanfaatkan media secara optimal dalam proses"
        f" pembelajaran interaktif dan merawat inventaris secara berkala."
    )
  elif any(
      kw in k_lower
      for kw in [
          "kebijakan",
          "sop",
          "peraturan",
          "tim",
          "pencegahan",
          "perundungan",
          "inklusi",
          "toleransi",
          "kesetaraan",
          "keamanan",
      ]
  ):
    return (
        f"1. Membentuk Tim Satuan Tugas khusus untuk mengelola program:"
        f" {k_clean}.\n2. Menyusun draf SOP, pedoman, atau tata tertib sekolah"
        f" yang transparan dan akomodatif bagi seluruh warga sekolah.\n3."
        f" Menyelenggarakan sosialisasi kebijakan kepada dewan guru, peserta"
        f" didik, dan orang tua/wali murid.\n4. Melaksanakan pemantauan rutin,"
        f" layanan pengaduan, dan evaluasi berkala efektivitas kebijakan."
    )
  else:
    return (
        f"1. Membentuk tim pelaksana kegiatan untuk merancang program:"
        f" {k_clean}.\n2. Menyusun jadwal operasional, instrumen pendukung,"
        f" serta alokasi sumber daya yang diperlukan.\n3. Melaksanakan kegiatan"
        f" secara aktif, partisipatif, dan terstruktur di satuan"
        f" pendidikan.\n4. Melakukan monitoring pelaksanaan, evaluasi"
        f" pencapaian program, dan penyusunan laporan pertanggungjawaban."
    )


# --- SIDEBAR PENGATURAN BERSIH ---
with st.sidebar:
  st.markdown("### ⚙️ Pengaturan Sumber Data")
  st.markdown(
      "<p style='font-size: 13px; color: #94a3b8 !important;'>Pilih"
      " rekomendasi dalam Rapor Pendidikan.</p>",
      unsafe_allow_html=True,
  )

  rekomendasi_dict = {
      "Rekom Keseluruhan": "2.1 REKOM. KESELURUHAN",
      "Rekom Prioritas": "2.2 REKOM. PRIORITAS",
  }

  pilihan_label = st.selectbox(
      "Pilih Rekomendasi",
      list(rekomendasi_dict.keys()),
      index=None,
      placeholder="Pilih Rekomendasi...",
      label_visibility="collapsed",
  )

  sheet_pilihan = (
      rekomendasi_dict[pilihan_label] if pilihan_label else None
  )

  st.markdown("---")
  st.markdown("### 🏫 Info Sekolah Aktif")
  st.markdown(
      f"<p style='font-size: 13px; color: #cbd5e1 !important;'><b>Sekolah:</b><br>{st.session_state.nama_sekolah}</p>",
      unsafe_allow_html=True,
  )
  st.markdown(
      f"<p style='font-size: 13px; color: #cbd5e1 !important;'><b>NPSN:</b><br>{st.session_state.npsn}</p>",
      unsafe_allow_html=True,
  )

  if st.button("🔄 Ganti Sekolah / Keluar", use_container_width=True):
    st.session_state.logged_in = False
    st.rerun()

# --- HERO BANNER UTAMA SEJAJAR & PROFESIONAL ---
logo_b64 = get_image_base64("logo.png")
if logo_b64:
  logo_html = (
      f'<div class="hero-logo-container"><img'
      f' src="data:image/png;base64,{logo_b64}" alt="Logo ASPRI RKT"></div>'
  )
else:
  logo_html = (
      '<div class="hero-logo-container"><span'
      ' style="font-size: 38px;">📊</span></div>'
  )

st.markdown(
    f"""
    <div class="hero-banner">
        {logo_html}
        <div class="hero-content">
            <div class="hero-title">ASPRI RKT</div>
            <div class="hero-subtitle">Automasi Prioritas Rencana Kerja Tahunan</div>
            <div class="hero-desc">Aplikasi Penyusunan Rencana Kerja Tahunan (RKT) Berbasis Data Rapor Pendidikan - Pengembang: Yustinus Budi Setyanta - PS Cabdin Bangkalan</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# --- HEADER UPLOAD DOKUMEN ---
st.markdown(
    """
    <div style="background: #0f172a !important; color: #ffffff !important; padding: 10px 16px; border-radius: 10px 10px 0 0; font-weight: 800; font-size: 1.05rem; letter-spacing: -0.2px;">
        📂 Unggah Dokumen Rapor Pendidikan
    </div>
    """,
    unsafe_allow_html=True,
)
uploaded_file = st.file_uploader(
    "Pilih atau seret file Rapor Pendidikan (.xlsx)",
    type=["xlsx"],
    label_visibility="collapsed",
)

if uploaded_file is not None:
  if not sheet_pilihan:
    st.warning(
        "⚠️ Silakan pilih jenis rekomendasi terlebih dahulu melalui menu di"
        " Sidebar sebelah kiri sebelum memproses data."
    )
  else:
    st.markdown(
        f"""
        <div style="background-color: #1e293b !important; border: 1px solid #3b82f6; padding: 10px 16px; border-radius: 8px; margin-bottom: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.02);">
            <p style="color: #60a5fa !important; font-size: 14px !important; font-weight: 700 !important; margin: 0;">✨ File Rapor Pendidikan berhasil diunggah dengan rekomendasi <b>{pilihan_label}</b> dan siap diproses!</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    try:
      xls_raw = pd.ExcelFile(uploaded_file)

      # Membaca data dan membersihkan kolom/baris kosong
      df_rekom = pd.read_excel(uploaded_file, sheet_name=sheet_pilihan, header=5)
      if df_rekom.shape[1] > 1:
        df_rekom = df_rekom.iloc[:, 1:]

      df_rekom = df_rekom.dropna(how="all").reset_index(drop=True)

      if len(df_rekom) > 0:
        first_col = df_rekom.columns[0]
        if (
            "Unnamed" in str(first_col)
            or str(df_rekom.iat[0, 0]).lower() == "none"
            or pd.isna(df_rekom.iat[0, 0])
        ):
          df_rekom = df_rekom.iloc[1:].reset_index(drop=True)

      # Membuat salinan dataframe khusus untuk pratinjau dengan kolom 'No' bernomor 1, 2, 3...
      df_preview = df_rekom.copy()
      df_preview.insert(0, "No", range(1, len(df_preview) + 1))

      # --- KOTAK PRATINJAU DATA ---
      st.markdown(
          f"""
          <div style="background: #1e293b !important; border: 1px solid #475569; border-radius: 10px; margin-bottom: 16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.03);">
              <div style="background: #0f172a !important; color: #ffffff !important; padding: 10px 16px; border-radius: 9px 9px 0 0; font-weight: 800; font-size: 1.05rem; letter-spacing: -0.2px;">
                  📋 Pratinjau Data: {pilihan_label} ({sheet_pilihan})
              </div>
              <div style="padding: 12px 16px; color: #f8fafc !important;">
          """,
          unsafe_allow_html=True,
      )
      st.dataframe(df_preview, use_container_width=True, hide_index=True)
      st.markdown("</div></div>", unsafe_allow_html=True)

      st.markdown(
          "<div style='margin: 10px 0;'></div>", unsafe_allow_html=True
      )

      # --- KOTAK EKSEKUSI ---
      st.markdown(
          """
          <div style="background: #1e293b !important; border: 1px solid #475569; border-radius: 10px; margin-bottom: 16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.03);">
              <div style="background: #0f172a !important; color: #ffffff !important; padding: 10px 16px; border-radius: 9px 9px 0 0; font-weight: 800; font-size: 1.05rem; letter-spacing: -0.2px;">
                  ⚡ Eksekusi Proses Otomatisasi
              </div>
              <div style="padding: 12px 16px; color: #f8fafc !important;">
          """,
          unsafe_allow_html=True,
      )

      if st.button(
          "🚀 Proses dan Susun RKT Secara Otomatis",
          type="primary",
          use_container_width=True,
      ):
        with st.spinner("Sedang memproses dan menyusun RKT secara instan..."):
          rkt_rows = []
          no_counter = 1

          for idx, row in df_rekom.iterrows():
            identifikasi = row.iloc[0] if len(row) > 0 else None
            if (
                pd.notna(identifikasi)
                and str(identifikasi).strip() != ""
                and str(identifikasi).lower() != "none"
            ):
              level_1 = (
                  str(row.iloc[3])
                  if len(row) > 3 and pd.notna(row.iloc[3])
                  else ""
              )
              level_2 = (
                  str(row.iloc[4])
                  if len(row) > 4 and pd.notna(row.iloc[4])
                  else ""
              )
              akar_masalah = f"{level_1} - {level_2}" if level_2 else level_1

              kegiatan_benahi = (
                  str(row.iloc[5])
                  if len(row) > 5 and pd.notna(row.iloc[5])
                  else ""
              )
              kegiatan_arkas_val = (
                  str(row.iloc[8])
                  if len(row) > 8 and pd.notna(row.iloc[8])
                  else ""
              )

              kb_lower = kegiatan_benahi.lower()
              ka_lower = kegiatan_arkas_val.lower()

              # 1. Logika Prediksi Biaya
              if (
                  any(
                      kw in kb_lower
                      for kw in [
                          "komunitas belajar",
                          "diskusi",
                          "refleksi",
                          "belajar mandiri",
                          "koordinasi internal",
                      ]
                  )
                  and not any(
                      kw in kb_lower
                      for kw in [
                          "narasumber",
                          "pelatihan eksternal",
                          "alat",
                          "bahan",
                          "pengadaan",
                      ]
                  )
              ):
                butuh_biaya = "Tidak"
              elif "tidak" in ka_lower and "biaya" in ka_lower:
                butuh_biaya = "Tidak"
              else:
                butuh_biaya = "Ya"

              # 2. Penjelasan Implementasi Cerdas & Unik per Baris
              penjelasan_implementasi = generate_penjelasan_operasional_cerdas(
                  kegiatan_benahi, akar_masalah
              )

              rkt_rows.append({
                  "No": no_counter,
                  "Identifikasi (Indikator Prioritas)": str(identifikasi),
                  "Akar Masalah": akar_masalah,
                  "Kegiatan Benahi": kegiatan_benahi,
                  "Penjelasan Implementasi Kegiatan": penjelasan_implementasi,
                  "Apakah Kegiatan Membutuhkan Biaya?": butuh_biaya,
              })
              no_counter += 1

          df_rkt_hasil = pd.DataFrame(rkt_rows)

          st.markdown(
              f"""
              <div style="background-color: #1e293b !important; border: 1px solid #3b82f6; padding: 10px 16px; border-radius: 8px; margin: 15px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.02);">
                  <p style="color: #60a5fa !important; font-size: 14px !important; font-weight: 700 !important; margin: 0;">🎉 Berhasil menyusun {len(df_rkt_hasil)} baris usulan RKT dengan penjelasan operasional yang unik di setiap baris!</p>
              </div>
              """,
              unsafe_allow_html=True,
          )

          # --- KOTAK HASIL LEMBAR KERJA ---
          st.markdown(
              """
              <div style="background: #1e293b !important; border: 1px solid #475569; border-radius: 10px; margin-bottom: 16px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.03);">
                  <div style="background: #0f172a !important; color: #ffffff !important; padding: 10px 16px; border-radius: 9px 9px 0 0; font-weight: 800; font-size: 1.05rem; letter-spacing: -0.2px;">
                      📋 Hasil Lembar Kerja RKT Tersusun
                  </div>
                  <div style="padding: 12px 16px; color: #f8fafc !important;">
              """,
              unsafe_allow_html=True,
          )
          edited_rkt = st.data_editor(
              df_rkt_hasil,
              num_rows="dynamic",
              key="editor_rkt",
              use_container_width=True,
          )
          st.markdown("</div></div>", unsafe_allow_html=True)

          # Proses konversi ke Excel dengan openpyxl (mulai dari baris ke-6)
          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            edited_rkt.to_excel(
                writer, index=False, sheet_name="Lembar Kerja RKT", startrow=5
            )

          output.seek(0)
          wb = openpyxl.load_workbook(output)
          ws_rkt = wb["Lembar Kerja RKT"]

          ws_rkt["A1"] = "LEMBAR RENCANA KERJA TAHUNAN (RKT)"
          ws_rkt["A1"].font = Font(name="Calibri", size=14, bold=True)
          ws_rkt["A3"] = f"Nama Satuan Pendidikan : {st.session_state.nama_sekolah}"
          ws_rkt["A3"].font = Font(name="Calibri", size=11, italic=True)
          ws_rkt["A4"] = f"NPSN : {st.session_state.npsn}"
          ws_rkt["A4"].font = Font(name="Calibri", size=11, italic=True)

          # Pengaturan Lebar Kolom Otomatis yang Proporsional & Rapi
          col_widths_rkt = {
              "A": 6,
              "B": 28,
              "C": 28,
              "D": 32,
              "E": 65,
              "F": 22,
          }
          for col_let, w in col_widths_rkt.items():
            ws_rkt.column_dimensions[col_let].width = w

          header_row_rkt = 6
          ws_rkt.row_dimensions[header_row_rkt].height = 30

          header_fill = PatternFill(
              start_color="1F4E78", end_color="1F4E78", fill_type="solid"
          )
          header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
          thin_border = Border(
              left=Side(style="thin", color="808080"),
              right=Side(style="thin", color="808080"),
              top=Side(style="thin", color="808080"),
              bottom=Side(style="thin", color="808080"),
          )

          for row in ws_rkt.iter_rows(
              min_row=header_row_rkt,
              max_row=ws_rkt.max_row,
              min_col=1,
              max_col=ws_rkt.max_column,
          ):
            for cell in row:
              cell.border = thin_border
              if cell.row == header_row_rkt:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(
                    wrap_text=True, vertical="center", horizontal="center"
                )
              else:
                cell.alignment = Alignment(
                    wrap_text=True, vertical="top", horizontal="left"
                )
                cell.font = Font(name="Calibri", size=11)

          # --- TAMBAHKAN BLOK CATATAN & TANDA TANGAN DI BAGIAN BAWAH ---
          last_row = ws_rkt.max_row
          note_row = last_row + 2

          ws_rkt.merge_cells(
              start_row=note_row,
              start_column=1,
              end_row=note_row + 1,
              end_column=6,
          )
          note_cell = ws_rkt.cell(row=note_row, column=1)
          note_cell.value = (
              "Dokumen Rencana Kerja Tahunan (RKT) ini telah disusun berdasarkan"
              " tahapan Perencanaan Berbasis Data (PBD) dengan menggunakan data"
              " Rapor Pendidikan dan data dari sumber lain yang relevan. Proses"
              " penyusunan RKT ini telah dilakukan secara kolaboratif bersama"
              " komite sekolah, warga sekolah dan pihak lain yang relevan untuk"
              " memperbaiki kualitas layanan pendidikan."
          )
          note_cell.font = Font(name="Calibri", size=10, italic=True)
          note_cell.alignment = Alignment(
              wrap_text=True, vertical="top", horizontal="left"
          )

          sig_row = note_row + 4

          # Kiri (Mengetahui)
          ws_rkt.cell(row=sig_row, column=2, value="Mengetahui,").font = Font(
              name="Calibri", size=11, bold=True
          )
          ws_rkt.cell(
              row=sig_row + 1,
              column=2,
              value=(
                  "(dapat ditandatangani oleh pihak yang relevan sesuai"
                  " konteks satdik, misalnya pihak yayasan/komite dan lainnya)"
              ),
          ).font = Font(name="Calibri", size=9, italic=True)
          ws_rkt.cell(
              row=sig_row + 5,
              column=2,
              value="Nama Lengkap: ....................................................",
          ).font = Font(name="Calibri", size=11)

          # Kanan (Mengesahkan)
          ws_rkt.cell(
              row=sig_row,
              column=5,
              value="[Nama Daerah], [tanggal] [bulan] [tahun]",
          ).font = Font(name="Calibri", size=11)
          ws_rkt.cell(
              row=sig_row + 1, column=5, value="Mengesahkan"
          ).font = Font(name="Calibri", size=11, bold=True)
          ws_rkt.cell(
              row=sig_row + 2,
              column=5,
              value=f"Kepala {st.session_state.nama_sekolah}",
          ).font = Font(name="Calibri", size=11, bold=True)
          ws_rkt.cell(
              row=sig_row + 5,
              column=5,
              value="Nama Lengkap: ....................................................",
          ).font = Font(name="Calibri", size=11)
          ws_rkt.cell(
              row=sig_row + 6,
              column=5,
              value="NIP: ....................................................",
          ).font = Font(name="Calibri", size=11)

          final_output = io.BytesIO()
          wb.save(final_output)
          processed_data = final_output.getvalue()

          st.markdown(
              """
              <div style="background-color: #1e293b !important; border: 2px solid #3b82f6; padding: 18px; border-radius: 10px; margin-top: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); text-align: center;">
                  <p style="color: #60a5fa !important; font-size: 16px !important; font-weight: 700 !important; margin-bottom: 10px;">📥 Dokumen RKT Resmi Telah Siap untuk Diunduh</p>
              </div>
              """,
              unsafe_allow_html=True,
          )

          st.download_button(
              label="📥 Unduh File RKT Resmi (Format Excel Ber-styling)",
              data=processed_data,
              file_name="RKT_ASPRI_Resmi.xlsx",
              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
              type="primary",
              use_container_width=True,
          )

      st.markdown("</div></div>", unsafe_allow_html=True)

    except Exception as e:
      st.error(
          f"Terjadi kesalahan saat memproses file. Pastikan format file sesuai."
          f" Detail error: {e}"
      )