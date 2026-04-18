import streamlit as st
import pandas as pd
import numpy as np
from scipy.stats import norm
import io

# 1. KONFIGURASI HALAMAN
st.set_page_config(page_title="Audit Sampling BPK", layout="wide")

# Tampilan Header Biru BPK
st.markdown("""
    <div style='background-color:#004a99; padding:20px; border-radius:10px; margin-bottom:20px'>
        <h2 style='color:white; margin:0; text-align:center;'>🏛️ Sistem Sampling Audit Digital</h2>
        <p style='color:white; text-align:center; margin:5px 0 0 0;'>Pemeriksaan Kepatuhan - BPK RI</p>
    </div>
    """, unsafe_allow_html=True)

# 2. FUNGSI FORMATTING
def format_idr(n):
    return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if isinstance(n, (int, float)) else n

def clean_val(t):
    try:
        # Menangani format ribuan titik dan desimal koma
        t_str = str(t).replace(".", "").replace(",", ".")
        return float(t_str)
    except:
        return 0.0

# --- SIDEBAR PARAMETER ---
st.sidebar.header("⚙️ Parameter Audit")
nama_akun_audit = st.sidebar.text_input("Nama Akun Audit", "Belanja Barang dan Jasa")
tm_raw = st.sidebar.text_input("Tolerable Misstatement (TM)", "50.000.000,00")
dr_pct = st.sidebar.number_input("Risiko Deteksi (DR) %", value=7.0, step=0.5)
n_st = st.sidebar.slider("Jumlah Strata", 3, 10, 10)

# --- UTAMA: DOWNLOAD TEMPLATE ---
st.subheader("1️⃣ Persiapkan Data")
st.write("Gunakan template ini agar kolom (Kode, Nama OPD, Nama Akun, Keterangan, Nilai) sesuai.")

template_buffer = io.BytesIO()
df_template = pd.DataFrame({
    'Kode': ['1.02.01', '1.02.02'],
    'Nama OPD': ['Dinas Kesehatan', 'Dinas Pendidikan'],
    'Nama Akun': ['Belanja Barang', 'Belanja Jasa'],
    'Keterangan': ['Pembelian Obat', 'Honorarium'],
    'Nilai': [75000000, 120000000]
})
with pd.ExcelWriter(template_buffer, engine='openpyxl') as writer:
    df_template.to_excel(writer, index=False)

st.download_button(
    label="📥 Download Template Excel",
    data=template_buffer.getvalue(),
    file_name="Template_Populasi_Audit.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.markdown("---")

# --- UTAMA: UPLOAD & PROSES ---
st.subheader("2️⃣ Upload & Eksekusi")
uploaded_file = st.file_uploader("Upload File Excel Populasi (Pastikan ada kolom 'Nilai')", type=["xlsx"])

if uploaded_file:
    # Membaca file dengan engine openpyxl untuk .xlsx
    df_pop = pd.read_excel(uploaded_file)
    
    if 'Nilai' in df_pop.columns:
        df_pop['Nilai'] = pd.to_numeric(df_pop['Nilai'], errors='coerce').fillna(0)
        st.success(f"✅ Data '{nama_akun_audit}' berhasil dimuat!")
        
        if st.button("🚀 Jalankan Proses Sampling"):
            with st.spinner('Menghitung presisi dan coverage...'):
                tm = clean_val(tm_raw)
                z = round(abs(norm.ppf((dr_pct/100)/2)), 4)
                
                # Mapping Target ACov sesuai arahan Pak Syakur
                if dr_pct <= 10: target_acov, cat = 0.51, "RENDAH (High Assurance)"
                elif dr_pct <= 20: target_acov, cat = 0.31, "MENENGAH (Moderate)"
                else: target_acov, cat = 0.21, "TINGGI (Low Assurance)"

                # Stratifikasi (Strata 1 = Nilai Terbesar)
                bins = np.linspace(df_pop['Nilai'].min(), df_pop['Nilai'].max(), n_st + 1)
                df_pop['Strata'] = pd.cut(df_pop['Nilai'], bins=bins, labels=list(range(n_st, 0, -1)), include_lowest=True)
                df_pop['Strata'] = df_pop['Strata'].astype(int)
                
                st_h = df_pop.groupby('Strata')['Nilai'].agg(['count', 'std', 'sum', 'min', 'max']).fillna(0)
                st_h['W'] = st_h['count'] * st_h['std']
                st_h['n_h'] = 0
                
                # Iterasi Presisi & Coverage
                n_iter, prec, acov, loops = max(int(len(df_pop)*0.05), n_st * 2), 9e15, 0.0, 0
                while (prec > tm or acov < target_acov) and (n_iter <= len(df_pop)):
                    loops += 1
                    total_w = st_h['W'].sum()
                    if total_w > 0:
                        st_h['n_h'] = (st_h['W'] / total_w * n_iter).round().fillna(0).astype(int)
                    else:
                        st_h['n_h'] = (st_h['count'] / len(df_pop) * n_iter).round().fillna(0).astype(int)
                    
                    st_h['n_h'] = st_h.apply(lambda r: min(max(1, int(r['n_h'])), int(r['count'])), axis=1)
                    st_h['Var'] = (st_h['std']**2 / st_h['n_h']) * (1 - st_h['n_h']/st_h['count'])
                    prec = z * np.sqrt((st_h['count']**2 * st_h['Var'].fillna(0)).sum())
                    acov = (st_h['n_h'] / st_h['count'] * st_h['sum']).sum() / df_pop['Nilai'].sum()
                    if (prec > tm or acov < target_acov): n_iter += max(5, int(len(df_pop)*0.02))
                    else: break
                    if loops > 200: break

                # Sampel Akhir
                df_s = pd.concat([df_pop[df_pop['Strata'] == i].sample(n=int(row['n_h'])) for i, row in st_h.iterrows() if row['n_h'] > 0])
                
                # Dashboard Hasil
                st.markdown("### 📊 Ringkasan Statistik")
                c1, c2, c3 = st.columns(3)
                c1.metric("Kategori DR", cat)
                c2.metric("Realisasi Coverage", f"{acov*100:.2f}%")
                c3.metric("Precision A'", format_idr(prec))

                st_h['Sum_S'] = df_s.groupby('Strata')['Nilai'].sum()
                kkp = st_h.sort_index().reset_index()[['Strata', 'min', 'max', 'count', 'n_h', 'sum', 'Sum_S']]
                kkp.columns = ['Strata', 'Batas Bawah', 'Batas Atas', 'Jml Pop', 'n Sampel', 'Nilai Pop', 'Nilai Sampel']
                st.table(kkp.style.format({c: format_idr for c in ['Batas Bawah', 'Batas Atas', 'Nilai Pop', 'Nilai Sampel']}))
                
                # Download Hasil Excel
                res_buffer = io.BytesIO()
                with pd.ExcelWriter(res_buffer, engine='openpyxl
