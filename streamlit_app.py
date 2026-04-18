import streamlit as st
import pandas as pd
import numpy as np
from scipy.stats import norm
import io

# 1. KONFIGURASI HALAMAN (Mobile & Web Friendly)
st.set_page_config(page_title="Audit Sampling BPK", layout="centered")

# CSS Kustom untuk menyesuaikan elemen di layar HP
st.markdown("""
    <style>
    .main-header {
        background-color:#004a99; 
        padding:15px; 
        border-radius:10px; 
        margin-bottom:20px;
        color: white;
        text-align: center;
    }
    /* Menghilangkan padding berlebih di mobile */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    </style>
    <div class="main-header">
        <h2 style='margin:0;'>🏛️ Sistem Sampling Digital</h2>
        <p style='margin:0; font-size: 0.9em;'>Pemeriksaan Kepatuhan - BPK RI</p>
    </div>
    """, unsafe_allow_html=True)

# 2. FUNGSI HELPER
def format_idr(n):
    return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if isinstance(n, (int, float)) else n

def clean_val(t):
    try:
        t_str = str(t).replace(".", "").replace(",", ".")
        return float(t_str)
    except:
        return 0.0

# --- PANEL PARAMETER ---
st.subheader("⚙️ Parameter Audit")
# Menggunakan kolom yang akan menumpuk (stack) otomatis di Mobile
col_1, col_2 = st.columns([1, 1])

with col_1:
    nama_akun_audit = st.text_input("Nama Akun", "Belanja Barang dan Jasa")
    tm_raw = st.text_input("TM (Tolerable Misstatement)", "50.000.000,00")

with col_2:
    dr_pct = st.number_input("DR (Risiko Deteksi) %", value=7.0, step=0.5)
    n_st = st.slider("Jumlah Strata", 3, 10, 10)

st.divider()

# --- BAGIAN 1: PREPARASI ---
st.subheader("1️⃣ Persiapan Data")
with st.expander("Lihat Petunjuk & Download Template"):
    st.write("Pastikan file Excel memiliki kolom: **Kode, Nama OPD, Nama Akun, Keterangan, Nilai**.")
    
    template_buffer = io.BytesIO()
    df_template = pd.DataFrame({
        'Kode': ['1.02.01', '1.02.02'],
        'Nama OPD': ['Dinas Kesehatan', 'Dinas Pendidikan'],
        'Nama Akun': ['Belanja Barang', 'Belanja Jasa'],
        'Keterangan': ['A', 'B'],
        'Nilai': [75000000, 120000000]
    })
    with pd.ExcelWriter(template_buffer, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 Download Template Excel", 
        data=template_buffer.getvalue(), 
        file_name="Template_Audit.xlsx",
        use_container_width=True # Agar tombol memenuhi lebar layar di HP
    )

st.divider()

# --- BAGIAN 2: EKSEKUSI ---
st.subheader("2️⃣ Upload & Proses")
uploaded_file = st.file_uploader("Upload File Populasi (.xlsx)", type=["xlsx"])

if uploaded_file:
    df_pop = pd.read_excel(uploaded_file)
    
    if 'Nilai' in df_pop.columns:
        df_pop['Nilai'] = pd.to_numeric(df_pop['Nilai'], errors='coerce').fillna(0)
        st.success(f"✅ Data dimuat: {len(df_pop)} Baris")
        
        # Tombol besar yang nyaman ditekan di HP
        if st.button("🚀 JALANKAN SAMPLING", use_container_width=True):
            with st.spinner('Memproses statistik...'):
                tm = clean_val(tm_raw)
                z = round(abs(norm.ppf((dr_pct/100)/2)), 4)
                
                # Logic Target ACov
                if dr_pct <= 10: target_acov, cat = 0.51, "RENDAH"
                elif dr_pct <= 20: target_acov, cat = 0.31, "MENENGAH"
                else: target_acov, cat = 0.21, "TINGGI"

                # Stratifikasi
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

                # Hasil Akhir
                df_s = pd.concat([df_pop[df_pop['Strata'] == i].sample(n=int(row['n_h'])) for i, row in st_h.iterrows() if row['n_h'] > 0])
                
                # Output untuk Mobile: Menggunakan Metric
                st.markdown("### 📊 Hasil Analisis")
                m1, m2 = st.columns(2)
                m1.metric("Risiko DR", cat)
                m2.metric("Coverage", f"{acov*100:.1f}%")
                st.metric("Precision Achieved (A')", format_idr(prec))

                # Tabel diletakkan dalam expander agar tidak memakan ruang di HP
                with st.expander("Lihat Rincian Per Strata"):
                    st_h['Sum_S'] = df_s.groupby('Strata')['Nilai'].sum()
                    kkp = st_h.sort_index().reset_index()[['Strata', 'min', 'max', 'count', 'n_h', 'sum', 'Sum_S']]
                    kkp.columns = ['Strata', 'Min', 'Max', 'N Pop', 'n Sampel', 'Nilai Pop', 'Nilai Sampel']
                    st.dataframe(kkp.style.format({c: format_idr for c in ['Min', 'Max', 'Nilai Pop', 'Nilai Sampel']}))
                
                # Tombol Download Akhir
                res_buffer = io.BytesIO()
                with pd.ExcelWriter(res_buffer, engine='openpyxl') as writer:
                    df_s.to_excel(writer, index=False)
                st.download_button(
                    label="📥 DOWNLOAD HASIL SAMPEL (EXCEL)", 
                    data=res_buffer.getvalue(), 
                    file_name=f"Sampel_{nama_akun_audit}.xlsx",
                    use_container_width=True,
                    button_style='primary'
                )
    else:
        st.error("Kolom 'Nilai' tidak ditemukan!")
