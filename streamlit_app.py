#@title 🏛️ APLIKASI SAMPLING AUDIT BPK RI (KOLOM CUSTOM) { display-mode: "form" }
#@markdown ---
#@markdown **PETUNJUK PENGGUNAAN:**
#@markdown 1. Klik menu **Runtime > Run all** untuk mengaktifkan aplikasi.
#@markdown 2. Klik **"Download Template"** untuk mendapatkan format Excel yang sesuai (Kode, OPD, Akun, Keterangan, Nilai).
#@markdown 3. Isi data ke template, lalu klik **"Upload Data"** dan terakhir **"Daftar Sampel"**.
#@markdown ---

import pandas as pd
import numpy as np
from scipy.stats import norm
from google.colab import files
import io, warnings
import ipywidgets as widgets
from IPython.display import display, clear_output

warnings.filterwarnings('ignore')

# --- 1. LOGIKA FORMAT & CLEANING ---
def format_idr(n):
    return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if isinstance(n, (int, float)) else n

def clean_val(t):
    try: return float(str(t).replace(".", "").replace(",", "."))
    except: return 0.0

# --- 2. UI COMPONENTS ---
style = {'description_width': '180px'}
layout = widgets.Layout(width='420px')

nama_akun_input = widgets.Text(value="Belanja Barang dan Jasa", description='Nama Akun Audit:', style=style, layout=layout)
tm_input = widgets.Text(value="50.000.000,00", description='Tolerable Misstatement (A):', style=style, layout=layout)
dr_input = widgets.FloatText(value=7.0, description='Risiko Deteksi (DR) %:', style=style, layout=layout)
strata_input = widgets.Dropdown(options=range(3, 11), value=10, description='Jumlah Strata:', style=style, layout=layout)

btn_template = widgets.Button(description="Download Template", button_style='info', icon='download')
btn_upload = widgets.Button(description="Upload Data", button_style='primary', icon='upload')
btn_hitung = widgets.Button(description="Daftar Sampel", button_style='success', icon='check')
btn_reset = widgets.Button(description="Reset", button_style='danger', icon='trash')

out = widgets.Output()
memori = {'df': None}

# --- 3. FUNGSI ACTION ---

def on_template(b):
    with out:
        # Template sesuai kolom yang diminta: Kode, Nama OPD, Nama Akun, Keterangan, dan Nilai
        df_temp = pd.DataFrame({
            'Kode': ['1.02.01', '1.02.02', '1.02.03'], 
            'Nama OPD': ['Dinas Kesehatan', 'Dinas Pendidikan', 'Dinas PU'], 
            'Nama Akun': ['Belanja Barang', 'Belanja Jasa', 'Belanja Pemeliharaan'], 
            'Keterangan': ['Pembelian Obat', 'Honorarium', 'Semen & Pasir'],
            'Nilai': [75000000, 120000000, 45000000]
        })
        fname = "Template_Sampling_BPK.xlsx"
        df_temp.to_excel(fname, index=False)
        files.download(fname)
        print("✅ Template berhasil didownload. Pastikan kolom 'Nilai' terisi angka tanpa titik ribuan.")

def on_upload(b):
    with out:
        clear_output()
        up = files.upload()
        if up:
            # Membaca data dengan asumsi standar Excel
            df = pd.read_excel(io.BytesIO(up[list(up.keys())[0]]))
            if 'Nilai' in df.columns:
                df['Nilai'] = pd.to_numeric(df['Nilai'], errors='coerce').fillna(0)
                memori['df'] = df
                print(f"✅ Data '{nama_akun_input.value}' berhasil dimuat!")
            else: 
                print("❌ Error: Kolom 'Nilai' tidak ditemukan! Gunakan template agar header sesuai.")

def on_hitung(b):
    with out:
        clear_output()
        if memori['df'] is None: 
            print("❌ Harap upload data terlebih dahulu!"); return
        
        df, n_st, dr_pct = memori['df'].copy(), strata_input.value, dr_input.value
        tm = clean_val(tm_input.value)
        z = round(abs(norm.ppf((dr_pct/100)/2)), 4)
        
        # Penentuan Target ACov berdasarkan DR
        if dr_pct <= 10: target_acov, cat = 0.51, "RENDAH (High Assurance)"
        elif dr_pct <= 20: target_acov, cat = 0.31, "MENENGAH (Moderate)"
        else: target_acov, cat = 0.21, "TINGGI (Low Assurance)"

        # Stratifikasi
        bins = np.linspace(df['Nilai'].min(), df['Nilai'].max(), n_st + 1)
        df['Strata'] = pd.cut(df['Nilai'], bins=bins, labels=list(range(n_st, 0, -1)), include_lowest=True)
        df['Strata'] = df['Strata'].astype(int)
        
        st_h = df.groupby('Strata')['Nilai'].agg(['count', 'std', 'sum', 'min', 'max']).fillna(0)
        st_h['W'] = st_h['count'] * st_h['std']
        st_h['n_h'] = 0 
        
        # Iterasi Penentuan Jumlah Sampel Optimal
        n_iter, prec, acov, loops = max(int(len(df)*0.05), n_st * 2), 9e15, 0.0, 0
        while (prec > tm or acov < target_acov) and (n_iter <= len(df)):
            loops += 1
            total_w = st_h['W'].sum()
            if total_w > 0:
                st_h['n_h'] = (st_h['W'] / total_w * n_iter).round().fillna(0).astype(int)
            else:
                st_h['n_h'] = (st_h['count'] / len(df) * n_iter).round().fillna(0).astype(int)
            
            st_h['n_h'] = st_h.apply(lambda r: min(max(1, int(r['n_h'])), int(r['count'])) if r['count'] > 0 else 0, axis=1)
            st_h['Var'] = (st_h['std']**2 / st_h['n_h']) * (1 - st_h['n_h']/st_h['count'])
            prec = z * np.sqrt((st_h['count']**2 * st_h['Var'].fillna(0)).sum())
            acov = (st_h['n_h'] / st_h['count'] * st_h['sum']).sum() / df['Nilai'].sum()
            
            if (prec > tm or acov < target_acov): 
                n_iter += max(5, int(len(df)*0.02))
            else: break
            if loops > 200: break

        # Pengambilan Sampel Acak Berstrata
        df_s = pd.concat([df[df['Strata'] == i].sample(n=int(row['n_h'])) for i, row in st_h.iterrows() if row['n_h'] > 0])
        
        # --- Tampilan Hasil ---
        print(f"🏛️ AKUN: {nama_akun_input.value} | Risiko: {cat} | Iterasi: {loops}")
        
        summary_df = pd.DataFrame({
            "Indikator": ["Target Minimal ACov", "Realisasi ACov", "Precision Achieved (A')"], 
            "Hasil": [f"> {target_acov*100-1:.0f}%", f"{df_s['Nilai'].sum()/df['Nilai'].sum()*100:,.2f}%", format_idr(prec)]
        })
        display(summary_df)
        
        st_h['Sum_S'] = df_s.groupby('Strata')['Nilai'].sum()
        kkp = st_h.sort_index().reset_index()[['Strata', 'min', 'max', 'count', 'n_h', 'sum', 'Sum_S']]
        kkp.columns = ['Strata', 'Batas Bawah', 'Batas Atas', 'Jml Pop', 'n Sampel', 'Nilai Pop', 'Nilai Sampel']
        
        total = pd.DataFrame([['TOTAL', '-', '-', kkp['Jml Pop'].sum(), kkp['n Sampel'].sum(), kkp['Nilai Pop'].sum(), df_s['Nilai'].sum()]], columns=kkp.columns)
        display(pd.concat([kkp, total], ignore_index=True).style.format({c: format_idr for c in ['Batas Bawah', 'Batas Atas', 'Nilai Pop', 'Nilai Sampel']}))
        
        fname = f"Sampel_{nama_akun_input.value.replace(' ','_')}.xlsx"
        df_s.to_excel(fname, index=False)
        files.download(fname)

# --- 4. RUN APPS ---
btn_template.on_click(on_template); btn_upload.on_click(on_upload); btn_hitung.on_click(on_hitung)
btn_reset.on_click(lambda b: out.clear_output())

display(widgets.VBox([
    widgets.HTML("<div style='background-color:#004a99; padding:10px; border-radius:5px;'><h2 style='color:white; margin:0;'>🏛️ KKP DIGITAL BPK: Sampling Alat Bantu</h2></div>"),
    widgets.Label(" "),
    widgets.VBox([nama_akun_input, tm_input, dr_input, strata_input]), 
    widgets.Label(" "),
    widgets.HBox([btn_template, btn_upload, btn_hitung, btn_reset]), 
    out
]))
