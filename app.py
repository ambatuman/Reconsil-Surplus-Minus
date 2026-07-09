import streamlit as st
import pandas as pd
import io

# Setup Page & Header
st.set_page_config(page_title="Stok Tek Reconsil Tools", page_icon="⚙️", layout="wide")

# ==========================================
# STIKER "HALAH NYOCOT"
# ==========================================
# Pastikan file gambar stiker lu sudah di-upload ke Github di satu folder yang sama.
try:
    # Ganti string di bawah kalau ekstensi gambar lu beda (misal: "halah_nyocot.jpg")
    st.image("halah_nyocot.png", width=200) 
except:
    pass # Kalau gambarnya belum diupload/salah nama, lewatin aja biar web ga error

st.title("⚙️ Stok Tek Reconsil Tools")
st.write("Aplikasi online untuk mencocokkan temuan surplus, minus, dan unrecorded secara otomatis untuk multi-lokasi.")

# File Uploader
uploaded_file = st.file_uploader("Upload File Raw Excel Hasil Stock Take (.xlsx)", type=["xlsx"])

def proses_semua_sheet(file_path):
    xl = pd.ExcelFile(file_path)
    findings_list = []
    
    # Looping semua sheet di dalam excel untuk antisipasi
    for sheet in xl.sheet_names:
        df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        
        # 1. Cari Header Otomatis (Cari baris yang ada tulisan PN dan LOC)
        header_row = -1
        for i, row in df_raw.head(20).iterrows():
            row_str = [str(x).upper().strip() for x in row.values]
            if 'PN' in row_str and ('LOC' in row_str or 'LOCATION' in row_str):
                header_row = i
                break
                
        if header_row == -1:
            continue # Lewatin sheet kalau bukan berisi data stock
            
        # Potong dataframe sesuai posisi header
        df = df_raw.iloc[header_row+1:].copy()
        df.columns = [str(x).upper().strip() for x in df_raw.iloc[header_row].values]
        
        # Standarisasi nama kolom biar seragam
        df = df.rename(columns={
            'LOCATION': 'LOC', 
            'BIN ACTUAL/FOUND': 'BIN', 
            'BIN ACTUAL': 'BIN'
        })
        
        req_cols = ['LOC', 'PN']
        if not all(c in df.columns for c in req_cols): 
            continue
            
        # Bikin kolom kosong kalau datanya nggak ada
        if 'BIN' not in df.columns: df['BIN'] = ''
        if 'SN' not in df.columns: df['SN'] = ''
        if 'QTY ACTUAL' not in df.columns: df['QTY ACTUAL'] = 0
        
        # Bersihin Serial Number
        df['SN'] = df['SN'].fillna('').astype(str).str.upper().str.strip().replace('NAN', '')
        
        # 2. Tarik Data Utama (Minus, Surplus, Not Found, Unrecorded)
        if 'RESULT' in df.columns and 'DIFF' in df.columns:
            df['DIFF'] = pd.to_numeric(df['DIFF'], errors='coerce').fillna(0)
            df['QTY ACTUAL'] = pd.to_numeric(df['QTY ACTUAL'], errors='coerce').fillna(0)
            
            for _, row in df.iterrows():
                res = str(row['RESULT']).upper().strip()
                if res in ['MINUS', 'NOT FOUND', 'SURPLUS', 'UNRECORDED']:
                    qty = row['DIFF']
                    # Khusus Unrecorded biasanya pake Qty Actual kalau Diff-nya beda/nol
                    if res == 'UNRECORDED': 
                        qty = row['QTY ACTUAL'] if row['QTY ACTUAL'] != 0 else qty
                    
                    findings_list.append({
                        'LOC': str(row['LOC']).strip(),
                        'BIN': str(row['BIN']).strip(),
                        'PN': str(row['PN']).strip(),
                        'SN': str(row['SN']).strip(),
                        'RESULT': res,
                        'QTY': qty
                    })
                    
        # 3. Tarik Data dari Sheet Unrecorded (Antisipasi kalau filenya ada sheet khusus unrecorded)
        elif 'QTY' in df.columns and 'UNRECORD' in sheet.upper():
            df['QTY'] = pd.to_numeric(df['QTY'], errors='coerce').fillna(0)
            for _, row in df.iterrows():
                findings_list.append({
                    'LOC': str(row['LOC']).strip(),
                    'BIN': str(row['BIN']).strip(),
                    'PN': str(row['PN']).strip(),
                    'SN': str(row['SN']).strip(),
                    'RESULT': 'UNRECORDED',
                    'QTY': row['QTY']
                })

    # 4. Olah, Rapihkan, & Hitung Data
    df_f = pd.DataFrame(findings_list)
    if df_f.empty: 
        return df_f
    
    # Pisahin QTY sesuai jenis temuannya ke dalam kolom masing-masing
    df_f['QTY : Surplus'] = df_f.apply(lambda x: x['QTY'] if x['RESULT'] == 'SURPLUS' and x['QTY'] > 0 else 0, axis=1)
    df_f['QTY : Unrecorded'] = df_f.apply(lambda x: x['QTY'] if x['RESULT'] == 'UNRECORDED' and x['QTY'] > 0 else 0, axis=1)
    df_f['QTY : Minus'] = df_f.apply(lambda x: abs(x['QTY']) if x['RESULT'] == 'MINUS' else 0, axis=1)
    df_f['QTY : Not Found'] = df_f.apply(lambda x: abs(x['QTY']) if x['RESULT'] == 'NOT FOUND' else 0, axis=1)
    
    # Grouping data berdasarkan LOC, PN, dan SN.
    # Jika BIN-nya beda (misal Minus di Bin A, Unrecord di Bin B), teks BIN akan digabung pisah koma.
    grouped = df_f.groupby(['LOC', 'PN', 'SN']).agg({
        'BIN': lambda x: ', '.join(sorted(set(str(v) for v in x if str(v).strip() not in ['', 'nan']))),
        'QTY : Surplus': 'sum',
        'QTY : Minus': 'sum',
        'QTY : Not Found': 'sum',
        'QTY : Unrecorded': 'sum'
    }).reset_index()
    
    # Susun ulang posisi kolom sesuai request lu (LOC, BIN, PN, SN, QTY...)
    grouped = grouped[['LOC', 'BIN', 'PN', 'SN', 'QTY : Surplus', 'QTY : Minus', 'QTY : Not Found', 'QTY : Unrecorded']]
    
    # Hitung Total QTY (Net Discrepancy)
    grouped['Total QTY'] = (grouped['QTY : Surplus'] + grouped['QTY : Unrecorded']) - (grouped['QTY : Minus'] + grouped['QTY : Not Found'])
    
    # Logic Status Rekonsiliasi
    def cek_status(row):
        minus = row['QTY : Minus'] + row['QTY : Not Found']
        lebih = row['QTY : Surplus'] + row['QTY : Unrecorded']
        
        if minus > 0 and lebih > 0: return '🟢 Match Ditemukan'
        elif minus > 0 and lebih == 0: return '🔴 Minus/Not Found'
        elif minus == 0 and lebih > 0: return '🟡 Surplus/Unrecord'
        else: return '⚪ Clear'
        
    grouped['Status'] = grouped.apply(cek_status, axis=1)
    
    # Sortir agar yang warnanya Hijau (Match) auto naik ke paling atas
    grouped = grouped.sort_values(by='Status').reset_index(drop=True)
    return grouped

if uploaded_file is not None:
    with st.spinner("Lagi ngegas baca dan nyocokin data..."):
        try:
            df_hasil = proses_semua_sheet(uploaded_file)
            
            if df_hasil.empty:
                st.warning("⚠️ Data temuan (Minus/Surplus/NotFound) nggak ketemu di file Excel lu.")
            else:
                st.success("🎉 Rekonsiliasi Selesai!")
                
                # Fitur tambahan: Filter drop-down biar lu bisa milih mau liat LOC mana aja
                lokasi_unik = df_hasil['LOC'].unique().tolist()
                lokasi_terpilih = st.multiselect("📍 Filter Tabel Berdasarkan Lokasi:", options=lokasi_unik, default=lokasi_unik)
                
                df_tampil = df_hasil[df_hasil['LOC'].isin(lokasi_terpilih)]
                
                st.metric(label="Total Item Terekonsiliasi (Berdasarkan Filter)", value=f"{len(df_tampil)} Baris")
                st.dataframe(df_tampil, use_container_width=True)
                
                # Tombol Download
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_tampil.to_excel(writer, sheet_name="Hasil Rekonsil", index=False)
                
                st.write("---")
                st.download_button(
                    label="📥 Download Hasil Rekonsiliasi (.xlsx)",
                    data=buffer.getvalue(),
                    file_name="Master_Hasil_Rekonsiliasi.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Waduh ada error nih bos: {str(e)}")
