import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Auto Recon Tools", page_icon="⚙️", layout="wide")

st.title("😹😹😹 Automated SO Reconciliation Tools by AmbatuMan 😹😹😹")
st.write("Rekonsil surplus, minus, dan unrecorded? Di sini aja.")

st.info("""
💡 **Aturan Penamaan Sheet Excel:**
Biar aplikasi bisa mendeteksi lokasi secara otomatis, pastikan *nama sheet* di file Excel-nya pake format:
* **DATA [Nama Lokasi]** (Contoh: `DATA K166`, `DATA S1`, `DATA K245`)
* **UNRECORD [Nama Lokasi]** (Contoh: `UNRECORD K166`, `UNRECORD S1`, `UNRECORD K245`)
""")

# File Uploader
uploaded_file = st.file_uploader("Upload File Raw Excel Hasil SO di sini (.xlsx)", type=["xlsx"])

def proses_rekonsiliasi(xl, data_sheet, unrecord_sheet):
    try:
        # 1. Baca DATA sheet
        df_data = pd.read_excel(xl, sheet_name=data_sheet)
        df_data.columns = [str(c).strip().upper() for c in df_data.columns]
        
        # Validasi kolom wajib
        required_cols = ['PN', 'RESULT', 'DIFF']
        if not all(col in df_data.columns for col in required_cols):
            st.error(f"Sheet '{data_sheet}' kekurangan kolom wajib (PN, Result, atau Diff).")
            return None
            
        if 'SN' not in df_data.columns:
            df_data['SN'] = ''
            
        # Standarisasi data
        df_data['PN'] = df_data['PN'].astype(str).str.strip().str.upper()
        df_data['SN'] = df_data['SN'].fillna('').astype(str).str.strip().str.upper()
        
        # Filter data temuan saja
        findings = df_data[df_data['RESULT'].isin(['NOT FOUND', 'MINUS', 'SURPLUS'])].copy()
        findings['DIFF'] = pd.to_numeric(findings['DIFF'], errors='coerce').fillna(0)
        
        # Hitung Qty minus & surplus
        findings['Qty_Minus_NotFound'] = findings.apply(lambda x: abs(x['DIFF']) if x['DIFF'] < 0 else 0, axis=1)
        findings['Qty_Surplus'] = findings.apply(lambda x: x['DIFF'] if x['DIFF'] > 0 else 0, axis=1)
        
        grouped_findings = findings.groupby(['PN', 'SN']).agg(
            Qty_Minus_NotFound=('Qty_Minus_NotFound', 'sum'),
            Qty_Surplus=('Qty_Surplus', 'sum')
        ).reset_index()
        
        # 2. Baca UNRECORD sheet jika ada
        if unrecord_sheet:
            df_unrec = pd.read_excel(xl, sheet_name=unrecord_sheet)
            df_unrec.columns = [str(c).strip().upper() for c in df_unrec.columns]
            
            if 'PN' in df_unrec.columns and 'QTY' in df_unrec.columns:
                if 'SN' not in df_unrec.columns:
                    df_unrec['SN'] = ''
                df_unrec['PN'] = df_unrec['PN'].astype(str).str.strip().str.upper()
                df_unrec['SN'] = df_unrec['SN'].fillna('').astype(str).str.strip().str.upper()
                df_unrec['QTY'] = pd.to_numeric(df_unrec['QTY'], errors='coerce').fillna(0)
                
                grouped_unrec = df_unrec.groupby(['PN', 'SN']).agg(
                    Qty_Unrecorded=('QTY', 'sum')
                ).reset_index()
                
                # Gabungkan data K166/Lokasi dengan Unrecorded
                recon = pd.merge(grouped_findings, grouped_unrec, on=['PN', 'SN'], how='outer').fillna(0)
            else:
                recon = grouped_findings.copy()
                recon['Qty_Unrecorded'] = 0
        else:
            recon = grouped_findings.copy()
            recon['Qty_Unrecorded'] = 0
            
        # 3. Logika Penentuan Match Status
        def cek_status_match(row):
            minus = row['Qty_Minus_NotFound']
            lebih = row['Qty_Surplus'] + row['Qty_Unrecorded']
            if minus > 0 and lebih > 0:
                return 'Match Ditemukan (Ada Minus & Lebih)'
            elif minus > 0 and lebih == 0:
                return 'Hanya Minus/Not Found'
            elif minus == 0 and lebih > 0:
                return 'Hanya Surplus/Unrecord'
            else:
                return 'Clear'
                
        recon['Status_Rekonsiliasi'] = recon.apply(cek_status_match, axis=1)
        
        # Urutkan dari yang ada "Match Ditemukan" biar enak dibaca paling atas
        recon = recon.sort_values(by='Status_Rekonsiliasi', ascending=False).reset_index(drop=True)
        
        # Rename kolom biar rapih saat di download
        recon.columns = ['Part Number (PN)', 'Serial Number (SN)', 'Total Qty Minus / Not Found', 'Total Qty Surplus', 'Total Qty Unrecorded', 'Status Rekonsiliasi']
        return recon
    except Exception as e:
        st.error(f"Gagal memproses sheet {data_sheet}: {str(e)}")
        return None

if uploaded_file is not None:
    # Membaca list dari seluruh sheet di Excel
    xl = pd.ExcelFile(uploaded_file)
    sheet_names = xl.sheet_names
    
    # Deteksi lokasi yang ada secara otomatis
    lokasi_list = []
    for s in sheet_names:
        match = re.match(r"^DATA\s+(.+)$", s, re.IGNORECASE)
        if match:
            loc_code = match.group(1).strip()
            # Cari pasangan sheet unrecord-nya (case-insensitive)
            unrec_sheet = None
            for u_s in sheet_names:
                if loc_code.lower() in u_s.lower() and "unrecord" in u_s.lower():
                    unrec_sheet = u_s
                    break
            lokasi_list.append({
                "lokasi": loc_code,
                "data_sheet": s,
                "unrecord_sheet": unrec_sheet
            })
            
    if len(lokasi_list) == 0:
        st.warning("⚠️ Tidak mendeteksi adanya sheet dengan format 'DATA [Nama Lokasi]'. Mohon cek kembali penamaan sheet Anda.")
    else:
        st.success(f"🎉 Berhasil mendeteksi {len(lokasi_list)} lokasi: {', '.join([l['lokasi'] for l in lokasi_list])}")
        
        # Proses seluruh data ke dalam dictionary dataframes
        hasil_proses = {}
        for item in lokasi_list:
            with st.spinner(f"Memproses lokasi {item['lokasi']}..."):
                df_hasil = proses_rekonsiliasi(xl, item['data_sheet'], item['unrecord_sheet'])
                if df_hasil is not None:
                    hasil_proses[item['lokasi']] = df_hasil
                    
        # Tampilkan Preview Data di Web Browser
        if hasil_proses:
            st.write("### 📊 Preview Hasil Rekonsiliasi")
            tabs = st.tabs(list(hasil_proses.keys()))
            for index, (loc, df) in enumerate(hasil_proses.items()):
                with tabs[index]:
                    # Hitung summary singkat untuk ditaruh di atas preview
                    total_match = len(df[df['Status Rekonsiliasi'] == 'Match Ditemukan (Ada Minus & Lebih)'])
                    st.metric(label=f"Match Potensial Ditemukan di {loc}", value=f"{total_match} Items")
                    st.dataframe(df, use_container_width=True)
            
            # Generator Excel untuk download massal (Semua lokasi jadi 1 file excel beda sheet)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                for loc, df in hasil_proses.items():
                    # Nama sheet excel maksimal 31 karakter
                    sheet_name = f"Hasil_{loc}"[:31]
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            st.write("---")
            st.download_button(
                label="📥 Download Hasil Rekonsiliasi (.xlsx)",
                data=buffer.getvalue(),
                file_name="Hasil_Rekonsilnya_lohya_😹😹😹.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
