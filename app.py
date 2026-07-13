import streamlit as st
import pandas as pd
import io

# Setup Page & Header
st.set_page_config(page_title="Stok Tek Reconsil Tools", page_icon="⚙️", layout="wide")

col1, col2, col3 = st.columns ([1, 2, 1])
with col1:
    st.image("halah_nyocot.jpg", width=80)
with col2:
    st.title("⚙️ Stok Tek Reconsil Tools")
with col3:
    st.image("halah_nyocot.jpg", width=80)
    
st.write("Aplikasi pintar untuk mencari pasangan barang yang selisih (otomatis menyembunyikan temuan yang tidak memiliki pasangan silang).")

uploaded_file = st.file_uploader("Upload File Raw Excel Hasil Stock Take (.xlsx)", type=["xlsx"])

def proses_semua_sheet(file_path):
    xl = pd.ExcelFile(file_path)
    findings_list = []
    
    for sheet in xl.sheet_names:
        df_raw = pd.read_excel(xl, sheet_name=sheet, header=None)
        
        # 1. Cari Header Otomatis
        header_row = -1
        for i, row in df_raw.head(20).iterrows():
            row_str = [str(x).upper().strip() for x in row.values]
            if 'PN' in row_str and ('LOC' in row_str or 'LOCATION' in row_str):
                header_row = i
                break
                
        if header_row == -1:
            continue 
            
        df = df_raw.iloc[header_row+1:].copy()
        df.columns = [str(x).upper().strip() for x in df_raw.iloc[header_row].values]
        
        df = df.rename(columns={
            'LOCATION': 'LOC', 
            'BIN ACTUAL/FOUND': 'BIN', 
            'BIN ACTUAL': 'BIN'
        })
        
        req_cols = ['LOC', 'PN']
        if not all(c in df.columns for c in req_cols): 
            continue
            
        if 'BIN' not in df.columns: df['BIN'] = ''
        if 'SN' not in df.columns: df['SN'] = ''
        if 'QTY ACTUAL' not in df.columns: df['QTY ACTUAL'] = 0
        
        df['SN'] = df['SN'].fillna('').astype(str).str.upper().str.strip().replace('NAN', '')
        
        # 2. Tarik Data Utama 
        if 'RESULT' in df.columns and 'DIFF' in df.columns:
            df['DIFF'] = pd.to_numeric(df['DIFF'], errors='coerce').fillna(0)
            df['QTY ACTUAL'] = pd.to_numeric(df['QTY ACTUAL'], errors='coerce').fillna(0)
            
            for _, row in df.iterrows():
                res = str(row['RESULT']).upper().strip()
                if res in ['MINUS', 'NOT FOUND', 'SURPLUS', 'UNRECORDED']:
                    qty = row['DIFF']
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
                    
        # 3. Tarik Data Unrecorded
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

    df_f = pd.DataFrame(findings_list)
    if df_f.empty: 
        return df_f
    
    # 4. Olah & Pecah Kolom
    df_f['Surplus'] = df_f.apply(lambda x: x['QTY'] if x['RESULT'] == 'SURPLUS' and x['QTY'] > 0 else 0, axis=1)
    df_f['Unrecorded'] = df_f.apply(lambda x: x['QTY'] if x['RESULT'] == 'UNRECORDED' and x['QTY'] > 0 else 0, axis=1)
    df_f['Minus'] = df_f.apply(lambda x: abs(x['QTY']) if x['RESULT'] == 'MINUS' else 0, axis=1)
    df_f['Not Found'] = df_f.apply(lambda x: abs(x['QTY']) if x['RESULT'] == 'NOT FOUND' else 0, axis=1)
    
    df_f['BIN Surplus'] = df_f.apply(lambda x: x['BIN'] if x['RESULT'] == 'SURPLUS' else '', axis=1)
    df_f['BIN Unrecorded'] = df_f.apply(lambda x: x['BIN'] if x['RESULT'] == 'UNRECORDED' else '', axis=1)
    df_f['BIN Minus'] = df_f.apply(lambda x: x['BIN'] if x['RESULT'] == 'MINUS' else '', axis=1)
    df_f['BIN Not Found'] = df_f.apply(lambda x: x['BIN'] if x['RESULT'] == 'NOT FOUND' else '', axis=1)
    
    def gabung_bin(series):
        vals = sorted(set(str(v).strip() for v in series if str(v).strip() not in ['', 'nan', 'None']))
        return ', '.join(vals) if vals else '-'

    # 5. Grouping berdasarkan LOC, PN, dan SN
    grouped = df_f.groupby(['LOC', 'PN', 'SN']).agg({
        'Surplus': 'sum',
        'Minus': 'sum',
        'Not Found': 'sum',
        'Unrecorded': 'sum',
        'BIN Surplus': gabung_bin,
        'BIN Minus': gabung_bin,
        'BIN Not Found': gabung_bin,
        'BIN Unrecorded': gabung_bin
    }).reset_index()
    
    # 6. FILTERING: Wajib ada Minus/Not Found DAN Surplus/Unrecorded (TIDAK HARUS IMPAS)
    kondisi_kurang = (grouped['Minus'] > 0) | (grouped['Not Found'] > 0)
    kondisi_lebih = (grouped['Surplus'] > 0) | (grouped['Unrecorded'] > 0)
    
    # Tendang yang jomblo
    grouped = grouped[kondisi_kurang & kondisi_lebih].copy()
    
    if grouped.empty:
        return grouped
    
    # Hitung Total QTY
    grouped['Total QTY'] = (grouped['Surplus'] + grouped['Unrecorded']) - (grouped['Minus'] + grouped['Not Found'])
    
    # Kasih status detail biar lu gampang bacanya
    def detail_status(qty):
        if qty == 0:
            return '🟢 Match (Impas)'
        else:
            return '🟡 Match (Ada Selisih)'
            
    grouped['Status'] = grouped['Total QTY'].apply(detail_status)
    
    # 7. Susun Ulang Kolom
    kolom_final = [
        'LOC', 'PN', 'SN', 
        'BIN Minus', 'Minus', 
        'BIN Not Found', 'Not Found', 
        'BIN Surplus', 'Surplus', 
        'BIN Unrecorded', 'Unrecorded', 
        'Total QTY', 'Status'
    ]
    
    # Urutin yang impas ditaruh di atas
    grouped = grouped.sort_values(by=['Status', 'LOC', 'PN']).reset_index(drop=True)
    grouped = grouped[kolom_final]
    return grouped

if uploaded_file is not None:
    with st.spinner("Lagi ngegas baca data dan nyari pasangan..."):
        try:
            df_hasil = proses_semua_sheet(uploaded_file)
            
            if df_hasil.empty:
                st.warning("⚠️ Yah, ga ada satupun temuan yang bisa dipasangkan di file lu (semuanya jomblo).")
            else:
                st.success("🎉 Cihuy! Berhasil nemuin pasangan rekonsiliasi.")
                
                lokasi_unik = df_hasil['LOC'].unique().tolist()
                lokasi_terpilih = st.multiselect("📍 Filter Tabel Berdasarkan Lokasi:", options=lokasi_unik, default=lokasi_unik)
                
                df_tampil = df_hasil[df_hasil['LOC'].isin(lokasi_terpilih)]
                
                st.metric(label="Total Item Terekonsiliasi (Berdasarkan Filter)", value=f"{len(df_tampil)} Baris")
                st.dataframe(df_tampil, use_container_width=True)
                
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
