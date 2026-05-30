import pandas as pd
import numpy as np
from collections import defaultdict
import re
import os

class LPDProcessor:
    def __init__(self):
        self.detected_cabang = None
        self.debug_mode = True
        
        # 1. DEFAULT SKU (25 kode hardcode)
        self.default_skus = [
            'TO', 'TB', 'TBL', 'TS', 'TOS4', 'TBS', 'TLS', 'TWS',
            'TOSB', 'TWSB', 
            'FK', 'FUK', 'FKS', 'PCB', 'PCC', 'PCK', 'PCP',
            'KNT', 'KNC', 'KNW', 'YPB', 'YPC', 'YPS',
            'PNN', 'PNT', 'ORI', 'DB', 'TA'
        ]
        
        # 2. MAPPING DEFAULT
        self.sku_mapping = {
            'TO': 'TO', 'TB': 'TB', 'TBL': 'TBL', 'TS': 'TS',
            'TOS4': 'TOS4', 'TBS': 'TBS', 'TLS': 'TLS', 'TWS': 'TWS',
            'TOSB': 'TOSB', 'TWSB': 'TWSB',
            'FK': 'FK', 'FUK': 'FUK', 'FKS': 'FKS', 
            'PCB': 'PCB', 'PCC': 'PCC', 'PCK': 'PCK', 'PCP': 'PCP',
            'KNT': 'KNT', 'KNC': 'KNC', 'KNW': 'KNW',
            'YPB': 'YPB', 'YPC': 'YPC', 'YPS': 'YPS',
            'PNN': 'PNN', 'PNT': 'PNT',
            'ORI': 'ORI', 'DB': 'DB', 'TA': 'TA',
        }
        
        # 3. DYNAMIC SKU (akan diisi dari data)
        self.dynamic_skus = []  # Awalnya kosong
        
        # 4. ALL SKU (gabungan default + dynamic)
        self.all_skus = self.default_skus.copy()  # Mulai dari default
        # =============================================
        
        # Template columns - akan di-update nanti
        self.template_columns = []
            
    def detect_file_format(self, df):
        """Deteksi format file - VERSI BARU YANG LEBIH AKURAT"""
        if df is None or df.empty:
            return "UNKNOWN"
        
        column_names = [str(col).upper().strip() for col in df.columns]
        
        print(f"\n[ANALISIS FORMAT FILE - DETAIL]")
        print(f"Jumlah kolom: {len(column_names)}")
        print(f"Kolom: {column_names}")
        
        # Analisis mendalam setiap kolom
        for i, col_name in enumerate(column_names):
            if col_name:  # Hanya tampilkan kolom yang punya nama
                print(f"  Kolom {i}: '{col_name}'")
        
        # Cari kata kunci spesifik
        has_kd_item = any('KD' in col or 'ITEM' in col for col in column_names)
        has_tanggal = any('TANGGAL' in col or 'TGL' in col or 'DATE' in col for col in column_names)
        has_total = any('TOTAL' in col for col in column_names)
        has_jml = any('JML' in col or 'JUMLAH' in col for col in column_names)
        has_pelanggan = any('PELANGGAN' in col or 'CUSTOMER' in col or 'NAMA' in col for col in column_names)
        
        print(f"\n[ANALISIS KATA KUNCI]")
        print(f"  - Punya 'KD/ITEM': {has_kd_item}")
        print(f"  - Punya 'TANGGAL': {has_tanggal}")
        print(f"  - Punya 'TOTAL': {has_total}")
        print(f"  - Punya 'JML/JUMLAH': {has_jml}")
        print(f"  - Punya 'PELANGGAN': {has_pelanggan}")
        
        # LOGIKA DETEKSI BARU:
        # 1. KONSINYASI: Ada kolom TOTAL (nilai rupiah) + struktur banyak kolom
        # 2. PENJUALAN: Tidak ada TOTAL, hanya JML + struktur sederhana
        
        if has_total:
            print(f"  → DETEKSI: FORMAT KONSINYASI (ada kolom TOTAL untuk nilai rupiah)")
            return "KONSINYASI"
        elif has_jml and has_kd_item and has_pelanggan:
            print(f"  → DETEKSI: FORMAT PENJUALAN (ada JML, KD.ITEM, PELANGGAN)")
            return "PENJUALAN"
        elif len(df.columns) >= 8:  # Banyak kolom biasanya konsinyasi
            print(f"  → DETEKSI: FORMAT KONSINYASI (struktur kompleks, {len(df.columns)} kolom)")
            return "KONSINYASI"
        elif len(df.columns) <= 5:  # Sedikit kolom biasanya penjualan
            print(f"  → DETEKSI: FORMAT PENJUALAN (struktur sederhana, {len(df.columns)} kolom)")
            return "PENJUALAN"
        else:
            print(f"  → DETEKSI: FORMAT TIDAK DIKENALI")
            return "UNKNOWN"
    
    # ==================== METODE EKSTRAKSI DATA ====================
    
    def extract_lpd_data(self, file_path):
        try:
            print(f"\n=== EKSTRAK DATA DARI FILE ===")
            print(f"File: {os.path.basename(file_path)}")
            
            # Baca file Excel
            try:
                xls = pd.ExcelFile(file_path)
                sheet_name = xls.sheet_names[0]
                print(f"Sheet: {sheet_name}")
                
                # Baca tanpa header dulu untuk analisis
                df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # Cari baris header
                header_row = None
                for i in range(min(15, len(df_raw))):  # Cek 15 baris pertama
                    row_values = df_raw.iloc[i].astype(str).str.upper().tolist()
                    # Cari baris yang mengandung kata kunci penting
                    if (any('KD' in val or 'ITEM' in val for val in row_values) and
                        any('NAMA' in val or 'PELANGGAN' in val for val in row_values)):
                        header_row = i
                        print(f"✓ Header ditemukan di baris {i+1}")
                        break
                
                if header_row is None:
                    header_row = 0
                    print("⚠ Header tidak ditemukan, gunakan baris 0")
                
                # Baca dengan header yang benar
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
                
            except Exception as e:
                print(f"Error baca Excel: {e}, coba CSV")
                df = pd.read_csv(file_path, encoding='utf-8')
            
            # Deteksi format file berdasarkan struktur
            file_type = self.detect_file_format(df)
            
            print(f"\n✓ Format file terdeteksi: {file_type}")
            
            # Tampilkan struktur untuk debugging
            print(f"\n[DEBUG STRUKTUR KOLOM]")
            for i, col in enumerate(df.columns):
                col_str = str(col)
                if col_str.strip() == '':
                    print(f"  {i}. (KOLOM KOSONG)")
                else:
                    print(f"  {i}. '{col_str}'")
            
            # Mapping kolom berdasarkan format yang terdeteksi
            tanggal_col = None
            kd_item_col = None
            nama_col = None
            jml_col = None
            total_col = None
            
            # CARI KOLOM BERDASARKAN NAMA (case-insensitive)
            for col in df.columns:
                col_str = str(col).upper().strip()
                
                if col_str == '':  # Skip kolom kosong
                    continue
                    
                # Tanggal
                if tanggal_col is None and any(word in col_str for word in ['TANGGAL', 'DATE', 'TGL', 'WAKTU']):
                    tanggal_col = col
                    print(f"✓ Kolom Tanggal: '{col}'")
                
                # Kode Item
                if kd_item_col is None and any(word in col_str for word in ['KD', 'KODE', 'ITEM', 'BARANG', 'PRODUK', 'SKU']):
                    kd_item_col = col
                    print(f"✓ Kolom Kd. Item: '{col}'")
                
                # Nama
                if nama_col is None and any(word in col_str for word in ['NAMA', 'PELANGGAN', 'CUSTOMER', 'CUST', 'TOKO', 'OUTLET']):
                    nama_col = col
                    print(f"✓ Kolom Nama: '{col}'")
                
                # Jumlah
                if jml_col is None and any(word in col_str for word in ['JML', 'JUMLAH', 'QTY', 'QUANTITY', 'KUANTITAS']):
                    jml_col = col
                    print(f"✓ Kolom Jumlah: '{col}'")
                
                # Total
                if total_col is None and any(word in col_str for word in ['TOTAL', 'HARGA', 'NILAI', 'AMOUNT', 'RUPIAH']):
                    total_col = col
                    print(f"✓ Kolom Total: '{col}'")
            
            # JIKA KOLOM TIDAK DITEMUKAN, GUNAKAN HEURISTIK BERDASARKAN FORMAT
            if file_type == "KONSINYASI":
                # HEURISTIK KONSINYASI: Kd.Item biasanya di posisi tertentu
                if kd_item_col is None and len(df.columns) > 1:
                    # Coba kolom B (index 1)
                    potential_cols = []
                    for idx, col in enumerate(df.columns):
                        col_str = str(col).upper()
                        if idx == 1 or ('KD' not in col_str and 'ITEM' not in col_str and 'TANGGAL' not in col_str):
                            potential_cols.append((idx, col))
                    
                    if potential_cols:
                        # Ambil kolom dengan data yang terlihat seperti kode item
                        for idx, col in potential_cols:
                            sample_values = df[col].dropna().astype(str).str.upper().head(10).tolist()
                            # Cek apakah isinya seperti kode item (TO, TB, TBL, dll)
                            if any(any(sku in val for sku in self.all_skus) for val in sample_values):
                                kd_item_col = col
                                print(f"✓ Kd.Item terdeteksi heuristik: '{col}' (isi: {sample_values[:3]}...)")
                                break
                        
                        if kd_item_col is None and potential_cols:
                            kd_item_col = potential_cols[0][1]  # Ambil kolom pertama yang potensial
                            print(f"⚠ Kd.Item asumsi: '{kd_item_col}'")
                
                # Nama biasanya setelah Kd.Item
                if nama_col is None and kd_item_col:
                    kd_idx = list(df.columns).index(kd_item_col)
                    if len(df.columns) > kd_idx + 1:
                        nama_col = df.columns[kd_idx + 1]
                        print(f"✓ Nama asumsi (setelah Kd.Item): '{nama_col}'")
                
                # Jumlah biasanya setelah Nama
                if jml_col is None and nama_col:
                    nama_idx = list(df.columns).index(nama_col)
                    if len(df.columns) > nama_idx + 1:
                        jml_col = df.columns[nama_idx + 1]
                        print(f"✓ Jumlah asumsi (setelah Nama): '{jml_col}'")
            
            elif file_type == "PENJUALAN":
                # HEURISTIK PENJUALAN: Tanggal di awal
                if tanggal_col is None and len(df.columns) > 0:
                    # Coba kolom pertama yang berisi tanggal format
                    for col in df.columns:
                        sample_val = str(df[col].iloc[0]) if len(df) > 0 else ""
                        # Cek apakah ini format tanggal
                        if any(sep in sample_val for sep in ['-', '/', ':']) and any(year in sample_val for year in ['2024', '2025', '2026']):
                            tanggal_col = col
                            print(f"✓ Tanggal terdeteksi dari format: '{col}'")
                            break
                    
                    if tanggal_col is None:
                        tanggal_col = df.columns[0]
                        print(f"⚠ Tanggal asumsi (kolom pertama): '{tanggal_col}'")
                
                # Kd.Item biasanya setelah Tanggal
                if kd_item_col is None and tanggal_col:
                    tanggal_idx = list(df.columns).index(tanggal_col)
                    if len(df.columns) > tanggal_idx + 1:
                        kd_item_col = df.columns[tanggal_idx + 1]
                        print(f"✓ Kd.Item asumsi (setelah Tanggal): '{kd_item_col}'")
            
            # FALLBACK: Jika masih belum ditemukan, gunakan logika posisi default
            if kd_item_col is None and len(df.columns) > 1:
                kd_item_col = df.columns[1]
                print(f"⚠ Kd.Item fallback (kolom B): '{kd_item_col}'")
            
            if nama_col is None and len(df.columns) > 2:
                nama_col = df.columns[2]
                print(f"⚠ Nama fallback (kolom C): '{nama_col}'")
            
            if jml_col is None and len(df.columns) > 3:
                jml_col = df.columns[3]
                print(f"⚠ Jumlah fallback (kolom D): '{jml_col}'")
            
            # Bersihkan data
            df_clean = df.copy()
            
            # Hapus baris kosong (semua kolom NaN)
            df_clean = df_clean.dropna(how='all')
            
            # Hapus baris yang tidak memiliki data penting
            if kd_item_col in df_clean.columns:
                df_clean = df_clean[df_clean[kd_item_col].notna()]
            if nama_col in df_clean.columns:
                df_clean = df_clean[df_clean[nama_col].notna()]
            
            # Konversi jumlah ke numeric - PERBAIKAN BESAR
            if jml_col in df_clean.columns:
                print(f"\n[KONVERSI JUMLAH]")
                print(f"✓ Kolom jumlah ditemukan: '{jml_col}'")
                
                # Tampilkan contoh data sebelum konversi
                sample_data = df_clean[jml_col].head(10).tolist()
                print(f"  Contoh 10 data pertama: {sample_data}")
                
                # Coba berbagai metode konversi
                original_dtype = str(df_clean[jml_col].dtype)
                print(f"  Tipe data asli: {original_dtype}")
                
                # Method 1: Coba konversi langsung ke numeric
                df_clean[jml_col] = pd.to_numeric(df_clean[jml_col], errors='coerce')
                
                # Cek hasil konversi
                nan_count = df_clean[jml_col].isna().sum()
                print(f"  Setelah konversi: {nan_count} nilai NaN")
                
                # Method 2: Jika masih banyak NaN, coba ekstrak angka dari string
                if nan_count > len(df_clean) * 0.5:  # Jika lebih dari 50% NaN
                    print(f"  ⚠ Banyak data gagal dikonversi, coba parsing string...")
                    
                    def extract_number(x):
                        if pd.isna(x):
                            return 0
                        x_str = str(x)
                        # Cari angka (termasuk desimal)
                        numbers = re.findall(r'[\d\.]+', x_str)
                        if numbers:
                            try:
                                return float(numbers[0])
                            except:
                                return 0
                        return 0
                    
                    df_clean[jml_col] = df_clean[jml_col].apply(extract_number)
                
                # Fill NaN dengan 0
                df_clean[jml_col] = df_clean[jml_col].fillna(0)
                
                # Konversi ke integer jika memungkinkan
                if df_clean[jml_col].apply(float.is_integer).all():
                    df_clean[jml_col] = df_clean[jml_col].astype(int)
                
                print(f"✓ Konversi selesai:")
                print(f"  - Total baris: {len(df_clean)}")
                print(f"  - Total jumlah barang: {df_clean[jml_col].sum():.0f}")
                print(f"  - Rata-rata per baris: {df_clean[jml_col].mean():.2f}")
                
            else:
                print(f"\n[ERROR: KOLOM JUMLAH]")
                print(f"✗ Kolom '{jml_col}' tidak ditemukan di DataFrame!")
                print(f"  Kolom yang tersedia: {list(df_clean.columns)}")
                
                # Cari kolom yang mungkin berisi jumlah
                numeric_cols = []
                for col in df_clean.columns:
                    try:
                        if pd.api.types.is_numeric_dtype(df_clean[col]):
                            numeric_cols.append(col)
                    except:
                        pass
                
                if numeric_cols:
                    print(f"  Kolom numerik yang ditemukan: {numeric_cols}")
                    # Ambil kolom numerik pertama yang bukan tanggal atau total
                    for col in numeric_cols:
                        if 'TOTAL' not in str(col).upper() and 'TANGGAL' not in str(col).upper():
                            jml_col = col
                            print(f"  ⚠ Gunakan '{col}' sebagai jumlah")
                            break
                
                if jml_col is None or jml_col not in df_clean.columns:
                    print(f"  ⚠ Buat kolom jumlah default = 1")
                    df_clean['JUMLAH_DEFAULT'] = 1
                    jml_col = 'JUMLAH_DEFAULT'
            
            print(f"\n✓ Data berhasil diekstrak:")
            print(f"  - Jumlah baris: {len(df_clean)}")
            print(f"  - Jumlah kolom: {len(df_clean.columns)}")
            
            # DETEKSI KODE ITEM DINAMIS DARI DATA
            print(f"\n🔍 [AUTO-DETECT KODE ITEM DARI DATA]")

            if kd_item_col and kd_item_col in df_clean.columns:
                # Ambil semua kode item unik dari data
                unique_items = df_clean[kd_item_col].dropna().astype(str).str.strip().str.upper()
                unique_items = unique_items[unique_items != '']
                
                # Bersihkan kode dan FILTER HANYA YANG VALID
                cleaned_items = []
                for item in unique_items:
                    clean = re.sub(r'\s+', '', item)
                    clean = re.sub(r'[^A-Z0-9]', '', clean)
                    
                    # FILTER: Hanya tambahkan jika valid (bukan nama kolom)
                    if clean and clean not in cleaned_items:
                        # Skip jika ini nama kolom
                        skip_columns = [
                            'KDITEM', 'KODEITEM', 'ITEM', 'KODE', 'KD',
                            'NAMA', 'PELANGGAN', 'CUSTOMER',
                            'JML', 'JUMLAH', 'QTY',
                            'TOTAL', 'HARGA', 'NILAI',
                            'TANGGAL', 'DATE', 'TGL'
                        ]
                        
                        if clean in skip_columns:
                            print(f"  ⚠ Skip '{item}' (nama kolom)")
                            continue
                        
                        # Skip jika format tidak valid untuk SKU
                        if (len(clean) < 2 or len(clean) > 6 or 
                            clean.isdigit() or 
                            not any(c.isalpha() for c in clean)):
                            print(f"  ⚠ Skip '{item}' (format tidak valid)")
                            continue
                        
                        # Skip jika mengandung "ITEM" dan pendek
                        if 'ITEM' in clean and len(clean) <= 6:
                            print(f"  ⚠ Skip '{item}' (nama kolom 'ITEM')")
                            continue
                            
                        # VALID! Tambahkan
                        cleaned_items.append(clean)
                        print(f"  ✓ Kode item valid: '{clean}'")
                
                # SIMPAN SEBAGAI DYNAMIC SKUS
                self.dynamic_skus = sorted(cleaned_items)
                
                # GABUNGKAN DENGAN DEFAULT SKUS
                combined_skus = self.default_skus.copy()
                for sku in self.dynamic_skus:
                    if sku not in combined_skus:  # Hanya tambah jika belum ada
                        combined_skus.append(sku)
                        # Tambah ke mapping juga
                        self.sku_mapping[sku] = sku
                
                self.all_skus = sorted(combined_skus)
                
                print(f"✓ Kode item terdeteksi:")
                print(f"  - Default SKU: {len(self.default_skus)} kode")
                print(f"  - Dynamic SKU: {len(self.dynamic_skus)} kode")
                print(f"  - Total SKU: {len(self.all_skus)} kode")
                
                # Tampilkan kode baru yang ditemukan
                new_skus = [sku for sku in self.dynamic_skus if sku not in self.default_skus]
                if new_skus:
                    print(f"  - Kode baru ditemukan: {new_skus}")
                
                # BUAT TEMPLATE COLUMNS
                self.template_columns = [
                    'NAMA PELANGGAN', 'TOTAL PENJUALAN'
                ] + self.all_skus + [
                    'SKU YANG SUDAH MASUK', 'SKU YANG BELUM MASUK', 
                    'JML SKU MASUK', '%', 'JML SKU TIDAK MASUK', '%', 'PERINGKAT SKU'
                ]
                
                print(f"📋 Template kolom: {len(self.template_columns)} kolom ({len(self.all_skus)} SKU)")
            
            return {
                'df': df_clean,
                'columns': {
                    'tanggal': tanggal_col,
                    'kd_item': kd_item_col,
                    'nama': nama_col,
                    'jml': jml_col,
                    'total': total_col
                },
                'file_type': file_type
            }
            
        except Exception as e:
            print(f"✗ Error saat ekstrak data: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
    
    def group_by_customer(self, data):
        """Kelompokkan data by customer - VERSI DIPERBAIKI DENGAN DEBUG"""
        if not data:
            print("✗ Data kosong di group_by_customer")
            return {}
        
        df = data['df']
        cols = data['columns']
        file_type = data.get('file_type', 'Unknown')
        
        print(f"\n{'='*60}")
        print(f"[DEBUG GROUP BY CUSTOMER - {file_type}]")
        print(f"{'='*60}")
        print(f"DataFrame shape: {df.shape}")
        print(f"Kolom mapping:")
        print(f"  - Tanggal: {cols.get('tanggal', 'Tidak ditemukan')}")
        print(f"  - Kd.Item: {cols.get('kd_item', 'Tidak ditemukan')}")
        print(f"  - Nama: {cols.get('nama', 'Tidak ditemukan')}")
        print(f"  - Jumlah: {cols.get('jml', 'Tidak ditemukan')}")
        print(f"  - Total: {cols.get('total', 'Tidak ditemukan')}")
        
        # Pastikan kolom yang diperlukan ada
        required_cols = ['nama', 'kd_item', 'jml']
        for col_key in required_cols:
            col_name = cols.get(col_key)
            if col_name is None:
                print(f"✗ Kolom {col_key} TIDAK DITEMUKAN!")
                return {}
            if col_name not in df.columns:
                print(f"✗ Kolom {col_key} ('{col_name}') TIDAK ADA di DataFrame!")
                print(f"  Kolom yang ada: {list(df.columns)}")
                return {}
        
        nama_col = cols['nama']
        kd_item_col = cols['kd_item']
        jml_col = cols['jml']
        
        try:
            # TAMPILKAN SAMPEL DATA UNTUK DEBUG
            print(f"\n[SAMPEL DATA RAW - 10 BARIS PERTAMA]")
            sample_df = df[[nama_col, kd_item_col, jml_col]].head(10)
            for idx, row in sample_df.iterrows():
                print(f"  Baris {idx}: {row[nama_col]} | {row[kd_item_col]} | {row[jml_col]}")
            
            # Group by customer dan kode item
            print(f"\n[MEMPROSES GROUPING...]")
            grouped = df.groupby([nama_col, kd_item_col])[jml_col].sum().reset_index()
            
            print(f"  Jumlah grup: {len(grouped)}")
            
            customer_data = defaultdict(lambda: {'total_sales': 0, 'skus': defaultdict(float)})
            
            # Counter untuk statistik
            sku_counter = defaultdict(int)
            
            for idx, row in grouped.iterrows():
                customer = row[nama_col]
                kode_item = str(row[kd_item_col]).strip().upper() if pd.notna(row[kd_item_col]) else ""
                qty = float(row[jml_col]) if pd.notna(row[jml_col]) else 0
                
                # Debug setiap 10 baris
                if idx < 10:  # Tampilkan 10 baris pertama
                    print(f"  Processing {idx}: {customer[:20]}... | {kode_item} | {qty}")
                
                # Bersihkan kode item
                # Hapus spasi, titik, karakter aneh
                kode_item_clean = re.sub(r'\s+', '', kode_item)  # Hapus spasi
                kode_item_clean = re.sub(r'[^A-Z0-9]', '', kode_item_clean)  # Hanya A-Z 0-9

                # ================ CEK VALIDITAS KODE ITEM ================
                # 1. Skip jika kosong
                if not kode_item_clean or kode_item_clean == '':
                    continue

                # 2. Skip jika ini nama kolom
                # List nama kolom yang harus di-skip (setelah dibersihkan)
                skip_columns = [
                    'KDITEM', 'KODEITEM', 'ITEM', 'KODE', 'KD',
                    'NAMA', 'PELANGGAN', 'CUSTOMER',
                    'JML', 'JUMLAH', 'QTY',
                    'TOTAL', 'HARGA', 'NILAI',
                    'TANGGAL', 'DATE', 'TGL'
                ]

                if kode_item_clean in skip_columns:
                    print(f"  ⚠ Skip '{kode_item}' (nama kolom)")
                    continue

                # 3. Skip jika format tidak valid untuk SKU
                # SKU biasanya: 2-6 karakter, mengandung huruf, tidak semua angka
                if (len(kode_item_clean) < 2 or len(kode_item_clean) > 6 or 
                    kode_item_clean.isdigit() or 
                    not any(c.isalpha() for c in kode_item_clean)):
                    print(f"  ⚠ Skip '{kode_item_clean}' (format tidak valid untuk SKU)")
                    continue

                # 4. Skip jika mengandung "ITEM" dan pendek (biasanya nama kolom)
                if 'ITEM' in kode_item_clean and len(kode_item_clean) <= 6:
                    print(f"  ⚠ Skip '{kode_item_clean}' (nama kolom 'ITEM')")
                    continue
                    
                                        
                # Map kode item
                sku = self.sku_mapping.get(kode_item_clean, kode_item_clean)
                
                # JIKA kode baru tidak ada di mapping, TAMBAHKAN ke dynamic_skus
                if kode_item_clean and kode_item_clean not in self.sku_mapping:
                    self.sku_mapping[kode_item_clean] = kode_item_clean
                    if kode_item_clean not in self.dynamic_skus:
                        self.dynamic_skus.append(kode_item_clean)
                        print(f"  ⚠ Tambah kode baru: '{kode_item_clean}'")
                
                # Validasi: cek apakah ini kode baru atau default
                if sku in self.default_skus:
                    # Kode default (normal)
                    pass
                elif sku in self.dynamic_skus:
                    # Kode dynamic (baru)
                    print(f"  ℹ Kode dynamic: '{sku}'")
                else:
                    # Kode tidak dikenal (jarang terjadi)
                    print(f"  ⚠ Kode tidak dikenal: '{sku}'")
                # ==============================================
                
                if pd.notna(customer) and qty > 0:
                    customer_name = str(customer).strip().upper()
                    customer_data[customer_name]['total_sales'] += qty
                    customer_data[customer_name]['skus'][sku] += qty
                    
                    sku_counter[sku] += 1
            
            print(f"\n[STATISTIK HASIL GROUPING]")
            print(f"✓ Total customer unik: {len(customer_data)}")
            print(f"✓ Distribusi SKU:")
            for sku, count in sorted(sku_counter.items(), key=lambda x: x[1], reverse=True):
                if count > 0:
                    print(f"  - {sku}: {count} customer")
            
            # Tampilkan sample hasil
            print(f"\n[SAMPEL HASIL - 3 CUSTOMER PERTAMA]")
            customers = list(customer_data.keys())[:3]
            for cust in customers:
                data = customer_data[cust]
                skus_with_qty = {sku: qty for sku, qty in data['skus'].items() if qty > 0}
                print(f"  {cust[:30]}...: {len(skus_with_qty)} SKU, Total: {data['total_sales']:.0f}")
                for sku, qty in list(skus_with_qty.items())[:5]:  # Tampilkan 5 pertama
                    print(f"    - {sku}: {qty}")
                if len(skus_with_qty) > 5:
                    print(f"    ... dan {len(skus_with_qty)-5} SKU lainnya")
            
            return dict(customer_data)
            
        except Exception as e:
            print(f"\n✗ ERROR dalam group_by_customer: {str(e)}")
            import traceback
            traceback.print_exc()
            
            # Debug tambahan
            print(f"\n[DEBUG INFO]")
            print(f"  Nama kolom: {nama_col}, tipe: {type(df[nama_col].iloc[0]) if len(df) > 0 else 'N/A'}")
            print(f"  Kd.Item kolom: {kd_item_col}, tipe: {type(df[kd_item_col].iloc[0]) if len(df) > 0 else 'N/A'}")
            print(f"  Jumlah kolom: {jml_col}, tipe: {type(df[jml_col].iloc[0]) if len(df) > 0 else 'N/A'}")
            
            return {}
    
    def validate_data_extraction(self, df, cols):
        """Validasi hasil ekstraksi data"""
        print(f"\n[VALIDASI DATA EKSTRAKSI]")
        
        valid = True
        
        # 1. Cek kolom yang diperlukan
        required = ['nama', 'kd_item', 'jml']
        for key in required:
            if key not in cols or cols[key] is None:
                print(f"✗ Kolom {key} tidak ditemukan")
                valid = False
            elif cols[key] not in df.columns:
                print(f"✗ Kolom {key} ('{cols[key]}') tidak ada di DataFrame")
                valid = False
        
        # 2. Cek data tidak kosong
        if len(df) == 0:
            print(f"✗ DataFrame kosong")
            valid = False
        
        # 3. Cek sample data
        if valid and len(df) > 0:
            print(f"✓ Sample data (baris 1-3):")
            for i in range(min(3, len(df))):
                row_info = []
                for key in ['nama', 'kd_item', 'jml']:
                    if key in cols and cols[key] is not None:
                        val = df.iloc[i][cols[key]]
                        row_info.append(f"{key}: {val}")
                print(f"  Baris {i}: {', '.join(row_info)}")
        
        # 4. Hitung statistik
        if valid and 'jml' in cols and cols['jml'] in df.columns:
            total_qty = df[cols['jml']].sum()
            avg_qty = df[cols['jml']].mean()
            print(f"✓ Statistik jumlah:")
            print(f"  - Total: {total_qty:.0f}")
            print(f"  - Rata-rata: {avg_qty:.2f}")
            print(f"  - Min: {df[cols['jml']].min():.0f}")
            print(f"  - Max: {df[cols['jml']].max():.0f}")
        
        return valid

    def create_template_data(self, customer_data):
        """Buat template data dari customer_data"""
        if not customer_data:
            print("Tidak ada data customer untuk dibuat template")
            return pd.DataFrame(columns=self.template_columns)
        
        if not self.template_columns:
            # Buat template columns jika belum ada
            self.template_columns = [
                'NAMA PELANGGAN', 'TOTAL PENJUALAN'
            ] + self.all_skus + [
                'SKU YANG SUDAH MASUK', 'SKU YANG BELUM MASUK', 
                'JML SKU MASUK', '%', 'JML SKU TIDAK MASUK', '%', 'PERINGKAT SKU'
            ]
        
        print(f"📊 Membuat template dengan {len(self.all_skus)} SKU:")
        print(f"  - Default: {len(self.default_skus)} SKU")
        print(f"  - Dynamic: {len(self.dynamic_skus)} SKU")
        
        template_rows = []
        
        # Sort customers by Jumlah SKU yang dibeli (bukan total sales)
        # Buat dulu list dengan jumlah SKU untuk sorting
        customers_with_sku_count = []
        
        for customer_name, data in customer_data.items():
            # Hitung jumlah SKU yang dibeli
            sku_count = 0
            for sku in self.all_skus:  # <- GUNAKAN self.all_skus (gabungan)
                if data['skus'].get(sku, 0) > 0:
                    sku_count += 1
            
            customers_with_sku_count.append({
                'name': customer_name,
                'data': data,
                'sku_count': sku_count,
                'total_sales': data['total_sales']
            })
        
        # Urutkan berdasarkan: 1. Jumlah SKU (desc), 2. Total sales (desc)
        sorted_customers = sorted(
            customers_with_sku_count,
            key=lambda x: (-x['sku_count'], -x['total_sales'])
        )
        
        ranking = 1
        for customer_info in sorted_customers:
            customer_name = customer_info['name']
            data = customer_info['data']
            
            row = [customer_name, data['total_sales']]
            
            # Tambahkan quantity untuk setiap SKU
            for sku in self.all_skus:
                row.append(data['skus'].get(sku, 0))
            
            # HITUNG SKU YANG SUDAH MASUK (dibeli) DAN YANG BELUM MASUK
            sku_sudah_masuk = []
            sku_belum_masuk = []
            
            for sku in self.all_skus:
                qty = data['skus'].get(sku, 0)
                if qty > 0:
                    sku_sudah_masuk.append(sku)
                else:
                    sku_belum_masuk.append(sku)
            
            # Gabungkan menjadi string daftar nama produk
            sku_sudah_str = ', '.join(sku_sudah_masuk) if sku_sudah_masuk else '-'
            sku_belum_str = ', '.join(sku_belum_masuk) if sku_belum_masuk else '-'
            
            # Hitung jumlah SKU dan persentase berdasarkan total
            jumlah_sku_masuk = len(sku_sudah_masuk)
            TOTAL_SKU = len(self.all_skus)  # Total semua SKU 
            
            # Hitung persentase
            if TOTAL_SKU > 0:
                percentage_masuk = (jumlah_sku_masuk / TOTAL_SKU) * 100
                percentage_belum = 100 - percentage_masuk
            else:
                percentage_masuk = 0
                percentage_belum = 0
            
            jumlah_sku_tidak_masuk = len(sku_belum_masuk)
            
            row.extend([
                sku_sudah_str,  # Nama SKU yang sudah masuk
                sku_belum_str,  # Nama SKU yang belum masuk
                f"{jumlah_sku_masuk}/{TOTAL_SKU}",  # Format: 11/20
                f"{percentage_masuk:.1f}%",  # Persentase masuk
                f"{jumlah_sku_tidak_masuk}/{TOTAL_SKU}",  # Format: 9/20
                f"{percentage_belum:.1f}%",  # Persentase tidak masuk
                ranking  # Peringkat berdasarkan jumlah SKU
            ])
            
            template_rows.append(row)
            ranking += 1
        
        print(f"✓ Template data dibuat: {len(template_rows)} baris, {len(self.template_columns)} kolom")
        
        return pd.DataFrame(template_rows, columns=self.template_columns)
            
    def add_total_pencapaian(self, template_df):
        """Add TOTAL PENCAPAIAN row"""
        if template_df.empty:
            return template_df
        
        try:
            # Ambil baris customer (bukan TOTAL / %)
            customer_rows = template_df[
                ~template_df['NAMA PELANGGAN'].str.contains('TOTAL|%', na=False)
            ].copy()
            
            if customer_rows.empty:
                return template_df
            
            total_customers = len(customer_rows)
            
            # Buat baris TOTAL PENCAPAIAN
            total_row = ["TOTAL PENCAPAIAN", total_customers]
            
            # Untuk setiap SKU, hitung jumlah customer yang membeli
            for sku in self.all_skus:
                customers_with_sku = (customer_rows[sku] > 0).sum()
                total_row.append(f"{customers_with_sku} dari {total_customers}")
            
            # HITUNG SKU YANG SUDAH MASUK DAN BELUM MASUK (dari seluruh customer)
            sku_sudah_masuk_all = []
            sku_belum_masuk_all = []
            
            for sku in self.all_skus:
                if (customer_rows[sku] > 0).any():
                    sku_sudah_masuk_all.append(sku)
                else:
                    sku_belum_masuk_all.append(sku)
            
            # Gabungkan menjadi string daftar nama
            sku_sudah_str = ', '.join(sku_sudah_masuk_all) if sku_sudah_masuk_all else '-'
            sku_belum_str = ', '.join(sku_belum_masuk_all) if sku_belum_masuk_all else '-'
            
            # Hitung jumlah dan persentase untuk TOTAL berdasarkan 20 SKU
            jumlah_sku_masuk_total = len(sku_sudah_masuk_all)
            TOTAL_SKU = len(self.all_skus)  # Total semua SKU (20)
            
            # Hitung persentase
            if TOTAL_SKU > 0:
                percentage_masuk_total = (jumlah_sku_masuk_total / TOTAL_SKU) * 100
                percentage_belum_total = 100 - percentage_masuk_total
            else:
                percentage_masuk_total = 0
                percentage_belum_total = 0
            
            jumlah_sku_tidak_masuk_total = len(sku_belum_masuk_all)
            
            total_row.extend([
                sku_sudah_str,
                sku_belum_str,
                f"{jumlah_sku_masuk_total}/{TOTAL_SKU}",  # Format: 11/20
                f"{percentage_masuk_total:.1f}%",  # Persentase masuk
                f"{jumlah_sku_tidak_masuk_total}/{TOTAL_SKU}",  # Format: 9/20
                f"{percentage_belum_total:.1f}%",  # Persentase tidak masuk
                ""
            ])
            
            # Pastikan panjang total_row sesuai
            expected_columns = len(template_df.columns)
            if len(total_row) != expected_columns:
                if len(total_row) < expected_columns:
                    total_row.extend([""] * (expected_columns - len(total_row)))
                else:
                    total_row = total_row[:expected_columns]
            
            # Tambahkan baris TOTAL PENCAPAIAN
            total_df = pd.DataFrame([total_row], columns=template_df.columns)
            result_df = pd.concat([template_df, total_df], ignore_index=True)
            
            print(f"✓ TOTAL PENCAPAIAN ditambahkan")
            
            return result_df
            
        except Exception as e:
            print(f"✗ Error dalam add_total_pencapaian: {str(e)}")
            return template_df
    
    def add_percentage_row(self, template_df):
        """Add % PENCAPAIAN row setelah TOTAL"""
        if template_df.empty:
            return template_df
        
        try:
            # Cari index baris TOTAL PENCAPAIAN
            total_indices = template_df[template_df['NAMA PELANGGAN'] == 'TOTAL PENCAPAIAN'].index
            
            if len(total_indices) == 0:
                return template_df
            
            total_idx = total_indices[0]
            
            # Ambil baris customer (sebelum TOTAL)
            customer_rows = template_df.iloc[:total_idx]
            
            if len(customer_rows) == 0:
                return template_df
            
            total_customers = len(customer_rows)
            
            # Buat baris % PENCAPAIAN
            percentage_row = ["% PENCAPAIAN", ""]
            
            # Hitung persentase untuk setiap SKU
            for sku in self.all_skus:
                customers_with_sku = (customer_rows[sku] > 0).sum()
                
                if total_customers > 0:
                    percentage = (customers_with_sku / total_customers) * 100
                else:
                    percentage = 0
                
                percentage_row.append(f"{percentage:.1f}%")
            
            # Kolom achievement kosong untuk baris %
            # Sesuaikan dengan jumlah kolom yang ada sekarang (2 untuk SKU nama, 4 untuk jumlah+persen)
            percentage_row.extend(["", "", "", "", "", "", ""])
            
            # Pastikan panjang percentage_row sesuai
            expected_columns = len(template_df.columns)
            if len(percentage_row) != expected_columns:
                if len(percentage_row) < expected_columns:
                    percentage_row.extend([""] * (expected_columns - len(percentage_row)))
                else:
                    percentage_row = percentage_row[:expected_columns]
            
            # Buat DataFrame untuk baris %
            percentage_df = pd.DataFrame([percentage_row], columns=template_df.columns)
            
            # Gabungkan: semua data sebelum TOTAL + TOTAL + % PENCAPAIAN + sisanya
            before_total = template_df.iloc[:total_idx]
            the_total = template_df.iloc[total_idx:total_idx+1]
            after_total = template_df.iloc[total_idx+1:]
            
            result_df = pd.concat([before_total, the_total, percentage_df, after_total], ignore_index=True)
            
            print(f"✓ % PENCAPAIAN ditambahkan")
            
            return result_df
            
        except Exception as e:
            print(f"✗ Error dalam add_percentage_row: {str(e)}")
            return template_df
        
    def process_excel_to_template(self, file_path):
        """Main processing function dengan validasi"""
        try:
            print(f"\n{'='*60}")
            print(f"PROSES FILE: {os.path.basename(file_path)}")
            print(f"{'='*60}")
            
            # 1. Ekstrak data dari file
            lpd_data = self.extract_lpd_data(file_path)
            if lpd_data is None:
                print("✗ Gagal mengekstrak data dari file")
                return self.create_fallback_template(file_path)
            
            # 1b. VALIDASI DATA EKSTRAKSI
            is_valid = self.validate_data_extraction(lpd_data['df'], lpd_data['columns'])
            if not is_valid:
                print("✗ Data ekstraksi tidak valid!")
                return self.create_fallback_template(file_path)
            
            # 2. Kelompokkan data by customer
            customer_data = self.group_by_customer(lpd_data)
            
            if not customer_data:
                print("✗ Tidak ada data customer yang valid")
                return self.create_fallback_template(file_path)
            
            # 3. Buat template data
            template_df = self.create_template_data(customer_data)
            if template_df.empty:
                print("✗ Gagal membuat template data")
                return self.create_fallback_template(file_path)
            
            # 4. Tambahkan TOTAL PENCAPAIAN
            template_df = self.add_total_pencapaian(template_df)
            
            # 5. Tambahkan kolom SISTEM di awal
            sistem_type = lpd_data.get('file_type', 'TIDAK DIKETAHUI')
            template_df.insert(0, 'SISTEM', sistem_type)
            
            print(f"\n{'='*60}")
            print(f"PROSES SELESAI!")
            print(f"✓ Total baris: {len(template_df)}")
            print(f"✓ Total kolom: {len(template_df.columns)}")
            print(f"✓ Sistem: {sistem_type}")
            print(f"✓ File output siap disimpan")
            print(f"{'='*60}")
            
            return template_df
            
        except Exception as e:
            print(f"✗ Error dalam process_excel_to_template: {str(e)}")
            import traceback
            traceback.print_exc()
            return self.create_fallback_template(file_path)
            
    def create_fallback_template(self, file_path=None):
        """Create fallback template when processing fails"""
        print("Membuat fallback template...")
        
        customers = [
            'PT MULTI BUAH SEGAR',
            'DUTA BUAH LESTARI', 
            'GRAND LUCKY MARKET',
            'FARMERS MARKET JAYA',
            'TOKO BUAH MURAH'
        ]
        
        template_rows = []
        
        for i, customer in enumerate(customers):
            total_sales = np.random.randint(500, 5000)
            row = [customer, total_sales]
            
            sku_quantities = []
            sku_sudah_masuk = []
            sku_belum_masuk = []
            
            for sku in self.all_skus:
                if np.random.random() > 0.7:
                    qty = np.random.randint(10, 100)
                    sku_quantities.append(qty)
                    sku_sudah_masuk.append(sku)
                else:
                    sku_quantities.append(0)
                    sku_belum_masuk.append(sku)
            
            row.extend(sku_quantities)
            
            # Gabungkan menjadi string daftar nama
            sku_sudah_str = ', '.join(sku_sudah_masuk) if sku_sudah_masuk else '-'
            sku_belum_str = ', '.join(sku_belum_masuk) if sku_belum_masuk else '-'
            
            # Hitung jumlah dan persentase berdasarkan 20 SKU
            jumlah_sku_masuk = len(sku_sudah_masuk)
            TOTAL_SKU = len(self.all_skus)  # Total semua SKU (20)
            
            # Hitung persentase
            if TOTAL_SKU > 0:
                percentage_masuk = (jumlah_sku_masuk / TOTAL_SKU) * 100
                percentage_belum = 100 - percentage_masuk
            else:
                percentage_masuk = 0
                percentage_belum = 0
            
            jumlah_sku_tidak_masuk = len(sku_belum_masuk)
            
            row.extend([
                sku_sudah_str,
                sku_belum_str,
                f"{jumlah_sku_masuk}/{TOTAL_SKU}",  # Format: 11/20
                f"{percentage_masuk:.1f}%",  # Persentase masuk
                f"{jumlah_sku_tidak_masuk}/{TOTAL_SKU}",  # Format: 9/20
                f"{percentage_belum:.1f}%",  # Persentase tidak masuk
                i + 1
            ])
            
            template_rows.append(row)
        
        template_df = pd.DataFrame(template_rows, columns=self.template_columns)
        
        # Ganti CABANG dengan SISTEM
        # Coba tebak sistem dari nama file jika ada
        sistem_type = "TIDAK DIKETAHUI"
        if file_path:
            file_name = os.path.basename(file_path).upper()
            if "KONSINYASI" in file_name:
                sistem_type = "KONSINYASI"
            elif "PENJUALAN" in file_name:
                sistem_type = "PENJUALAN"
        
        # Tambahkan kolom SISTEM
        template_df.insert(0, 'SISTEM', sistem_type)
        
        template_df = self.add_total_pencapaian(template_df)
        
        print(f"✓ Fallback template dibuat: {len(template_df)} baris")
        print(f"✓ Sistem: {sistem_type}")
        
        return template_df


# ===== TEST CODE =====
if __name__ == "__main__":
    processor = LPDProcessor()
    
    # Test dengan file konsinyasi
    print("\n" + "="*60)
    print("TEST FILE KONSINYASI")
    print("="*60)
    test_file_kons = "KONSINYASI - NGL (1-10 DES) - ver 2.xlsx"
    
    if os.path.exists(test_file_kons):
        result_kons = processor.process_excel_to_template(test_file_kons)
        output_name_kons = "hasil_konsinyasi_test.xlsx"
        result_kons.to_excel(output_name_kons, index=False)
        print(f"✅ Konsinyasi berhasil → {output_name_kons}")
    else:
        print(f"⚠ File {test_file_kons} tidak ditemukan")
    
    # Test dengan file penjualan
    print("\n" + "="*60)
    print("TEST FILE PENJUALAN")
    print("="*60)
    processor2 = LPDProcessor()
    test_file_penj = "PENJUALAN - UTM (1-10 DES) - ver 2 (1).xlsx"
    
    if os.path.exists(test_file_penj):
        result_penj = processor2.process_excel_to_template(test_file_penj)
        output_name_penj = "hasil_penjualan_test.xlsx"
        result_penj.to_excel(output_name_penj, index=False)
        print(f"✅ Penjualan berhasil → {output_name_penj}")
    else:
        print(f"⚠ File {test_file_penj} tidak ditemukan")