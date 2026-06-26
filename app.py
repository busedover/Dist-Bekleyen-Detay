import streamlit as st
import pandas as pd
import numpy as np

# Sayfa ayarları
st.set_page_config(page_title="CPD Order & Stock Allocator Dashboard", layout="wide")

st.title("📦 CPD Sipariş Karşılama ve NIV Dashboard")
st.subheader("Distribütör Bekleyen Sipariş, Katalog, Stok ve Fiyat Analizi")

# --- DOSYA YÜKLEME ALANI (4 Dosya) ---
st.sidebar.header("📂 Excel Dosyalarını Yükleyin")
orders_file = st.sidebar.file_uploader("1. Bekleyen Siparişler Excel'i", type=["xlsx", "xls"])
catalog_file = st.sidebar.file_uploader("2. Katalog Raporu Excel'i (Köprü)", type=["xlsx", "xls"])
stock_file = st.sidebar.file_uploader("3. Güncel Stok Excel'i", type=["xlsx", "xls"])
prices_file = st.sidebar.file_uploader("4. Güncel Fiyat Listesi Excel'i (Fiyatlar)", type=["xlsx", "xls"])

if orders_file and catalog_file and stock_file and prices_file:
    # Verileri oku
    df_orders = pd.read_excel(orders_file, engine="openpyxl")
    df_catalog = pd.read_excel(catalog_file, engine="openpyxl")
    df_stock = pd.read_excel(stock_file, engine="openpyxl")
    df_prices_raw = pd.read_excel(prices_file, engine="openpyxl")
    
    # Kolon isimlerindeki boşlukları temizleme
    df_orders.columns = df_orders.columns.str.strip()
    df_catalog.columns = df_catalog.columns.str.strip()
    df_stock.columns = df_stock.columns.str.strip()
    df_prices_raw.columns = df_prices_raw.columns.str.strip()
    
    st.success("✅ Tüm Excel dosyaları başarıyla yüklendi!")
    
    # Beklenen Kritik Kolon İsimleri Tanımları
    siparis_barkod_col = "Barkod"
    katalog_material_col = "Material"
    katalog_ean_col = "EAN Cod-UM"
    stok_material_col = "Material"
    stok_net_avail_col = "Net avail."
    
    # Fiyat dosyasındaki tam sütun isimleriniz
    fiyat_barkod_col = "EAN Cod-UM"
    fiyat_deger_col = "Catal.price"
    
    # Kolon Kontrolleri
    catalog_ok = katalog_material_col in df_catalog.columns and katalog_ean_col in df_catalog.columns
    stock_ok = stok_material_col in df_stock.columns and stok_net_avail_col in df_stock.columns
    orders_ok = siparis_barkod_col in df_orders.columns and "Sipariş Miktarı" in df_orders.columns
    prices_ok = fiyat_barkod_col in df_prices_raw.columns and fiyat_deger_col in df_prices_raw.columns
    
    if catalog_ok and stock_ok and orders_ok and prices_ok:
        # --- VERİ TİPİ TEMİZLEME VE UYUMLAŞTIRMA (Kritik Aşama) ---
        # Barkod kolonlarını string (yazı) formatına getiriyoruz ve boşlukları siliyoruz
        df_orders[siparis_barkod_col] = df_orders[siparis_barkod_col].astype(str).str.strip().str.split('.').str[0]
        df_catalog[katalog_ean_col] = df_catalog[katalog_ean_col].astype(str).str.strip().str.split('.').str[0]
        df_prices_raw[fiyat_barkod_col] = df_prices_raw[fiyat_barkod_col].astype(str).str.strip().str.split('.').str[0]
        
        # Malzeme (Material) kolonlarını temizleme (Örn: başında sıfır olan sayıları eşitlemek için)
        df_catalog[katalog_material_col] = df_catalog[katalog_material_col].astype(str).str.strip().str.lstrip('0')
        df_stock[stok_material_col] = df_stock[stok_material_col].astype(str).str.strip().str.lstrip('0')
        
        # Fiyat listesini temizleyelim ve "Barkod" - "Fiyat" olarak standartlaştıralım
        df_prices = df_prices_raw[[fiyat_barkod_col, fiyat_deger_col]].dropna().drop_duplicates(subset=[fiyat_barkod_col])
        df_prices.rename(columns={fiyat_barkod_col: "Barkod", fiyat_deger_col: "Fiyat"}, inplace=True)

        # 1. Stok dosyasında Malzeme bazında toplam 'Net avail.' miktarını hesapla
        df_stock_grouped = df_stock.groupby(stok_material_col)[stok_net_avail_col].sum().reset_index()
        
        # 2. Katalog dosyasından Material -> EAN (Barkod) eşleşmesini al
        df_cat_clean = df_catalog[[katalog_material_col, katalog_ean_col]].dropna().drop_duplicates()
        
        # 3. Kataloğu Stok ile birleştirerek Barkod bazlı stok durumunu elde et
        df_barcode_stock = pd.merge(df_cat_clean, df_stock_grouped, on=katalog_material_col, how="inner")
        
        # Aynı barkoda sahip birden fazla malzeme kodu olma ihtimaline karşı barkod bazında stokları topla
        df_final_stock = df_barcode_stock.groupby(katalog_ean_col)[stok_net_avail_col].sum().reset_index()
        df_final_stock.rename(columns={katalog_ean_col: "Barkod"}, inplace=True)
        
        # 4. Sipariş dosyasına otomatik fiyatları eşleştirme
        if "Fiyat" in df_orders.columns:
            df_orders = df_orders.drop(columns=["Fiyat"])
        df_orders_with_price = pd.merge(df_orders, df_prices, on="Barkod", how="left")
        
        # 5. Sipariş verilerini Barkod bazlı son stok durumu ile birleştir
        df_merged = pd.merge(df_orders_with_price, df_final_stock, on="Barkod", how="left")
        df_merged[stok_net_avail_col] = df_merged[stok_net_avail_col].fillna(0)
        
        # Karşılama hesaplama mantığı
        df_merged['Karşılanan Adet'] = np.minimum(df_merged['Sipariş Miktarı'], df_merged[stok_net_avail_col])
        df_merged['Fiyat'] = df_merged['Fiyat'].fillna(0)
        
        # NIV Tutar Hesaplamaları
        df_merged['Toplam Talep Edilen NIV'] = df_merged['Sipariş Miktarı'] * df_merged['Fiyat']
        df_merged['Karşılanan NIV'] = df_merged['Karşılanan Adet'] * df_merged['Fiyat']
        df_merged['Kayıp (Karşılanamayan) NIV'] = (df_merged['Sipariş Miktarı'] - df_merged['Karşılanan Adet']) * df_merged['Fiyat']
        df_merged['Fill Rate %'] = (df_merged['Karşılanan Adet'] / df_merged['Sipariş Miktarı'] * 100).fillna(0)

        # KPI Kartları
        total_requested_niv = df_merged['Toplam Talep Edilen NIV'].sum()
        total_allocated_niv = df_merged['Karşılanan NIV'].sum()
        total_lost_niv = df_merged['Kayıp (Karşılanamayan) NIV'].sum()
        overall_fill_rate = (total_allocated_niv / total_requested_niv * 100) if total_requested_niv > 0 else 0
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Toplam Bekleyen Sipariş NIV", f"₺{total_requested_niv:,.2f}")
        col2.metric("Karşılanabilir (Allocated) NIV", f"₺{total_allocated_niv:,.2f}", delta=f"{overall_fill_rate:.1f}% Fill Rate")
        col3.metric("Kayıp (Stoksuzluk) NIV", f"₺{total_lost_niv:,.2f}", delta_color="inverse")
        
        # Sonuç Tablosu
        st.markdown("---")
        st.subheader("📋 Karşılama Detay Tablosu")
        # Dinamik kolon kontrolü
        display_cols = []
        possible_cols = {
            'Müşteri Adı': 'Müşteri Adı',
            'Barkod': 'Barkod',
            'Ürün Adı': 'Ürün Adı',
            'Sipariş Miktarı': 'Sipariş Miktarı',
            stok_net_avail_col: stok_net_avail_col,
            'Karşılanan Adet': 'Karşılanan Adet',
            'Fiyat': 'Fiyat',
            'Toplam Talep Edilen NIV': 'Toplam Talep Edilen NIV',
            'Karşılanan NIV': 'Karşılanan NIV'
        }
        for col_key, col_val in possible_cols.items():
            if col_val in df_merged.columns:
                display_cols.append(col_val)
                
        st.dataframe(df_merged[display_cols].style.format({
            'Fiyat': '₺{:,.2f}' if 'Fiyat' in df_merged.columns else '{}',
            'Toplam Talep Edilen NIV': '₺{:,.2f}' if 'Toplam Talep Edilen NIV' in df_merged.columns else '{}',
            'Karşılanan NIV': '₺{:,.2f}' if 'Karşılanan NIV' in df_merged.columns else '{}'
        }))
    else:
        st.warning("⚠ Dosyalar yüklendi fakat aşağıdaki kolon isimlerinde uyuşmazlık var:")
        # Detaylı Hata Gösterimi
        if not orders_ok:
            st.error(f"❌ Sipariş Dosyasında Sorun Var!")
            st.write(f"Aranan Kolonlar: **'{siparis_barkod_col}'** ve **'Sipariş Miktarı'**")
            st.write(f"Dosyanızdaki Mevcut Kolonlar: {list(df_orders.columns)}")
        if not catalog_ok:
            st.error(f"❌ Katalog Dosyasında Sorun Var!")
            st.write(f"Aranan Kolonlar: **'{katalog_material_col}'** ve **'{katalog_ean_col}'**")
            st.write(f"Dosyanızdaki Mevcut Kolonlar: {list(df_catalog.columns)}")
        if not stock_ok:
            st.error(f"❌ Stok Dosyasında Sorun Var!")
            st.write(f"Aranan Kolonlar: **'{stok_material_col}'** ve **'{stok_net_avail_col}'**")
            st.write(f"Dosyanızdaki Mevcut Kolonlar: {list(df_stock.columns)}")
        if not prices_ok:
            st.error(f"❌ Fiyat Dosyasında Sorun Var!")
            st.write(f"Aranan Kolonlar: **'{fiyat_barkod_col}'** ve **'{fiyat_deger_col}'**")
            st.write(f"Dosyanızdaki Mevcut Kolonlar: {list(df_prices_raw.columns)}")
else:
    st.info("💡 Lütfen sol menüden 'Sipariş', 'Katalog', 'Stok' ve 'Fiyat' excel dosyalarını yükleyin. Analiz otomatik olarak başlayacaktır.")
