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
        # --- VERİ TEMİZLEME VE FORMAT STANDARTLAŞTIRMA ---
        # 1. Sipariş Barkod Temizliği
        df_orders[siparis_barkod_col] = df_orders[siparis_barkod_col].astype(str).str.strip().str.split('.').str[0]
        
        # 2. Katalog Temizliği (Material ve EAN)
        df_catalog[katalog_material_col] = df_catalog[katalog_material_col].astype(str).str.strip().str.lstrip('0')
        df_catalog[katalog_ean_col] = df_catalog[katalog_ean_col].astype(str).str.strip().str.split('.').str[0]
        
        # 3. Stok Temizliği (Material ve Net avail. sayısal dönüşümü)
        df_stock[stok_material_col] = df_stock[stok_material_col].astype(str).str.strip().str.lstrip('0')
        df_stock[stok_net_avail_col] = pd.to_numeric(df_stock[stok_net_avail_col], errors='coerce').fillna(0)
        
        # 4. Fiyat Listesi Temizliği
        df_prices_raw[fiyat_barkod_col] = df_prices_raw[fiyat_barkod_col].astype(str).str.strip().str.split('.').str[0]
        df_prices = df_prices_raw[[fiyat_barkod_col, fiyat_deger_col]].dropna().drop_duplicates(subset=[fiyat_barkod_col])
        df_prices.rename(columns={fiyat_barkod_col: "Barkod", fiyat_deger_col: "Fiyat"}, inplace=True)

        # --- ASIL CPD KÖPRÜ MANTIK ZİNCİRİ ---
        # 1. Adım: Stok dosyasındaki malzemelerin toplam stoklarını gruplayarak alalım
        df_stock_grouped = df_stock.groupby(stok_material_col)[stok_net_avail_col].sum().reset_index()
        
        # 2. Adım: Katalogdan sadece Material ve EAN (Barkod) sütunlarını çekelim
        df_cat_bridge = df_catalog[[katalog_material_col, katalog_ean_col]].dropna().drop_duplicates()
        
        # 3. Adım: Katalog köprüsüyle Stokları "Material" üzerinden birleştirelim
        df_merged_stock = pd.merge(df_cat_bridge, df_stock_grouped, on=katalog_material_col, how="inner")
        
        # 4. Adım: Aynı barkoda (EAN Cod-UM) karşılık gelen tüm malzeme stoklarını TOPLAYALIM
        df_barcode_stock_sum = df_merged_stock.groupby(katalog_ean_col)[stok_net_avail_col].sum().reset_index()
        
        # Kolon ismini standartlaştırarak siparişe hazır hale getirelim
        df_barcode_stock_sum.rename(columns={katalog_ean_col: "Barkod"}, inplace=True)

        # --- SİPARİŞ BİRLEŞTİRME VE FİYATLANDIRMA ---
        # 5. Adım: Siparişe Fiyatları Barkod üzerinden bağlayalım
        if "Fiyat" in df_orders.columns:
            df_orders = df_orders.drop(columns=["Fiyat"])
        df_orders_with_price = pd.merge(df_orders, df_prices, on="Barkod", how="left")
        
        # 6. Adım: Siparişe konsolide edilmiş barkod stoğunu bağlayalım
        df_final = pd.merge(df_orders_with_price, df_barcode_stock_sum, on="Barkod", how="left")
        df_final[stok_net_avail_col] = df_final[stok_net_avail_col].fillna(0)
        
        # --- HESAPLAMALAR ---
        df_final['Karşılanan Adet'] = np.minimum(df_final['Sipariş Miktarı'], df_final[stok_net_avail_col])
        df_final['Fiyat'] = df_final['Fiyat'].fillna(0)
        
        df_final['Toplam Talep Edilen NIV'] = df_final['Sipariş Miktarı'] * df_final['Fiyat']
        df_final['Karşılanan NIV'] = df_final['Karşılanan Adet'] * df_final['Fiyat']
        df_final['Kayıp (Karşılanamayan) NIV'] = (df_final['Sipariş Miktarı'] - df_final['Karşılanan Adet']) * df_final['Fiyat']
        df_final['Fill Rate %'] = (df_final['Karşılanan Adet'] / df_final['Sipariş Miktarı'] * 100).fillna(0)

        # --- YAZARAK ARAMA ÖZELLİKLİ FİLTRELEME ALANI ---
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Yazarak Arayın & Filtreleyin")
        
        tum_musteriler = sorted(df_final['Müşteri Adı'].dropna().unique().tolist())
        secilen_musteriler = st.sidebar.multiselect(
            "🏢 Müşteri Adı Yazın / Seçin", 
            options=tum_musteriler,
            placeholder="Müşteri ismi yazın..."
        )
        
        tum_barkodlar = sorted(df_final['Barkod'].dropna().unique().tolist())
        secilen_barkodlar = st.sidebar.multiselect(
            "🏷 Barkod Yazın / Seçin", 
            options=tum_barkodlar,
            placeholder="Barkod yazın..."
        )
        
        # Filtreleme uygulaması
        df_filtered = df_final.copy()
        if len(secilen_musteriler) > 0:
            df_filtered = df_filtered[df_filtered['Müşteri Adı'].isin(secilen_musteriler)]
        if len(secilen_barkodlar) > 0:
            df_filtered = df_filtered[df_filtered['Barkod'].isin(secilen_barkodlar)]

        # --- KPI KARTLARI ---
        total_requested_niv = df_filtered['Toplam Talep Edilen NIV'].sum()
        total_allocated_niv = df_filtered['Karşılanan NIV'].sum()
        total_lost_niv = df_filtered['Kayıp (Karşılanamayan) NIV'].sum()
        overall_fill_rate = (total_allocated_niv / total_requested_niv * 100) if total_requested_niv > 0 else 0
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Toplam Bekleyen Sipariş NIV", f"₺{total_requested_niv:,.2f}")
        col2.metric("Karşılanabilir (Allocated) NIV", f"₺{total_allocated_niv:,.2f}", delta=f"{overall_fill_rate:.1f}% Fill Rate")
        col3.metric("Kayıp (Stoksuzluk) NIV", f"₺{total_lost_niv:,.2f}", delta_color="inverse")
        
        # --- TABLO GÖSTERİMİ ---
        st.markdown("---")
        st.subheader("📋 Karşılama Detay Tablosu")
        
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
            if col_val in df_filtered.columns:
                display_cols.append(col_val)
                
        # Tablo stil formatlamasını pürüzsüzce kapatıyoruz
        st.dataframe(df_filtered[display_cols].style.format({
            'Fiyat': '₺{:,.2f}' if 'Fiyat' in df_filtered.columns else '{}',
            'Toplam Talep Edilen NIV': '₺{:,.2f}' if 'Toplam Talep Edilen NIV' in df_filtered.columns else '{}',
            'Karşılanan NIV': '₺{:,.2f}' if 'Karşılanan NIV' in df_filtered.columns else '{}'
        }))
    else:
        st.warning("⚠ Yüklenen Excel dosyalarındaki kolon isimlerini kontrol edin:")
        if not orders_ok:
            st.error(f"❌ Sipariş Dosyasında Sorun Var! Aranan: '{siparis_barkod_col}' ve 'Sipariş Miktarı'")
        if not catalog_ok:
            st.error(f"❌ Katalog Dosyasında Sorun Var! Aranan: '{katalog_material_col}' ve '{katalog_ean_col}'")
        if not stock_ok:
            st.error(f"❌ Stok Dosyasında Sorun Var! Aranan: '{stok_material_col}' ve '{stok_net_avail_col}'")
        if not prices_ok:
            st.error(f"❌ Fiyat Dosyasında Sorun Var! Aranan: '{fiyat_barkod_col}' ve '{fiyat_deger_col}'")
else:
    st.info("💡 Lütfen sol menüden 'Sipariş', 'Katalog', 'Stok' ve 'Fiyat' excel dosyalarını yükleyin. Eşleştirmeler otomatik tamamlanacaktır.")
