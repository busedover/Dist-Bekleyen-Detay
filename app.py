import streamlit as st
import pandas as pd
import numpy as np
import re

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
        
        # --- BARKOD TEMİZLEME MOTORU ---
        def gelismis_barkod_temizle(seri):
            s_str = seri.astype(str).str.strip()
            # Baş ve sondaki noktaları at
            s_str = s_str.apply(lambda x: re.sub(r'^\.+|\.+$', '', x))
            # Bilimsel gösterim ayarı
            def e_notation_duzelt(val):
                if 'e' in val.lower():
                    try:
                        return str(int(float(val)))
                    except:
                        pass
                return val
            s_str = s_str.apply(e_notation_duzelt)
            s_str = s_str.apply(lambda x: x.split('.')[0] if '.' in x else x)
            # Sadece sayıları koru
            s_str = s_str.apply(lambda x: re.sub(r'\D', '', x))
            return s_str

        # --- MALZEME TEMİZLEME MOTORU ---
        def malzeme_kodunu_temizle(seri):
            # Sayısal değerlere zorlamadan önce string temizliği
            s_str = seri.astype(str).str.strip()
            s_str = s_str.apply(lambda x: x.split('.')[0] if '.' in x else x)
            # Başındaki sıfırları ve harf dışı boşlukları temizle
            s_str = s_str.str.lstrip('0')
            return s_str

        # --- VERİ TEMİZLEME VE FORMAT STANDARTLAŞTIRMA ---
        df_orders[siparis_barkod_col] = gelismis_barkod_temizle(df_orders[siparis_barkod_col])
        
        # Malzemeleri metin tabanlı temizleyelim (Integer zorlaması eşleşmeyi bozmasın diye string olarak eşitleyeceğiz)
        df_catalog[katalog_material_col] = malzeme_kodunu_temizle(df_catalog[katalog_material_col])
        df_catalog[katalog_ean_col] = gelismis_barkod_temizle(df_catalog[katalog_ean_col])
        
        df_stock[stok_material_col] = malzeme_kodunu_temizle(df_stock[stok_material_col])
        df_stock[stok_net_avail_col] = pd.to_numeric(df_stock[stok_net_avail_col], errors='coerce').fillna(0)
        
        df_prices_raw[fiyat_barkod_col] = gelismis_barkod_temizle(df_prices_raw[fiyat_barkod_col])
        df_prices = df_prices_raw[[fiyat_barkod_col, fiyat_deger_col]].dropna().drop_duplicates(subset=[fiyat_barkod_col])
        df_prices.rename(columns={fiyat_barkod_col: "Barkod", fiyat_deger_col: "Fiyat"}, inplace=True)

        # --- CANLI ADIM ADIM TEŞHİS PANELİ ---
        st.info("🔍 **Eşleşme Teşhis Paneli** (Adım Adım Hata Tespiti)")
        
        t1, t2, t3, t4 = st.columns(4)
        t1.metric("Katalogdaki Satır Sayısı", len(df_catalog))
        t2.metric("Stoktaki Satır Sayısı", len(df_stock))
        
        # Ortak malzeme kodu sayısı
        katalog_malzemeler = set(df_catalog[katalog_material_col].dropna().unique())
        stok_malzemeler = set(df_stock[stok_material_col].dropna().unique())
        ortak_malzemeler = katalog_malzemeler.intersection(stok_malzemeler)
        
        t3.metric("Ortak Malzeme Kodu", len(ortak_malzemeler))
        
        # 1. Aşama: Stok Gruplama
        df_stock_grouped = df_stock.groupby(stok_material_col)[stok_net_avail_col].sum().reset_index()
        
        # 2. Aşama: Katalog Köprü Temizliği
        df_cat_bridge = df_catalog[[katalog_material_col, katalog_ean_col]].dropna().drop_duplicates()
        
        # 3. Aşama: Katalog ve Stok Birleştirme
        df_merged_stock = pd.merge(df_cat_bridge, df_stock_grouped, on=katalog_material_col, how="inner")
        
        t4.metric("Eşleşen Stok Satırı", len(df_merged_stock))

        # Hata Detayı Verme
        if len(ortak_malzemeler) == 0:
            st.error("❌ TEŞHİS: Katalog ve Stok dosyalarındaki Material (Malzeme) kodları hiçbir şekilde uyuşmuyor!")
            st.write("Katalogdaki Örnek Malzeme Kodları:", list(katalog_malzemeler)[:5])
            st.write("Stoktaki Örnek Malzeme Kodları:", list(stok_malzemeler)[:5])
        elif len(df_merged_stock) == 0:
            st.warning("⚠ TEŞHİS: Ortak malzemeler var ancak birleştirme aşamasında eleniyorlar. Lütfen formatları inceleyin.")

        # --- ASIL CPD KÖPRÜ MANTIK ZİNCİRİ ---
        # Aynı barkoda (EAN Cod-UM) karşılık gelen tüm malzeme stoklarını topluyoruz
        df_barcode_stock_sum = df_merged_stock.groupby(katalog_ean_col)[stok_net_avail_col].sum().reset_index()
        df_barcode_stock_sum.rename(columns={katalog_ean_col: "Barkod"}, inplace=True)

        # --- SİPARİŞ BİRLEŞTİRME VE FİYATLANDIRMA ---
        if "Fiyat" in df_orders.columns:
            df_orders = df_orders.drop(columns=["Fiyat"])
        df_orders_with_price = pd.merge(df_orders, df_prices, on="Barkod", how="left")
        
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
