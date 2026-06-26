import streamlit as st
import pandas as pd
import numpy as np
import os

# Sayfa ayarları
st.set_page_config(page_title="CPD Order & Stock Allocator Dashboard", layout="wide")

st.title("📦 CPD Sipariş Karşılama ve NIV Dashboard")
st.subheader("Distribütör Bekleyen Sipariş, Katalog ve Stok Durumu Analizi")

# --- ARKA PLANDAKİ SİZİN YÜKLEDİĞİNİZ FİYAT LİSTESİNİ OKUMA ---
# Dosya ismini yüklediğiniz "fiyatlar.xlsx" olarak güncelledik
fiyat_dosya_adi = "fiyatlar.xlsx"

@st.cache_data
def fiyat_listesini_yukle():
    if os.path.exists(fiyat_dosya_adi):
        df_fiyat = pd.read_excel(fiyat_dosya_adi, engine="openpyxl")
        df_fiyat.columns = df_fiyat.columns.str.strip()
        if "Barkod" in df_fiyat.columns and "Fiyat" in df_fiyat.columns:
            return df_fiyat[["Barkod", "Fiyat"]].drop_duplicates(subset=["Barkod"])
    return None

df_prices = fiyat_listesini_yukle()

if df_prices is not None:
    st.sidebar.success("✅ Güncel Fiyat Listesi sistemden otomatik yüklendi!")
else:
    st.sidebar.error(f"❌ '{fiyat_dosya_adi}' dosyası GitHub deposunda bulunamadı! Lütfen bu dosyayı GitHub'a yükleyin.")

# --- DOSYA YÜKLEME ALANI (Kullanıcının yükleyeceği 3 dosya) ---
st.sidebar.header("📂 Excel Dosyalarını Yükleyin")
orders_file = st.sidebar.file_uploader("1. Bekleyen Siparişler Excel'i", type=["xlsx", "xls"])
catalog_file = st.sidebar.file_uploader("2. Katalog Raporu Excel'i (Köprü)", type=["xlsx", "xls"])
stock_file = st.sidebar.file_uploader("3. Güncel Stok Excel'i", type=["xlsx", "xls"])

if orders_file and catalog_file and stock_file and df_prices is not None:
    # Verileri oku
    df_orders = pd.read_excel(orders_file, engine="openpyxl")
    df_catalog = pd.read_excel(catalog_file, engine="openpyxl")
    df_stock = pd.read_excel(stock_file, engine="openpyxl")
    
    # Kolon isimlerindeki boşlukları temizleme
    df_orders.columns = df_orders.columns.str.strip()
    df_catalog.columns = df_catalog.columns.str.strip()
    df_stock.columns = df_stock.columns.str.strip()
    
    st.success("✅ Tüm Excel dosyaları başarıyla yüklendi!")
    
    # Beklenen Kritik Kolon İsimleri Tanımları
    siparis_barkod_col = "Barkod"
    katalog_material_col = "Material"
    katalog_ean_col = "EAN Cod-UM"
    stok_material_col = "Material"
    stok_net_avail_col = "Net avail." # Tam belirttiğiniz 'Net avail.' sütun ismi
    
    # Kolon Kontrolleri
    catalog_ok = katalog_material_col in df_catalog.columns and katalog_ean_col in df_catalog.columns
    stock_ok = stok_material_col in df_stock.columns and stok_net_avail_col in df_stock.columns
    orders_ok = siparis_barkod_col in df_orders.columns and "Sipariş Miktarı" in df_orders.columns
    
    if catalog_ok and stock_ok and orders_ok:
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
        display_cols = ['Müşteri Adı', 'Barkod', 'Ürün Adı', 'Sipariş Miktarı', stok_net_avail_col, 'Karşılanan Adet', 'Fiyat', 'Toplam Talep Edilen NIV', 'Karşılanan NIV']
        st.dataframe(df_merged[display_cols].style.format({
            'Fiyat': '₺{:,.2f}',
            'Toplam Talep Edilen NIV': '₺{:,.2f}',
            'Karşılanan NIV': '₺{:,.2f}'
        }))
    else:
        st.warning("⚠ Yüklenen Excel dosyalarındaki kolon isimlerini kontrol edin:")
        if not orders_ok:
            st.write(f"- Sipariş dosyasında '{siparis_barkod_col}' ve 'Sipariş Miktarı' kolonları olmalı.")
        if not catalog_ok:
            st.write(f"- Katalog dosyasında '{katalog_material_col}' ve '{katalog_ean_col}' kolonları olmalı.")
        if not stock_ok:
            st.write(f"- Stok dosyasında '{stok_material_col}' ve '{stok_net_avail_col}' kolonları olmalı.")
else:
