import streamlit as st
import pandas as pd
import numpy as np
 
# Sayfa ayarları
st.set_page_config(page_title="CPD Order & Stock Allocator Dashboard", layout="wide")
 
st.title("📦 CPD Sipariş Karşılama ve NIV Dashboard")
st.subheader("Distribütör Bekleyen Sipariş & Stok Durumu Analizi")
 
# 1. Dosya Yükleme Alanı
st.sidebar.header("📂 Excel Dosyalarını Yükleyin")
orders_file = st.sidebar.file_uploader("1. Bekleyen Siparişler Excel'i", type=["xlsx", "xls"])
stock_file = st.sidebar.file_uploader("2. Güncel Stok Excel'i", type=["xlsx", "xls"])
 
if orders_file and stock_file:
    # Verileri oku
    df_orders = pd.read_excel(orders_file)
    df_stock = pd.read_excel(stock_file)
    # Kolon isimlerini temizleme (boşlukları silme)
    df_orders.columns = df_orders.columns.str.strip()
    df_stock.columns = df_stock.columns.str.strip()
    st.success("✅ Bekleyen siparişler ve Stok verileri başarıyla yüklendi!")
    # NOT: Stok excelindeki kolon isimlerine göre buralar güncellenecektir.
    # Varsayılan olarak Barkod ve Stok Adeti kolonlarını arıyoruz.
    stok_barkod_col = "Barkod" 
    stok_adet_col = "Stok Adeti"
    if stok_barkod_col in df_stock.columns and stok_adet_col in df_stock.columns:
        df_stock_grouped = df_stock.groupby(stok_barkod_col)[stok_adet_col].sum().reset_index()
        # Sipariş ve Stok Birleştirme (Merge)
        df_merged = pd.merge(df_orders, df_stock_grouped, on="Barkod", how="left")
        df_merged[stok_adet_col] = df_merged[stok_adet_col].fillna(0)
        # Karşılama hesaplama mantığı
        df_merged['Karşılanan Adet'] = np.minimum(df_merged['Sipariş Miktarı'], df_merged[stok_adet_col])
        df_merged['Hesaplanan_Fiyat'] = df_merged['Fiyat'].fillna(0)
        # NIV Tutar Hesaplamaları
        df_merged['Toplam Talep Edilen NIV'] = df_merged['Sipariş Miktarı'] * df_merged['Hesaplanan_Fiyat']
        df_merged['Karşılanan NIV'] = df_merged['Karşılanan Adet'] * df_merged['Hesaplanan_Fiyat']
        df_merged['Kayıp (Karşılanamayan) NIV'] = (df_merged['Sipariş Miktarı'] - df_merged['Karşılanan Adet']) * df_merged['Hesaplanan_Fiyat']
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
        display_cols = ['Müşteri Adı', 'Barkod', 'Ürün Adı', 'Sipariş Miktarı', stok_adet_col, 'Karşılanan Adet', 'Toplam Talep Edilen NIV', 'Karşılanan NIV']
        st.dataframe(df_merged[display_cols].style.format({
            'Toplam Talep Edilen NIV': '₺{:,.2f}',
            'Karşılanan NIV': '₺{:,.2f}'
        }))
    else:
        st.warning(f"⚠ Stok excelinde '{stok_barkod_col}' veya '{stok_adet_col}' kolonları bulunamadı. Lütfen kolon isimlerini kontrol edin.")
else:
    st.info("💡 Lütfen sol menüden 'Sipariş' ve 'Stok' excel dosyalarını yükleyin.")
