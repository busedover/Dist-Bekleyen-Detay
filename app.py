import streamlit as st
   import pandas as pd
   import numpy as np

   st.set_page_config(page_title="Dist Bekleyen Detay", page_icon="📊", layout="wide")

   st.title("📊 Distribütör Bekleyen Detay Analiz Dashboard'u")
   st.write("Uygulamamıza hoş geldiniz! Burası analiz raporlarımızı görselleştireceğimiz alan.")

   # Örnek Göstergeler (KPI)
   kpi1, kpi2, kpi3 = st.columns(3)
   kpi1.metric(label="Bekleyen Sipariş Adedi", value="1,240 Adet", delta="-5%")
   kpi2.metric(label="Toplam Bekleyen Tutar", value="450,000 TL", delta="12%")
   kpi3.metric(label="Ortalama Bekleme Süresi", value="4.2 Gün", delta="-1.5 Gün")

   # Örnek Grafik Verisi
   st.subheader("📈 Günlük Bekleme Dağılımı")
   chart_data = pd.DataFrame(
       np.random.randn(10, 2),
       columns=['Makyaj', 'Cilt Bakımı']
   )
   st.bar_chart(chart_data)
