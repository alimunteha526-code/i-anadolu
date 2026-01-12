import streamlit as st
import pandas as pd
import dataframe_image as dfi
import os

# Sayfa başlığı
st.set_page_config(page_title="Excel Bölge Ayıklayıcı", layout="wide")
st.title("📍 İç Anadolu Mağaza Raporlayıcı")

# 1. Dosya Yükleme Alanı
yuklenen_dosya = st.file_uploader("Excel Dosyasını Buraya Bırakın", type=['xlsx', 'csv'])

if yuklenen_dosya is not None:
    # Excel'i oku (Senin dosyanın formatına göre ilk 2 satırı atlıyoruz)
    df = pd.read_excel(yuklenen_dosya, skiprows=2)
    
    # Sütun isimlerini temizle (Başlardaki ve sonlardaki boşlukları siler)
    df.columns = df.columns.str.strip()

    # 2. İç Anadolu Filtrelemesi
    if 'Bölge' in df.columns:
        # Sadece "İÇ ANADOLU" olanları al
        filtreli_df = df[df['Bölge'].str.contains('İÇ ANADOLU', na=False, case=False)]
        
        # 3. İlk 17 Sütunu Seç
        final_df = filtreli_df.iloc[:, :17]

        st.success(f"İç Anadolu bölgesine ait {len(final_df)} mağaza bulundu!")
        st.write("Önizleme (İlk 17 Sütun):")
        st.dataframe(final_df)

        # 4. Fotoğraf Olarak İndirme Butonu
        if st.button("Tabloyu Fotoğrafa Dönüştür"):
            with st.spinner('Resim oluşturuluyor...'):
                resim_adi = "ic_anadolu_rapor.png"
                # Tabloyu resme çevirme işlemi
                dfi.export(final_df, resim_adi)
                
                with open(resim_adi, "rb") as file:
                    st.download_button(
                        label="🖼️ Fotoğrafı İndir",
                        data=file,
                        file_name="ic_anadolu_magaza_listesi.png",
                        mime="image/png"
                    )
    else:
        st.error("Hata: Dosyada 'Bölge' isimli bir sütun bulunamadı!")