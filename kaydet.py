import streamlit as st
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="Mağaza Analiz", layout="wide")

# 1. Mağaza Listesini Tanımlayalım
varsayilan_magazalar = ["M38001", "M38003", "M38002", "M38005", "M38004", "M42001"]

st.title("📊 Mağaza Bazlı Excel Ayıklayıcı")

# Yan menüde mağaza seçimi yapabilmen için bir alan
secilen_kodlar = st.multiselect(
    "Raporlanacak Mağazaları Seçin veya Yazın:",
    options=varsayilan_magazalar,
    default=varsayilan_magazalar
)

yuklenen_dosya = st.file_uploader("Excel Dosyanızı Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    # Excel'i oku
    df = pd.read_excel(yuklenen_dosya, skiprows=2)
    df.columns = df.columns.str.strip()

    if 'Üst Birim' in df.columns:
        # Seçilen kodlara göre filtrele
        filtreli_df = df[df['Üst Birim'].isin(secilen_kodlar)]
        
        # İlk 17 sütunu al
        final_df = filtreli_df.iloc[:, :17]

        if not final_df.empty:
            st.write(f"### Sonuç Tablosu ({len(final_df)} Kayıt)")
            st.dataframe(final_df)

            # Fotoğraf Dönüştürme
            if st.button("🖼️ Fotoğraf Olarak İndir"):
                with st.spinner('Görsel oluşturuluyor...'):
                    # Görseli oluştur
                    dfi.export(final_df, 'tablo_cikti.png', table_conversion='chrome')
                    
                    with open("tablo_cikti.png", "rb") as file:
                        st.download_button(
                            label="Fotoğrafı Kaydet",
                            data=file,
                            file_name="magaza_raporu.png",
                            mime="image/png"
                        )
        else:
            st.warning("Seçilen kodlara ait veri bulunamadı.")
    else:
        st.error("Dosyada 'Üst Birim' sütunu bulunamadı. Lütfen doğru dosyayı yüklediğinizden emin olun.")
