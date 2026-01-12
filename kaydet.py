import streamlit as st
import pandas as pd
import dataframe_image as dfi
import os

st.set_page_config(page_title="Mağaza Koduna Göre Rapor", layout="wide")
st.title("📊 Özel Mağaza Analiz Raporu")

# İstediğin mağaza kodlarını bir liste olarak tanımlayalım
hedef_magazalar = ["M38001", "M38003"]

yuklenen_dosya = st.file_uploader("Excel Dosyasını Yükleyin", type=['xlsx', 'csv'])

if yuklenen_dosya is not None:
    # Excel'i oku (İlk 2 satırı atla, 3. satırı başlık yap)
    df = pd.read_excel(yuklenen_dosya, skiprows=2)
    df.columns = df.columns.str.strip() # Sütun isimlerindeki boşlukları temizle

    # 1. Mağaza Koduna (Üst Birim sütununa) göre filtrele
    # 'isin' fonksiyonu listedeki kodların tamamını arar
    if 'Üst Birim' in df.columns:
        filtreli_df = df[df['Üst Birim'].isin(hedef_magazalar)]
        
        # 2. İlk 17 Sütunu Seç
        final_df = filtreli_df.iloc[:, :17]

        if not final_df.empty:
            st.success(f"Seçilen {len(final_df)} mağaza bulundu.")
            st.table(final_df) # Önizleme için tabloyu göster

            # 3. Fotoğraf Oluşturma
            if st.button("Fotoğraf Olarak İndir"):
                with st.spinner('Fotoğraf hazırlanıyor...'):
                    resim_yolu = "magaza_rapor.png"
                    # Tabloyu resme dönüştür
                    dfi.export(final_df, resim_yolu)
                    
                    with open(resim_yolu, "rb") as file:
                        st.download_button(
                            label="🖼️ Fotoğrafı Kaydet",
                            data=file,
                            file_name="ozel_magaza_raporu.png",
                            mime="image/png"
                        )
        else:
            st.warning("Belirlediğiniz mağaza kodları dosyada bulunamadı. Lütfen kodları kontrol edin.")
    else:
        st.error("Hata: Dosyada 'Üst Birim' (mağaza kodu) sütunu bulunamadı!")
