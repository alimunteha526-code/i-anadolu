import streamlit as st
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="Mağaza Raporu", layout="wide")

# Mağaza kodları listesi
varsayilan_magazalar = ["M38001", "M38003", "M38002", "M38005", "M38004", "M42001"]

st.title("📊 Mağaza Koduna Özel Rapor")

secilen_kodlar = st.multiselect(
    "Raporlanacak Mağazaları Seçin:",
    options=varsayilan_magazalar,
    default=varsayilan_magazalar
)

yuklenen_dosya = st.file_uploader("Excel Dosyasını Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    # Excel'i oku (İlk 2 satır başlık değil, onları atla)
    df = pd.read_excel(yuklenen_dosya, skiprows=2)
    df.columns = df.columns.str.strip() # Sütun isimlerini temizle

    if 'Üst Birim' in df.columns:
        # 1. ADIM: "Üst Birim" sütununun kaçıncı sırada olduğunu bul
        ust_birim_index = df.columns.get_loc('Üst Birim')

        # 2. ADIM: Bu indexten başlayarak 17 sütun al (Öncesini otomatik siler)
        # Örn: Üst Birim 3. sütunsa, 3'ten 20'ye kadar olanları alır
        final_df = df.iloc[:, ust_birim_index : ust_birim_index + 17]

        # 3. ADIM: Mağaza kodlarına göre filtrele
        final_df = final_df[final_df['Üst Birim'].isin(secilen_kodlar)]

        if not final_df.empty:
            st.write("### Ayıklanan Tablo (İlk Sütun: Üst Birim)")
            st.dataframe(final_df)

            if st.button("🖼️ Fotoğraf Olarak İndir"):
                with st.spinner('Fotoğraf hazırlanıyor...'):
                    resim_yolu = "ozel_cikti.png"
                    # Tabloyu resme dönüştür
                    dfi.export(final_df, resim_yolu)
                    
                    with open(resim_yolu, "rb") as file:
                        st.download_button(
                            label="Fotoğrafı Kaydet",
                            data=file,
                            file_name="magaza_ozel_rapor.png",
                            mime="image/png"
                        )
        else:
            st.warning("Seçilen mağaza kodlarına uygun veri bulunamadı.")
    else:
        st.error("Dosyada 'Üst Birim' sütunu bulunamadı!")
