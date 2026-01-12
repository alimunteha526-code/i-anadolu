import streamlit as st
import pandas as pd
import dataframe_image as dfi

st.set_page_config(page_title="Mağaza Raporu", layout="wide")

varsayilan_magazalar = ["M38001", "M38003", "M38002", "M38005", "M38004", "M42001"]

st.title("📊 Mağaza Koduna Özel Rapor")

secilen_kodlar = st.multiselect("Raporlanacak Mağazaları Seçin:", options=varsayilan_magazalar, default=varsayilan_magazalar)

yuklenen_dosya = st.file_uploader("Excel Dosyasını Yükleyin", type=['xlsx'])

if yuklenen_dosya is not None:
    df = pd.read_excel(yuklenen_dosya, skiprows=2)
    df.columns = df.columns.str.strip()

    if 'Üst Birim' in df.columns:
        ust_birim_index = df.columns.get_loc('Üst Birim')
        # Üst birimden başla ve 17 sütun al
        final_df = df.iloc[:, ust_birim_index : ust_birim_index + 17].copy()
        
        # Seçilen mağazalara göre filtrele
        final_df = final_df[final_df['Üst Birim'].isin(secilen_kodlar)]

        if not final_df.empty:
            # --- YÜZDE BİÇİMLENDİRME ---
            # Sütun isminde "Oranı" veya "Hedef" geçenleri bul ve formatla
            oran_sutunlari = [col for col in final_df.columns if 'Oran' in col or 'Hedef' in col]
            
            # Görselleştirme için stil oluşturma
            styled_df = final_df.style.format({col: "{:.1%}" for col in oran_sutunlari})
            
            st.write("### Ayıklanan Tablo (% Biçimli)")
            st.write(styled_df)

            if st.button("🖼️ Fotoğraf Olarak İndir"):
                with st.spinner('Fotoğraf hazırlanıyor...'):
                    resim_yolu = "ozel_cikti.png"
                    # Stil verilmiş tabloyu (styled_df) resme dönüştürüyoruz
                    dfi.export(styled_df, resim_yolu)
                    
                    with open(resim_yolu, "rb") as file:
                        st.download_button(
                            label="Fotoğrafı Kaydet",
                            data=file,
                            file_name="magaza_ozel_rapor.png",
                            mime="image/png"
                        )
        else:
            st.warning("Veri bulunamadı.")
